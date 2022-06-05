from flask import Flask, redirect, render_template, request, abort
from flask_sqlalchemy import SQLAlchemy

from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage, TemplateSendMessage, ButtonsTemplate, URIAction, MessageAction, FollowEvent, QuickReply, QuickReplyButton
)

import os
import pandas as pd
from datetime import datetime, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import requests

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///info.db'
db = SQLAlchemy(app)

class Post(db.Model):
    id = db.Column(db.String, primary_key=True)
    start_bool = db.Column(db.Integer, nullable=False, default=0)
    rest_bool = db.Column(db.Integer, nullable=False, default=0)
    name = db.Column(db.String(10), nullable=False)
    teacher = db.Column(db.String(10), nullable=False)
    date = db.Column(db.String, nullable=True)
    day = db.Column(db.String(1), nullable=True)
    start = db.Column(db.String, nullable=True)
    end = db.Column(db.String, nullable=True)
    rest = db.Column(db.Integer, nullable=False, default=0)
    time = db.Column(db.Float, nullable=True)
    stay = db.Column(db.String, nullable=True)


YOUR_CHANNEL_ACCESS_TOKEN = '***アクセストークン***'
YOUR_CHANNEL_SECRET =  '***シークレットキー***'

line_bot_api = LineBotApi(YOUR_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(YOUR_CHANNEL_SECRET)

def auth():
    SP_CREDENTIAL_FILE = 'seacret.json'
    SP_SCOPE = [
        'https://spreadsheets.google.com/feeds',
        'https://www.googleapis.com/auth/drive'
    ]

    SP_SHEET_KEY = '***スプレッドシートシークレットキー***'
    SP_SHEET = '***シート名***'

    credentials = ServiceAccountCredentials.from_json_keyfile_name(SP_CREDENTIAL_FILE, SP_SCOPE)
    gc = gspread.authorize(credentials)

    worksheet = gc.open_by_key(SP_SHEET_KEY).worksheet(SP_SHEET)
    return worksheet

def get_h_m_s(seconds):
    m, s = divmod(seconds, 60)
    h, m = map(int, divmod(m, 60))
    return str(h) + ':' + str(m).zfill(2)

def punch_in(id):
    days = ['月', '火', '水', '木', '金', '土', '日']
    timestamp = datetime.now() + timedelta(hours=9)

    post = Post.query.get(id)
    post.start_bool = 1
    post.date = timestamp.strftime('%m/%d')
    post.day = days[timestamp.weekday()]
    post.start = timestamp.strftime('%H:%M')
    name = post.name

    db.session.commit()
    line_bot_api.push_message(id, TextSendMessage(text=f"{name}さんの入室を記録しました。今日も1日頑張ってください！"))

def punch_out(id):
    worksheet = auth()
    df = worksheet.get_all_values(value_render_option='FORMULA')
    timestamp = datetime.now() + timedelta(hours=9)
    punch_out = timestamp.strftime('%H:%M')
    post = Post.query.get(id)
    post.end = punch_out
    post.rest = get_h_m_s(post.rest)
    post.stay = f'=F{len(df)+1}-E{len(df)+1}-G{len(df)+1}'

    date = post.date
    day = post.day
    name = post.name
    teacher = post.teacher
    start = post.start
    end = post.end
    rest = post.rest
    stay = post.stay

    row = [date, day, name, teacher, start, end, rest, stay]
    worksheet.append_row(row, value_input_option='USER_ENTERED')

    post.start_bool = 0
    post.rest = 0
    db.session.commit()

    line_bot_api.push_message(id, TextSendMessage(text=f"{name}さんの退室を記録しました。今日もお疲れ様でした！"))
    df = worksheet.get_all_values(value_render_option='FORMULA')
    if len(df) == 21:
        url = '***GASのスクリプト実行***'
        requests.get(url)
        init(worksheet, df)

def rest_start(id):
    post = Post.query.get(id)
    post.rest_bool = 1
    post.time = time.time()
    name = post.name
    db.session.commit()
    line_bot_api.push_message(id, TextSendMessage(text=f"{name}さんの休憩を記録しました。"))

def rest_end(id):
    post = Post.query.get(id)
    start = post.time
    end = time.time()
    rest_time = end - start
    post.rest += rest_time
    name = post.name
    post.rest_bool = 0
    db.session.commit()
    line_bot_api.push_message(id, TextSendMessage(text=f"{name}さんの外出時間を記録しました。おかえりなさいませ！"))

def init(worksheet, df):
  for i in range(6,21):
    df[i] = ['', '', '', '', '', '', '', '', '']
  worksheet.update(df, value_input_option='USER_ENTERED')

def register_template(user_id, display_name):
    message_template = TemplateSendMessage(
        alt_text="研究室登録",
        template=ButtonsTemplate(
            text="以下より研究室登録をしてください。",
            title=f"ようこそ、{display_name}さん",
            image_size="cover",
            thumbnail_image_url="https://cdn.pixabay.com/photo/2022/05/16/18/17/sheep-7200918_1280.jpg",
            actions=[
                URIAction(
                    uri=f"https://management-bot-20220519.herokuapp.com/?id={user_id}&name={display_name}",
                    label="登録"
                )
            ]
        )
    )
    return message_template

def IN(name):
    message_template = TemplateSendMessage(
        alt_text="入室登録",
        template=ButtonsTemplate(
            text="以下より入室登録ができます。",
            title=f"こんにちは、{name}さん",
            image_size="cover",
            thumbnail_image_url="https://cdn.pixabay.com/photo/2022/05/16/18/17/sheep-7200918_1280.jpg",
            actions=[
                MessageAction(
                    type = "message",
                    label = "入室する",
                    text = "入室する"
                )
            ]
        )
    )
    return message_template

def OUT1(name):
    message_template = TemplateSendMessage(
        alt_text="退室OR休憩登録",
        template=ButtonsTemplate(
            text="以下より休憩もしくは退室登録ができます。",
            title=f"お疲れ様です、{name}さん",
            image_size="cover",
            thumbnail_image_url="https://cdn.pixabay.com/photo/2022/05/16/18/17/sheep-7200918_1280.jpg",
            actions=[
                MessageAction(
                    type = "message",
                    label = "休憩する",
                    text = "休憩する"
                ),
                MessageAction(
                    type = "message",
                    label = "退室する",
                    text = "退室する"
                )
            ]
        )
    )
    return message_template

def OUT2(name):
    message_template = TemplateSendMessage(
        alt_text="休憩終了登録",
        template=ButtonsTemplate(
            text="以下より休憩を終了できます。",
            title=f"おかえりなさいませ、{name}さん",
            image_size="cover",
            thumbnail_image_url="https://cdn.pixabay.com/photo/2022/05/16/18/17/sheep-7200918_1280.jpg",
            actions=[
                MessageAction(
                    type = "message",
                    label = "休憩終了",
                    text = "休憩終了"
                ),
            ]
        )
    )
    return message_template

@app.route("/", methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        req = request.args
        post = {"user_id": req.get("id"),
                "display_name": req.get("name"),
                "DB": Post.query.all()}
        return render_template('index.html', post=post)
    
    elif request.method == 'POST':
        display_name = request.form.get('display_name')
        id = request.form.get('id')
        name = request.form.get('name')
        teacher = request.form.get('teacher')
        info = Post(id=id, name=name, teacher=teacher)

        db.session.add(info)
        db.session.commit()

        post = {
            "id": id,
            "name": name,
            "teacher": teacher
        }

        line_bot_api.push_message(id, TextSendMessage(text=f"{name}さんの登録が完了しました！"))

        return render_template('completion.html', post=post)

@app.route("/delete/<string:id>")
def delete(id):
    post = Post.query.get(id)
    db.session.delete(post)
    db.session.commit()
    return redirect('/')

@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']

    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        print("Invalid signature. Please check your channel access token/channel secret.")
        abort(400)

    return 'OK'



@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    profile = line_bot_api.get_profile(event.source.user_id)
    user_id = profile.user_id
    display_name = profile.display_name
    if event.message.text == "登録":
        messages = register_template(user_id, display_name)
        line_bot_api.reply_message(
            event.reply_token,
            messages
        )

    elif event.message.text == "打刻":
        post = Post.query.get(user_id)
        # まだ入室していない場合
        if not post.start_bool:
            messages = IN(post.name)
            line_bot_api.reply_message(
                event.reply_token,
                messages
            )
        # 既に入室していて休憩を開始していない場合、退室OR休憩開始
        elif not post.rest_bool:
            messages = OUT1(post.name)
            line_bot_api.reply_message(
                event.reply_token,
                messages
            )

        # 休憩している場合
        else:
            messages = OUT2(post.name)
            line_bot_api.reply_message(
                event.reply_token,
                messages
            )


    elif event.message.text == "入室する":
        post = Post.query.get(user_id)
        if not post.start_bool and not post.rest_bool:
            punch_in(user_id)

    elif event.message.text == "休憩する":
        post = Post.query.get(user_id)
        if post.start_bool and not post.rest_bool:
            rest_start(user_id)

    elif event.message.text == "休憩終了":
        post = Post.query.get(user_id)
        if post.start_bool and post.rest_bool:
            rest_end(user_id)

    elif event.message.text == "退室する":
        post = Post.query.get(user_id)
        if post.start_bool and not post.rest_bool:
            punch_out(user_id)

    elif event.message.text == "管理表":
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="***スプレッドシートURL***"))
    else:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text=event.message.text))


if __name__ == "__main__":
    port  = os.getenv("PORT")
    app.run(host="0.0.0.0" ,port=port)
