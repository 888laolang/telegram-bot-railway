import asyncio
import pandas as pd
import os
import time
from datetime import datetime, timedelta
import pytz
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, filters, ContextTypes

# 配置
EXCEL_FILE = "卡池.xlsx"
BOT_TOKEN = os.getenv("8319773854:AAHOi3S3ya4u40Uhj-XQnfWPfwug0Y7dFaw")
CHAT_ID = int(os.getenv("7136882977"))  # 你的 Telegram 用户ID
REPLY_SEPARATOR = "|"

# 全局变量
keywords = {}
last_modified = 0

def load_keywords():
    global keywords, last_modified
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        combined = {}
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name)
            if df.shape[1] < 16:
                continue
            keys = df.iloc[:, 6].astype(str)  # G列
            replies = df.iloc[:, 15].astype(str)  # P列
            temp_df = pd.DataFrame({'key': keys, 'reply': replies})
            temp_df = temp_df[temp_df['key'].notna() & temp_df['reply'].notna()]
            temp_df = temp_df[(temp_df['key'].str.lower() != 'nan') & (temp_df['reply'].str.lower() != 'nan')]
            sheet_dict = dict(zip(temp_df['key'], temp_df['reply']))
            combined.update(sheet_dict)
        keywords = combined
        last_modified = os.path.getmtime(EXCEL_FILE)
        print(f"[{time.strftime('%H:%M:%S')}] 已加载关键词：{len(keywords)} 条")
    except Exception as e:
        print("读取 Excel 出错:", e)

def check_excel_update():
    global last_modified
    try:
        mtime = os.path.getmtime(EXCEL_FILE)
        if mtime != last_modified:
            load_keywords()
    except FileNotFoundError:
        pass

async def delete_message_later(message):
    await asyncio.sleep(90)
    try:
        await message.delete()
    except:
        pass

async def auto_reply(update: Update, context: ContextTypes.DEFAULT_TYPE):
    check_excel_update()
    text = update.message.text.strip()
    matched_replies = []
    for key, reply_text in keywords.items():
        if key in text:
            parts = [msg.strip() for msg in reply_text.split(REPLY_SEPARATOR) if msg.strip()]
            matched_replies.extend(parts)
    if matched_replies:
        full_reply = "\n".join(matched_replies)
        sent_msg = await update.message.reply_text(full_reply, disable_notification=True)
        asyncio.create_task(delete_message_later(sent_msg))

async def scheduled_send(context):
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name)
            if df.shape[1] < 16:
                continue
            df_c0 = df[df.iloc[:, 2] == 0]  # C列=0
            if not df_c0.empty:
                df_c0 = df_c0.sort_values(by=df_c0.columns[15].map(lambda x: str(x)[0] if pd.notna(x) else ""))
                messages = df_c0.iloc[:, 15].dropna().astype(str).tolist()
                if messages:
                    text = f"【{sheet_name}】\n" + "\n".join(messages)
                    await context.bot.send_message(chat_id=CHAT_ID, text=text, disable_notification=True)
    except Exception as e:
        print("定时任务出错:", e)

async def schedule_daily_task(app):
    tz = pytz.timezone("Asia/Shanghai")
    while True:
        now = datetime.now(tz)
        target = tz.localize(datetime(now.year, now.month, now.day, 7, 0, 0))
        if now > target:
            target += timedelta(days=1)
        wait_seconds = (target - now).total_seconds()
        await asyncio.sleep(wait_seconds)
        await scheduled_send(app)

if __name__ == '__main__':
    load_keywords()
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, auto_reply))
    asyncio.get_event_loop().create_task(schedule_daily_task(app))
    app.run_polling()
