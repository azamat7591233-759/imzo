"""
Obyasnitel boti (sxema bo'yicha) + barmoq bilan imzo (Telegram Mini App)

O'rnatish:  pip install python-telegram-bot gspread openpyxl Pillow
Ishga tushirish:  python bot_sheets.py

Saqlash:
  - credentials.json fayli bo'lsa  -> Google Sheets'ga yozadi (real vaqtda)
  - bo'lmasa                       -> arizalar.xlsx fayliga yozadi

Imzo sahifasi (sign.html) HTTPS manzilda turishi kerak (masalan GitHub Pages).
Uning manzilini WEBAPP_URL ga yozing.
"""
import io
import os
import json
import asyncio
import threading
from datetime import datetime
from telegram import (Update, InlineKeyboardButton, InlineKeyboardMarkup,
                      KeyboardButton, ReplyKeyboardMarkup, ReplyKeyboardRemove,
                      WebAppInfo)
from telegram.ext import (Application, CommandHandler, MessageHandler,
                          CallbackQueryHandler, ContextTypes, filters)
from http.server import BaseHTTPRequestHandler, HTTPServer

# ---------------- SOZLAMALAR ----------------
TOKEN = os.getenv("BOT_TOKEN", "9053618634:AAHRvyBkPUJ9Bn_WRqxcsaiP7Dy_XiZDuto")   # BotFather tokeni
ACCESS_CODE = os.getenv("ACCESS_CODE", "12A33")           # kirish kodi
# Jadval manzilidagi /d/ dan keyingi va /edit dan oldingi uzun kod:
# docs.google.com/spreadsheets/d/  SHU_YERDAGI_KOD  /edit
SHEET_ID = os.getenv("SHEET_ID", "19SzqbUUUPYlhWHlJGCFYWzP_O-sGjRczdMtHHIkOg3o")
# Imzo sahifasining HTTPS manzili (sign.html joylangan joy)
WEBAPP_URL = os.getenv("WEBAPP_URL", "https://https://azamat7591233-759.github.io/imzo/sign.html")
CREDS_FILE = "credentials.json"
if not os.path.exists(CREDS_FILE) and os.path.exists("/etc/secrets/credentials.json"):
     CREDS_FILE ="/etc/secrets/credentials.json"                                                
EXCEL_FILE = "arizalar.xlsx"
AUTH_FILE = "authorized.json"  # kodni kiritgan foydalanuvchilar
# Rasmlar va imzolar yuboriladigan guruh/kanal ID (bot u yerda admin bo'lishi kerak).
# Masalan: -1001234567890. None bo'lsa jadvalga file_id yoziladi.
PHOTO_CHAT_ID = None
# --------------------------------------------

BACK_TXT = "⬅️ Назад"

# (kalit, savol)
STEPS = [
    ("date", "📅 Sanani tasdiqlang yoki boshqa sanani yozing (KK.OO.YYYY):"),
    ("dept", "🏢 Otdelni yozing:"),
    ("operator", "👤 Operator F.I.O:"),
    ("photo", "📷 Rasmni yuboring:"),
    ("comment", "💬 Izoh yozing (yoki «Izohsiz» tugmasini bosing):"),
    ("shift", "🕐 Smenani tanlang:"),
    ("brigadir", "👷 Brigadir F.I.O:"),
    ("lider", "⭐ Lider F.I.O:"),
    ("dkk", "🔍 ДКК F.I.O:"),
    ("sign", "✍️ Pastdagi tugmani bosing va imzoni barmoq bilan chizing:"),
]
# Jadvaldagi ustunlar tartibi: A..J (K, L — qo'shimcha)
HEADERS = ["Дата", "Отдел", "оператор Ф.И.О", "расм", "Комментария", "смена",
           "Бригадир", "Лидер", "ДКК Ф.И.О", "Подпись", "Telegram ID", "Vaqt"]


def today():
    return datetime.now().strftime("%d.%m.%Y")


# ---------- foydalanuvchilar ruxsati ----------
def load_auth():
    if os.path.exists(AUTH_FILE):
        with open(AUTH_FILE) as f:
            return set(json.load(f))
    return set()


def save_auth(users):
    with open(AUTH_FILE, "w") as f:
        json.dump(list(users), f)


AUTHORIZED = load_auth()


# ---------- imzoni rasmga aylantirish ----------
def render_signature(strokes, w=300, h=150, scale=3):
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (w * scale, h * scale), "white")
    dr = ImageDraw.Draw(img)
    for s in strokes:
        pts = [(s[i] * scale, s[i + 1] * scale) for i in range(0, len(s) - 1, 2)]
        if len(pts) == 1:
            x, y = pts[0]
            dr.ellipse((x - 3, y - 3, x + 3, y + 3), fill="black")
        elif len(pts) > 1:
            dr.line(pts, fill="black", width=5, joint="curve")
    bio = io.BytesIO()
    img.save(bio, "PNG")
    bio.seek(0)
    return bio


# ---------- saqlash ----------
def save_row(row):
    if os.path.exists(CREDS_FILE):
        import gspread
        gc = gspread.service_account(filename=CREDS_FILE)
        ws = gc.open_by_key(SHEET_ID).sheet1
        if not ws.get_all_values():
            ws.append_row(HEADERS)
        ws.append_row(row,)
    else:
        from openpyxl import Workbook, load_workbook
        if os.path.exists(EXCEL_FILE):
            wb = load_workbook(EXCEL_FILE)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.append(HEADERS)
        ws.append(row)
        wb.save(EXCEL_FILE)


def msg_link(message_id):
    cid = str(PHOTO_CHAT_ID)
    if cid.startswith("-100"):
        return f"https://t.me/c/{cid[4:]}/{message_id}"
    return None


# ---------- savol yuborish ----------
async def ask(context, chat_id):
    idx = context.user_data["step"]
    key, text = STEPS[idx]

    if key == "sign":
        kb = ReplyKeyboardMarkup(
            [[KeyboardButton("✍️ Imzo chizish", web_app=WebAppInfo(url=WEBAPP_URL))],
             [KeyboardButton(BACK_TXT)]],
            resize_keyboard=True)
        await context.bot.send_message(chat_id, text, reply_markup=kb)
        return

    rows = []
    if key == "date":
        rows.append([InlineKeyboardButton(f"✅ {today()}", callback_data="date_ok")])
    if key == "comment":
        rows.append([InlineKeyboardButton("Izohsiz", callback_data="skip_comment")])
    if key == "shift":
        rows.append([InlineKeyboardButton("1-смена", callback_data="shift:1-смена")])
        rows.append([InlineKeyboardButton("2-смена", callback_data="shift:2-смена")])
        rows.append([InlineKeyboardButton("3-смена", callback_data="shift:3-смена")])
    if idx > 0:
        rows.append([InlineKeyboardButton(BACK_TXT, callback_data="back")])
        await context.bot.send_message(
        chat_id, text,
        reply_markup=InlineKeyboardMarkup(rows) if rows else None)


async def next_step(context, chat_id):
    context.user_data["step"] += 1
    if context.user_data["step"] >= len(STEPS):
        await finish(update, context, chat_id)
    else:
        await ask(context, chat_id)

async def finish(update, context, chat_id): 
  d = context.user_data["data"]
  photo_file = await context.bot.get_file(d["photo"])
  photo_url = f"https://api.telegram.org/file/bot{TOKEN}/{photo_file.file_path}"

  sign_file = await context.bot.get_file(d["sign"])
  sign_url = f"https://api.telegram.org/file/bot{TOKEN}/{sign_file.file_path}"

  row = [
    d["date"],
    d["dept"],
    d["operator"],
    f'=IMAGE("{photo_url}")',
    d["comment"],
    d["shift"],
    d["brigadier"],
    d["lider"],
    d["dkk"],
    f'=IMAGE("{sign_url}")',
    str(chat_id),
    datetime.now().strftime("%d.%m.%Y %H:%M"),
]

try:
    await asyncio.to_thread(save_row, row)
  text = "✅ Ma'lumotlar jadvalga muvaffaqiyatli saqlandi!"
except Exception as e:
    text = f"❌ Jadvalga yozishda xato: {e}"

rows = []
rows.append([InlineKeyboardButton("➕ Yangi ariza", callback_data="now")])

await context.bot.send_message(
    chat_id,
    text,
    reply_markup=InlineKeyboardMarkup(rows)
)
context.user_data.clear()
await context.bot.send_message(
        chat_id, text,
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("➕ Yangi ariza", callback_data="new")]]))


async def begin_form(context, chat_id):
    context.user_data["step"] = 0
    context.user_data["data"] = {}
    await ask(context, chat_id)


# ---------- handlerlar ----------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    chat_id = update.effective_chat.id
    uid = update.effective_user.id
    if uid in AUTHORIZED:
        await context.bot.send_message(chat_id, "🔄 Yangi ariza", reply_markup=ReplyKeyboardRemove())
        await begin_form(context, chat_id)
    else:
        context.user_data["step"] = -1  # kod kutilmoqda
        await update.message.reply_text("🔑 Kodni kiriting:", reply_markup=ReplyKeyboardRemove())


async def on_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    step = context.user_data.get("step")
    chat_id = update.effective_chat.id
    text = update.message.text.strip()

    if step is None:
        await update.message.reply_text("Boshlash uchun /start bosing.")
        return

    if step == -1:  # kod
        if text == ACCESS_CODE:
            AUTHORIZED.add(update.effective_user.id)
            save_auth(AUTHORIZED)
            await begin_form(context, chat_id)
        else:
            await update.message.reply_text("❌ Kod noto'g'ri, qayta kiriting.")
        return

    key = STEPS[step][0]

    if key == "sign":
        if text == BACK_TXT:
            await update.message.reply_text("↩️", reply_markup=ReplyKeyboardRemove())
            context.user_data["step"] = step - 1
            await ask(context, chat_id)
        else:
            await update.message.reply_text("Pastdagi «✍️ Imzo chizish» tugmasini bosing.")
        return

    if key in ("photo", "shift"):
        await update.message.reply_text("Iltimos, so'ralgan narsani yuboring (rasm yoki tugma).")
        return
    if key == "date":
        try:
            datetime.strptime(text, "%d.%m.%Y")
        except ValueError:
            await update.message.reply_text("Sana noto'g'ri. Masalan: 05.11.2026")
            return
    context.user_data["data"][key] = text
    await next_step(context, chat_id)


async def on_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    step = context.user_data.get("step")
    if step is None or step < 0 or STEPS[step][0] != "photo":
        return
    context.user_data["data"]["photo"] = update.message.photo[-1].file_id
    await next_step(context, update.effective_chat.id)


async def on_webapp_data(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Mini App'dan kelgan imzo (chiziqlar) -> rasm."""
    step = context.user_data.get("step")
    if step is None or step < 0 or STEPS[step][0] != "sign":
        return
    try:
        payload = json.loads(update.message.web_app_data.data)
        strokes = [[int(v) for v in s] for s in payload["s"] if isinstance(s, list) and s]
        if not strokes or sum(len(s) for s in strokes) > 5000:
            raise ValueError("bad strokes")
        img = await asyncio.to_thread(render_signature, strokes)
    except Exception:
        await update.message.reply_text("❌ Imzo o'qilmadi, qayta chizing.")
        return

    msg = await update.message.reply_photo(
        img, caption="✍️ Imzo qabul qilindi", reply_markup=ReplyKeyboardRemove())
    context.user_data["data"]["sign"] = msg.photo[-1].file_id
    await next_step(context, update.effective_chat.id)


async def on_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    chat_id = q.message.chat_id

    if q.data == "new":
        if update.effective_user.id in AUTHORIZED:
            await begin_form(context, chat_id)
        return
    step = context.user_data.get("step")
    if step is None or step < 0:
        return

    if q.data == "back":
        context.user_data["step"] = max(0, step - 1)
        await ask(context, chat_id)
    elif q.data == "skip_comment":
        context.user_data["data"]["comment"] = "-"
        await next_step(context, chat_id)
    elif q.data == "date_ok":
        context.user_data["data"]["date"] = today()
        await next_step(context, chat_id)
    elif q.data.startswith("shift:"):
        context.user_data["data"]["shift"] = q.data.split(":", 1)[1]
        await next_step(context, chat_id)


def start_health_server():
    """Render Web Service ochiq port talab qiladi. Kompyuterda ishlamaydi (PORT yo'q)."""
    port = os.getenv("PORT")
    if not port:
        return
    from http.server import BaseHTTPRequestHandler, HTTPServer
    
class Health(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"ok")

    def do_HEAD(self):
        self.send_response(200)
        self.end_headers()

    def log_message(self, *args):
        pass

port= int(os.getenv("PORT",1000))
server = HTTPServer(("0.0.0.0", int(port)), Health)
threading.Thread(target=server.serve_forever, daemon=True).start()

def main():
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(on_button))
    app.add_handler(MessageHandler(filters.StatusUpdate.WEB_APP_DATA, on_webapp_data))
    app.add_handler(MessageHandler(filters.PHOTO, on_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text))
    app.run_polling()


if __name__ == "__main__":
    main()
