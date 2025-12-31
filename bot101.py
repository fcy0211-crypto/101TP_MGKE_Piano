import asyncio
import sqlite3
from datetime import datetime

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import (
    Message, CallbackQuery,
    ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton,
    FSInputFile
)

from openpyxl import Workbook
from openpyxl.styles import Font

# ================= НАСТРОЙКИ =================
BOT_TOKEN = "8397597216:AAFtzivDMoNxcRU06vp8wobfG6NU28BkIgs"

ADMIN_USERNAME = "Glabak0200"  # БЕЗ @
ADMIN_CHAT_ID = None

DB_FILE = "attendance.db"
EXCEL_FILE = "rapport_101tp.xlsx"

STUDENTS = [
    "Бабук Владислав",
    "Гарцуев Ростислав",
    "Глинская Милена",
    "Демьянко Надежда",
    "Касьянюк Глеб",
    "Мигутский Тимур",
    "Михальчик Илья",
    "Полторако Артём",
    "Русецкая Кристина",
    "Серяков Игорь",
    "Шаболтас Матвей"
]

REASONS = [
    "по заявлению",
    "по болезни",
    "по неуважительной причине"
]

bot = Bot(BOT_TOKEN)
dp = Dispatcher()

# ================= БАЗА =================
def db():
    return sqlite3.connect(DB_FILE)

def init_db():
    with db() as conn:
        conn.execute("""
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            student TEXT,
            status TEXT,
            reason TEXT,
            author TEXT
        )
        """)
        conn.commit()

# ================= ДАТА =================
def today():
    return datetime.now().strftime("%Y-%m-%d")

# ================= EXCEL =================
def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Рапортичка"

    ws.append(["Дата", "ФИО", "Статус", "Причина", "Кто отметил"])
    for c in ws[1]:
        c.font = Font(bold=True)

    with db() as conn:
        cur = conn.cursor()
        cur.execute("""
        SELECT date, student, status, reason, author
        FROM attendance
        ORDER BY date, student
        """)
        for row in cur.fetchall():
            ws.append(row)

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 30

    wb.save(EXCEL_FILE)

# ================= КЛАВИАТУРА =================
def main_menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📋 Отметить отсутствующих")],
            [KeyboardButton(text="✏ Редактировать рапортичку")],
            [KeyboardButton(text="📤 Выгрузить рапортичку")],
            [KeyboardButton(text="📨 Отправить админу")],
            [KeyboardButton(text="🗑 Очистить рапортичку")]
        ],
        resize_keyboard=True
    )

# ================= START =================
@dp.message(Command("start"))
async def start(msg: Message):
    global ADMIN_CHAT_ID

    if msg.from_user.username == ADMIN_USERNAME:
        ADMIN_CHAT_ID = msg.chat.id
        await msg.answer("✅ Ты назначен администратором")

    await msg.answer(
        "📘 Рапортичка группы 101 тп",
        reply_markup=main_menu()
    )

# ================= ОТМЕТКА =================
@dp.message(F.text == "📋 Отметить отсутствующих")
async def mark(msg: Message):
    kb = [
        [InlineKeyboardButton(text=s, callback_data=f"st|{s}")]
        for s in STUDENTS
    ]
    await msg.answer(
        f"📅 Дата: {today()}",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb)
    )

@dp.callback_query(F.data.startswith("st|"))
async def choose_reason(call: CallbackQuery):
    student = call.data.split("|")[1]
    kb = [
        [InlineKeyboardButton(text=r, callback_data=f"rs|{student}|{r}")]
        for r in REASONS
    ]
    await call.message.answer(
        f"{student}\nПричина отсутствия:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb)
    )

@dp.callback_query(F.data.startswith("rs|"))
async def save(call: CallbackQuery):
    _, student, reason = call.data.split("|", 2)
    with db() as conn:
        conn.execute("""
        INSERT INTO attendance
        (date, student, status, reason, author)
        VALUES (?, ?, 'отсутствовал', ?, ?)
        """, (
            today(),
            student,
            reason,
            call.from_user.username or call.from_user.full_name
        ))
        conn.commit()
    await call.message.answer("✅ Отмечено")

# ================= РЕДАКТИРОВАНИЕ =================
@dp.message(F.text == "✏ Редактировать рапортичку")
async def edit(msg: Message):
    with db() as conn:
        rows = conn.execute(
            "SELECT id, date, student FROM attendance"
        ).fetchall()

    if not rows:
        await msg.answer("Нет записей")
        return

    kb = [
        [InlineKeyboardButton(
            text=f"{r[1]} — {r[2]}",
            callback_data=f"ed|{r[0]}"
        )] for r in rows
    ]
    await msg.answer(
        "Выбери запись:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb)
    )

@dp.callback_query(F.data.startswith("ed|"))
async def edit_reason(call: CallbackQuery):
    rec_id = call.data.split("|")[1]
    kb = [
        [InlineKeyboardButton(
            text=r,
            callback_data=f"upd|{rec_id}|{r}"
        )] for r in REASONS
    ]
    await call.message.answer(
        "Новая причина:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb)
    )

@dp.callback_query(F.data.startswith("upd|"))
async def update(call: CallbackQuery):
    _, rec_id, reason = call.data.split("|", 2)
    with db() as conn:
        conn.execute(
            "UPDATE attendance SET reason=? WHERE id=?",
            (reason, rec_id)
        )
        conn.commit()
    await call.message.answer("✏ Обновлено")

# ================= ВЫГРУЗКА =================
@dp.message(F.text == "📤 Выгрузить рапортичку")
async def export(msg: Message):
    export_excel()
    await msg.answer_document(
        FSInputFile(EXCEL_FILE),
        caption="📊 Рапортичка группы 101 тп"
    )

# ================= АДМИН =================
@dp.message(F.text == "📨 Отправить админу")
async def send_admin(msg: Message):
    if not ADMIN_CHAT_ID:
        await msg.answer("❌ Администратор не написал /start")
        return
    export_excel()
    await bot.send_document(
        ADMIN_CHAT_ID,
        FSInputFile(EXCEL_FILE),
        caption="📨 Рапортичка"
    )
    await msg.answer("✅ Отправлено")

# ================= ОЧИСТКА =================
@dp.message(F.text == "🗑 Очистить рапортичку")
async def clear(msg: Message):
    with db() as conn:
        conn.execute("DELETE FROM attendance")
        conn.commit()
    await msg.answer("🗑 Рапортичка очищена")

# ================= ЗАПУСК =================
async def main():
    init_db()
    print("Бот запущен")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
