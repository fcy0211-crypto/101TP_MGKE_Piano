import asyncio
import sqlite3
from datetime import datetime, timedelta

from aiogram import Bot, Dispatcher
from aiogram.types import (
    Message, CallbackQuery,
    ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton,
    FSInputFile
)
from aiogram.filters import Command

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# ================= НАСТРОЙКИ =================
BOT_TOKEN = "8397597216:AAFtzivDMoNxcRU06vp8wobfG6NU28BkIgs"
ADMIN_USERNAME = "Glabak0200"

DB_NAME = "attendance.db"
EXCEL_NAME = "rapport.xlsx"

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

HOURS = [1, 2, 3, 4, 5, 6]

bot = Bot(BOT_TOKEN)
dp = Dispatcher()
ADMIN_CHAT_ID = None

# ================= БАЗА =================
def db():
    return sqlite3.connect(DB_NAME)

def init_db():
    with db() as con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            student TEXT,
            reason TEXT,
            hours INTEGER,
            author TEXT,
            deleted_at TEXT
        )
        """)
        con.commit()

def today():
    return datetime.now().strftime("%Y-%m-%d")

def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# ================= EXCEL =================
def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Рапортичка"

    headers = ["Дата", "ФИО", "Статус", "Причина", "Часы", "Кто отметил"]
    ws.append(headers)

    for c in ws[1]:
        c.font = Font(bold=True)

    with db() as con:
        rows = con.execute("""
        SELECT date, student, reason, hours, author
        FROM attendance
        WHERE deleted_at IS NULL
        ORDER BY date
        """).fetchall()

    for r in rows:
        ws.append([r[0], r[1], "отсутствовал", r[2], r[3], r[4]])

    wb.save(EXCEL_NAME)

# ================= КЛАВИАТУРА =================
def menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="➕ Отметить")],
            [KeyboardButton(text="✏ Редактировать")],
            [KeyboardButton(text="📤 Выгрузить")]
        ],
        resize_keyboard=True
    )

# ================= START =================
@dp.message(Command("start"))
async def start(msg: Message):
    await msg.answer("📘 Рапортичка", reply_markup=menu())

# ================= ОТМЕТКА =================
@dp.message(lambda m: m.text == "➕ Отметить")
async def choose_student(msg: Message):
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=s, callback_data=f"s{i}")]
            for i, s in enumerate(STUDENTS)
        ]
    )
    await msg.answer("Выбери студента:", reply_markup=kb)

@dp.callback_query(lambda c: c.data.startswith("s"))
async def choose_reason(call: CallbackQuery):
    await call.answer()

    idx = int(call.data[1:])
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=r, callback_data=f"r{idx}|{i}")]
            for i, r in enumerate(REASONS)
        ]
    )
    await call.message.answer(STUDENTS[idx], reply_markup=kb)

@dp.callback_query(lambda c: c.data.startswith("r"))
async def choose_hours(call: CallbackQuery):
    await call.answer()

    s_idx, r_idx = call.data[1:].split("|")
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=str(h), callback_data=f"h{s_idx}|{r_idx}|{h}")]
            for h in HOURS
        ]
    )
    await call.message.answer("Сколько часов отсутствовал?", reply_markup=kb)

@dp.callback_query(lambda c: c.data.startswith("h"))
async def save(call: CallbackQuery):
    await call.answer()

    s_idx, r_idx, hours = call.data[1:].split("|")

    with db() as con:
        con.execute("""
        INSERT INTO attendance (date, student, reason, hours, author, deleted_at)
        VALUES (?, ?, ?, ?, ?, NULL)
        """, (
            today(),
            STUDENTS[int(s_idx)],
            REASONS[int(r_idx)],
            int(hours),
            call.from_user.username
        ))
        con.commit()

    await call.message.answer("✅ Отмечено")

# ================= ВЫГРУЗКА =================
@dp.message(lambda m: m.text == "📤 Выгрузить")
async def export(msg: Message):
    export_excel()
    await msg.answer_document(FSInputFile(EXCEL_NAME))

# ================= ЗАПУСК =================
async def main():
    init_db()
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
