import asyncio
import sqlite3
from datetime import datetime, timedelta

from aiogram import Bot, Dispatcher, F
from aiogram.types import (
    Message, CallbackQuery,
    InlineKeyboardMarkup, InlineKeyboardButton
)
from aiogram.filters import Command
from openpyxl import Workbook
from openpyxl.styles import Font

# ================= НАСТРОЙКИ =================
TOKEN = "8397597216:AAFtzivDMoNxcRU06vp8wobfG6NU28BkIgs"
DB_NAME = "attendance.db"
EXCEL_NAME = "report.xlsx"
ADMIN_USERNAME = "Glabak0200"  # без @

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
            author TEXT,
            deleted_at TEXT
        )
        """)

# ================= ВСПОМОГАТЕЛЬНОЕ =================
def today():
    return datetime.now().strftime("%Y-%m-%d")

# ================= КНОПКИ =================
def main_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="➕ Отметить отсутствующих", callback_data="mark")],
        [InlineKeyboardButton(text="📄 Выгрузить Excel", callback_data="export")],
        [InlineKeyboardButton(text="♻ Восстановить (30 дней)", callback_data="restore")]
    ])

def students_kb():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=s, callback_data=f"student|{i}")]
        for i, s in enumerate(STUDENTS)
    ])

def reasons_kb(student):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=r, callback_data=f"reason|{student}|{r}")]
        for r in REASONS
    ])

# ================= EXCEL =================
def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Рапортичка"

    headers = ["Дата", "ФИО", "Статус", "Причина", "Кто отметил"]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)

    date = today()

    with db() as con:
        rows = con.execute("""
        SELECT student, reason, author
        FROM attendance
        WHERE date = ? AND deleted_at IS NULL
        """, (date,)).fetchall()

    absent = {r[0]: (r[1], r[2]) for r in rows}

    for s in STUDENTS:
        if s in absent:
            reason, author = absent[s]
            ws.append([date, s, "отсутствовал", reason, author])
        else:
            ws.append([date, s, "присутствовал", "", ""])

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 30

    ws.auto_filter.ref = f"A1:E{ws.max_row}"
    wb.save(EXCEL_NAME)

# ================= BOT =================
bot = Bot(TOKEN)
dp = Dispatcher()

@dp.message(Command("start"))
async def start(msg: Message):
    await msg.answer("📋 Рапортичка группы 101 тп", reply_markup=main_kb())

@dp.callback_query(F.data == "mark")
async def mark(call: CallbackQuery):
    await call.message.answer("Выбери учащегося:", reply_markup=students_kb())
    await call.answer()

@dp.callback_query(F.data.startswith("student|"))
async def choose_student(call: CallbackQuery):
    idx = int(call.data.split("|")[1])
    student = STUDENTS[idx]
    await call.message.answer(f"{student}\nВыбери причину:", reply_markup=reasons_kb(student))
    await call.answer()

@dp.callback_query(F.data.startswith("reason|"))
async def save(call: CallbackQuery):
    _, student, reason = call.data.split("|", 2)

    with db() as con:
        # мягкое удаление старой записи
        con.execute("""
        UPDATE attendance
        SET deleted_at = ?
        WHERE date = ? AND student = ? AND deleted_at IS NULL
        """, (datetime.now().isoformat(), today(), student))

        con.execute("""
        INSERT INTO attendance (date, student, reason, author, deleted_at)
        VALUES (?, ?, ?, ?, NULL)
        """, (
            today(),
            student,
            reason,
            call.from_user.username or call.from_user.full_name
        ))

    await call.message.answer(f"✅ {student} отмечен: {reason}")
    await call.answer()

@dp.callback_query(F.data == "export")
async def export(call: CallbackQuery):
    export_excel()
    await call.message.answer_document(open(EXCEL_NAME, "rb"))
    await call.answer()

@dp.callback_query(F.data == "restore")
async def restore(call: CallbackQuery):
    limit = (datetime.now() - timedelta(days=30)).isoformat()
    with db() as con:
        con.execute("""
        UPDATE attendance
        SET deleted_at = NULL
        WHERE deleted_at IS NOT NULL AND deleted_at >= ?
        """, (limit,))
    await call.message.answer("♻ Записи восстановлены (до 30 дней)")
    await call.answer()

# ================= АВТОСТАРТ =================
async def main():
    init_db()
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
