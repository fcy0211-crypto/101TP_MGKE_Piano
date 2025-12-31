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
ADMIN_USERNAME = "Glabak0200"  # без @

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
            author TEXT,
            deleted_at TEXT
        )
        """)
        con.commit()

def today():
    return datetime.now().strftime("%Y-%m-%d")

def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# ================= EXCEL (С ЦВЕТАМИ) =================
def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Рапортичка"

    header_fill = PatternFill("solid", fgColor="DDDDDD")
    green_fill = PatternFill("solid", fgColor="C6EFCE")
    red_fill = PatternFill("solid", fgColor="FFC7CE")

    headers = ["Дата", "ФИО", "Статус", "Причина", "Кто отметил"]
    ws.append(headers)

    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = header_fill

    with db() as con:
        dates = con.execute("""
        SELECT DISTINCT date FROM attendance
        WHERE deleted_at IS NULL
        ORDER BY date
        """).fetchall()

    for (date,) in dates:
        with db() as con:
            rows = con.execute("""
            SELECT student, reason, author
            FROM attendance
            WHERE date = ? AND deleted_at IS NULL
            """, (date,)).fetchall()

        absent = {r[0]: (r[1], r[2]) for r in rows}

        for student in sorted(STUDENTS):
            if student in absent:
                reason, author = absent[student]
                ws.append([date, student, "отсутствовал", reason, author])
                for c in ws[ws.max_row]:
                    c.fill = red_fill
            else:
                ws.append([date, student, "присутствовал", "", ""])
                for c in ws[ws.max_row]:
                    c.fill = green_fill

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 30

    ws.auto_filter.ref = f"A1:E{ws.max_row}"
    wb.save(EXCEL_NAME)

# ================= КЛАВИАТУРА =================
def menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="➕ Отметить")],
            [KeyboardButton(text="✏ Редактировать")],
            [KeyboardButton(text="📤 Выгрузить")],
            [KeyboardButton(text="📨 Админу")],
            [KeyboardButton(text="🗑 Очистить")],
            [KeyboardButton(text="♻ Восстановить")]
        ],
        resize_keyboard=True
    )

# ================= START =================
@dp.message(Command("start"))
async def start(msg: Message):
    global ADMIN_CHAT_ID
    if msg.from_user.username == ADMIN_USERNAME:
        ADMIN_CHAT_ID = msg.chat.id
        await msg.answer("✅ Ты администратор")

    await msg.answer("📘 Рапортичка 101 тп", reply_markup=menu())

# ================= ОТМЕТКА =================
@dp.message(lambda m: m.text == "➕ Отметить")
async def choose_student(msg: Message):
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=s, callback_data=f"s{i}")]
            for i, s in enumerate(STUDENTS)
        ]
    )
    await msg.answer(f"Дата: {today()}", reply_markup=kb)

@dp.callback_query(lambda c: c.data.startswith("s"))
async def choose_reason(call: CallbackQuery):
    idx = int(call.data[1:])
    student = STUDENTS[idx]

    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=r, callback_data=f"r{idx}|{i}")]
            for i, r in enumerate(REASONS)
        ]
    )
    await call.message.answer(student, reply_markup=kb)

@dp.callback_query(lambda c: c.data.startswith("r"))
async def save(call: CallbackQuery):
    left, reason_idx = call.data[1:].split("|")
    student = STUDENTS[int(left)]
    reason = REASONS[int(reason_idx)]

    with db() as con:
        con.execute("""
        INSERT INTO attendance (date, student, reason, author, deleted_at)
        VALUES (?, ?, ?, ?, NULL)
        """, (
            today(),
            student,
            reason,
            call.from_user.username or call.from_user.full_name
        ))
        con.commit()

    await call.message.answer("✅ Отмечено")

# ================= РЕДАКТИРОВАНИЕ =================
@dp.message(lambda m: m.text == "✏ Редактировать")
async def edit(msg: Message):
    with db() as con:
        rows = con.execute("""
        SELECT id, date, student, reason
        FROM attendance
        WHERE deleted_at IS NULL
        """).fetchall()

    if not rows:
        await msg.answer("Нет записей")
        return

    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(
                text=f"{r[1]} | {r[2]}",
                callback_data=f"e{r[0]}"
            )] for r in rows
        ]
    )
    await msg.answer("Выбери запись:", reply_markup=kb)

@dp.callback_query(lambda c: c.data.startswith("e"))
async def edit_reason(call: CallbackQuery):
    rec_id = int(call.data[1:])
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=r, callback_data=f"u{rec_id}|{i}")]
            for i, r in enumerate(REASONS)
        ]
    )
    await call.message.answer("Новая причина:", reply_markup=kb)

@dp.callback_query(lambda c: c.data.startswith("u"))
async def update(call: CallbackQuery):
    rec_id, reason_idx = call.data[1:].split("|")
    reason = REASONS[int(reason_idx)]

    with db() as con:
        con.execute(
            "UPDATE attendance SET reason=? WHERE id=?",
            (reason, int(rec_id))
        )
        con.commit()

    await call.message.answer("✏ Обновлено")

# ================= ВЫГРУЗКА =================
@dp.message(lambda m: m.text == "📤 Выгрузить")
async def export(msg: Message):
    export_excel()
    await msg.answer_document(FSInputFile(EXCEL_NAME))

# ================= АДМИН =================
@dp.message(lambda m: m.text == "📨 Админу")
async def send_admin(msg: Message):
    if not ADMIN_CHAT_ID:
        await msg.answer("Админ не активен")
        return

    export_excel()
    await bot.send_document(
        ADMIN_CHAT_ID,
        FSInputFile(EXCEL_NAME),
        caption="📊 Рапортичка"
    )
    await msg.answer("✅ Отправлено")

# ================= ОЧИСТКА =================
@dp.message(lambda m: m.text == "🗑 Очистить")
async def clear(msg: Message):
    with db() as con:
        con.execute(
            "UPDATE attendance SET deleted_at=? WHERE deleted_at IS NULL",
            (now(),)
        )
        con.commit()
    await msg.answer("🗑 Очищено (восстановимо 30 дней)")

# ================= ВОССТАНОВЛЕНИЕ =================
@dp.message(lambda m: m.text == "♻ Восстановить")
async def restore(msg: Message):
    limit = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
    with db() as con:
        con.execute("""
        UPDATE attendance
        SET deleted_at=NULL
        WHERE deleted_at IS NOT NULL
        AND deleted_at >= ?
        """, (limit,))
        con.commit()
    await msg.answer("♻ Восстановлено")

# ================= ЗАПУСК =================
async def main():
    init_db()
    print("Бот запущен")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
