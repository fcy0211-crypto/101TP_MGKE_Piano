import asyncio
import sqlite3
from datetime import datetime, timedelta

from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import (
    ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton,
    FSInputFile
)

from openpyxl import Workbook
from openpyxl.styles import Font

# ================== НАСТРОЙКИ ==================
BOT_TOKEN = "8397597216:AAFtzivDMoNxcRU06vp8wobfG6NU28BkIgs"
ADMIN_USERNAME = "Glabak0200"  # без @

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
ADMIN_CHAT_ID = None

# ================== БАЗА ==================
def db():
    return sqlite3.connect(DB_FILE)

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

# ================== EXCEL ==================
def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Рапортичка"

    ws.append(["Дата", "ФИО", "Статус", "Причина", "Кто отметил"])
    for c in ws[1]:
        c.font = Font(bold=True)

    with db() as con:
        rows = con.execute("""
        SELECT date, student, 'отсутствовал', reason, author
        FROM attendance
        WHERE deleted_at IS NULL
        ORDER BY date, student
        """).fetchall()

    for r in rows:
        ws.append(r)

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 30

    wb.save(EXCEL_FILE)

# ================== КЛАВИАТУРА ==================
def menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📋 Отметить отсутствующих")],
            [KeyboardButton(text="✏ Редактировать рапортичку")],
            [KeyboardButton(text="📤 Выгрузить рапортичку")],
            [KeyboardButton(text="📨 Отправить админу")],
            [KeyboardButton(text="🗑 Очистить рапортичку")],
            [KeyboardButton(text="♻ Восстановить за месяц")]
        ],
        resize_keyboard=True
    )

# ================== START ==================
@dp.message(Command("start"))
async def start(msg: types.Message):
    global ADMIN_CHAT_ID
    if msg.from_user.username == ADMIN_USERNAME:
        ADMIN_CHAT_ID = msg.chat.id
        await msg.answer("✅ Ты назначен администратором")

    await msg.answer("📘 Рапортичка группы 101 тп", reply_markup=menu())

# ================== ОТМЕТКА ==================
@dp.message(lambda m: m.text == "📋 Отметить отсутствующих")
async def mark(msg: types.Message):
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=s, callback_data=f"st|{s}")]
            for s in STUDENTS
        ]
    )
    await msg.answer(f"Дата: {today()}", reply_markup=kb)

@dp.callback_query(lambda c: c.data.startswith("st|"))
async def choose_reason(call: types.CallbackQuery):
    student = call.data.split("|", 1)[1]
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=r, callback_data=f"rs|{student}|{r}")]
            for r in REASONS
        ]
    )
    await call.message.answer(student, reply_markup=kb)

@dp.callback_query(lambda c: c.data.startswith("rs|"))
async def save(call: types.CallbackQuery):
    _, student, reason = call.data.split("|", 2)
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
    await call.message.answer("✅ Записано")

# ================== РЕДАКТИРОВАНИЕ ==================
@dp.message(lambda m: m.text == "✏ Редактировать рапортичку")
async def edit(msg: types.Message):
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
                text=f"{r[1]} | {r[2]} | {r[3]}",
                callback_data=f"ed|{r[0]}"
            )] for r in rows
        ]
    )
    await msg.answer("Выбери запись:", reply_markup=kb)

@dp.callback_query(lambda c: c.data.startswith("ed|"))
async def edit_reason(call: types.CallbackQuery):
    rec_id = call.data.split("|")[1]
    kb = InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(
                text=r,
                callback_data=f"upd|{rec_id}|{r}"
            )] for r in REASONS
        ]
    )
    await call.message.answer("Новая причина:", reply_markup=kb)

@dp.callback_query(lambda c: c.data.startswith("upd|"))
async def update(call: types.CallbackQuery):
    _, rec_id, reason = call.data.split("|", 2)
    with db() as con:
        con.execute(
            "UPDATE attendance SET reason=? WHERE id=?",
            (reason, rec_id)
        )
        con.commit()
    await call.message.answer("✏ Обновлено")

# ================== ВЫГРУЗКА ==================
@dp.message(lambda m: m.text == "📤 Выгрузить рапортичку")
async def export(msg: types.Message):
    export_excel()
    await msg.answer_document(FSInputFile(EXCEL_FILE))

# ================== АДМИН ==================
@dp.message(lambda m: m.text == "📨 Отправить админу")
async def send_admin(msg: types.Message):
    if not ADMIN_CHAT_ID:
        await msg.answer("❌ Администратор не активен")
        return
    export_excel()
    await bot.send_document(
        ADMIN_CHAT_ID,
        FSInputFile(EXCEL_FILE),
        caption="📨 Итоговая рапортичка"
    )
    await msg.answer("✅ Отправлено админу")

# ================== ОЧИСТКА ==================
@dp.message(lambda m: m.text == "🗑 Очистить рапортичку")
async def clear(msg: types.Message):
    with db() as con:
        con.execute(
            "UPDATE attendance SET deleted_at=? WHERE deleted_at IS NULL",
            (now(),)
        )
        con.commit()
    await msg.answer("🗑 Очищено (можно восстановить 30 дней)")

# ================== ВОССТАНОВЛЕНИЕ ==================
@dp.message(lambda m: m.text == "♻ Восстановить за месяц")
async def restore(msg: types.Message):
    limit = (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
    with db() as con:
        con.execute("""
        UPDATE attendance
        SET deleted_at = NULL
        WHERE deleted_at IS NOT NULL
        AND deleted_at >= ?
        """, (limit,))
        con.commit()
    await msg.answer("♻ Восстановление выполнено")

# ================== ЗАПУСК ==================
async def main():
    init_db()
    print("Бот запущен")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
