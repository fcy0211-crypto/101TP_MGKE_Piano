import asyncio
import sqlite3
from datetime import datetime

from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command
from aiogram.types import (
    Message, CallbackQuery,
    InlineKeyboardMarkup, InlineKeyboardButton,
    ReplyKeyboardMarkup, KeyboardButton,
    FSInputFile
)

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo

# ===== ПОДКЛЮЧЕНИЕ МОДУЛЯ РЕДАКТИРОВАНИЯ =====
from edit_attendance import (
    edit_choose_date,
    edit_choose_student,
    edit_choose_action,
    edit_choose_reason,
    edit_set_reason,
    edit_set_present
)

# ================== НАСТРОЙКИ ==================
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

# ================== БАЗА ==================
def db():
    return sqlite3.connect(DB_FILE)


def init_db():
    with db() as conn:
        c = conn.cursor()

        c.execute("""
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            full_name TEXT UNIQUE
        )
        """)

        c.execute("""
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            student_id INTEGER,
            status TEXT,
            reason TEXT,
            author TEXT,
            deleted_at TEXT,
            updated_at TEXT
        )
        """)

        for s in STUDENTS:
            c.execute(
                "INSERT OR IGNORE INTO students (full_name) VALUES (?)",
                (s,)
            )
        conn.commit()

# ================== ДАТА ==================
def current_date():
    return datetime.now().strftime("%Y-%m-%d")

# ================== EXCEL ==================
def update_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Рапортичка"

    headers = ["Дата", "ФИО", "Статус", "Причина", "Кто отметил"]
    ws.append(headers)

    for i in range(1, 6):
        ws.cell(row=1, column=i).font = Font(bold=True)
        ws.cell(row=1, column=i).alignment = Alignment(horizontal="center")

    with db() as conn:
        c = conn.cursor()

        c.execute("""
        SELECT DISTINCT date FROM attendance
        WHERE deleted_at IS NULL
        ORDER BY date
        """)
        dates = [d[0] for d in c.fetchall()]

        c.execute("SELECT id, full_name FROM students")
        students = c.fetchall()

        for d in dates:
            for sid, name in students:
                c.execute("""
                SELECT status, reason, author
                FROM attendance
                WHERE date=? AND student_id=? AND deleted_at IS NULL
                """, (d, sid))
                row = c.fetchone()

                if row:
                    status, reason, author = row
                else:
                    status, reason, author = "присутствовал", "", ""

                ws.append([d, name, status, reason, author])

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 28

    table = Table(displayName="Attendance", ref=f"A1:E{ws.max_row}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showRowStripes=True
    )
    ws.add_table(table)

    wb.save(EXCEL_FILE)

# ================== КЛАВИАТУРА ==================
def main_menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📋 Отметить отсутствующих")],
            [KeyboardButton(text="✏ Редактировать рапортичку")],
            [KeyboardButton(text="📤 Выгрузить рапортичку")],
            [KeyboardButton(text="📨 Отправить рапортичку админу")],
        ],
        resize_keyboard=True
    )

# ================== START ==================
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

# ================== ОТМЕТКА ==================
@dp.message(F.text == "📋 Отметить отсутствующих")
async def mark(msg: Message):
    kb = []
    with db() as conn:
        c = conn.cursor()
        c.execute("SELECT id, full_name FROM students")
        for sid, name in c.fetchall():
            kb.append([InlineKeyboardButton(
                text=name,
                callback_data=f"student_{sid}"
            )])

    await msg.answer(
        f"📅 Дата: {current_date()}",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb)
    )

@dp.callback_query(F.data.startswith("student_"))
async def choose_reason(call: CallbackQuery):
    sid = call.data.split("_")[1]
    kb = [[InlineKeyboardButton(
        text=r,
        callback_data=f"reason_{sid}_{r}"
    )] for r in REASONS]

    await call.message.answer(
        "Причина отсутствия:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb)
    )

@dp.callback_query(F.data.startswith("reason_"))
async def save(call: CallbackQuery):
    _, sid, reason = call.data.split("_", 2)

    with db() as conn:
        conn.execute("""
        INSERT INTO attendance (date, student_id, status, reason, author)
        VALUES (?, ?, 'отсутствовал', ?, ?)
        """, (
            current_date(),
            sid,
            reason,
            call.from_user.username or call.from_user.full_name
        ))
        conn.commit()

    update_excel()
    await call.message.answer("✅ Отмечено")

# ================== РЕДАКТИРОВАНИЕ ==================
@dp.message(F.text == "✏ Редактировать рапортичку")
async def edit(msg: Message):
    await edit_choose_date(msg)

@dp.callback_query(F.data.startswith("edit_date_"))
async def cb_edit_date(call: CallbackQuery):
    await edit_choose_student(call)

@dp.callback_query(F.data.startswith("edit_student_"))
async def cb_edit_student(call: CallbackQuery):
    await edit_choose_action(call)

@dp.callback_query(F.data.startswith("edit_reason_") and not F.data.startswith("edit_reason_set"))
async def cb_edit_reason(call: CallbackQuery):
    await edit_choose_reason(call)

@dp.callback_query(F.data.startswith("edit_reason_set_"))
async def cb_edit_reason_set(call: CallbackQuery):
    await edit_set_reason(call)
    update_excel()

@dp.callback_query(F.data.startswith("edit_present_"))
async def cb_edit_present(call: CallbackQuery):
    await edit_set_present(call)
    update_excel()

# ================== ВЫГРУЗКА ==================
@dp.message(F.text == "📤 Выгрузить рапортичку")
async def export(msg: Message):
    update_excel()
    await msg.answer_document(
        FSInputFile(EXCEL_FILE),
        caption="📤 Общая рапортичка"
    )

# ================== АДМИН ==================
@dp.message(F.text == "📨 Отправить рапортичку админу")
async def send_admin(msg: Message):
    if not ADMIN_CHAT_ID:
        await msg.answer("❌ Администратор не авторизовался (/start)")
        return

    update_excel()
    await bot.send_document(
        ADMIN_CHAT_ID,
        FSInputFile(EXCEL_FILE),
        caption="📨 Рапортичка"
    )
    await msg.answer("✅ Отправлено")

# ================== ЗАПУСК ==================
async def main():
    init_db()
    print("Бот запущен")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
