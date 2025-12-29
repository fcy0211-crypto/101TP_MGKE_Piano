import asyncio
import sqlite3

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

from time_service import get_current_date  # ⬅️ ВАЖНО

# ================== НАСТРОЙКИ ==================
BOT_TOKEN = "8397597216:AAFtzivDMoNxcRU06vp8wobfG6NU28BkIgs"
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

# ================== БАЗА ДАННЫХ ==================
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
            author TEXT
        )
        """)

        for s in STUDENTS:
            c.execute(
                "INSERT OR IGNORE INTO students (full_name) VALUES (?)",
                (s,)
            )

        conn.commit()

# ================== EXCEL ==================
def update_excel_file():
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

        c.execute("SELECT DISTINCT date FROM attendance ORDER BY date")
        dates = [d[0] for d in c.fetchall()]

        c.execute("SELECT id, full_name FROM students")
        students = c.fetchall()

        for d in dates:
            for sid, name in students:
                c.execute("""
                SELECT status, reason, author
                FROM attendance
                WHERE date=? AND student_id=?
                """, (d, sid))
                row = c.fetchone()

                if row:
                    status, reason, author = row
                else:
                    status, reason, author = "присутствовал", "", ""

                ws.append([d, name, status, reason, author])

    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

    table = Table(displayName="Attendance", ref=f"A1:E{ws.max_row}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showRowStripes=True
    )
    ws.add_table(table)

    wb.save(EXCEL_FILE)

# ================== КЛАВИАТУРЫ ==================
def main_menu():
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📋 Отметить отсутствующих")],
            [KeyboardButton(text="📤 Выгрузить рапортичку")],
            [KeyboardButton(text="🗑 Очистить рапортичку")]
        ],
        resize_keyboard=True
    )

# ================== ХЕНДЛЕРЫ ==================
@dp.message(Command("start"))
async def start(msg: Message):
    await msg.answer(
        "📘 Рапортичка группы 101 тп",
        reply_markup=main_menu()
    )

# -------- ОТМЕТИТЬ ОТСУТСТВУЮЩИХ --------
@dp.message(F.text == "📋 Отметить отсутствующих")
async def mark_menu(msg: Message):
    kb = []
    with db() as conn:
        c = conn.cursor()
        c.execute("SELECT id, full_name FROM students")
        for sid, name in c.fetchall():
            kb.append([
                InlineKeyboardButton(
                    text=name,
                    callback_data=f"student_{sid}"
                )
            ])

    await msg.answer(
        f"📅 Дата: {get_current_date()}",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb)
    )

@dp.callback_query(F.data.startswith("student_"))
async def choose_reason(call: CallbackQuery):
    sid = call.data.split("_")[1]
    kb = [
        [InlineKeyboardButton(
            text=r,
            callback_data=f"reason_{sid}_{r}"
        )] for r in REASONS
    ]
    await call.message.answer(
        "Укажи причину отсутствия:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb)
    )

@dp.callback_query(F.data.startswith("reason_"))
async def save_attendance(call: CallbackQuery):
    _, sid, reason = call.data.split("_", 2)

    with db() as conn:
        c = conn.cursor()
        c.execute("""
        INSERT INTO attendance (date, student_id, status, reason, author)
        VALUES (?, ?, ?, ?, ?)
        """, (
            get_current_date(),
            sid,
            "отсутствовал",
            reason,
            call.from_user.username or call.from_user.full_name
        ))
        conn.commit()

    update_excel_file()
    await call.message.answer("✅ Отмечено")

# -------- ВЫГРУЗКА --------
@dp.message(F.text == "📤 Выгрузить рапортичку")
async def export_menu(msg: Message):
    update_excel_file()
    await msg.answer_document(
        FSInputFile(EXCEL_FILE),
        caption="📤 Общая синхронизированная рапортичка группы 101 тп"
    )

# -------- ОЧИСТКА --------
@dp.message(F.text == "🗑 Очистить рапортичку")
async def clear_menu(msg: Message):
    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="❌ Нет", callback_data="clear_no")],
        [InlineKeyboardButton(text="✅ Да", callback_data="clear_yes")]
    ])
    await msg.answer(
        "⚠ Очистить ВСЮ рапортичку?",
        reply_markup=kb
    )

@dp.callback_query(F.data == "clear_yes")
async def confirm_clear(call: CallbackQuery):
    with db() as conn:
        conn.execute("DELETE FROM attendance")
        conn.commit()

    update_excel_file()
    await call.message.answer("🗑 Рапортичка очищена")

@dp.callback_query(F.data == "clear_no")
async def cancel_clear(call: CallbackQuery):
    await call.message.answer("Отмена")

# ================== ЗАПУСК ==================
async def main():
    init_db()
    while True:
        try:
            print("🤖 Бот запущен")
            await dp.start_polling(bot)
        except Exception as e:
            print("Ошибка:", e)
            await asyncio.sleep(5)

if __name__ == "__main__":
    asyncio.run(main())
