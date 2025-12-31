import sqlite3
from aiogram import F
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, CallbackQuery, Message

DB_FILE = "attendance.db"

REASONS = [
    "по заявлению",
    "по болезни",
    "по неуважительной причине"
]

def db():
    return sqlite3.connect(DB_FILE)

# ===== КНОПКА В ПАНЕЛИ =====
def edit_menu_button():
    from aiogram.types import KeyboardButton
    return KeyboardButton(text="✏ Редактировать рапортичку")

# ===== ВЫБОР ДАТЫ =====
async def edit_choose_date(msg: Message):
    with db() as conn:
        c = conn.cursor()
        c.execute("""
        SELECT DISTINCT date FROM attendance
        WHERE deleted_at IS NULL
        ORDER BY date DESC
        """)
        dates = c.fetchall()

    kb = [
        [InlineKeyboardButton(text=d[0], callback_data=f"edit_date_{d[0]}")]
        for d in dates
    ]

    await msg.answer(
        "📅 Выбери дату для редактирования:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb)
    )

# ===== ВЫБОР СТУДЕНТА =====
async def edit_choose_student(call: CallbackQuery):
    date = call.data.replace("edit_date_", "")

    with db() as conn:
        c = conn.cursor()
        c.execute("""
        SELECT s.id, s.full_name
        FROM students s
        JOIN attendance a ON a.student_id = s.id
        WHERE a.date=? AND a.deleted_at IS NULL
        """, (date,))
        students = c.fetchall()

    kb = [
        [InlineKeyboardButton(
            text=name,
            callback_data=f"edit_student_{date}_{sid}"
        )]
        for sid, name in students
    ]

    await call.message.answer(
        f"👤 Кто отсутствовал {date}?",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb)
    )

# ===== ВЫБОР ДЕЙСТВИЯ =====
async def edit_choose_action(call: CallbackQuery):
    _, date, sid = call.data.split("_", 2)

    kb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(
            text="🔄 Изменить причину",
            callback_data=f"edit_reason_{date}_{sid}"
        )],
        [InlineKeyboardButton(
            text="✅ Сделать присутствующим",
            callback_data=f"edit_present_{date}_{sid}"
        )]
    ])

    await call.message.answer(
        "Что изменить?",
        reply_markup=kb
    )

# ===== ИЗМЕНЕНИЕ ПРИЧИНЫ =====
async def edit_choose_reason(call: CallbackQuery):
    _, date, sid = call.data.split("_", 2)

    kb = [
        [InlineKeyboardButton(
            text=r,
            callback_data=f"edit_reason_set_{date}_{sid}_{r}"
        )] for r in REASONS
    ]

    await call.message.answer(
        "Выбери новую причину:",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=kb)
    )

async def edit_set_reason(call: CallbackQuery):
    _, date, sid, reason = call.data.split("_", 3)

    with db() as conn:
        conn.execute("""
        UPDATE attendance
        SET reason=?, updated_at=datetime('now')
        WHERE date=? AND student_id=? AND deleted_at IS NULL
        """, (reason, date, sid))
        conn.commit()

    await call.message.answer("✏ Причина обновлена")

# ===== СДЕЛАТЬ ПРИСУТСТВУЮЩИМ =====
async def edit_set_present(call: CallbackQuery):
    _, date, sid = call.data.split("_", 2)

    with db() as conn:
        conn.execute("""
        DELETE FROM attendance
        WHERE date=? AND student_id=? AND deleted_at IS NULL
        """, (date, sid))
        conn.commit()

    await call.message.answer("✅ Отметка удалена (присутствовал)")
