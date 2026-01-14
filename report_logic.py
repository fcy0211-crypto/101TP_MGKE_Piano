import sqlite3
from datetime import datetime, timedelta
from time_service import get_current_date

DB_FILE = ""


def db():
    return sqlite3.connect(DB_FILE)


# ================== МИГРАЦИЯ БД ==================
def migrate():
    with db() as conn:
        c = conn.cursor()
        c.execute("PRAGMA table_info(attendance)")
        cols = [col[1] for col in c.fetchall()]

        if "deleted_at" not in cols:
            c.execute("ALTER TABLE attendance ADD COLUMN deleted_at TEXT")

        if "updated_at" not in cols:
            c.execute("ALTER TABLE attendance ADD COLUMN updated_at TEXT")

        conn.commit()


# ================== ИЗМЕНЕНИЕ ПРИЧИНЫ ==================
def change_reason(attendance_id: int, new_reason: str):
    with db() as conn:
        conn.execute("""
        UPDATE attendance
        SET reason = ?, updated_at = ?
        WHERE id = ? AND deleted_at IS NULL
        """, (new_reason, get_current_date(), attendance_id))
        conn.commit()


# ================== МЯГКАЯ ОЧИСТКА ==================
def soft_clear():
    with db() as conn:
        conn.execute("""
        UPDATE attendance
        SET deleted_at = ?
        WHERE deleted_at IS NULL
        """, (get_current_date(),))
        conn.commit()


# ================== ВОССТАНОВЛЕНИЕ (30 ДНЕЙ) ==================
def restore_last_30_days():
    limit = (datetime.now() - timedelta(days=30)).date().isoformat()
    with db() as conn:
        conn.execute("""
        UPDATE attendance
        SET deleted_at = NULL
        WHERE deleted_at >= ?
        """, (limit,))
        conn.commit()


# ================== ПРОВЕРКА КОНЦА МЕСЯЦА ==================
def is_last_day_of_month() -> bool:
    today = datetime.now().date()
    return (today + timedelta(days=1)).month != today.month


# ================== ДАННЫЕ ДЛЯ ИТОГОВОЙ ВЫГРУЗКИ ==================
def get_month_data(year: int, month: int):
    start = f"{year}-{month:02d}-01"
    end_month = month + 1 if month < 12 else 1
    end_year = year if month < 12 else year + 1
    end = f"{end_year}-{end_month:02d}-01"

    with db() as conn:
        return conn.execute("""
        SELECT a.date, s.full_name, a.status, a.reason, a.author
        FROM attendance a
        JOIN students s ON s.id = a.student_id
        WHERE a.date >= ? AND a.date < ?
        AND a.deleted_at IS NULL
        ORDER BY a.date
        """, (start, end)).fetchall()
