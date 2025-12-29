from datetime import datetime, timezone, timedelta

# === НАСТРОЙКА ЧАСОВОГО ПОЯСА ===
# Москва = UTC+3
TZ = timezone(timedelta(hours=3))


def get_current_date() -> str:
    """
    Возвращает актуальную дату по реальному времени сервера
    с учётом часового пояса.
    Формат: YYYY-MM-DD
    """
    return datetime.now(TZ).date().isoformat()
