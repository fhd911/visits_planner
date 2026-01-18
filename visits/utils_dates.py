from __future__ import annotations

from datetime import date, timedelta
from hijridate import Gregorian

AR_WEEKDAYS = ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس"]


def week_start_date(start_week1_sunday: date, week_no: int) -> date:
    return start_week1_sunday + timedelta(days=(week_no - 1) * 7)


def hijri_str(d: date) -> str:
    h = Gregorian.fromdate(d).to_hijri()
    return f"{h.year:04d}/{h.month:02d}/{h.day:02d}"


def week_rows(start_week1_sunday: date, week_no: int):
    start = week_start_date(start_week1_sunday, week_no)
    rows = []
    for i in range(5):
        g = start + timedelta(days=i)
        rows.append(
            {
                "weekday": i,
                "weekday_name": AR_WEEKDAYS[i],
                "greg_date": g,
                "hijri_date": hijri_str(g),
            }
        )
    return rows
