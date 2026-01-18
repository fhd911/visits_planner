from __future__ import annotations

from datetime import date, datetime
from django import template

try:
    from hijri_converter import Gregorian
except Exception:
    Gregorian = None

register = template.Library()


@register.filter(name="to_hijri")
def to_hijri(value):
    """
    يحول تاريخ ميلادي (date/datetime) إلى نص هجري: 1447-07-28
    """
    if not value:
        return ""
    if isinstance(value, datetime):
        value = value.date()
    if not isinstance(value, date):
        return str(value)

    if Gregorian is None:
        # fallback إذا المكتبة غير مثبتة
        return value.strftime("%Y-%m-%d")

    h = Gregorian(value.year, value.month, value.day).to_hijri()
    # تنسيق: YYYY-MM-DD
    return f"{h.year:04d}-{h.month:02d}-{h.day:02d}"
