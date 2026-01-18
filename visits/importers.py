from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from openpyxl import load_workbook
from django.db import transaction

from .models import School, Principal, Supervisor, Assignment


@dataclass
class ImportStats:
    created: int = 0
    updated: int = 0
    skipped: int = 0


def _cell(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _load_rows(file) -> list[dict[str, Any]]:
    wb = load_workbook(file, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [_cell(h) for h in rows[0]]
    out: list[dict[str, Any]] = []
    for r in rows[1:]:
        d = {}
        for i, h in enumerate(headers):
            if not h:
                continue
            d[h] = r[i] if i < len(r) else None
        out.append(d)
    return out


@transaction.atomic
def import_schools(file, gender: str) -> ImportStats:
    """
    المدارس:
    A: اسم المدرسة
    B: الرقم الإحصائي
    C: نوع التعليم
    D: المرحلة
    """
    st = ImportStats()
    rows = _load_rows(file)
    for row in rows:
        name = _cell(row.get("اسم المدرسة"))
        stat_code = _cell(row.get("الرقم الإحصائي"))
        edu = _cell(row.get("نوع التعليم"))
        stage = _cell(row.get("المرحلة"))

        if not stat_code or not name:
            st.skipped += 1
            continue

        obj, created = School.objects.update_or_create(
            stat_code=stat_code,
            defaults={
                "name": name,
                "gender": gender,
                "education_type": edu,
                "stage": stage,
                "is_active": True,
            },
        )
        if created:
            st.created += 1
        else:
            st.updated += 1
    return st


@transaction.atomic
def import_principals(file) -> ImportStats:
    """
    ملف المدراء:
    A: السجل المدني
    B: اسم المدير
    C: رقم الجوال
    D: الرقم الإحصائي
    E: قطاع المدرسة
    """
    st = ImportStats()
    rows = _load_rows(file)
    for row in rows:
        stat_code = _cell(row.get("الرقم الإحصائي"))
        if not stat_code:
            st.skipped += 1
            continue

        school = School.objects.filter(stat_code=stat_code).first()
        if not school:
            st.skipped += 1
            continue

        national_id = _cell(row.get("السجل المدني"))
        full_name = _cell(row.get("اسم المدير"))
        mobile = _cell(row.get("رقم الجوال"))
        sector = _cell(row.get("قطاع المدرسة"))

        if not full_name:
            st.skipped += 1
            continue

        obj, created = Principal.objects.update_or_create(
            school=school,
            defaults={
                "national_id": national_id,
                "full_name": full_name,
                "mobile": mobile,
                "sector": sector,
            },
        )
        if created:
            st.created += 1
        else:
            st.updated += 1
    return st


@transaction.atomic
def import_supervisors(file) -> ImportStats:
    """
    ملف المشرفين (حسب صورتك الأولى):
    A: السجل المدني
    B: اسم المشرف
    C: القسم
    """
    st = ImportStats()
    rows = _load_rows(file)
    for row in rows:
        national_id = _cell(row.get("السجل المدني"))
        full_name = _cell(row.get("اسم المشرف"))
        department = _cell(row.get("القسم"))

        if not national_id or not full_name:
            st.skipped += 1
            continue

        obj, created = Supervisor.objects.update_or_create(
            national_id=national_id,
            defaults={
                "full_name": full_name,
                "department": department,
                "is_active": True,
            },
        )
        if created:
            st.created += 1
        else:
            st.updated += 1
    return st


@transaction.atomic
def import_assignments(file) -> ImportStats:
    """
    ملف الإسناد (حسب صورتك الأولى):
    A: السجل المدني (مشرف)
    D: الرقم الإحصائي (مدرسة)
    """
    st = ImportStats()
    rows = _load_rows(file)
    for row in rows:
        sup_id = _cell(row.get("السجل المدني"))
        stat_code = _cell(row.get("الرقم الإحصائي"))

        if not sup_id or not stat_code:
            st.skipped += 1
            continue

        supervisor = Supervisor.objects.filter(national_id=sup_id).first()
        school = School.objects.filter(stat_code=stat_code).first()

        if not supervisor or not school:
            st.skipped += 1
            continue

        obj, created = Assignment.objects.update_or_create(
            supervisor=supervisor,
            school=school,
            defaults={"is_active": True},
        )
        if created:
            st.created += 1
        else:
            st.updated += 1

    return st
