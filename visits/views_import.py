# visits/views_import.py  (FINAL ✅)
# صفحة الاستيراد + معالجة Excel بشكل ذكي
# Requires: openpyxl

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from django.contrib import messages
from django.db import transaction
from django.http import HttpRequest, HttpResponse
from django.shortcuts import render

from openpyxl import load_workbook

from .forms import ImportExcelForm
from .models import Assignment, Principal, School, Supervisor


# ============================================================================
# Stats
# ============================================================================
@dataclass
class ImportStats:
    created: int = 0
    updated: int = 0
    skipped: int = 0


# ============================================================================
# Helpers
# ============================================================================
def _norm(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _digits(v: Any) -> str:
    """
    يرجع الأرقام فقط من أي قيمة (مفيد للهوية / الجوال):
    - 1020103717 -> 1020103717
    - "1020103717 " -> 1020103717
    - 70228.0 -> 70228
    - "70228-1" -> 702281
    """
    s = _norm(v)
    if not s:
        return ""
    s = s.replace(".0", "").strip()
    return re.sub(r"\D+", "", s)


def _code(v: Any) -> str:
    """
    ✅ كود المدرسة / الرقم الإحصائي:
    - يحافظ على الحروف مثل M3964353
    - يحول 70228.0 إلى 70228
    - يحذف الفراغات فقط
    """
    s = _norm(v)
    if not s:
        return ""
    s = s.replace(".0", "").strip()
    s = s.replace(" ", "")
    return s


def _to_bool(v: Any) -> bool:
    s = _norm(v).lower()
    if s in {"1", "true", "yes", "y", "نعم"}:
        return True
    if s in {"0", "false", "no", "n", "لا"}:
        return False
    return True  # default


def _canon_header(h: str) -> str:
    """
    تحويل الهيدر (عربي/إنجليزي) إلى مفاتيح قياسية موحدة.
    """
    x = _norm(h).lower()
    x = x.replace("ـ", "").replace("_", " ").strip()

    # ---------- Schools ----------
    if x in {"stat_code", "stat code", "الرقم الإحصائي", "الرقم الاحصائي", "رقم احصائي"}:
        return "stat_code"
    if x in {"name", "اسم المدرسة", "المدرسة"}:
        return "name"
    if x in {"education_type", "education type", "type", "نوع التعليم"}:
        return "education_type"
    if x in {"stage", "المرحلة"}:
        return "stage"
    if x in {"gender", "الجنس"}:
        return "gender"
    if x in {"is_active", "active", "نشط"}:
        return "is_active"

    # ---------- Principals ----------
    if x in {"school_stat_code", "school stat code", "رقم المدرسة", "رقم احصائي المدرسة", "الرقم الإحصائي"}:
        return "school_stat_code"
    if x in {"full_name", "full name", "الاسم", "اسم القائد", "اسم القائدة", "اسم المدير", "اسم المديرة"}:
        return "full_name"
    if x in {"national_id", "national id", "السجل المدني", "رقم الهوية", "الهوية"}:
        return "national_id"
    if x in {"mobile", "الجوال", "رقم الجوال"}:
        return "mobile"
    if x in {"sector", "القطاع"}:
        return "sector"
    if x in {"department", "القسم", "الإدارة"}:
        return "department"

    # ---------- Supervisors ----------
    if x in {"supervisor_national_id", "supervisor national id", "رقم هوية المشرف"}:
        return "supervisor_national_id"
    if x in {"supervisor_name", "supervisor name", "اسم المشرف", "المشرف"}:
        return "supervisor_name"

    # ---------- Assignments ----------
    if x in {"school", "school_stat_code", "school stat code", "الرقم الإحصائي", "الرقم الاحصائي"}:
        return "school_stat_code"
    if x in {"supervisor", "supervisor_national_id", "السجل المدني", "رقم الهوية", "الهوية"}:
        return "supervisor_national_id"

    return _norm(h)


def _sheet_rows(file) -> list[dict]:
    """
    يقرأ الشيت ويعيد list[dict] بحيث مفاتيح الأعمدة تكون:
    - مفاتيح أصلية + مفاتيح قياسية (canonical) لنفس القيم
    """
    wb = load_workbook(filename=file, data_only=True)
    ws = wb.active

    out: list[dict] = []
    headers_raw: list[str] = []

    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if i == 1:
            headers_raw = [_norm(x) for x in row]
            continue

        if not any(x is not None and _norm(x) != "" for x in row):
            continue

        rec: dict = {}

        # ضع القيم بمفاتيحها الأصلية
        for j in range(len(headers_raw)):
            key = headers_raw[j] if j < len(headers_raw) else f"col_{j}"
            rec[key] = row[j] if j < len(row) else None

        # ثم أضف مفاتيح canonical لنفس القيم
        for j in range(len(headers_raw)):
            canon = _canon_header(headers_raw[j])
            if canon and canon not in rec:
                rec[canon] = row[j] if j < len(row) else None

        out.append(rec)

    return out


# ============================================================================
# Importers
# ============================================================================
def _import_schools(file, gender: str) -> ImportStats:
    """
    الأعمدة المتوقعة (بالاسم):
      stat_code | name | education_type | stage | is_active (اختياري)
    """
    st = ImportStats()
    rows = _sheet_rows(file)

    for r in rows:
        stat_code = _code(r.get("stat_code"))  # ✅ يحافظ على M...
        name = _norm(r.get("name"))

        if not stat_code or not name:
            st.skipped += 1
            continue

        defaults = {
            "name": name,
            "gender": gender,
            "education_type": _norm(r.get("education_type")),
            "stage": _norm(r.get("stage")),
            "is_active": _to_bool(r.get("is_active")) if "is_active" in r else True,
        }

        obj, created = School.objects.update_or_create(stat_code=stat_code, defaults=defaults)
        if created:
            st.created += 1
        else:
            st.updated += 1

    return st


def _import_principals(file) -> ImportStats:
    """
    الأعمدة المتوقعة:
      school_stat_code | full_name | national_id (اختياري) | mobile (اختياري) | sector (اختياري)
    """
    st = ImportStats()
    rows = _sheet_rows(file)

    for r in rows:
        school_stat_code = _code(r.get("school_stat_code"))  # ✅ قد يكون M...
        full_name = _norm(r.get("full_name"))

        if not school_stat_code or not full_name:
            st.skipped += 1
            continue

        school = School.objects.filter(stat_code=school_stat_code).first()
        if not school:
            st.skipped += 1
            continue

        defaults = {
            "full_name": full_name,
            "national_id": _digits(r.get("national_id")),
            "mobile": _digits(r.get("mobile")),
            "sector": _norm(r.get("sector")),
        }

        obj, created = Principal.objects.update_or_create(school=school, defaults=defaults)
        if created:
            st.created += 1
        else:
            st.updated += 1

    return st


def _import_supervisors(file) -> ImportStats:
    """
    الأعمدة المتوقعة:
      national_id | full_name | department (اختياري) | is_active (اختياري)
    """
    st = ImportStats()
    rows = _sheet_rows(file)

    for r in rows:
        national_id = _digits(r.get("national_id"))  # ✅ أرقام فقط
        full_name = _norm(r.get("full_name"))

        if not national_id or not full_name:
            st.skipped += 1
            continue

        defaults = {
            "full_name": full_name,
            "department": _norm(r.get("department")),
            "is_active": _to_bool(r.get("is_active")) if "is_active" in r else True,
        }

        obj, created = Supervisor.objects.update_or_create(national_id=national_id, defaults=defaults)
        if created:
            st.created += 1
        else:
            st.updated += 1

    return st


def _import_assignments(file) -> ImportStats:
    """
    ✅ يدعم ملف الإسنادات بأي من هذي الصيغ:

    1) القياسية:
      supervisor_national_id | school_stat_code | is_active (اختياري)

    2) العربية الشائعة:
      السجل المدني | اسم المشرف | الرقم الإحصائي

    ✅ ملاحظة مهمة:
    - الهوية = أرقام فقط
    - الرقم الإحصائي للمدرسة = قد يحتوي حرف M لذلك نستخدم _code وليس _digits
    """
    st = ImportStats()
    rows = _sheet_rows(file)

    for r in rows:
        # ✅ 1) هوية المشرف: نأخذها من أكثر من مكان
        sup_nid = _digits(r.get("supervisor_national_id")) or _digits(r.get("national_id"))

        # لو ملفك ملخبط: الهوية داخل "اسم المشرف"
        if not sup_nid:
            sup_nid = (
                _digits(r.get("supervisor_name"))
                or _digits(r.get("اسم المشرف"))
                or _digits(r.get("المشرف"))
            )

        # ✅ 2) الرقم الإحصائي للمدرسة: لا نستخدم digits
        school_stat_code = (
            _code(r.get("school_stat_code"))
            or _code(r.get("stat_code"))
            or _code(r.get("الرقم الإحصائي"))
            or _code(r.get("الرقم الاحصائي"))
        )

        if not sup_nid or not school_stat_code:
            st.skipped += 1
            continue

        supervisor = Supervisor.objects.filter(national_id=sup_nid).first()
        school = School.objects.filter(stat_code=school_stat_code).first()

        if not supervisor or not school:
            st.skipped += 1
            continue

        defaults = {
            "is_active": _to_bool(r.get("is_active")) if "is_active" in r else True,
        }

        obj, created = Assignment.objects.update_or_create(
            supervisor=supervisor,
            school=school,
            defaults=defaults,
        )

        if created:
            st.created += 1
        else:
            st.updated += 1

    return st


# ============================================================================
# View
# ============================================================================
def manager_import_view(request: HttpRequest) -> HttpResponse:
    results: dict[str, ImportStats] = {}

    if request.method == "POST":
        form = ImportExcelForm(request.POST, request.FILES)
        if form.is_valid():
            with transaction.atomic():
                if form.cleaned_data.get("schools_boys"):
                    results["المدارس (بنين)"] = _import_schools(form.cleaned_data["schools_boys"], "boys")

                if form.cleaned_data.get("schools_girls"):
                    results["المدارس (بنات)"] = _import_schools(form.cleaned_data["schools_girls"], "girls")

                if form.cleaned_data.get("principals"):
                    results["مديرو المدارس"] = _import_principals(form.cleaned_data["principals"])

                if form.cleaned_data.get("supervisors"):
                    results["المشرفون"] = _import_supervisors(form.cleaned_data["supervisors"])

                if form.cleaned_data.get("assignments"):
                    results["الإسنادات"] = _import_assignments(form.cleaned_data["assignments"])

            messages.success(request, "تمت عملية الاستيراد بنجاح.")
        else:
            messages.error(request, "تحقق من الملفات المرفوعة (ارفع ملفًا واحدًا على الأقل).")
    else:
        form = ImportExcelForm()

    return render(request, "visits/manager_import.html", {"form": form, "results": results})
