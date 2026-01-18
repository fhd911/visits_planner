# visits/views_import.py  (كود كامل: صفحة الاستيراد + معالجة Excel)
# ملاحظة: يتطلب openpyxl

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional

from django.contrib import messages
from django.db import transaction
from django.http import HttpRequest, HttpResponse
from django.shortcuts import render

from openpyxl import load_workbook

from .forms import ImportExcelForm
from .models import Assignment, Principal, School, Supervisor


@dataclass
class ImportStats:
    created: int = 0
    updated: int = 0
    skipped: int = 0


def _norm(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _to_bool(v) -> bool:
    s = _norm(v).lower()
    if s in {"1", "true", "yes", "y", "نعم"}:
        return True
    if s in {"0", "false", "no", "n", "لا"}:
        return False
    return True  # افتراضي


def _sheet_rows(file) -> list[dict]:
    wb = load_workbook(filename=file, data_only=True)
    ws = wb.active
    headers = []
    out = []
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if i == 1:
            headers = [_norm(x) for x in row]
            continue
        if not any(x is not None and _norm(x) != "" for x in row):
            continue
        rec = {headers[j]: row[j] for j in range(len(headers))}
        out.append(rec)
    return out


def _import_schools(file, gender: str) -> ImportStats:
    """
    الأعمدة المتوقعة (بالاسم):
      stat_code | name | education_type | stage | is_active (اختياري)
    """
    st = ImportStats()
    rows = _sheet_rows(file)

    for r in rows:
        stat_code = _norm(r.get("stat_code"))
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
        school_stat_code = _norm(r.get("school_stat_code"))
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
            "national_id": _norm(r.get("national_id")),
            "mobile": _norm(r.get("mobile")),
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
        national_id = _norm(r.get("national_id"))
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
    الأعمدة المتوقعة:
      supervisor_national_id | school_stat_code | is_active (اختياري)
    """
    st = ImportStats()
    rows = _sheet_rows(file)

    for r in rows:
        sup_nid = _norm(r.get("supervisor_national_id"))
        school_stat_code = _norm(r.get("school_stat_code"))
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
            supervisor=supervisor, school=school, defaults=defaults
        )
        if created:
            st.created += 1
        else:
            st.updated += 1

    return st


def manager_import_view(request: HttpRequest) -> HttpResponse:
    results = {}

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
