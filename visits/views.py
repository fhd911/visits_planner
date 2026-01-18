# visits/views.py
from __future__ import annotations

from io import BytesIO
from datetime import date, timedelta
from typing import Optional

from django.contrib import messages
from django.contrib.admin.views.decorators import staff_member_required
from django.core.paginator import Paginator
from django.db import transaction
from django.db.models import Q
from django.http import HttpRequest, HttpResponse, JsonResponse
from django.shortcuts import render, redirect, get_object_or_404
from django.urls import reverse
from django.utils import timezone
from django.views.decorators.http import require_POST

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from .forms import ImportExcelForm
from .models import (
    Assignment,
    Plan,
    PlanDay,
    PlanWeek,
    School,
    Supervisor,
    UnlockRequest,
    Principal,
)

# =====================
# ثوابت
# =====================
WEEKDAYS = [
    (0, "الأحد"),
    (1, "الإثنين"),
    (2, "الثلاثاء"),
    (3, "الأربعاء"),
    (4, "الخميس"),
]

WEEKDAY_MAP = dict(WEEKDAYS)

SESSION_SUP_ID = "visits_sup_id"


# =====================
# Helpers عامة
# =====================
def _digits(s: str) -> str:
    """يرجع الأرقام فقط من أي نص"""
    return "".join(ch for ch in (s or "") if ch.isdigit())


def _safe_int(v: object, default: int = 1) -> int:
    try:
        return int(str(v).strip())
    except Exception:
        return default


def _is_ajax(request: HttpRequest) -> bool:
    return request.headers.get("X-Requested-With") == "XMLHttpRequest"


def _get_mobile_digits(sup: Supervisor) -> str:
    return _digits(getattr(sup, "mobile", "") or "")


def _supervisor_last4(sup: Supervisor) -> Optional[str]:
    d = _get_mobile_digits(sup)
    return d[-4:] if len(d) >= 4 else None


def _get_logged_in_supervisor(request: HttpRequest) -> Optional[Supervisor]:
    sid = request.session.get(SESSION_SUP_ID)
    if not sid:
        return None
    try:
        sid_int = int(sid)
    except Exception:
        return None
    return Supervisor.objects.filter(id=sid_int, is_active=True).first()


def _require_supervisor(request: HttpRequest) -> Supervisor:
    sup = _get_logged_in_supervisor(request)
    if not sup:
        raise Supervisor.DoesNotExist
    return sup


def _find_supervisor_by_nid(nid: str) -> Optional[Supervisor]:
    """
    يبحث بالمشرف عبر national_id (أساسي)
    ويدعم civil_registry لو كان موجود في الموديل.
    """
    nid = _digits(nid)
    if not nid:
        return None

    qs = Supervisor.objects.filter(is_active=True)
    query = Q()

    try:
        Supervisor._meta.get_field("national_id")
        query |= Q(national_id=nid)
    except Exception:
        pass

    try:
        Supervisor._meta.get_field("civil_registry")
        query |= Q(civil_registry=nid)
    except Exception:
        pass

    if query == Q():
        return None

    return qs.filter(query).first()


def _sup_nid_value(sup: Supervisor) -> str:
    """إرجاع الهوية سواء national_id أو civil_registry"""
    try:
        Supervisor._meta.get_field("national_id")
        v = getattr(sup, "national_id", "") or ""
        if v:
            return v
    except Exception:
        pass
    return getattr(sup, "civil_registry", "") or ""


def _inject_week_no(plan: Plan) -> None:
    """
    ✅ حل سريع وعملي:
    بعض القوالب عندك تستخدم plan.week_no مباشرة
    بينما في الموديل غالباً week = FK (PlanWeek)
    هنا نضيف week_no بشكل مؤقت للكائن حتى القوالب تشتغل بدون تغيير الموديل.
    """
    try:
        plan.week_no = plan.week.week_no  # type: ignore[attr-defined]
    except Exception:
        plan.week_no = None  # type: ignore[attr-defined]


# ==========================================================
# ✅ PlanWeek helpers (يدوي بالكامل)
# ==========================================================
def _get_active_weeks_qs():
    """الأسابيع الفعّالة للمشرف = ليست إجازة"""
    return PlanWeek.objects.filter(is_break=False).order_by("week_no")


def _get_all_weeks_qs():
    return PlanWeek.objects.all().order_by("week_no")


def _get_default_week_no() -> int:
    w = _get_active_weeks_qs().first()
    return w.week_no if w else 1


def _week_exists(week_no: int) -> bool:
    return PlanWeek.objects.filter(week_no=week_no).exists()


def _resolve_week_or_404(week_no: int, *, allow_inactive: bool = False) -> PlanWeek:
    qs = PlanWeek.objects.all()
    if not allow_inactive:
        qs = qs.filter(is_break=False)
    return get_object_or_404(qs, week_no=week_no)


def _build_week_choices(active_only: bool = True) -> list[tuple[int, str]]:
    qs = _get_active_weeks_qs() if active_only else _get_all_weeks_qs()
    out: list[tuple[int, str]] = []
    for w in qs:
        label = f"الأسبوع {w.week_no}"
        if getattr(w, "title", ""):
            label += f" — {w.title}"
        if getattr(w, "is_break", False):
            label += " (إجازة)"
        out.append((w.week_no, label))
    return out


def _build_day_dates_from_week(week_obj: PlanWeek) -> dict[int, date]:
    """تواريخ الأحد..الخميس بناءً على start_sunday المدخلة يدويًا"""
    start = week_obj.start_sunday
    return {
        0: start,
        1: start + timedelta(days=1),
        2: start + timedelta(days=2),
        3: start + timedelta(days=3),
        4: start + timedelta(days=4),
    }


def _plan_url(week_no: int) -> str:
    return f"{reverse('visits:plan')}?week={week_no}"


def _admin_dashboard_url(week_no: int, *, show_all: bool = False) -> str:
    url = f"{reverse('visits:admin_dashboard')}?week={week_no}"
    if show_all:
        url += "&all=1"
    return url


# =====================
# Status helpers
# =====================
def _status_label(plan: Plan) -> str:
    if plan.status == Plan.STATUS_APPROVED:
        return "معتمدة"
    if plan.status == Plan.STATUS_UNLOCK_REQUESTED:
        return "طلب فك اعتماد"
    return "مسودة"


def _status_code(plan: Plan) -> str:
    if plan.status == Plan.STATUS_APPROVED:
        return "approved"
    if plan.status == Plan.STATUS_UNLOCK_REQUESTED:
        return "unlock"
    return "draft"


def _status_css(plan: Plan) -> str:
    return _status_code(plan)


def _plan_filled_count(plan: Plan) -> int:
    day_map = {d.weekday: d for d in plan.days.all()}
    return sum(1 for wd, _ in WEEKDAYS if day_map.get(wd) and day_map[wd].school_id)


# =====================
# Login / Logout
# =====================
def login_view(request: HttpRequest) -> HttpResponse:
    """
    يقبل حقول الفورم بأي من هذه الأسماء:
    - رقم الهوية: nid / national_id / civil_registry
    - آخر4: last4 / phone_last4
    """
    if request.method == "POST":
        nid = _digits(
            (
                request.POST.get("nid")
                or request.POST.get("national_id")
                or request.POST.get("civil_registry")
                or ""
            ).strip()
        )
        last4 = _digits((request.POST.get("last4") or request.POST.get("phone_last4") or "").strip())

        if len(nid) != 10:
            messages.error(request, "فضلاً أدخل رقم الهوية بشكل صحيح.")
            return render(request, "visits/login.html")

        if len(last4) != 4:
            messages.error(request, "فضلاً أدخل آخر 4 أرقام من الجوال (4 أرقام).")
            return render(request, "visits/login.html")

        sup = _find_supervisor_by_nid(nid)
        if not sup:
            messages.error(request, "المشرف غير موجود أو غير مفعل.")
            return render(request, "visits/login.html")

        sup_last4 = _supervisor_last4(sup)
        if not sup_last4:
            messages.error(request, "لا يمكن التحقق لأن رقم جوال المشرف غير محفوظ. راجع الإدارة لإضافته.")
            return render(request, "visits/login.html")

        if sup_last4 != last4:
            messages.error(request, "بيانات التحقق غير صحيحة (آخر 4 أرقام من الجوال).")
            return render(request, "visits/login.html")

        request.session[SESSION_SUP_ID] = sup.id
        request.session.modified = True
        return redirect(_plan_url(_get_default_week_no()))

    return render(request, "visits/login.html")


def logout_view(request: HttpRequest) -> HttpResponse:
    request.session.pop(SESSION_SUP_ID, None)
    messages.success(request, "تم تسجيل الخروج.")
    return redirect("visits:login")


# =====================
# Plan (Supervisor)
# =====================
def plan_view(request: HttpRequest) -> HttpResponse:
    try:
        supervisor = _require_supervisor(request)
    except Supervisor.DoesNotExist:
        return redirect("visits:login")

    week_no = _safe_int(
        request.GET.get("week") or request.POST.get("week") or _get_default_week_no(),
        default=_get_default_week_no(),
    )
    week_obj = _resolve_week_or_404(week_no, allow_inactive=False)

    plan, _ = Plan.objects.get_or_create(supervisor=supervisor, week=week_obj)
    _inject_week_no(plan)

    schools = School.objects.filter(
        assignments__supervisor=supervisor,
        assignments__is_active=True,
        is_active=True,
    ).order_by("name")

    days_map = {d.weekday: d for d in plan.days.all().select_related("school")}
    week_choices = _build_week_choices(active_only=True)
    day_dates = _build_day_dates_from_week(week_obj)

    if request.method == "POST":
        if plan.status in (Plan.STATUS_APPROVED, Plan.STATUS_UNLOCK_REQUESTED):
            messages.info(
                request,
                "لا يمكن تعديل الخطة الآن. إذا كانت معتمدة اطلب فك الاعتماد أولاً.",
            )
            return redirect(_plan_url(week_obj.week_no))

        action = (request.POST.get("action") or "save").strip()  # save | approve

        for wd, _ in WEEKDAYS:
            sid = (request.POST.get(f"school_{wd}") or "").strip()
            vtype = (request.POST.get(f"visit_{wd}") or "in").strip()
            if vtype not in ("in", "remote"):
                vtype = "in"

            if sid:
                PlanDay.objects.update_or_create(
                    plan=plan,
                    weekday=wd,
                    defaults={"school_id": sid, "visit_type": vtype},
                )
            else:
                PlanDay.objects.filter(plan=plan, weekday=wd).delete()

        plan.saved_at = timezone.now()

        if action == "approve":
            if plan.is_fully_filled():
                plan.status = Plan.STATUS_APPROVED
                plan.approved_at = timezone.now()
                plan.save(update_fields=["saved_at", "status", "approved_at"])
                messages.success(request, "تم اعتماد الخطة بنجاح ✅")
            else:
                plan.save(update_fields=["saved_at"])
                messages.warning(request, "تم الحفظ، لكن لا يمكن الاعتماد قبل اكتمال جميع الأيام.")
        else:
            plan.save(update_fields=["saved_at"])
            messages.success(request, "تم حفظ الخطة ✅")

        return redirect(_plan_url(week_obj.week_no))

    return render(
        request,
        "visits/plan.html",
        {
            "plan": plan,
            "schools": schools,
            "week": week_obj.week_no,
            "week_obj": week_obj,
            "sup": supervisor,
            "weekdays": WEEKDAYS,
            "days_map": days_map,
            "week_choices": week_choices,
            "day_dates": day_dates,
            "today": date.today(),
        },
    )


# =====================
# Export: supervisor plan (Excel)
# =====================
def export_plan_excel(request: HttpRequest) -> HttpResponse:
    try:
        supervisor = _require_supervisor(request)
    except Supervisor.DoesNotExist:
        return redirect("visits:login")

    week_no = _safe_int(request.GET.get("week") or _get_default_week_no(), default=_get_default_week_no())
    week_obj = _resolve_week_or_404(week_no, allow_inactive=False)

    plan = get_object_or_404(Plan, supervisor=supervisor, week=week_obj)
    _inject_week_no(plan)

    wb = Workbook()
    ws = wb.active
    ws.title = f"الأسبوع {week_obj.week_no}"
    ws.sheet_view.rightToLeft = True

    title_font = Font(name="Cairo", bold=True, size=14)
    bold_font = Font(name="Cairo", bold=True, size=12)
    normal_font = Font(name="Cairo", size=12)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    header_fill = PatternFill("solid", fgColor="F1F5F9")
    title_fill = PatternFill("solid", fgColor="E8F5E9")

    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:C1")
    ws["A1"] = f"خطة الأسبوع رقم {week_obj.week_no}"
    ws["A1"].font = title_font
    ws["A1"].alignment = center
    ws["A1"].fill = title_fill

    ws.merge_cells("A2:C2")
    ws["A2"] = f"المشرف: {supervisor.full_name} — الهوية: {_sup_nid_value(supervisor)}"
    ws["A2"].font = bold_font
    ws["A2"].alignment = center

    ws.append(["", "", ""])

    header_row = 4
    headers = ["اليوم", "المدرسة", "نوع الزيارة"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=h)
        cell.font = bold_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    days = {d.weekday: d for d in plan.days.all().select_related("school")}
    row = header_row + 1

    for wd, wd_name in WEEKDAYS:
        d = days.get(wd)
        school_name = d.school.name if d and d.school else "—"
        visit_label = d.get_visit_type_display() if d else "—"

        ws.cell(row=row, column=1, value=wd_name).font = normal_font
        ws.cell(row=row, column=2, value=school_name).font = normal_font
        ws.cell(row=row, column=3, value=visit_label).font = normal_font

        ws.cell(row=row, column=1).alignment = center
        ws.cell(row=row, column=2).alignment = right
        ws.cell(row=row, column=3).alignment = center

        for c in range(1, 4):
            ws.cell(row=row, column=c).border = border

        row += 1

    widths = {1: 16, 2: 55, 3: 18}
    for col_i, w in widths.items():
        ws.column_dimensions[get_column_letter(col_i)].width = w

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 22

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    filename = f"خطة_الأسبوع_{week_obj.week_no}_{_sup_nid_value(supervisor)}.xlsx"
    resp = HttpResponse(
        bio.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resp["Content-Disposition"] = f'attachment; filename="{filename}"'
    return resp


# =====================
# Unlock request (Supervisor)
# =====================
@require_POST
def request_unlock_view(request: HttpRequest) -> HttpResponse:
    try:
        supervisor = _require_supervisor(request)
    except Supervisor.DoesNotExist:
        return redirect("visits:login")

    plan_id = _safe_int(request.POST.get("plan") or 0, default=0)
    plan = get_object_or_404(Plan, id=plan_id, supervisor=supervisor)
    _inject_week_no(plan)

    if plan.status != Plan.STATUS_APPROVED:
        messages.info(request, "لا يمكن طلب فك اعتماد إلا لخطة معتمدة.")
        return redirect(_plan_url(plan.week.week_no))

    req, created = UnlockRequest.objects.get_or_create(plan=plan)
    if not created and req.status == UnlockRequest.STATUS_PENDING:
        messages.info(request, "يوجد طلب فك اعتماد سابق لهذه الخطة (بانتظار الإدارة).")
        return redirect(_plan_url(plan.week.week_no))

    req.status = UnlockRequest.STATUS_PENDING
    req.resolved_at = None
    req.save(update_fields=["status", "resolved_at"])

    plan.status = Plan.STATUS_UNLOCK_REQUESTED
    plan.save(update_fields=["status"])

    messages.success(request, "تم إرسال طلب فك الاعتماد ✅")
    return redirect(_plan_url(plan.week.week_no))


# =====================
# Admin dashboard
# =====================
@staff_member_required
def admin_dashboard_view(request: HttpRequest) -> HttpResponse:
    show_all = (request.GET.get("all") or "0").strip().lower() in ("1", "true", "yes")

    weeks_qs = _get_all_weeks_qs() if show_all else _get_active_weeks_qs()
    week_choices = _build_week_choices(active_only=(not show_all))

    default_week = weeks_qs.first()
    if not default_week:
        messages.warning(request, "لا يوجد أسابيع في جدول PlanWeek. أضف الأسابيع أولاً.")
        return render(request, "visits/admin_dashboard.html", {"rows": [], "page_obj": None})

    week_no = _safe_int(request.GET.get("week") or default_week.week_no, default=default_week.week_no)
    if not _week_exists(week_no):
        week_no = default_week.week_no

    week_obj = _resolve_week_or_404(week_no, allow_inactive=show_all)

    q = (request.GET.get("q") or "").strip()

    status = (request.GET.get("status") or "all").strip().lower()
    if status not in ("all", "approved", "draft", "unlock", "not_full"):
        status = "all"

    page_sizes = [10, 20, 30, 50]
    page_size = _safe_int(request.GET.get("ps") or 10, default=10)
    if page_size not in page_sizes:
        page_size = 10

    page = _safe_int(request.GET.get("page") or 1, default=1)

    base_qs = (
        Plan.objects.filter(week=week_obj)
        .select_related("supervisor", "week")
        .prefetch_related("days__school")
        .order_by("supervisor__full_name")
    )

    if q:
        cond = Q(supervisor__full_name__icontains=q)
        try:
            Supervisor._meta.get_field("national_id")
            cond |= Q(supervisor__national_id__icontains=q)
        except Exception:
            pass
        base_qs = base_qs.filter(cond)

    plans_all = list(base_qs)
    for p in plans_all:
        _inject_week_no(p)

    # KPIs
    kpi_total = len(plans_all)
    kpi_approved = sum(1 for p in plans_all if p.status == Plan.STATUS_APPROVED)
    kpi_drafts = sum(1 for p in plans_all if p.status == Plan.STATUS_DRAFT)
    kpi_unlock = sum(1 for p in plans_all if p.status == Plan.STATUS_UNLOCK_REQUESTED)
    kpi_filled = sum(1 for p in plans_all if p.is_fully_filled())
    kpi_not_filled = kpi_total - kpi_filled

    # Filter by status
    plans_filtered = plans_all
    if status == "approved":
        plans_filtered = [p for p in plans_all if p.status == Plan.STATUS_APPROVED]
    elif status == "draft":
        plans_filtered = [p for p in plans_all if p.status == Plan.STATUS_DRAFT]
    elif status == "unlock":
        plans_filtered = [p for p in plans_all if p.status == Plan.STATUS_UNLOCK_REQUESTED]
    elif status == "not_full":
        plans_filtered = [p for p in plans_all if not p.is_fully_filled()]

    rows = []
    for p in plans_filtered:
        day_map = {d.weekday: d for d in p.days.all()}
        filled = sum(1 for wd, _ in WEEKDAYS if day_map.get(wd) and day_map[wd].school_id)
        rows.append(
            {
                "plan": p,
                "sup": p.supervisor,
                "filled": filled,
                "is_full": (filled == 5),
                "day_map": day_map,
            }
        )

    paginator = Paginator(rows, page_size)
    page_obj = paginator.get_page(page)

    return render(
        request,
        "visits/admin_dashboard.html",
        {
            "rows": page_obj.object_list,
            "page_obj": page_obj,
            "week": week_obj.week_no,
            "week_obj": week_obj,
            "q": q,
            "status": status,
            "page_size": page_size,
            "page_sizes": page_sizes,
            "weekdays": WEEKDAYS,
            "week_choices": week_choices,
            "kpi_total": kpi_total,
            "kpi_approved": kpi_approved,
            "kpi_drafts": kpi_drafts,
            "kpi_unlock": kpi_unlock,
            "kpi_filled": kpi_filled,
            "kpi_not_filled": kpi_not_filled,
            "show_all": show_all,
        },
    )


# =====================
# ✅ Admin: Plan Details (PAGE)
# =====================
@staff_member_required
def admin_plan_detail_view(request: HttpRequest, plan_id: int) -> HttpResponse:
    """
    ✅ صفحة تفاصيل خطة مشرف
    - تُستخدم كرابط من لوحة الإدارة
    - وتحل NoReverseMatch
    """
    plan = get_object_or_404(
        Plan.objects.select_related("supervisor", "week").prefetch_related("days__school"),
        id=plan_id,
    )
    _inject_week_no(plan)

    sup = plan.supervisor
    day_map = {d.weekday: d for d in plan.days.all().select_related("school")}
    filled = sum(1 for wd, _ in WEEKDAYS if day_map.get(wd) and day_map[wd].school_id)
    day_dates = _build_day_dates_from_week(plan.week)

    return render(
        request,
        "visits/admin_plan_detail.html",
        {
            "plan": plan,
            "sup": sup,
            "weekdays": WEEKDAYS,
            "day_map": day_map,
            "filled": filled,
            "week": plan.week.week_no,
            "week_obj": plan.week,
            "day_dates": day_dates,
        },
    )


# =====================
# Admin: Export SINGLE plan Excel
# =====================
@staff_member_required
def admin_plan_export_excel(request: HttpRequest, plan_id: int) -> HttpResponse:
    plan = get_object_or_404(
        Plan.objects.select_related("supervisor", "week").prefetch_related("days__school"),
        id=plan_id,
    )
    _inject_week_no(plan)

    sup = plan.supervisor
    week_no = plan.week.week_no

    wb = Workbook()
    ws = wb.active
    ws.title = f"الأسبوع {week_no}"
    ws.sheet_view.rightToLeft = True

    title_font = Font(name="Cairo", bold=True, size=14)
    bold_font = Font(name="Cairo", bold=True, size=12)
    normal_font = Font(name="Cairo", size=12)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    header_fill = PatternFill("solid", fgColor="F1F5F9")
    title_fill = PatternFill("solid", fgColor="E8F5E9")

    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:C1")
    ws["A1"] = f"خطة الأسبوع رقم {week_no}"
    ws["A1"].font = title_font
    ws["A1"].alignment = center
    ws["A1"].fill = title_fill

    ws.merge_cells("A2:C2")
    ws["A2"] = f"المشرف: {sup.full_name} — الهوية: {_sup_nid_value(sup)}"
    ws["A2"].font = bold_font
    ws["A2"].alignment = center

    ws.append(["", "", ""])

    header_row = 4
    headers = ["اليوم", "المدرسة", "نوع الزيارة"]
    for col, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=h)
        cell.font = bold_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    days_map = {d.weekday: d for d in plan.days.all().select_related("school")}
    row = header_row + 1

    for wd, wd_name in WEEKDAYS:
        d = days_map.get(wd)
        school_name = d.school.name if d and d.school else "—"
        visit_label = d.get_visit_type_display() if d else "—"

        ws.cell(row=row, column=1, value=wd_name).font = normal_font
        ws.cell(row=row, column=2, value=school_name).font = normal_font
        ws.cell(row=row, column=3, value=visit_label).font = normal_font

        ws.cell(row=row, column=1).alignment = center
        ws.cell(row=row, column=2).alignment = right
        ws.cell(row=row, column=3).alignment = center

        for c in range(1, 4):
            ws.cell(row=row, column=c).border = border

        row += 1

    widths = {1: 16, 2: 55, 3: 18}
    for col_i, w in widths.items():
        ws.column_dimensions[get_column_letter(col_i)].width = w

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 22

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    filename = f"خطة_الأسبوع_{week_no}_{_sup_nid_value(sup)}.xlsx"
    resp = HttpResponse(
        bio.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resp["Content-Disposition"] = f'attachment; filename="{filename}"'
    return resp


# =====================
# Admin approve / back to draft (AJAX + Redirect)
# =====================
@staff_member_required
@require_POST
def admin_plan_approve_view(request: HttpRequest, plan_id: int) -> HttpResponse:
    plan = get_object_or_404(
        Plan.objects.select_related("supervisor", "week").prefetch_related("days"),
        id=plan_id,
    )
    _inject_week_no(plan)

    if plan.status == Plan.STATUS_APPROVED:
        msg = "الخطة معتمدة مسبقًا."
        filled = _plan_filled_count(plan)
        if _is_ajax(request):
            return JsonResponse(
                {
                    "ok": True,
                    "message": msg,
                    "plan_id": plan.id,
                    "status_label": _status_label(plan),
                    "status_css": _status_css(plan),
                    "status_code": _status_code(plan),
                    "filled": filled,
                    "is_full": (filled == 5),
                    "week_no": plan.week.week_no,
                },
                status=200,
            )
        messages.info(request, msg)
        return redirect(_admin_dashboard_url(plan.week.week_no))

    if not plan.is_fully_filled():
        msg = "لا يمكن اعتماد الخطة قبل اكتمال جميع الأيام (5/5)."
        if _is_ajax(request):
            return JsonResponse({"ok": False, "message": msg, "plan_id": plan.id}, status=400)
        messages.warning(request, msg)
        return redirect(_admin_dashboard_url(plan.week.week_no))

    plan.status = Plan.STATUS_APPROVED
    plan.approved_at = timezone.now()
    plan.save(update_fields=["status", "approved_at"])

    msg = "تم اعتماد الخطة ✅"
    filled = _plan_filled_count(plan)

    if _is_ajax(request):
        return JsonResponse(
            {
                "ok": True,
                "message": msg,
                "plan_id": plan.id,
                "status_label": _status_label(plan),
                "status_css": _status_css(plan),
                "status_code": _status_code(plan),
                "filled": filled,
                "is_full": (filled == 5),
                "week_no": plan.week.week_no,
            },
            status=200,
        )

    messages.success(request, msg)
    return redirect(_admin_dashboard_url(plan.week.week_no))


@staff_member_required
@require_POST
def admin_plan_back_to_draft_view(request: HttpRequest, plan_id: int) -> HttpResponse:
    plan = get_object_or_404(
        Plan.objects.select_related("week").prefetch_related("days"),
        id=plan_id,
    )
    _inject_week_no(plan)

    if plan.status == Plan.STATUS_DRAFT:
        msg = "الخطة بالفعل مسودة."
        filled = _plan_filled_count(plan)
        if _is_ajax(request):
            return JsonResponse(
                {
                    "ok": True,
                    "message": msg,
                    "plan_id": plan.id,
                    "status_label": _status_label(plan),
                    "status_css": _status_css(plan),
                    "status_code": _status_code(plan),
                    "filled": filled,
                    "is_full": (filled == 5),
                    "week_no": plan.week.week_no,
                },
                status=200,
            )
        messages.info(request, msg)
        return redirect(_admin_dashboard_url(plan.week.week_no))

    plan.status = Plan.STATUS_DRAFT
    plan.approved_at = None
    plan.save(update_fields=["status", "approved_at"])

    UnlockRequest.objects.filter(plan=plan).delete()

    msg = "تم إرجاع الخطة إلى مسودة ✅"
    filled = _plan_filled_count(plan)

    if _is_ajax(request):
        return JsonResponse(
            {
                "ok": True,
                "message": msg,
                "plan_id": plan.id,
                "status_label": _status_label(plan),
                "status_css": _status_css(plan),
                "status_code": _status_code(plan),
                "filled": filled,
                "is_full": (filled == 5),
                "week_no": plan.week.week_no,
            },
            status=200,
        )

    messages.success(request, msg)
    return redirect(_admin_dashboard_url(plan.week.week_no))


# =====================
# Admin export week Excel
# =====================
@staff_member_required
def admin_export_week_excel(request: HttpRequest) -> HttpResponse:
    show_all = (request.GET.get("all") or "0").strip().lower() in ("1", "true", "yes")

    week_no = _safe_int(request.GET.get("week") or _get_default_week_no(), default=_get_default_week_no())
    week_obj = _resolve_week_or_404(week_no, allow_inactive=show_all)

    wb = Workbook()
    ws = wb.active
    ws.title = f"الأسبوع {week_obj.week_no}"
    ws.sheet_view.rightToLeft = True

    ws.append(["المشرف", "اليوم", "المدرسة", "نوع الزيارة"])

    qs = (
        PlanDay.objects.filter(plan__week=week_obj)
        .select_related("plan__supervisor", "school")
        .order_by("plan__supervisor__full_name", "weekday")
    )

    for d in qs:
        day_label = WEEKDAY_MAP.get(d.weekday, str(d.weekday))
        ws.append(
            [
                d.plan.supervisor.full_name,
                day_label,
                d.school.name if d.school else "—",
                d.get_visit_type_display(),
            ]
        )

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)

    resp = HttpResponse(
        bio.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resp["Content-Disposition"] = f'attachment; filename="week_{week_obj.week_no}.xlsx"'
    return resp


# ==========================================================
# ✅ Admin Import (Excel)
# ==========================================================
def _import_stat() -> dict:
    return {"created": 0, "updated": 0, "skipped": 0}


def _norm_key(k: str) -> str:
    """توحيد أسماء الأعمدة"""
    k = (k or "").strip().lower()
    k = k.replace(" ", "").replace("-", "").replace("_", "")
    return k


def _cell_str(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _read_excel_dicts(file_obj) -> list[dict]:
    wb = load_workbook(file_obj, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    headers = [_norm_key(_cell_str(x)) for x in rows[0]]
    out = []
    for r in rows[1:]:
        row = {}
        for i, h in enumerate(headers):
            if not h:
                continue
            row[h] = r[i] if i < len(r) else None
        out.append(row)
    return out


def _pick(row: dict, keys: list[str], default: str = "") -> str:
    """يجيب قيمة من row حسب قائمة مفاتيح محتملة"""
    for k in keys:
        nk = _norm_key(k)
        if nk in row and row[nk] is not None:
            v = _cell_str(row[nk])
            if v != "":
                return v
    return default


def import_schools_excel(file_obj, *, gender: str) -> dict:
    st = _import_stat()
    rows = _read_excel_dicts(file_obj)

    for row in rows:
        stat_code = _pick(row, ["stat_code", "الرقمالاحصائي", "الرقمالإحصائي", "رقمالمدرسة", "code"])
        name = _pick(row, ["name", "اسمالمدرسة", "المدرسة", "schoolname"])

        stat_code = _cell_str(stat_code)
        name = _cell_str(name)

        if not stat_code or not name:
            st["skipped"] += 1
            continue

        _, created = School.objects.update_or_create(
            stat_code=stat_code,
            defaults={"name": name, "gender": gender, "is_active": True},
        )
        st["created" if created else "updated"] += 1

    return st


def import_supervisors_excel(file_obj) -> dict:
    st = _import_stat()
    rows = _read_excel_dicts(file_obj)

    for row in rows:
        nid = _pick(row, ["national_id", "السجلالمدني", "الهوية", "رقمالهوية", "nid"])
        full_name = _pick(row, ["full_name", "الاسم", "اسمالمشرف", "المشرف", "name"])
        mobile = _pick(row, ["mobile", "الجوال", "رقمالجوال", "الهاتف", "phone"])

        nid = _digits(_cell_str(nid))
        full_name = _cell_str(full_name)
        mobile = _cell_str(mobile)

        if len(nid) != 10 or not full_name:
            st["skipped"] += 1
            continue

        _, created = Supervisor.objects.update_or_create(
            national_id=nid,
            defaults={"full_name": full_name, "mobile": (mobile or None), "is_active": True},
        )
        st["created" if created else "updated"] += 1

    return st


def import_principals_excel(file_obj) -> dict:
    st = _import_stat()
    rows = _read_excel_dicts(file_obj)

    for row in rows:
        stat_code = _pick(row, ["stat_code", "school_stat_code", "رقمالمدرسة", "الرقمالإحصائي", "الرقمالاحصائي"])
        full_name = _pick(row, ["full_name", "الاسم", "اسمالمدير", "المدير"])
        mobile = _pick(row, ["mobile", "الجوال", "رقمالجوال", "phone"])

        stat_code = _cell_str(stat_code)
        full_name = _cell_str(full_name)
        mobile = _cell_str(mobile)

        if not stat_code or not full_name:
            st["skipped"] += 1
            continue

        school = School.objects.filter(stat_code=stat_code).first()
        if not school:
            stat_digits = _digits(stat_code)
            if stat_digits:
                school = School.objects.filter(stat_code=stat_digits).first()

        if not school:
            st["skipped"] += 1
            continue

        _, created = Principal.objects.update_or_create(
            school=school,
            defaults={"full_name": full_name, "mobile": (mobile or None)},
        )
        st["created" if created else "updated"] += 1

    return st


def import_assignments_excel(file_obj) -> dict:
    """يدعم ملف الإسنادات (مشرف-مدرسة)"""
    st = _import_stat()
    rows = _read_excel_dicts(file_obj)

    for row in rows:
        sup_nid = _pick(row, ["هويةالمشرف", "رقمالمشرف", "national_id", "السجلالمدني"])
        stat_code = _pick(row, ["stat_code", "school_stat_code", "رقمالمدرسة", "الرقمالإحصائي", "الرقمالاحصائي"])

        sup_nid = _digits(_cell_str(sup_nid))
        stat_code = _cell_str(stat_code)

        if len(sup_nid) != 10 or not stat_code:
            st["skipped"] += 1
            continue

        supervisor = Supervisor.objects.filter(national_id=sup_nid).first()
        if not supervisor:
            st["skipped"] += 1
            continue

        school = School.objects.filter(stat_code=stat_code).first()
        if not school:
            stat_digits = _digits(stat_code)
            if stat_digits:
                school = School.objects.filter(stat_code=stat_digits).first()

        if not school:
            st["skipped"] += 1
            continue

        obj, created = Assignment.objects.update_or_create(
            supervisor=supervisor,
            school=school,
            defaults={"is_active": True},
        )
        if created:
            st["created"] += 1
        else:
            if not obj.is_active:
                obj.is_active = True
                obj.save(update_fields=["is_active"])
            st["updated"] += 1

    return st


@staff_member_required
def admin_import_view(request: HttpRequest) -> HttpResponse:
    results: dict | None = None

    if request.method == "POST":
        form = ImportExcelForm(request.POST, request.FILES)
        if form.is_valid():
            results = {}
            try:
                with transaction.atomic():
                    if form.cleaned_data.get("schools_boys"):
                        results["مدارس البنين"] = import_schools_excel(
                            form.cleaned_data["schools_boys"], gender="boys"
                        )

                    if form.cleaned_data.get("schools_girls"):
                        results["مدارس البنات"] = import_schools_excel(
                            form.cleaned_data["schools_girls"], gender="girls"
                        )

                    if form.cleaned_data.get("supervisors"):
                        results["المشرفين"] = import_supervisors_excel(form.cleaned_data["supervisors"])

                    if form.cleaned_data.get("principals"):
                        results["مديري المدارس"] = import_principals_excel(form.cleaned_data["principals"])

                    if form.cleaned_data.get("assignments"):
                        results["الإسنادات (مشرف-مدرسة)"] = import_assignments_excel(
                            form.cleaned_data["assignments"]
                        )

                messages.success(request, "تم تنفيذ الاستيراد ✅")
            except Exception as e:
                messages.error(request, f"فشل الاستيراد: {e}")
        else:
            messages.error(request, "تحقق من الملفات: ارفع ملفًا واحدًا على الأقل.")
    else:
        form = ImportExcelForm()

    return render(request, "visits/manager_import.html", {"form": form, "results": results})
