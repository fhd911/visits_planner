from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from functools import wraps
from io import BytesIO
from typing import Any, Optional

from django.contrib import messages
from django.contrib.admin.views.decorators import staff_member_required
from django.contrib.auth import authenticate, login, logout
from django.core.paginator import Paginator
from django.db.models import Count, Q
from django.http import HttpRequest, HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone
from django.views.decorators.http import require_POST

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .forms import WeeklyLetterLinkForm
from .models import (
    Assignment,
    Plan,
    PlanDay,
    PlanWeek,
    Principal,
    School,
    Sector,
    SiteSetting,
    Supervisor,
    SupervisorNotification,
    UnlockRequest,
    WeeklyLetterLink,
)
from .services.drive_service import list_files_in_folder


# =============================================================================
# Constants
# =============================================================================
WEEKDAYS = [
    (0, "الأحد"),
    (1, "الإثنين"),
    (2, "الثلاثاء"),
    (3, "الأربعاء"),
    (4, "الخميس"),
]
WEEKDAY_MAP = dict(WEEKDAYS)
SESSION_SUP_ID = "visits_sup_id"


# =============================================================================
# Generic helpers
# =============================================================================
def _digits(s: str) -> str:
    return "".join(ch for ch in (s or "") if ch.isdigit())


def _safe_int(v: object, default: int = 1) -> int:
    try:
        return int(str(v).strip())
    except Exception:
        return default


def _cell_str(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _is_ajax(request: HttpRequest) -> bool:
    return request.headers.get("X-Requested-With") == "XMLHttpRequest"


def _bool_from_post(value: object, default: bool = True) -> bool:
    if value is None:
        return default
    return str(value).strip().lower() in ("1", "true", "yes", "on")


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
    nid = _digits(nid)
    if not nid:
        return None
    return Supervisor.objects.filter(national_id=nid, is_active=True).first()


def _sup_nid_value(sup: Supervisor) -> str:
    return getattr(sup, "national_id", "") or ""


def _inject_week_no(plan: Plan) -> None:
    try:
        plan.week_no = plan.week.week_no  # type: ignore[attr-defined]
    except Exception:
        plan.week_no = None  # type: ignore[attr-defined]


def _supervisor_school_ids(supervisor: Supervisor) -> set[int]:
    return set(
        Assignment.objects.filter(
            supervisor=supervisor,
            is_active=True,
        ).values_list("school_id", flat=True)
    )


def _supervisor_schools_qs(supervisor: Supervisor):
    return (
        School.objects.filter(
            assignments__supervisor=supervisor,
            assignments__is_active=True,
            is_active=True,
        )
        .distinct()
        .order_by("name")
    )


# =============================================================================
# Admin/supervisor isolation
# =============================================================================
def _has_supervisor_session(request: HttpRequest) -> bool:
    return bool(request.session.get(SESSION_SUP_ID))


def _redirect_supervisor_from_admin(request: HttpRequest) -> HttpResponse | None:
    if _has_supervisor_session(request):
        messages.warning(request, "ليس لديك صلاحية لدخول صفحات الإدارة أثناء تسجيلك كمشرف.")
        week_no = _safe_int(request.GET.get("week") or _get_default_week_no(), default=_get_default_week_no())
        return redirect(_plan_url(week_no))
    return None


def admin_only_view(view_func):
    @wraps(view_func)
    def _wrapped(request: HttpRequest, *args, **kwargs):
        supervisor_redirect = _redirect_supervisor_from_admin(request)
        if supervisor_redirect:
            return supervisor_redirect
        return staff_member_required(view_func)(request, *args, **kwargs)
    return _wrapped


# =============================================================================
# Site maintenance helpers
# =============================================================================
def _get_site_setting() -> SiteSetting:
    return SiteSetting.get_solo()


def _maintenance_message(setting: SiteSetting) -> str:
    return (
        setting.maintenance_message
        or "الموقع مغلق مؤقتًا للصيانة، وسيعود العمل خلال وقت قريب."
    )


def _parse_dt_local(value: str) -> Optional[datetime]:
    value = (value or "").strip()
    if not value:
        return None

    fmts = (
        "%Y-%m-%dT%H:%M",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
    )
    for fmt in fmts:
        try:
            dt = datetime.strptime(value, fmt)
            return timezone.make_aware(dt, timezone.get_current_timezone())
        except Exception:
            continue
    return None


def _format_dt_local(value: Optional[datetime]) -> str:
    if not value:
        return ""
    return timezone.localtime(value).strftime("%Y-%m-%dT%H:%M")


def _maintenance_is_active(setting: SiteSetting, *, persist: bool = False) -> bool:
    active = bool(setting.is_maintenance_mode)
    now = timezone.now()

    starts_at = getattr(setting, "maintenance_starts_at", None)
    ends_at = getattr(setting, "maintenance_ends_at", None)

    if not active:
        return False

    if starts_at and now < starts_at:
        return False

    if ends_at and now >= ends_at:
        if persist:
            setting.is_maintenance_mode = False
            setting.save(update_fields=["is_maintenance_mode", "updated_at"])
        return False

    return True


def _maintenance_context() -> dict[str, Any]:
    setting = _get_site_setting()
    active = _maintenance_is_active(setting, persist=True)
    return {
        "site_setting": setting,
        "maintenance_message": _maintenance_message(setting),
        "expected_return_text": setting.expected_return_text or "",
        "allow_admin_only": setting.allow_admin_only,
        "maintenance_is_active": active,
        "maintenance_starts_at": getattr(setting, "maintenance_starts_at", None),
        "maintenance_ends_at": getattr(setting, "maintenance_ends_at", None),
        "maintenance_starts_at_value": _format_dt_local(getattr(setting, "maintenance_starts_at", None)),
        "maintenance_ends_at_value": _format_dt_local(getattr(setting, "maintenance_ends_at", None)),
        "maintenance_window_label": getattr(setting, "maintenance_window_label", "غير محدد"),
    }


def _is_admin_user(request: HttpRequest) -> bool:
    user = getattr(request, "user", None)
    return bool(user and user.is_authenticated and user.is_staff)


def _maintenance_allowed_for_request(request: HttpRequest, setting: SiteSetting) -> bool:
    if not _maintenance_is_active(setting, persist=True):
        return True

    if _is_admin_user(request):
        return True

    if request.path.startswith("/admin/"):
        return True

    return False


# =============================================================================
# Maintenance views
# =============================================================================
def maintenance_page_view(request: HttpRequest) -> HttpResponse:
    setting = _get_site_setting()
    if not _maintenance_is_active(setting, persist=True):
        if request.user.is_authenticated and request.user.is_staff:
            return redirect("visits:admin_dashboard")
        return redirect("visits:login")

    context = _maintenance_context()
    return render(request, "visits/maintenance.html", context)


@admin_only_view
def admin_maintenance_settings_view(request: HttpRequest) -> HttpResponse:
    setting = _get_site_setting()
    _maintenance_is_active(setting, persist=True)
    return render(
        request,
        "visits/admin_maintenance_settings.html",
        {
            "site_setting": setting,
            "maintenance_starts_at_value": _format_dt_local(getattr(setting, "maintenance_starts_at", None)),
            "maintenance_ends_at_value": _format_dt_local(getattr(setting, "maintenance_ends_at", None)),
            "maintenance_window_label": getattr(setting, "maintenance_window_label", "غير محدد"),
            "maintenance_is_active": _maintenance_is_active(setting, persist=False),
        },
    )


@admin_only_view
@require_POST
def admin_toggle_maintenance_view(request: HttpRequest) -> HttpResponse:
    setting = _get_site_setting()

    enable_raw = (
        request.POST.get("enable")
        or request.POST.get("is_maintenance_mode")
        or request.POST.get("maintenance")
        or ""
    ).strip().lower()

    if enable_raw in ("1", "true", "yes", "on", "enable"):
        setting.is_maintenance_mode = True
        msg = "تم تفعيل وضع الصيانة بنجاح."
    elif enable_raw in ("0", "false", "no", "off", "disable"):
        setting.is_maintenance_mode = False
        msg = "تم إيقاف وضع الصيانة بنجاح."
    else:
        setting.is_maintenance_mode = not setting.is_maintenance_mode
        msg = "تم تحديث حالة وضع الصيانة بنجاح."

    setting.allow_admin_only = _bool_from_post(
        request.POST.get("allow_admin_only"),
        default=setting.allow_admin_only,
    )
    setting.save()
    active = _maintenance_is_active(setting, persist=True)

    if _is_ajax(request):
        return JsonResponse(
            {
                "ok": True,
                "message": msg,
                "is_maintenance_mode": setting.is_maintenance_mode,
                "maintenance_is_active": active,
                "allow_admin_only": setting.allow_admin_only,
                "maintenance_window_label": getattr(setting, "maintenance_window_label", "غير محدد"),
            },
            status=200,
        )

    messages.success(request, msg)
    return redirect("visits:admin_maintenance_settings")


@admin_only_view
@require_POST
def admin_update_maintenance_message_view(request: HttpRequest) -> HttpResponse:
    setting = _get_site_setting()

    starts_at_raw = (
        request.POST.get("maintenance_starts_at")
        or request.POST.get("starts_at")
        or ""
    )
    ends_at_raw = (
        request.POST.get("maintenance_ends_at")
        or request.POST.get("ends_at")
        or ""
    )

    starts_at = _parse_dt_local(starts_at_raw)
    ends_at = _parse_dt_local(ends_at_raw)

    if starts_at_raw.strip() and starts_at is None:
        messages.error(request, "صيغة تاريخ بداية الصيانة غير صحيحة.")
        return redirect("visits:admin_maintenance_settings")

    if ends_at_raw.strip() and ends_at is None:
        messages.error(request, "صيغة تاريخ نهاية الصيانة غير صحيحة.")
        return redirect("visits:admin_maintenance_settings")

    setting.site_name = (request.POST.get("site_name") or setting.site_name or "بوابة الزيارات").strip()
    setting.maintenance_message = (
        request.POST.get("maintenance_message")
        or request.POST.get("message")
        or ""
    ).strip() or None
    setting.expected_return_text = (
        request.POST.get("expected_return_text")
        or request.POST.get("expected_return")
        or ""
    ).strip() or None
    setting.maintenance_starts_at = starts_at
    setting.maintenance_ends_at = ends_at
    setting.allow_admin_only = _bool_from_post(
        request.POST.get("allow_admin_only"),
        default=setting.allow_admin_only,
    )
    setting.save()
    active = _maintenance_is_active(setting, persist=True)

    msg = "تم تحديث رسالة وإعدادات الصيانة بنجاح."

    if _is_ajax(request):
        return JsonResponse(
            {
                "ok": True,
                "message": msg,
                "site_name": setting.site_name,
                "maintenance_message": setting.maintenance_message or "",
                "expected_return_text": setting.expected_return_text or "",
                "allow_admin_only": setting.allow_admin_only,
                "is_maintenance_mode": setting.is_maintenance_mode,
                "maintenance_is_active": active,
                "maintenance_starts_at": _format_dt_local(setting.maintenance_starts_at),
                "maintenance_ends_at": _format_dt_local(setting.maintenance_ends_at),
                "maintenance_window_label": getattr(setting, "maintenance_window_label", "غير محدد"),
            },
            status=200,
        )

    messages.success(request, msg)
    return redirect("visits:admin_maintenance_settings")


# =============================================================================
# Google Drive helpers
# =============================================================================
def _extract_drive_folder_id(url: str) -> str | None:
    if not url:
        return None

    url = url.strip()
    patterns = [
        r"/folders/([a-zA-Z0-9_-]+)",
        r"[?&]id=([a-zA-Z0-9_-]+)",
    ]
    for pattern in patterns:
        m = re.search(pattern, url)
        if m:
            return m.group(1)
    return None


def _find_supervisor_letter_in_folder(*, folder_id: str, national_id: str) -> str | None:
    files = list_files_in_folder(folder_id, page_size=200)

    national_id = _digits(national_id)
    if not national_id:
        return None

    for file_obj in files:
        name = (file_obj.get("name") or "").strip()
        if national_id in name:
            return file_obj.get("webViewLink")

    return None


# =============================================================================
# Week helpers
# =============================================================================
def _get_active_weeks_qs():
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
    start = week_obj.start_sunday
    return {
        0: start,
        1: start + timedelta(days=1),
        2: start + timedelta(days=2),
        3: start + timedelta(days=3),
        4: start + timedelta(days=4),
    }


def _current_week_obj() -> Optional[PlanWeek]:
    today = timezone.localdate()
    return (
        PlanWeek.objects.filter(start_sunday__lte=today, is_break=False)
        .order_by("-start_sunday", "-week_no")
        .first()
    )


def _plan_url(week_no: int) -> str:
    return f"{reverse('visits:plan')}?week={week_no}"


def _notifications_url(week_no: int | None = None) -> str:
    url = reverse("visits:notifications")
    return f"{url}?week={week_no}" if week_no else url


def _supervisor_visit_status_url(week_no: int) -> str:
    return f"{reverse('visits:supervisor_visit_status')}?week={week_no}"


def _admin_dashboard_url(
    week_no: int,
    *,
    show_all: bool = False,
    q: str = "",
    status: str = "all",
    ps: int | None = None,
    page: int | None = None,
) -> str:
    params = [f"week={week_no}"]
    if show_all:
        params.append("all=1")
    if q:
        params.append(f"q={q}")
    if status and status != "all":
        params.append(f"status={status}")
    if ps:
        params.append(f"ps={ps}")
    if page:
        params.append(f"page={page}")
    return f"{reverse('visits:admin_dashboard')}?{'&'.join(params)}"


def _admin_plan_detail_url(plan_id: int, *, week_no: int | None = None, next_url: str = "") -> str:
    url = reverse("visits:admin_plan_detail", args=[plan_id])
    params = []
    if week_no:
        params.append(f"week={week_no}")
    if next_url:
        params.append(f"next={next_url}")
    return f"{url}?{'&'.join(params)}" if params else url


def _resolve_admin_return_url(
    request: HttpRequest,
    *,
    plan: Plan,
    default_week_no: int,
    show_all: bool,
    q: str,
    status_filter: str,
    ps: int,
    page: int,
) -> str:
    next_url = (request.POST.get("next") or request.GET.get("next") or "").strip()
    if next_url:
        return next_url

    back_to = (request.POST.get("back_to") or request.GET.get("back_to") or "").strip().lower()
    if back_to == "detail":
        return _admin_plan_detail_url(
            plan.id,
            week_no=plan.week.week_no,
            next_url=_admin_dashboard_url(
                default_week_no,
                show_all=show_all,
                q=q,
                status=status_filter,
                ps=ps,
                page=page,
            ),
        )

    return _admin_dashboard_url(
        default_week_no,
        show_all=show_all,
        q=q,
        status=status_filter,
        ps=ps,
        page=page,
    )


def _login_page_context() -> dict:
    return {
        "week_choices": _build_week_choices(active_only=False),
        "today": timezone.localdate(),
    }


# =============================================================================
# Status helpers
# =============================================================================
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


def _day_is_filled(d: Optional[PlanDay]) -> bool:
    if not d:
        return False
    if getattr(d, "school_id", None):
        return True
    return getattr(d, "visit_type", "") == getattr(PlanDay, "VISIT_NONE", "none")


def _plan_filled_count(plan: Plan) -> int:
    day_map = {d.weekday: d for d in plan.days.all()}
    return sum(1 for wd, _ in WEEKDAYS if _day_is_filled(day_map.get(wd)))


def _plan_visit_counts(plan: Plan) -> dict[str, int]:
    assigned_school_ids = set(
        Assignment.objects.filter(
            supervisor=plan.supervisor,
            is_active=True,
            school__is_active=True,
        ).values_list("school_id", flat=True)
    )

    visited_school_ids = set(
        PlanDay.objects.filter(
            plan=plan,
            visit_type=getattr(PlanDay, "VISIT_IN", "in"),
            school_id__isnull=False,
            visited=True,
        ).values_list("school_id", flat=True)
    )

    visited_school_ids &= assigned_school_ids
    total_assigned = len(assigned_school_ids)
    visited_count = len(visited_school_ids)
    remaining_count = max(total_assigned - visited_count, 0)

    return {
        "assigned_count": total_assigned,
        "visited_count": visited_count,
        "remaining_count": remaining_count,
    }


def _notify_supervisor(
    *,
    supervisor: Supervisor,
    plan: Plan | None,
    notif_type: str,
    title: str,
    message: str | None = None,
) -> None:
    SupervisorNotification.objects.create(
        supervisor=supervisor,
        plan=plan,
        notif_type=notif_type,
        title=title,
        message=message or "",
    )


def _parse_admin_action_state(request: HttpRequest) -> dict[str, Any]:
    show_all = (request.POST.get("all") or request.GET.get("all") or "0").strip().lower() in ("1", "true", "yes")
    return {
        "show_all": show_all,
        "q": (request.POST.get("q") or request.GET.get("q") or "").strip(),
        "status_filter": (request.POST.get("status") or request.GET.get("status") or "all").strip(),
        "ps": _safe_int(request.POST.get("ps") or request.GET.get("ps") or 10, default=10),
        "page": _safe_int(request.POST.get("page") or request.GET.get("page") or 1, default=1),
    }


def _build_plan_action_meta(plan: Plan) -> dict[str, Any]:
    filled = _plan_filled_count(plan)
    status_code = _status_code(plan)
    return {
        "plan_id": plan.id,
        "week_no": plan.week.week_no,
        "status_code": status_code,
        "status_css": _status_css(plan),
        "status_label": _status_label(plan),
        "admin_note": plan.admin_note or "",
        "filled": filled,
        "is_full": filled == 5,
        "can_approve": status_code != "approved" and filled == 5,
        "can_back_to_draft": status_code != "draft",
        "can_unlock_approve": status_code == "unlock",
        "can_unlock_reject": status_code == "unlock",
    }


def _plan_ajax_payload(plan: Plan, message: str, ok: bool = True, *, errors: list[str] | None = None) -> dict:
    payload = {
        "ok": ok,
        "message": message,
        "errors": errors or [],
    }
    payload.update(_build_plan_action_meta(plan))
    return payload


def _admin_json_response(
    request: HttpRequest,
    *,
    plan: Plan,
    message: str,
    ok: bool,
    http_status: int,
    errors: list[str] | None = None,
) -> JsonResponse | None:
    if not _is_ajax(request):
        return None
    return JsonResponse(_plan_ajax_payload(plan, message, ok=ok, errors=errors), status=http_status)


def _build_chart_counts(week_obj: PlanWeek) -> dict:
    base = (
        PlanDay.objects.filter(plan__week=week_obj)
        .values("weekday", "visit_type")
        .annotate(total=Count("id"))
        .order_by("weekday", "visit_type")
    )

    in_map = {wd: 0 for wd, _ in WEEKDAYS}
    remote_map = {wd: 0 for wd, _ in WEEKDAYS}
    none_map = {wd: 0 for wd, _ in WEEKDAYS}

    for item in base:
        wd = item["weekday"]
        vt = item["visit_type"]
        total = item["total"]

        if vt == getattr(PlanDay, "VISIT_IN", "in"):
            in_map[wd] = total
        elif vt == getattr(PlanDay, "VISIT_REMOTE", "remote"):
            remote_map[wd] = total
        elif vt == getattr(PlanDay, "VISIT_NONE", "none"):
            none_map[wd] = total

    return {
        "chart_labels": [name for _, name in WEEKDAYS],
        "chart_in_values": [in_map[wd] for wd, _ in WEEKDAYS],
        "chart_remote_values": [remote_map[wd] for wd, _ in WEEKDAYS],
        "chart_none_values": [none_map[wd] for wd, _ in WEEKDAYS],
        "chart_in_total": sum(in_map.values()),
        "chart_remote_total": sum(remote_map.values()),
        "chart_none_total": sum(none_map.values()),
    }


# =============================================================================
# Excel helpers
# =============================================================================
def _gender_label(value: str) -> str:
    v = (value or "").strip().lower()
    if v in ("boys", "male", "m", "بنين"):
        return "بنين"
    if v in ("girls", "female", "f", "بنات"):
        return "بنات"
    return value or "—"


def _excel_response(wb: Workbook, filename: str) -> HttpResponse:
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    response = HttpResponse(
        bio.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


def _build_plan_excel_workbook(plan: Plan) -> Workbook:
    _inject_week_no(plan)

    wb = Workbook()
    ws = wb.active
    ws.title = f"الأسبوع {plan.week.week_no}"
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
    ws["A1"] = f"خطة الأسبوع رقم {plan.week.week_no}"
    ws["A1"].font = title_font
    ws["A1"].alignment = center
    ws["A1"].fill = title_fill

    ws.merge_cells("A2:C2")
    ws["A2"] = f"المشرف: {plan.supervisor.full_name} — الهوية: {_sup_nid_value(plan.supervisor)}"
    ws["A2"].font = bold_font
    ws["A2"].alignment = center

    ws.append(["", "", ""])

    headers = ["اليوم", "المدرسة/السبب", "نوع اليوم"]
    header_row = 4
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = bold_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    days = {d.weekday: d for d in plan.days.all().select_related("school")}
    row = header_row + 1
    none_val = getattr(PlanDay, "VISIT_NONE", "none")

    for wd, wd_name in WEEKDAYS:
        d = days.get(wd)
        if d and d.visit_type == none_val:
            reason = d.get_no_visit_reason_display() if getattr(d, "no_visit_reason", None) else "بدون زيارة"
            if getattr(d, "note", None):
                reason = f"{reason} — {d.note}"
            school_or_reason = reason
        else:
            school_or_reason = d.school.name if d and d.school else "—"

        visit_label = d.get_visit_type_display() if d else "—"

        ws.cell(row=row, column=1, value=wd_name).font = normal_font
        ws.cell(row=row, column=2, value=school_or_reason).font = normal_font
        ws.cell(row=row, column=3, value=visit_label).font = normal_font

        ws.cell(row=row, column=1).alignment = center
        ws.cell(row=row, column=2).alignment = right
        ws.cell(row=row, column=3).alignment = center

        for c in range(1, 4):
            ws.cell(row=row, column=c).border = border
        row += 1

    for col_i, width in {1: 16, 2: 55, 3: 18}.items():
        ws.column_dimensions[get_column_letter(col_i)].width = width

    return wb


def _build_supervisor_assignments_excel_workbook(supervisor: Supervisor) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "المدارس المسندة"
    ws.sheet_view.rightToLeft = True

    title_font = Font(name="Cairo", bold=True, size=14)
    bold_font = Font(name="Cairo", bold=True, size=12)
    normal_font = Font(name="Cairo", size=11)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    header_fill = PatternFill("solid", fgColor="F1F5F9")
    title_fill = PatternFill("solid", fgColor="E8F5E9")

    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    assignments = (
        Assignment.objects.filter(
            supervisor=supervisor,
            is_active=True,
            school__is_active=True,
        )
        .select_related("school")
        .order_by("school__name")
    )

    school_ids = [a.school_id for a in assignments]
    principals_map = {p.school_id: p for p in Principal.objects.filter(school_id__in=school_ids)}

    ws.merge_cells("A1:G1")
    ws["A1"] = "المدارس المسندة للمشرف التربوي"
    ws["A1"].font = title_font
    ws["A1"].alignment = center
    ws["A1"].fill = title_fill

    ws.merge_cells("A2:G2")
    ws["A2"] = f"المشرف: {supervisor.full_name} — الهوية: {_sup_nid_value(supervisor)}"
    ws["A2"].font = bold_font
    ws["A2"].alignment = center

    ws.merge_cells("A3:G3")
    ws["A3"] = f"عدد المدارس: {assignments.count()}"
    ws["A3"].font = bold_font
    ws["A3"].alignment = center

    headers = ["م", "الرقم الإحصائي", "اسم المدرسة", "الجنس", "مدير المدرسة", "جوال المدير", "حالة الإسناد"]
    header_row = 5
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = bold_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    row_idx = header_row + 1
    for i, assignment in enumerate(assignments, start=1):
        school = assignment.school
        principal = principals_map.get(school.id)
        values = [
            i,
            school.stat_code or "—",
            school.name or "—",
            _gender_label(getattr(school, "gender", "")),
            getattr(principal, "full_name", "") or "—",
            getattr(principal, "mobile", "") or "—",
            "نشط" if assignment.is_active else "غير نشط",
        ]
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col, value=value)
            cell.font = normal_font
            cell.border = border
            cell.alignment = center if col in (1, 2, 4, 6, 7) else right
        row_idx += 1

    for col_i, width in {1: 8, 2: 18, 3: 42, 4: 12, 5: 28, 6: 18, 7: 14}.items():
        ws.column_dimensions[get_column_letter(col_i)].width = width

    return wb


def _build_admin_week_excel_workbook(week_obj: PlanWeek, plans) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = f"الأسبوع {week_obj.week_no}"
    ws.sheet_view.rightToLeft = True

    title_font = Font(name="Cairo", bold=True, size=14)
    bold_font = Font(name="Cairo", bold=True, size=12)
    normal_font = Font(name="Cairo", size=11)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    header_fill = PatternFill("solid", fgColor="F1F5F9")
    title_fill = PatternFill("solid", fgColor="E8F5E9")

    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:J1")
    ws["A1"] = f"تصدير خطط المشرفين — الأسبوع {week_obj.week_no}"
    ws["A1"].font = title_font
    ws["A1"].alignment = center
    ws["A1"].fill = title_fill

    headers = ["م", "اسم المشرف", "السجل المدني", "الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الحالة", "مكتملة"]
    header_row = 3
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = bold_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    none_val = getattr(PlanDay, "VISIT_NONE", "none")
    row_idx = header_row + 1

    for i, plan in enumerate(plans, start=1):
        day_map = {d.weekday: d for d in plan.days.all()}
        day_values: list[str] = []
        filled = 0

        for wd, _wd_name in WEEKDAYS:
            d = day_map.get(wd)
            text = "—"
            if d:
                if d.visit_type == none_val:
                    reason = d.get_no_visit_reason_display() if getattr(d, "no_visit_reason", None) else "بدون زيارة"
                    if getattr(d, "note", None):
                        reason = f"{reason} — {d.note}"
                    text = reason
                    filled += 1
                elif d.school:
                    text = d.school.name
                    filled += 1
            day_values.append(text)

        row_values = [
            i,
            plan.supervisor.full_name,
            _sup_nid_value(plan.supervisor),
            *day_values,
            _status_label(plan),
            "نعم" if filled == 5 else f"لا ({filled}/5)",
        ]

        for col, value in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col, value=value)
            cell.font = normal_font
            cell.border = border
            cell.alignment = center if col in (1, 3, 9, 10) else right

        row_idx += 1

    for col_i, width in {1: 8, 2: 24, 3: 18, 4: 26, 5: 26, 6: 26, 7: 26, 8: 26, 9: 16, 10: 14}.items():
        ws.column_dimensions[get_column_letter(col_i)].width = width

    return wb


# =============================================================================
# Auth views
# =============================================================================
def login_view(request: HttpRequest) -> HttpResponse:
    if request.user.is_authenticated and request.user.is_staff:
        return redirect("visits:admin_dashboard")

    setting = _get_site_setting()
    if _maintenance_is_active(setting, persist=True) and not _maintenance_allowed_for_request(request, setting):
        return redirect("visits:maintenance_page")

    context = _login_page_context()

    if request.method == "POST":
        nid = _digits(
            (
                request.POST.get("nid")
                or request.POST.get("national_id")
                or request.POST.get("civil_registry")
                or ""
            ).strip()
        )
        last4 = _digits(
            (
                request.POST.get("last4")
                or request.POST.get("phone_last4")
                or request.POST.get("mobile_last4")
                or ""
            ).strip()
        )

        if len(nid) != 10:
            messages.error(request, "فضلاً أدخل رقم الهوية بشكل صحيح.")
            return render(request, "visits/login.html", context)

        if len(last4) != 4:
            messages.error(request, "فضلاً أدخل آخر 4 أرقام من الجوال بشكل صحيح.")
            return render(request, "visits/login.html", context)

        sup = _find_supervisor_by_nid(nid)
        if not sup:
            messages.error(request, "المشرف غير موجود أو غير مفعل.")
            return render(request, "visits/login.html", context)

        sup_last4 = _supervisor_last4(sup)
        if not sup_last4:
            messages.error(request, "لا يمكن التحقق لأن رقم جوال المشرف غير محفوظ. راجع الإدارة لإضافته.")
            return render(request, "visits/login.html", context)

        if sup_last4 != last4:
            messages.error(request, "بيانات التحقق غير صحيحة.")
            return render(request, "visits/login.html", context)

        logout(request)
        request.session.flush()
        request.session[SESSION_SUP_ID] = sup.id
        request.session.modified = True
        messages.success(request, f"مرحبًا بك {sup.full_name}.")
        return redirect("visits:supervisor_dashboard")

    return render(request, "visits/login.html", context)


def admin_login_view(request: HttpRequest) -> HttpResponse:
    supervisor_redirect = _redirect_supervisor_from_admin(request)
    if supervisor_redirect:
        return supervisor_redirect

    if request.user.is_authenticated and request.user.is_staff:
        return redirect("visits:admin_dashboard")

    setting = _get_site_setting()
    if _maintenance_is_active(setting, persist=True) and not _maintenance_allowed_for_request(request, setting):
        return redirect("visits:maintenance_page")

    if request.method == "POST":
        username = (request.POST.get("username") or "").strip()
        password = request.POST.get("password") or ""

        if not username or not password:
            messages.error(request, "فضلاً أدخل اسم المستخدم وكلمة المرور.")
            return render(request, "visits/admin_login.html", {})

        user = authenticate(request, username=username, password=password)
        if not user:
            messages.error(request, "بيانات الدخول غير صحيحة.")
            return render(request, "visits/admin_login.html", {})

        if not user.is_staff:
            messages.error(request, "هذا الحساب لا يملك صلاحية الدخول للإدارة.")
            return render(request, "visits/admin_login.html", {})

        request.session.pop(SESSION_SUP_ID, None)
        login(request, user)
        messages.success(request, "تم تسجيل دخول الإدارة بنجاح.")
        return redirect("visits:admin_dashboard")

    return render(request, "visits/admin_login.html", {})


def logout_view(request: HttpRequest) -> HttpResponse:
    was_supervisor = bool(request.session.get(SESSION_SUP_ID))
    was_admin = bool(request.user.is_authenticated)

    request.session.pop(SESSION_SUP_ID, None)
    logout(request)

    if was_supervisor and not was_admin:
        messages.success(request, "تم تسجيل خروج المشرف بنجاح.")
        return redirect("visits:login")

    if was_admin:
        messages.success(request, "تم تسجيل خروج الإدارة بنجاح.")
        return redirect("visits:admin_login")

    messages.success(request, "تم تسجيل الخروج.")
    return redirect("visits:login")


def supervisor_dashboard_view(request: HttpRequest) -> HttpResponse:
    setting = _get_site_setting()
    if _maintenance_is_active(setting, persist=True) and not _maintenance_allowed_for_request(request, setting):
        return redirect("visits:maintenance_page")

    try:
        _require_supervisor(request)
    except Supervisor.DoesNotExist:
        return redirect("visits:login")
    return redirect(_plan_url(_get_default_week_no()))


# =============================================================================
# Letter views
# =============================================================================
def print_assignment_letter_view(request: HttpRequest) -> HttpResponse:
    nid = _digits((request.GET.get("nid") or "").strip())
    week = _safe_int(request.GET.get("week") or 0, default=0)

    if len(nid) != 10:
        messages.error(request, "يرجى إدخال السجل المدني بشكل صحيح.")
        return redirect("visits:login")

    if not PlanWeek.objects.filter(week_no=week).exists():
        messages.error(request, "يرجى اختيار أسبوع صحيح.")
        return redirect("visits:login")

    sup = _find_supervisor_by_nid(nid)
    if not sup:
        messages.error(request, "لم يتم العثور على مشرف بهذا السجل المدني.")
        return redirect("visits:login")

    link_obj = WeeklyLetterLink.objects.filter(week__week_no=week, is_active=True).select_related("week").first()
    if not link_obj or not link_obj.drive_url:
        messages.warning(request, "لا يوجد رابط خطاب منشور لهذا الأسبوع.")
        return redirect("visits:login")

    folder_id = _extract_drive_folder_id(link_obj.drive_url)
    if not folder_id:
        messages.error(request, "رابط الأسبوع غير صالح. يجب أن يكون رابط مجلد Google Drive صحيح.")
        return redirect("visits:login")

    try:
        file_url = _find_supervisor_letter_in_folder(folder_id=folder_id, national_id=nid)
    except Exception as e:
        messages.error(request, f"تعذر البحث عن خطاب المشرف داخل Google Drive: {e}")
        return redirect("visits:login")

    if not file_url:
        messages.warning(
            request,
            "لم يتم العثور على خطاب مطابق للسجل المدني داخل مجلد هذا الأسبوع. "
            "تأكد أن اسم الملف يحتوي السجل المدني للمشرف."
        )
        return redirect("visits:login")

    return redirect(file_url)


def current_week_letter_redirect_view(request: HttpRequest) -> HttpResponse:
    try:
        supervisor = _require_supervisor(request)
    except Supervisor.DoesNotExist:
        return redirect("visits:login")

    week_obj = _current_week_obj()
    if not week_obj:
        messages.error(request, "لم يتم العثور على أسبوع حالي.")
        return redirect("visits:supervisor_dashboard")

    link_obj = WeeklyLetterLink.objects.filter(week=week_obj, is_active=True).first()
    if not link_obj or not link_obj.drive_url:
        messages.warning(request, "لا يوجد رابط خطاب منشور لهذا الأسبوع.")
        return redirect(_plan_url(week_obj.week_no))

    folder_id = _extract_drive_folder_id(link_obj.drive_url)
    if not folder_id:
        messages.error(request, "رابط الأسبوع غير صالح. يجب أن يكون رابط مجلد Google Drive صحيح.")
        return redirect(_plan_url(week_obj.week_no))

    nid = _digits(_sup_nid_value(supervisor))
    if len(nid) != 10:
        messages.error(request, "السجل المدني للمشرف غير محفوظ بشكل صحيح.")
        return redirect(_plan_url(week_obj.week_no))

    try:
        file_url = _find_supervisor_letter_in_folder(folder_id=folder_id, national_id=nid)
    except Exception as e:
        messages.error(request, f"تعذر البحث عن الخطاب داخل Google Drive: {e}")
        return redirect(_plan_url(week_obj.week_no))

    if not file_url:
        messages.warning(
            request,
            "لم يتم العثور على خطابك داخل مجلد هذا الأسبوع. "
            "تأكد أن اسم الملف في Google Drive يحتوي السجل المدني."
        )
        return redirect(_plan_url(week_obj.week_no))

    return redirect(file_url)


@admin_only_view
def weekly_letters_drive_view(request: HttpRequest, week_number: int) -> HttpResponse:
    link_obj = WeeklyLetterLink.objects.filter(
        week__week_no=week_number,
        is_active=True,
    ).select_related("week").first()

    week_folder = None
    rows = []

    if link_obj and link_obj.drive_url:
        folder_id = _extract_drive_folder_id(link_obj.drive_url)
        if folder_id:
            try:
                files = list_files_in_folder(folder_id, page_size=300)
                week_folder = {
                    "id": folder_id,
                    "name": f"الأسبوع {week_number}",
                    "webViewLink": link_obj.drive_url,
                }
                for f in files:
                    filename = (f.get("name") or "").strip()
                    rows.append(
                        {
                            "school_code": filename.removesuffix(".pdf").strip(),
                            "name": filename,
                            "url": f.get("webViewLink"),
                            "file_id": f.get("id"),
                            "mime_type": f.get("mimeType"),
                        }
                    )
            except Exception as e:
                messages.error(request, f"تعذر قراءة ملفات Google Drive: {e}")
        else:
            messages.error(request, "رابط Google Drive المحفوظ غير صالح.")
    else:
        messages.warning(request, "لا يوجد رابط نشط محفوظ لهذا الأسبوع.")

    return render(
        request,
        "visits/weekly_letters_drive.html",
        {
            "week_number": week_number,
            "week_folder": week_folder,
            "rows": rows,
            "link_obj": link_obj,
        },
    )


# =============================================================================
# Supervisor plan views
# =============================================================================
def plan_view(request: HttpRequest) -> HttpResponse:
    setting = _get_site_setting()
    if _maintenance_is_active(setting, persist=True) and not _maintenance_allowed_for_request(request, setting):
        return redirect("visits:maintenance_page")

    try:
        supervisor = _require_supervisor(request)
    except Supervisor.DoesNotExist:
        return redirect("visits:login")

    week_no = _safe_int(request.GET.get("week") or request.POST.get("week") or _get_default_week_no(), default=_get_default_week_no())
    week_obj = _resolve_week_or_404(week_no, allow_inactive=False)

    plan, _ = Plan.objects.get_or_create(supervisor=supervisor, week=week_obj)
    _inject_week_no(plan)

    schools = _supervisor_schools_qs(supervisor)
    allowed_school_ids = _supervisor_school_ids(supervisor)
    days_map = {d.weekday: d for d in plan.days.all().select_related("school")}
    week_choices = _build_week_choices(active_only=True)
    day_dates = _build_day_dates_from_week(week_obj)
    notifications = supervisor.notifications.select_related("plan").order_by("-created_at")[:8]
    unread_count = supervisor.notifications.filter(is_read=False).count()
    week_letter = WeeklyLetterLink.objects.filter(week=week_obj, is_active=True).first()

    if request.method == "POST":
        if plan.status in (Plan.STATUS_APPROVED, Plan.STATUS_UNLOCK_REQUESTED):
            messages.info(request, "لا يمكن تعديل الخطة الآن. إذا كانت معتمدة فاطلب فك الاعتماد أولًا.")
            return redirect(_plan_url(week_obj.week_no))

        action = (request.POST.get("action") or "save").strip()

        for wd, _ in WEEKDAYS:
            sid = _safe_int((request.POST.get(f"school_{wd}") or "").strip(), default=0)
            vtype = (request.POST.get(f"visit_{wd}") or getattr(PlanDay, "VISIT_IN", "in")).strip()

            allowed_visit_types = {
                getattr(PlanDay, "VISIT_IN", "in"),
                getattr(PlanDay, "VISIT_REMOTE", "remote"),
                getattr(PlanDay, "VISIT_NONE", "none"),
            }
            if vtype not in allowed_visit_types:
                vtype = getattr(PlanDay, "VISIT_IN", "in")

            reason = (request.POST.get(f"reason_{wd}") or "").strip() or None
            note = (request.POST.get(f"note_{wd}") or "").strip() or None

            if vtype == getattr(PlanDay, "VISIT_NONE", "none"):
                sid = 0

            if sid and sid not in allowed_school_ids:
                messages.warning(request, f"تم تجاهل مدرسة غير مسندة للمشرف في يوم {WEEKDAY_MAP.get(wd, wd)}.")
                PlanDay.objects.filter(plan=plan, weekday=wd).delete()
                continue

            if (not sid) and (vtype != getattr(PlanDay, "VISIT_NONE", "none")):
                PlanDay.objects.filter(plan=plan, weekday=wd).delete()
                continue

            old_day = PlanDay.objects.filter(plan=plan, weekday=wd).first()
            visited = getattr(old_day, "visited", False) if old_day else False
            visited_at = getattr(old_day, "visited_at", None) if old_day else None
            visit_note = getattr(old_day, "visit_note", None) if old_day else None

            if vtype != getattr(PlanDay, "VISIT_IN", "in") or not sid:
                visited = False
                visited_at = None
                visit_note = None

            defaults = {
                "visit_type": vtype,
                "note": note,
                "visited": visited,
                "visited_at": visited_at,
                "visit_note": visit_note,
            }

            if sid:
                defaults["school_id"] = sid
                defaults["no_visit_reason"] = None
            else:
                defaults["school"] = None
                defaults["no_visit_reason"] = reason

            PlanDay.objects.update_or_create(plan=plan, weekday=wd, defaults=defaults)

        plan.saved_at = timezone.now()

        if action == "approve":
            if plan.is_fully_filled():
                plan.status = Plan.STATUS_APPROVED
                plan.approved_at = timezone.now()
                plan.admin_note = None
                plan.save(update_fields=["saved_at", "status", "approved_at", "admin_note"])
                messages.success(request, "تم اعتماد الخطة بنجاح.")
            else:
                plan.save(update_fields=["saved_at"])
                messages.warning(request, "تم الحفظ، لكن لا يمكن الاعتماد قبل اكتمال جميع الأيام.")
        else:
            plan.save(update_fields=["saved_at"])
            messages.success(request, "تم حفظ الخطة بنجاح.")

        return redirect(_plan_url(week_obj.week_no))

    visit_counts = _plan_visit_counts(plan)

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
            "today": timezone.localdate(),
            "notifications": notifications,
            "unread_count": unread_count,
            "visit_none_value": getattr(PlanDay, "VISIT_NONE", "none"),
            "visit_in_value": getattr(PlanDay, "VISIT_IN", "in"),
            "visit_remote_value": getattr(PlanDay, "VISIT_REMOTE", "remote"),
            "week_letter": week_letter,
            "assigned_count": visit_counts["assigned_count"],
            "visited_count": visit_counts["visited_count"],
            "remaining_count": visit_counts["remaining_count"],
        },
    )


def supervisor_visit_status_view(request: HttpRequest) -> HttpResponse:
    try:
        supervisor = _require_supervisor(request)
    except Supervisor.DoesNotExist:
        return redirect("visits:login")

    week_no = _safe_int(request.GET.get("week") or _get_default_week_no(), default=_get_default_week_no())
    week_obj = _resolve_week_or_404(week_no, allow_inactive=False)

    plan = get_object_or_404(
        Plan.objects.select_related("supervisor", "week").prefetch_related("days__school"),
        supervisor=supervisor,
        week=week_obj,
    )
    _inject_week_no(plan)

    school_days = [d for d in plan.days.all() if d.visit_type == getattr(PlanDay, "VISIT_IN", "in") and d.school_id]
    visited_days = [d for d in school_days if getattr(d, "visited", False)]
    unvisited_days = [d for d in school_days if not getattr(d, "visited", False)]

    notifications = supervisor.notifications.select_related("plan").order_by("-created_at")[:8]
    unread_count = supervisor.notifications.filter(is_read=False).count()
    visit_counts = _plan_visit_counts(plan)

    assigned_school_ids = set(
        Assignment.objects.filter(
            supervisor=supervisor,
            is_active=True,
            school__is_active=True,
        ).values_list("school_id", flat=True)
    )
    remaining_school_ids = assigned_school_ids - set(d.school_id for d in visited_days if d.school_id)
    remaining_schools = list(School.objects.filter(id__in=remaining_school_ids, is_active=True).order_by("name"))

    return render(
        request,
        "visits/supervisor_visit_status.html",
        {
            "sup": supervisor,
            "plan": plan,
            "week": week_obj.week_no,
            "week_obj": week_obj,
            "school_days": school_days,
            "visited_days": visited_days,
            "unvisited_days": unvisited_days,
            "remaining_schools": remaining_schools,
            "assigned_count": visit_counts["assigned_count"],
            "visited_count": visit_counts["visited_count"],
            "remaining_count": visit_counts["remaining_count"],
            "notifications": notifications,
            "unread_count": unread_count,
        },
    )


@require_POST
def toggle_day_visited_view(request: HttpRequest, day_id: int) -> HttpResponse:
    try:
        supervisor = _require_supervisor(request)
    except Supervisor.DoesNotExist:
        return redirect("visits:login")

    day = get_object_or_404(
        PlanDay.objects.select_related("plan", "plan__supervisor", "plan__week", "school"),
        id=day_id,
        plan__supervisor=supervisor,
    )

    if day.visit_type != getattr(PlanDay, "VISIT_IN", "in") or not day.school_id:
        msg = "لا يمكن تتبع هذا اليوم لأنه ليس زيارة حضورية مدرسية."
        if _is_ajax(request):
            return JsonResponse({"ok": False, "message": msg}, status=400)
        messages.warning(request, msg)
        return redirect(_supervisor_visit_status_url(day.plan.week.week_no))

    visited_raw = (request.POST.get("visited") or "").strip()
    note = (request.POST.get("visit_note") or request.POST.get("note") or "").strip() or None

    if visited_raw in ("1", "true", "yes", "on"):
        day.visited = True
        day.visited_at = timezone.now()
        day.visit_note = note
        msg = "تم تسجيل الزيارة بنجاح."
    else:
        day.visited = False
        day.visited_at = None
        day.visit_note = None
        msg = "تم إلغاء تسجيل الزيارة."

    day.save()
    visit_counts = _plan_visit_counts(day.plan)

    if _is_ajax(request):
        return JsonResponse(
            {
                "ok": True,
                "message": msg,
                "day_id": day.id,
                "plan_id": day.plan.id,
                "visited": day.visited,
                "visited_at": day.visited_at.strftime("%Y-%m-%d %H:%M") if day.visited_at else "",
                "visit_note": day.visit_note or "",
                "assigned_count": visit_counts["assigned_count"],
                "visited_count": visit_counts["visited_count"],
                "remaining_count": visit_counts["remaining_count"],
            },
            status=200,
        )

    messages.success(request, msg)
    return redirect(_supervisor_visit_status_url(day.plan.week.week_no))


# =============================================================================
# Notifications
# =============================================================================
def notifications_view(request: HttpRequest) -> HttpResponse:
    setting = _get_site_setting()
    if _maintenance_is_active(setting, persist=True) and not _maintenance_allowed_for_request(request, setting):
        return redirect("visits:maintenance_page")

    try:
        supervisor = _require_supervisor(request)
    except Supervisor.DoesNotExist:
        return redirect("visits:login")

    week_no = _safe_int(request.GET.get("week") or _get_default_week_no(), default=_get_default_week_no())
    notifications_qs = supervisor.notifications.select_related("plan", "plan__week").all()

    paginator = Paginator(notifications_qs, 20)
    page = _safe_int(request.GET.get("page") or 1, default=1)
    page_obj = paginator.get_page(page)

    unread_count = supervisor.notifications.filter(is_read=False).count()

    return render(
        request,
        "visits/notifications.html",
        {
            "notifications": page_obj.object_list,
            "page_obj": page_obj,
            "sup": supervisor,
            "unread_count": unread_count,
            "week": week_no,
        },
    )


@require_POST
def mark_notification_read_view(request: HttpRequest, notif_id: int) -> HttpResponse:
    try:
        supervisor = _require_supervisor(request)
    except Supervisor.DoesNotExist:
        return redirect("visits:login")

    notif = get_object_or_404(SupervisorNotification, id=notif_id, supervisor=supervisor)
    if not notif.is_read:
        notif.is_read = True
        notif.save(update_fields=["is_read"])

    next_url = request.POST.get("next") or request.GET.get("next") or _notifications_url()
    return redirect(next_url)


@require_POST
def mark_all_notifications_read_view(request: HttpRequest) -> HttpResponse:
    try:
        supervisor = _require_supervisor(request)
    except Supervisor.DoesNotExist:
        return redirect("visits:login")

    supervisor.notifications.filter(is_read=False).update(is_read=True)
    messages.success(request, "تم تعليم جميع الإشعارات كمقروءة.")
    next_url = request.POST.get("next") or request.GET.get("next") or _notifications_url()
    return redirect(next_url)


# =============================================================================
# Supervisor exports
# =============================================================================
def export_plan_excel(request: HttpRequest) -> HttpResponse:
    try:
        supervisor = _require_supervisor(request)
    except Supervisor.DoesNotExist:
        return redirect("visits:login")

    week_no = _safe_int(request.GET.get("week") or _get_default_week_no(), default=_get_default_week_no())
    week_obj = _resolve_week_or_404(week_no, allow_inactive=False)

    plan = get_object_or_404(
        Plan.objects.select_related("supervisor", "week").prefetch_related("days__school"),
        supervisor=supervisor,
        week=week_obj,
    )
    wb = _build_plan_excel_workbook(plan)
    filename = f"خطة_الأسبوع_{week_obj.week_no}_{_sup_nid_value(supervisor)}.xlsx"
    return _excel_response(wb, filename)


def export_supervisor_assignments_excel(request: HttpRequest) -> HttpResponse:
    try:
        supervisor = _require_supervisor(request)
    except Supervisor.DoesNotExist:
        return redirect("visits:login")

    wb = _build_supervisor_assignments_excel_workbook(supervisor)
    filename = f"المدارس_المسندة_{_sup_nid_value(supervisor)}.xlsx"
    return _excel_response(wb, filename)


# =============================================================================
# Admin schools
# =============================================================================
@admin_only_view
def admin_school_list_view(request: HttpRequest) -> HttpResponse:
    q = _cell_str(request.GET.get("q"))
    gender = _cell_str(request.GET.get("gender"))
    sector_id = _safe_int(request.GET.get("sector") or 0, default=0)
    only_active = request.GET.get("active", "1") == "1"

    qs = School.objects.select_related("sector").order_by("name")

    if only_active:
        qs = qs.filter(is_active=True)
    if q:
        qs = qs.filter(Q(name__icontains=q) | Q(stat_code__icontains=q))
    if gender in ("boys", "girls"):
        qs = qs.filter(gender=gender)
    if sector_id:
        qs = qs.filter(sector_id=sector_id)

    sectors = Sector.objects.filter(is_active=True).order_by("name")
    paginator = Paginator(qs, 30)
    page_obj = paginator.get_page(_safe_int(request.GET.get("page") or 1, default=1))

    return render(
        request,
        "visits/admin_school_list.html",
        {
            "rows": page_obj.object_list,
            "page_obj": page_obj,
            "q": q,
            "gender": gender,
            "sector_id": sector_id,
            "only_active": only_active,
            "sectors": sectors,
            "kpi_total": qs.count(),
            "kpi_active": qs.filter(is_active=True).count(),
            "kpi_inactive": qs.filter(is_active=False).count(),
        },
    )


@admin_only_view
@require_POST
def admin_school_save_view(request: HttpRequest) -> HttpResponse:
    school_id = _safe_int(request.POST.get("school_id") or 0, default=0)

    name = _cell_str(request.POST.get("name"))
    stat_code = _cell_str(request.POST.get("stat_code"))
    gender = _cell_str(request.POST.get("gender"))
    sector_id = _safe_int(request.POST.get("sector") or request.POST.get("sector_id") or 0, default=0)
    is_active = _bool_from_post(request.POST.get("is_active"), default=True)

    if not name:
        messages.error(request, "اسم المدرسة مطلوب.")
        return redirect("visits:admin_school_list")

    sector = Sector.objects.filter(id=sector_id).first() if sector_id else None

    if school_id:
        school = get_object_or_404(School, id=school_id)
        school.name = name
        if hasattr(school, "stat_code"):
            school.stat_code = stat_code
        if hasattr(school, "gender") and gender:
            school.gender = gender
        if hasattr(school, "sector"):
            school.sector = sector
        if hasattr(school, "is_active"):
            school.is_active = is_active
        school.save()
        messages.success(request, "تم تحديث المدرسة بنجاح.")
    else:
        data = {"name": name}
        if hasattr(School, "stat_code"):
            data["stat_code"] = stat_code
        if hasattr(School, "gender"):
            data["gender"] = gender or "boys"
        if hasattr(School, "sector"):
            data["sector"] = sector
        if hasattr(School, "is_active"):
            data["is_active"] = is_active
        School.objects.create(**data)
        messages.success(request, "تمت إضافة المدرسة بنجاح.")

    return redirect("visits:admin_school_list")


@admin_only_view
@require_POST
def admin_school_toggle_active_view(request: HttpRequest, school_id: int) -> HttpResponse:
    school = get_object_or_404(School, id=school_id)
    school.is_active = not bool(school.is_active)
    school.save(update_fields=["is_active"])
    messages.success(request, "تم تحديث حالة المدرسة.")
    return redirect("visits:admin_school_list")


# =============================================================================
# Admin supervisors
# =============================================================================
@admin_only_view
def admin_supervisor_list_view(request: HttpRequest) -> HttpResponse:
    q = _cell_str(request.GET.get("q"))
    gender = _cell_str(request.GET.get("gender"))
    sector_id = _safe_int(request.GET.get("sector") or 0, default=0)
    only_active = request.GET.get("active", "1") == "1"

    qs = (
        Supervisor.objects.select_related("sector")
        .annotate(
            active_assignment_count=Count("assignments", filter=Q(assignments__is_active=True), distinct=True),
            total_assignment_count=Count("assignments", distinct=True),
        )
        .order_by("full_name")
    )

    if only_active:
        qs = qs.filter(is_active=True)
    if q:
        qs = qs.filter(Q(full_name__icontains=q) | Q(national_id__icontains=q))
    if gender in ("boys", "girls"):
        qs = qs.filter(gender=gender)
    if sector_id:
        qs = qs.filter(sector_id=sector_id)

    sectors = Sector.objects.filter(is_active=True).order_by("name")
    paginator = Paginator(qs, 30)
    page_obj = paginator.get_page(_safe_int(request.GET.get("page") or 1, default=1))

    return render(
        request,
        "visits/admin_supervisor_list.html",
        {
            "rows": page_obj.object_list,
            "page_obj": page_obj,
            "q": q,
            "gender": gender,
            "sector_id": sector_id,
            "only_active": only_active,
            "sectors": sectors,
            "kpi_total": qs.count(),
            "kpi_active": qs.filter(is_active=True).count(),
            "kpi_inactive": qs.filter(is_active=False).count(),
        },
    )


@admin_only_view
@require_POST
def admin_supervisor_save_view(request: HttpRequest) -> HttpResponse:
    supervisor_id = _safe_int(request.POST.get("supervisor_id") or 0, default=0)

    full_name = _cell_str(request.POST.get("full_name"))
    national_id = _cell_str(request.POST.get("national_id"))
    mobile = _cell_str(request.POST.get("mobile"))
    gender = _cell_str(request.POST.get("gender"))
    sector_id = _safe_int(request.POST.get("sector") or request.POST.get("sector_id") or 0, default=0)
    is_active = _bool_from_post(request.POST.get("is_active"), default=True)

    if not full_name:
        messages.error(request, "اسم المشرف مطلوب.")
        return redirect("visits:admin_supervisor_list")

    sector = Sector.objects.filter(id=sector_id).first() if sector_id else None

    if supervisor_id:
        supervisor = get_object_or_404(Supervisor, id=supervisor_id)
        supervisor.full_name = full_name
        if hasattr(supervisor, "national_id"):
            supervisor.national_id = national_id
        if hasattr(supervisor, "mobile"):
            supervisor.mobile = mobile
        if hasattr(supervisor, "gender") and gender:
            supervisor.gender = gender
        if hasattr(supervisor, "sector"):
            supervisor.sector = sector
        if hasattr(supervisor, "is_active"):
            supervisor.is_active = is_active
        supervisor.save()
        messages.success(request, "تم تحديث المشرف بنجاح.")
    else:
        data = {"full_name": full_name}
        if hasattr(Supervisor, "national_id"):
            data["national_id"] = national_id
        if hasattr(Supervisor, "mobile"):
            data["mobile"] = mobile
        if hasattr(Supervisor, "gender"):
            data["gender"] = gender or "boys"
        if hasattr(Supervisor, "sector"):
            data["sector"] = sector
        if hasattr(Supervisor, "is_active"):
            data["is_active"] = is_active
        Supervisor.objects.create(**data)
        messages.success(request, "تمت إضافة المشرف بنجاح.")

    return redirect("visits:admin_supervisor_list")


@admin_only_view
@require_POST
def admin_supervisor_toggle_active_view(request: HttpRequest, supervisor_id: int) -> HttpResponse:
    supervisor = get_object_or_404(Supervisor, id=supervisor_id)
    supervisor.is_active = not bool(supervisor.is_active)
    supervisor.save(update_fields=["is_active"])
    messages.success(request, "تم تحديث حالة المشرف.")
    return redirect("visits:admin_supervisor_list")


def _supervisor_assignment_scope(supervisor: Supervisor):
    qs = School.objects.filter(is_active=True)
    gender = getattr(supervisor, "gender", "")
    if gender in ("boys", "girls"):
        qs = qs.filter(gender=gender)
    if getattr(supervisor, "sector_id", None):
        qs = qs.filter(sector_id=supervisor.sector_id)
    return qs.order_by("name")


def _globally_assigned_school_ids(*, exclude_supervisor_id: int | None = None) -> set[int]:
    qs = Assignment.objects.filter(is_active=True).exclude(school_id__isnull=True)
    if exclude_supervisor_id:
        qs = qs.exclude(supervisor_id=exclude_supervisor_id)
    return set(qs.values_list("school_id", flat=True))


def _find_existing_school_for_assignment(*, stat_code: str, name: str, gender: str, sector_id: int | None):
    qs = School.objects.all()
    if stat_code:
        obj = qs.filter(stat_code=stat_code).first()
        if obj:
            return obj
    if name:
        qs = qs.filter(name=name)
        if gender in ("boys", "girls"):
            qs = qs.filter(gender=gender)
        if sector_id:
            qs = qs.filter(sector_id=sector_id)
        return qs.first()
    return None


# =============================================================================
# Admin assignments
# =============================================================================
@admin_only_view
def admin_assignments_overview_view(request: HttpRequest) -> HttpResponse:
    q = _cell_str(request.GET.get("q"))
    gender = _cell_str(request.GET.get("gender"))
    sector_id = _safe_int(request.GET.get("sector") or 0, default=0)
    only_active = request.GET.get("active", "1") == "1"

    supervisors = (
        Supervisor.objects.select_related("sector")
        .annotate(
            active_assignment_count=Count(
                "assignments",
                filter=Q(assignments__is_active=True, assignments__school__is_active=True),
                distinct=True,
            ),
            total_assignment_count=Count("assignments", distinct=True),
        )
        .order_by("full_name")
    )

    if only_active:
        supervisors = supervisors.filter(is_active=True)
    if q:
        supervisors = supervisors.filter(Q(full_name__icontains=q) | Q(national_id__icontains=q))
    if gender in ("boys", "girls"):
        supervisors = supervisors.filter(gender=gender)
    if sector_id:
        supervisors = supervisors.filter(sector_id=sector_id)

    sectors = Sector.objects.filter(is_active=True).order_by("name")

    return render(
        request,
        "visits/admin_assignments_overview.html",
        {
            "rows": supervisors,
            "q": q,
            "gender": gender,
            "sector_id": sector_id,
            "only_active": only_active,
            "sectors": sectors,
            "kpi_supervisors": supervisors.count(),
            "kpi_assignments": Assignment.objects.filter(is_active=True, school__is_active=True).count(),
            "kpi_schools": School.objects.filter(is_active=True).count(),
        },
    )


@admin_only_view
def admin_supervisor_assignments_view(request: HttpRequest, supervisor_id: int) -> HttpResponse:
    supervisor = get_object_or_404(Supervisor.objects.select_related("sector"), id=supervisor_id)
    q = _cell_str(request.GET.get("q"))

    assigned_qs = (
        Assignment.objects.filter(
            supervisor=supervisor,
            is_active=True,
            school__isnull=False,
            school__is_active=True,
        )
        .select_related("school")
        .order_by("school__name")
    )

    if q:
        assigned_qs = assigned_qs.filter(
            Q(school__name__icontains=q) | Q(school__stat_code__icontains=q)
        )

    assigned_items = list(assigned_qs)
    school_ids = [a.school_id for a in assigned_items if a.school_id]
    principals_map = {p.school_id: p for p in Principal.objects.filter(school_id__in=school_ids)}

    rows = []
    for i, assignment in enumerate(assigned_items, start=1):
        school = assignment.school
        principal = principals_map.get(school.id) if school else None
        rows.append(
            {
                "index": i,
                "assignment": assignment,
                "school": school,
                "principal": principal,
                "is_active": assignment.is_active,
                "stat_code": getattr(school, "stat_code", "") or "—",
                "school_name": getattr(school, "name", "") or "—",
                "gender_label": _gender_label(getattr(school, "gender", "")),
                "principal_name": getattr(principal, "full_name", "") or "—",
                "principal_mobile": getattr(principal, "mobile", "") or "—",
            }
        )

    available_scope_qs = _supervisor_assignment_scope(supervisor)
    busy_school_ids = _globally_assigned_school_ids(exclude_supervisor_id=supervisor.id)
    available_schools_qs = available_scope_qs.exclude(id__in=busy_school_ids)

    if q:
        available_schools_qs = available_schools_qs.filter(
            Q(name__icontains=q) | Q(stat_code__icontains=q)
        )

    available_schools_qs = available_schools_qs.order_by("name")
    available_schools = list(available_schools_qs[:100])

    total_assigned_count = Assignment.objects.filter(
        supervisor=supervisor,
        is_active=True,
        school__isnull=False,
        school__is_active=True,
    ).count()
    total_available_count = available_scope_qs.exclude(id__in=busy_school_ids).count()

    return render(
        request,
        "visits/admin_supervisor_assignments.html",
        {
            "supervisor": supervisor,
            "supervisor_id": supervisor.id,
            "sup": supervisor,
            "assigned_items": assigned_items,
            "available_schools": available_schools,
            "q": q,
            "stats": {
                "assigned_count": total_assigned_count,
                "available_count": total_available_count,
                "search_count": len(available_schools),
            },
            # توافق مع أي قالب قديم
            "rows": rows,
            "kpi_total": total_assigned_count,
            "kpi_active": total_assigned_count,
            "kpi_inactive": 0,
            "kpi_available": total_available_count,
        },
    )


@admin_only_view
@require_POST
def admin_add_assignment_view(request: HttpRequest, supervisor_id: int) -> HttpResponse:
    supervisor = get_object_or_404(Supervisor, id=supervisor_id)

    next_url = (
        request.POST.get("next")
        or reverse("visits:admin_supervisor_assignments", args=[supervisor.id])
    )

    action = _cell_str(request.POST.get("action"))
    school_id = _safe_int(request.POST.get("school_id") or request.POST.get("school") or 0, default=0)
    new_school_name = _cell_str(request.POST.get("new_school_name") or request.POST.get("school_name") or request.POST.get("name"))
    new_stat_code = _cell_str(request.POST.get("new_stat_code") or request.POST.get("stat_code"))
    new_gender = _cell_str(request.POST.get("new_gender") or request.POST.get("gender")) or getattr(supervisor, "gender", "")

    school: School | None = None

    if action == "add_existing":
        if not school_id:
            messages.error(request, "يرجى اختيار مدرسة متاحة.")
            return redirect(next_url)
        school = get_object_or_404(School, id=school_id, is_active=True)
    elif action == "create_school":
        if not new_school_name:
            messages.error(request, "اسم المدرسة مطلوب.")
            return redirect(next_url)
        if not new_stat_code:
            messages.error(request, "الرقم الإحصائي مطلوب عند إضافة مدرسة جديدة.")
            return redirect(next_url)
        if new_gender not in ("boys", "girls"):
            new_gender = getattr(supervisor, "gender", "") or "boys"
        school = _find_existing_school_for_assignment(
            stat_code=new_stat_code,
            name=new_school_name,
            gender=new_gender,
            sector_id=getattr(supervisor, "sector_id", None),
        )
        if school is None:
            create_data = {
                "name": new_school_name,
                "stat_code": new_stat_code,
                "gender": new_gender,
                "is_active": True,
            }
            if getattr(supervisor, "sector_id", None):
                create_data["sector_id"] = supervisor.sector_id
            school = School.objects.create(**create_data)
    else:
        if school_id:
            school = get_object_or_404(School, id=school_id, is_active=True)
        elif new_school_name:
            if not new_stat_code:
                messages.error(request, "الرقم الإحصائي مطلوب عند إضافة مدرسة جديدة.")
                return redirect(next_url)
            if new_gender not in ("boys", "girls"):
                new_gender = getattr(supervisor, "gender", "") or "boys"
            school = _find_existing_school_for_assignment(
                stat_code=new_stat_code,
                name=new_school_name,
                gender=new_gender,
                sector_id=getattr(supervisor, "sector_id", None),
            )
            if school is None:
                create_data = {
                    "name": new_school_name,
                    "stat_code": new_stat_code,
                    "gender": new_gender,
                    "is_active": True,
                }
                if getattr(supervisor, "sector_id", None):
                    create_data["sector_id"] = supervisor.sector_id
                school = School.objects.create(**create_data)
        else:
            messages.error(request, "يرجى اختيار مدرسة أو إدخال بيانات مدرسة جديدة.")
            return redirect(next_url)

    if school is None:
        messages.error(request, "تعذر تحديد المدرسة المطلوبة.")
        return redirect(next_url)

    if getattr(supervisor, "gender", "") in ("boys", "girls") and getattr(school, "gender", "") not in ("", supervisor.gender):
        messages.error(request, "لا يمكن إسناد مدرسة بجنس مختلف عن جنس المشرف.")
        return redirect(next_url)

    if getattr(supervisor, "sector_id", None) and getattr(school, "sector_id", None) and school.sector_id != supervisor.sector_id:
        messages.error(request, "لا يمكن إسناد مدرسة من قطاع مختلف عن قطاع المشرف.")
        return redirect(next_url)

    existing_elsewhere = (
        Assignment.objects.filter(school=school, is_active=True, school__is_active=True)
        .exclude(supervisor=supervisor)
        .select_related("supervisor")
        .first()
    )
    if existing_elsewhere:
        messages.error(
            request,
            f"هذه المدرسة مسندة بالفعل إلى المشرف {existing_elsewhere.supervisor.full_name}.",
        )
        return redirect(next_url)

    assignment, created = Assignment.objects.get_or_create(
        supervisor=supervisor,
        school=school,
        defaults={"is_active": True},
    )

    if created:
        messages.success(request, "تمت إضافة الإسناد بنجاح.")
    else:
        if not assignment.is_active:
            assignment.is_active = True
            assignment.save(update_fields=["is_active"])
            messages.success(request, "تمت إعادة تفعيل الإسناد بنجاح.")
        else:
            messages.info(request, "هذا الإسناد موجود مسبقًا.")

    return redirect(next_url)


@admin_only_view
@require_POST
def admin_delete_assignment_view(request: HttpRequest, assignment_id: int) -> HttpResponse:
    assignment = get_object_or_404(Assignment.objects.select_related("supervisor"), id=assignment_id)

    next_url = (
        request.POST.get("next")
        or reverse("visits:admin_supervisor_assignments", args=[assignment.supervisor_id])
    )

    if hasattr(assignment, "is_active"):
        assignment.is_active = False
        assignment.save(update_fields=["is_active"])
        messages.success(request, "تم إلغاء الإسناد بنجاح.")
    else:
        assignment.delete()
        messages.success(request, "تم حذف الإسناد بنجاح.")

    return redirect(next_url)


# =============================================================================
# Admin sectors
# =============================================================================
@admin_only_view
def admin_sector_list_view(request: HttpRequest) -> HttpResponse:
    q = _cell_str(request.GET.get("q"))
    show_inactive = request.GET.get("inactive") == "1"

    sectors = (
        Sector.objects.annotate(
            schools_count=Count("schools", distinct=True),
            supervisors_count=Count("supervisors", distinct=True),
        )
        .order_by("name")
    )

    if q:
        sectors = sectors.filter(name__icontains=q)
    if not show_inactive:
        sectors = sectors.filter(is_active=True)

    return render(
        request,
        "visits/admin_sector_list.html",
        {
            "rows": sectors,
            "q": q,
            "show_inactive": show_inactive,
        },
    )


@admin_only_view
@require_POST
def admin_sector_save_view(request: HttpRequest) -> HttpResponse:
    sector_id = _safe_int(request.POST.get("sector_id") or 0, default=0)
    name = _cell_str(request.POST.get("name"))

    if not name:
        messages.error(request, "اسم القطاع مطلوب.")
        return redirect("visits:admin_sector_list")

    if sector_id:
        sector = get_object_or_404(Sector, id=sector_id)
        sector.name = name
        sector.save()
        messages.success(request, "تم تحديث بيانات القطاع بنجاح.")
    else:
        if Sector.objects.filter(name=name).exists():
            messages.error(request, "يوجد قطاع بهذا الاسم بالفعل.")
            return redirect("visits:admin_sector_list")
        Sector.objects.create(name=name, is_active=True)
        messages.success(request, "تمت إضافة القطاع بنجاح.")

    return redirect("visits:admin_sector_list")


@admin_only_view
@require_POST
def admin_sector_toggle_active_view(request: HttpRequest, sector_id: int) -> HttpResponse:
    sector = get_object_or_404(Sector, id=sector_id)
    sector.is_active = not bool(sector.is_active)
    sector.save(update_fields=["is_active"])

    if sector.is_active:
        messages.success(request, f"تم تفعيل القطاع «{sector.name}».")
    else:
        messages.success(request, f"تم تعطيل القطاع «{sector.name}».")
    return redirect("visits:admin_sector_list")


# =============================================================================
# Unlock request
# =============================================================================
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
        messages.info(request, "يوجد طلب فك اعتماد سابق لهذه الخطة بانتظار الإدارة.")
        return redirect(_plan_url(plan.week.week_no))

    req.status = UnlockRequest.STATUS_PENDING
    req.resolved_at = None
    req.save(update_fields=["status", "resolved_at"])

    plan.status = Plan.STATUS_UNLOCK_REQUESTED
    plan.save(update_fields=["status"])

    messages.success(request, "تم إرسال طلب فك الاعتماد.")
    return redirect(_plan_url(plan.week.week_no))


# =============================================================================
# Admin dashboard / detail / exports
# =============================================================================
@admin_only_view
def admin_dashboard_view(request: HttpRequest) -> HttpResponse:
    show_all = (request.GET.get("all") or "0").strip().lower() in ("1", "true", "yes")

    weeks_qs = _get_all_weeks_qs() if show_all else _get_active_weeks_qs()
    week_choices = _build_week_choices(active_only=(not show_all))
    default_week = weeks_qs.first()

    if not default_week:
        messages.warning(request, "لا يوجد أسابيع في جدول PlanWeek. أضف الأسابيع أولًا.")
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
        base_qs = base_qs.filter(Q(supervisor__full_name__icontains=q) | Q(supervisor__national_id__icontains=q))

    plans_all = list(base_qs)
    for p in plans_all:
        _inject_week_no(p)

    supervisor_ids = [p.supervisor_id for p in plans_all]
    assignment_counts = {}
    if supervisor_ids:
        assignment_counts = dict(
            Assignment.objects.filter(
                supervisor_id__in=supervisor_ids,
                is_active=True,
                school__is_active=True,
            )
            .values("supervisor_id")
            .annotate(c=Count("id"))
            .values_list("supervisor_id", "c")
        )

    kpi_total = len(plans_all)
    kpi_approved = sum(1 for p in plans_all if p.status == Plan.STATUS_APPROVED)
    kpi_drafts = sum(1 for p in plans_all if p.status == Plan.STATUS_DRAFT)
    kpi_unlock = sum(1 for p in plans_all if p.status == Plan.STATUS_UNLOCK_REQUESTED)
    kpi_filled = sum(1 for p in plans_all if p.is_fully_filled())
    kpi_not_filled = kpi_total - kpi_filled

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
        filled = sum(1 for wd, _ in WEEKDAYS if _day_is_filled(day_map.get(wd)))
        rows.append(
            {
                "plan": p,
                "sup": p.supervisor,
                "filled": filled,
                "is_full": filled == 5,
                "day_map": day_map,
                "assignment_count": assignment_counts.get(p.supervisor_id, 0),
            }
        )

    paginator = Paginator(rows, page_size)
    page_obj = paginator.get_page(page)
    chart_data = _build_chart_counts(week_obj)
    site_setting = _get_site_setting()
    _maintenance_is_active(site_setting, persist=True)

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
            "week_letter": WeeklyLetterLink.objects.filter(week=week_obj).first(),
            "site_setting": site_setting,
            **chart_data,
        },
    )


@admin_only_view
def admin_plan_detail_view(request: HttpRequest, plan_id: int) -> HttpResponse:
    plan = get_object_or_404(
        Plan.objects.select_related("supervisor", "week").prefetch_related("days__school"),
        id=plan_id,
    )
    _inject_week_no(plan)

    day_map = {d.weekday: d for d in plan.days.all().select_related("school")}
    filled = sum(1 for wd, _ in WEEKDAYS if _day_is_filled(day_map.get(wd)))
    day_dates = _build_day_dates_from_week(plan.week)
    visit_counts = _plan_visit_counts(plan)

    show_all = (request.GET.get("all") or "0").strip().lower() in ("1", "true", "yes")
    q = (request.GET.get("q") or "").strip()
    status_filter = (request.GET.get("status") or "all").strip()
    ps = _safe_int(request.GET.get("ps") or 10, default=10)
    page = _safe_int(request.GET.get("page") or 1, default=1)

    next_url = (request.GET.get("next") or "").strip()
    if not next_url:
        next_url = _admin_dashboard_url(
            plan.week.week_no,
            show_all=show_all,
            q=q,
            status=status_filter,
            ps=ps,
            page=page,
        )

    return render(
        request,
        "visits/admin_plan_detail.html",
        {
            "plan": plan,
            "sup": plan.supervisor,
            "weekdays": WEEKDAYS,
            "day_map": day_map,
            "filled": filled,
            "week": plan.week.week_no,
            "week_obj": plan.week,
            "day_dates": day_dates,
            "next_url": next_url,
            "back_to": "detail",
            "show_all": show_all,
            "q": q,
            "status": status_filter,
            "page_size": ps,
            "page_number": page,
            "assigned_count": visit_counts["assigned_count"],
            "visited_count": visit_counts["visited_count"],
            "remaining_count": visit_counts["remaining_count"],
        },
    )


@admin_only_view
def admin_plan_export_excel(request: HttpRequest, plan_id: int) -> HttpResponse:
    plan = get_object_or_404(
        Plan.objects.select_related("supervisor", "week").prefetch_related("days__school"),
        id=plan_id,
    )
    wb = _build_plan_excel_workbook(plan)
    filename = f"خطة_الأسبوع_{plan.week.week_no}_{_sup_nid_value(plan.supervisor)}.xlsx"
    return _excel_response(wb, filename)


@admin_only_view
def admin_export_supervisor_assignments_excel(request: HttpRequest, supervisor_id: int) -> HttpResponse:
    supervisor = get_object_or_404(Supervisor, id=supervisor_id, is_active=True)
    wb = _build_supervisor_assignments_excel_workbook(supervisor)
    filename = f"المدارس_المسندة_{_sup_nid_value(supervisor)}.xlsx"
    return _excel_response(wb, filename)


@admin_only_view
def admin_export_week_excel(request: HttpRequest) -> HttpResponse:
    show_all = (request.GET.get("all") or "0").strip().lower() in ("1", "true", "yes")
    weeks_qs = _get_all_weeks_qs() if show_all else _get_active_weeks_qs()
    default_week = weeks_qs.first()

    if not default_week:
        messages.warning(request, "لا يوجد أسابيع متاحة للتصدير.")
        return redirect("visits:admin_dashboard")

    week_no = _safe_int(request.GET.get("week") or default_week.week_no, default=default_week.week_no)
    if not _week_exists(week_no):
        week_no = default_week.week_no

    week_obj = _resolve_week_or_404(week_no, allow_inactive=show_all)
    plans = (
        Plan.objects.filter(week=week_obj)
        .select_related("supervisor", "week")
        .prefetch_related("days__school")
        .order_by("supervisor__full_name")
    )
    wb = _build_admin_week_excel_workbook(week_obj, plans)
    filename = f"خطط_الأسبوع_{week_obj.week_no}.xlsx"
    return _excel_response(wb, filename)


# =============================================================================
# Admin actions
# =============================================================================
@admin_only_view
@require_POST
def admin_plan_approve_view(request: HttpRequest, plan_id: int) -> HttpResponse:
    plan = get_object_or_404(Plan.objects.select_related("supervisor", "week").prefetch_related("days"), id=plan_id)
    _inject_week_no(plan)

    state = _parse_admin_action_state(request)
    return_url = _resolve_admin_return_url(
        request,
        plan=plan,
        default_week_no=plan.week.week_no,
        show_all=state["show_all"],
        q=state["q"],
        status_filter=state["status_filter"],
        ps=state["ps"],
        page=state["page"],
    )

    if plan.status == Plan.STATUS_APPROVED:
        msg = "الخطة معتمدة مسبقًا."
        ajax_response = _admin_json_response(request, plan=plan, message=msg, ok=True, http_status=200)
        if ajax_response:
            return ajax_response
        messages.info(request, msg)
        return redirect(return_url)

    if not plan.is_fully_filled():
        msg = "لا يمكن اعتماد الخطة قبل اكتمال جميع الأيام (5/5)."
        ajax_response = _admin_json_response(request, plan=plan, message=msg, ok=False, http_status=400, errors=[msg])
        if ajax_response:
            return ajax_response
        messages.warning(request, msg)
        return redirect(return_url)

    plan.status = Plan.STATUS_APPROVED
    plan.approved_at = timezone.now()
    plan.admin_note = None
    plan.save(update_fields=["status", "approved_at", "admin_note"])
    UnlockRequest.objects.filter(plan=plan).delete()

    _notify_supervisor(
        supervisor=plan.supervisor,
        plan=plan,
        notif_type=SupervisorNotification.TYPE_APPROVED,
        title="تم اعتماد خطتك",
        message=f"تم اعتماد خطة الأسبوع {plan.week.week_no} بنجاح.",
    )

    msg = "تم اعتماد الخطة."
    ajax_response = _admin_json_response(request, plan=plan, message=msg, ok=True, http_status=200)
    if ajax_response:
        return ajax_response

    messages.success(request, msg)
    return redirect(return_url)


@admin_only_view
@require_POST
def admin_plan_back_to_draft_view(request: HttpRequest, plan_id: int) -> HttpResponse:
    plan = get_object_or_404(Plan.objects.select_related("supervisor", "week").prefetch_related("days"), id=plan_id)
    _inject_week_no(plan)

    note = (request.POST.get("note") or "").strip()
    state = _parse_admin_action_state(request)
    return_url = _resolve_admin_return_url(
        request,
        plan=plan,
        default_week_no=plan.week.week_no,
        show_all=state["show_all"],
        q=state["q"],
        status_filter=state["status_filter"],
        ps=state["ps"],
        page=state["page"],
    )

    if plan.status == Plan.STATUS_DRAFT:
        msg = "الخطة بالفعل مسودة."
        ajax_response = _admin_json_response(request, plan=plan, message=msg, ok=True, http_status=200)
        if ajax_response:
            return ajax_response
        messages.info(request, msg)
        return redirect(return_url)

    plan.status = Plan.STATUS_DRAFT
    plan.approved_at = None
    plan.admin_note = note or None
    plan.save(update_fields=["status", "approved_at", "admin_note"])
    UnlockRequest.objects.filter(plan=plan).delete()

    _notify_supervisor(
        supervisor=plan.supervisor,
        plan=plan,
        notif_type=SupervisorNotification.TYPE_RETURNED,
        title="تمت إعادة الخطة للمراجعة",
        message=f"تمت إعادة خطة الأسبوع {plan.week.week_no} للمراجعة." + (f" الملاحظة: {note}" if note else ""),
    )

    msg = "تم إرجاع الخطة إلى مسودة."
    ajax_response = _admin_json_response(request, plan=plan, message=msg, ok=True, http_status=200)
    if ajax_response:
        return ajax_response

    messages.success(request, msg)
    return redirect(return_url)


@admin_only_view
@require_POST
def admin_send_notification_view(request: HttpRequest, plan_id: int) -> HttpResponse:
    plan = get_object_or_404(Plan.objects.select_related("supervisor", "week").prefetch_related("days"), id=plan_id)
    _inject_week_no(plan)

    note = (request.POST.get("note") or "").strip()
    title = (request.POST.get("title") or "تنبيه إداري").strip()

    state = _parse_admin_action_state(request)
    return_url = _resolve_admin_return_url(
        request,
        plan=plan,
        default_week_no=plan.week.week_no,
        show_all=state["show_all"],
        q=state["q"],
        status_filter=state["status_filter"],
        ps=state["ps"],
        page=state["page"],
    )

    if not note:
        msg = "نص التنبيه مطلوب."
        ajax_response = _admin_json_response(request, plan=plan, message=msg, ok=False, http_status=400, errors=[msg])
        if ajax_response:
            return ajax_response
        messages.warning(request, msg)
        return redirect(return_url)

    _notify_supervisor(
        supervisor=plan.supervisor,
        plan=plan,
        notif_type=SupervisorNotification.TYPE_ADMIN_ALERT,
        title=title,
        message=note,
    )

    msg = "تم إرسال التنبيه إلى المشرف."
    ajax_response = _admin_json_response(request, plan=plan, message=msg, ok=True, http_status=200)
    if ajax_response:
        return ajax_response

    messages.success(request, msg)
    return redirect(return_url)


@admin_only_view
@require_POST
def admin_send_notification_all_view(request: HttpRequest) -> HttpResponse:
    show_all = (request.POST.get("all") or request.GET.get("all") or "0").strip().lower() in ("1", "true", "yes")

    weeks_qs = _get_all_weeks_qs() if show_all else _get_active_weeks_qs()
    default_week = weeks_qs.first()

    if not default_week:
        messages.warning(request, "لا يوجد أسابيع متاحة.")
        return redirect("visits:admin_dashboard")

    week_no = _safe_int(request.POST.get("week") or request.GET.get("week") or default_week.week_no, default=default_week.week_no)
    if not _week_exists(week_no):
        week_no = default_week.week_no

    week_obj = _resolve_week_or_404(week_no, allow_inactive=show_all)
    q = (request.POST.get("q") or request.GET.get("q") or "").strip()
    status_filter = (request.POST.get("status") or request.GET.get("status") or "all").strip().lower()
    ps = _safe_int(request.POST.get("ps") or request.GET.get("ps") or 10, default=10)
    page = _safe_int(request.POST.get("page") or request.GET.get("page") or 1, default=1)
    title = (request.POST.get("title") or "تنبيه إداري عام").strip()
    note = (request.POST.get("note") or "").strip()

    return_url = _admin_dashboard_url(
        week_obj.week_no,
        show_all=show_all,
        q=q,
        status=status_filter,
        ps=ps,
        page=page,
    )

    if not note:
        msg = "نص التنبيه العام مطلوب."
        if _is_ajax(request):
            return JsonResponse({"ok": False, "message": msg, "errors": [msg]}, status=400)
        messages.warning(request, msg)
        return redirect(return_url)

    plans_qs = Plan.objects.filter(week=week_obj).select_related("supervisor").order_by("supervisor__full_name")
    if q:
        plans_qs = plans_qs.filter(Q(supervisor__full_name__icontains=q) | Q(supervisor__national_id__icontains=q))

    plans = list(plans_qs)

    if status_filter == "approved":
        plans = [p for p in plans if p.status == Plan.STATUS_APPROVED]
    elif status_filter == "draft":
        plans = [p for p in plans if p.status == Plan.STATUS_DRAFT]
    elif status_filter == "unlock":
        plans = [p for p in plans if p.status == Plan.STATUS_UNLOCK_REQUESTED]
    elif status_filter == "not_full":
        plans = [p for p in plans if not p.is_fully_filled()]

    sent_count = 0
    touched_supervisor_ids: set[int] = set()

    for plan in plans:
        sup = plan.supervisor
        if not sup or not getattr(sup, "is_active", False):
            continue
        if sup.id in touched_supervisor_ids:
            continue
        _notify_supervisor(
            supervisor=sup,
            plan=plan,
            notif_type=SupervisorNotification.TYPE_ADMIN_ALERT,
            title=title,
            message=note,
        )
        touched_supervisor_ids.add(sup.id)
        sent_count += 1

    msg = f"تم إرسال التنبيه العام إلى {sent_count} مشرف/ة."
    if _is_ajax(request):
        return JsonResponse({"ok": True, "message": msg, "sent_count": sent_count, "week_no": week_obj.week_no}, status=200)

    messages.success(request, msg)
    return redirect(return_url)


@admin_only_view
@require_POST
def admin_unlock_approve_view(request: HttpRequest, plan_id: int) -> HttpResponse:
    plan = get_object_or_404(Plan.objects.select_related("supervisor", "week").prefetch_related("days"), id=plan_id)
    _inject_week_no(plan)

    state = _parse_admin_action_state(request)
    return_url = _resolve_admin_return_url(
        request,
        plan=plan,
        default_week_no=plan.week.week_no,
        show_all=state["show_all"],
        q=state["q"],
        status_filter=state["status_filter"],
        ps=state["ps"],
        page=state["page"],
    )

    unlock = get_object_or_404(UnlockRequest, plan=plan)

    if unlock.status != UnlockRequest.STATUS_PENDING:
        msg = "تمت معالجة طلب فك الاعتماد سابقًا."
        ajax_response = _admin_json_response(request, plan=plan, message=msg, ok=False, http_status=400, errors=[msg])
        if ajax_response:
            return ajax_response
        messages.warning(request, msg)
        return redirect(return_url)

    unlock.status = UnlockRequest.STATUS_APPROVED
    unlock.resolved_at = timezone.now()
    unlock.save(update_fields=["status", "resolved_at"])

    plan.status = Plan.STATUS_DRAFT
    plan.approved_at = None
    plan.save(update_fields=["status", "approved_at"])

    _notify_supervisor(
        supervisor=plan.supervisor,
        plan=plan,
        notif_type=SupervisorNotification.TYPE_UNLOCK_APPROVED,
        title="تمت الموافقة على فك اعتماد الخطة",
        message=f"تمت الموافقة على فك اعتماد خطة الأسبوع {plan.week.week_no}. يمكنك التعديل الآن.",
    )

    msg = "تمت الموافقة على فك الاعتماد وإرجاع الخطة إلى مسودة."
    ajax_response = _admin_json_response(request, plan=plan, message=msg, ok=True, http_status=200)
    if ajax_response:
        return ajax_response

    messages.success(request, msg)
    return redirect(return_url)


@admin_only_view
@require_POST
def admin_unlock_reject_view(request: HttpRequest, plan_id: int) -> HttpResponse:
    plan = get_object_or_404(Plan.objects.select_related("supervisor", "week").prefetch_related("days"), id=plan_id)
    _inject_week_no(plan)

    state = _parse_admin_action_state(request)
    return_url = _resolve_admin_return_url(
        request,
        plan=plan,
        default_week_no=plan.week.week_no,
        show_all=state["show_all"],
        q=state["q"],
        status_filter=state["status_filter"],
        ps=state["ps"],
        page=state["page"],
    )

    unlock = get_object_or_404(UnlockRequest, plan=plan)
    note = (request.POST.get("note") or "").strip()

    if unlock.status != UnlockRequest.STATUS_PENDING:
        msg = "تمت معالجة طلب فك الاعتماد سابقًا."
        ajax_response = _admin_json_response(request, plan=plan, message=msg, ok=False, http_status=400, errors=[msg])
        if ajax_response:
            return ajax_response
        messages.warning(request, msg)
        return redirect(return_url)

    unlock.status = UnlockRequest.STATUS_REJECTED
    unlock.resolved_at = timezone.now()
    unlock.save(update_fields=["status", "resolved_at"])

    plan.status = Plan.STATUS_APPROVED
    plan.save(update_fields=["status"])

    _notify_supervisor(
        supervisor=plan.supervisor,
        plan=plan,
        notif_type=SupervisorNotification.TYPE_UNLOCK_REJECTED,
        title="تم رفض طلب فك اعتماد الخطة",
        message=f"تم رفض طلب فك اعتماد خطة الأسبوع {plan.week.week_no}." + (f" الملاحظة: {note}" if note else ""),
    )

    msg = "تم رفض طلب فك الاعتماد."
    ajax_response = _admin_json_response(request, plan=plan, message=msg, ok=True, http_status=200)
    if ajax_response:
        return ajax_response

    messages.success(request, msg)
    return redirect(return_url)


# =============================================================================
# Weekly letter links admin
# =============================================================================
@admin_only_view
def weekly_letter_links_list_view(request: HttpRequest) -> HttpResponse:
    rows = WeeklyLetterLink.objects.select_related("week").order_by("week__week_no")
    form = WeeklyLetterLinkForm()
    return render(
        request,
        "visits/weekly_letter_links_list.html",
        {
            "rows": rows,
            "form": form,
            "edit_obj": None,
        },
    )


@admin_only_view
def weekly_letter_link_create_view(request: HttpRequest) -> HttpResponse:
    if request.method == "POST":
        form = WeeklyLetterLinkForm(request.POST)
        if form.is_valid():
            obj = form.save(commit=False)
            if obj.is_active:
                WeeklyLetterLink.objects.filter(
                    week=obj.week,
                    is_active=True,
                ).exclude(id=obj.id).update(is_active=False)
            obj.save()
            messages.success(request, "تم حفظ رابط الأسبوع بنجاح.")
            return redirect("visits:weekly_letter_links_list")

        messages.error(request, "تعذر حفظ رابط الأسبوع. تحقق من الحقول.")
    else:
        form = WeeklyLetterLinkForm()

    return render(
        request,
        "visits/weekly_letter_link_form.html",
        {
            "form": form,
            "edit_obj": None,
        },
    )


@admin_only_view
def weekly_letter_link_edit_view(request: HttpRequest, pk: int) -> HttpResponse:
    obj = get_object_or_404(WeeklyLetterLink, pk=pk)

    if request.method == "POST":
        form = WeeklyLetterLinkForm(request.POST, instance=obj)
        if form.is_valid():
            updated = form.save(commit=False)
            if updated.is_active:
                WeeklyLetterLink.objects.filter(week=updated.week, is_active=True).exclude(id=updated.id).update(is_active=False)
            updated.save()
            messages.success(request, "تم تحديث رابط الأسبوع بنجاح.")
            return redirect("visits:weekly_letter_links_list")
        messages.error(request, "تعذر تحديث الرابط. تحقق من الحقول.")
    else:
        form = WeeklyLetterLinkForm(instance=obj)

    rows = WeeklyLetterLink.objects.select_related("week").order_by("week__week_no")
    return render(
        request,
        "visits/weekly_letter_links.html",
        {
            "rows": rows,
            "form": form,
            "edit_obj": obj,
        },
    )


@admin_only_view
@require_POST
def weekly_letter_link_delete_view(request: HttpRequest, pk: int) -> HttpResponse:
    obj = get_object_or_404(WeeklyLetterLink, pk=pk)
    obj.delete()
    messages.success(request, "تم حذف الرابط.")
    return redirect("visits:weekly_letter_links_list")