from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from django.conf import settings
from django.core.mail import EmailMultiAlternatives
from django.db import transaction
from django.urls import reverse
from django.utils import timezone

from visits.models import (
    EmailNotificationLog,
    EmailNotificationPreference,
    Plan,
    PlanClosedDay,
    PlanDay,
    PlanWeek,
    Supervisor,
)

WEEKDAYS = [
    (0, "الأحد"),
    (1, "الإثنين"),
    (2, "الثلاثاء"),
    (3, "الأربعاء"),
    (4, "الخميس"),
]


@dataclass(frozen=True)
class EmailSendResult:
    sent: bool
    skipped: bool
    log_id: int | None
    message: str


EVENT_TO_PREF_FIELD = {
    EmailNotificationLog.EVENT_PLAN_APPROVED: "plan_approved",
    EmailNotificationLog.EVENT_PLAN_RETURNED: "plan_returned",
    EmailNotificationLog.EVENT_UNLOCK_APPROVED: "unlock_result",
    EmailNotificationLog.EVENT_UNLOCK_REJECTED: "unlock_result",
    EmailNotificationLog.EVENT_ADMIN_ALERT: "admin_alert",
    EmailNotificationLog.EVENT_CONTROL_FOLLOWUP: "control_followup",
    EmailNotificationLog.EVENT_INCOMPLETE_REMINDER: "incomplete_reminder",
    EmailNotificationLog.EVENT_WEEKLY_SUMMARY: "weekly_summary",
    EmailNotificationLog.EVENT_UNLOCK_REQUESTED: "unlock_result",
}


def get_or_create_email_preferences(supervisor: Supervisor) -> EmailNotificationPreference:
    pref, _created = EmailNotificationPreference.objects.get_or_create(supervisor=supervisor)
    return pref


def _supervisor_email(supervisor: Supervisor) -> str:
    return (getattr(supervisor, "email", None) or "").strip().lower()


def _week_label(plan: Plan | None) -> str:
    if not plan or not getattr(plan, "week", None):
        return "الأسبوع الحالي"
    week = plan.week
    semester = getattr(week, "semester", None)
    year = getattr(week, "academic_year", None) or getattr(semester, "academic_year", None)
    semester_title = getattr(semester, "title", None) or ""
    year_name = getattr(year, "name", None) or ""
    semester_week_no = getattr(week, "semester_week_no", None) or getattr(week, "week_no", None)
    parts = []
    if semester_title:
        parts.append(str(semester_title))
    if year_name:
        parts.append(str(year_name))
    if semester_week_no:
        parts.append(f"الأسبوع {semester_week_no}")
    return " — ".join(parts) or f"الأسبوع {getattr(week, 'week_no', '')}"


def _absolute_url(request, route_name: str, *args, **kwargs) -> str:
    try:
        path = reverse(route_name, args=args, kwargs=kwargs)
    except Exception:
        return ""
    if request is not None:
        try:
            return request.build_absolute_uri(path)
        except Exception:
            pass
    base_url = getattr(settings, "SITE_BASE_URL", "") or getattr(settings, "VISITS_SITE_BASE_URL", "")
    return f"{base_url.rstrip('/')}{path}" if base_url else path


def _log_email(
    *,
    supervisor: Supervisor,
    event_type: str,
    subject: str,
    body: str,
    plan: Plan | None = None,
    recipient: str | None = None,
    status: str,
    error_message: str | None = None,
) -> EmailNotificationLog:
    return EmailNotificationLog.objects.create(
        supervisor=supervisor,
        plan=plan,
        event_type=event_type,
        recipient_email=recipient or _supervisor_email(supervisor) or None,
        subject=subject[:250],
        body_preview=(body or "")[:1000],
        status=status,
        error_message=error_message,
        sent_at=timezone.now() if status == EmailNotificationLog.STATUS_SENT else None,
    )


def _preference_allows(supervisor: Supervisor, event_type: str, *, force: bool = False) -> bool:
    if force:
        return True
    if not getattr(supervisor, "email_notifications_enabled", True):
        return False
    pref = get_or_create_email_preferences(supervisor)
    field = EVENT_TO_PREF_FIELD.get(event_type)
    if not field:
        return True
    return bool(getattr(pref, field, True))


def send_supervisor_email(
    *,
    supervisor: Supervisor,
    event_type: str,
    subject: str,
    body: str,
    plan: Plan | None = None,
    request=None,
    force: bool = False,
) -> EmailSendResult:
    """
    يرسل رسالة بريدية للمشرف مع تسجيل نتيجة الإرسال.
    لا يرفع استثناءً للواجهة؛ كل فشل يسجل في EmailNotificationLog.
    """
    recipient = _supervisor_email(supervisor)

    if not recipient:
        log = _log_email(
            supervisor=supervisor,
            plan=plan,
            event_type=event_type,
            subject=subject,
            body=body,
            recipient=None,
            status=EmailNotificationLog.STATUS_SKIPPED,
            error_message="لا يوجد بريد إلكتروني مسجل للمشرف.",
        )
        return EmailSendResult(False, True, log.id, "لا يوجد بريد إلكتروني مسجل.")

    require_verified = bool(getattr(settings, "VISITS_REQUIRE_VERIFIED_EMAIL", False))
    if require_verified and not getattr(supervisor, "email_verified", False):
        log = _log_email(
            supervisor=supervisor,
            plan=plan,
            event_type=event_type,
            subject=subject,
            body=body,
            recipient=recipient,
            status=EmailNotificationLog.STATUS_SKIPPED,
            error_message="البريد غير موثق.",
        )
        return EmailSendResult(False, True, log.id, "البريد غير موثق.")

    if not _preference_allows(supervisor, event_type, force=force):
        log = _log_email(
            supervisor=supervisor,
            plan=plan,
            event_type=event_type,
            subject=subject,
            body=body,
            recipient=recipient,
            status=EmailNotificationLog.STATUS_SKIPPED,
            error_message="التنبيه غير مفعل حسب تفضيلات المشرف.",
        )
        return EmailSendResult(False, True, log.id, "التنبيه غير مفعل.")

    try:
        from_email = getattr(settings, "DEFAULT_FROM_EMAIL", None) or getattr(settings, "SERVER_EMAIL", None)
        msg = EmailMultiAlternatives(subject=subject, body=body, from_email=from_email, to=[recipient])
        msg.send(fail_silently=False)
        log = _log_email(
            supervisor=supervisor,
            plan=plan,
            event_type=event_type,
            subject=subject,
            body=body,
            recipient=recipient,
            status=EmailNotificationLog.STATUS_SENT,
        )
        return EmailSendResult(True, False, log.id, "تم الإرسال.")
    except Exception as exc:
        log = _log_email(
            supervisor=supervisor,
            plan=plan,
            event_type=event_type,
            subject=subject,
            body=body,
            recipient=recipient,
            status=EmailNotificationLog.STATUS_FAILED,
            error_message=str(exc),
        )
        return EmailSendResult(False, False, log.id, f"فشل الإرسال: {exc}")


def _format_message(title: str, lines: list[str], url: str | None = None) -> str:
    clean_lines = [line for line in lines if line]
    body = [title, "", *clean_lines]
    if url:
        body.extend(["", f"رابط المتابعة: {url}"])
    body.extend(["", "هذه رسالة آلية من منصة خطة الزيارات."])
    return "\n".join(body)


def send_plan_approved_email(plan: Plan, request=None) -> EmailSendResult:
    url = _absolute_url(request, "visits:supervisor_previous_plan_detail", plan.id)
    subject = f"تم اعتماد خطة {_week_label(plan)}"
    body = _format_message(
        "تم اعتماد خطتك الأسبوعية.",
        [
            f"الخطة: {_week_label(plan)}",
            f"وقت الاعتماد: {timezone.localtime(plan.approved_at or timezone.now()).strftime('%Y-%m-%d %H:%M')}",
        ],
        url,
    )
    return send_supervisor_email(
        supervisor=plan.supervisor,
        plan=plan,
        event_type=EmailNotificationLog.EVENT_PLAN_APPROVED,
        subject=subject,
        body=body,
        request=request,
    )


def send_plan_returned_email(plan: Plan, note: str | None = None, request=None) -> EmailSendResult:
    url = _absolute_url(request, "visits:plan")
    subject = f"تمت إعادة خطة {_week_label(plan)} للتعديل"
    body = _format_message(
        "تمت إعادة خطتك الأسبوعية للتعديل.",
        [
            f"الخطة: {_week_label(plan)}",
            f"ملاحظة الإدارة: {note}" if note else "يرجى مراجعة الخطة واستكمال التعديل المطلوب.",
        ],
        url,
    )
    return send_supervisor_email(
        supervisor=plan.supervisor,
        plan=plan,
        event_type=EmailNotificationLog.EVENT_PLAN_RETURNED,
        subject=subject,
        body=body,
        request=request,
    )


def send_unlock_request_received_email(plan: Plan, request=None) -> EmailSendResult:
    subject = f"تم استلام طلب فك اعتماد خطة {_week_label(plan)}"
    body = _format_message(
        "تم استلام طلب فك اعتماد الخطة.",
        [
            f"الخطة: {_week_label(plan)}",
            "سيتم إشعارك بعد معالجة الطلب من الإدارة.",
        ],
        _absolute_url(request, "visits:plan"),
    )
    return send_supervisor_email(
        supervisor=plan.supervisor,
        plan=plan,
        event_type=EmailNotificationLog.EVENT_UNLOCK_REQUESTED,
        subject=subject,
        body=body,
        request=request,
    )


def send_unlock_approved_email(plan: Plan, request=None) -> EmailSendResult:
    subject = f"تم قبول طلب فك اعتماد خطة {_week_label(plan)}"
    body = _format_message(
        "تم قبول طلب فك الاعتماد.",
        [
            f"الخطة: {_week_label(plan)}",
            "يمكنك الآن تعديل الخطة ثم حفظها واعتمادها من جديد.",
        ],
        _absolute_url(request, "visits:plan"),
    )
    return send_supervisor_email(
        supervisor=plan.supervisor,
        plan=plan,
        event_type=EmailNotificationLog.EVENT_UNLOCK_APPROVED,
        subject=subject,
        body=body,
        request=request,
    )


def send_unlock_rejected_email(plan: Plan, note: str | None = None, request=None) -> EmailSendResult:
    subject = f"تم رفض طلب فك اعتماد خطة {_week_label(plan)}"
    body = _format_message(
        "تم رفض طلب فك الاعتماد.",
        [
            f"الخطة: {_week_label(plan)}",
            f"ملاحظة الإدارة: {note}" if note else "بقيت الخطة على حالتها المعتمدة.",
        ],
        _absolute_url(request, "visits:supervisor_previous_plan_detail", plan.id),
    )
    return send_supervisor_email(
        supervisor=plan.supervisor,
        plan=plan,
        event_type=EmailNotificationLog.EVENT_UNLOCK_REJECTED,
        subject=subject,
        body=body,
        request=request,
    )


def send_admin_alert_email(plan: Plan, title: str, message: str, request=None) -> EmailSendResult:
    subject = title or f"تنبيه إداري بشأن خطة {_week_label(plan)}"
    body = _format_message(
        subject,
        [
            f"الخطة: {_week_label(plan)}",
            message or "يرجى مراجعة التنبيه في المنصة.",
        ],
        _absolute_url(request, "visits:notifications"),
    )
    return send_supervisor_email(
        supervisor=plan.supervisor,
        plan=plan,
        event_type=EmailNotificationLog.EVENT_ADMIN_ALERT,
        subject=subject,
        body=body,
        request=request,
    )


def send_control_followup_email(supervisor: Supervisor, title: str, message: str, plan: Plan | None = None, request=None) -> EmailSendResult:
    subject = title or "ملاحظة رقابية تحتاج متابعة"
    body = _format_message(
        subject,
        [message or "توجد ملاحظة رقابية تحتاج إفادة أو متابعة."],
        _absolute_url(request, "visits:supervisor_control_followups"),
    )
    return send_supervisor_email(
        supervisor=supervisor,
        plan=plan,
        event_type=EmailNotificationLog.EVENT_CONTROL_FOLLOWUP,
        subject=subject,
        body=body,
        request=request,
    )


def _day_is_filled(day: PlanDay | None) -> bool:
    if not day:
        return False
    if day.school_id:
        return True
    if day.visit_type == PlanDay.VISIT_NONE and day.no_visit_reason:
        return True
    return False


def plan_work_completion(plan: Plan) -> dict[str, int]:
    closed = set(
        PlanClosedDay.objects.filter(
            week=plan.week,
            is_active=True,
            count_as_completed=True,
        ).values_list("weekday", flat=True)
    )
    required_weekdays = [wd for wd, _name in WEEKDAYS if wd not in closed]
    days = {d.weekday: d for d in plan.days.all()}
    filled = sum(1 for wd in required_weekdays if _day_is_filled(days.get(wd)))
    return {
        "required": len(required_weekdays),
        "filled": min(filled, len(required_weekdays)),
        "missing": max(len(required_weekdays) - filled, 0),
        "closed": len(closed),
    }


def send_incomplete_plan_email(plan: Plan, *, request=None) -> EmailSendResult:
    stats = plan_work_completion(plan)
    subject = f"تذكير باستكمال خطة {_week_label(plan)}"
    body = _format_message(
        "تذكير باستكمال الخطة الأسبوعية.",
        [
            f"الخطة: {_week_label(plan)}",
            f"المكتمل: {stats['filled']} / {stats['required']} أيام عمل",
            f"الأيام المتبقية: {stats['missing']}",
            "يرجى استكمال الخطة قبل نهاية المهلة.",
        ],
        _absolute_url(request, "visits:plan"),
    )
    return send_supervisor_email(
        supervisor=plan.supervisor,
        plan=plan,
        event_type=EmailNotificationLog.EVENT_INCOMPLETE_REMINDER,
        subject=subject,
        body=body,
        request=request,
    )


def send_incomplete_plan_reminders(*, week: PlanWeek | None = None, request=None) -> dict[str, int]:
    """
    يرسل تذكيرات للخطط غير المكتملة في الأسبوع الحالي المفتوح.
    يستخدم من لوحة الإدارة عند الحاجة.
    """
    if week is None:
        week = PlanWeek.objects.filter(is_current=True, is_open_for_supervisors=True, is_break=False).first()
    if not week:
        return {"sent": 0, "skipped": 0, "failed": 0, "total": 0}

    active_supervisors = Supervisor.objects.filter(is_active=True).order_by("full_name")
    totals = {"sent": 0, "skipped": 0, "failed": 0, "total": 0}

    with transaction.atomic():
        for supervisor in active_supervisors:
            plan, _created = Plan.objects.get_or_create(supervisor=supervisor, week=week)
            stats = plan_work_completion(plan)
            if stats["required"] <= 0 or stats["missing"] <= 0 or plan.status == Plan.STATUS_APPROVED:
                continue
            totals["total"] += 1
            result = send_incomplete_plan_email(plan, request=request)
            if result.sent:
                totals["sent"] += 1
            elif result.skipped:
                totals["skipped"] += 1
            else:
                totals["failed"] += 1
    return totals
