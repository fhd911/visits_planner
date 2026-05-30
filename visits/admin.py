from __future__ import annotations

from datetime import timedelta
from typing import Any

from django.contrib import admin
from django.contrib.admin import RelatedOnlyFieldListFilter
from django.core.exceptions import ValidationError
from django.db import models
from django.utils import timezone
from django.utils.html import format_html

from .models import (
    AcademicYear,
    Assignment,
    Semester,
    Plan,
    PlanClosedDay,
    PlanDay,
    PlanWeek,
    Principal,
    School,
    Supervisor,
    EmailNotificationLog,
    EmailNotificationPreference,
    SupervisorNotification,
    UnlockRequest,
    WeeklyLetterLink,
)

try:
    from .models import ControlFollowUp
except Exception:  # يسمح للملف بالعمل حتى قبل تطبيق موديل المتابعة الرقابية
    ControlFollowUp = None  # type: ignore[assignment]

try:
    from .models import ControlFollowUpAction
except Exception:
    ControlFollowUpAction = None  # type: ignore[assignment]


admin.site.site_header = "لوحة إدارة الزيارات"
admin.site.site_title = "إدارة الزيارات"
admin.site.index_title = "الإدارة الرئيسية"


# =============================================================================
# Helpers
# =============================================================================
def _badge(text: str, *, bg: str, color: str, weight: int = 800) -> str:
    return format_html(
        '<span style="display:inline-flex;align-items:center;gap:4px;'
        'padding:3px 10px;border-radius:999px;'
        'background:{};color:{};font-weight:{};white-space:nowrap;">{}</span>',
        bg,
        color,
        weight,
        text,
    )


def _model_has_field(model: type[models.Model], field_name: str) -> bool:
    try:
        model._meta.get_field(field_name)
        return True
    except Exception:
        return False


def _append_existing_fields(model: type[models.Model], names: list[str]) -> list[str]:
    return [name for name in names if _model_has_field(model, name)]


def _short_text(value: str | None, limit: int = 55) -> str:
    value = (value or "").strip()
    if not value:
        return "—"
    return value if len(value) <= limit else f"{value[:limit]}…"


def _safe_update_model_fields(obj: models.Model, updates: dict[str, Any]) -> None:
    """تحديث حقول موجودة فقط؛ مفيد عند اختلاف نسخة models.py أثناء التطوير."""
    update_fields: list[str] = []
    for field_name, value in updates.items():
        if _model_has_field(obj.__class__, field_name):
            setattr(obj, field_name, value)
            update_fields.append(field_name)
    if update_fields:
        obj.save(update_fields=update_fields)


def _current_user_or_none(request):
    user = getattr(request, "user", None)
    if user and getattr(user, "is_authenticated", False):
        return user
    return None


def _notify_unlock_decision(request, unlock_request: UnlockRequest, *, approved: bool) -> None:
    """يرسل إشعارًا للمشرف عند معالجة طلب فك الاعتماد من Django Admin."""
    plan = unlock_request.plan
    supervisor = plan.supervisor

    if approved:
        title = "تم قبول طلب فك اعتماد الخطة"
        message = (
            f"تمت الموافقة على فك اعتماد خطة الأسبوع {plan.week.week_no}. "
            "يمكنكم تعديل الخطة ثم إعادة اعتمادها بعد الانتهاء."
        )
        notif_type = SupervisorNotification.TYPE_UNLOCK_APPROVED
    else:
        title = "تم رفض طلب فك اعتماد الخطة"
        message = (
            f"تم رفض طلب فك اعتماد خطة الأسبوع {plan.week.week_no}. "
            "تبقى الخطة معتمدة ما لم تردكم معالجة أخرى من الإدارة."
        )
        notif_type = SupervisorNotification.TYPE_UNLOCK_REJECTED

    try:
        SupervisorNotification.objects.create(
            supervisor=supervisor,
            plan=plan,
            notif_type=notif_type,
            title=title,
            message=message,
        )
    except Exception:
        # لا نجعل فشل الإشعار يعطل حفظ القرار الإداري.
        pass


def _sync_unlock_control_followup(request, unlock_request: UnlockRequest, *, decision: str) -> None:
    """يربط معالجة طلب فك الاعتماد بسجل المتابعة الرقابية إن كان الموديل موجودًا."""
    if ControlFollowUp is None:
        return

    plan = unlock_request.plan
    supervisor = plan.supervisor
    week = plan.week
    unique_key = f"unlock_request:{supervisor_id_part(supervisor)}:{week.week_no}:{plan.id}"

    title = f"طلب فك اعتماد خطة الأسبوع {week.week_no}"
    description = f"طلب فك اعتماد مقدم من المشرف/ {supervisor.full_name} لخطة الأسبوع {week.week_no}."

    defaults = {
        "issue_type": getattr(ControlFollowUp, "ISSUE_UNLOCK_REQUEST", "unlock_request"),
        "status": getattr(ControlFollowUp, "STATUS_OPEN", "open"),
        "supervisor": supervisor,
        "plan": plan,
        "week": week,
        "title": title,
        "description": description,
        "created_by": _current_user_or_none(request),
    }

    try:
        followup, _created = ControlFollowUp.objects.get_or_create(
            unique_key=unique_key,
            defaults=defaults,
        )

        if decision == "approved":
            followup.accept_processing("تم قبول فك الاعتماد من لوحة Django Admin.")
        elif decision == "rejected":
            followup.close_administratively("تم رفض فك الاعتماد من لوحة Django Admin.")
        else:
            followup.status = getattr(ControlFollowUp, "STATUS_OPEN", "open")
            followup.save(update_fields=["status", "updated_at"])

        if ControlFollowUpAction is not None:
            try:
                ControlFollowUpAction.objects.create(
                    followup=followup,
                    action_type="unlock_approved" if decision == "approved" else "unlock_rejected" if decision == "rejected" else "unlock_pending",
                    from_status="",
                    to_status=followup.status,
                    actor_user=_current_user_or_none(request),
                    note="تمت مزامنة قرار فك الاعتماد من Django Admin.",
                )
            except Exception:
                pass
    except Exception:
        pass


def supervisor_id_part(supervisor: Supervisor) -> str:
    try:
        return str(supervisor.pk or supervisor.national_id or "0")
    except Exception:
        return "0"


# =============================================================================
# Supervisors / Schools / Assignments
# =============================================================================
@admin.register(Supervisor)
class SupervisorAdmin(admin.ModelAdmin):
    list_display = ("full_name", "national_id", "mobile", "last4", "is_active")
    search_fields = ("full_name", "national_id", "mobile")
    list_filter = ("is_active",)
    ordering = ("full_name",)

    @admin.display(description="آخر 4")
    def last4(self, obj: Supervisor):
        try:
            return obj.mobile_last4() or "—"
        except Exception:
            return "—"


@admin.register(School)
class SchoolAdmin(admin.ModelAdmin):
    list_display = ("name", "stat_code", "gender", "is_active")
    search_fields = ("name", "stat_code")
    list_filter = ("is_active", "gender")
    ordering = ("name",)


@admin.register(Principal)
class PrincipalAdmin(admin.ModelAdmin):
    list_display = ("full_name", "school", "mobile")
    search_fields = ("full_name", "mobile", "school__name", "school__stat_code")
    autocomplete_fields = ("school",)


@admin.register(Assignment)
class AssignmentAdmin(admin.ModelAdmin):
    list_display = ("supervisor", "school", "is_active")
    list_filter = ("is_active",)
    search_fields = (
        "supervisor__full_name",
        "supervisor__national_id",
        "school__name",
        "school__stat_code",
    )
    autocomplete_fields = ("supervisor", "school")
    ordering = ("supervisor__full_name", "school__name")



# =============================================================================
# Academic years / semesters / closed days
# =============================================================================
@admin.register(AcademicYear)
class AcademicYearAdmin(admin.ModelAdmin):
    list_display = ("name", "starts_at", "ends_at", "is_current", "is_active")
    list_filter = ("is_current", "is_active")
    search_fields = ("name",)
    ordering = ("-starts_at", "-id")


@admin.register(Semester)
class SemesterAdmin(admin.ModelAdmin):
    list_display = (
        "academic_year",
        "title",
        "number",
        "starts_at",
        "ends_at",
        "weeks_count",
        "is_current",
        "is_open",
    )
    list_filter = ("academic_year", "number", "is_current", "is_open")
    search_fields = ("title", "academic_year__name")
    autocomplete_fields = ("academic_year",)
    ordering = ("academic_year", "number")


@admin.register(PlanClosedDay)
class PlanClosedDayAdmin(admin.ModelAdmin):
    list_display = (
        "week",
        "weekday",
        "date",
        "reason_type",
        "reason_title",
        "count_as_completed",
        "is_active",
    )
    list_filter = ("reason_type", "is_active", "count_as_completed", "week")
    search_fields = ("reason_title", "week__title", "week__week_no")
    autocomplete_fields = ("week",)
    ordering = ("week__week_no", "weekday")


class PlanClosedDayInline(admin.TabularInline):
    model = PlanClosedDay
    extra = 0
    fields = ("weekday", "date", "reason_type", "reason_title", "count_as_completed", "is_active")
    ordering = ("weekday",)


# =============================================================================
# Weeks / Plans
# =============================================================================
@admin.register(PlanWeek)
class PlanWeekAdmin(admin.ModelAdmin):
    list_display = (
        "week_no",
        "display_label_admin",
        "academic_year",
        "semester",
        "semester_week_no",
        "start_sunday",
        "end_thursday",
        "title",
        "current_badge",
        "open_badge",
        "break_badge",
    )
    list_filter = ("academic_year", "semester", "is_current", "is_open_for_supervisors", "is_break")
    search_fields = ("week_no", "title", "semester__title", "academic_year__name")
    autocomplete_fields = ("academic_year", "semester")
    ordering = ("week_no",)
    inlines = [PlanClosedDayInline]

    @admin.display(description="العرض للمشرف")
    def display_label_admin(self, obj: PlanWeek):
        return getattr(obj, "display_label", f"الأسبوع {obj.week_no}")

    @admin.display(description="الأسبوع الحالي")
    def current_badge(self, obj: PlanWeek):
        if getattr(obj, "is_current", False):
            return _badge("✅ الحالي", bg="#dcfce7", color="#166534")
        return _badge("—", bg="#f1f5f9", color="#475569", weight=600)

    @admin.display(description="مفتوح للمشرفين")
    def open_badge(self, obj: PlanWeek):
        if getattr(obj, "is_open_for_supervisors", False):
            return _badge("مفتوح", bg="#dbeafe", color="#1d4ed8")
        return _badge("مغلق", bg="#f1f5f9", color="#475569", weight=600)

    @admin.display(description="نهاية الأسبوع (الخميس)")
    def end_thursday(self, obj: PlanWeek):
        try:
            return obj.start_sunday + timedelta(days=4)
        except Exception:
            return "—"

    @admin.display(description="نوع الأسبوع")
    def break_badge(self, obj: PlanWeek):
        if obj.is_break:
            return _badge("⛔ إجازة", bg="#fee2e2", color="#991b1b")
        return _badge("✅ فعّال", bg="#dcfce7", color="#166534")


class PlanDayInline(admin.TabularInline):
    model = PlanDay
    extra = 0
    fields = ("weekday", "school", "visit_type", "no_visit_reason", "note", "visited")
    autocomplete_fields = ("school",)
    ordering = ("weekday",)


@admin.register(Plan)
class PlanAdmin(admin.ModelAdmin):
    list_display = (
        "supervisor",
        "week_no_display",
        "week_start_display",
        "week_break_hint",
        "status_badge",
        "filled_badge",
        "saved_at",
        "approved_at",
        "unlocked_at_display",
    )

    list_filter = (
        "status",
        ("week", RelatedOnlyFieldListFilter),
    )

    search_fields = (
        "supervisor__full_name",
        "supervisor__national_id",
        "week__week_no",
        "week__title",
    )

    autocomplete_fields = ("supervisor", "week")
    ordering = ("-week__week_no", "-id")
    inlines = [PlanDayInline]

    def save_model(self, request, obj: Plan, form, change):
        if obj.week and getattr(obj.week, "is_break", False):
            raise ValidationError("لا يمكن إنشاء/تعديل خطة على أسبوع مُحدد كإجازة.")
        super().save_model(request, obj, form, change)

    @admin.display(description="الأسبوع", ordering="week__week_no")
    def week_no_display(self, obj: Plan):
        return obj.week.week_no if obj.week_id else "—"

    @admin.display(description="بداية الأسبوع", ordering="week__start_sunday")
    def week_start_display(self, obj: Plan):
        return obj.week.start_sunday if obj.week_id else "—"

    @admin.display(description="ملاحظة")
    def week_break_hint(self, obj: Plan):
        if not obj.week_id:
            return "—"
        if obj.week.is_break:
            return _badge("⛔ أسبوع إجازة", bg="#fee2e2", color="#991b1b")
        return _badge("✅ أسبوع فعّال", bg="#dcfce7", color="#166534")

    @admin.display(description="الحالة")
    def status_badge(self, obj: Plan):
        status_unlocked = getattr(Plan, "STATUS_UNLOCKED", "unlocked")

        if obj.status == Plan.STATUS_APPROVED:
            return _badge("✅ معتمدة", bg="#dcfce7", color="#166534", weight=700)

        if obj.status == Plan.STATUS_UNLOCK_REQUESTED:
            return _badge("🔓 طلب فك اعتماد", bg="#ffedd5", color="#7c2d12", weight=700)

        if obj.status == status_unlocked:
            return _badge("🛠 مفكوكة للتعديل", bg="#e0f2fe", color="#075985", weight=700)

        return _badge("📝 مسودة", bg="#dbeafe", color="#0c4a6e", weight=700)

    @admin.display(description="الامتلاء")
    def filled_badge(self, obj: Plan):
        try:
            filled_weekdays = set(
                obj.days.filter(
                    models.Q(school__isnull=False) | models.Q(visit_type=PlanDay.VISIT_NONE)
                ).values_list("weekday", flat=True)
            )
            if obj.week_id:
                filled_weekdays.update(
                    PlanClosedDay.objects.filter(
                        week=obj.week,
                        is_active=True,
                        count_as_completed=True,
                    ).values_list("weekday", flat=True)
                )
            count = len(filled_weekdays)
        except Exception:
            count = 0

        if count >= 5:
            return _badge("5/5", bg="#dcfce7", color="#166534")
        return _badge(f"{count}/5", bg="#fff7ed", color="#7c2d12")

    @admin.display(description="فك الاعتماد")
    def unlocked_at_display(self, obj: Plan):
        value = getattr(obj, "unlocked_at", None)
        return value or "—"


@admin.register(PlanDay)
class PlanDayAdmin(admin.ModelAdmin):
    list_display = ("plan", "weekday", "school", "visit_type", "no_visit_reason", "visited")
    list_filter = ("weekday", "visit_type", "no_visit_reason", "visited")
    search_fields = (
        "plan__supervisor__full_name",
        "plan__supervisor__national_id",
        "school__name",
        "school__stat_code",
    )
    autocomplete_fields = ("plan", "school")
    ordering = ("plan__week__week_no", "weekday")


# =============================================================================
# Unlock requests
# =============================================================================
@admin.register(UnlockRequest)
class UnlockRequestAdmin(admin.ModelAdmin):
    list_display = (
        "plan",
        "supervisor_display",
        "week_display",
        "status_badge",
        "reason_short",
        "created_at",
        "resolved_at",
    )
    list_filter = ("status", "created_at")
    search_fields = (
        "plan__supervisor__full_name",
        "plan__supervisor__national_id",
        "plan__week__week_no",
    )
    autocomplete_fields = ("plan",)

    def get_readonly_fields(self, request, obj=None):
        readonly = ["created_at"]
        readonly += _append_existing_fields(UnlockRequest, ["resolved_at"])
        return tuple(dict.fromkeys(readonly))

    def get_fieldsets(self, request, obj=None):
        main_fields = ["plan", "status"]
        main_fields += _append_existing_fields(UnlockRequest, ["reason", "request_reason", "admin_note"])

        processed_fields = _append_existing_fields(
            UnlockRequest,
            ["resolved_by", "processed_by", "handled_by"],
        )

        fieldsets = [(None, {"fields": tuple(main_fields)})]
        date_fields = _append_existing_fields(UnlockRequest, ["created_at", "resolved_at"])
        if processed_fields or date_fields:
            fieldsets.append(("المعالجة والتاريخ", {"fields": tuple(processed_fields + date_fields)}))
        return tuple(fieldsets)

    @admin.display(description="المشرف")
    def supervisor_display(self, obj: UnlockRequest):
        try:
            return obj.plan.supervisor.full_name
        except Exception:
            return "—"

    @admin.display(description="الأسبوع")
    def week_display(self, obj: UnlockRequest):
        try:
            return obj.plan.week.week_no
        except Exception:
            return "—"

    @admin.display(description="سبب الطلب")
    def reason_short(self, obj: UnlockRequest):
        return _short_text(
            getattr(obj, "reason", None)
            or getattr(obj, "request_reason", None)
            or getattr(obj, "admin_note", None)
        )

    @admin.display(description="الحالة")
    def status_badge(self, obj: UnlockRequest):
        if obj.status == UnlockRequest.STATUS_PENDING:
            return _badge("⏳ معلّق", bg="#fef9c3", color="#854d0e")
        if obj.status == UnlockRequest.STATUS_APPROVED:
            return _badge("✅ مقبول", bg="#dcfce7", color="#166534")
        return _badge("❌ مرفوض", bg="#fee2e2", color="#991b1b")

    def save_model(self, request, obj: UnlockRequest, form, change):
        old_status = None
        if change and obj.pk:
            old_status = UnlockRequest.objects.filter(pk=obj.pk).values_list("status", flat=True).first()

        # تعبئة المعالج إن وُجد الحقل في نسخة models.py.
        user = _current_user_or_none(request)
        if obj.status in (UnlockRequest.STATUS_APPROVED, UnlockRequest.STATUS_REJECTED):
            if _model_has_field(UnlockRequest, "resolved_by") and not getattr(obj, "resolved_by_id", None):
                obj.resolved_by = user
            if _model_has_field(UnlockRequest, "processed_by") and not getattr(obj, "processed_by_id", None):
                obj.processed_by = user
            if _model_has_field(UnlockRequest, "handled_by") and not getattr(obj, "handled_by_id", None):
                obj.handled_by = user

        super().save_model(request, obj, form, change)

        if old_status == obj.status:
            return

        plan = obj.plan
        status_unlocked = getattr(Plan, "STATUS_UNLOCKED", "unlocked")

        if obj.status == UnlockRequest.STATUS_PENDING:
            if plan.status != Plan.STATUS_UNLOCK_REQUESTED:
                plan.status = Plan.STATUS_UNLOCK_REQUESTED
                plan.save(update_fields=["status"])

            if obj.resolved_at is not None:
                obj.resolved_at = None
                obj.save(update_fields=["resolved_at"])

            _sync_unlock_control_followup(request, obj, decision="pending")
            return

        if obj.status == UnlockRequest.STATUS_APPROVED:
            # لا ترجع الخطة إلى مسودة عادية إذا كانت نسخة Plan تدعم حالة مفكوكة للتعديل.
            plan_updates = {
                "status": status_unlocked if hasattr(Plan, "STATUS_UNLOCKED") else Plan.STATUS_DRAFT,
                "approved_at": None,
                "unlocked_at": timezone.now(),
                "unlocked_by": user,
                "unlock_reason": getattr(obj, "reason", None) or getattr(obj, "request_reason", None) or "",
            }

            existing_updates = {
                name: value
                for name, value in plan_updates.items()
                if name == "status" or _model_has_field(Plan, name)
            }

            for field_name, value in existing_updates.items():
                setattr(plan, field_name, value)
            plan.save(update_fields=list(existing_updates.keys()))

            if obj.resolved_at is None:
                obj.resolved_at = timezone.now()
                obj.save(update_fields=["resolved_at"])

            _notify_unlock_decision(request, obj, approved=True)
            _sync_unlock_control_followup(request, obj, decision="approved")
            return

        if obj.status == UnlockRequest.STATUS_REJECTED:
            if plan.status != Plan.STATUS_APPROVED:
                plan.status = Plan.STATUS_APPROVED
                plan.save(update_fields=["status"])

            if obj.resolved_at is None:
                obj.resolved_at = timezone.now()
                obj.save(update_fields=["resolved_at"])

            _notify_unlock_decision(request, obj, approved=False)
            _sync_unlock_control_followup(request, obj, decision="rejected")


# =============================================================================
# Notifications / Letters
# =============================================================================
@admin.register(SupervisorNotification)
class SupervisorNotificationAdmin(admin.ModelAdmin):
    list_display = ("supervisor", "notif_type", "title", "is_read", "created_at")
    list_filter = ("notif_type", "is_read", "created_at")
    search_fields = (
        "supervisor__full_name",
        "supervisor__national_id",
        "title",
        "message",
    )
    autocomplete_fields = ("supervisor", "plan")
    readonly_fields = ("created_at",)
    ordering = ("-created_at",)


@admin.register(WeeklyLetterLink)
class WeeklyLetterLinkAdmin(admin.ModelAdmin):
    list_display = (
        "week",
        "week_start",
        "title",
        "is_active_badge",
        "open_link",
        "updated_at",
    )
    list_filter = ("is_active", "week__is_break")
    search_fields = ("title", "note", "week__week_no", "week__title", "drive_url")
    autocomplete_fields = ("week",)
    ordering = ("week__week_no",)
    readonly_fields = ("created_at", "updated_at")

    fieldsets = (
        (None, {"fields": ("week", "title", "drive_url", "note", "is_active")}),
        ("التوقيت", {"fields": ("created_at", "updated_at")}),
    )

    @admin.display(description="بداية الأسبوع")
    def week_start(self, obj: WeeklyLetterLink):
        try:
            return obj.week.start_sunday
        except Exception:
            return "—"

    @admin.display(description="الحالة")
    def is_active_badge(self, obj: WeeklyLetterLink):
        if obj.is_active:
            return _badge("✅ نشط", bg="#dcfce7", color="#166534")
        return _badge("⛔ غير نشط", bg="#e5e7eb", color="#374151")

    @admin.display(description="الرابط")
    def open_link(self, obj: WeeklyLetterLink):
        if not obj.drive_url:
            return "—"
        return format_html(
            '<a href="{}" target="_blank" rel="noopener noreferrer">فتح الرابط</a>',
            obj.drive_url,
        )


# =============================================================================
# Control follow-ups
# =============================================================================
if ControlFollowUp is not None:

    if ControlFollowUpAction is not None:

        class ControlFollowUpActionInline(admin.TabularInline):
            model = ControlFollowUpAction
            extra = 0
            fields = (
                "action_type",
                "from_status",
                "to_status",
                "actor_user",
                "actor_supervisor",
                "note",
                "created_at",
            )
            readonly_fields = ("created_at",)
            autocomplete_fields = ("actor_user", "actor_supervisor")
            can_delete = False

    else:
        ControlFollowUpActionInline = None  # type: ignore[assignment]

    @admin.register(ControlFollowUp)
    class ControlFollowUpAdmin(admin.ModelAdmin):
        list_display = (
            "title",
            "issue_badge",
            "status_badge",
            "supervisor",
            "week",
            "notification_count",
            "last_notification_at",
            "updated_at",
        )
        list_filter = ("status", "issue_type", "week", "last_notification_at", "created_at")
        search_fields = (
            "title",
            "description",
            "admin_note",
            "supervisor_response",
            "admin_review_note",
            "supervisor__full_name",
            "supervisor__national_id",
        )
        readonly_fields = (
            "unique_key",
            "notification_count",
            "last_notification_at",
            "supervisor_response_at",
            "resolved_at",
            "closed_at",
            "created_at",
            "updated_at",
        )
        autocomplete_fields = ("supervisor", "plan", "week", "created_by")
        ordering = ("-updated_at", "-created_at")
        if ControlFollowUpActionInline is not None:
            inlines = [ControlFollowUpActionInline]

        @admin.display(description="نوع الحالة")
        def issue_badge(self, obj):
            label = obj.get_issue_type_display()
            if obj.issue_type == getattr(ControlFollowUp, "ISSUE_UNLOCK_REQUEST", "unlock_request"):
                return _badge(label, bg="#ffedd5", color="#7c2d12", weight=700)
            if obj.issue_type == getattr(ControlFollowUp, "ISSUE_INCOMPLETE_PLAN", "incomplete_plan"):
                return _badge(label, bg="#fef9c3", color="#854d0e", weight=700)
            return _badge(label, bg="#e0f2fe", color="#075985", weight=700)

        @admin.display(description="حالة المتابعة")
        def status_badge(self, obj):
            if obj.status == getattr(ControlFollowUp, "STATUS_OPEN", "open"):
                return _badge("مفتوحة", bg="#fef9c3", color="#854d0e", weight=700)
            if obj.status == getattr(ControlFollowUp, "STATUS_NOTIFIED", "notified"):
                return _badge("تم التنبيه", bg="#ffedd5", color="#7c2d12", weight=700)
            if obj.status == getattr(ControlFollowUp, "STATUS_PENDING_ADMIN", "pending_admin"):
                return _badge("بانتظار الإدارة", bg="#dbeafe", color="#0c4a6e", weight=700)
            if obj.status == getattr(ControlFollowUp, "STATUS_PROCESSED", "processed"):
                return _badge("تمت المعالجة", bg="#dcfce7", color="#166534", weight=700)
            return _badge("مغلقة", bg="#e5e7eb", color="#374151", weight=700)


if ControlFollowUpAction is not None:

    @admin.register(ControlFollowUpAction)
    class ControlFollowUpActionAdmin(admin.ModelAdmin):
        list_display = (
            "followup",
            "action_type",
            "from_status",
            "to_status",
            "actor_user",
            "actor_supervisor",
            "created_at",
        )
        list_filter = ("action_type", "created_at")
        search_fields = (
            "followup__title",
            "followup__supervisor__full_name",
            "followup__supervisor__national_id",
            "note",
        )
        readonly_fields = ("created_at",)
        autocomplete_fields = ("followup", "actor_user", "actor_supervisor")
        ordering = ("-created_at",)



@admin.register(EmailNotificationPreference)
class EmailNotificationPreferenceAdmin(admin.ModelAdmin):
    list_display = (
        "supervisor",
        "plan_approved",
        "plan_returned",
        "unlock_result",
        "admin_alert",
        "control_followup",
        "incomplete_reminder",
        "weekly_summary",
        "updated_at",
    )
    list_filter = (
        "plan_approved",
        "plan_returned",
        "unlock_result",
        "admin_alert",
        "control_followup",
        "incomplete_reminder",
        "weekly_summary",
    )
    search_fields = ("supervisor__full_name", "supervisor__national_id", "supervisor__email")
    autocomplete_fields = ("supervisor",)


@admin.register(EmailNotificationLog)
class EmailNotificationLogAdmin(admin.ModelAdmin):
    list_display = (
        "created_at",
        "supervisor",
        "event_type",
        "recipient_email",
        "status",
        "sent_at",
    )
    list_filter = ("event_type", "status", "created_at", "sent_at")
    search_fields = (
        "supervisor__full_name",
        "supervisor__national_id",
        "recipient_email",
        "subject",
        "error_message",
    )
    readonly_fields = (
        "supervisor",
        "plan",
        "event_type",
        "recipient_email",
        "subject",
        "body_preview",
        "status",
        "error_message",
        "sent_at",
        "created_at",
    )
    date_hierarchy = "created_at"
    autocomplete_fields = ("supervisor", "plan")

    def has_add_permission(self, request):
        return False

    def has_change_permission(self, request, obj=None):
        return False
