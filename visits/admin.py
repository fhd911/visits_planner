from __future__ import annotations

from datetime import timedelta

from django.contrib import admin
from django.contrib.admin import RelatedOnlyFieldListFilter
from django.core.exceptions import ValidationError
from django.db import models
from django.utils import timezone
from django.utils.html import format_html

from .models import (
    Assignment,
    Plan,
    PlanDay,
    PlanWeek,
    Principal,
    School,
    Supervisor,
    SupervisorNotification,
    UnlockRequest,
    WeeklyLetterLink,
)

admin.site.site_header = "لوحة إدارة الزيارات"
admin.site.site_title = "إدارة الزيارات"
admin.site.index_title = "الإدارة الرئيسية"


def _badge(text: str, *, bg: str, color: str, weight: int = 800) -> str:
    return format_html(
        '<span style="padding:3px 10px;border-radius:999px;'
        'background:{};color:{};font-weight:{};">{}</span>',
        bg,
        color,
        weight,
        text,
    )


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


@admin.register(PlanWeek)
class PlanWeekAdmin(admin.ModelAdmin):
    list_display = ("week_no", "start_sunday", "end_thursday", "title", "break_badge")
    list_filter = ("is_break",)
    search_fields = ("week_no", "title")
    ordering = ("week_no",)

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
    fields = ("weekday", "school", "visit_type", "no_visit_reason", "note")
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
        if obj.status == Plan.STATUS_APPROVED:
            return _badge("✅ معتمدة", bg="#dcfce7", color="#166534", weight=700)
        if obj.status == Plan.STATUS_UNLOCK_REQUESTED:
            return _badge("🔓 فك اعتماد", bg="#ffedd5", color="#7c2d12", weight=700)
        return _badge("📝 مسودة", bg="#dbeafe", color="#0c4a6e", weight=700)

    @admin.display(description="الامتلاء")
    def filled_badge(self, obj: Plan):
        try:
            count = obj.days.filter(
                models.Q(school__isnull=False) | models.Q(visit_type=PlanDay.VISIT_NONE)
            ).count()
        except Exception:
            count = 0

        if count == 5:
            return _badge("5/5", bg="#dcfce7", color="#166534")
        return _badge(f"{count}/5", bg="#fff7ed", color="#7c2d12")


@admin.register(PlanDay)
class PlanDayAdmin(admin.ModelAdmin):
    list_display = ("plan", "weekday", "school", "visit_type", "no_visit_reason")
    list_filter = ("weekday", "visit_type", "no_visit_reason")
    search_fields = (
        "plan__supervisor__full_name",
        "plan__supervisor__national_id",
        "school__name",
        "school__stat_code",
    )
    autocomplete_fields = ("plan", "school")
    ordering = ("plan__week__week_no", "weekday")


@admin.register(UnlockRequest)
class UnlockRequestAdmin(admin.ModelAdmin):
    list_display = ("plan", "status_badge", "created_at", "resolved_at")
    list_filter = ("status", "created_at")
    search_fields = ("plan__supervisor__full_name", "plan__supervisor__national_id")
    autocomplete_fields = ("plan",)
    readonly_fields = ("created_at",)

    fieldsets = (
        (None, {"fields": ("plan", "status")}),
        ("التاريخ", {"fields": ("created_at", "resolved_at")}),
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

        super().save_model(request, obj, form, change)

        if old_status == obj.status:
            return

        plan = obj.plan

        if obj.status == UnlockRequest.STATUS_PENDING:
            if plan.status != Plan.STATUS_UNLOCK_REQUESTED:
                plan.status = Plan.STATUS_UNLOCK_REQUESTED
                plan.save(update_fields=["status"])

            if obj.resolved_at is not None:
                obj.resolved_at = None
                obj.save(update_fields=["resolved_at"])

        elif obj.status == UnlockRequest.STATUS_APPROVED:
            if plan.status != Plan.STATUS_DRAFT:
                plan.status = Plan.STATUS_DRAFT
                plan.save(update_fields=["status"])

            if obj.resolved_at is None:
                obj.resolved_at = timezone.now()
                obj.save(update_fields=["resolved_at"])

        elif obj.status == UnlockRequest.STATUS_REJECTED:
            if plan.status != Plan.STATUS_APPROVED:
                plan.status = Plan.STATUS_APPROVED
                plan.save(update_fields=["status"])

            if obj.resolved_at is None:
                obj.resolved_at = timezone.now()
                obj.save(update_fields=["resolved_at"])


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