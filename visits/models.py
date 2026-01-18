from __future__ import annotations

from django.db import models
from django.core.validators import MinValueValidator, MaxValueValidator
from django.utils import timezone


# =========================
# المدارس
# =========================
class School(models.Model):
    class Meta:
        verbose_name = "مدرسة"
        verbose_name_plural = "المدارس"
        ordering = ["name"]
        indexes = [
            models.Index(fields=["stat_code"]),
            models.Index(fields=["is_active"]),
        ]

    GENDER_CHOICES = [("boys", "بنين"), ("girls", "بنات")]

    stat_code = models.CharField("الرقم الإحصائي", max_length=32, unique=True)
    name = models.CharField("اسم المدرسة", max_length=255)
    gender = models.CharField("النوع", max_length=10, choices=GENDER_CHOICES)
    is_active = models.BooleanField("نشطة", default=True)

    def __str__(self) -> str:
        return f"{self.name} ({self.stat_code})"


# =========================
# مدير المدرسة
# =========================
class Principal(models.Model):
    class Meta:
        verbose_name = "مدير مدرسة"
        verbose_name_plural = "مديرو المدارس"
        ordering = ["full_name"]

    school = models.OneToOneField(
        School,
        on_delete=models.CASCADE,
        related_name="principal",
        verbose_name="المدرسة",
    )
    full_name = models.CharField("اسم المدير", max_length=255)
    mobile = models.CharField("الجوال", max_length=20)

    def __str__(self) -> str:
        return f"{self.full_name} — {self.school.name}"


# =========================
# المشرف
# =========================
class Supervisor(models.Model):
    class Meta:
        verbose_name = "مشرف"
        verbose_name_plural = "المشرفون"
        ordering = ["full_name"]
        indexes = [
            models.Index(fields=["national_id"]),
            models.Index(fields=["is_active"]),
        ]

    national_id = models.CharField("السجل المدني", max_length=20, unique=True)
    full_name = models.CharField("اسم المشرف", max_length=255)

    mobile = models.CharField("جوال المشرف", max_length=20, blank=True, null=True)
    is_active = models.BooleanField("نشط", default=True)

    def __str__(self) -> str:
        return self.full_name

    @staticmethod
    def _digits(v: str) -> str:
        return "".join(ch for ch in (v or "") if ch.isdigit())

    def mobile_last4(self) -> str | None:
        d = self._digits(self.mobile or "")
        return d[-4:] if len(d) >= 4 else None


# =========================
# إسناد مدرسة لمشرف
# =========================
class Assignment(models.Model):
    class Meta:
        verbose_name = "إسناد مدرسة"
        verbose_name_plural = "إسنادات المدارس"
        constraints = [
            models.UniqueConstraint(fields=["supervisor", "school"], name="uniq_assignment_sup_school"),
        ]
        ordering = ["supervisor__full_name", "school__name"]

    supervisor = models.ForeignKey(
        Supervisor,
        on_delete=models.CASCADE,
        related_name="assignments",
        verbose_name="المشرف",
    )
    school = models.ForeignKey(
        School,
        on_delete=models.CASCADE,
        related_name="assignments",
        verbose_name="المدرسة",
    )
    is_active = models.BooleanField("نشط", default=True)

    def __str__(self) -> str:
        return f"{self.supervisor} ← {self.school}"


# =========================
# ✅ جدول الأسابيع (يدوي بالكامل)
# =========================
class PlanWeek(models.Model):
    class Meta:
        verbose_name = "أسبوع خطة"
        verbose_name_plural = "أسابيع الخطة"
        ordering = ["week_no"]
        indexes = [
            models.Index(fields=["week_no"]),
            models.Index(fields=["is_break"]),
        ]

    week_no = models.PositiveSmallIntegerField(
        "رقم الأسبوع",
        unique=True,
        validators=[MinValueValidator(1), MaxValueValidator(60)],
    )

    start_sunday = models.DateField("بداية الأسبوع (الأحد)")
    title = models.CharField("ملاحظة/اسم", max_length=120, blank=True, null=True)
    is_break = models.BooleanField("إجازة/توقف؟", default=False)

    def __str__(self) -> str:
        extra = " (إجازة)" if self.is_break else ""
        t = f" — {self.title}" if self.title else ""
        return f"الأسبوع {self.week_no} يبدأ {self.start_sunday}{t}{extra}"


# =========================
# ✅ خطة أسبوعية
# =========================
class Plan(models.Model):
    class Meta:
        verbose_name = "خطة أسبوعية"
        verbose_name_plural = "الخطط الأسبوعية"
        constraints = [
            models.UniqueConstraint(fields=["supervisor", "week"], name="uniq_plan_supervisor_week"),
        ]
        ordering = ["-id"]
        indexes = [
            models.Index(fields=["status"]),
            models.Index(fields=["week"]),
        ]

    STATUS_DRAFT = "draft"
    STATUS_APPROVED = "approved"
    STATUS_UNLOCK_REQUESTED = "unlock"

    STATUS_CHOICES = [
        (STATUS_DRAFT, "مسودة"),
        (STATUS_APPROVED, "معتمدة"),
        (STATUS_UNLOCK_REQUESTED, "طلب فك اعتماد"),
    ]

    supervisor = models.ForeignKey(
        Supervisor,
        on_delete=models.CASCADE,
        related_name="plans",
        verbose_name="المشرف",
    )

    # ✅ إلزامي
    week = models.ForeignKey(
        PlanWeek,
        on_delete=models.CASCADE,
        related_name="plans",
        verbose_name="الأسبوع",
    )

    status = models.CharField("الحالة", max_length=20, choices=STATUS_CHOICES, default=STATUS_DRAFT)
    saved_at = models.DateTimeField("وقت الحفظ", null=True, blank=True)
    approved_at = models.DateTimeField("وقت الاعتماد", null=True, blank=True)

    def __str__(self) -> str:
        return f"{self.supervisor} — أسبوع {self.week.week_no}"

    def is_fully_filled(self) -> bool:
        """
        ✅ التعديل:
        اليوم يعتبر "معبّى" إذا:
        - فيه مدرسة (school_id موجود)
        أو
        - نوع الزيارة = بدون زيارة (none)
        """
        needed = {0, 1, 2, 3, 4}
        filled = {
            d.weekday
            for d in self.days.all()
            if (d.school_id is not None) or (d.visit_type == PlanDay.VISIT_NONE)
        }
        return needed.issubset(filled)


# =========================
# ✅ تفاصيل أيام الخطة
# =========================
class PlanDay(models.Model):
    class Meta:
        verbose_name = "يوم خطة"
        verbose_name_plural = "أيام الخطة"
        constraints = [
            models.UniqueConstraint(fields=["plan", "weekday"], name="uniq_planday_plan_weekday"),
        ]
        ordering = ["weekday"]
        indexes = [
            models.Index(fields=["weekday"]),
            models.Index(fields=["visit_type"]),
        ]

    WEEKDAY_CHOICES = [
        (0, "الأحد"),
        (1, "الإثنين"),
        (2, "الثلاثاء"),
        (3, "الأربعاء"),
        (4, "الخميس"),
    ]

    # ✅ أنواع الزيارة (مع "بدون زيارة")
    VISIT_IN = "in"
    VISIT_REMOTE = "remote"
    VISIT_NONE = "none"

    VISIT_CHOICES = [
        (VISIT_IN, "حضوري"),
        (VISIT_REMOTE, "عن بعد"),
        (VISIT_NONE, "بدون زيارة مدرسية"),
    ]

    # ✅ أسباب جاهزة عند عدم وجود زيارة
    REASON_MEETING = "meeting"
    REASON_TRAINING = "training"
    REASON_VISIT = "event"
    REASON_OFFICE = "office"
    REASON_OTHER = "other"

    NO_VISIT_REASON_CHOICES = [
        (REASON_MEETING, "اجتماع"),
        (REASON_TRAINING, "تدريب"),
        (REASON_VISIT, "لقاء/فعالية"),
        (REASON_OFFICE, "عمل مكتبي"),
        (REASON_OTHER, "أخرى"),
    ]

    plan = models.ForeignKey(
        Plan,
        related_name="days",
        on_delete=models.CASCADE,
        verbose_name="الخطة",
    )
    weekday = models.PositiveSmallIntegerField("اليوم", choices=WEEKDAY_CHOICES)

    # ✅ nullable ممتاز
    school = models.ForeignKey(
        School,
        on_delete=models.PROTECT,
        null=True,
        blank=True,
        verbose_name="المدرسة",
    )

    visit_type = models.CharField("نوع اليوم", max_length=10, choices=VISIT_CHOICES, default=VISIT_IN)

    # ✅ سبب عدم الزيارة (يظهر فقط إذا visit_type = none)
    no_visit_reason = models.CharField(
        "سبب عدم الزيارة",
        max_length=20,
        choices=NO_VISIT_REASON_CHOICES,
        null=True,
        blank=True,
    )

    # ✅ ملاحظة اختيارية (مفيدة خصوصًا مع "أخرى")
    note = models.CharField("ملاحظة", max_length=120, null=True, blank=True)

    def __str__(self) -> str:
        return f"{self.plan} — {self.get_weekday_display()} — {self.school or '—'}"

    def clean(self):
        """
        (اختياري) يمكنك إضافة تحقق هنا لاحقًا.
        تركته بسيط عشان ما يسبب لك مشاكل أثناء الاستيراد/التعديل.
        """
        super().clean()


# =========================
# طلب فك اعتماد
# =========================
class UnlockRequest(models.Model):
    class Meta:
        verbose_name = "طلب فك اعتماد"
        verbose_name_plural = "طلبات فك الاعتماد"
        constraints = [
            models.UniqueConstraint(fields=["plan"], name="uniq_unlock_request_plan"),
        ]
        ordering = ["-created_at"]
        indexes = [
            models.Index(fields=["status"]),
        ]

    STATUS_PENDING = "pending"
    STATUS_APPROVED = "approved"
    STATUS_REJECTED = "rejected"

    STATUS_CHOICES = [
        (STATUS_PENDING, "معلّق"),
        (STATUS_APPROVED, "مقبول"),
        (STATUS_REJECTED, "مرفوض"),
    ]

    plan = models.OneToOneField(
        Plan,
        on_delete=models.CASCADE,
        related_name="unlock_request",
        verbose_name="الخطة",
    )
    status = models.CharField("الحالة", max_length=20, choices=STATUS_CHOICES, default=STATUS_PENDING)
    created_at = models.DateTimeField("تاريخ الطلب", auto_now_add=True)
    resolved_at = models.DateTimeField("تاريخ المعالجة", null=True, blank=True)

    def __str__(self) -> str:
        return f"فك اعتماد — {self.plan} — {self.get_status_display()}"

    def mark_resolved(self):
        if not self.resolved_at:
            self.resolved_at = timezone.now()

    def approve(self):
        self.status = self.STATUS_APPROVED
        self.mark_resolved()
        self.save(update_fields=["status", "resolved_at"])

    def reject(self):
        self.status = self.STATUS_REJECTED
        self.mark_resolved()
        self.save(update_fields=["status", "resolved_at"])
