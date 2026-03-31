from __future__ import annotations

import random
from datetime import timedelta

from django.core.exceptions import ValidationError
from django.core.validators import MaxValueValidator, MinValueValidator
from django.db import models
from django.utils import timezone


# =========================
# Helpers داخلية
# =========================
def _digits(v: str) -> str:
    return "".join(ch for ch in (v or "") if ch.isdigit())


def _clean_text(v: str | None) -> str:
    return (v or "").strip()


# =========================
# إعدادات الموقع العامة
# =========================
class SiteSetting(models.Model):
    class Meta:
        verbose_name = "إعدادات الموقع"
        verbose_name_plural = "إعدادات الموقع"

    site_name = models.CharField("اسم الموقع", max_length=150, default="بوابة الزيارات")
    is_maintenance_mode = models.BooleanField("وضع الصيانة مفعّل", default=False)

    maintenance_message = models.TextField(
        "رسالة الصيانة",
        blank=True,
        null=True,
        default="الموقع مغلق مؤقتًا للصيانة، وسيعود العمل خلال وقت قريب.",
    )

    allow_admin_only = models.BooleanField(
        "السماح للإدارة فقط أثناء الصيانة",
        default=True,
        help_text="عند التفعيل: يمكن للإدارة فقط الدخول أثناء وضع الصيانة.",
    )

    expected_return_text = models.CharField(
        "وقت العودة المتوقع (نصي)",
        max_length=200,
        blank=True,
        null=True,
        help_text="مثال: اليوم الساعة 3 مساءً أو خلال ساعة تقريبًا",
    )

    maintenance_starts_at = models.DateTimeField(
        "بداية الصيانة",
        null=True,
        blank=True,
        help_text="تاريخ ووقت بدء الصيانة.",
    )

    maintenance_ends_at = models.DateTimeField(
        "نهاية الصيانة",
        null=True,
        blank=True,
        help_text="تاريخ ووقت انتهاء الصيانة وإعادة فتح الموقع.",
    )

    updated_at = models.DateTimeField("آخر تحديث", auto_now=True)
    created_at = models.DateTimeField("تاريخ الإنشاء", auto_now_add=True)

    def __str__(self) -> str:
        return self.site_name or "إعدادات الموقع"

    def clean(self):
        super().clean()
        self.site_name = _clean_text(self.site_name) or "بوابة الزيارات"
        self.maintenance_message = _clean_text(self.maintenance_message) or None
        self.expected_return_text = _clean_text(self.expected_return_text) or None

        errors = {}

        if self.maintenance_starts_at and self.maintenance_ends_at:
            if self.maintenance_ends_at <= self.maintenance_starts_at:
                errors["maintenance_ends_at"] = "يجب أن يكون وقت نهاية الصيانة بعد وقت بدايتها."

        if errors:
            raise ValidationError(errors)

    @property
    def is_currently_in_maintenance_window(self) -> bool:
        now = timezone.now()

        if self.maintenance_starts_at and self.maintenance_ends_at:
            return self.maintenance_starts_at <= now <= self.maintenance_ends_at

        if self.maintenance_starts_at and not self.maintenance_ends_at:
            return now >= self.maintenance_starts_at

        if not self.maintenance_starts_at and self.maintenance_ends_at:
            return now <= self.maintenance_ends_at

        return False

    @property
    def maintenance_window_label(self) -> str:
        local_start = timezone.localtime(self.maintenance_starts_at) if self.maintenance_starts_at else None
        local_end = timezone.localtime(self.maintenance_ends_at) if self.maintenance_ends_at else None

        if local_start and local_end:
            return (
                f"من {local_start.strftime('%Y-%m-%d %I:%M %p')} "
                f"إلى {local_end.strftime('%Y-%m-%d %I:%M %p')}"
            )

        if local_start:
            return f"تبدأ في {local_start.strftime('%Y-%m-%d %I:%M %p')}"

        if local_end:
            return f"تنتهي في {local_end.strftime('%Y-%m-%d %I:%M %p')}"

        return "غير محدد"

    @classmethod
    def get_solo(cls) -> "SiteSetting":
        obj, _created = cls.objects.get_or_create(pk=1)
        return obj

    def save(self, *args, **kwargs):
        self.pk = 1
        self.full_clean()
        return super().save(*args, **kwargs)


# =========================
# القطاعات
# =========================
class Sector(models.Model):
    class Meta:
        verbose_name = "قطاع"
        verbose_name_plural = "القطاعات"
        ordering = ["name"]
        indexes = [
            models.Index(fields=["name"]),
            models.Index(fields=["is_active"]),
        ]

    name = models.CharField("اسم القطاع", max_length=150, unique=True)
    is_active = models.BooleanField("نشط", default=True)

    def __str__(self) -> str:
        return self.name

    def clean(self):
        super().clean()
        self.name = _clean_text(self.name)

        if not self.name:
            raise ValidationError({"name": "اسم القطاع مطلوب."})

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)


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
            models.Index(fields=["gender"]),
            models.Index(fields=["sector"]),
            models.Index(fields=["is_active"]),
        ]

    GENDER_CHOICES = [("boys", "بنين"), ("girls", "بنات")]

    stat_code = models.CharField("الرقم الإحصائي", max_length=32, unique=True)
    name = models.CharField("اسم المدرسة", max_length=255)
    gender = models.CharField("النوع", max_length=10, choices=GENDER_CHOICES)
    sector = models.ForeignKey(
        Sector,
        on_delete=models.PROTECT,
        related_name="schools",
        verbose_name="القطاع",
        null=True,
        blank=True,
    )
    is_active = models.BooleanField("نشطة", default=True)

    def __str__(self) -> str:
        sector_name = self.sector.name if self.sector_id else "بدون قطاع"
        return f"{self.name} ({self.stat_code}) — {sector_name}"

    def clean(self):
        super().clean()
        self.stat_code = _clean_text(self.stat_code)
        self.name = _clean_text(self.name)

        if self.stat_code:
            digits = _digits(self.stat_code)
            self.stat_code = digits or self.stat_code

        errors = {}

        if not self.stat_code:
            errors["stat_code"] = "الرقم الإحصائي مطلوب."

        if not self.name:
            errors["name"] = "اسم المدرسة مطلوب."

        if self.sector_id and not self.sector.is_active:
            errors["sector"] = "لا يمكن ربط المدرسة بقطاع غير نشط."

        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)


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
    mobile = models.CharField("الجوال", max_length=20, blank=True, null=True)

    def __str__(self) -> str:
        return f"{self.full_name} — {self.school.name}"

    def clean(self):
        super().clean()
        self.full_name = _clean_text(self.full_name)
        self.mobile = _clean_text(self.mobile) or None

        if self.mobile:
            self.mobile = _digits(self.mobile) or self.mobile

        if not self.full_name:
            raise ValidationError({"full_name": "اسم المدير مطلوب."})

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)


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
            models.Index(fields=["gender"]),
            models.Index(fields=["sector"]),
            models.Index(fields=["is_active"]),
            models.Index(fields=["email_notifications_enabled"]),
        ]

    GENDER_CHOICES = [("boys", "بنين"), ("girls", "بنات")]

    national_id = models.CharField("السجل المدني", max_length=20, unique=True)
    full_name = models.CharField("اسم المشرف", max_length=255)
    mobile = models.CharField("جوال المشرف", max_length=20, blank=True, null=True)

    email = models.EmailField("البريد الإلكتروني", blank=True, null=True)
    email_notifications_enabled = models.BooleanField("تفعيل التنبيهات البريدية", default=True)
    email_verified = models.BooleanField("تم التحقق من البريد", default=False)

    gender = models.CharField("النوع", max_length=10, choices=GENDER_CHOICES, null=True, blank=True)
    sector = models.ForeignKey(
        Sector,
        on_delete=models.PROTECT,
        related_name="supervisors",
        verbose_name="القطاع",
        null=True,
        blank=True,
    )
    is_active = models.BooleanField("نشط", default=True)

    def __str__(self) -> str:
        return self.full_name

    @staticmethod
    def _digits(v: str) -> str:
        return "".join(ch for ch in (v or "") if ch.isdigit())

    def mobile_last4(self) -> str | None:
        d = self._digits(self.mobile or "")
        return d[-4:] if len(d) >= 4 else None

    def clean(self):
        super().clean()
        self.national_id = self._digits(self.national_id or "")
        self.full_name = _clean_text(self.full_name)
        self.mobile = _clean_text(self.mobile) or None
        self.email = _clean_text(self.email) or None

        if self.mobile:
            self.mobile = self._digits(self.mobile) or self.mobile

        if self.email:
            self.email = self.email.lower()

        errors = {}
        if not self.national_id:
            errors["national_id"] = "السجل المدني مطلوب."
        elif len(self.national_id) != 10:
            errors["national_id"] = "السجل المدني يجب أن يتكون من 10 أرقام."

        if not self.full_name:
            errors["full_name"] = "اسم المشرف مطلوب."

        if self.sector_id and not self.sector.is_active:
            errors["sector"] = "لا يمكن ربط المشرف بقطاع غير نشط."

        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)


# =========================
# OTP البريد الإلكتروني
# =========================
class EmailOTP(models.Model):
    class Meta:
        verbose_name = "رمز تحقق البريد"
        verbose_name_plural = "رموز تحقق البريد"
        ordering = ["-created_at"]
        indexes = [
            models.Index(fields=["email", "purpose"]),
            models.Index(fields=["is_used"]),
            models.Index(fields=["expires_at"]),
        ]

    PURPOSE_EMAIL_VERIFICATION = "email_verification"
    PURPOSE_LOGIN = "login"
    PURPOSE_PASSWORD_RESET = "password_reset"

    PURPOSE_CHOICES = [
        (PURPOSE_EMAIL_VERIFICATION, "التحقق من البريد"),
        (PURPOSE_LOGIN, "تسجيل الدخول"),
        (PURPOSE_PASSWORD_RESET, "إعادة تعيين كلمة المرور"),
    ]

    supervisor = models.ForeignKey(
        Supervisor,
        on_delete=models.CASCADE,
        related_name="email_otps",
        verbose_name="المشرف",
        null=True,
        blank=True,
    )
    email = models.EmailField("البريد الإلكتروني", db_index=True)
    code = models.CharField("رمز التحقق", max_length=6)
    purpose = models.CharField(
        "الغرض",
        max_length=50,
        choices=PURPOSE_CHOICES,
        default=PURPOSE_EMAIL_VERIFICATION,
    )
    is_used = models.BooleanField("تم الاستخدام", default=False)

    created_at = models.DateTimeField("تاريخ الإنشاء", auto_now_add=True)
    expires_at = models.DateTimeField("ينتهي في")
    verified_at = models.DateTimeField("وقت التحقق", null=True, blank=True)

    def __str__(self) -> str:
        return f"{self.email} - {self.code}"

    @property
    def is_expired(self) -> bool:
        return timezone.now() > self.expires_at

    @classmethod
    def generate_code(cls) -> str:
        return f"{random.randint(0, 999999):06d}"

    @classmethod
    def create_otp(
        cls,
        email: str,
        purpose: str = PURPOSE_EMAIL_VERIFICATION,
        expiry_minutes: int = 10,
        supervisor: Supervisor | None = None,
    ) -> "EmailOTP":
        email = _clean_text(email).lower()

        cls.objects.filter(
            email=email,
            purpose=purpose,
            is_used=False,
            expires_at__gt=timezone.now(),
        ).update(is_used=True)

        code = cls.generate_code()
        return cls.objects.create(
            supervisor=supervisor,
            email=email,
            code=code,
            purpose=purpose,
            expires_at=timezone.now() + timedelta(minutes=expiry_minutes),
        )

    def mark_as_used(self):
        self.is_used = True
        self.verified_at = timezone.now()
        self.save(update_fields=["is_used", "verified_at"])

    def clean(self):
        super().clean()
        self.email = _clean_text(self.email).lower()
        self.code = _digits(self.code)

        errors = {}

        if not self.email:
            errors["email"] = "البريد الإلكتروني مطلوب."

        if not self.code:
            errors["code"] = "رمز التحقق مطلوب."
        elif len(self.code) != 6:
            errors["code"] = "رمز التحقق يجب أن يكون 6 أرقام."

        if self.expires_at and self.verified_at and self.verified_at < self.created_at:
            errors["verified_at"] = "وقت التحقق غير صحيح."

        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)


# =========================
# إسناد مدرسة لمشرف
# =========================
class Assignment(models.Model):
    class Meta:
        verbose_name = "إسناد مدرسة"
        verbose_name_plural = "إسنادات المدارس"
        constraints = [
            models.UniqueConstraint(
                fields=["supervisor", "school"],
                name="uniq_assignment_sup_school",
            ),
        ]
        ordering = ["supervisor__full_name", "school__name"]
        indexes = [
            models.Index(fields=["is_active"]),
            models.Index(fields=["created_at"]),
        ]

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
    created_at = models.DateTimeField("تاريخ الإسناد", auto_now_add=True)
    updated_at = models.DateTimeField("آخر تحديث", auto_now=True)

    def __str__(self) -> str:
        return f"{self.supervisor} ← {self.school}"

    def clean(self):
        super().clean()
        errors = {}

        if self.supervisor_id and not self.supervisor.is_active:
            errors["supervisor"] = "لا يمكن إسناد مدرسة إلى مشرف غير نشط."

        if self.school_id and not self.school.is_active:
            errors["school"] = "لا يمكن إسناد مدرسة غير نشطة."

        if self.supervisor_id and self.school_id:
            if self.supervisor.gender and self.school.gender and self.supervisor.gender != self.school.gender:
                errors["school"] = "لا يمكن إسناد مدرسة بنات إلى مشرف بنين أو العكس."

            if self.supervisor.sector_id and self.school.sector_id:
                if self.supervisor.sector_id != self.school.sector_id:
                    errors["school"] = "لا يمكن إسناد مدرسة من قطاع مختلف لهذا المشرف."

        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)


# =========================
# جدول الأسابيع
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

    def clean(self):
        super().clean()
        self.title = _clean_text(self.title) or None

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)


# =========================
# الخطة الأسبوعية
# =========================
class Plan(models.Model):
    class Meta:
        verbose_name = "خطة أسبوعية"
        verbose_name_plural = "الخطط الأسبوعية"
        constraints = [
            models.UniqueConstraint(
                fields=["supervisor", "week"],
                name="uniq_plan_supervisor_week",
            ),
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
    week = models.ForeignKey(
        PlanWeek,
        on_delete=models.CASCADE,
        related_name="plans",
        verbose_name="الأسبوع",
    )

    status = models.CharField(
        "الحالة",
        max_length=20,
        choices=STATUS_CHOICES,
        default=STATUS_DRAFT,
    )
    saved_at = models.DateTimeField("وقت الحفظ", null=True, blank=True)
    approved_at = models.DateTimeField("وقت الاعتماد", null=True, blank=True)
    admin_note = models.TextField("ملاحظة الإدارة", blank=True, null=True)

    def __str__(self) -> str:
        return f"{self.supervisor} — أسبوع {self.week.week_no}"

    def is_fully_filled(self) -> bool:
        needed = {0, 1, 2, 3, 4}
        filled = {
            d.weekday
            for d in self.days.all()
            if (d.school_id is not None) or (d.visit_type == PlanDay.VISIT_NONE)
        }
        return needed.issubset(filled)

    def clean(self):
        super().clean()
        self.admin_note = _clean_text(self.admin_note) or None

        errors = {}

        if self.week_id and self.week.is_break:
            errors["week"] = "لا يمكن إنشاء خطة على أسبوع مخصص كإجازة أو توقف."

        if self.status == self.STATUS_APPROVED and not self.approved_at:
            self.approved_at = timezone.now()

        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)


# =========================
# تفاصيل أيام الخطة
# =========================
class PlanDay(models.Model):
    class Meta:
        verbose_name = "يوم خطة"
        verbose_name_plural = "أيام الخطة"
        constraints = [
            models.UniqueConstraint(
                fields=["plan", "weekday"],
                name="uniq_planday_plan_weekday",
            ),
        ]
        ordering = ["weekday"]
        indexes = [
            models.Index(fields=["weekday"]),
            models.Index(fields=["visit_type"]),
            models.Index(fields=["visited"]),
        ]

    WEEKDAY_CHOICES = [
        (0, "الأحد"),
        (1, "الإثنين"),
        (2, "الثلاثاء"),
        (3, "الأربعاء"),
        (4, "الخميس"),
    ]

    VISIT_IN = "in"
    VISIT_REMOTE = "remote"
    VISIT_NONE = "none"

    VISIT_CHOICES = [
        (VISIT_IN, "حضوري"),
        (VISIT_REMOTE, "عن بعد"),
        (VISIT_NONE, "بدون زيارة مدرسية"),
    ]

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

    school = models.ForeignKey(
        School,
        on_delete=models.PROTECT,
        null=True,
        blank=True,
        verbose_name="المدرسة",
    )

    visit_type = models.CharField(
        "نوع اليوم",
        max_length=10,
        choices=VISIT_CHOICES,
        default=VISIT_IN,
    )

    no_visit_reason = models.CharField(
        "سبب عدم الزيارة",
        max_length=20,
        choices=NO_VISIT_REASON_CHOICES,
        null=True,
        blank=True,
    )

    note = models.CharField("ملاحظة", max_length=120, null=True, blank=True)

    visited = models.BooleanField("تمت الزيارة", default=False)
    visited_at = models.DateTimeField("وقت تنفيذ الزيارة", null=True, blank=True)
    visit_note = models.CharField("ملاحظة تنفيذ الزيارة", max_length=200, null=True, blank=True)

    def __str__(self) -> str:
        return f"{self.plan} — {self.get_weekday_display()} — {self.school or '—'}"

    def clean(self):
        super().clean()

        self.note = _clean_text(self.note) or None
        self.no_visit_reason = _clean_text(self.no_visit_reason) or None
        self.visit_note = _clean_text(self.visit_note) or None

        errors = {}

        if self.visit_type == self.VISIT_NONE:
            if self.school_id is not None:
                errors["school"] = "لا يمكن اختيار مدرسة عندما يكون اليوم بدون زيارة مدرسية."
            if not self.no_visit_reason:
                errors["no_visit_reason"] = "حدد سبب عدم الزيارة عند اختيار (بدون زيارة مدرسية)."
        else:
            if self.school_id is None:
                errors["school"] = "يجب اختيار مدرسة عندما يكون اليوم حضوريًا أو عن بُعد."
            if self.no_visit_reason:
                errors["no_visit_reason"] = "سبب عدم الزيارة يُستخدم فقط مع خيار (بدون زيارة مدرسية)."

        if self.school_id and self.plan_id and self.plan.supervisor_id:
            is_assigned = Assignment.objects.filter(
                supervisor=self.plan.supervisor,
                school=self.school,
                is_active=True,
            ).exists()
            if not is_assigned:
                errors["school"] = "هذه المدرسة ليست ضمن المدارس المسندة لهذا المشرف."

        if self.school_id and not self.school.is_active:
            errors["school"] = "لا يمكن اختيار مدرسة غير نشطة."

        if self.visit_type == self.VISIT_NONE:
            if self.visited:
                errors["visited"] = "لا يمكن تعليم يوم بدون زيارة مدرسية على أنه تمت زيارته."
            if self.visited_at is not None:
                errors["visited_at"] = "لا يمكن تحديد وقت زيارة ليوم بدون زيارة مدرسية."

        if self.visit_type == self.VISIT_REMOTE:
            if self.visited:
                errors["visited"] = "حالة (تمت الزيارة) مخصصة للزيارات الحضورية المدرسية فقط."
            if self.visited_at is not None:
                errors["visited_at"] = "وقت الزيارة يستخدم للزيارة الحضورية فقط."

        if self.visit_type == self.VISIT_IN and self.school_id is not None:
            if self.visited and self.visited_at is None:
                self.visited_at = timezone.now()
            if not self.visited and self.visited_at is not None:
                self.visited_at = None
        else:
            if self.visit_note:
                errors["visit_note"] = "ملاحظة تنفيذ الزيارة تُستخدم فقط مع الزيارة الحضورية المدرسية."

        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)


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
    status = models.CharField(
        "الحالة",
        max_length=20,
        choices=STATUS_CHOICES,
        default=STATUS_PENDING,
    )
    created_at = models.DateTimeField("تاريخ الطلب", auto_now_add=True)
    resolved_at = models.DateTimeField("تاريخ المعالجة", null=True, blank=True)

    def __str__(self) -> str:
        return f"فك اعتماد — {self.plan} — {self.get_status_display()}"

    def clean(self):
        super().clean()
        errors = {}

        if self.plan_id:
            if self.plan.status not in (Plan.STATUS_APPROVED, Plan.STATUS_UNLOCK_REQUESTED):
                errors["plan"] = "لا يمكن إنشاء طلب فك اعتماد إلا لخطة معتمدة أو عليها طلب فك اعتماد."

        if self.status == self.STATUS_PENDING and self.resolved_at is not None:
            errors["resolved_at"] = "لا ينبغي تحديد تاريخ المعالجة لطلب ما زال معلقًا."

        if self.status in (self.STATUS_APPROVED, self.STATUS_REJECTED) and self.resolved_at is None:
            self.resolved_at = timezone.now()

        if errors:
            raise ValidationError(errors)

    def mark_resolved(self):
        if not self.resolved_at:
            self.resolved_at = timezone.now()

    def approve(self):
        if self.status != self.STATUS_PENDING:
            return
        self.status = self.STATUS_APPROVED
        self.mark_resolved()
        self.save(update_fields=["status", "resolved_at"])

    def reject(self):
        if self.status != self.STATUS_PENDING:
            return
        self.status = self.STATUS_REJECTED
        self.mark_resolved()
        self.save(update_fields=["status", "resolved_at"])

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)


# =========================
# إشعارات المشرف
# =========================
class SupervisorNotification(models.Model):
    class Meta:
        verbose_name = "إشعار مشرف"
        verbose_name_plural = "إشعارات المشرفين"
        ordering = ["-created_at"]
        indexes = [
            models.Index(fields=["supervisor", "is_read"]),
            models.Index(fields=["created_at"]),
            models.Index(fields=["notif_type"]),
        ]

    TYPE_APPROVED = "approved"
    TYPE_RETURNED = "returned"
    TYPE_UNLOCK_APPROVED = "unlock_approved"
    TYPE_UNLOCK_REJECTED = "unlock_rejected"
    TYPE_ADMIN_ALERT = "admin_alert"

    TYPE_CHOICES = [
        (TYPE_APPROVED, "تم اعتماد الخطة"),
        (TYPE_RETURNED, "تمت إعادة الخطة للمراجعة"),
        (TYPE_UNLOCK_APPROVED, "تم قبول طلب فك الاعتماد"),
        (TYPE_UNLOCK_REJECTED, "تم رفض طلب فك الاعتماد"),
        (TYPE_ADMIN_ALERT, "تنبيه إداري"),
    ]

    supervisor = models.ForeignKey(
        Supervisor,
        on_delete=models.CASCADE,
        related_name="notifications",
        verbose_name="المشرف",
    )
    plan = models.ForeignKey(
        Plan,
        on_delete=models.CASCADE,
        related_name="notifications",
        verbose_name="الخطة",
        null=True,
        blank=True,
    )
    notif_type = models.CharField(
        "نوع الإشعار",
        max_length=30,
        choices=TYPE_CHOICES,
    )
    title = models.CharField("عنوان الإشعار", max_length=200)
    message = models.TextField("نص الإشعار", blank=True, null=True)
    is_read = models.BooleanField("مقروء", default=False)
    created_at = models.DateTimeField("تاريخ الإنشاء", auto_now_add=True)

    def __str__(self) -> str:
        return f"{self.supervisor} — {self.title}"

    def clean(self):
        super().clean()
        self.title = _clean_text(self.title)
        self.message = _clean_text(self.message) or None

        errors = {}

        if not self.title:
            errors["title"] = "عنوان الإشعار مطلوب."

        if self.plan_id and self.supervisor_id and self.plan.supervisor_id != self.supervisor_id:
            errors["plan"] = "الخطة المحددة لا تتبع هذا المشرف."

        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)


# =========================
# روابط الخطابات الأسبوعية
# =========================
class WeeklyLetterLink(models.Model):
    class Meta:
        verbose_name = "رابط خطاب أسبوعي"
        verbose_name_plural = "روابط الخطابات الأسبوعية"
        ordering = ["week__week_no"]
        indexes = [
            models.Index(fields=["is_active"]),
        ]

    week = models.OneToOneField(
        PlanWeek,
        on_delete=models.CASCADE,
        related_name="letter_link",
        verbose_name="الأسبوع",
    )
    title = models.CharField("عنوان الخطاب", max_length=200, blank=True, null=True)
    drive_url = models.URLField("رابط الخطاب على Google Drive", max_length=1000)
    note = models.CharField("ملاحظة", max_length=255, blank=True, null=True)
    is_active = models.BooleanField("نشط", default=True)
    updated_at = models.DateTimeField("آخر تحديث", auto_now=True)
    created_at = models.DateTimeField("تاريخ الإنشاء", auto_now_add=True)

    def __str__(self) -> str:
        return self.title or f"خطاب الأسبوع {self.week.week_no}"

    def clean(self):
        super().clean()
        self.title = _clean_text(self.title) or None
        self.note = _clean_text(self.note) or None
        self.drive_url = _clean_text(self.drive_url)

        errors = {}

        if not self.drive_url:
            errors["drive_url"] = "رابط الخطاب مطلوب."

        if self.week_id and self.week.is_break:
            errors["week"] = "لا يمكن ربط خطاب بأسبوع إجازة أو توقف."

        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)