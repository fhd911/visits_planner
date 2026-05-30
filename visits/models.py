from __future__ import annotations

import random
from datetime import timedelta

from django.conf import settings

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
# قائد/قائدة المدرسة
# =========================
class Principal(models.Model):
    class Meta:
        verbose_name = "قائد مدرسة"
        verbose_name_plural = "قادة المدارس"
        ordering = ["school__name", "full_name"]
        indexes = [
            models.Index(fields=["national_id"]),
            models.Index(fields=["gender"]),
            models.Index(fields=["stage"]),
            models.Index(fields=["faris_status"]),
            models.Index(fields=["current_work"]),
            models.Index(fields=["is_active"]),
            models.Index(fields=["last_imported_at"]),
        ]

    GENDER_BOYS = "boys"
    GENDER_GIRLS = "girls"

    GENDER_CHOICES = [
        (GENDER_BOYS, "بنين"),
        (GENDER_GIRLS, "بنات"),
    ]

    school = models.OneToOneField(
        School,
        on_delete=models.CASCADE,
        related_name="principal",
        verbose_name="المدرسة",
    )

    full_name = models.CharField("اسم القائد/ة", max_length=255)
    national_id = models.CharField(
        "السجل المدني",
        max_length=20,
        blank=True,
        null=True,
        db_index=True,
        help_text="يحفظ داخليًا للمطابقة والتحقق، ولا يُعرض في بوابة الاطلاع.",
    )

    gender = models.CharField(
        "الجنس",
        max_length=10,
        choices=GENDER_CHOICES,
        blank=True,
        null=True,
    )

    stage = models.CharField("المرحلة", max_length=120, blank=True, null=True)
    faris_status = models.CharField("حالة فارس", max_length=120, blank=True, null=True)
    current_work = models.CharField("العمل الحالي", max_length=160, blank=True, null=True)

    specialization = models.CharField("التخصص", max_length=160, blank=True, null=True)
    qualification = models.CharField("المؤهل", max_length=160, blank=True, null=True)
    rank = models.CharField("الرتبة", max_length=160, blank=True, null=True)

    mobile = models.CharField("الجوال", max_length=20, blank=True, null=True)

    source_filename = models.CharField(
        "اسم ملف الاستيراد",
        max_length=255,
        blank=True,
        null=True,
    )
    last_imported_at = models.DateTimeField("آخر استيراد", null=True, blank=True)

    is_active = models.BooleanField("نشط", default=True)

    note = models.CharField("ملاحظة", max_length=255, blank=True, null=True)

    created_at = models.DateTimeField("تاريخ الإنشاء", auto_now_add=True)
    updated_at = models.DateTimeField("آخر تحديث", auto_now=True)

    def __str__(self) -> str:
        school_name = self.school.name if self.school_id else "بدون مدرسة"
        return f"{self.full_name} — {school_name}"

    @staticmethod
    def normalize_gender(value: str | None) -> str | None:
        value = _clean_text(value)
        if not value:
            return None

        if value in {"بنين", "boys", "boy", "male", "ذكر"}:
            return Principal.GENDER_BOYS

        if value in {"بنات", "girls", "girl", "female", "أنثى"}:
            return Principal.GENDER_GIRLS

        return None

    def clean(self):
        super().clean()

        self.full_name = _clean_text(self.full_name)
        self.national_id = _digits(self.national_id or "") or None
        self.gender = self.normalize_gender(self.gender) or self.gender or None

        self.stage = _clean_text(self.stage) or None
        self.faris_status = _clean_text(self.faris_status) or None
        self.current_work = _clean_text(self.current_work) or None
        self.specialization = _clean_text(self.specialization) or None
        self.qualification = _clean_text(self.qualification) or None
        self.rank = _clean_text(self.rank) or None
        self.mobile = _clean_text(self.mobile) or None
        self.source_filename = _clean_text(self.source_filename) or None
        self.note = _clean_text(self.note) or None

        if self.mobile:
            self.mobile = _digits(self.mobile) or self.mobile

        errors = {}

        if not self.full_name:
            errors["full_name"] = "اسم القائد/ة مطلوب."

        if self.national_id and len(self.national_id) != 10:
            errors["national_id"] = "السجل المدني يجب أن يتكون من 10 أرقام."

        if self.gender and self.gender not in dict(self.GENDER_CHOICES):
            errors["gender"] = "الجنس غير صحيح."

        if self.school_id and self.gender and self.school.gender and self.school.gender != self.gender:
            errors["gender"] = "جنس القائد/ة لا يطابق نوع المدرسة."

        if self.school_id and not self.school.is_active:
            errors["school"] = "لا يمكن ربط قائد/ة بمدرسة غير نشطة."

        if errors:
            raise ValidationError(errors)

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
# العام الدراسي والفصول
# =========================
class AcademicYear(models.Model):
    class Meta:
        verbose_name = "عام دراسي"
        verbose_name_plural = "الأعوام الدراسية"
        ordering = ["-starts_at", "-id"]
        indexes = [
            models.Index(fields=["is_current"]),
            models.Index(fields=["is_active"]),
        ]

    name = models.CharField("العام الدراسي", max_length=30, unique=True)
    starts_at = models.DateField("تاريخ بداية العام", null=True, blank=True)
    ends_at = models.DateField("تاريخ نهاية العام", null=True, blank=True)
    is_current = models.BooleanField("العام الحالي", default=False)
    is_active = models.BooleanField("نشط", default=True)
    created_at = models.DateTimeField("تاريخ الإنشاء", auto_now_add=True)
    updated_at = models.DateTimeField("آخر تحديث", auto_now=True)

    def __str__(self) -> str:
        return self.name

    def clean(self):
        super().clean()
        self.name = _clean_text(self.name)
        errors = {}

        if not self.name:
            errors["name"] = "اسم العام الدراسي مطلوب."

        if self.starts_at and self.ends_at and self.ends_at <= self.starts_at:
            errors["ends_at"] = "يجب أن يكون تاريخ نهاية العام بعد تاريخ بدايته."

        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        self.full_clean()
        result = super().save(*args, **kwargs)
        if self.is_current:
            AcademicYear.objects.exclude(pk=self.pk).update(is_current=False)
        return result


class Semester(models.Model):
    class Meta:
        verbose_name = "فصل دراسي"
        verbose_name_plural = "الفصول الدراسية"
        constraints = [
            models.UniqueConstraint(
                fields=["academic_year", "number"],
                name="uniq_semester_academic_year_number",
            ),
        ]
        ordering = ["academic_year__starts_at", "number"]
        indexes = [
            models.Index(fields=["number"]),
            models.Index(fields=["is_current"]),
            models.Index(fields=["is_open"]),
        ]

    FIRST = 1
    SECOND = 2

    SEMESTER_CHOICES = [
        (FIRST, "الفصل الدراسي الأول"),
        (SECOND, "الفصل الدراسي الثاني"),
    ]

    academic_year = models.ForeignKey(
        AcademicYear,
        on_delete=models.CASCADE,
        related_name="semesters",
        verbose_name="العام الدراسي",
    )
    number = models.PositiveSmallIntegerField("الفصل الدراسي", choices=SEMESTER_CHOICES)
    title = models.CharField("اسم الفصل", max_length=120, blank=True, null=True)
    starts_at = models.DateField("تاريخ بداية الفصل")
    ends_at = models.DateField("تاريخ نهاية الفصل", null=True, blank=True)
    weeks_count = models.PositiveSmallIntegerField(
        "عدد الأسابيع",
        default=19,
        validators=[MinValueValidator(1), MaxValueValidator(25)],
    )
    is_current = models.BooleanField("الفصل الحالي", default=False)
    is_open = models.BooleanField("مفتوح", default=False)
    created_at = models.DateTimeField("تاريخ الإنشاء", auto_now_add=True)
    updated_at = models.DateTimeField("آخر تحديث", auto_now=True)

    def __str__(self) -> str:
        year = self.academic_year.name if self.academic_year_id else "—"
        return f"{self.title or self.get_number_display()} — {year}"

    def clean(self):
        super().clean()
        self.title = _clean_text(self.title) or None
        errors = {}

        if self.ends_at and self.starts_at and self.ends_at <= self.starts_at:
            errors["ends_at"] = "يجب أن يكون تاريخ نهاية الفصل بعد تاريخ بدايته."

        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        if not _clean_text(self.title):
            self.title = self.get_number_display()
        self.full_clean()
        result = super().save(*args, **kwargs)
        if self.is_current:
            Semester.objects.exclude(pk=self.pk).update(is_current=False)
        return result


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
            models.Index(fields=["is_current"]),
            models.Index(fields=["is_open_for_supervisors"]),
            models.Index(fields=["semester", "semester_week_no"]),
        ]

    week_no = models.PositiveSmallIntegerField(
        "رقم الأسبوع",
        unique=True,
        validators=[MinValueValidator(1), MaxValueValidator(60)],
    )
    academic_year = models.ForeignKey(
        AcademicYear,
        on_delete=models.SET_NULL,
        related_name="weeks",
        verbose_name="العام الدراسي",
        null=True,
        blank=True,
    )
    semester = models.ForeignKey(
        Semester,
        on_delete=models.SET_NULL,
        related_name="weeks",
        verbose_name="الفصل الدراسي",
        null=True,
        blank=True,
    )
    semester_week_no = models.PositiveSmallIntegerField(
        "رقم الأسبوع داخل الفصل",
        null=True,
        blank=True,
        validators=[MinValueValidator(1), MaxValueValidator(25)],
    )
    start_sunday = models.DateField("بداية الأسبوع (الأحد)")
    title = models.CharField("ملاحظة/اسم", max_length=120, blank=True, null=True)
    is_break = models.BooleanField("إجازة/توقف؟", default=False)
    is_current = models.BooleanField("الأسبوع الحالي للمشرفين", default=False)
    is_open_for_supervisors = models.BooleanField("مفتوح للمشرفين", default=False)

    def __str__(self) -> str:
        extra = " (إجازة)" if self.is_break else ""
        if self.semester_id and self.semester_week_no:
            label = f"{self.semester} — الأسبوع {self.semester_week_no}"
        else:
            label = f"الأسبوع {self.week_no}"
        t = f" — {self.title}" if self.title else ""
        return f"{label} يبدأ {self.start_sunday}{t}{extra}"

    @property
    def end_thursday(self):
        return self.start_sunday + timedelta(days=4)

    @property
    def display_label(self) -> str:
        if self.semester_id and self.semester_week_no:
            return f"{self.semester} — الأسبوع {self.semester_week_no}"
        if self.title:
            return f"الأسبوع {self.week_no} — {self.title}"
        return f"الأسبوع {self.week_no}"

    def clean(self):
        super().clean()
        self.title = _clean_text(self.title) or None
        errors = {}

        if self.semester_id and not self.academic_year_id:
            self.academic_year = self.semester.academic_year

        if self.semester_id and self.academic_year_id:
            if self.semester.academic_year_id != self.academic_year_id:
                errors["semester"] = "الفصل الدراسي لا يتبع العام الدراسي المحدد."

        if self.semester_week_no and self.semester_id:
            if self.semester_week_no > self.semester.weeks_count:
                errors["semester_week_no"] = "رقم الأسبوع داخل الفصل يتجاوز عدد أسابيع الفصل."

        if self.is_break:
            self.is_current = False
            self.is_open_for_supervisors = False

        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        self.full_clean()
        result = super().save(*args, **kwargs)
        if self.is_current:
            PlanWeek.objects.exclude(pk=self.pk).update(is_current=False)
        if self.is_open_for_supervisors:
            PlanWeek.objects.exclude(pk=self.pk).update(is_open_for_supervisors=False)
        return result



# =========================
# الأيام المغلقة داخل الأسبوع
# =========================
class PlanClosedDay(models.Model):
    class Meta:
        verbose_name = "يوم مغلق في الخطة"
        verbose_name_plural = "الأيام المغلقة في الخطط"
        constraints = [
            models.UniqueConstraint(
                fields=["week", "weekday"],
                name="uniq_plan_closed_day_week_weekday",
            ),
        ]
        ordering = ["week__week_no", "weekday"]
        indexes = [
            models.Index(fields=["weekday"]),
            models.Index(fields=["reason_type"]),
            models.Index(fields=["is_active"]),
        ]

    NATIONAL_DAY = "national_day"
    FOUNDATION_DAY = "foundation_day"
    OFFICIAL_HOLIDAY = "official_holiday"
    STUDY_SUSPENDED = "study_suspended"
    OTHER = "other"

    REASON_CHOICES = [
        (NATIONAL_DAY, "اليوم الوطني"),
        (FOUNDATION_DAY, "يوم التأسيس"),
        (OFFICIAL_HOLIDAY, "إجازة رسمية"),
        (STUDY_SUSPENDED, "تعليق دراسة"),
        (OTHER, "أخرى"),
    ]

    WEEKDAY_CHOICES = [
        (0, "الأحد"),
        (1, "الإثنين"),
        (2, "الثلاثاء"),
        (3, "الأربعاء"),
        (4, "الخميس"),
    ]

    week = models.ForeignKey(
        PlanWeek,
        on_delete=models.CASCADE,
        related_name="closed_days",
        verbose_name="الأسبوع",
    )
    weekday = models.PositiveSmallIntegerField("اليوم", choices=WEEKDAY_CHOICES)
    date = models.DateField("التاريخ", null=True, blank=True)
    reason_type = models.CharField(
        "نوع الإغلاق",
        max_length=30,
        choices=REASON_CHOICES,
        default=OFFICIAL_HOLIDAY,
    )
    reason_title = models.CharField("سبب الإغلاق", max_length=150, default="إجازة رسمية")
    count_as_completed = models.BooleanField("يُحسب مكتملًا", default=True)
    is_active = models.BooleanField("نشط", default=True)
    created_at = models.DateTimeField("تاريخ الإنشاء", auto_now_add=True)
    updated_at = models.DateTimeField("آخر تحديث", auto_now=True)

    def __str__(self) -> str:
        return f"{self.week.display_label} — {self.get_weekday_display()} — {self.reason_title}"

    def clean(self):
        super().clean()
        self.reason_title = _clean_text(self.reason_title) or "إجازة رسمية"

        if self.week_id and self.date is None:
            self.date = self.week.start_sunday + timedelta(days=int(self.weekday or 0))

        if self.week_id and self.date:
            expected = self.week.start_sunday + timedelta(days=int(self.weekday or 0))
            if self.date != expected:
                raise ValidationError({"date": "تاريخ اليوم المغلق لا يوافق اليوم المحدد داخل الأسبوع."})

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
    STATUS_UNLOCKED = "unlocked"

    STATUS_CHOICES = [
        (STATUS_DRAFT, "مسودة"),
        (STATUS_APPROVED, "معتمدة"),
        (STATUS_UNLOCK_REQUESTED, "طلب فك اعتماد"),
        (STATUS_UNLOCKED, "مفكوكة للتعديل"),
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

    unlocked_at = models.DateTimeField("وقت فك الاعتماد", null=True, blank=True)
    unlocked_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="unlocked_visit_plans",
        verbose_name="فُك الاعتماد بواسطة",
    )
    unlocked_reason = models.TextField(
        "سبب فك الاعتماد",
        blank=True,
        null=True,
        help_text="يحفظ سبب المشرف وملاحظة الإدارة عند قبول فك الاعتماد.",
    )

    def __str__(self) -> str:
        return f"{self.supervisor} — أسبوع {self.week.week_no}"

    def is_fully_filled(self) -> bool:
        needed = {0, 1, 2, 3, 4}
        filled = {
            d.weekday
            for d in self.days.all()
            if (d.school_id is not None) or (d.visit_type == PlanDay.VISIT_NONE)
        }
        if self.week_id:
            filled.update(
                PlanClosedDay.objects.filter(
                    week=self.week,
                    is_active=True,
                    count_as_completed=True,
                ).values_list("weekday", flat=True)
            )
        return needed.issubset(filled)

    def clean(self):
        super().clean()
        self.admin_note = _clean_text(self.admin_note) or None
        self.unlocked_reason = _clean_text(self.unlocked_reason) or None

        errors = {}

        if self.week_id and self.week.is_break:
            errors["week"] = "لا يمكن إنشاء خطة على أسبوع مخصص كإجازة أو توقف."

        if self.status == self.STATUS_APPROVED and not self.approved_at:
            self.approved_at = timezone.now()

        if self.status == self.STATUS_UNLOCKED and not self.unlocked_at:
            self.unlocked_at = timezone.now()

        if self.status != self.STATUS_UNLOCKED and self.status != self.STATUS_UNLOCK_REQUESTED:
            # لا نمسح unlocked_reason/unlocked_at حتى يبقى الأثر التاريخي محفوظًا.
            pass

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
    REASON_OFFICIAL_HOLIDAY = "official_holiday"
    REASON_OTHER = "other"

    NO_VISIT_REASON_CHOICES = [
        (REASON_MEETING, "اجتماع"),
        (REASON_TRAINING, "تدريب"),
        (REASON_VISIT, "لقاء/فعالية"),
        (REASON_OFFICE, "عمل مكتبي"),
        (REASON_OFFICIAL_HOLIDAY, "إجازة رسمية / يوم مغلق"),
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
    reason = models.TextField("سبب طلب فك الاعتماد", blank=True, null=True)
    admin_note = models.TextField("ملاحظة الإدارة على طلب الفك", blank=True, null=True)
    previous_status = models.CharField("حالة الخطة قبل الطلب", max_length=30, blank=True, null=True)
    resolved_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="resolved_unlock_requests",
        verbose_name="تمت المعالجة بواسطة",
    )
    created_at = models.DateTimeField("تاريخ الطلب", auto_now_add=True)
    resolved_at = models.DateTimeField("تاريخ المعالجة", null=True, blank=True)

    def __str__(self) -> str:
        return f"فك اعتماد — {self.plan} — {self.get_status_display()}"

    def clean(self):
        super().clean()
        self.reason = _clean_text(self.reason) or None
        self.admin_note = _clean_text(self.admin_note) or None
        self.previous_status = _clean_text(self.previous_status) or None

        errors = {}

        if self.plan_id and self.status == self.STATUS_PENDING:
            allowed_statuses = (
                Plan.STATUS_APPROVED,
                Plan.STATUS_UNLOCK_REQUESTED,
            )
            if self.plan.status not in allowed_statuses:
                errors["plan"] = "لا يمكن إنشاء طلب فك اعتماد جديد إلا لخطة معتمدة أو عليها طلب فك اعتماد."

        if self.status == self.STATUS_PENDING and not self.reason:
            errors["reason"] = "سبب طلب فك الاعتماد مطلوب."

        if self.status == self.STATUS_PENDING and self.resolved_at is not None:
            errors["resolved_at"] = "لا ينبغي تحديد تاريخ المعالجة لطلب ما زال معلقًا."

        if self.status in (self.STATUS_APPROVED, self.STATUS_REJECTED) and self.resolved_at is None:
            self.resolved_at = timezone.now()

        if errors:
            raise ValidationError(errors)

    def mark_resolved(self):
        if not self.resolved_at:
            self.resolved_at = timezone.now()

    def approve(self, *, user=None, admin_note: str = ""):
        if self.status != self.STATUS_PENDING:
            return
        self.status = self.STATUS_APPROVED
        self.admin_note = _clean_text(admin_note) or self.admin_note
        if user is not None and getattr(user, "is_authenticated", False):
            self.resolved_by = user
        self.mark_resolved()
        self.save(update_fields=["status", "admin_note", "resolved_by", "resolved_at"])

    def reject(self, *, user=None, admin_note: str = ""):
        if self.status != self.STATUS_PENDING:
            return
        self.status = self.STATUS_REJECTED
        self.admin_note = _clean_text(admin_note) or self.admin_note
        if user is not None and getattr(user, "is_authenticated", False):
            self.resolved_by = user
        self.mark_resolved()
        self.save(update_fields=["status", "admin_note", "resolved_by", "resolved_at"])

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

# =========================
# سجل المتابعة الرقابية
# =========================
class ControlFollowUp(models.Model):
    class Meta:
        verbose_name = "متابعة رقابية"
        verbose_name_plural = "سجل المتابعة الرقابية"
        ordering = ["-updated_at", "-created_at", "-id"]
        indexes = [
            models.Index(fields=["unique_key"]),
            models.Index(fields=["issue_type"]),
            models.Index(fields=["status"]),
            models.Index(fields=["status", "issue_type"]),
            models.Index(fields=["supervisor", "status"]),
            models.Index(fields=["week", "issue_type"]),
            models.Index(fields=["last_notification_at"]),
            models.Index(fields=["created_at"]),
        ]

    ISSUE_INCOMPLETE_PLAN = "incomplete_plan"
    ISSUE_NOT_SAVED_PLAN = "not_saved_plan"
    ISSUE_UNLOCK_REQUEST = "unlock_request"
    ISSUE_UNCOVERED_SCHOOLS = "uncovered_schools"

    ISSUE_CHOICES = [
        (ISSUE_INCOMPLETE_PLAN, "خطة غير مكتملة"),
        (ISSUE_NOT_SAVED_PLAN, "لم يحفظ خطة الأسبوع"),
        (ISSUE_UNLOCK_REQUEST, "طلب فك اعتماد"),
        (ISSUE_UNCOVERED_SCHOOLS, "مدارس غير مغطاة"),
    ]

    STATUS_OPEN = "open"
    STATUS_NOTIFIED = "notified"
    STATUS_PENDING_ADMIN = "pending_admin"
    STATUS_PROCESSED = "processed"
    STATUS_CLOSED = "closed"

    STATUS_CHOICES = [
        (STATUS_OPEN, "مفتوحة"),
        (STATUS_NOTIFIED, "تم التنبيه"),
        (STATUS_PENDING_ADMIN, "بانتظار مراجعة الإدارة"),
        (STATUS_PROCESSED, "تمت المعالجة"),
        (STATUS_CLOSED, "مغلقة إداريًا"),
    ]

    unique_key = models.CharField(
        "المفتاح الفريد للحالة",
        max_length=220,
        unique=True,
        db_index=True,
        help_text="يمنع تكرار نفس الحالة الرقابية للمشرف والأسبوع والخطة.",
    )
    issue_type = models.CharField(
        "نوع الحالة",
        max_length=40,
        choices=ISSUE_CHOICES,
        db_index=True,
    )
    status = models.CharField(
        "حالة المتابعة",
        max_length=30,
        choices=STATUS_CHOICES,
        default=STATUS_OPEN,
        db_index=True,
    )

    supervisor = models.ForeignKey(
        Supervisor,
        on_delete=models.CASCADE,
        related_name="control_followups",
        verbose_name="المشرف",
    )
    plan = models.ForeignKey(
        Plan,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="control_followups",
        verbose_name="الخطة المرتبطة",
    )
    week = models.ForeignKey(
        PlanWeek,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="control_followups",
        verbose_name="الأسبوع",
    )

    title = models.CharField("عنوان الحالة", max_length=220)
    description = models.TextField("وصف الحالة", blank=True, null=True)
    admin_note = models.TextField("ملاحظة/تنبيه الإدارة", blank=True, null=True)

    supervisor_response = models.TextField("إفادة المشرف", blank=True, null=True)
    supervisor_response_at = models.DateTimeField("تاريخ إفادة المشرف", null=True, blank=True)

    admin_review_note = models.TextField("ملاحظة مراجعة الإدارة", blank=True, null=True)
    notification_count = models.PositiveIntegerField("عدد التنبيهات", default=0)
    last_notification_at = models.DateTimeField("آخر تنبيه", null=True, blank=True)

    resolved_at = models.DateTimeField("تاريخ المعالجة", null=True, blank=True)
    closed_at = models.DateTimeField("تاريخ الإغلاق", null=True, blank=True)

    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="created_control_followups",
        verbose_name="أنشئت بواسطة",
    )

    created_at = models.DateTimeField("تاريخ الرصد", auto_now_add=True)
    updated_at = models.DateTimeField("آخر تحديث", auto_now=True)

    def __str__(self) -> str:
        return f"{self.get_issue_type_display()} — {self.supervisor}"

    @property
    def is_active(self) -> bool:
        return self.status in {
            self.STATUS_OPEN,
            self.STATUS_NOTIFIED,
            self.STATUS_PENDING_ADMIN,
        }

    def clean(self):
        super().clean()
        self.unique_key = _clean_text(self.unique_key)
        self.title = _clean_text(self.title)
        self.description = _clean_text(self.description) or None
        self.admin_note = _clean_text(self.admin_note) or None
        self.supervisor_response = _clean_text(self.supervisor_response) or None
        self.admin_review_note = _clean_text(self.admin_review_note) or None

        errors = {}

        if not self.unique_key:
            errors["unique_key"] = "المفتاح الفريد للحالة مطلوب."

        if not self.title:
            errors["title"] = "عنوان الحالة مطلوب."

        if self.plan_id and self.supervisor_id and self.plan.supervisor_id != self.supervisor_id:
            errors["plan"] = "الخطة المحددة لا تتبع هذا المشرف."

        if self.plan_id and self.week_id and self.plan.week_id != self.week_id:
            errors["week"] = "الأسبوع المحدد لا يطابق أسبوع الخطة."

        if self.status == self.STATUS_PENDING_ADMIN and not self.supervisor_response:
            errors["supervisor_response"] = "لا يمكن جعل الحالة بانتظار مراجعة الإدارة دون إفادة المشرف."

        if self.status == self.STATUS_PROCESSED and not self.resolved_at:
            self.resolved_at = timezone.now()

        if self.status == self.STATUS_CLOSED and not self.closed_at:
            self.closed_at = timezone.now()

        if self.status not in (self.STATUS_PROCESSED, self.STATUS_CLOSED):
            self.resolved_at = None
            if self.status != self.STATUS_CLOSED:
                self.closed_at = None

        if errors:
            raise ValidationError(errors)

    def mark_notified(self, note: str = ""):
        self.status = self.STATUS_NOTIFIED
        self.notification_count = (self.notification_count or 0) + 1
        self.last_notification_at = timezone.now()
        if note:
            self.admin_note = _clean_text(note) or None
        self.save(
            update_fields=[
                "status",
                "notification_count",
                "last_notification_at",
                "admin_note",
                "updated_at",
            ]
        )

    def submit_supervisor_response(self, response: str):
        self.supervisor_response = _clean_text(response)
        self.supervisor_response_at = timezone.now()
        self.status = self.STATUS_PENDING_ADMIN
        self.save(
            update_fields=[
                "supervisor_response",
                "supervisor_response_at",
                "status",
                "updated_at",
            ]
        )

    def accept_processing(self, note: str = ""):
        self.status = self.STATUS_PROCESSED
        self.resolved_at = timezone.now()
        if note:
            self.admin_review_note = _clean_text(note) or None
        self.save(update_fields=["status", "resolved_at", "admin_review_note", "updated_at"])

    def return_to_supervisor(self, note: str = ""):
        self.status = self.STATUS_NOTIFIED
        if note:
            self.admin_review_note = _clean_text(note) or None
        self.notification_count = (self.notification_count or 0) + 1
        self.last_notification_at = timezone.now()
        self.save(
            update_fields=[
                "status",
                "admin_review_note",
                "notification_count",
                "last_notification_at",
                "updated_at",
            ]
        )

    def close_administratively(self, note: str = ""):
        self.status = self.STATUS_CLOSED
        self.closed_at = timezone.now()
        if note:
            self.admin_review_note = _clean_text(note) or None
        self.save(update_fields=["status", "closed_at", "admin_review_note", "updated_at"])

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)



# =========================
# إجراءات المتابعة الرقابية
# =========================
class ControlFollowUpAction(models.Model):
    class Meta:
        verbose_name = "إجراء متابعة رقابية"
        verbose_name_plural = "إجراءات المتابعة الرقابية"
        ordering = ["-created_at", "-id"]
        indexes = [
            models.Index(fields=["action_type"]),
            models.Index(fields=["created_at"]),
        ]

    ACTION_DETECTED = "detected"
    ACTION_UPDATED_DETECTION = "updated_detection"
    ACTION_NOTIFIED = "notified"
    ACTION_SUPERVISOR_RESPONSE = "supervisor_response"
    ACTION_ACCEPTED = "accepted"
    ACTION_RETURNED = "returned"
    ACTION_CLOSED = "closed"
    ACTION_UNLOCK_REQUESTED = "unlock_requested"
    ACTION_UNLOCK_APPROVED = "unlock_approved"
    ACTION_UNLOCK_REJECTED = "unlock_rejected"
    ACTION_PLAN_REAPPROVED = "plan_reapproved"

    ACTION_CHOICES = [
        (ACTION_DETECTED, "رصد الحالة"),
        (ACTION_UPDATED_DETECTION, "تحديث الرصد"),
        (ACTION_NOTIFIED, "إرسال تنبيه"),
        (ACTION_SUPERVISOR_RESPONSE, "إفادة المشرف"),
        (ACTION_ACCEPTED, "قبول المعالجة"),
        (ACTION_RETURNED, "إعادة للمشرف"),
        (ACTION_CLOSED, "إغلاق إداري"),
        (ACTION_UNLOCK_REQUESTED, "طلب فك اعتماد"),
        (ACTION_UNLOCK_APPROVED, "قبول فك الاعتماد"),
        (ACTION_UNLOCK_REJECTED, "رفض فك الاعتماد"),
        (ACTION_PLAN_REAPPROVED, "إعادة اعتماد بعد الفك"),
    ]

    followup = models.ForeignKey(
        ControlFollowUp,
        on_delete=models.CASCADE,
        related_name="actions",
        verbose_name="حالة المتابعة",
    )
    action_type = models.CharField("نوع الإجراء", max_length=40, choices=ACTION_CHOICES)
    from_status = models.CharField("الحالة السابقة", max_length=30, blank=True, null=True)
    to_status = models.CharField("الحالة الجديدة", max_length=30, blank=True, null=True)
    actor_user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="control_followup_actions",
        verbose_name="المستخدم المنفذ",
    )
    actor_supervisor = models.ForeignKey(
        Supervisor,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="control_followup_actions",
        verbose_name="المشرف المنفذ",
    )
    note = models.TextField("ملاحظة الإجراء", blank=True, null=True)
    metadata = models.JSONField("بيانات إضافية", default=dict, blank=True)
    created_at = models.DateTimeField("تاريخ الإجراء", auto_now_add=True)

    def __str__(self) -> str:
        return f"{self.get_action_type_display()} — {self.followup}"

    def clean(self):
        super().clean()
        self.note = _clean_text(self.note) or None
        self.from_status = _clean_text(self.from_status) or None
        self.to_status = _clean_text(self.to_status) or None
        if self.metadata is None:
            self.metadata = {}

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)



# =========================
# تفضيلات وسجل التنبيهات البريدية
# =========================
class EmailNotificationPreference(models.Model):
    class Meta:
        verbose_name = "تفضيلات التنبيهات البريدية"
        verbose_name_plural = "تفضيلات التنبيهات البريدية"
        indexes = [
            models.Index(fields=["supervisor"]),
        ]

    supervisor = models.OneToOneField(
        Supervisor,
        on_delete=models.CASCADE,
        related_name="email_preferences",
        verbose_name="المشرف",
    )

    plan_approved = models.BooleanField("تنبيه اعتماد الخطة", default=True)
    plan_returned = models.BooleanField("تنبيه إرجاع الخطة للتعديل", default=True)
    unlock_result = models.BooleanField("تنبيه نتيجة طلب فك الاعتماد", default=True)
    admin_alert = models.BooleanField("التنبيهات الإدارية", default=True)
    control_followup = models.BooleanField("الملاحظات الرقابية", default=True)
    incomplete_reminder = models.BooleanField("تذكير عدم اكتمال الخطة", default=True)
    weekly_summary = models.BooleanField("ملخص أسبوعي", default=False)

    updated_at = models.DateTimeField("آخر تحديث", auto_now=True)

    def __str__(self) -> str:
        return f"تفضيلات البريد — {self.supervisor}"


class EmailNotificationLog(models.Model):
    class Meta:
        verbose_name = "سجل بريد إلكتروني"
        verbose_name_plural = "سجل البريد الإلكتروني"
        ordering = ["-created_at", "-id"]
        indexes = [
            models.Index(fields=["supervisor", "event_type"]),
            models.Index(fields=["status"]),
            models.Index(fields=["created_at"]),
            models.Index(fields=["sent_at"]),
        ]

    EVENT_PLAN_APPROVED = "plan_approved"
    EVENT_PLAN_RETURNED = "plan_returned"
    EVENT_UNLOCK_REQUESTED = "unlock_requested"
    EVENT_UNLOCK_APPROVED = "unlock_approved"
    EVENT_UNLOCK_REJECTED = "unlock_rejected"
    EVENT_ADMIN_ALERT = "admin_alert"
    EVENT_CONTROL_FOLLOWUP = "control_followup"
    EVENT_INCOMPLETE_REMINDER = "incomplete_reminder"
    EVENT_WEEKLY_SUMMARY = "weekly_summary"

    EVENT_CHOICES = [
        (EVENT_PLAN_APPROVED, "اعتماد الخطة"),
        (EVENT_PLAN_RETURNED, "إرجاع الخطة للتعديل"),
        (EVENT_UNLOCK_REQUESTED, "طلب فك اعتماد"),
        (EVENT_UNLOCK_APPROVED, "قبول فك الاعتماد"),
        (EVENT_UNLOCK_REJECTED, "رفض فك الاعتماد"),
        (EVENT_ADMIN_ALERT, "تنبيه إداري"),
        (EVENT_CONTROL_FOLLOWUP, "ملاحظة رقابية"),
        (EVENT_INCOMPLETE_REMINDER, "تذكير بعدم اكتمال الخطة"),
        (EVENT_WEEKLY_SUMMARY, "ملخص أسبوعي"),
    ]

    STATUS_SENT = "sent"
    STATUS_SKIPPED = "skipped"
    STATUS_FAILED = "failed"

    STATUS_CHOICES = [
        (STATUS_SENT, "أُرسلت"),
        (STATUS_SKIPPED, "لم تُرسل"),
        (STATUS_FAILED, "فشل الإرسال"),
    ]

    supervisor = models.ForeignKey(
        Supervisor,
        on_delete=models.CASCADE,
        related_name="email_logs",
        verbose_name="المشرف",
    )
    plan = models.ForeignKey(
        Plan,
        on_delete=models.SET_NULL,
        related_name="email_logs",
        verbose_name="الخطة",
        null=True,
        blank=True,
    )
    event_type = models.CharField("نوع الرسالة", max_length=40, choices=EVENT_CHOICES)
    recipient_email = models.EmailField("البريد المرسل إليه", blank=True, null=True)
    subject = models.CharField("عنوان الرسالة", max_length=250)
    body_preview = models.TextField("مختصر الرسالة", blank=True, null=True)
    status = models.CharField("حالة الإرسال", max_length=20, choices=STATUS_CHOICES, default=STATUS_SKIPPED)
    error_message = models.TextField("رسالة الخطأ", blank=True, null=True)
    sent_at = models.DateTimeField("وقت الإرسال", null=True, blank=True)
    created_at = models.DateTimeField("وقت الإنشاء", auto_now_add=True)

    def __str__(self) -> str:
        return f"{self.get_event_type_display()} — {self.supervisor} — {self.get_status_display()}"

    def clean(self):
        super().clean()
        self.subject = _clean_text(self.subject)
        self.body_preview = _clean_text(self.body_preview) or None
        self.error_message = _clean_text(self.error_message) or None
        if self.recipient_email:
            self.recipient_email = _clean_text(self.recipient_email).lower()

    def save(self, *args, **kwargs):
        self.full_clean()
        return super().save(*args, **kwargs)
