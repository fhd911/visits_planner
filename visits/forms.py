from __future__ import annotations

from django import forms

from .models import WeeklyLetterLink


class ImportExcelForm(forms.Form):
    schools_boys = forms.FileField(
        label="ملف المدارس (بنين)",
        required=False,
        help_text="Excel .xlsx",
    )
    schools_girls = forms.FileField(
        label="ملف المدارس (بنات)",
        required=False,
        help_text="Excel .xlsx",
    )
    principals = forms.FileField(
        label="ملف مديري المدارس",
        required=False,
        help_text="Excel .xlsx",
    )
    supervisors = forms.FileField(
        label="ملف المشرفين",
        required=False,
        help_text="Excel .xlsx",
    )
    assignments = forms.FileField(
        label="ملف الإسنادات (مشرف-مدرسة)",
        required=False,
        help_text="Excel .xlsx",
    )

    def clean(self):
        data = super().clean()
        if not any(
            data.get(k)
            for k in ["schools_boys", "schools_girls", "principals", "supervisors", "assignments"]
        ):
            raise forms.ValidationError("ارفع ملفًا واحدًا على الأقل.")
        return data


class WeeklyLetterLinkForm(forms.ModelForm):
    class Meta:
        model = WeeklyLetterLink
        fields = ["week", "title", "drive_url", "note", "is_active"]
        widgets = {
            "week": forms.Select(
                attrs={
                    "class": "form-select",
                }
            ),
            "title": forms.TextInput(
                attrs={
                    "class": "form-control",
                    "placeholder": "مثال: خطاب الأسبوع الخامس",
                }
            ),
            "drive_url": forms.URLInput(
                attrs={
                    "class": "form-control",
                    "placeholder": "https://drive.google.com/...",
                    "dir": "ltr",
                }
            ),
            "note": forms.Textarea(
                attrs={
                    "class": "form-control",
                    "placeholder": "ملاحظة اختيارية",
                    "rows": 4,
                }
            ),
            "is_active": forms.CheckboxInput(
                attrs={
                    "class": "form-check-input",
                }
            ),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        qs = self.fields["week"].queryset.filter(is_break=False).order_by("week_no")
        used_week_ids = list(WeeklyLetterLink.objects.values_list("week_id", flat=True))

        if self.instance and self.instance.pk and self.instance.week_id:
            used_week_ids = [wid for wid in used_week_ids if wid != self.instance.week_id]

        self.fields["week"].queryset = qs.exclude(pk__in=used_week_ids)

        self.fields["week"].label_from_instance = (
            lambda obj: f"الأسبوع {obj.week_no} — يبدأ {obj.start_sunday}"
            + (f" — {obj.title}" if obj.title else "")
        )

    def clean_drive_url(self):
        url = (self.cleaned_data.get("drive_url") or "").strip()
        if not url:
            raise forms.ValidationError("رابط الخطاب مطلوب.")
        return url