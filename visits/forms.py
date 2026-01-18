# visits/forms.py  (كود كامل)

from __future__ import annotations

from django import forms


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
