from __future__ import annotations

from django.contrib import messages
from django.contrib.admin.views.decorators import staff_member_required
from django.http import HttpRequest, HttpResponse
from django.shortcuts import render
from django.urls import reverse

from .forms import ImportExcelForm
from .importers import import_schools, import_principals, import_supervisors, import_assignments


def home_view(request: HttpRequest) -> HttpResponse:
    return HttpResponse(
        f"""
        <html lang="ar" dir="rtl">
        <head>
          <meta charset="utf-8">
          <title>بوابة خطط الزيارات</title>
        </head>
        <body style="font-family:Tahoma,Arial; padding:24px;">
          <h2>بوابة خطط الزيارات</h2>
          <p>روابط سريعة:</p>
          <ul>
            <li><a href="{reverse('admin:index')}">لوحة الإدارة</a></li>
            <li><a href="{reverse('visits:manager_import')}">استيراد ملفات Excel</a></li>
            <li><a href="/plan/?sup=ضع_سجل_المشرف&week=1">خطة المشرف (تجربة)</a></li>
          </ul>
        </body>
        </html>
        """,
    )


@staff_member_required
def import_view(request: HttpRequest) -> HttpResponse:
    if request.method == "POST":
        form = ImportExcelForm(request.POST, request.FILES)
        if form.is_valid():
            results = {}

            if f := form.cleaned_data.get("schools_boys"):
                results["schools_boys"] = import_schools(f, gender="boys")
            if f := form.cleaned_data.get("schools_girls"):
                results["schools_girls"] = import_schools(f, gender="girls")
            if f := form.cleaned_data.get("principals"):
                results["principals"] = import_principals(f)
            if f := form.cleaned_data.get("supervisors"):
                results["supervisors"] = import_supervisors(f)
            if f := form.cleaned_data.get("assignments"):
                results["assignments"] = import_assignments(f)

            messages.success(request, "تم الاستيراد بنجاح.")
            return render(request, "visits/manager_import.html", {"form": ImportExcelForm(), "results": results})

    else:
        form = ImportExcelForm()

    return render(request, "visits/manager_import.html", {"form": form})
