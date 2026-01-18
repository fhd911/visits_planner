from __future__ import annotations

from django.http import HttpRequest, HttpResponse
from openpyxl import Workbook

from .models import Plan, Supervisor


def plan_export_view(request: HttpRequest) -> HttpResponse:
    sup = request.GET.get("sup")
    week = request.GET.get("week")

    if not sup or not week:
        return HttpResponse("بيانات غير مكتملة", status=400)

    supervisor = Supervisor.objects.filter(national_id=sup).first()
    if not supervisor:
        return HttpResponse("مشرف غير موجود", status=404)

    plan = Plan.objects.filter(supervisor=supervisor, week_no=week).first()
    if not plan:
        return HttpResponse("لا توجد خطة", status=404)

    wb = Workbook()
    ws = wb.active
    ws.title = "الخطة"

    ws.append(["اليوم", "المدرسة"])

    for d in plan.days.select_related("school").order_by("weekday"):
        ws.append([d.get_weekday_display(), d.school.name if d.school else ""])

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = f'attachment; filename="plan_week_{plan.week_no}.xlsx"'
    wb.save(response)
    return response
