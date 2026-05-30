from __future__ import annotations

from datetime import datetime, timedelta
from typing import Any

from django.contrib import messages
from django.contrib.admin.views.decorators import staff_member_required
from django.db import transaction
from django.http import HttpRequest, HttpResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone
from django.views.decorators.http import require_http_methods

from .models import AcademicYear, Semester, PlanWeek, PlanClosedDay


WEEKDAYS = [
    (0, "الأحد"),
    (1, "الإثنين"),
    (2, "الثلاثاء"),
    (3, "الأربعاء"),
    (4, "الخميس"),
]


def _clean(value: object) -> str:
    return str(value or "").strip()


def _safe_int(value: object, default: int = 0) -> int:
    try:
        return int(str(value).strip())
    except Exception:
        return default


def _parse_date(value: object):
    value = _clean(value)
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except Exception:
        return None


def _date_value(value) -> str:
    return value.strftime("%Y-%m-%d") if value else ""


def _next_available_week_no(start: int | None = None) -> int:
    used = set(PlanWeek.objects.values_list("week_no", flat=True))
    n = max(1, int(start or 1))
    if not start:
        n = (max(used) + 1) if used else 1
    while n in used:
        n += 1
    return n


def _semester_default_end(starts_at, weeks_count: int):
    if not starts_at:
        return None
    # الأسبوع يبدأ الأحد وينتهي الخميس، ونهاية الفصل تكون الخميس من آخر أسبوع.
    return starts_at + timedelta(days=(weeks_count - 1) * 7 + 4)


def _generate_weeks_for_semester(
    *,
    semester: Semester,
    start_sunday,
    weeks_count: int,
    first_week_no: int | None = None,
) -> int:
    """
    ينشئ أسابيع الفصل أو يحدثها إن كانت موجودة.
    يعيد رقم الأسبوع التالي المتاح بعد آخر أسبوع تم إنشاؤه.
    """
    if not start_sunday:
        raise ValueError("تاريخ بداية أول أسبوع مطلوب.")

    weeks_count = max(1, min(int(weeks_count or semester.weeks_count or 19), 25))
    next_week_no = _next_available_week_no(first_week_no)

    for idx in range(1, weeks_count + 1):
        week_start = start_sunday + timedelta(days=(idx - 1) * 7)
        existing = PlanWeek.objects.filter(semester=semester, semester_week_no=idx).first()

        if existing:
            existing.academic_year = semester.academic_year
            existing.start_sunday = week_start
            existing.title = existing.title or f"الأسبوع {idx}"
            existing.is_break = False
            existing.save()
            continue

        week_no = _next_available_week_no(next_week_no)
        PlanWeek.objects.create(
            week_no=week_no,
            academic_year=semester.academic_year,
            semester=semester,
            semester_week_no=idx,
            start_sunday=week_start,
            title=f"الأسبوع {idx}",
            is_break=False,
            is_current=False,
            is_open_for_supervisors=False,
        )
        next_week_no = week_no + 1

    return next_week_no


def _set_current_week(week: PlanWeek, *, open_for_supervisors: bool = True) -> None:
    with transaction.atomic():
        PlanWeek.objects.exclude(pk=week.pk).update(
            is_current=False,
            is_open_for_supervisors=False,
        )
        week.is_break = False
        week.is_current = True
        week.is_open_for_supervisors = bool(open_for_supervisors)
        week.save()

        if week.semester_id:
            Semester.objects.exclude(pk=week.semester_id).update(is_current=False)
            semester = week.semester
            semester.is_current = True
            semester.is_open = True
            semester.save()

            _set_current_year(semester.academic_year)
        elif week.academic_year_id:
            _set_current_year(week.academic_year)




def _set_current_year(year: AcademicYear) -> None:
    """يجعل عامًا واحدًا فقط هو العام الحالي."""
    AcademicYear.objects.exclude(pk=year.pk).update(is_current=False)
    year.is_current = True
    year.is_active = True
    year.save()


def _set_current_semester(semester: Semester) -> None:
    """يجعل فصلًا واحدًا فقط هو الفصل الحالي ويربط عامه كعام حالي."""
    Semester.objects.exclude(pk=semester.pk).update(is_current=False)
    semester.is_current = True
    semester.is_open = True
    semester.save()
    _set_current_year(semester.academic_year)

def _selected_objects(request: HttpRequest) -> dict[str, Any]:
    current_week = (
        PlanWeek.objects.select_related("academic_year", "semester", "semester__academic_year")
        .filter(is_current=True)
        .first()
    )
    current_year = AcademicYear.objects.filter(is_current=True).first()
    current_semester = Semester.objects.select_related("academic_year").filter(is_current=True).first()

    year_id = _safe_int(request.GET.get("year") or request.POST.get("year_id"), 0)
    semester_id = _safe_int(request.GET.get("semester") or request.POST.get("semester_id"), 0)
    week_id = _safe_int(request.GET.get("week_id") or request.POST.get("week_id"), 0)

    selected_year = None
    if year_id:
        selected_year = AcademicYear.objects.filter(pk=year_id).first()
    if not selected_year:
        selected_year = current_year or AcademicYear.objects.order_by("-starts_at", "-id").first()

    semesters_qs = Semester.objects.select_related("academic_year").all()
    if selected_year:
        semesters_qs = semesters_qs.filter(academic_year=selected_year)
    semesters = list(semesters_qs.order_by("number", "starts_at"))

    selected_semester = None
    if semester_id:
        selected_semester = Semester.objects.filter(pk=semester_id).first()
    if not selected_semester and current_semester and selected_year and current_semester.academic_year_id == selected_year.id:
        selected_semester = current_semester
    if not selected_semester and semesters:
        selected_semester = semesters[0]

    weeks_qs = PlanWeek.objects.select_related("academic_year", "semester").all()
    if selected_semester:
        weeks_qs = weeks_qs.filter(semester=selected_semester)
    elif selected_year:
        weeks_qs = weeks_qs.filter(academic_year=selected_year)
    weeks = list(weeks_qs.order_by("week_no"))

    selected_week = None
    if week_id:
        selected_week = PlanWeek.objects.filter(pk=week_id).first()
    if not selected_week and current_week:
        if not weeks or any(w.id == current_week.id for w in weeks):
            selected_week = current_week
    if not selected_week and weeks:
        selected_week = weeks[0]

    return {
        "current_week": current_week,
        "current_year": current_year,
        "current_semester": current_semester,
        "selected_year": selected_year,
        "selected_semester": selected_semester,
        "selected_week": selected_week,
        "semesters": semesters,
        "weeks": weeks,
    }


def _redirect_to_page(*, year_id: int | None = None, semester_id: int | None = None, week_id: int | None = None) -> HttpResponse:
    params = []
    if year_id:
        params.append(f"year={year_id}")
    if semester_id:
        params.append(f"semester={semester_id}")
    if week_id:
        params.append(f"week_id={week_id}")
    url = reverse("visits:admin_academic_plan")
    return redirect(f"{url}?{'&'.join(params)}" if params else url)


@staff_member_required(login_url="visits:admin_login")
@require_http_methods(["GET", "POST"])
def admin_academic_plan_view(request: HttpRequest) -> HttpResponse:
    """صفحة مستقلة لإدارة العام الدراسي والفصول والأسابيع والأيام المغلقة."""

    if request.method == "POST":
        action = _clean(request.POST.get("action"))

        try:
            with transaction.atomic():
                if action == "create_full_year":
                    year_name = _clean(request.POST.get("year_name"))
                    weeks_count = _safe_int(request.POST.get("weeks_count"), 19) or 19
                    sem1_start = _parse_date(request.POST.get("semester1_start"))
                    sem2_start = _parse_date(request.POST.get("semester2_start"))
                    first_week_no = _safe_int(request.POST.get("first_week_no"), 0) or None

                    if not year_name:
                        messages.error(request, "اسم العام الدراسي مطلوب.")
                        return _redirect_to_page()
                    if not sem1_start or not sem2_start:
                        messages.error(request, "تاريخ بداية الفصل الأول والفصل الثاني مطلوبان.")
                        return _redirect_to_page()

                    year, _created = AcademicYear.objects.update_or_create(
                        name=year_name,
                        defaults={
                            "starts_at": _parse_date(request.POST.get("year_starts_at")) or sem1_start,
                            "ends_at": _parse_date(request.POST.get("year_ends_at")) or _semester_default_end(sem2_start, weeks_count),
                            "is_current": request.POST.get("make_current") == "1",
                            "is_active": True,
                        },
                    )

                    sem1, _ = Semester.objects.update_or_create(
                        academic_year=year,
                        number=Semester.FIRST,
                        defaults={
                            "title": "الفصل الدراسي الأول",
                            "starts_at": sem1_start,
                            "ends_at": _semester_default_end(sem1_start, weeks_count),
                            "weeks_count": weeks_count,
                            "is_current": False,
                            "is_open": False,
                        },
                    )
                    sem2, _ = Semester.objects.update_or_create(
                        academic_year=year,
                        number=Semester.SECOND,
                        defaults={
                            "title": "الفصل الدراسي الثاني",
                            "starts_at": sem2_start,
                            "ends_at": _semester_default_end(sem2_start, weeks_count),
                            "weeks_count": weeks_count,
                            "is_current": False,
                            "is_open": False,
                        },
                    )

                    if request.POST.get("generate_weeks") == "1":
                        next_no = _generate_weeks_for_semester(
                            semester=sem1,
                            start_sunday=sem1_start,
                            weeks_count=weeks_count,
                            first_week_no=first_week_no,
                        )
                        _generate_weeks_for_semester(
                            semester=sem2,
                            start_sunday=sem2_start,
                            weeks_count=weeks_count,
                            first_week_no=next_no,
                        )

                    if request.POST.get("make_current") == "1":
                        _set_current_year(year)

                    messages.success(request, "تم إنشاء/تحديث العام الدراسي والفصلين بنجاح.")
                    return _redirect_to_page(year_id=year.id)

                if action == "create_year":
                    name = _clean(request.POST.get("name"))
                    if not name:
                        messages.error(request, "اسم العام الدراسي مطلوب.")
                        return _redirect_to_page()

                    year = AcademicYear.objects.create(
                        name=name,
                        starts_at=_parse_date(request.POST.get("starts_at")),
                        ends_at=_parse_date(request.POST.get("ends_at")),
                        is_current=False,
                        is_active=True,
                    )
                    if request.POST.get("is_current") == "1":
                        _set_current_year(year)
                    messages.success(request, "تمت إضافة العام الدراسي بنجاح.")
                    return _redirect_to_page(year_id=year.id)

                if action == "set_current_year":
                    year = get_object_or_404(AcademicYear, pk=_safe_int(request.POST.get("year_id"), 0))
                    _set_current_year(year)
                    messages.success(request, "تم تعيين العام الدراسي كعام حالي وإلغاء تعيين باقي الأعوام.")
                    return _redirect_to_page(year_id=year.id)

                if action == "create_semester":
                    year = get_object_or_404(AcademicYear, pk=_safe_int(request.POST.get("year_id"), 0))
                    number = _safe_int(request.POST.get("number"), 1)
                    starts_at = _parse_date(request.POST.get("starts_at"))
                    weeks_count = _safe_int(request.POST.get("weeks_count"), 19) or 19
                    if number not in (Semester.FIRST, Semester.SECOND):
                        messages.error(request, "الفصل الدراسي غير صحيح.")
                        return _redirect_to_page(year_id=year.id)
                    if not starts_at:
                        messages.error(request, "تاريخ بداية الفصل مطلوب.")
                        return _redirect_to_page(year_id=year.id)

                    semester, _created = Semester.objects.update_or_create(
                        academic_year=year,
                        number=number,
                        defaults={
                            "title": _clean(request.POST.get("title")) or dict(Semester.SEMESTER_CHOICES).get(number),
                            "starts_at": starts_at,
                            "ends_at": _parse_date(request.POST.get("ends_at")) or _semester_default_end(starts_at, weeks_count),
                            "weeks_count": weeks_count,
                            "is_current": False,
                            "is_open": request.POST.get("is_open") == "1",
                        },
                    )
                    if request.POST.get("is_current") == "1":
                        _set_current_semester(semester)
                    messages.success(request, "تم حفظ الفصل الدراسي بنجاح.")
                    return _redirect_to_page(year_id=year.id, semester_id=semester.id)

                if action == "set_current_semester":
                    semester = get_object_or_404(Semester, pk=_safe_int(request.POST.get("semester_id"), 0))
                    _set_current_semester(semester)
                    messages.success(request, "تم تعيين الفصل الدراسي كفصل حالي وإلغاء تعيين باقي الفصول.")
                    return _redirect_to_page(year_id=semester.academic_year_id, semester_id=semester.id)

                if action == "generate_semester_weeks":
                    semester = get_object_or_404(Semester, pk=_safe_int(request.POST.get("semester_id"), 0))
                    start_sunday = _parse_date(request.POST.get("start_sunday")) or semester.starts_at
                    weeks_count = _safe_int(request.POST.get("weeks_count"), semester.weeks_count) or semester.weeks_count
                    first_week_no = _safe_int(request.POST.get("first_week_no"), 0) or None
                    _generate_weeks_for_semester(
                        semester=semester,
                        start_sunday=start_sunday,
                        weeks_count=weeks_count,
                        first_week_no=first_week_no,
                    )
                    messages.success(request, "تم توليد أسابيع الفصل الدراسي بنجاح.")
                    return _redirect_to_page(year_id=semester.academic_year_id, semester_id=semester.id)

                if action == "link_existing_weeks":
                    semester = get_object_or_404(
                        Semester,
                        pk=_safe_int(request.POST.get("semester_id"), 0),
                    )

                    from_week_no = _safe_int(request.POST.get("from_week_no"), 1) or 1
                    to_week_no = _safe_int(request.POST.get("to_week_no"), from_week_no) or from_week_no
                    first_semester_week_no = _safe_int(request.POST.get("first_semester_week_no"), 1) or 1

                    if from_week_no < 1 or to_week_no < 1:
                        messages.error(request, "نطاق الأسابيع غير صحيح.")
                        return _redirect_to_page(
                            year_id=semester.academic_year_id,
                            semester_id=semester.id,
                        )

                    if to_week_no < from_week_no:
                        messages.error(request, "رقم نهاية النطاق يجب أن يكون أكبر من أو يساوي رقم البداية.")
                        return _redirect_to_page(
                            year_id=semester.academic_year_id,
                            semester_id=semester.id,
                        )

                    expected_week_numbers = list(range(from_week_no, to_week_no + 1))
                    weeks = list(
                        PlanWeek.objects.filter(
                            week_no__gte=from_week_no,
                            week_no__lte=to_week_no,
                        ).order_by("week_no")
                    )
                    existing_numbers = {week.week_no for week in weeks}
                    missing_numbers = [num for num in expected_week_numbers if num not in existing_numbers]

                    if missing_numbers:
                        messages.error(
                            request,
                            "يوجد أسابيع غير منشأة داخل النطاق المحدد: "
                            + ", ".join(str(num) for num in missing_numbers[:12])
                            + (" ..." if len(missing_numbers) > 12 else "")
                            + ". عدّل النطاق أو أنشئ الأسابيع الناقصة أولًا.",
                        )
                        return _redirect_to_page(
                            year_id=semester.academic_year_id,
                            semester_id=semester.id,
                        )

                    if not weeks:
                        messages.warning(request, "لا توجد أسابيع ضمن النطاق المحدد لربطها بالفصل.")
                        return _redirect_to_page(
                            year_id=semester.academic_year_id,
                            semester_id=semester.id,
                        )

                    max_semester_week_no = first_semester_week_no + len(weeks) - 1
                    if max_semester_week_no > 25:
                        messages.error(request, "رقم الأسبوع داخل الفصل لا يمكن أن يتجاوز 25 أسبوعًا.")
                        return _redirect_to_page(
                            year_id=semester.academic_year_id,
                            semester_id=semester.id,
                        )

                    if max_semester_week_no > semester.weeks_count:
                        semester.weeks_count = max_semester_week_no
                        semester.save(update_fields=["weeks_count", "updated_at"])

                    for offset, week in enumerate(weeks):
                        semester_week_no = first_semester_week_no + offset
                        PlanWeek.objects.filter(pk=week.pk).update(
                            academic_year=semester.academic_year,
                            semester=semester,
                            semester_week_no=semester_week_no,
                        )

                    messages.success(
                        request,
                        f"تم ربط {len(weeks)} أسبوعًا موجودًا بالفصل الدراسي دون إنشاء أسابيع جديدة.",
                    )
                    return _redirect_to_page(
                        year_id=semester.academic_year_id,
                        semester_id=semester.id,
                        week_id=weeks[0].id,
                    )

                if action == "set_current_week":
                    week = get_object_or_404(PlanWeek, pk=_safe_int(request.POST.get("week_id"), 0))
                    _set_current_week(week, open_for_supervisors=request.POST.get("open_for_supervisors") == "1")
                    messages.success(request, "تم تعيين الأسبوع الحالي للمشرفين بنجاح.")
                    return _redirect_to_page(year_id=week.academic_year_id, semester_id=week.semester_id, week_id=week.id)

                if action == "toggle_week_open":
                    week = get_object_or_404(PlanWeek, pk=_safe_int(request.POST.get("week_id"), 0))
                    if week.is_open_for_supervisors:
                        week.is_open_for_supervisors = False
                        week.save(update_fields=["is_open_for_supervisors"])
                        messages.success(request, "تم إغلاق الأسبوع عن المشرفين.")
                    else:
                        _set_current_week(week, open_for_supervisors=True)
                        messages.success(request, "تم فتح الأسبوع وتعيينه كأسبوع حالي، وإغلاق بقية الأسابيع تلقائيًا.")
                    return _redirect_to_page(year_id=week.academic_year_id, semester_id=week.semester_id, week_id=week.id)

                if action == "toggle_week_break":
                    week = get_object_or_404(PlanWeek, pk=_safe_int(request.POST.get("week_id"), 0))
                    week.is_break = not bool(week.is_break)
                    if week.is_break:
                        week.is_current = False
                        week.is_open_for_supervisors = False
                    week.save()
                    messages.success(request, "تم تحديث حالة الأسبوع.")
                    return _redirect_to_page(year_id=week.academic_year_id, semester_id=week.semester_id, week_id=week.id)

                if action == "close_day":
                    week = get_object_or_404(PlanWeek, pk=_safe_int(request.POST.get("week_id"), 0))
                    if week.is_break:
                        messages.error(request, "لا يمكن إغلاق يوم داخل أسبوع محدد كإجازة كاملة.")
                        return _redirect_to_page(year_id=week.academic_year_id, semester_id=week.semester_id, week_id=week.id)
                    weekday = _safe_int(request.POST.get("weekday"), 0)
                    if weekday not in dict(WEEKDAYS):
                        messages.error(request, "اليوم المحدد غير صحيح.")
                        return _redirect_to_page(year_id=week.academic_year_id, semester_id=week.semester_id, week_id=week.id)

                    reason_type = _clean(request.POST.get("reason_type")) or PlanClosedDay.OFFICIAL_HOLIDAY
                    reason_title = _clean(request.POST.get("reason_title")) or dict(PlanClosedDay.REASON_CHOICES).get(reason_type, "إجازة رسمية")
                    closed_date = _parse_date(request.POST.get("date"))

                    PlanClosedDay.objects.update_or_create(
                        week=week,
                        weekday=weekday,
                        defaults={
                            "date": closed_date,
                            "reason_type": reason_type,
                            "reason_title": reason_title,
                            "count_as_completed": True,
                            "is_active": True,
                        },
                    )
                    messages.success(request, "تم إغلاق اليوم داخل الأسبوع بنجاح.")
                    return _redirect_to_page(year_id=week.academic_year_id, semester_id=week.semester_id, week_id=week.id)

                if action == "open_day":
                    closed_day = get_object_or_404(PlanClosedDay, pk=_safe_int(request.POST.get("closed_day_id"), 0))
                    week = closed_day.week
                    closed_day.is_active = False
                    closed_day.save()
                    messages.success(request, "تم فتح اليوم وإلغاء الإغلاق النشط.")
                    return _redirect_to_page(year_id=week.academic_year_id, semester_id=week.semester_id, week_id=week.id)

                messages.warning(request, "لم يتم التعرف على الإجراء المطلوب.")
                return _redirect_to_page()

        except Exception as exc:
            messages.error(request, f"تعذر تنفيذ العملية: {exc}")
            return _redirect_to_page()

    data = _selected_objects(request)
    selected_week = data["selected_week"]

    closed_days_by_weekday = {}
    if selected_week:
        closed_days_by_weekday = {
            item.weekday: item
            for item in PlanClosedDay.objects.filter(week=selected_week, is_active=True).order_by("weekday")
        }

    current_week_closed_count = 0
    if data.get("current_week"):
        current_week_closed_count = PlanClosedDay.objects.filter(week=data["current_week"], is_active=True).count()

    selected_week_closed_count = len(closed_days_by_weekday)

    day_rows = []
    for weekday, weekday_name in WEEKDAYS:
        day_date = selected_week.start_sunday + timedelta(days=weekday) if selected_week else None
        day_rows.append(
            {
                "weekday": weekday,
                "weekday_name": weekday_name,
                "date": day_date,
                "date_value": _date_value(day_date),
                "closed_day": closed_days_by_weekday.get(weekday),
            }
        )

    week_rows = []
    for week in data["weeks"]:
        active_closed_count = week.closed_days.filter(is_active=True).count()
        week_rows.append(
            {
                "week": week,
                "active_closed_count": active_closed_count,
                "status_label": "إجازة/توقف" if week.is_break else "مفتوح" if week.is_open_for_supervisors else "مغلق",
            }
        )

    all_weeks_qs = PlanWeek.objects.order_by("week_no")
    first_existing_week = all_weeks_qs.first()
    last_existing_week = all_weeks_qs.last()
    existing_weeks_count = all_weeks_qs.count()
    min_existing_week_no = first_existing_week.week_no if first_existing_week else 1
    max_existing_week_no = last_existing_week.week_no if last_existing_week else 0

    context = {
        **data,
        "years": AcademicYear.objects.order_by("-is_current", "-starts_at", "-id"),
        "semester_choices": Semester.SEMESTER_CHOICES,
        "reason_choices": PlanClosedDay.REASON_CHOICES,
        "week_rows": week_rows,
        "day_rows": day_rows,
        "today": timezone.localdate(),
        "selected_week_closed_count": selected_week_closed_count,
        "current_week_closed_count": current_week_closed_count,
        "existing_weeks_count": existing_weeks_count,
        "min_existing_week_no": min_existing_week_no,
        "max_existing_week_no": max_existing_week_no,
        "page_title": "إدارة الخطة الدراسية",
        "page_subtitle": "إعداد العام الدراسي والفصول والأسابيع والأيام المغلقة.",
    }
    return render(request, "visits/admin_academic_plan.html", context)
