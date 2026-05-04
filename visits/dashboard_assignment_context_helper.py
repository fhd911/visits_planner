from django.db.models import Count

from .models import Assignment, School

try:
    from .models import AssignmentReviewLog
except Exception:
    AssignmentReviewLog = None


def get_assignment_dashboard_context():
    """
    مؤشرات مختصرة للوحة الإدارة الرئيسية.

    الاستخدام داخل admin_dashboard_view:
        context.update(get_assignment_dashboard_context())
    """
    active_assignments = Assignment.objects.filter(is_active=True)

    assigned_active_school_ids = (
        active_assignments
        .filter(school__is_active=True)
        .values_list("school_id", flat=True)
        .distinct()
    )

    unassigned_schools_count = (
        School.objects
        .filter(is_active=True)
        .exclude(id__in=assigned_active_school_ids)
        .count()
    )

    duplicate_assignments_count = (
        active_assignments
        .values("school_id")
        .annotate(total=Count("id"))
        .filter(total__gt=1, school_id__isnull=False)
        .count()
    )

    inactive_supervisor_assignments_count = (
        active_assignments
        .filter(supervisor__is_active=False)
        .count()
    )

    inactive_school_assignments_count = (
        active_assignments
        .filter(school__is_active=False)
        .count()
    )

    latest_assignment_log = None
    if AssignmentReviewLog is not None:
        latest_assignment_log = (
            AssignmentReviewLog.objects
            .select_related("school", "supervisor", "user")
            .order_by("-created_at", "-id")
            .first()
        )

    return {
        "dashboard_unassigned_schools_count": unassigned_schools_count,
        "dashboard_duplicate_assignments_count": duplicate_assignments_count,
        "dashboard_inactive_supervisor_assignments_count": inactive_supervisor_assignments_count,
        "dashboard_inactive_school_assignments_count": inactive_school_assignments_count,
        "dashboard_latest_assignment_log": latest_assignment_log,
    }
