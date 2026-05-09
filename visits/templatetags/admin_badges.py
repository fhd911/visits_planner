from django import template

register = template.Library()


@register.simple_tag
def pending_unlock_requests_count() -> int:
    """
    عدد طلبات فك الاعتماد التي ما زالت بانتظار القرار الإداري.
    لا يحسب الطلب إلا إذا كان:
    - UnlockRequest.status = pending
    - Plan.status = unlock_requested
    """
    try:
        from visits.models import UnlockRequest, Plan

        pending_status = getattr(UnlockRequest, "STATUS_PENDING", "pending")
        unlock_status = getattr(Plan, "STATUS_UNLOCK_REQUESTED", "unlock_requested")

        return UnlockRequest.objects.filter(
            status=pending_status,
            plan__status=unlock_status,
        ).count()
    except Exception:
        return 0


@register.simple_tag
def pending_unlock_count() -> int:
    return pending_unlock_requests_count()


@register.simple_tag
def unlock_requests_count() -> int:
    return pending_unlock_requests_count()


@register.simple_tag
def admin_pending_unlock_count() -> int:
    return pending_unlock_requests_count()
