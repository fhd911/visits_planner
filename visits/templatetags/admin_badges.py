from __future__ import annotations

from django import template
from django.urls import reverse

from visits.models import Plan, UnlockRequest

register = template.Library()


def _pending_unlock_requests_qs():
    """
    طلبات فك الاعتماد التي لا تزال بانتظار قرار إداري.

    نعتمد على شرطين معًا حتى لا يظهر رقم قديم بعد قبول/رفض الطلب:
    1) UnlockRequest.status = pending
    2) Plan.status = unlock_requested
    """
    pending_status = getattr(UnlockRequest, "STATUS_PENDING", "pending")
    unlock_plan_status = getattr(Plan, "STATUS_UNLOCK_REQUESTED", "unlock_requested")

    return UnlockRequest.objects.filter(
        status=pending_status,
        plan__status=unlock_plan_status,
    )


def _pending_unlock_requests_count() -> int:
    try:
        return _pending_unlock_requests_qs().count()
    except Exception:
        return 0


@register.simple_tag
def pending_unlock_requests_count() -> int:
    """عدد طلبات فك الاعتماد المعلقة لعرضها في الهيدر."""
    return _pending_unlock_requests_count()


@register.simple_tag
def has_pending_unlock_requests() -> bool:
    """هل يوجد طلب فك اعتماد معلق؟"""
    return _pending_unlock_requests_count() > 0


@register.simple_tag
def pending_unlock_requests_url() -> str:
    """رابط مباشر للوحة الإدارة مع فلتر طلبات فك الاعتماد."""
    try:
        return f"{reverse('visits:admin_dashboard')}?status=unlock"
    except Exception:
        return "/manager/dashboard/?status=unlock"
