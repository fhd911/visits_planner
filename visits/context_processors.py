# visits/context_processors.py

def supervisor_control_followups_badge(request):
    """
    يوفر عداد الملاحظات الرقابية النشطة لجميع قوالب المشرف.
    يستخدم لإظهار زر "الملاحظات" في الشريط العلوي فقط عند وجود ملاحظات تحتاج متابعة.
    """
    try:
        sid = request.session.get("visits_sup_id")
        if not sid:
            return {
                "supervisor_active_control_followups_count": 0,
                "has_supervisor_active_control_followups": False,
            }

        from .models import ControlFollowUp

        active_statuses = [
            ControlFollowUp.STATUS_OPEN,
            ControlFollowUp.STATUS_NOTIFIED,
            ControlFollowUp.STATUS_PENDING_ADMIN,
        ]
        count = ControlFollowUp.objects.filter(
            supervisor_id=sid,
            status__in=active_statuses,
        ).count()
        return {
            "supervisor_active_control_followups_count": count,
            "has_supervisor_active_control_followups": count > 0,
        }
    except Exception:
        return {
            "supervisor_active_control_followups_count": 0,
            "has_supervisor_active_control_followups": False,
        }
