from django.urls import path
from django.shortcuts import redirect
from django.http import HttpResponse
from django.contrib import messages

from . import views
from . import views_import
from . import views_assignment_review



# =========================================================
# مسارات توافقية للصفحات الجديدة
# الهدف: لا يتوقف python manage.py check إذا كان views.py ناقصًا في بعض الدوال.
# إذا كانت الدالة موجودة في views.py تُستخدم تلقائيًا، وإذا لم تكن موجودة يُستخدم بديل آمن.
# =========================================================

def _view_or(name, fallback):
    return getattr(views, name, fallback)


def _is_staff_user(request):
    return bool(
        getattr(request, "user", None)
        and request.user.is_authenticated
        and request.user.is_staff
    )


def _fallback_admin_reports_view(request):
    if not _is_staff_user(request):
        return redirect("visits:admin_login")
    return views.admin_dashboard_view(request)


def _fallback_admin_control_report_view(request, report_type="incomplete_plans", *args, **kwargs):
    if not _is_staff_user(request):
        return redirect("visits:admin_login")
    return redirect("visits:admin_reports")


def _fallback_admin_control_report_export_excel_view(request, report_type="incomplete_plans", *args, **kwargs):
    if not _is_staff_user(request):
        return redirect("visits:admin_login")
    response = HttpResponse(content_type="text/csv; charset=utf-8")
    response["Content-Disposition"] = f'attachment; filename="control_report_{report_type}.csv"'
    response.write("\ufeff")
    response.write("التقرير,الحالة\n")
    response.write(f"{report_type},لا توجد بيانات مصدرة\n")
    return response


def _fallback_admin_control_report_notify_view(request, report_type="incomplete_plans", *args, **kwargs):
    if not _is_staff_user(request):
        return redirect("visits:admin_login")
    if request.method == "POST":
        messages.info(request, "مسار التنبيه موجود، لكن دالة التنبيه التفصيلية غير مفعلة في views.py.")
    return redirect("visits:admin_reports")


def _fallback_admin_control_followups_view(request):
    if not _is_staff_user(request):
        return redirect("visits:admin_login")
    return redirect("visits:admin_reports")


def _fallback_admin_control_followups_export_excel_view(request):
    if not _is_staff_user(request):
        return redirect("visits:admin_login")
    response = HttpResponse(content_type="text/csv; charset=utf-8")
    response["Content-Disposition"] = 'attachment; filename="control_followups.csv"'
    response.write("\ufeff")
    response.write("سجل الملحوظات\n")
    return response


def _fallback_admin_control_followup_notify_view(request, pk=None, *args, **kwargs):
    if not _is_staff_user(request):
        return redirect("visits:admin_login")
    if request.method == "POST":
        messages.info(request, "مسار تنبيه الملحوظة موجود، لكن دالة التنفيذ غير مفعلة في views.py.")
    return redirect("visits:admin_control_followups")


def _fallback_admin_control_followup_update_view(request, pk=None, *args, **kwargs):
    if not _is_staff_user(request):
        return redirect("visits:admin_login")
    if request.method == "POST":
        messages.info(request, "مسار تحديث الملحوظة موجود، لكن دالة التنفيذ غير مفعلة في views.py.")
    return redirect("visits:admin_control_followups")


def _fallback_supervisor_control_followups_view(request):
    return redirect("visits:supervisor_dashboard")


def _supervisor_control_followup_respond_route(request, pk=None, followup_id=None, *args, **kwargs):
    target_id = followup_id if followup_id is not None else pk
    func = getattr(views, "supervisor_control_followup_respond_view", None)
    if func is None:
        func = getattr(views, "supervisor_control_followup_response_view", None)
    if func is None:
        return redirect("visits:supervisor_control_followups")
    return func(request, target_id)

app_name = "visits"

urlpatterns = [
    # =========================================================
    # صفحة الصيانة
    # =========================================================
    path("maintenance/", views.maintenance_page_view, name="maintenance_page"),

    # =========================================================
    # دخول / خروج
    # =========================================================
    path("", views.login_view, name="login"),
    path("admin-login/", views.admin_login_view, name="admin_login"),
    path("logout/", views.logout_view, name="logout"),

    # =========================================================
    # بوابة المشرف
    # =========================================================
    path("dashboard/", views.supervisor_dashboard_view, name="supervisor_dashboard"),
    path("email-settings/", views.supervisor_email_settings_view, name="supervisor_email_settings"),
    path("email-settings/verify/", views.supervisor_email_verify_view, name="supervisor_email_verify"),
    path("email-settings/resend-otp/", views.supervisor_email_resend_otp_view, name="supervisor_email_resend_otp"),
    path("email-settings/toggle/", views.toggle_email_notifications_view, name="toggle_email_notifications"),

    path("print-assignment-letter/", views.print_assignment_letter_view, name="print_assignment_letter"),
    path("weekly-letter/", views.current_week_letter_redirect_view, name="current_week_letter"),

    path("plan/", views.plan_view, name="plan"),
    path("plan/export/", views.export_plan_excel, name="plan_export"),
    path(
        "plan/export/planned-schools/",
        views.export_plan_planned_schools_excel,
        name="plan_export_planned_schools",
    ),
    path(
        "plan/export/unplanned-schools/",
        views.export_plan_unplanned_schools_excel,
        name="plan_export_unplanned_schools",
    ),
    path(
        "plan/assignments/export/",
        views.export_supervisor_assignments_excel,
        name="export_supervisor_assignments_excel",
    ),
    path("plan/unlock/", views.request_unlock_view, name="request_unlock"),

    path("control-followups/", _view_or("supervisor_control_followups_view", _fallback_supervisor_control_followups_view), name="supervisor_control_followups"),
    path(
        "control-followups/<int:pk>/respond/",
        _supervisor_control_followup_respond_route,
        name="supervisor_control_followup_respond",
    ),

    # تتبع تنفيذ الزيارات للمشرف
    path("plan/visit-status/", views.supervisor_visit_status_view, name="supervisor_visit_status"),
    path("plan/day/<int:day_id>/toggle-visited/", views.toggle_day_visited_view, name="toggle_day_visited"),

    # إشعارات المشرف
    path("notifications/", views.notifications_view, name="notifications"),
    path("notifications/read/<int:notif_id>/", views.mark_notification_read_view, name="mark_notification_read"),
    path("notifications/read-all/", views.mark_all_notifications_read_view, name="mark_all_notifications_read"),

    # =========================================================
    # الإدارة - لوحة الخطط
    # =========================================================
    path("manager/dashboard/", views.admin_dashboard_view, name="admin_dashboard"),
    path("manager/reports/", _view_or("admin_reports_view", _fallback_admin_reports_view), name="admin_reports"),
    path(
        "manager/reports/control/<str:report_type>/",
        _view_or("admin_control_report_view", _fallback_admin_control_report_view),
        name="admin_control_report",
    ),
    path(
        "manager/reports/control/<str:report_type>/export.xlsx",
        _view_or("admin_control_report_export_excel_view", _fallback_admin_control_report_export_excel_view),
        name="admin_control_report_export_excel",
    ),
    path(
        "manager/reports/control/<str:report_type>/notify/",
        _view_or("admin_control_report_notify_view", _fallback_admin_control_report_notify_view),
        name="admin_control_report_notify",
    ),
    path("manager/control-followups/", _view_or("admin_control_followups_view", _fallback_admin_control_followups_view), name="admin_control_followups"),
    path(
        "manager/control-followups/export.xlsx",
        _view_or("admin_control_followups_export_excel_view", _fallback_admin_control_followups_export_excel_view),
        name="admin_control_followups_export_excel",
    ),
    path(
        "manager/control-followups/<int:pk>/notify/",
        _view_or("admin_control_followup_notify_view", _fallback_admin_control_followup_notify_view),
        name="admin_control_followup_notify",
    ),
    path(
        "manager/control-followups/<int:pk>/update/",
        _view_or("admin_control_followup_update_view", _fallback_admin_control_followup_update_view),
        name="admin_control_followup_update",
    ),
    path("manager/plan/<int:plan_id>/", views.admin_plan_detail_view, name="admin_plan_detail"),
    path("manager/export-week/", views.admin_export_week_excel, name="admin_export_week"),
    path(
        "manager/export-all-plans/",
        views.admin_export_all_plans_excel,
        name="admin_export_all_plans_excel",
    ),
    path(
        "manager/export-week-visit-summary/",
        views.admin_export_week_visit_summary_excel,
        name="admin_export_week_visit_summary",
    ),
    path("manager/plan-export/<int:plan_id>/", views.admin_plan_export_excel, name="admin_plan_export_excel"),
    path("manager/plan-approve/<int:plan_id>/", views.admin_plan_approve_view, name="admin_plan_approve"),
    path("manager/plan-draft/<int:plan_id>/", views.admin_plan_back_to_draft_view, name="admin_plan_back_to_draft"),

    # تنبيهات الإدارة
    path("manager/plan-notify/<int:plan_id>/", views.admin_send_notification_view, name="admin_send_notification"),
    path("manager/notify-all/", views.admin_send_notification_all_view, name="admin_send_notification_all"),

    # =========================================================
    # الإدارة - متابعة الزيارات العامة
    # =========================================================
    path(
        "manager/visit-followup/",
        views.admin_visit_followup_dashboard_view,
        name="admin_visit_followup_dashboard",
    ),
    path(
        "manager/visit-followup/export-excel/",
        views.admin_visit_followup_export_excel_view,
        name="admin_visit_followup_export_excel",
    ),
    path(
        "manager/visit-followup/notify/<int:supervisor_id>/",
        views.admin_notify_supervisor_visit_followup_view,
        name="admin_notify_supervisor_visit_followup",
    ),
    path(
        "manager/visit-followup/notify-all/",
        views.admin_notify_all_supervisors_visit_followup_view,
        name="admin_notify_all_supervisors_visit_followup",
    ),

    # طلبات فك الاعتماد
    path("manager/unlock-approve/<int:plan_id>/", views.admin_unlock_approve_view, name="admin_unlock_approve"),
    path("manager/unlock-reject/<int:plan_id>/", views.admin_unlock_reject_view, name="admin_unlock_reject"),

    # =========================================================
    # الإدارة - الصيانة
    # =========================================================
    path("manager/maintenance/", views.admin_maintenance_settings_view, name="admin_maintenance_settings"),
    path("manager/maintenance/toggle/", views.admin_toggle_maintenance_view, name="admin_toggle_maintenance"),
    path(
        "manager/maintenance/message/",
        views.admin_update_maintenance_message_view,
        name="admin_update_maintenance_message",
    ),

    # =========================================================
    # الإدارة - المدارس
    # =========================================================
    path("manager/schools/", views.admin_school_list_view, name="admin_school_list"),
    path("manager/schools/save/", views.admin_school_save_view, name="admin_school_save"),
    path(
        "manager/schools/<int:school_id>/toggle-active/",
        views.admin_school_toggle_active_view,
        name="admin_school_toggle_active",
    ),

    # =========================================================
    # الإدارة - المشرفون
    # =========================================================
    path("manager/supervisors/", views.admin_supervisor_list_view, name="admin_supervisor_list"),
    path("manager/supervisors/save/", views.admin_supervisor_save_view, name="admin_supervisor_save"),
    path(
        "manager/supervisors/<int:supervisor_id>/toggle-active/",
        views.admin_supervisor_toggle_active_view,
        name="admin_supervisor_toggle_active",
    ),

    # =========================================================
    # الإدارة - الإسناد
    # =========================================================
    # مراجعة الإسناد
    path(
        "manager/assignment-review/",
        views_assignment_review.admin_assignment_review_view,
        name="admin_assignment_review",
    ),
    path(
        "manager/assignment-review/export.xlsx",
        views_assignment_review.admin_assignment_review_export_view,
        name="admin_assignment_review_export",
    ),

    # سجل عمليات معالجة الإسناد
    path(
        "manager/assignment-review/logs/",
        views_assignment_review.admin_assignment_review_logs_view,
        name="admin_assignment_review_logs",
    ),
    path(
        "manager/assignment-review/logs/export.xlsx",
        views_assignment_review.admin_assignment_review_logs_export_view,
        name="admin_assignment_review_logs_export",
    ),

    # معالجة الإسناد المكرر
    path(
        "manager/assignment-review/duplicates/<int:school_id>/",
        views_assignment_review.admin_assignment_duplicate_resolve_view,
        name="admin_assignment_duplicate_resolve",
    ),
    path(
        "manager/assignment-review/duplicates/<int:school_id>/keep/",
        views_assignment_review.admin_assignment_duplicate_keep_view,
        name="admin_assignment_duplicate_keep",
    ),

    # معالجة مشرف غير نشط لديه مدارس
    path(
        "manager/assignment-review/inactive-supervisors/<int:supervisor_id>/",
        views_assignment_review.admin_assignment_inactive_supervisor_resolve_view,
        name="admin_assignment_inactive_supervisor_resolve",
    ),
    path(
        "manager/assignment-review/inactive-supervisors/<int:supervisor_id>/export.xlsx",
        views_assignment_review.admin_assignment_inactive_supervisor_export_view,
        name="admin_assignment_inactive_supervisor_export",
    ),
    path(
        "manager/assignment-review/inactive-supervisors/<int:supervisor_id>/disable/",
        views_assignment_review.admin_assignment_inactive_supervisor_disable_view,
        name="admin_assignment_inactive_supervisor_disable",
    ),

    # معالجة مدرسة معطلة لها إسناد نشط
    path(
        "manager/assignment-review/inactive-schools/<int:school_id>/",
        views_assignment_review.admin_assignment_inactive_school_resolve_view,
        name="admin_assignment_inactive_school_resolve",
    ),
    path(
        "manager/assignment-review/inactive-schools/<int:school_id>/export.xlsx",
        views_assignment_review.admin_assignment_inactive_school_export_view,
        name="admin_assignment_inactive_school_export",
    ),
    path(
        "manager/assignment-review/inactive-schools/<int:school_id>/disable/",
        views_assignment_review.admin_assignment_inactive_school_disable_view,
        name="admin_assignment_inactive_school_disable",
    ),

    path("manager/assignments/", views.admin_assignments_overview_view, name="admin_assignments_overview"),
    path(
        "manager/assignments/unassigned/export/",
        views.admin_export_unassigned_schools_excel,
        name="admin_export_unassigned_schools_excel",
    ),
    path(
        "manager/supervisors/<int:supervisor_id>/assignments/",
        views.admin_supervisor_assignments_view,
        name="admin_supervisor_assignments",
    ),
    path(
        "manager/supervisors/<int:supervisor_id>/assignments/add/",
        views.admin_add_assignment_view,
        name="admin_add_assignment",
    ),
    path(
        "manager/assignments/<int:assignment_id>/delete/",
        views.admin_delete_assignment_view,
        name="admin_delete_assignment",
    ),
    path(
        "manager/supervisors/<int:supervisor_id>/assignments/export/",
        views.admin_export_supervisor_assignments_excel,
        name="admin_export_supervisor_assignments_excel",
    ),

    # =========================================================
    # الإدارة - القطاعات
    # =========================================================
    path("manager/sectors/", views.admin_sector_list_view, name="admin_sector_list"),
    path("manager/sectors/save/", views.admin_sector_save_view, name="admin_sector_save"),
    path(
        "manager/sectors/<int:sector_id>/toggle-active/",
        views.admin_sector_toggle_active_view,
        name="admin_sector_toggle_active",
    ),

    # =========================================================
    # الإدارة - روابط الخطابات
    # =========================================================
    path("manager/weekly-letter-links/", views.weekly_letter_links_list_view, name="weekly_letter_links_list"),
    path("manager/weekly-letter-links/add/", views.weekly_letter_link_create_view, name="weekly_letter_link_create"),
    path("manager/weekly-letter-links/<int:pk>/edit/", views.weekly_letter_link_edit_view, name="weekly_letter_link_edit"),
    path("manager/weekly-letter-links/<int:pk>/delete/", views.weekly_letter_link_delete_view, name="weekly_letter_link_delete"),
    path("manager/weekly-letters-drive/<int:week_number>/", views.weekly_letters_drive_view, name="weekly_letters_drive"),

    # =========================================================
    # الإدارة - الاستيراد
    # =========================================================
    path("manager/import/", views_import.manager_import_view, name="admin_import"),
    path("manager/import/rejected.xlsx", views_import.download_rejected_view, name="download_rejected"),

    # استيراد المدارس وإسنادها للمشرفين
    path(
        "manager/import/schools-supervisors/",
        views_import.admin_schools_with_supervisors_import_view,
        name="admin_schools_with_supervisors_import",
    ),
    path(
        "manager/import/schools-supervisors/template/",
        views_import.admin_schools_with_supervisors_import_template_view,
        name="admin_schools_with_supervisors_import_template",
    ),

    # تصدير المدارس بالمشرفين
    path(
        "manager/import/schools-supervisors/export.xlsx",
        views_import.admin_schools_with_supervisors_export_view,
        name="admin_schools_with_supervisors_export",
    ),
]