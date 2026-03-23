from django.urls import path

from . import views
from . import views_import

app_name = "visits"

urlpatterns = [
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
    path("print-assignment-letter/", views.print_assignment_letter_view, name="print_assignment_letter"),
    path("weekly-letter/", views.current_week_letter_redirect_view, name="current_week_letter"),

    path("plan/", views.plan_view, name="plan"),
    path("plan/export/", views.export_plan_excel, name="plan_export"),
    path(
        "plan/assignments/export/",
        views.export_supervisor_assignments_excel,
        name="export_supervisor_assignments_excel",
    ),
    path("plan/unlock/", views.request_unlock_view, name="request_unlock"),

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
    path("manager/plan/<int:plan_id>/", views.admin_plan_detail_view, name="admin_plan_detail"),
    path("manager/export-week/", views.admin_export_week_excel, name="admin_export_week"),
    path("manager/plan-export/<int:plan_id>/", views.admin_plan_export_excel, name="admin_plan_export_excel"),
    path("manager/plan-approve/<int:plan_id>/", views.admin_plan_approve_view, name="admin_plan_approve"),
    path("manager/plan-draft/<int:plan_id>/", views.admin_plan_back_to_draft_view, name="admin_plan_back_to_draft"),

    # تنبيهات الإدارة
    path("manager/plan-notify/<int:plan_id>/", views.admin_send_notification_view, name="admin_send_notification"),
    path("manager/notify-all/", views.admin_send_notification_all_view, name="admin_send_notification_all"),

    # طلبات فك الاعتماد
    path("manager/unlock-approve/<int:plan_id>/", views.admin_unlock_approve_view, name="admin_unlock_approve"),
    path("manager/unlock-reject/<int:plan_id>/", views.admin_unlock_reject_view, name="admin_unlock_reject"),

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
    path("manager/assignments/", views.admin_assignments_overview_view, name="admin_assignments_overview"),

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
]