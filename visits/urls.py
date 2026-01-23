# visits/urls.py
from django.urls import path

from . import views
from . import views_import  # ✅ صفحة الاستيراد + تحميل المتجاهل

app_name = "visits"

urlpatterns = [
    # ==================================================
    # 1) Supervisor Auth / Entry
    # ==================================================
    path("", views.login_view, name="login"),
    path("logout/", views.logout_view, name="logout"),

    # ==================================================
    # 2) Supervisor Plan
    # ==================================================
    path("plan/", views.plan_view, name="plan"),
    path("plan/export/", views.export_plan_excel, name="plan_export"),
    path("plan/unlock/", views.request_unlock_view, name="request_unlock"),  # POST

    # ==================================================
    # 3) Management Dashboard (Admin/Manager)
    # ==================================================
    path("manager/dashboard/", views.admin_dashboard_view, name="admin_dashboard"),

    # ✅ صفحة تفاصيل خطة مشرف
    path("manager/plan/<int:plan_id>/", views.admin_plan_detail_view, name="admin_plan_detail"),

    # ✅ Export week Excel (لكل الأسبوع)
    path("manager/export-week/", views.admin_export_week_excel, name="admin_export_week"),

    # ✅ Export plan Excel (خطة واحدة)
    path("manager/plan-export/<int:plan_id>/", views.admin_plan_export_excel, name="admin_plan_export_excel"),

    # ✅ Admin approve / back to draft (AJAX + POST)
    path("manager/plan-approve/<int:plan_id>/", views.admin_plan_approve_view, name="admin_plan_approve"),
    path("manager/plan-draft/<int:plan_id>/", views.admin_plan_back_to_draft_view, name="admin_plan_back_to_draft"),

    # ==================================================
    # 4) Import (✅ الجديد + تحميل المتجاهل)
    # ==================================================
    path("manager/import/", views_import.manager_import_view, name="admin_import"),

    # ✅ تحميل ملف المتجاهل بعد الاستيراد
    path("manager/import/rejected.xlsx", views_import.download_rejected_view, name="download_rejected"),
]
