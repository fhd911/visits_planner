from __future__ import annotations

from django.shortcuts import redirect
from django.urls import reverse
from django.utils.deprecation import MiddlewareMixin

from .models import SiteSetting


class MaintenanceModeMiddleware(MiddlewareMixin):
    """
    يتحكم في إظهار وضع الصيانة على مستوى الموقع بالكامل.

    السلوك:
    - إذا لم تكن الصيانة مفعلة: يمنع بقاء صفحة /maintenance/ مفتوحة ويعيد التوجيه.
    - إذا كانت الصيانة مفعلة:
        - يسمح للإدارة بالدخول إذا allow_admin_only=True
        - يمنع بقية المستخدمين ويوجههم إلى صفحة الصيانة
    - إذا انتهت نافذة الصيانة الزمنية:
        - يتم إلغاء تفعيل وضع الصيانة تلقائيًا
    """

    def process_request(self, request):
        try:
            setting = SiteSetting.get_solo()
        except Exception:
            return None

        self._auto_disable_if_window_expired(setting)

        path = request.path or ""
        maintenance_url = reverse("visits:maintenance_page")
        login_url = reverse("visits:login")
        admin_login_url = reverse("visits:admin_login")
        admin_dashboard_url = reverse("visits:admin_dashboard")

        allowed_exact_paths = {
            maintenance_url,
            admin_login_url,
            reverse("visits:logout"),
        }

        allowed_prefixes = (
            "/static/",
            "/media/",
            "/admin/",
        )

        if any(path.startswith(prefix) for prefix in allowed_prefixes):
            return None

        if not setting.is_maintenance_mode:
            if path == maintenance_url:
                if self._is_admin_user(request):
                    return redirect(admin_dashboard_url)
                return redirect(login_url)
            return None

        if path in allowed_exact_paths:
            return None

        if self._is_admin_user(request) and setting.allow_admin_only:
            return None

        return redirect(maintenance_url)

    @staticmethod
    def _is_admin_user(request) -> bool:
        user = getattr(request, "user", None)
        return bool(user and user.is_authenticated and user.is_staff)

    @staticmethod
    def _auto_disable_if_window_expired(setting: SiteSetting) -> None:
        """
        إذا كانت الصيانة مفعلة، وكانت نافذة الصيانة محددة وانتهت،
        يتم إلغاء تفعيل الصيانة تلقائيًا.
        """
        if not setting.is_maintenance_mode:
            return

        if setting.maintenance_ends_at and not setting.is_currently_in_maintenance_window:
            update_fields = ["is_maintenance_mode", "updated_at"]

            setting.is_maintenance_mode = False

            if setting.expected_return_text:
                setting.expected_return_text = None
                update_fields.append("expected_return_text")

            setting.save(update_fields=update_fields)