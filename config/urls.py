from __future__ import annotations

from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path("admin/", admin.site.urls),

    # كل روابط التطبيق تحت /
    path("", include("visits.urls")),
]
