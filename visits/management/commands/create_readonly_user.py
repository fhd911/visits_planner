from __future__ import annotations

from django.contrib.auth import get_user_model
from django.contrib.auth.models import Group
from django.core.management.base import BaseCommand, CommandError

from visits.models import Sector


class Command(BaseCommand):
    help = "إنشاء أو تحديث مستخدم لبوابة الاطلاع: مدير قسم أو مدير وحدة."

    def add_arguments(self, parser):
        parser.add_argument("username")
        parser.add_argument("password")
        parser.add_argument(
            "--role",
            choices=["department", "unit"],
            required=True,
            help="department = مدير القسم، unit = مدير وحدة",
        )
        parser.add_argument(
            "--sector",
            action="append",
            default=[],
            help="رقم القطاع لمدير الوحدة. يمكن تكراره أكثر من مرة.",
        )
        parser.add_argument("--first-name", default="")
        parser.add_argument("--last-name", default="")
        parser.add_argument("--email", default="")

    def handle(self, *args, **options):
        User = get_user_model()
        username = options["username"]
        password = options["password"]
        role = options["role"]

        user, created = User.objects.get_or_create(username=username)
        user.set_password(password)
        user.is_active = True
        user.is_staff = False
        user.is_superuser = False
        user.first_name = options.get("first_name") or user.first_name
        user.last_name = options.get("last_name") or user.last_name
        user.email = options.get("email") or user.email
        user.save()

        # إزالة المجموعات القديمة الخاصة بالاطلاع فقط حتى لا تتضارب الصلاحية.
        readonly_prefixes = ("مدير القسم", "department_manager", "مدير وحدة:", "unit_manager:")
        for group in list(user.groups.all()):
            if group.name in {"مدير القسم", "department_manager", "readonly_department_manager"} or group.name.startswith(readonly_prefixes):
                user.groups.remove(group)

        if role == "department":
            group, _ = Group.objects.get_or_create(name="مدير القسم")
            user.groups.add(group)
            scope_label = "جميع القطاعات"
        else:
            sector_values = options.get("sector") or []
            if not sector_values:
                raise CommandError("مدير الوحدة يحتاج --sector برقم القطاع.")

            sectors = []
            for value in sector_values:
                try:
                    sector_id = int(value)
                except Exception as exc:
                    raise CommandError(f"رقم القطاع غير صحيح: {value}") from exc

                sector = Sector.objects.filter(id=sector_id).first()
                if not sector:
                    raise CommandError(f"لم يتم العثور على قطاع بالرقم: {sector_id}")
                sectors.append(sector)

            for sector in sectors:
                group, _ = Group.objects.get_or_create(name=f"مدير وحدة:{sector.id}")
                user.groups.add(group)

            scope_label = "، ".join(f"{s.name} ({s.id})" for s in sectors)

        self.stdout.write(self.style.SUCCESS("تم حفظ مستخدم بوابة الاطلاع بنجاح."))
        self.stdout.write(f"المستخدم: {username}")
        self.stdout.write(f"الدور: {'مدير القسم' if role == 'department' else 'مدير وحدة'}")
        self.stdout.write(f"النطاق: {scope_label}")
        self.stdout.write("رابط الدخول: /viewer-login/")
