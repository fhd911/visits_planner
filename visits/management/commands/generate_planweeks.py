from __future__ import annotations

from datetime import datetime, timedelta
from django.core.management.base import BaseCommand
from visits.models import PlanWeek, PlanConfig


class Command(BaseCommand):
    help = "Generate PlanWeek rows automatically from a start Sunday (week1) and weeks_count"

    def add_arguments(self, parser):
        parser.add_argument("--start", type=str, help="Start date for week 1 Sunday (YYYY-MM-DD)")
        parser.add_argument("--count", type=int, help="How many weeks to generate")
        parser.add_argument(
            "--breaks",
            type=str,
            default="",
            help="Comma separated break week numbers e.g. 6,7,8",
        )
        parser.add_argument(
            "--overwrite",
            action="store_true",
            help="Delete existing PlanWeek rows before generating",
        )

    def handle(self, *args, **options):
        start_str = options.get("start")
        count = options.get("count")
        breaks_str = (options.get("breaks") or "").strip()
        overwrite = bool(options.get("overwrite"))

        # ✅ لو ما أعطيت start/count → خذها من PlanConfig
        cfg = PlanConfig.objects.order_by("-id").first()
        if not start_str:
            if not cfg:
                self.stdout.write(self.style.ERROR("No PlanConfig found and --start not provided."))
                return
            start_date = cfg.start_week1_sunday
        else:
            start_date = datetime.strptime(start_str, "%Y-%m-%d").date()

        if not count:
            count = cfg.weeks_count if cfg else 19

        # ✅ تأكيد أنه الأحد
        # Python weekday: Mon=0 .. Sun=6
        if start_date.weekday() != 6:
            self.stdout.write(self.style.WARNING(
                f"⚠️ start date {start_date} is not Sunday. (weekday={start_date.weekday()})"
            ))

        break_weeks = set()
        if breaks_str:
            for x in breaks_str.split(","):
                x = x.strip()
                if x.isdigit():
                    break_weeks.add(int(x))

        if overwrite:
            PlanWeek.objects.all().delete()
            self.stdout.write(self.style.WARNING("Deleted all existing PlanWeek rows."))

        created = 0
        updated = 0

        for i in range(1, count + 1):
            start = start_date + timedelta(weeks=i - 1)
            is_break = (i in break_weeks)

            obj, was_created = PlanWeek.objects.update_or_create(
                week_no=i,
                defaults={
                    "start_sunday": start,
                    "is_break": is_break,
                    "title": "إجازة" if is_break else "",
                },
            )
            created += 1 if was_created else 0
            updated += 0 if was_created else 1

        self.stdout.write(self.style.SUCCESS(
            f"✅ Done. weeks={count} | created={created} | updated={updated} | breaks={sorted(break_weeks)}"
        ))
