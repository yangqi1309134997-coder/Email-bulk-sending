from datetime import datetime, timedelta
from typing import Optional


# Common email domain -> timezone mapping
DOMAIN_TIMEZONES = {
    "qq.com": "Asia/Shanghai",
    "163.com": "Asia/Shanghai",
    "126.com": "Asia/Shanghai",
    "sina.com": "Asia/Shanghai",
    "sohu.com": "Asia/Shanghai",
    "yeah.net": "Asia/Shanghai",
    "aliyun.com": "Asia/Shanghai",
    "foxmail.com": "Asia/Shanghai",
    "gmail.com": "America/New_York",
    "yahoo.com": "America/Los_Angeles",
    "hotmail.com": "America/New_York",
    "outlook.com": "America/New_York",
    "aol.com": "America/New_York",
    "icloud.com": "America/Los_Angeles",
    "mail.ru": "Europe/Moscow",
    "yandex.ru": "Europe/Moscow",
    "web.de": "Europe/Berlin",
    "gmx.de": "Europe/Berlin",
    "gmx.net": "Europe/Berlin",
}

# Best sending hours (local time) per timezone group
PEAK_HOURS = {
    "Asia/Shanghai": [(9, 11), (14, 16), (19, 21)],
    "America/New_York": [(9, 11), (14, 16)],
    "America/Los_Angeles": [(9, 11), (14, 16)],
    "Europe/Berlin": [(9, 11), (14, 16)],
    "Europe/Moscow": [(9, 11), (14, 16)],
}

DEFAULT_PEAK_HOURS = [(9, 11), (14, 16)]


def infer_timezone(email: str) -> str:
    domain = email.split("@")[-1].lower()
    return DOMAIN_TIMEZONES.get(domain, "Asia/Shanghai")


def get_peak_hours(timezone: str) -> list[tuple[int, int]]:
    return PEAK_HOURS.get(timezone, DEFAULT_PEAK_HOURS)


def is_peak_hour(email: str, current_hour: Optional[int] = None) -> bool:
    tz = infer_timezone(email)
    peak = get_peak_hours(tz)
    hour = current_hour if current_hour is not None else datetime.utcnow().hour
    for start, end in peak:
        if start <= hour < end:
            return True
    return False


def compute_smart_schedule(
    recipients: list[dict],
    start_time: Optional[datetime] = None,
    emails_per_slot: int = 50,
) -> list[dict]:
    """Group recipients by timezone peak slots and assign send times.

    Returns list of {"time": datetime, "recipients": list[dict]}.
    """
    from collections import defaultdict

    if not recipients:
        return []

    start = start_time or datetime.utcnow()

    # Group by timezone
    tz_groups: dict[str, list[dict]] = defaultdict(list)
    for r in recipients:
        tz = infer_timezone(r.get("email", ""))
        tz_groups[tz].append(r)

    schedule = []
    for tz, group in tz_groups.items():
        peak = get_peak_hours(tz)
        # Find next available peak slot starting from start_time
        slot_time = start
        remaining = list(group)

        while remaining:
            # Find next peak hour on slot_time's day
            assigned_this_day = False
            for slot_start, slot_end in peak:
                candidate = slot_time.replace(hour=slot_start, minute=0, second=0, microsecond=0)
                if candidate < start:
                    candidate = start
                # Assign up to emails_per_slot in this slot
                batch = remaining[:emails_per_slot]
                remaining = remaining[emails_per_slot:]
                schedule.append({"time": candidate, "recipients": batch})
                assigned_this_day = True
                if not remaining:
                    break
            if remaining:
                # Move to next day
                slot_time += timedelta(days=1)

    schedule.sort(key=lambda x: x["time"])
    return schedule
