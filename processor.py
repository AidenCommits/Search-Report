
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo


# --------------------------------------------------
# CONFIGURATION
# --------------------------------------------------

TIMEZONE = ZoneInfo("America/New_York")


# --------------------------------------------------
# CALCULATE REPORT ANALYTICS
# --------------------------------------------------

def calculate_analytics(records, now=None):
    """
    Calculate search labor metrics and projections.

    Reporting week runs Monday through Sunday.

    Returns a dictionary containing all metrics
    required by report.py.
    """

    if now is None:
        now = datetime.now(TIMEZONE)

    if now.tzinfo is None:
        now = now.replace(tzinfo=TIMEZONE)
    else:
        now = now.astimezone(TIMEZONE)

    # --------------------------------------------------
    # REPORTING PERIOD
    # --------------------------------------------------

    week_start = (
        now - timedelta(days=now.weekday())
    ).replace(hour=0, minute=0, second=0, microsecond=0)

    week_end = week_start + timedelta(days=7)

    # --------------------------------------------------
    # FILTER CURRENT WEEK
    # --------------------------------------------------

    weekly_records = []

    for record in records:
        timestamp = record["timestamp"]

        if timestamp.tzinfo is None:
            timestamp = timestamp.replace(tzinfo=TIMEZONE)
        else:
            timestamp = timestamp.astimezone(TIMEZONE)

        if week_start <= timestamp < week_end:
            weekly_records.append(record)

    # --------------------------------------------------
    # WEEKLY METRICS
    # --------------------------------------------------

    weekly_total = round(
        sum(r["total_cost"] for r in weekly_records), 2
    )

    weekly_searches = len(weekly_records)

    weekly_labor_hours = round(
        sum(r["labor_hours"] for r in weekly_records), 2
    )

    average_search_cost = round(
        weekly_total / weekly_searches, 2
    ) if weekly_searches else 0.0

    average_search_duration = round(
        sum(r["duration_minutes"] for r in weekly_records)
        / weekly_searches, 2
    ) if weekly_searches else 0.0

    # --------------------------------------------------
    # DAILY COSTS - CURRENT WEEK
    # --------------------------------------------------

    daily_costs = {}

    for day_offset in range(7):
        day = (week_start + timedelta(days=day_offset)).date()
        daily_costs[day.isoformat()] = 0.0

    for record in weekly_records:
        timestamp = record["timestamp"]

        if timestamp.tzinfo is None:
            timestamp = timestamp.replace(tzinfo=TIMEZONE)
        else:
            timestamp = timestamp.astimezone(TIMEZONE)

        day_key = timestamp.date().isoformat()

        daily_costs[day_key] += record["total_cost"]

    daily_costs = {
        day: round(cost, 2)
        for day, cost in daily_costs.items()
    }

    # --------------------------------------------------
    # HISTORICAL METRICS
    # --------------------------------------------------

    total_searches = len(records)

    historical_total = round(
        sum(r["total_cost"] for r in records), 2
    )

    historical_labor_hours = round(
        sum(r["labor_hours"] for r in records), 2
    )

    # --------------------------------------------------
    # ANNUALIZED PROJECTION
    # --------------------------------------------------

    # Use completed weeks to avoid projecting from
    # a partially completed reporting week.

    completed_records = []

    for record in records:
        timestamp = record["timestamp"]

        if timestamp.tzinfo is None:
            timestamp = timestamp.replace(tzinfo=TIMEZONE)
        else:
            timestamp = timestamp.astimezone(TIMEZONE)

        if timestamp < week_start:
            completed_records.append(record)

    if completed_records:
        first_timestamp = min(
            r["timestamp"].replace(tzinfo=TIMEZONE)
            if r["timestamp"].tzinfo is None
            else r["timestamp"].astimezone(TIMEZONE)
            for r in completed_records
        )

        first_week_start = (
            first_timestamp - timedelta(days=first_timestamp.weekday())
        ).replace(hour=0, minute=0, second=0, microsecond=0)

        completed_weeks = max(
            1,
            (week_start.date() - first_week_start.date()).days // 7
        )

        completed_total = sum(
            r["total_cost"] for r in completed_records
        )

        average_weekly_cost = completed_total / completed_weeks
        annualized_projection = round(average_weekly_cost * 52, 2)

    else:
        completed_weeks = 0
        average_weekly_cost = 0.0
        annualized_projection = 0.0

    # --------------------------------------------------
    # MONTHLY ANNUAL PROJECTION
    # --------------------------------------------------

    monthly_projection = {}

    for month in range(1, 13):
        monthly_projection[month] = round(
            annualized_projection * (month / 12), 2
        )

    # --------------------------------------------------
    # RETURN REPORT DATA
    # --------------------------------------------------

    return {
        "week_start": week_start,
        "week_end": week_end - timedelta(microseconds=1),

        "weekly_total": weekly_total,
        "weekly_searches": weekly_searches,
        "weekly_labor_hours": weekly_labor_hours,
        "average_search_cost": average_search_cost,
        "average_search_duration": average_search_duration,

        "daily_costs": daily_costs,

        "total_searches": total_searches,
        "historical_total": historical_total,
        "historical_labor_hours": historical_labor_hours,

        "completed_weeks": completed_weeks,
        "average_weekly_cost": round(average_weekly_cost, 2),
        "annualized_projection": annualized_projection,
        "monthly_projection": monthly_projection
    }
