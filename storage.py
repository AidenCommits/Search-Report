import csv
import os
import uuid

from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

from datetime import datetime
from pathlib import Path


# --------------------------------------------------
# FILE PATHS
# --------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent

DATA_DIR = BASE_DIR / "data"
DATA_FILE = DATA_DIR / "searches.csv"

FIELDNAMES = [
    "id",
    "timestamp",
    "employees",
    "hourly_rate",
    "duration_minutes",
    "labor_hours",
    "total_cost"
]

TIMEZONE = ZoneInfo("America/New_York")

WEEKDAYS = [
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday"
]

# --------------------------------------------------
# INITIALIZE STORAGE
# --------------------------------------------------

def initialize_storage():
    """Create the data directory and CSV file if needed."""

    DATA_DIR.mkdir(parents=True, exist_ok=True)

    if not DATA_FILE.exists():
        with open(DATA_FILE, "w", newline="", encoding="utf-8") as file:
            writer = csv.DictWriter(
                file,
                fieldnames=FIELDNAMES
            )
            writer.writeheader()


# --------------------------------------------------
# SAVE SEARCH INCIDENT
# --------------------------------------------------


def save_search(
    employees,
    hourly_rate,
    duration_minutes,
    day
):
    """
    Save a search incident under the selected weekday
    of the current reporting week.
    """

    initialize_storage()

    now = datetime.now(TIMEZONE)

    # Find Monday of the current week
    monday = (
        now - timedelta(days=now.weekday())
    ).replace(
        hour=0,
        minute=0,
        second=0,
        microsecond=0
    )

    # Determine the selected date
    day_index = WEEKDAYS.index(day)

    search_date = monday + timedelta(days=day_index)

    # Use noon as the event time since only the day
    # is being selected, not a specific clock time.
    search_date = search_date.replace(hour=12)

    labor_hours = employees * (duration_minutes / 60)
    total_cost = labor_hours * hourly_rate

    record = {
        "id": str(uuid.uuid4()),
        "timestamp": search_date.isoformat(),
        "employees": employees,
        "hourly_rate": round(hourly_rate, 2),
        "duration_minutes": duration_minutes,
        "labor_hours": round(labor_hours, 4),
        "total_cost": round(total_cost, 2)
    }

    with open(DATA_FILE, "a", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(
            file,
            fieldnames=FIELDNAMES
        )
        writer.writerow(record)

    return record


# --------------------------------------------------
# LOAD SEARCH HISTORY
# --------------------------------------------------

def load_searches():
    """
    Load all saved search incidents.

    Returns a list of dictionaries with numeric values
    converted to the appropriate types.
    """

    initialize_storage()

    records = []

    with open(DATA_FILE, "r", newline="", encoding="utf-8") as file:
        reader = csv.DictReader(file)

        for row in reader:
            records.append({
                "id": row["id"],
                "timestamp": datetime.fromisoformat(
                    row["timestamp"]
                ),
                "employees": int(row["employees"]),
                "hourly_rate": float(row["hourly_rate"]),
                "duration_minutes": float(row["duration_minutes"]),
                "labor_hours": float(row["labor_hours"]),
                "total_cost": float(row["total_cost"])
            })

    return records
