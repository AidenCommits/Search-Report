
from gui import launch_gui
from storage import save_search, load_searches, initialize_storage
from report import generate_report


# --------------------------------------------------
# REPORT GENERATION
# --------------------------------------------------

def generate_search_report(
    employees,
    hourly_rate,
    minutes,
    day
):
    """
    Save a search incident and generate
    the complete Excel management report.

    Called by the GUI when Generate Report
    is clicked.
    """

    # --------------------------------------------------
    # SAVE NEW SEARCH INCIDENT
    # --------------------------------------------------

    # A saved incident awaiting a successful report.
    # This prevents duplicate records if Excel fails.
    if not hasattr(generate_search_report, "pending_record"):
        generate_search_report.pending_record = None

    if generate_search_report.pending_record is None:

        record = save_search(
            employees,
            hourly_rate,
            minutes,
            day
        )

        generate_search_report.pending_record = record

        print("\n--- NEW SEARCH RECORDED ---")
        print("Employees:", record["employees"])
        print("Hourly Rate:", record["hourly_rate"])
        print("Duration:", record["duration_minutes"])
        print("Labor Hours:", record["labor_hours"])
        print("Total Cost:", record["total_cost"])
        print("---------------------------\n")

    # --------------------------------------------------
    # LOAD ALL SEARCH RECORDS
    # --------------------------------------------------

    records = load_searches()

    # --------------------------------------------------
    # GENERATE EXCEL REPORT
    # --------------------------------------------------

    # If report generation fails, pending_record
    # remains set and the next click retries the
    # report without saving another incident.
    report_path = generate_report(records)

    # Clear pending record after successful generation.
    generate_search_report.pending_record = None

    return report_path


# --------------------------------------------------
# APPLICATION ENTRY POINT
# --------------------------------------------------

def main():
    """Initialize storage and launch the application."""

    initialize_storage()

    print("\nSearch Report Generator")
    print("Storage initialized.")
    print("Launching GUI...\n")

    launch_gui(generate_search_report)


if __name__ == "__main__":
    main()
