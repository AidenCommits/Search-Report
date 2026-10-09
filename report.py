
import io
from datetime import datetime
from pathlib import Path

import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break

from processor import calculate_analytics


# --------------------------------------------------
# CONFIGURATION
# --------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
REPORT_DIR = BASE_DIR / "reports"

NAVY = "172B4D"
BLUE = "2563EB"
LIGHT_BLUE = "DBEAFE"
LIGHT_GRAY = "F1F5F9"
WHITE = "FFFFFF"
DARK = "1E293B"
GRAY = "64748B"

CURRENCY_FORMAT = '"$"#,##0.00'
NUMBER_FORMAT = '#,##0.00'

WEEKDAYS = [
    "Monday", "Tuesday", "Wednesday",
    "Thursday", "Friday", "Saturday", "Sunday"
]

MONTHS = [
    "January", "February", "March",
    "April", "May", "June",
    "July", "August", "September",
    "October", "November", "December"
]


# --------------------------------------------------
# PRINT FORMATTING
# --------------------------------------------------

def configure_print_settings(ws, orientation="landscape"):
    """Configure consistent Excel print settings."""

    ws.page_setup.orientation = orientation
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER

    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.autoPageBreaks = False

    ws.print_options.horizontalCentered = True

    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.35
    ws.page_margins.bottom = 0.35
    ws.page_margins.header = 0.15
    ws.page_margins.footer = 0.15

    ws.oddFooter.center.text = "Page &P of &N"
    ws.oddFooter.center.size = 9

    ws.oddFooter.right.text = "Search Labor Cost Report"
    ws.oddFooter.right.size = 9


# --------------------------------------------------
# EXCEL STYLING HELPERS
# --------------------------------------------------

def section_header(ws, row, title, start=1, end=8):
    """Create a full-width section heading."""

    ws.merge_cells(
        start_row=row,
        start_column=start,
        end_row=row,
        end_column=end
    )

    cell = ws.cell(row=row, column=start)
    cell.value = title

    cell.font = Font(
        name="Aptos",
        size=12,
        bold=True,
        color=WHITE
    )

    cell.fill = PatternFill(
        "solid",
        fgColor=NAVY
    )

    cell.alignment = Alignment(
        vertical="center",
        indent=1
    )

    ws.row_dimensions[row].height = 27


def metric(ws, row, label, value, number_format=None):
    """Write one dashboard metric."""

    label_cell = ws.cell(row=row, column=1)
    value_cell = ws.cell(row=row, column=3)

    label_cell.value = label
    value_cell.value = value

    label_cell.font = Font(
        name="Aptos",
        size=11,
        color=DARK
    )

    value_cell.font = Font(
        name="Aptos",
        size=12,
        bold=True,
        color=BLUE
    )

    if number_format:
        value_cell.number_format = number_format

    background = LIGHT_GRAY if row % 2 == 0 else WHITE

    for col in range(1, 5):
        ws.cell(row, col).fill = PatternFill(
            "solid",
            fgColor=background
        )

    ws.row_dimensions[row].height = 24


def table_header(ws, row, headers):
    """Create formatted column headings."""

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col)

        cell.value = header
        cell.fill = PatternFill(
            "solid",
            fgColor=NAVY
        )

        cell.font = Font(
            name="Aptos",
            bold=True,
            color=WHITE
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    ws.row_dimensions[row].height = 25


def style_table_row(ws, row, start_col, end_col):
    """Apply alternating table row backgrounds."""

    background = LIGHT_GRAY if row % 2 == 0 else WHITE

    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)

        cell.fill = PatternFill(
            "solid",
            fgColor=background
        )

        cell.font = Font(
            name="Aptos",
            size=10,
            color=DARK
        )

    ws.row_dimensions[row].height = 21


# --------------------------------------------------
# MATPLOTLIB HELPERS
# --------------------------------------------------

def currency_axis(value, position):
    """Format chart axis values as dollars."""
    return f"${value:,.0f}"


def style_chart(ax, title, ylabel):
    """Apply consistent chart styling."""

    ax.set_facecolor("white")

    ax.set_title(
        title,
        fontsize=14,
        fontweight="bold",
        color="#172B4D",
        loc="left",
        pad=17
    )

    ax.set_ylabel(
        ylabel,
        fontsize=10,
        color="#64748B",
        labelpad=10
    )

    ax.yaxis.set_major_formatter(
        FuncFormatter(currency_axis)
    )

    ax.grid(
        axis="y",
        color="#E2E8F0",
        linewidth=0.8,
        zorder=0
    )

    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)

    ax.spines["bottom"].set_color("#CBD5E1")

    ax.tick_params(
        axis="both",
        length=0,
        labelsize=9,
        pad=7,
        colors="#64748B"
    )


def embed_chart(ws, fig, anchor, width=550):
    """
    Convert a Matplotlib figure into a PNG image
    and embed it into the worksheet.

    Chart width is controlled to fit the printable
    dashboard area.
    """

    buffer = io.BytesIO()

    fig.savefig(
        buffer,
        format="png",
        dpi=150,
        bbox_inches="tight",
        facecolor="white"
    )

    plt.close(fig)
    buffer.seek(0)

    image = ExcelImage(buffer)

    scale = width / image.width
    image.width = width
    image.height = int(image.height * scale)

    if not hasattr(ws, "_chart_buffers"):
        ws._chart_buffers = []

    ws._chart_buffers.append(buffer)

    ws.add_image(image, anchor)


# --------------------------------------------------
# DAILY SEARCH COST CHART
# --------------------------------------------------

def create_daily_cost_chart(ws, results):
    """Create the weekly search cost bar chart."""

    days = [
        "Mon", "Tue", "Wed", "Thu",
        "Fri", "Sat", "Sun"
    ]

    costs = list(results["daily_costs"].values())

    fig, ax = plt.subplots(figsize=(8.5, 4.1))

    bars = ax.bar(
        days,
        costs,
        width=0.55,
        color="#2563EB",
        zorder=3
    )

    style_chart(
        ax,
        "DAILY SEARCH LABOR COST",
        "Labor Cost ($)"
    )

    maximum = max(costs, default=0)

    ax.set_ylim(
        0,
        max(maximum * 1.35, 10)
    )

    for bar, cost in zip(bars, costs):
        if cost <= 0:
            continue

        ax.annotate(
            f"${cost:,.2f}",
            xy=(
                bar.get_x() + bar.get_width() / 2,
                bar.get_height()
            ),
            xytext=(0, 6),
            textcoords="offset points",
            ha="center",
            fontsize=8,
            fontweight="bold",
            color="#172B4D"
        )

    fig.tight_layout(pad=1.5)

    embed_chart(ws, fig, "D25", width=550)


# --------------------------------------------------
# ANNUAL PROJECTION CHART
# --------------------------------------------------

def create_projection_chart(ws, results):
    """Create the cumulative annual cost projection."""

    months = [
        "Jan", "Feb", "Mar", "Apr",
        "May", "Jun", "Jul", "Aug",
        "Sep", "Oct", "Nov", "Dec"
    ]

    costs = [
        results["monthly_projection"][month]
        for month in range(1, 13)
    ]

    fig, ax = plt.subplots(figsize=(8.5, 4.1))

    x_values = list(range(12))

    ax.plot(
        x_values,
        costs,
        color="#2563EB",
        linewidth=2.8,
        marker="o",
        markersize=4,
        zorder=3
    )

    ax.fill_between(
        x_values,
        costs,
        color="#DBEAFE",
        alpha=0.65,
        zorder=2
    )

    ax.set_xticks(x_values)
    ax.set_xticklabels(months)

    style_chart(
        ax,
        "PROJECTED ANNUAL SEARCH COST",
        "Cumulative Cost ($)"
    )

    maximum = max(costs, default=0)

    ax.set_ylim(
        0,
        max(maximum * 1.20, 10)
    )

    if maximum > 0:
        ax.annotate(
            f"${costs[-1]:,.2f}",
            xy=(11, costs[-1]),
            xytext=(-10, 12),
            textcoords="offset points",
            ha="right",
            fontsize=10,
            fontweight="bold",
            color="#172B4D"
        )

    fig.tight_layout(pad=1.5)

    embed_chart(ws, fig, "D46", width=550)


# --------------------------------------------------
# MANAGEMENT DASHBOARD
# --------------------------------------------------

def create_dashboard(wb, results):
    """Build the two-page management dashboard."""

    ws = wb.active
    ws.title = "Management Dashboard"

    ws.sheet_view.showGridLines = False

    # Column widths
    for col in range(1, 9):
        ws.column_dimensions[
            get_column_letter(col)
        ].width = 14

    ws.column_dimensions["A"].width = 33
    ws.column_dimensions["B"].width = 17
    ws.column_dimensions["C"].width = 18

    # --------------------------------------------------
    # TITLE
    # --------------------------------------------------

    ws.merge_cells("A1:H2")

    title = ws["A1"]
    title.value = "SEARCH LABOR COST REPORT"

    title.font = Font(
        name="Aptos Display",
        size=20,
        bold=True,
        color=WHITE
    )

    title.fill = PatternFill(
        "solid",
        fgColor=NAVY
    )

    title.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws.row_dimensions[1].height = 27
    ws.row_dimensions[2].height = 27

    ws.merge_cells("A3:H3")

    ws["A3"] = (
        f"Reporting Period: "
        f"{results['week_start']:%B %d, %Y} - "
        f"{results['week_end']:%B %d, %Y}"
    )

    ws["A3"].font = Font(
        name="Aptos",
        size=11,
        italic=True,
        color=GRAY
    )

    ws["A3"].alignment = Alignment(
        horizontal="center"
    )

    # --------------------------------------------------
    # WEEKLY METRICS
    # --------------------------------------------------

    section_header(
        ws, 5,
        "WEEKLY SEARCH METRICS"
    )

    metric(
        ws, 7,
        "Total Search Cost",
        results["weekly_total"],
        CURRENCY_FORMAT
    )

    metric(
        ws, 8,
        "Search Incidents",
        results["weekly_searches"]
    )

    metric(
        ws, 9,
        "Labor Hours Lost",
        results["weekly_labor_hours"],
        NUMBER_FORMAT
    )

    metric(
        ws, 10,
        "Average Cost per Search",
        results["average_search_cost"],
        CURRENCY_FORMAT
    )

    metric(
        ws, 11,
        "Average Search Duration (min)",
        results["average_search_duration"],
        NUMBER_FORMAT
    )

    # --------------------------------------------------
    # HISTORICAL METRICS
    # --------------------------------------------------

    section_header(
        ws, 13,
        "HISTORICAL & PROJECTED COSTS"
    )

    metric(
        ws, 15,
        "Historical Search Cost",
        results["historical_total"],
        CURRENCY_FORMAT
    )

    metric(
        ws, 16,
        "Total Search Incidents",
        results["total_searches"]
    )

    metric(
        ws, 17,
        "Total Labor Hours Lost",
        results["historical_labor_hours"],
        NUMBER_FORMAT
    )

    metric(
        ws, 18,
        "Average Weekly Cost",
        results["average_weekly_cost"],
        CURRENCY_FORMAT
    )

    metric(
        ws, 19,
        "Annualized Projection",
        results["annualized_projection"],
        CURRENCY_FORMAT
    )

    metric(
        ws, 20,
        "Completed Reporting Weeks",
        results["completed_weeks"]
    )

    # --------------------------------------------------
    # DAILY COST BREAKDOWN
    # --------------------------------------------------

    section_header(
        ws, 23,
        "DAILY SEARCH COST BREAKDOWN"
    )

    table_header(
        ws, 25,
        ["Day", "Search Cost"]
    )

    daily_values = list(
        results["daily_costs"].values()
    )

    for index, day in enumerate(WEEKDAYS):
        row = 26 + index

        ws.cell(row, 1).value = day
        ws.cell(row, 2).value = daily_values[index]

        ws.cell(row, 2).number_format = CURRENCY_FORMAT

        style_table_row(ws, row, 1, 2)

    create_daily_cost_chart(ws, results)

    # --------------------------------------------------
    # PAGE 2: ANNUAL PROJECTION
    # --------------------------------------------------

    section_header(
        ws, 44,
        "ANNUAL COST PROJECTION"
    )

    table_header(
        ws, 46,
        ["Month", "Projected Cost"]
    )

    for month, name in enumerate(MONTHS, start=1):
        row = 46 + month

        ws.cell(row, 1).value = name
        ws.cell(row, 2).value = (
            results["monthly_projection"][month]
        )

        ws.cell(row, 2).number_format = CURRENCY_FORMAT

        style_table_row(ws, row, 1, 2)

    create_projection_chart(ws, results)

    # Projection notes
    ws.merge_cells("A61:H61")

    ws["A61"] = (
        "Annual projection is based on average search "
        "costs from completed reporting weeks."
    )

    ws["A61"].font = Font(
        name="Aptos",
        size=10,
        italic=True,
        color=GRAY
    )

    if results["completed_weeks"] == 0:
        ws.merge_cells("A62:H62")

        ws["A62"] = (
            "No completed reporting weeks available. "
            "Annual projections will populate after "
            "the first completed week."
        )

        ws["A62"].font = Font(
            name="Aptos",
            size=10,
            italic=True,
            color=GRAY
        )

    # --------------------------------------------------
    # DASHBOARD PRINT SETTINGS
    # --------------------------------------------------

    configure_print_settings(
        ws,
        orientation="landscape"
    )

    ws.print_area = "A1:H62"

    # Fit horizontally while preserving manual
    # vertical page breaks.
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    # Page 1 ends at row 43.
    # Page 2 starts at row 44.
    ws.row_breaks.append(Break(id=22))
    ws.row_breaks.append(Break(id=43))

    ws.freeze_panes = "A5"

    return ws


# --------------------------------------------------
# SEARCH HISTORY
# --------------------------------------------------

def create_history(wb, records):
    """Create the printable search incident history."""

    ws = wb.create_sheet("Search History")
    ws.sheet_view.showGridLines = False

    headers = [
        "Date",
        "Day",
        "Employees",
        "Hourly Rate",
        "Duration (min)",
        "Labor Hours",
        "Total Cost",
        "Record ID"
    ]

    # --------------------------------------------------
    # TITLE
    # --------------------------------------------------

    ws.merge_cells("A1:H2")

    title = ws["A1"]
    title.value = "SEARCH INCIDENT HISTORY"

    title.font = Font(
        name="Aptos Display",
        size=19,
        bold=True,
        color=WHITE
    )

    title.fill = PatternFill(
        "solid",
        fgColor=NAVY
    )

    title.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws.row_dimensions[1].height = 27
    ws.row_dimensions[2].height = 27

    # --------------------------------------------------
    # TABLE
    # --------------------------------------------------

    table_header(ws, 4, headers)

    sorted_records = sorted(
        records,
        key=lambda record: record["timestamp"],
        reverse=True
    )

    for row_index, record in enumerate(
        sorted_records,
        start=5
    ):
        timestamp = record["timestamp"]

        values = [
            timestamp.date(),
            timestamp.strftime("%A"),
            record["employees"],
            record["hourly_rate"],
            record["duration_minutes"],
            record["labor_hours"],
            record["total_cost"],
            record["id"]
        ]

        for col_index, value in enumerate(
            values,
            start=1
        ):
            ws.cell(
                row=row_index,
                column=col_index,
                value=value
            )

        style_table_row(
            ws,
            row_index,
            1,
            8
        )

        ws.cell(row_index, 1).number_format = "mm/dd/yyyy"
        ws.cell(row_index, 4).number_format = CURRENCY_FORMAT
        ws.cell(row_index, 5).number_format = NUMBER_FORMAT
        ws.cell(row_index, 6).number_format = NUMBER_FORMAT
        ws.cell(row_index, 7).number_format = CURRENCY_FORMAT

    # Column widths
    widths = [
        16, 16, 15, 17,
        19, 17, 17, 40
    ]

    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[
            get_column_letter(index)
        ].width = width

    ws.freeze_panes = "A5"

    if sorted_records:
        ws.auto_filter.ref = (
            f"A4:H{4 + len(sorted_records)}"
        )

    # --------------------------------------------------
    # HISTORY PRINT SETTINGS
    # --------------------------------------------------

    configure_print_settings(
        ws,
        orientation="landscape"
    )

    ws.print_area = f"A1:H{max(5, ws.max_row)}"

    # Repeat title and column headers on every page.
    ws.print_title_rows = "1:4"

    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    return ws


# --------------------------------------------------
# GENERATE COMPLETE REPORT
# --------------------------------------------------

def generate_report(records):
    """
    Generate the Excel Search Report.

    Returns:
        Path to the generated workbook.
    """

    REPORT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    # Calculate analytics
    results = calculate_analytics(records)

    # Create workbook
    wb = Workbook()

    create_dashboard(wb, results)
    create_history(wb, records)

    # Unique timestamp to prevent overwriting.
    timestamp = datetime.now().strftime(
        "%Y-%m-%d_%H-%M-%S_%f"
    )

    report_path = (
        REPORT_DIR /
        f"Search_Report_{timestamp}.xlsx"
    )

    wb.save(report_path)

    print("\n--- SEARCH REPORT ---")
    print(
        "Reporting Period:",
        results["week_start"],
        "to",
        results["week_end"]
    )
    print("Weekly Cost:", results["weekly_total"])
    print("Weekly Searches:", results["weekly_searches"])
    print("Labor Hours:", results["weekly_labor_hours"])
    print(
        "Annual Projection:",
        results["annualized_projection"]
    )
    print("Report Saved:", report_path)
    print("---------------------\n")

    return report_path
