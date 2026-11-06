"""Lloeva Beauty business dashboard.

Reads the Google Sheet (STOCKS, SALES, REVENUE, COMMISION tabs) with the service
account in .env and serves a report at http://127.0.0.1:5009.

Run:  python app.py
"""

import json
import re
import threading
import time
from collections import OrderedDict
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

import gspread
from flask import Flask, abort, render_template, request, send_file, url_for
from google.oauth2.service_account import Credentials
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

BASE_DIR = Path(__file__).resolve().parent
ENV_FILE = BASE_DIR / ".env"
PORT = 5009
CACHE_SECONDS = 300
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SHEET_TABS = ["STOCKS", "SALES", "REVENUE", "COMMISION"]
WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
TOP_SERVICES = 10
EXPORT_MAX_ROWS = 5000
EXPORT_MAX_COLUMNS = 20
BRAND_ROSE = "C2255C"
BRAND_BLUSH = "FFF0F5"

# Filter and page tabs, in display order
VIEWS = [
    ("weekly", "Weekly"),
    ("monthly", "Monthly"),
    ("quarterly", "Quarterly"),
    ("custom", "Custom"),
]
DEFAULT_VIEW = "quarterly"
PAGE_TABS = [
    ("stocks", "Stocks"),
    ("sales", "Sales and revenue"),
    ("margin", "Commission and margin"),
]

app = Flask(__name__)
app.config["TEMPLATES_AUTO_RELOAD"] = True  # page edits show on refresh


# --------------------------------------------------------------------------
# Google Sheet access
# --------------------------------------------------------------------------

def read_env():
    """Return (service account credentials, sheet id) from .env.

    The key is stored as multi-line JSON, which standard .env parsers can't
    read, so the file is parsed here.
    """
    text = ENV_FILE.read_text(encoding="utf-8-sig")

    def value(key):
        match = re.search(rf"^{key}\s*=\s*(.*)$", text, re.MULTILINE)
        return match.group(1).strip().strip("'\"") if match else ""

    sheet_id = value("SHEET_ID")
    if not sheet_id:
        raise RuntimeError("SHEET_ID is missing from .env")

    match = re.search(
        r"^GOOGLE_APPLICATION_CREDENTIALS_JSON\s*=\s*['\"]?(\{.*?^\}|\{[^\n]*\})",
        text,
        re.MULTILINE | re.DOTALL,
    )
    if match:
        info = json.loads(match.group(1), strict=False)
        return Credentials.from_service_account_info(info, scopes=SCOPES), sheet_id

    key_path = value("GOOGLE_APPLICATION_CREDENTIALS")
    if key_path:
        path = Path(key_path)
        path = path if path.is_absolute() else BASE_DIR / path
        return Credentials.from_service_account_file(str(path), scopes=SCOPES), sheet_id

    raise RuntimeError("No service account credentials found in .env")


def fetch_sheet():
    """Read all four tabs in one API call. Returns {tab name: rows}."""
    creds, sheet_id = read_env()
    spreadsheet = gspread.authorize(creds).open_by_key(sheet_id)
    response = spreadsheet.values_batch_get(
        SHEET_TABS,
        params={
            "valueRenderOption": "UNFORMATTED_VALUE",
            "dateTimeRenderOption": "SERIAL_NUMBER",
        },
    )
    return {
        tab: block.get("values", [])
        for tab, block in zip(SHEET_TABS, response["valueRanges"])
    }


# --------------------------------------------------------------------------
# Parsing helpers
# --------------------------------------------------------------------------

def num(value):
    """Return a float for numeric cells, otherwise None."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.replace(",", "").strip())
        except ValueError:
            return None
    return None


def to_date(value):
    """Convert a Sheets date serial or date text to a date."""
    if isinstance(value, (int, float)) and not isinstance(value, bool) and value > 30000:
        return date(1899, 12, 30) + timedelta(days=int(value))
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%Y", "%d %b %Y"):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                pass
    return None


def cell(row, index):
    return row[index] if index < len(row) else None


def header_index(rows):
    """Find the header row: the one starting with PARTICULARS."""
    for index, row in enumerate(rows):
        if row and str(row[0]).strip().upper() == "PARTICULARS":
            return index
    raise ValueError("Could not find the PARTICULARS header row")


def parse_grid(rows):
    """Parse SALES or REVENUE into (dates, {item: {date: value}})."""
    head = header_index(rows)
    columns = [
        (index, parsed)
        for index, parsed in ((i, to_date(v)) for i, v in enumerate(rows[head]))
        if index > 0 and parsed
    ]
    items = OrderedDict()
    for row in rows[head + 1:]:
        name = str(cell(row, 0) or "").strip()
        if not name:
            continue
        values = items.setdefault(name, {})
        for index, day in columns:
            values[day] = values.get(day, 0.0) + (num(cell(row, index)) or 0.0)
    return [day for _, day in columns], items


def parse_stocks(rows):
    stock = OrderedDict()
    for row in rows[header_index(rows) + 1:]:
        name = str(cell(row, 0) or "").strip()
        if not name:
            continue
        qty_cell = cell(row, 1)
        price_cell = cell(row, 4)
        qty, amount, cost = num(qty_cell), num(cell(row, 2)), num(cell(row, 3))
        if qty is None and qty_cell not in (None, "") and amount and cost:
            qty = amount / cost  # e.g. "30 ml" perfume sold in portions
        if amount is None and qty and cost:
            amount = qty * cost
        stock[name] = {
            "qty": qty,
            "amount": amount,
            "cost": cost,
            "price": num(price_cell),
            "used": str(price_cell or "").strip().lower() == "used",
        }
    return stock


def parse_commission(rows):
    commission = {}
    for row in rows:
        day = to_date(cell(row, 0))
        if day:
            commission[day] = commission.get(day, 0.0) + (num(cell(row, 1)) or 0.0)
    return commission


def parse_sheet(raw):
    """Turn the raw tabs into the data every report is built from."""
    _, sales = parse_grid(raw["SALES"])
    all_dates, revenue = parse_grid(raw["REVENUE"])
    active = [
        day for day in all_dates
        if sum(values.get(day, 0.0) for values in revenue.values()) > 0
    ]
    if not active:
        raise ValueError("The REVENUE tab has no sales yet")
    return {
        "stock": parse_stocks(raw["STOCKS"]),
        "sales": sales,
        "revenue": revenue,
        "commission": parse_commission(raw["COMMISION"]),
        # Blank days after the last sale, such as today, are left out
        "dates": [day for day in all_dates if active[0] <= day <= active[-1]],
    }


def label(day):
    return f"{day.day} {day.strftime('%b')}"


def share(part, whole):
    return part / whole * 100 if whole else 0


# --------------------------------------------------------------------------
# Date filter
# --------------------------------------------------------------------------

def period_options(dates):
    """List the weeks, months and quarters in the data, newest first."""
    options = {"weekly": OrderedDict(), "monthly": OrderedDict(), "quarterly": OrderedDict()}
    for day in reversed(dates):
        monday = day - timedelta(days=day.weekday())
        sunday = monday + timedelta(days=6)
        options["weekly"].setdefault(monday.isoformat(), {
            "label": f"{label(monday)} to {label(sunday)} {sunday.year}",
            "start": monday,
            "end": sunday,
        })

        month_start = day.replace(day=1)
        next_month = (month_start + timedelta(days=32)).replace(day=1)
        options["monthly"].setdefault(month_start.strftime("%Y-%m"), {
            "label": day.strftime("%B %Y"),
            "start": month_start,
            "end": next_month - timedelta(days=1),
        })

        quarter = (day.month - 1) // 3 + 1
        first_month = (quarter - 1) * 3 + 1
        quarter_start = date(day.year, first_month, 1)
        quarter_end = (
            date(day.year + 1, 1, 1) if quarter == 4
            else date(day.year, first_month + 3, 1)
        ) - timedelta(days=1)
        options["quarterly"].setdefault(f"{day.year}-Q{quarter}", {
            "label": (
                f"Q{quarter} {day.year} "
                f"({quarter_start.strftime('%b')} to {quarter_end.strftime('%b')})"
            ),
            "start": quarter_start,
            "end": quarter_end,
        })
    return options


def parse_iso(text):
    try:
        return date.fromisoformat(text or "")
    except ValueError:
        return None


def resolve_filter(args, dates):
    """Read the filter from the URL. Anything invalid falls back to a default."""
    view = args.get("view", DEFAULT_VIEW)
    if view not in dict(VIEWS):
        view = DEFAULT_VIEW
    options = period_options(dates)

    if view == "custom":
        start = parse_iso(args.get("start")) or dates[0]
        end = parse_iso(args.get("end")) or dates[-1]
        if start > end:
            start, end = end, start
        chosen = {"key": "", "start": start, "end": end}
        choices = []
    else:
        periods = options[view]
        key = args.get("period")
        if key not in periods:
            key = next(iter(periods))
        chosen = {"key": key, **periods[key]}
        choices = [{"key": k, "label": p["label"]} for k, p in periods.items()]

    return {
        "view": view,
        "key": chosen["key"],
        "start": chosen["start"],
        "end": chosen["end"],
        "choices": choices,
        "min": dates[0].isoformat(),
        "max": dates[-1].isoformat(),
    }


# --------------------------------------------------------------------------
# Report
# --------------------------------------------------------------------------

def build_report(data, start, end):
    """Build the report for start to end, inclusive."""
    stock, sales = data["stock"], data["sales"]
    revenue, commission = data["revenue"], data["commission"]
    dates = [day for day in data["dates"] if start <= day <= end]
    to_date_range = [day for day in data["dates"] if day <= end]

    if dates:
        period = f"{label(dates[0])} to {label(dates[-1])} {dates[-1].year}"
    else:
        period = f"{label(start)} to {label(end)} {end.year}"

    def total(grid, name, days=None):
        values = grid.get(name, {})
        return sum(values.get(day, 0.0) for day in (dates if days is None else days))

    service_names = [name for name in revenue if name not in stock]
    product_names = [name for name in revenue if name in stock]

    # Daily figures
    daily = []
    for day in dates:
        service = sum(revenue[name].get(day, 0.0) for name in service_names)
        product = sum(revenue[name].get(day, 0.0) for name in product_names)
        monday = day - timedelta(days=day.weekday())
        daily.append({
            "date": day.isoformat(),
            "label": label(day),
            "weekday": WEEKDAYS[day.weekday()],
            "week": label(monday),
            "service": service,
            "product": product,
            "revenue": service + product,
            "commission": commission.get(day),
        })

    total_revenue = sum(d["revenue"] for d in daily)
    if not total_revenue:
        return {"has_data": False, "period": period}

    service_revenue = sum(d["service"] for d in daily)
    product_revenue = sum(d["product"] for d in daily)
    total_commission = sum(d["commission"] or 0.0 for d in daily)
    day_count = len(daily)

    # Average by weekday
    weekdays = []
    for name in WEEKDAYS:
        values = [d["revenue"] for d in daily if d["weekday"] == name]
        if values:
            weekdays.append({
                "day": name,
                "average": sum(values) / len(values),
                "days": len(values),
            })

    # Service ranking
    ranked = sorted(
        (
            {"name": name, "revenue": total(revenue, name), "jobs": total(sales, name)}
            for name in service_names
        ),
        key=lambda item: -item["revenue"],
    )
    ranked = [item for item in ranked if item["revenue"] or item["jobs"]]
    services = ranked[:TOP_SERVICES]
    others = ranked[TOP_SERVICES:]
    if others:
        services = services + [{
            "name": f"{len(others)} other services",
            "revenue": sum(item["revenue"] for item in others),
            "jobs": sum(item["jobs"] for item in others),
            "other": True,
        }]
    top_five = sum(item["revenue"] for item in ranked[:5])

    # Commission as a share of service revenue
    commission_days = [d for d in daily if d["commission"] is not None]
    commission_service = sum(d["service"] for d in commission_days)
    commission_revenue = sum(d["revenue"] for d in commission_days)
    rates = [
        {
            "date": d["date"],
            "label": d["label"],
            "service": d["service"],
            "commission": d["commission"],
            "rate": d["commission"] / d["service"] * 100,
        }
        for d in commission_days
        if d["service"] > 0
    ]
    paid_rates = [r["rate"] for r in rates if r["commission"] > 0]

    # Stock and margins. Sales in the period drive cost and margin;
    # stock left counts everything sold up to the end of the period.
    invested = used_cost = cogs = costed_revenue = 0.0
    left_cost = left_retail = unsold_cost = 0.0
    lines = []
    products = []
    missing = []
    for name, item in stock.items():
        if item["used"]:
            used_cost += item["amount"] or 0.0
            continue
        cost, price, qty = item["cost"], item["price"], item["qty"]
        sold = total(sales, name)
        sold_to_date = total(sales, name, to_date_range)
        earned = total(revenue, name)
        invested += item["amount"] or 0.0
        if cost is None or qty is None:
            missing.append(name)
        if cost is not None:
            cogs += sold * cost
            costed_revenue += earned
        if earned or sold:
            profit = earned - sold * cost if cost is not None else None
            products.append({
                "name": name,
                "sold": sold,
                "revenue": earned,
                "profit": profit,
                "margin": share(profit, earned) if profit is not None else None,
            })
        if qty is None:
            continue
        left = qty - sold_to_date
        left_cost += left * (cost or 0.0)
        left_retail += left * (price or 0.0)
        if cost is not None and price:
            if sold == 0:
                unsold_cost += left * cost
            lines.append({
                "name": name,
                "qty": qty,
                "sold": sold,
                "sold_to_date": sold_to_date,
                "left": left,
                "cost": cost,
                "price": price,
                "tied_up": left * cost,
                "margin": (price - cost) / price * 100,
            })

    restock = sorted(
        (line for line in lines if line["sold_to_date"] > 0 and line["left"] <= 1),
        key=lambda line: (-line["sold_to_date"], line["left"]),
    )
    unsold = [line for line in lines if line["sold"] == 0]
    slow = sorted(unsold, key=lambda line: -line["tied_up"])
    low_margin = sorted(
        (line for line in lines if line["left"] > 0),
        key=lambda line: (line["margin"], -line["tied_up"]),
    )
    top_products = sorted(products, key=lambda p: -p["revenue"])
    top_profit = sorted(
        (p for p in products if p["profit"] is not None),
        key=lambda p: -p["profit"],
    )

    kept = total_revenue - cogs - total_commission

    # Data checks: a job with no money, or money with no job
    mismatches = []
    for name, values in revenue.items():
        for day in dates:
            earned = values.get(day, 0.0)
            count = sales.get(name, {}).get(day, 0.0)
            if (earned > 0) != (count > 0):
                mismatches.append({
                    "date": day.isoformat(),
                    "label": label(day),
                    "item": name,
                    "revenue": earned,
                    "units": count,
                })
    mismatches.sort(key=lambda row: row["date"], reverse=True)

    return {
        "has_data": True,
        "period": period,
        "days": day_count,
        "trading_days": sum(1 for d in daily if d["revenue"] > 0),
        "revenue": total_revenue,
        "service_revenue": service_revenue,
        "product_revenue": product_revenue,
        "service_share": share(service_revenue, total_revenue),
        "cogs": cogs,
        "cogs_share": share(cogs, total_revenue),
        "commission": total_commission,
        "commission_total_share": share(total_commission, total_revenue),
        "kept": kept,
        "kept_share": share(kept, total_revenue),
        "average_day": total_revenue / day_count,
        "best_day": max(daily, key=lambda d: d["revenue"]),
        "best_weekday": max(weekdays, key=lambda w: w["average"]),
        "worst_weekday": min(weekdays, key=lambda w: w["average"]),
        "daily": daily,
        "weekdays": weekdays,
        "services": services,
        "top_five_share": share(top_five, service_revenue),
        "jobs": sum(item["jobs"] for item in ranked),
        "rates": rates,
        "commission_rate": share(total_commission, commission_service),
        "commission_share": share(total_commission, commission_revenue),
        "rate_low": min(paid_rates) if paid_rates else 0,
        "rate_high": max(paid_rates) if paid_rates else 0,
        "invested": invested,
        "used_cost": used_cost,
        "left_cost": left_cost,
        "left_retail": left_retail,
        "sold_share": share(cogs, invested),
        "months_to_clear": left_cost / cogs * day_count / 30 if cogs else None,
        "product_margin": share(costed_revenue - cogs, costed_revenue),
        "product_profit": costed_revenue - cogs,
        "unsold_cost": unsold_cost,
        "unsold_lines": len(unsold),
        "priced_lines": len(lines),
        "restock": restock,
        "slow": slow,
        "lines": sorted(lines, key=lambda line: line["name"].lower()),
        "low_margin": low_margin,
        "top_products": top_products,
        "top_profit": top_profit,
        "mismatches": mismatches,
        "missing": missing,
    }


# --------------------------------------------------------------------------
# Table export
# --------------------------------------------------------------------------

NUMBER = re.compile(r"^-?\d[\d,]*(\.\d+)?$")
PERCENT = re.compile(r"^-?\d+(\.\d+)?%$")


def read_export_request():
    """Validate the table sent by the page. Rows arrive exactly as displayed."""
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        abort(400)
    headers, rows = data.get("headers"), data.get("rows")
    if not isinstance(headers, list) or not isinstance(rows, list):
        abort(400)
    if not 1 <= len(headers) <= EXPORT_MAX_COLUMNS or len(rows) > EXPORT_MAX_ROWS:
        abort(400)
    if any(not isinstance(row, list) or len(row) != len(headers) for row in rows):
        abort(400)
    return {
        "title": str(data.get("title") or "Report")[:120],
        "period": str(data.get("period") or "")[:80],
        "headers": [str(text) for text in headers],
        "rows": [[str(text) for text in row] for row in rows],
    }


def is_numeric(text):
    return bool(NUMBER.match(text) or PERCENT.match(text))


def build_excel(table):
    book = Workbook()
    sheet = book.active
    sheet.title = re.sub(r"[\\/*?:\[\]]", " ", table["title"])[:31] or "Report"

    sheet["A1"] = f"Lloeva Beauty: {table['title']}"
    sheet["A1"].font = Font(bold=True, size=14, color=BRAND_ROSE)
    sheet["A2"] = table["period"]
    sheet["A2"].font = Font(italic=True, color="7A5566")

    head_row = 4
    for column, text in enumerate(table["headers"], start=1):
        head = sheet.cell(row=head_row, column=column, value=text)
        head.font = Font(bold=True, color="FFFFFF")
        head.fill = PatternFill("solid", fgColor=BRAND_ROSE)
        head.alignment = Alignment(horizontal="left" if column == 1 else "center")

    for row_number, row in enumerate(table["rows"], start=head_row + 1):
        for column, text in enumerate(row, start=1):
            item = sheet.cell(row=row_number, column=column)
            if column > 1 and NUMBER.match(text):
                number = float(text.replace(",", ""))
                item.value = number
                item.number_format = "#,##0" if number.is_integer() else "#,##0.0"
            elif column > 1 and PERCENT.match(text):
                item.value = float(text.rstrip("%")) / 100
                item.number_format = "0%"
            else:
                item.value = text
                item.data_type = "s"  # text only, so no cell can act as a formula

    for column, text in enumerate(table["headers"], start=1):
        longest = max([len(text)] + [len(row[column - 1]) for row in table["rows"]])
        sheet.column_dimensions[get_column_letter(column)].width = min(longest + 3, 48)

    last_row = head_row + len(table["rows"])
    last_column = get_column_letter(len(table["headers"]))
    sheet.auto_filter.ref = f"A{head_row}:{last_column}{last_row}"
    sheet.freeze_panes = sheet.cell(row=head_row + 1, column=1)

    output = BytesIO()
    book.save(output)
    output.seek(0)
    return output


def build_pdf(table):
    headers, rows = table["headers"], table["rows"]
    size = landscape(A4) if len(headers) > 4 else A4
    margin = 14 * mm
    width = size[0] - 2 * margin
    rose = colors.HexColor(f"#{BRAND_ROSE}")
    blush = colors.HexColor(f"#{BRAND_BLUSH}")

    title_style = ParagraphStyle("title", fontName="Times-Bold", fontSize=19, leading=23, textColor=rose)
    period_style = ParagraphStyle("period", fontName="Helvetica", fontSize=9.5, leading=13, textColor=colors.HexColor("#7A5566"))
    left = ParagraphStyle("left", fontName="Helvetica", fontSize=8.5, leading=11)
    right = ParagraphStyle("right", parent=left, alignment=TA_RIGHT)
    head_left = ParagraphStyle("head_left", parent=left, fontName="Helvetica-Bold", textColor=colors.white)
    head_right = ParagraphStyle("head_right", parent=head_left, alignment=TA_RIGHT)

    numeric = [
        index > 0 and bool(rows) and all(is_numeric(row[index]) for row in rows)
        for index in range(len(headers))
    ]
    data = [[
        Paragraph(escape(text), head_right if numeric[index] else head_left)
        for index, text in enumerate(headers)
    ]]
    for row in rows:
        data.append([
            Paragraph(escape(text), right if numeric[index] else left)
            for index, text in enumerate(row)
        ])

    weights = [
        min(max([len(headers[index])] + [len(row[index]) for row in rows]), 40) + 2
        for index in range(len(headers))
    ]
    widths = [width * weight / sum(weights) for weight in weights]

    grid = Table(data, colWidths=widths, repeatRows=1)
    grid.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), rose),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, blush]),
        ("LINEBELOW", (0, 1), (-1, -1), 0.25, colors.HexColor("#FAE3EB")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))

    def footer(canvas, document):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#7A5566"))
        canvas.drawString(margin, 8 * mm, f"Lloeva Beauty. Exported {date.today():%d %b %Y}.")
        canvas.drawRightString(size[0] - margin, 8 * mm, f"Page {document.page}")
        canvas.restoreState()

    output = BytesIO()
    document = SimpleDocTemplate(
        output, pagesize=size, title=f"Lloeva Beauty: {table['title']}", author="Lloeva Beauty",
        leftMargin=margin, rightMargin=margin, topMargin=margin, bottomMargin=16 * mm,
    )
    document.build(
        [
            Paragraph(escape(f"Lloeva Beauty: {table['title']}"), title_style),
            Paragraph(escape(table["period"]), period_style),
            Spacer(1, 5 * mm),
            grid,
        ],
        onFirstPage=footer,
        onLaterPages=footer,
    )
    output.seek(0)
    return output


# --------------------------------------------------------------------------
# Cache and routes
# --------------------------------------------------------------------------

_cache = {"data": None, "loaded": 0.0, "updated": ""}
_lock = threading.Lock()


def get_data(force=False):
    """Return (parsed sheet, last updated text, warning)."""
    with _lock:
        fresh = time.time() - _cache["loaded"] < CACHE_SECONDS
        if _cache["data"] and fresh and not force:
            return _cache["data"], _cache["updated"], None
        try:
            _cache["data"] = parse_sheet(fetch_sheet())
            _cache["loaded"] = time.time()
            _cache["updated"] = datetime.now().strftime("%d %b %Y, %H:%M")
            return _cache["data"], _cache["updated"], None
        except Exception as error:
            app.logger.exception("Could not load the sheet")
            if _cache["data"]:
                warning = "Could not reach the sheet, so this is the last saved data."
                return _cache["data"], _cache["updated"], warning
            raise RuntimeError(f"{type(error).__name__}: {error}") from error


@app.template_filter("ksh")
def ksh(value):
    return f"{value:,.0f}"


@app.template_filter("qty")
def qty(value):
    return f"{value:,.1f}".rstrip("0").rstrip(".")


@app.route("/")
def index():
    force = request.args.get("refresh") == "1"
    try:
        data, updated, warning = get_data(force)
    except RuntimeError as error:
        return render_template("error.html", message=str(error)), 503

    current = resolve_filter(request.args, data["dates"])
    report = build_report(data, current["start"], current["end"])
    tab = request.args.get("tab")
    if tab not in dict(PAGE_TABS):
        tab = PAGE_TABS[0][0]

    keep = {"view": current["view"], "tab": tab}
    if current["view"] == "custom":
        keep.update(start=current["start"].isoformat(), end=current["end"].isoformat())
    elif current["key"]:
        keep["period"] = current["key"]

    return render_template(
        "index.html",
        r=report,
        f=current,
        views=VIEWS,
        tabs=PAGE_TABS,
        tab=tab,
        refresh_url=url_for("index", refresh=1, **keep),
        updated=updated,
        warning=warning,
    )


@app.post("/export/<kind>")
def export_table(kind):
    builders = {
        "xlsx": (build_excel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        "pdf": (build_pdf, "application/pdf"),
    }
    if kind not in builders:
        abort(404)
    build, mimetype = builders[kind]
    table = read_export_request()
    name = re.sub(r"[^\w ()-]+", " ", f"Lloeva Beauty - {table['title']} - {table['period']}")
    return send_file(
        build(table), mimetype=mimetype, as_attachment=True,
        download_name=f"{name.strip()}.{kind}",
    )


if __name__ == "__main__":
    print(f"Lloeva Beauty dashboard: http://127.0.0.1:{PORT}")
    app.run(host="127.0.0.1", port=PORT, debug=False)
