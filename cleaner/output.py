from io import BytesIO
"""Writes the cleaning result to a formatted Excel workbook."""
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT_PATH = Path("output/clean_orders.xlsx")

# Report keys in the order the client should read them, with plain-English labels.
REPORT_LABELS = {
    "rows_read": "Rows read from all files",
    "rows_clean": "Clean orders",
    "rows_rejected": "Rows rejected (see Rejected sheet)",
    "rows_with_issues": "Clean orders with a note in 'issues'",
    "duplicates_removed": "Duplicate rows removed",
    "conflicting_orders": "Orders with conflicting copies",
    "customer_changed": "Customer names fixed (case, spaces)",
    "product_changed": "Product names matched to the catalogue",
    "product_unknown": "Unknown products",
    "email_changed": "Emails fixed (case, spaces)",
    "email_invalid": "Invalid emails removed",
    "order_date_invalid": "Unreadable dates",
    "unit_price_from_text": "Prices converted from text",
    "unit_price_invalid": "Unreadable prices",
    "quantity_from_text": "Quantities converted from text",
    "quantity_invalid": "Unreadable quantities",
    "city_changed": "City names standardized",
    "city_invalid": "Unknown cities",
    "phone_changed": "Phone numbers standardized",
    "phone_invalid": "Invalid phone numbers removed",
}

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
DATE_FORMAT = "yyyy-mm-dd"
MONEY_FORMAT = '"$"#,##0.00'


def report_table(report, orders):
    """DataFrame with columns Metric, Value:
    1. every key from REPORT_LABELS that is in report, in REPORT_LABELS order, with its label
    2. any report key that has no label, under its own name (so nothing is lost)
    3. a last row 'Total revenue of clean orders (USD)' = sum of line_total, rounded to 2"""
    rows = [(label, report[key]) for key, label in REPORT_LABELS.items() if key in report]
    # Safety net: a new step's count still shows up even if nobody gave it a label.
    rows += [(key, value) for key, value in report.items() if key not in REPORT_LABELS]
    rows.append(("Total revenue of clean orders (USD)",
                 round(float(orders["line_total"].sum()), 2)))
    return pd.DataFrame(rows, columns=["Metric", "Value"])


def style_sheet(ws, date_columns=(), money_columns=(), wrap_columns=()):
    """Formats one openpyxl worksheet: styled and frozen header, filters,
    column widths, date/money formats and wrapped text by column name."""
    headers = [cell.value for cell in ws[1]]
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    ws.freeze_panes = "A2"               # header stays visible while scrolling
    ws.auto_filter.ref = ws.dimensions   # filter arrows over all data, e.g. 'A1:M518'

    for idx, name in enumerate(headers, start=1):
        letter = get_column_letter(idx)
        lengths = [len(str(c.value)) for c in ws[letter] if c.value is not None]
        ws.column_dimensions[letter].width = min(max(lengths, default=0) + 2, 60)

        for cell in ws[letter][1:]:      # data cells only, not the header
            if name in date_columns:
                cell.number_format = DATE_FORMAT
            elif name in money_columns:
                cell.number_format = MONEY_FORMAT
            elif name in wrap_columns:
                cell.alignment = Alignment(wrap_text=True, vertical="top")


def write_workbook(orders, rejected, report, path=OUTPUT_PATH):
    """Writes the Orders, Rejected and Report sheets to path and returns the path.
    Creates the output folder if it doesn't exist."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    _write_sheets(path, orders, rejected, report)
    return path


def workbook_bytes(orders, rejected, report):
    """The same workbook as write_workbook, but in memory (for a download button)."""
    buffer = BytesIO()
    _write_sheets(buffer, orders, rejected, report)
    return buffer.getvalue()


def _write_sheets(target, orders, rejected, report):
    """Writes and formats the three sheets. target is a file path or an in-memory buffer."""
    # One `with` block: the file is written and closed once, even if a step fails.
    with pd.ExcelWriter(target, engine="openpyxl") as writer:
        orders.to_excel(writer, sheet_name="Orders", index=False)
        rejected.to_excel(writer, sheet_name="Rejected", index=False)
        report_table(report, orders).to_excel(writer, sheet_name="Report", index=False)

        style_sheet(writer.sheets["Orders"],
                    date_columns={"order_date"},
                    money_columns={"unit_price", "line_total"})
        style_sheet(writer.sheets["Rejected"],
                    date_columns={"order_date"},
                    wrap_columns={"reason"})

        ws = writer.sheets["Report"]
        style_sheet(ws)
        # Only the revenue cell is money; the other values are row counts.
        ws.cell(row=ws.max_row, column=2).number_format = MONEY_FORMAT


if __name__ == "__main__":
    from cleaner.pipeline import run

    orders, rejected, report = run()
    path = write_workbook(orders, rejected, report)
    print(f"Saved {path}: {len(orders)} orders, {len(rejected)} rejected rows")