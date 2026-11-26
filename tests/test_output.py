from datetime import date

import pandas as pd
import pytest
from openpyxl import load_workbook

from cleaner.output import REPORT_LABELS, report_table, write_workbook


@pytest.fixture
def orders():
    return pd.DataFrame({
        "order_id": ["S-1", "S-2"],
        "order_date": [date(2026, 3, 24), date(2026, 1, 5)],
        "customer": ["Ana Smith", "Bob Jones"],
        "unit_price": [39.0, 1249.0],
        "quantity": [2, 1],
        "line_total": [78.0, 1249.0],
        "issues": ["", "missing city"],
    })


@pytest.fixture
def rejected():
    return pd.DataFrame({
        "source_file": ["store_south.xlsx"],
        "source_row": [23],
        "reason": ["quantity must be positive (0)"],
        "order_id": ["S-9"],
    })


@pytest.fixture
def report():
    return {"rows_read": 3, "rows_clean": 2, "rows_rejected": 1, "email_invalid": 0,
            "something_new": 7}


# ---------- report_table ----------

def test_report_uses_labels_in_fixed_order(report, orders):
    table = report_table(report, orders)
    assert list(table["Metric"][:4]) == [REPORT_LABELS["rows_read"], REPORT_LABELS["rows_clean"],
                                         REPORT_LABELS["rows_rejected"],
                                         REPORT_LABELS["email_invalid"]]


def test_report_keeps_unknown_keys_and_adds_revenue(report, orders):
    table = report_table(report, orders)
    assert ("something_new", 7) in list(table.itertuples(index=False, name=None))
    assert table.iloc[-1]["Metric"] == "Total revenue of clean orders (USD)"
    assert table.iloc[-1]["Value"] == 1327.0


# ---------- write_workbook ----------

@pytest.fixture
def workbook(tmp_path, orders, rejected, report):
    path = write_workbook(orders, rejected, report, tmp_path / "out" / "clean.xlsx")
    return load_workbook(path)


def test_three_sheets_in_order(workbook):
    assert workbook.sheetnames == ["Orders", "Rejected", "Report"]


def test_orders_sheet_content(workbook):
    ws = workbook["Orders"]
    assert [c.value for c in ws[1]] == ["order_id", "order_date", "customer", "unit_price",
                                       "quantity", "line_total", "issues"]
    assert ws.max_row == 3                       # header + 2 orders
    assert ws["A2"].value == "S-1"


def test_orders_sheet_formats(workbook):
    ws = workbook["Orders"]
    assert ws["B2"].number_format == "yyyy-mm-dd"      # order_date
    assert ws["D2"].number_format == '"$"#,##0.00'     # unit_price
    assert ws["F3"].number_format == '"$"#,##0.00'     # line_total
    assert ws["A1"].font.bold


def test_header_is_frozen_and_filtered(workbook):
    for ws in workbook.worksheets:
        assert ws.freeze_panes == "A2"
        assert ws.auto_filter.ref is not None


def test_rejected_reason_wraps(workbook):
    ws = workbook["Rejected"]
    assert ws["C1"].value == "reason"
    assert ws["C2"].alignment.wrap_text


def test_columns_are_wide_enough(workbook):
    ws = workbook["Rejected"]
    assert ws.column_dimensions["C"].width >= len("quantity must be positive (0)")