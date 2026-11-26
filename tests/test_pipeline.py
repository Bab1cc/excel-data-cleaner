from datetime import date

import pandas as pd
import pytest

from cleaner.pipeline import clean_orders, describe, row_problems


# ---------- describe ----------

def test_describe_with_value():
    assert describe("invalid", "date", " 31.02.2026 ") == "invalid date (31.02.2026)"


def test_describe_without_value():
    assert describe("invalid", "date", None) == "missing date"


# ---------- row_problems ----------

GOOD = {"order_id": "S-1", "order_date": date(2026, 3, 24), "customer": "Ana Smith",
        "email": "ana@example.com", "phone": "+15551112222", "city": "New York",
        "product": "USB-C Hub", "quantity": 2, "unit_price": 39.0}


def problems(clean_changes=None, raw_changes=None):
    clean = {**GOOD, **(clean_changes or {})}
    raw = {**GOOD, **(raw_changes or {})}
    return row_problems(clean, raw)


def test_good_row_has_no_problems():
    assert problems() == ([], [])


def test_zero_quantity_is_rejected():
    reasons, _ = problems({"quantity": 0})
    assert reasons == ["quantity must be positive (0)"]


def test_unparsable_values_show_the_original():
    reasons, _ = problems({"quantity": None, "unit_price": None},
                          {"quantity": "two", "unit_price": None})
    assert reasons == ["invalid quantity (two)", "missing price"]


def test_unknown_product_is_rejected():
    reasons, _ = problems({"product": None}, {"product": "Gaming Chair"})
    assert reasons == ["unknown product (Gaming Chair)"]


def test_several_reasons_are_all_listed():
    reasons, _ = problems({"order_date": None, "unit_price": -5.0},
                          {"order_date": "31.02.2026"})
    assert reasons == ["invalid date (31.02.2026)", "price must be positive (-5.0)"]


def test_bad_contact_details_are_issues_not_reasons():
    reasons, issues = problems({"email": None, "city": None, "customer": None},
                               {"email": "ana@example", "city": None})
    assert reasons == []
    assert issues == ["missing customer name",
                      "removed invalid email (ana@example)",
                      "missing city"]


# ---------- clean_orders ----------

def raw_row(order_id, quantity="2", price="$39.00", email="ana@example.com", row=2):
    return {"order_id": order_id, "order_date": "24.03.2026", "customer": " ana SMITH ",
            "email": email, "phone": "(555) 111-2222", "city": "NYC",
            "product": "usb-c hub", "quantity": quantity, "unit_price": price,
            "source_file": "store_south.xlsx", "source_row": row}


@pytest.fixture
def raw():
    return pd.DataFrame([
        raw_row("S-1", row=2),
        raw_row("S-2", quantity="0", row=3),
        raw_row("S-3", price=None, row=4),
        raw_row("S-4", email="ana@example", row=5),
    ])


def test_rows_are_split(raw):
    orders, rejected, _ = clean_orders(raw)
    assert list(orders["order_id"]) == ["S-1", "S-4"]
    assert list(rejected["order_id"]) == ["S-2", "S-3"]


def test_clean_orders_have_clean_values_and_total(raw):
    orders, _, _ = clean_orders(raw)
    first = orders.iloc[0]
    assert first["customer"] == "Ana Smith"
    assert first["order_date"] == date(2026, 3, 24)
    assert first["city"] == "New York"
    assert first["quantity"] == 2
    assert first["line_total"] == 78.0
    assert first["issues"] == ""


def test_kept_row_with_issue(raw):
    orders, _, _ = clean_orders(raw)
    assert orders.iloc[1]["issues"] == "removed invalid email (ana@example)"


def test_rejected_rows_keep_original_values_and_location(raw):
    _, rejected, _ = clean_orders(raw)
    first = rejected.iloc[0]
    assert first["reason"] == "quantity must be positive (0)"
    assert first["source_file"] == "store_south.xlsx"
    assert first["source_row"] == 3
    assert first["customer"] == " ana SMITH "        # original, not cleaned
    assert list(rejected.columns[:3]) == ["source_file", "source_row", "reason"]


def test_report_counts_rows(raw):
    _, _, report = clean_orders(raw)
    assert report["rows_read"] == 4
    assert report["rows_clean"] == 2
    assert report["rows_rejected"] == 2
    assert report["rows_with_issues"] == 1
    assert report["email_invalid"] == 1             # step reports are included too


def test_duplicates_are_moved_to_rejected():
    raw = pd.DataFrame([raw_row("S-1", row=2), raw_row("S-1", row=3)])
    orders, rejected, report = clean_orders(raw)
    assert list(orders["source_row"]) == [2]
    assert list(rejected["source_row"]) == [3]
    assert rejected.loc[0, "reason"] == "duplicate of order S-1 (kept store_south.xlsx row 2)"
    assert report["duplicates_removed"] == 1
    assert report["rows_clean"] == 1