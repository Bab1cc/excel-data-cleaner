from datetime import date, datetime

import pandas as pd
import pytest

from cleaner.values import (parse_date, parse_price, parse_quantity,
                            parse_value_columns)


# ---------- parse_date ----------

@pytest.mark.parametrize("raw, fmt, expected", [
    (datetime(2026, 3, 24), None, date(2026, 3, 24)),           # real Excel date
    ("24.03.2026", "%d.%m.%Y", date(2026, 3, 24)),              # store_south
    ("2/5/2026", "%m/%d/%Y", date(2026, 2, 5)),                 # online_shop: Feb 5
    ("2/5/2026", "%d/%m/%Y", date(2026, 5, 2)),                 # same text, other meaning
    (" 24.03.2026 ", "%d.%m.%Y", date(2026, 3, 24)),
    ("31.02.2026", "%d.%m.%Y", None),                           # no such day
    ("24.03.2026", "%m/%d/%Y", None),                           # wrong format
    ("24.03.2026", None, None),                                 # text where a date was expected
    (None, "%d.%m.%Y", None),
    (float("nan"), None, None),
])
def test_parse_date(raw, fmt, expected):
    assert parse_date(raw, fmt) == expected


# ---------- parse_price ----------

@pytest.mark.parametrize("raw, expected", [
    (39.0, 39.0),
    (1249, 1249.0),
    ("$39.00", 39.0),
    ("39.00", 39.0),
    ("39.00 USD", 39.0),
    ("$1,249.00", 1249.0),
    ("1,249.00 USD", 1249.0),
    (24.999, 25.0),
    ("free", None),
    ("", None),
    (None, None),
])
def test_parse_price(raw, expected):
    assert parse_price(raw) == expected


# ---------- parse_quantity ----------

@pytest.mark.parametrize("raw, expected", [
    (2, 2),
    (2.0, 2),          # online_shop stores whole numbers as decimals
    ("3", 3),
    ("2 pcs", 2),
    ("5 PCS", 5),
    ("1 pc", 1),
    (0, 0),            # parsed; rejecting it is the validation step's job
    (-1, -1),
    (2.5, None),       # half an item is not a quantity
    ("two", None),
    (None, None),
])
def test_parse_quantity(raw, expected):
    assert parse_quantity(raw) == expected


# ---------- parse_value_columns ----------

@pytest.fixture
def orders():
    return pd.DataFrame({
        "order_id": ["N-1", "S-1", "S-2", "W-1"],
        "order_date": [datetime(2026, 1, 15), "24.03.2026", "31.02.2026", "2/5/2026"],
        "quantity": [1, "2 pcs", "two", 2.0],
        "unit_price": [24.99, "$1,249.00", None, 59.9],
        "source_file": ["store_north.xlsx", "store_south.xlsx",
                        "store_south.xlsx", "online_shop.xlsx"],
    })


def test_values_are_parsed(orders):
    parsed, _ = parse_value_columns(orders)
    assert list(parsed["order_date"]) == [date(2026, 1, 15), date(2026, 3, 24),
                                          None, date(2026, 2, 5)]
    assert list(parsed["quantity"]) == [1, 2, None, 2]
    assert list(parsed["unit_price"]) == [24.99, 1249.0, None, 59.9]


def test_report_counts(orders):
    _, report = parse_value_columns(orders)
    assert report == {
        "order_date_invalid": 1,     # '31.02.2026'
        "unit_price_from_text": 1,   # '$1,249.00'
        "unit_price_invalid": 0,     # a missing price is missing, not invalid
        "quantity_from_text": 1,     # '2 pcs'
        "quantity_invalid": 1,       # 'two'
    }


def test_unknown_source_file_raises(orders):
    orders.loc[0, "source_file"] = "new_store.xlsx"
    with pytest.raises(ValueError, match="new_store.xlsx"):
        parse_value_columns(orders)


def test_original_is_not_modified(orders):
    original = orders.copy()
    parse_value_columns(orders)
    pd.testing.assert_frame_equal(orders, original)


def test_date_formats_can_be_passed_in(orders):
    formats = {"store_north.xlsx": None, "store_south.xlsx": "%d.%m.%Y",
               "online_shop.xlsx": "%d/%m/%Y"}            # online read as day first
    parsed, _ = parse_value_columns(orders, formats)
    assert parsed.loc[3, "order_date"] == date(2026, 5, 2)