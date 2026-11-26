import math

import pandas as pd
import pytest

from cleaner.text import (clean_email, clean_name, clean_text_columns,
                          clean_whitespace, match_product)


# ---------- clean_whitespace ----------

@pytest.mark.parametrize("raw, expected", [
    ("  John   Smith ", "John Smith"),
    ("already clean", "already clean"),
    ("", None),
    ("    ", None),
    (None, None),
    (math.nan, None),
    (123, "123"),
])
def test_clean_whitespace(raw, expected):
    assert clean_whitespace(raw) == expected


# ---------- clean_name ----------

@pytest.mark.parametrize("raw, expected", [
    ("  JOHN   smith ", "John Smith"),
    ("mary jones", "Mary Jones"),
    ("Linda  Davis", "Linda Davis"),
    ("Linda Davis", "Linda Davis"),
    ("   ", None),
    (None, None),
])
def test_clean_name(raw, expected):
    assert clean_name(raw) == expected


# ---------- match_product ----------

@pytest.mark.parametrize("raw, expected", [
    ("USB-C Hub", "USB-C Hub"),
    ("usb-c hub", "USB-C Hub"),
    ("  MECHANICAL KEYBOARD ", "Mechanical Keyboard"),
    ("27-inch  monitor", "27-inch Monitor"),
    ("Gaming Chair", None),
    (None, None),
])
def test_match_product(raw, expected):
    assert match_product(raw) == expected


# ---------- clean_email ----------

@pytest.mark.parametrize("raw, expected", [
    ("ana.smith12@example.com", "ana.smith12@example.com"),
    ("  ANA.SMITH12@Example.COM ", "ana.smith12@example.com"),
    ("ana.smith12example.com", None),    # no @
    ("ana.smith12@example", None),       # no domain ending
    ("ana smith@example.com", None),     # space inside
    ("", None),
    (None, None),
])
def test_clean_email(raw, expected):
    assert clean_email(raw) == expected


# ---------- clean_text_columns ----------

@pytest.fixture
def orders():
    return pd.DataFrame({
        "order_id": ["S-1", "S-2", "S-3", "S-4"],
        "customer": ["  ANA smith", "Bob Jones", None, "carl  lee"],
        "email": ["ANA@example.com", "bob@example", "carl@example.com", None],
        "product": ["usb-c hub", "Webcam HD", "Gaming Chair", "WEBCAM HD"],
    })


def test_values_are_cleaned(orders):
    cleaned, _ = clean_text_columns(orders)
    assert list(cleaned["customer"]) == ["Ana Smith", "Bob Jones", None, "Carl Lee"]
    assert list(cleaned["email"]) == ["ana@example.com", None, "carl@example.com", None]
    assert list(cleaned["product"]) == ["USB-C Hub", "Webcam HD", None, "Webcam HD"]


def test_report_counts(orders):
    _, report = clean_text_columns(orders)
    assert report == {
        "customer_changed": 2,   # '  ANA smith' and 'carl  lee'
        "product_changed": 2,    # 'usb-c hub' and 'WEBCAM HD'
        "product_unknown": 1,    # 'Gaming Chair'
        "email_changed": 1,      # 'ANA@example.com'
        "email_invalid": 1,      # 'bob@example' (an empty email is missing, not invalid)
    }


def test_original_is_not_modified(orders):
    original = orders.copy()
    clean_text_columns(orders)
    pd.testing.assert_frame_equal(orders, original)


def test_other_columns_are_untouched(orders):
    cleaned, _ = clean_text_columns(orders)
    assert list(cleaned["order_id"]) == ["S-1", "S-2", "S-3", "S-4"]