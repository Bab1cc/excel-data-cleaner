import pandas as pd
import pytest

from cleaner.contacts import clean_contact_columns, match_city, normalize_phone


# ---------- match_city ----------

@pytest.mark.parametrize("raw, expected", [
    ("New York", "New York"),
    ("new york", "New York"),
    ("  NYC ", "New York"),
    ("New York City", "New York"),
    ("NY", "New York"),
    ("L.A.", "Los Angeles"),
    ("la", "Los Angeles"),
    ("Chicago, IL", "Chicago"),
    ("CHICAGO", "Chicago"),
    ("Houston TX", "Houston"),
    ("San  Fran", "San Francisco"),
    ("Boston", None),        # not a known city: flag it, don't guess
    ("", None),
    (None, None),
])
def test_match_city(raw, expected):
    assert match_city(raw) == expected


# ---------- normalize_phone ----------

@pytest.mark.parametrize("raw, expected", [
    ("(555) 123-4567", "+15551234567"),
    ("555-123-4567", "+15551234567"),
    ("555.123.4567", "+15551234567"),
    ("+1 555 123 4567", "+15551234567"),
    ("1-555-123-4567", "+15551234567"),
    ("5551234567", "+15551234567"),
    (5551234567, "+15551234567"),        # Excel can store a phone as a number
    (5551234567.0, "+15551234567"),
    ("555-1234", None),                  # too short
    ("2-555-123-4567", None),            # 11 digits, but not a US country code
    ("call me", None),
    (None, None),
])
def test_normalize_phone(raw, expected):
    assert normalize_phone(raw) == expected


# ---------- clean_contact_columns ----------

@pytest.fixture
def orders():
    return pd.DataFrame({
        "order_id": ["S-1", "S-2", "S-3", "S-4"],
        "city": ["NYC", "Chicago", "Boston", None],
        "phone": ["(555) 123-4567", "+15550000000", "555-1234", None],
    })


def test_values_are_cleaned(orders):
    cleaned, _ = clean_contact_columns(orders)
    assert list(cleaned["city"]) == ["New York", "Chicago", None, None]
    assert list(cleaned["phone"]) == ["+15551234567", "+15550000000", None, None]


def test_report_counts(orders):
    _, report = clean_contact_columns(orders)
    assert report == {
        "city_changed": 1,       # 'NYC'
        "city_invalid": 1,       # 'Boston'
        "phone_changed": 1,      # '(555) 123-4567'
        "phone_invalid": 1,      # '555-1234'
    }


def test_original_is_not_modified(orders):
    original = orders.copy()
    clean_contact_columns(orders)
    pd.testing.assert_frame_equal(orders, original)