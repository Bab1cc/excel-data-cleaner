"""Standardizes cities and phone numbers."""
import re

import pandas as pd

from cleaner.text import _as_object, clean_whitespace

# Every city spelling seen in the data, grouped under its official name.
# 'NY' and 'LA' are documented assumptions (see docs/data_issues.md).
CITY_ALIASES = {
    "New York": ["New York", "NYC", "New York City", "NY"],
    "Los Angeles": ["Los Angeles", "LA", "L.A."],
    "Chicago": ["Chicago", "Chicago, IL"],
    "Houston": ["Houston", "Houston TX"],
    "San Francisco": ["San Francisco", "SF", "San Fran"],
}


def city_key(value):
    """Comparable key for a city: 'L.A.' -> 'la', ' Chicago,  IL ' -> 'chicago il'.
    Removes dots and commas, extra spaces, and lowercases. Empty -> ''."""
    if value is None or pd.isna(value):
        return ""
    text = re.sub(r"[.,]", "", str(value))
    return " ".join(text.split()).lower()


_CITY_LOOKUP = {
    city_key(alias): city
    for city, aliases in CITY_ALIASES.items()
    for alias in aliases
}



def match_city(value):
    """Official city name for a messy value: 'NYC' -> 'New York'.
    Empty or unknown city -> None."""
    key = city_key(value)
    if not key:
        return None
    return _CITY_LOOKUP.get(key)



def normalize_phone(value):
    """Any US phone format -> E.164: '(555) 123-4567' -> '+15551234567'.
    Accepts 10 digits, or 11 digits starting with the country code 1.
    Anything else, or empty -> None."""
    if value is None or pd.isna(value):
        return None
    # 5551234567.0 -> 5551234567, otherwise str() adds '.0' and the 0 becomes a digit.
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    digits = re.sub(r"\D", "", str(value))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]  # drop the US country code
    if len(digits) != 10:
        return None
    return "+1" + digits



def clean_contact_columns(df):
    """Returns (cleaned copy of df, report). The original df is not modified.

    Report keys:
      city_changed, phone_changed: cleaned value exists and differs from the original
      city_invalid, phone_invalid: had a value, but it could not be standardized
    """
    cleaned = df.copy()
    report = {}
    for col, clean in [("city", match_city), ("phone", normalize_phone)]:
        before = _as_object(df[col])
        after = _as_object(before.map(clean))
        had_value = before.map(clean_whitespace).notna()

        report[f"{col}_changed"] = int((after.notna() & (after != before)).sum())
        report[f"{col}_invalid"] = int((had_value & after.isna()).sum())
        cleaned[col] = after
    return cleaned, report



if __name__ == "__main__":
    from cleaner.reader import RAW_DIR, read_all

    orders = read_all(RAW_DIR)
    cleaned, report = clean_contact_columns(orders)
    for key, count in report.items():
        print(f"{key:>14}: {count}")
    print(cleaned["city"].value_counts(dropna=False).to_string())