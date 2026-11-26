"""Cleans text columns: customer names, product names and emails.

Every function takes one raw cell and returns the cleaned value, or None when
the value is empty or cannot be trusted. Nothing here deletes rows: deciding
what to do with a None is the validation step's job.
"""
import re

import pandas as pd

# The client's product catalogue: the only valid product names.
PRODUCTS = [
    "Wireless Mouse",
    "Mechanical Keyboard",
    "USB-C Hub",
    "27-inch Monitor",
    "Laptop Stand",
    "Webcam HD",
    "Noise Cancelling Headphones",
]

# letters/digits/._%+- , then @, then domain parts, then an ending of 2+ letters
EMAIL_PATTERN = re.compile(r"^[a-z0-9._%+-]+@[a-z0-9-]+(\.[a-z0-9-]+)*\.[a-z]{2,}$")

# Lookup key -> catalogue name, e.g. 'usb-c hub' -> 'USB-C Hub'. Built once from
# PRODUCTS, so adding a product to the catalogue is the only change ever needed.
_PRODUCT_KEYS = {" ".join(name.split()).lower(): name for name in PRODUCTS}


def clean_whitespace(value):
    """'  John   Smith ' -> 'John Smith'. Empty, blank or missing -> None."""
    if pd.isna(value):
        return None
    # split() with no argument also strips both ends and splits on tabs/newlines.
    text = " ".join(str(value).split())
    return text or None


def clean_name(value):
    """'  JOHN   smith ' -> 'John Smith'. Empty -> None."""
    text = clean_whitespace(value)
    return text.title() if text else None


def match_product(value):
    """Finds the catalogue name for a messy product value: 'usb-c  HUB' -> 'USB-C Hub'.
    Returns None for empty values and for products that are not in PRODUCTS."""
    text = clean_whitespace(value)
    if text is None:
        return None
    return _PRODUCT_KEYS.get(text.lower())


def clean_email(value):
    """'  ANA@Example.COM ' -> 'ana@example.com'. Invalid or empty -> None."""
    text = clean_whitespace(value)
    if text is None:
        return None
    text = text.lower()
    return text if EMAIL_PATTERN.match(text) else None


def clean_text_columns(df):
    """Returns (cleaned copy of df, report dict). The original df is not modified.

    Report keys:
      customer_changed, product_changed, email_changed:
          rows where the cleaned value exists and differs from the original
      product_unknown: rows that had a product, but it is not in the catalogue
      email_invalid:   rows that had an email, but it is not a valid address
    """
    cleaned = df.copy()
    report = {}
    # (column, cleaning function, name of the "could not clean" count or None)
    for col, clean, bad_key in [("customer", clean_name, None),
                                ("product", match_product, "product_unknown"),
                                ("email", clean_email, "email_invalid")]:
        before = _as_object(df[col])
        after = _as_object(before.map(clean))

        report[f"{col}_changed"] = int((after.notna() & (after != before)).sum())
        if bad_key:
            had_value = before.map(clean_whitespace).notna()
            report[bad_key] = int((had_value & after.isna()).sum())

        cleaned[col] = after
    return cleaned, report


def _as_object(series):
    """Plain Python objects with None for missing values (pandas may use NaN).
    Use it on a column before and after .map(), so comparisons with None work."""
    series = series.astype(object)
    return series.where(series.notna(), None)


if __name__ == "__main__":
    from cleaner.reader import RAW_DIR, read_all

    orders = read_all(RAW_DIR)
    cleaned, report = clean_text_columns(orders)
    for key, count in report.items():
        print(f"{key:>18}: {count}")
    print(cleaned[["customer", "email", "product"]].head(10).to_string())