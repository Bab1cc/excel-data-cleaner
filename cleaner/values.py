"""Converts dates, prices and quantities from mixed text/number cells to real values."""
import math
import re
from datetime import date, datetime

import pandas as pd

from cleaner.text import _as_object, clean_whitespace

# How each source writes its dates as text. None means the file stores real
# Excel dates. A text date like '2/5/2026' is ambiguous on its own, so the
# format must come from knowing the source, never from guessing.
SOURCE_DATE_FORMATS = {
    "store_north.xlsx": None,
    "store_south.xlsx": "%d.%m.%Y",
    "online_shop.xlsx": "%m/%d/%Y",
}

# Everything that is not part of the number in '$1,249.00 USD': 'usd', '$', ',' and spaces.
_PRICE_JUNK = re.compile(r"(?i)usd|\$|,|\s")
# A whole number, optionally followed by a unit: '2', '2 pcs', '5 PCS', '1 pc'.
_QTY_PATTERN = re.compile(r"(?i)^(-?\d+)\s*(pcs?|pieces?|units?)?$")


def parse_date(value, fmt):
    """Real Excel date (datetime) -> date. Text -> parsed with fmt.
    Empty, impossible ('31.02.2026') or text when fmt is None -> None."""
    if value is None or pd.isna(value):
        return None
    # datetime is a subclass of date, so it must be checked first.
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if fmt is None:
        return None  # text in a source that should only have real dates
    try:
        return datetime.strptime(str(value).strip(), fmt).date()
    except ValueError:
        return None  # wrong format or impossible day such as 31.02.


def parse_price(value):
    """39.0 -> 39.0, '$1,249.00' -> 1249.0, '39.00 USD' -> 39.0. Rounded to 2 decimals.
    Empty or not a number -> None."""
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return round(float(value), 2)
    try:
        number = float(_PRICE_JUNK.sub("", str(value)))
    except ValueError:
        return None
    # float() also accepts 'nan' and 'inf', which are not prices.
    return round(number, 2) if math.isfinite(number) else None


def parse_quantity(value):
    """2 -> 2, 2.0 -> 2, '2 pcs' -> 2. Not a whole number or empty -> None."""
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return int(value) if float(value).is_integer() else None
    match = _QTY_PATTERN.match(str(value).strip())
    return int(match.group(1)) if match else None


def parse_value_columns(df, date_formats=None):
    """Returns (parsed copy of df, report). Raises ValueError for a source_file
    that has no entry in SOURCE_DATE_FORMATS.

    Report keys:
      order_date_invalid: had a date value, but it could not be read
      unit_price_from_text, quantity_from_text: text cells converted to numbers
      unit_price_invalid, quantity_invalid: had a value, but it is not a number
    """
    formats = SOURCE_DATE_FORMATS if date_formats is None else date_formats
    unknown = set(df["source_file"]) - set(formats)
    if unknown:
        raise ValueError(
            f"No date format known for: {', '.join(sorted(unknown))}. "
            f"Add it to SOURCE_DATE_FORMATS."
        )

    parsed = df.copy()
    report = {}

    # Dates need two columns (value + its source), so .map() is not enough.
    before = _as_object(df["order_date"])
    after = _object_series(
        [parse_date(value, formats[source])
         for value, source in zip(before, df["source_file"])],
        df.index,
    )
    report["order_date_invalid"] = int((_has_value(before) & after.isna()).sum())
    parsed["order_date"] = after

    for col, parse in [("unit_price", parse_price), ("quantity", parse_quantity)]:
        before = _as_object(df[col])
        after = _object_series([parse(value) for value in before], df.index)
        was_text = before.map(lambda value: isinstance(value, str))
        report[f"{col}_from_text"] = int((was_text & after.notna()).sum())
        report[f"{col}_invalid"] = int((_has_value(before) & after.isna()).sum())
        parsed[col] = after

    return parsed, report


def _object_series(values, index):
    """A Series that keeps exactly these Python values. Series.map() would infer
    a dtype instead: ints with one None become floats (2 -> 2.0, None -> NaN)."""
    return _as_object(pd.Series(values, index=index, dtype=object))


def _has_value(series):
    """True where the original cell was not empty (blank text counts as empty)."""
    return series.map(clean_whitespace).notna()


if __name__ == "__main__":
    from cleaner.reader import RAW_DIR, read_all

    orders = read_all(RAW_DIR)
    parsed, report = parse_value_columns(orders)
    for key, count in report.items():
        print(f"{key:>22}: {count}")
    dates = parsed["order_date"].dropna()
    print("dates from", min(dates), "to", max(dates))