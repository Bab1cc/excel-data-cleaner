"""Reads the raw order files into one table with standard column names.

The reader only *finds* the data: it locates the real header, renames the
columns and throws away rows that are not orders (blank rows, TOTAL rows).
It does not convert or clean any values - that is the cleaner's job, and it
is much easier once every file has the same shape.
"""
from pathlib import Path

import pandas as pd

RAW_DIR = Path("data/raw")

STANDARD_COLUMNS = [
    "order_id",
    "order_date",
    "customer",
    "email",
    "phone",
    "city",
    "product",
    "quantity",
    "unit_price",
]

# Every header spelling seen in the raw files (after normalize_header),
# mapped to its standard name. See docs/data_issues.md, issue 9.
COLUMN_ALIASES = {
    "order id": "order_id",
    "order_id": "order_id",
    "id": "order_id",

    "order date": "order_date",
    "date": "order_date",

    "customer": "customer",
    "customer_name": "customer",
    "name": "customer",

    "email": "email",
    "e-mail": "email",
    "email address": "email",

    "phone": "phone",
    "phone_number": "phone",

    "city": "city",
    "location": "city",

    "product": "product",
    "item": "product",

    "qty": "quantity",
    "quantity": "quantity",

    "unit price": "unit_price",
    "price": "unit_price",
    "price (usd)": "unit_price",
}

# A row is the header when at least this many cells are known column names.
# 5 of 9 is more than half, so a title line or a data row can never pass,
# but a file with one renamed column is still found and the bad column is
# reported by name instead of a vague "no header".
MIN_HEADER_MATCHES = 5

# Values in the order ID column that mark a summary row, not an order.
SUMMARY_LABELS = {"total", "subtotal", "grand total"}

def normalize_header(name):
    """Turns a header cell into a comparable key: '  Order   ID ' -> 'order id'."""
    if pd.isna(name):
        return ""
    return " ".join(str(name).split()).lower()


def find_header_row(raw):
    """Returns the index of the header row in a sheet read with header=None."""
    for i in range(len(raw)):
        matches = sum(normalize_header(cell) in COLUMN_ALIASES for cell in raw.iloc[i])
        if matches >= MIN_HEADER_MATCHES:
            return i
    raise ValueError(
        f"No header row found: no row has {MIN_HEADER_MATCHES} or more known column names"
    )


def read_orders(path):
    """Reads one order file and returns its orders with standard column names.

    Values are returned exactly as Excel stores them. source_file and
    source_row (Excel row number) say where every row came from.
    """
    path = Path(path)
    # header=None: we find the header ourselves (title rows can sit above it).
    # dtype=object: no type guessing, so '2 pcs' and '$39.00' stay as text.
    raw = pd.read_excel(path, header=None, dtype=object)

    # 1. Find the header row.
    header_idx = find_header_row(raw)

    # 2. Translate column names; anything unknown is an error, not a guess.
    names = {}
    unknown = []
    for col, cell in raw.iloc[header_idx].items():
        key = normalize_header(cell)
        if key == "":
            continue
        if key in COLUMN_ALIASES:
            names[col] = COLUMN_ALIASES[key]
        else:
            unknown.append(str(cell).strip())
    if unknown:
        raise ValueError(f"{path.name}: unknown column(s): {', '.join(unknown)}")

    # Two different headers must not map to the same standard name.
    duplicates = {v for v in names.values() if list(names.values()).count(v) > 1}
    if duplicates:
        raise ValueError(f"{path.name}: columns mapped twice: {', '.join(sorted(duplicates))}")

    missing = [c for c in STANDARD_COLUMNS if c not in names.values()]
    if missing:
        raise ValueError(f"{path.name}: missing column(s): {', '.join(missing)}")

    # 3. Rows below the header, with standard names.
    df = raw.iloc[header_idx + 1:][list(names)].rename(columns=names)

    # 4. Remember where each row came from, *before* any row is dropped,
    #    otherwise the numbers would shift. Index 0 is Excel row 1.
    df["source_file"] = path.name
    df["source_row"] = df.index + 1

    # 5. Drop blank rows and TOTAL rows. Whitespace-only cells count as blank.
    data = df[STANDARD_COLUMNS]
    blank = data.apply(lambda col: col.map(_is_blank)).all(axis=1)
    total = data["order_id"].map(normalize_header).isin(SUMMARY_LABELS)
    df = df[~blank & ~total]

    # 6. Standard columns in a fixed order, clean 0..n-1 index.
    return df[STANDARD_COLUMNS + ["source_file", "source_row"]].reset_index(drop=True)


def read_all(folder):
    """Reads every .xlsx file in a folder into one table."""
    folder = Path(folder)
    # A typo in the folder name and an empty folder are different mistakes,
    # so they get different messages.
    if not folder.is_dir():
        raise FileNotFoundError(f"Folder not found: {folder}")
    # '~$name.xlsx' are lock files Excel creates while a file is open.
    paths = sorted(p for p in folder.glob("*.xlsx") if not p.name.startswith("~$"))
    if not paths:
        raise ValueError(f"No .xlsx files found in {folder}")
    return pd.concat([read_orders(p) for p in paths], ignore_index=True)


def _is_blank(value):
    return pd.isna(value) or (isinstance(value, str) and value.strip() == "")


if __name__ == "__main__":
    orders = read_all(RAW_DIR)
    print(orders.shape)
    print(orders.groupby("source_file").size())
