import pandas as pd
import pytest
from openpyxl import Workbook

from cleaner.reader import (STANDARD_COLUMNS, find_header_row, normalize_header,
                            read_all, read_orders)

SOUTH_HEADER = ["order_id", "date", "customer_name", "e-mail", "phone_number",
                "city", "product", "quantity", "price"]
ORDER = ["S-1", "24.03.2026", "Ana Smith", "ana@example.com", "555-111-2222",
         "NYC", "USB-C Hub", "2", "$39.00"]


def write_xlsx(path, rows):
    wb = Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(path)
    return path


# ---------- normalize_header ----------

@pytest.mark.parametrize("raw, expected", [
    ("Order ID", "order id"),
    ("  Order   ID ", "order id"),
    ("E-MAIL", "e-mail"),
    ("Price (USD)", "price (usd)"),
    (None, ""),
])
def test_normalize_header(raw, expected):
    assert normalize_header(raw) == expected


# ---------- find_header_row ----------

def test_header_in_first_row():
    raw = pd.DataFrame([SOUTH_HEADER, ORDER])
    assert find_header_row(raw) == 0


def test_header_below_title_rows():
    raw = pd.DataFrame([
        ["Online shop - orders report"] + [None] * 8,
        ["Generated: 04/01/2026"] + [None] * 8,
        [None] * 9,
        SOUTH_HEADER,
        ORDER,
    ])
    assert find_header_row(raw) == 3


def test_no_header_raises():
    raw = pd.DataFrame([["a", "b", "c"], [1, 2, 3]])
    with pytest.raises(ValueError, match="header"):
        find_header_row(raw)


# ---------- read_orders ----------

def test_columns_get_standard_names(tmp_path):
    path = write_xlsx(tmp_path / "south.xlsx", [SOUTH_HEADER, ORDER])
    df = read_orders(path)
    assert list(df.columns) == STANDARD_COLUMNS + ["source_file", "source_row"]
    assert df.loc[0, "customer"] == "Ana Smith"
    assert df.loc[0, "unit_price"] == "$39.00"


def test_values_are_not_converted(tmp_path):
    path = write_xlsx(tmp_path / "south.xlsx", [SOUTH_HEADER, ORDER])
    df = read_orders(path)
    assert df.loc[0, "quantity"] == "2"          # still text: cleaning comes later
    assert df.loc[0, "order_date"] == "24.03.2026"


def test_title_rows_are_skipped(tmp_path):
    path = write_xlsx(tmp_path / "online.xlsx", [
        ["Online shop - orders report"],
        ["Generated: 04/01/2026"],
        [],
        ["ID", "Date", "Name", "Email Address", "Phone", "Location", "Item", "Qty", "Price (USD)"],
        ["W-1", "2/5/2026", "Bob Jones", "bob@example.com", "5551112222", "LA", "Webcam HD", 1, 59.9],
    ])
    df = read_orders(path)
    assert len(df) == 1
    assert df.loc[0, "order_id"] == "W-1"
    assert df.loc[0, "city"] == "LA"


def test_empty_and_total_rows_are_dropped(tmp_path):
    path = write_xlsx(tmp_path / "south.xlsx", [
        SOUTH_HEADER,
        ORDER,
        [],
        ["S-2"] + ORDER[1:],
        [],
        ["TOTAL", None, None, None, None, None, None, None, 78.0],
    ])
    df = read_orders(path)
    assert list(df["order_id"]) == ["S-1", "S-2"]


def test_source_file_and_excel_row_are_kept(tmp_path):
    path = write_xlsx(tmp_path / "south.xlsx", [SOUTH_HEADER, ORDER, [], ["S-2"] + ORDER[1:]])
    df = read_orders(path)
    assert list(df["source_file"]) == ["south.xlsx", "south.xlsx"]
    assert list(df["source_row"]) == [2, 4]      # Excel row numbers, header is row 1


def test_unknown_column_raises(tmp_path):
    header = SOUTH_HEADER[:-1] + ["discount"]
    path = write_xlsx(tmp_path / "south.xlsx", [header, ORDER])
    with pytest.raises(ValueError, match="discount"):
        read_orders(path)

def test_missing_column_raises(tmp_path):
    header = SOUTH_HEADER[:-1]                      # no price column
    path = write_xlsx(tmp_path / "south.xlsx", [header, ORDER[:-1]])
    with pytest.raises(ValueError, match="unit_price"):
        read_orders(path)


def test_column_mapped_twice_raises(tmp_path):
    header = SOUTH_HEADER + ["Order Date"]          # 'date' and 'Order Date'
    path = write_xlsx(tmp_path / "south.xlsx", [header, ORDER + ["24.03.2026"]])
    with pytest.raises(ValueError, match="order_date"):
        read_orders(path)


# ---------- read_all ----------

def test_read_all_combines_files(tmp_path):
    write_xlsx(tmp_path / "a.xlsx", [SOUTH_HEADER, ORDER])
    write_xlsx(tmp_path / "b.xlsx", [SOUTH_HEADER, ["S-2"] + ORDER[1:], ["S-3"] + ORDER[1:]])
    df = read_all(tmp_path)
    assert len(df) == 3
    assert sorted(df["source_file"].unique()) == ["a.xlsx", "b.xlsx"]


def test_read_all_empty_folder_raises(tmp_path):
    with pytest.raises(ValueError, match="No .xlsx files"):
        read_all(tmp_path)


def test_read_all_missing_folder_raises(tmp_path):
    with pytest.raises(FileNotFoundError, match="Folder not found"):
        read_all(tmp_path / "no_such_folder")