from openpyxl import Workbook, load_workbook

from cleaner.__main__ import main

HEADER = ["order_id", "date", "customer_name", "e-mail", "phone_number",
          "city", "product", "quantity", "price"]


def write_south(path, rows):
    wb = Workbook()
    ws = wb.active
    for row in [HEADER] + rows:
        ws.append(row)
    wb.save(path)


def test_main_cleans_folder_and_saves_workbook(tmp_path, capsys):
    folder = tmp_path / "in"
    folder.mkdir()
    write_south(folder / "store_south.xlsx", [
        ["S-1", "24.03.2026", "Ana Smith", "ana@example.com", "555-111-2222",
         "NYC", "USB-C Hub", "2", "$39.00"],
        ["S-2", "24.03.2026", "Bob Jones", "bob@example.com", "555-111-3333",
         "LA", "Webcam HD", "0", "$59.90"],                  # rejected: quantity 0
    ])
    out = tmp_path / "result" / "clean.xlsx"

    assert main([str(folder), "-o", str(out)]) == 0

    printed = capsys.readouterr().out
    assert "Read 2 rows" in printed
    assert "1 clean orders" in printed
    assert "1 rejected rows" in printed
    assert load_workbook(out).sheetnames == ["Orders", "Rejected", "Report"]


def test_main_reports_error_without_traceback(tmp_path, capsys):
    assert main([str(tmp_path)]) == 1                         # folder with no .xlsx files

    captured = capsys.readouterr()
    assert captured.err.startswith("Error: No .xlsx files")
    assert captured.out == ""


def test_main_reports_missing_folder(tmp_path, capsys):
    assert main([str(tmp_path / "no_such_folder")]) == 1

    assert capsys.readouterr().err.startswith("Error: Folder not found")
