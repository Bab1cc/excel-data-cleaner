"""Runs every cleaning step and splits the result into clean orders and rejected rows."""
import pandas as pd

from cleaner.contacts import clean_contact_columns
from cleaner.reader import RAW_DIR, STANDARD_COLUMNS, read_all
from cleaner.text import clean_whitespace, clean_text_columns
from cleaner.values import parse_value_columns

SOURCE_COLUMNS = ["source_file", "source_row"]


def describe(problem, label, raw_value):
    """'invalid date (31.02.2026)' when there was a value, 'missing date' when there wasn't."""
    text = clean_whitespace(raw_value)
    if text:
        return f"{problem} {label} ({text})"
    return f"missing {label}"


def row_problems(clean, raw):
    """Returns (reasons, issues) for one row. `clean` and `raw` are the same row
    after and before cleaning (anything that supports row["column"]).

    reasons: problems that make the order unusable for sales reporting -> row is rejected
    issues:  problems with contact details only -> row is kept, issue is noted
    """
    reasons = []
    issues = []

    if clean["order_id"] is None:
        reasons.append("missing order ID")
    if clean["order_date"] is None:
        reasons.append(describe("invalid", "date", raw["order_date"]))
    if clean["product"] is None:
        reasons.append(describe("unknown", "product", raw["product"]))
    for col, label in [("quantity", "quantity"), ("unit_price", "price")]:
        value = clean[col]
        if value is None:
            reasons.append(describe("invalid", label, raw[col]))
        elif value <= 0:  # elif: None must never be compared with 0
            reasons.append(f"{label} must be positive ({value})")

    if clean["customer"] is None:
        issues.append("missing customer name")
    for col in ["email", "phone", "city"]:
        if clean[col] is None:
            issues.append(describe("removed invalid", col, raw[col]))

    return reasons, issues


def clean_orders(raw):
    """Runs all cleaning steps on the table from read_all().
    Returns (orders, rejected, report)."""
    # 1. Every step returns (df, report); the reports are merged into one.
    df, report = clean_text_columns(raw)
    df, values_report = parse_value_columns(df)
    report.update(values_report)
    df, contacts_report = clean_contact_columns(df)
    report.update(contacts_report)

    # 2. Order IDs only need their spaces trimmed.
    df["order_id"] = df["order_id"].map(clean_whitespace)

    # 3. One kind of "missing" in the whole table, so `is None` checks work.
    df = df.astype(object).where(df.notna(), None)

    # 4. Check every row against its clean and its original values.
    reasons = {}
    issues = {}
    for i in df.index:
        row_reasons, row_issues = row_problems(df.loc[i], raw.loc[i])
        reasons[i] = "; ".join(row_reasons)
        issues[i] = "; ".join(row_issues)
    reasons = pd.Series(reasons, dtype=object)
    issues = pd.Series(issues, dtype=object)
    is_rejected = reasons != ""

    # 5. Clean orders: cleaned values, the line total and the noted issues.
    orders = df[~is_rejected].copy()
    orders["line_total"] = [round(q * p, 2)
                            for q, p in zip(orders["quantity"], orders["unit_price"])]
    orders["issues"] = issues[~is_rejected]
    orders = orders[STANDARD_COLUMNS + ["line_total", "issues"] + SOURCE_COLUMNS]

    # 6. Rejected rows: the ORIGINAL values, so the client sees exactly what
    #    was in their file, with where it was and why it was rejected first.
    rejected = raw[is_rejected].copy()
    rejected["reason"] = reasons[is_rejected]
    rejected = rejected[SOURCE_COLUMNS + ["reason"] + STANDARD_COLUMNS]

    # 7. Clean 0..n-1 index and the row counts.
    orders = orders.reset_index(drop=True)
    rejected = rejected.reset_index(drop=True)
    report.update({
        "rows_read": len(raw),
        "rows_clean": len(orders),
        "rows_rejected": len(rejected),
        "rows_with_issues": int((orders["issues"] != "").sum()),
    })
    return orders, rejected, report


def run(folder=RAW_DIR):
    return clean_orders(read_all(folder))


if __name__ == "__main__":
    orders, rejected, report = run()
    for key, count in report.items():
        print(f"{key:>22}: {count}")
    print()
    print(rejected[["source_file", "source_row", "reason"]].to_string())