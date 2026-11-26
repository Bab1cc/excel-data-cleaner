"""Finds orders that appear more than once and keeps a single copy."""
import pandas as pd

# Columns that describe the sale itself. Copies of one order must agree on
# these; if they don't, we can't know which copy is right.
BUSINESS_COLUMNS = ["order_date", "customer", "product", "quantity", "unit_price"]
# Columns that only describe the customer; used to pick the most complete copy.
CONTACT_COLUMNS = ["customer", "email", "phone", "city"]
SOURCE_COLUMNS = ["source_file", "source_row"]
# When copies are equally complete, keep the one from the file that created the
# order: the stores are the originals, the online report re-exports some of them.
SOURCE_PRIORITY = ["store_north.xlsx", "store_south.xlsx", "online_shop.xlsx"]


def source_rank(source_file):
    """Position in SOURCE_PRIORITY; unknown files come last."""
    if source_file in SOURCE_PRIORITY:
        return SOURCE_PRIORITY.index(source_file)
    return len(SOURCE_PRIORITY)


def find_conflicts(group):
    """For rows that share one order ID: a list like ['quantity differs (2 vs 3)']
    for every BUSINESS_COLUMN with more than one distinct non-empty value."""
    conflicts = []
    for col in BUSINESS_COLUMNS:
        # pd.notna, not `is not None`: an empty cell can also be NaN.
        values = sorted({str(v) for v in group[col] if pd.notna(v)})
        if len(values) > 1:
            conflicts.append(f"{col} differs ({' vs '.join(values)})")
    return conflicts


def remove_duplicates(orders):
    """Keeps one row per order ID. Returns (unique, removed, report).

    unique:  orders with one row per order ID, in their original order
    removed: DataFrame with source_file, source_row, reason for every removed row,
             sorted by source_file, source_row
    report:  {'duplicates_removed': ..., 'conflicting_orders': ...}
    """
    # Rows without an order ID cannot be matched with anything, so they stay.
    # (groupby would silently drop them.)
    kept = list(orders.index[orders["order_id"].isna()])
    removed = []
    conflicting_orders = 0

    for order_id, group in orders.groupby("order_id", sort=False):
        if len(group) == 1:
            kept.append(group.index[0])
            continue

        conflicts = find_conflicts(group)
        if conflicts:
            # We can't tell which copy is right, so none of them is kept.
            conflicting_orders += 1
            reason = f"conflicting copies of order {order_id}: {'; '.join(conflicts)}"
            for _, row in group.iterrows():
                removed.append({"source_file": row["source_file"],
                                "source_row": row["source_row"],
                                "reason": reason})
            continue

        # Best copy first: fewest empty contact fields, then the most trusted
        # source, then the earliest row (so the result is always the same).
        missing = group[CONTACT_COLUMNS].isna().sum(axis=1)
        ranked = group.assign(_missing=missing, _rank=group["source_file"].map(source_rank))
        ranked = ranked.sort_values(["_missing", "_rank", "source_row"])

        best = ranked.iloc[0]
        kept.append(ranked.index[0])
        reason = (f"duplicate of order {order_id} "
                  f"(kept {best['source_file']} row {best['source_row']})")
        for _, row in ranked.iloc[1:].iterrows():
            removed.append({"source_file": row["source_file"],
                            "source_row": row["source_row"],
                            "reason": reason})

    unique = orders[orders.index.isin(kept)].reset_index(drop=True)

    # columns= keeps the right columns even when nothing was removed.
    removed = pd.DataFrame(removed, columns=SOURCE_COLUMNS + ["reason"])
    # int64 like the reader's source_row, so pipeline can merge on it even when empty.
    removed = removed.astype({"source_row": "int64"})
    removed = removed.sort_values(SOURCE_COLUMNS).reset_index(drop=True)

    report = {
        "duplicates_removed": int(removed["reason"].str.startswith("duplicate").sum()),
        "conflicting_orders": conflicting_orders,
    }
    return unique, removed, report