from datetime import date

import pandas as pd

from cleaner.duplicates import find_conflicts, remove_duplicates, source_rank


def order(order_id, source_file, source_row, **changes):
    row = {"order_id": order_id, "order_date": date(2026, 1, 15), "customer": "Ana Smith",
           "email": "ana@example.com", "phone": "+15551112222", "city": "New York",
           "product": "USB-C Hub", "quantity": 2, "unit_price": 39.0,
           "source_file": source_file, "source_row": source_row}
    row.update(changes)
    return row


def table(*rows):
    return pd.DataFrame(list(rows)).astype(object)


# ---------- source_rank ----------

def test_source_rank():
    assert source_rank("store_north.xlsx") < source_rank("online_shop.xlsx")
    assert source_rank("new_file.xlsx") == 3        # unknown files come last


# ---------- find_conflicts ----------

def test_identical_copies_have_no_conflicts():
    group = table(order("N-1", "store_north.xlsx", 2), order("N-1", "online_shop.xlsx", 9))
    assert find_conflicts(group) == []


def test_missing_value_is_not_a_conflict():
    group = table(order("N-1", "store_north.xlsx", 2, customer=None),
                  order("N-1", "online_shop.xlsx", 9))
    assert find_conflicts(group) == []


def test_different_quantity_is_a_conflict():
    group = table(order("N-1", "store_north.xlsx", 2, quantity=2),
                  order("N-1", "online_shop.xlsx", 9, quantity=3))
    assert find_conflicts(group) == ["quantity differs (2 vs 3)"]


# ---------- remove_duplicates ----------

def test_unique_orders_are_untouched():
    orders = table(order("N-1", "store_north.xlsx", 2), order("N-2", "store_north.xlsx", 3))
    unique, removed, report = remove_duplicates(orders)
    assert list(unique["order_id"]) == ["N-1", "N-2"]
    assert removed.empty
    assert report == {"duplicates_removed": 0, "conflicting_orders": 0}


def test_exact_duplicate_keeps_first_row():
    orders = table(order("N-1", "store_north.xlsx", 2), order("N-1", "store_north.xlsx", 3))
    unique, removed, _ = remove_duplicates(orders)
    assert list(unique["source_row"]) == [2]
    assert list(removed["source_row"]) == [3]
    assert removed.loc[0, "reason"] == "duplicate of order N-1 (kept store_north.xlsx row 2)"


def test_store_copy_wins_over_online_copy():
    orders = table(order("N-1", "online_shop.xlsx", 9), order("N-1", "store_north.xlsx", 50))
    unique, removed, _ = remove_duplicates(orders)
    assert unique.loc[0, "source_file"] == "store_north.xlsx"
    assert removed.loc[0, "source_file"] == "online_shop.xlsx"


def test_most_complete_copy_wins():
    orders = table(order("N-1", "store_north.xlsx", 2, email=None, city=None),
                   order("N-1", "online_shop.xlsx", 9))
    unique, _, _ = remove_duplicates(orders)
    assert unique.loc[0, "source_file"] == "online_shop.xlsx"


def test_conflicting_copies_are_all_removed():
    orders = table(order("N-1", "store_north.xlsx", 2, unit_price=39.0),
                   order("N-1", "online_shop.xlsx", 9, unit_price=45.0),
                   order("N-2", "store_north.xlsx", 3))
    unique, removed, report = remove_duplicates(orders)
    assert list(unique["order_id"]) == ["N-2"]
    assert len(removed) == 2
    assert removed.loc[0, "reason"] == ("conflicting copies of order N-1: "
                                        "unit_price differs (39.0 vs 45.0)")
    assert report == {"duplicates_removed": 0, "conflicting_orders": 1}