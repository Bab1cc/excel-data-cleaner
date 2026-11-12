"""Generates messy sample Excel files in data/raw/.

Simulates a small company whose three sales channels (two stores and an
online shop) each export orders in their own format, with typical real-world
problems: different column names, date formats and number formats, extra
whitespace, inconsistent spelling, duplicates, empty rows and invalid values.

The data is completely fake. Run from the project root:
    python scripts/make_messy_data.py
"""
import random
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import Workbook

SEED = 42
OUT_DIR = Path("data/raw")

FIRST_NAMES = ["James", "Mary", "John", "Patricia", "Robert", "Jennifer", "Michael",
               "Linda", "David", "Elizabeth", "William", "Susan", "Richard", "Jessica",
               "Thomas", "Sarah", "Daniel", "Karen", "Matthew", "Nancy"]
LAST_NAMES = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller",
              "Davis", "Rodriguez", "Martinez", "Wilson", "Anderson", "Taylor", "Moore"]

# canonical city -> spellings that appear in the exports
CITIES = {
    "New York": ["New York", "new york", "NYC", "New York City", "NY"],
    "Los Angeles": ["Los Angeles", "LA", "L.A.", "los angeles"],
    "Chicago": ["Chicago", "chicago", "CHICAGO", "Chicago, IL"],
    "Houston": ["Houston", "houston", "Houston TX"],
    "San Francisco": ["San Francisco", "SF", "San Fran", "san francisco"],
}

# product -> unit price in USD
PRODUCTS = {
    "Wireless Mouse": 24.99,
    "Mechanical Keyboard": 89.50,
    "USB-C Hub": 39.00,
    "27-inch Monitor": 1249.00,
    "Laptop Stand": 45.75,
    "Webcam HD": 59.90,
    "Noise Cancelling Headphones": 199.99,
}


def make_customers(rng, n):
    customers = []
    for _ in range(n):
        first, last = rng.choice(FIRST_NAMES), rng.choice(LAST_NAMES)
        digits = "".join(str(rng.randint(0, 9)) for _ in range(7))
        customers.append({
            "name": f"{first} {last}",
            "email": f"{first.lower()}.{last.lower()}{rng.randint(1, 99)}@example.com",
            "phone": "555" + digits,  # 10 digits, US style
            "city": rng.choice(list(CITIES)),
        })
    return customers


def make_orders(rng, customers, prefix, start_id, n):
    orders = []
    start = date(2026, 1, 1)
    for i in range(n):
        customer = rng.choice(customers)
        product = rng.choice(list(PRODUCTS))
        orders.append({
            "order_id": f"{prefix}-{start_id + i}",
            "date": start + timedelta(days=rng.randint(0, 89)),
            "product": product,
            "quantity": rng.choice([1, 1, 1, 2, 2, 3, 5]),
            "unit_price": PRODUCTS[product],
            **customer,
        })
    return orders


# ---------- ways of making a value messy ----------

def messy_text(rng, text):
    r = rng.random()
    if r < 0.10:
        return text.upper()
    if r < 0.20:
        return text.lower()
    if r < 0.32:
        return "  " + text + " "
    if r < 0.38:
        return text.replace(" ", "  ")
    return text


def messy_phone(rng, digits):
    a, b, c = digits[:3], digits[3:6], digits[6:]
    return rng.choice([
        f"({a}) {b}-{c}", f"{a}-{b}-{c}", f"{a}.{b}.{c}",
        f"+1 {a} {b} {c}", f"{a}{b}{c}", f"1-{a}-{b}-{c}",
    ])


def messy_email(rng, email):
    r = rng.random()
    if r < 0.04:
        return email.replace("@", "")            # invalid: no @
    if r < 0.07:
        return email.replace(".com", "")         # invalid: no domain ending
    if r < 0.15:
        return email.upper()
    if r < 0.22:
        return " " + email + "  "
    return email


def messy_city(rng, city):
    return rng.choice(CITIES[city])


def maybe_blank(rng, value, p=0.03):
    return None if rng.random() < p else value


# ---------- the three exports ----------

def write_store_north(rng, orders, path):
    """Clean-ish export: real Excel dates and numeric prices."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders"
    ws.append(["Order ID", "Order Date", "Customer", "Email", "Phone", "City",
               "Product", "Qty", "Unit Price"])
    for o in orders:
        ws.append([
            o["order_id"],
            datetime(o["date"].year, o["date"].month, o["date"].day),
            maybe_blank(rng, messy_text(rng, o["name"])),
            maybe_blank(rng, messy_email(rng, o["email"])),
            messy_phone(rng, o["phone"]),
            messy_city(rng, o["city"]),
            messy_text(rng, o["product"]),
            o["quantity"],
            o["unit_price"],
        ])
        if rng.random() < 0.04:                   # exact duplicate row
            ws.append([c.value for c in ws[ws.max_row]])
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].number_format = "dd/mm/yyyy"
    wb.save(path)


def write_store_south(rng, orders, path):
    """Dates and prices stored as text, European date format (day.month.year)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["order_id", "date", "customer_name", "e-mail", "phone_number",
               "city", "product", "quantity", "price"])
    for o in orders:
        qty = o["quantity"]
        if rng.random() < 0.08:
            qty = f"{qty} pcs"
        elif rng.random() < 0.02:
            qty = rng.choice([0, -1])
        price = o["unit_price"]
        price_text = rng.choice([f"${price:,.2f}", f"{price:.2f}", f"{price:,.2f} USD"])
        ws.append([
            o["order_id"],
            o["date"].strftime("%d.%m.%Y"),
            maybe_blank(rng, messy_text(rng, o["name"])),
            maybe_blank(rng, messy_email(rng, o["email"])),
            messy_phone(rng, o["phone"]),
            messy_city(rng, o["city"]),
            messy_text(rng, o["product"]),
            qty,
            maybe_blank(rng, price_text),
        ])
        if rng.random() < 0.05:                   # empty row in the middle
            ws.append([None] * 9)
    wb.save(path)


def write_online_shop(rng, orders, already_exported, path):
    """Report-style export: title rows above the header, US dates (month/day/year),
    a TOTAL row at the bottom and some orders that were already exported by a store."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Online Orders"
    ws.append(["Online shop - orders report"])
    ws.append([f"Generated: {date(2026, 4, 1):%m/%d/%Y}"])
    ws.append([])
    ws.append(["ID", "Date", "Name", "Email Address", "Phone", "Location",
               "Item", "Qty", "Price (USD)"])
    total = 0
    rows = orders + already_exported
    rng.shuffle(rows)
    for o in rows:
        d = o["date"]
        total += o["quantity"] * o["unit_price"]
        ws.append([
            o["order_id"],
            f"{d.month}/{d.day}/{d.year}",
            maybe_blank(rng, messy_text(rng, o["name"])),
            maybe_blank(rng, messy_email(rng, o["email"])),
            messy_phone(rng, o["phone"]),
            maybe_blank(rng, messy_city(rng, o["city"])),
            messy_text(rng, o["product"]),
            o["quantity"],
            o["unit_price"],
        ])
    ws.append([])
    ws.append(["TOTAL", None, None, None, None, None, None, None, round(total, 2)])
    wb.save(path)


def main():
    rng = random.Random(SEED)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    customers = make_customers(rng, 60)

    north = make_orders(rng, customers, "N", 1001, 180)
    south = make_orders(rng, customers, "S", 2001, 150)
    online = make_orders(rng, customers, "W", 3001, 200)

    write_store_north(rng, north, OUT_DIR / "store_north.xlsx")
    write_store_south(rng, south, OUT_DIR / "store_south.xlsx")
    write_online_shop(rng, online, rng.sample(north, 12), OUT_DIR / "online_shop.xlsx")

    for f in sorted(OUT_DIR.glob("*.xlsx")):
        print("created", f)


if __name__ == "__main__":
    main()