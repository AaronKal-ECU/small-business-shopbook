"""
Synthetic data generator for small-business-shopbook.

Every name, price, sale and order in the published workbook comes from this
script. Nothing in it is real business data. The random seed is fixed, so
running it again gives exactly the same rows.

    python tools/generate_data.py            # writes data/*.csv

The shop sells dash and audio fit-out parts (brackets, fascias, gauge pods,
panels). The catalogue keeps the workbook's three-level shape:

    Brand  ->  Size (vehicle class)  ->  Product (part + finish)

The simulation keeps stock honest: a product can only be sold (or written off,
or taken as owner drawings) after a RECEIVED supplier order has delivered it,
and never below zero.
"""
from __future__ import annotations

import csv
import random
from datetime import date, timedelta
from pathlib import Path

SEED = 2025
START = date(2025, 6, 2)   # first supplier order
END = date(2025, 8, 31)    # last trading day
ARRIVAL_DAYS = 5           # supplier lead time for a RECEIVED order

# --------------------------------------------------------------------------
# Catalogue
# --------------------------------------------------------------------------
# Parts in display order, with a base retail price range (AUD).
PARTS = {
    "Single-DIN Bracket": (29, 45),
    "Double-DIN Fascia": (49, 79),
    "Gauge Pod (Single)": (25, 39),
    "Gauge Pod (Triple)": (59, 95),
    "Phone Mount Bracket": (25, 39),
    "Switch Panel": (35, 59),
    "Speaker Spacer Rings": (19, 35),
    "USB/12V Socket Panel": (29, 49),
}

# Brand -> (vehicle classes, finishes, products per class, price multiplier)
BRANDS = {
    "Brand A": (["Hatch", "Sedan", "Wagon", "Ute", "SUV"],
                ["Matte Black ABS", "Gloss Black ABS", "Silver ABS"], 5, 1.00),
    "Brand B": (["Hatch", "Sedan", "SUV"],
                ["Matte Black ABS", "Carbon-look ABS", "Gloss Black ABS"], 7, 1.15),
    "Brand C": (["Ute", "SUV"],
                ["Matte Black ASA", "Carbon-look ASA", "Black PETG"], 6, 1.60),
}

RETAIL = [
    # (name, weight) - a handful of regulars do most of the retail buying
    ("Liam Carter", 9), ("Olivia Bennett", 8), ("Noah Fletcher", 7), ("Ava Mitchell", 6),
    ("Jack Sullivan", 5), ("Mia Robertson", 3), ("Ethan Walsh", 3), ("Chloe Harper", 3),
    ("Lucas Grant", 2), ("Grace Donovan", 2), ("Henry Lawson", 2), ("Zoe Whitfield", 2),
    ("Oscar Pearce", 2), ("Ruby Callahan", 2), ("Leo Matthews", 1), ("Isla Morgan", 1),
    ("Max Thornton", 1), ("Ella Drummond", 1), ("Archie Holt", 1), ("Sophie Langford", 1),
]
TRADE = [("Workshop 1", 6), ("Workshop 2", 4), ("Workshop 3", 3), ("Workshop 4", 2)]
TRADE_SHARE = 0.30          # share of sale lines that go to trade customers
WRITE_OFF = "WRITE-OFF"
OWNER_DRAWINGS = "OWNER DRAWINGS"


def build_catalogue(rng: random.Random) -> list[dict]:
    rows = []
    part_names = list(PARTS)
    for brand, (classes, finishes, per_class, mult) in BRANDS.items():
        for vclass in classes:
            combos = [(p, f) for p in part_names for f in finishes]
            picked = rng.sample(combos, per_class)
            picked.sort(key=lambda pf: (part_names.index(pf[0]), finishes.index(pf[1])))
            for part, finish in picked:
                lo, hi = PARTS[part]
                price = round(rng.randint(lo, hi) * mult * (1.1 if "ASA" in finish else 1.0)) - 0.05
                cost = round(price * rng.uniform(0.40, 0.55), 2)
                rows.append({"Brand": brand, "Size": vclass, "Product": f"{part} – {finish}",
                             "Cost": cost, "Price": price})
    return rows


def build_orders(rng: random.Random, catalogue: list[dict]) -> list[dict]:
    """Five supplier orders: three RECEIVED, one ORDERED, one PENDING."""
    ranged = rng.sample(catalogue, 36)           # the shop only ranges part of the catalogue
    plan = [
        (1, START, "RECEIVED", ranged[:26], (3, 6)),
        (2, date(2025, 6, 30), "RECEIVED", 14, (2, 5)),
        (3, date(2025, 7, 21), "RECEIVED", 14, (2, 5)),
        (4, date(2025, 8, 18), "ORDERED", 10, (2, 5)),
        (5, date(2025, 8, 29), "PENDING", 7, (2, 4)),
    ]
    orders = []
    for order_no, d, status, items, (qlo, qhi) in plan:
        if isinstance(items, int):
            items = rng.sample(ranged, items)
        for item in items:
            orders.append({"Order No.": order_no, "Brand": item["Brand"], "Size": item["Size"],
                           "Product": item["Product"], "Quantity": rng.randint(qlo, qhi),
                           "Unit Cost": item["Cost"], "Date": d, "Status": status})
    return orders


def key(row: dict) -> tuple[str, str, str]:
    return row["Brand"], row["Size"], row["Product"]


def build_sales(rng: random.Random, catalogue: list[dict], orders: list[dict]) -> list[dict]:
    price = {key(c): c for c in catalogue}
    arrivals: dict[date, list[dict]] = {}
    for o in orders:
        if o["Status"] == "RECEIVED":
            arrivals.setdefault(o["Date"] + timedelta(days=ARRIVAL_DAYS), []).append(o)

    stock: dict[tuple, int] = {}
    sales: list[dict] = []
    special = {WRITE_OFF: 3, OWNER_DRAWINGS: 5}

    d = START
    while d <= END:
        for o in arrivals.get(d, []):
            stock[key(o)] = stock.get(key(o), 0) + o["Quantity"]
        busy = d.weekday() >= 5                              # weekends are busier
        n_sales = rng.choices([0, 1, 2, 3], [2, 4, 3, 1] if busy else [4, 4, 1, 1])[0]
        for _ in range(n_sales):
            in_stock = [k for k, q in stock.items() if q > 0]
            if not in_stock:
                break
            k = rng.choice(in_stock)
            if rng.random() < TRADE_SHARE:
                who = rng.choices([t[0] for t in TRADE], [t[1] for t in TRADE])[0]
                qty = min(stock[k], rng.randint(2, 4))
                method = rng.choices(["PAYID", "UNPAID"], [75, 25])[0]
            else:
                who = rng.choices([c[0] for c in RETAIL], [c[1] for c in RETAIL])[0]
                qty = min(stock[k], rng.choices([1, 2], [85, 15])[0])
                method = rng.choices(["CASH", "PAYID", "UNPAID"], [52, 44, 4])[0]
            sales.append(_sale(d, who, k, qty, method, price))
            stock[k] -= qty
        # occasional write-off / owner drawings, spread across the period
        for who, left in special.items():
            if left and rng.random() < left / max(1, (END - d).days + 1) * 1.6:
                in_stock = [k for k, q in stock.items() if q > 0]
                if in_stock:
                    k = rng.choice(in_stock)
                    sales.append(_sale(d, who, k, 1, "OTHER", price))
                    stock[k] -= 1
                    special[who] -= 1
        d += timedelta(days=1)

    for i, s in enumerate(sales, start=1):
        s["Invoice"] = i
    return sales


def _sale(d, customer, k, qty, method, price) -> dict:
    c = price[k]
    return {"Invoice": None, "Date": d, "Customer": customer, "Brand": k[0], "Size": k[1],
            "Product": k[2], "Quantity": qty, "Unit Price": c["Price"], "Method": method,
            "Unit Cost": c["Cost"]}


def sale_total(s: dict) -> float:
    return 0.0 if s["Method"] == "OTHER" else s["Quantity"] * s["Unit Price"]


def build_withdrawals(rng: random.Random, sales: list[dict]) -> list[dict]:
    """Cash taken out of the till roughly every fortnight, never more than is in it."""
    reasons = ["Partner payout", "Supplier payment", "Bank deposit", "Partner payout",
               "Supplier payment", "Bank deposit"]
    out, withdrawn = [], 0.0
    d = date(2025, 6, 20)
    for reason in reasons:
        cash_in = sum(sale_total(s) for s in sales if s["Method"] == "CASH" and s["Date"] <= d)
        amount = int((cash_in - withdrawn) * rng.uniform(0.45, 0.75) / 10) * 10
        if amount > 0:
            out.append({"Date": d, "Cash Requested": amount, "Reason": reason})
            withdrawn += amount
        d += timedelta(days=rng.randint(12, 16))
        if d > END:
            break
    return out


def generate() -> dict[str, list[dict]]:
    rng = random.Random(SEED)
    catalogue = build_catalogue(rng)
    orders = build_orders(rng, catalogue)
    sales = build_sales(rng, catalogue, orders)
    withdrawals = build_withdrawals(rng, sales)
    return {"catalogue": catalogue, "orders": orders, "sales": sales, "withdrawals": withdrawals}


def write_csv(data: dict[str, list[dict]], out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    for name, rows in data.items():
        with open(out_dir / f"{name}.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)


if __name__ == "__main__":
    data = generate()
    write_csv(data, Path(__file__).resolve().parent.parent / "data")
    for name, rows in data.items():
        print(f"{name:12s} {len(rows):4d} rows")
