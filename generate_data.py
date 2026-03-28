import pandas as pd
import numpy as np
from faker import Faker
from datetime import datetime, timedelta
import random
import os

fake = Faker('en_AU')
np.random.seed(42)
random.seed(42)

# ── 5 AusMart Sydney stores ──────────────────────────────
STORES = [
    {"store_id": 1, "store_name": "AusMart CBD",         "suburb": "Sydney CBD",    "region": "CBD",          "manager": "Sarah Chen",    "target_monthly": 120000},
    {"store_id": 2, "store_name": "AusMart Parramatta",  "suburb": "Parramatta",    "region": "Western",      "manager": "James Wilson",  "target_monthly": 95000},
    {"store_id": 3, "store_name": "AusMart Chatswood",   "suburb": "Chatswood",     "region": "North Shore",  "manager": "Priya Patel",   "target_monthly": 88000},
    {"store_id": 4, "store_name": "AusMart Bondi",       "suburb": "Bondi",         "region": "Eastern",      "manager": "Michael Brown", "target_monthly": 82000},
    {"store_id": 5, "store_name": "AusMart Penrith",     "suburb": "Penrith",       "region": "Western",      "manager": "Lisa Nguyen",   "target_monthly": 75000},
]

# ── 50 products across 5 categories ─────────────────────
PRODUCTS = [
    {"product_id": 1,  "name": "Full Cream Milk 2L",        "category": "Grocery",     "cost": 2.20,  "price": 3.80,  "margin_pct": 0.42},
    {"product_id": 2,  "name": "Sourdough Bread 750g",      "category": "Grocery",     "cost": 2.50,  "price": 5.50,  "margin_pct": 0.55},
    {"product_id": 3,  "name": "Free Range Eggs 12pk",      "category": "Grocery",     "cost": 4.20,  "price": 7.50,  "margin_pct": 0.44},
    {"product_id": 4,  "name": "Chicken Breast 1kg",        "category": "Grocery",     "cost": 7.00,  "price": 13.00, "margin_pct": 0.46},
    {"product_id": 5,  "name": "Cheddar Cheese 500g",       "category": "Grocery",     "cost": 4.50,  "price": 8.50,  "margin_pct": 0.47},
    {"product_id": 6,  "name": "Greek Yoghurt 1kg",         "category": "Grocery",     "cost": 3.80,  "price": 7.00,  "margin_pct": 0.46},
    {"product_id": 7,  "name": "Pasta 500g",                "category": "Grocery",     "cost": 1.20,  "price": 2.80,  "margin_pct": 0.57},
    {"product_id": 8,  "name": "Olive Oil 500ml",           "category": "Grocery",     "cost": 5.50,  "price": 11.00, "margin_pct": 0.50},
    {"product_id": 9,  "name": "Canned Tomatoes 400g",      "category": "Grocery",     "cost": 0.90,  "price": 2.20,  "margin_pct": 0.59},
    {"product_id": 10, "name": "Orange Juice 2L",           "category": "Grocery",     "cost": 3.20,  "price": 6.50,  "margin_pct": 0.51},
    {"product_id": 11, "name": "Samsung 55 TV",             "category": "Electronics", "cost": 650,   "price": 999,   "margin_pct": 0.35},
    {"product_id": 12, "name": "iPhone 15 Case",            "category": "Electronics", "cost": 8.00,  "price": 29.99, "margin_pct": 0.73},
    {"product_id": 13, "name": "Wireless Earbuds",          "category": "Electronics", "cost": 45,    "price": 99,    "margin_pct": 0.55},
    {"product_id": 14, "name": "USB-C Cable 2m",            "category": "Electronics", "cost": 5.00,  "price": 19.99, "margin_pct": 0.75},
    {"product_id": 15, "name": "Portable Charger 20000mAh", "category": "Electronics", "cost": 22,    "price": 59.99, "margin_pct": 0.63},
    {"product_id": 16, "name": "Bluetooth Speaker",         "category": "Electronics", "cost": 35,    "price": 89.99, "margin_pct": 0.61},
    {"product_id": 17, "name": "Smart Watch",               "category": "Electronics", "cost": 120,   "price": 299,   "margin_pct": 0.60},
    {"product_id": 18, "name": "Laptop Stand",              "category": "Electronics", "cost": 18,    "price": 49.99, "margin_pct": 0.64},
    {"product_id": 19, "name": "Wireless Mouse",            "category": "Electronics", "cost": 15,    "price": 39.99, "margin_pct": 0.63},
    {"product_id": 20, "name": "HDMI Cable 3m",             "category": "Electronics", "cost": 6.00,  "price": 24.99, "margin_pct": 0.76},
    {"product_id": 21, "name": "Cotton T-Shirt",            "category": "Clothing",    "cost": 8.00,  "price": 29.99, "margin_pct": 0.73},
    {"product_id": 22, "name": "Denim Jeans",               "category": "Clothing",    "cost": 22,    "price": 79.99, "margin_pct": 0.73},
    {"product_id": 23, "name": "Winter Jacket",             "category": "Clothing",    "cost": 45,    "price": 149,   "margin_pct": 0.70},
    {"product_id": 24, "name": "Running Shoes",             "category": "Clothing",    "cost": 55,    "price": 159,   "margin_pct": 0.65},
    {"product_id": 25, "name": "Sports Socks 5pk",          "category": "Clothing",    "cost": 5.00,  "price": 19.99, "margin_pct": 0.75},
    {"product_id": 26, "name": "Polo Shirt",                "category": "Clothing",    "cost": 12,    "price": 49.99, "margin_pct": 0.76},
    {"product_id": 27, "name": "Yoga Pants",                "category": "Clothing",    "cost": 18,    "price": 64.99, "margin_pct": 0.72},
    {"product_id": 28, "name": "Baseball Cap",              "category": "Clothing",    "cost": 7.00,  "price": 24.99, "margin_pct": 0.72},
    {"product_id": 29, "name": "Leather Belt",              "category": "Clothing",    "cost": 10,    "price": 39.99, "margin_pct": 0.75},
    {"product_id": 30, "name": "Scarf",                     "category": "Clothing",    "cost": 8.00,  "price": 29.99, "margin_pct": 0.73},
    {"product_id": 31, "name": "Shampoo 400ml",             "category": "Health",      "cost": 3.50,  "price": 9.99,  "margin_pct": 0.65},
    {"product_id": 32, "name": "Vitamin C 500mg 60pk",      "category": "Health",      "cost": 6.00,  "price": 19.99, "margin_pct": 0.70},
    {"product_id": 33, "name": "Sunscreen SPF50+ 200ml",    "category": "Health",      "cost": 5.50,  "price": 16.99, "margin_pct": 0.68},
    {"product_id": 34, "name": "Hand Sanitiser 500ml",      "category": "Health",      "cost": 2.50,  "price": 7.99,  "margin_pct": 0.69},
    {"product_id": 35, "name": "Paracetamol 20pk",          "category": "Health",      "cost": 2.00,  "price": 6.99,  "margin_pct": 0.71},
    {"product_id": 36, "name": "Moisturiser 200ml",         "category": "Health",      "cost": 7.00,  "price": 22.99, "margin_pct": 0.70},
    {"product_id": 37, "name": "Toothbrush Electric",       "category": "Health",      "cost": 18,    "price": 49.99, "margin_pct": 0.64},
    {"product_id": 38, "name": "Floss Picks 150pk",         "category": "Health",      "cost": 2.20,  "price": 6.99,  "margin_pct": 0.69},
    {"product_id": 39, "name": "Face Mask 10pk",            "category": "Health",      "cost": 4.00,  "price": 12.99, "margin_pct": 0.69},
    {"product_id": 40, "name": "Deodorant Roll-on",         "category": "Health",      "cost": 2.80,  "price": 8.99,  "margin_pct": 0.69},
    {"product_id": 41, "name": "Coffee Mug",                "category": "Home",        "cost": 4.00,  "price": 14.99, "margin_pct": 0.73},
    {"product_id": 42, "name": "Frying Pan 28cm",           "category": "Home",        "cost": 18,    "price": 49.99, "margin_pct": 0.64},
    {"product_id": 43, "name": "Dish Soap 500ml",           "category": "Home",        "cost": 1.80,  "price": 5.99,  "margin_pct": 0.70},
    {"product_id": 44, "name": "Bath Towel Set",            "category": "Home",        "cost": 15,    "price": 44.99, "margin_pct": 0.67},
    {"product_id": 45, "name": "Scented Candle",            "category": "Home",        "cost": 6.00,  "price": 19.99, "margin_pct": 0.70},
    {"product_id": 46, "name": "Storage Basket",            "category": "Home",        "cost": 8.00,  "price": 24.99, "margin_pct": 0.68},
    {"product_id": 47, "name": "Picture Frame A4",          "category": "Home",        "cost": 5.00,  "price": 16.99, "margin_pct": 0.71},
    {"product_id": 48, "name": "Laundry Powder 2kg",        "category": "Home",        "cost": 6.50,  "price": 18.99, "margin_pct": 0.66},
    {"product_id": 49, "name": "Cutting Board",             "category": "Home",        "cost": 9.00,  "price": 27.99, "margin_pct": 0.68},
    {"product_id": 50, "name": "Plant Pot 20cm",            "category": "Home",        "cost": 5.50,  "price": 17.99, "margin_pct": 0.69},
]

# ── 20 staff members ─────────────────────────────────────
STAFF = [
    {"staff_id": 1,  "name": "Emma Thompson",   "store_id": 1, "role": "Sales Associate",  "salary": 52000, "start_date": "2022-03-15"},
    {"staff_id": 2,  "name": "Liam Johnson",    "store_id": 1, "role": "Senior Associate",  "salary": 61000, "start_date": "2021-07-01"},
    {"staff_id": 3,  "name": "Olivia Williams", "store_id": 1, "role": "Sales Associate",  "salary": 50000, "start_date": "2023-01-10"},
    {"staff_id": 4,  "name": "Noah Brown",      "store_id": 1, "role": "Team Leader",      "salary": 72000, "start_date": "2020-05-20"},
    {"staff_id": 5,  "name": "Ava Jones",       "store_id": 2, "role": "Sales Associate",  "salary": 51000, "start_date": "2022-09-01"},
    {"staff_id": 6,  "name": "William Davis",   "store_id": 2, "role": "Senior Associate",  "salary": 60000, "start_date": "2021-11-15"},
    {"staff_id": 7,  "name": "Sophia Miller",   "store_id": 2, "role": "Sales Associate",  "salary": 50000, "start_date": "2023-04-01"},
    {"staff_id": 8,  "name": "James Wilson",    "store_id": 2, "role": "Team Leader",      "salary": 70000, "start_date": "2020-08-10"},
    {"staff_id": 9,  "name": "Isabella Moore",  "store_id": 3, "role": "Sales Associate",  "salary": 52000, "start_date": "2022-06-01"},
    {"staff_id": 10, "name": "Benjamin Taylor", "store_id": 3, "role": "Senior Associate",  "salary": 62000, "start_date": "2021-03-20"},
    {"staff_id": 11, "name": "Mia Anderson",    "store_id": 3, "role": "Sales Associate",  "salary": 51000, "start_date": "2023-02-14"},
    {"staff_id": 12, "name": "Elijah Thomas",   "store_id": 3, "role": "Team Leader",      "salary": 71000, "start_date": "2020-11-01"},
    {"staff_id": 13, "name": "Charlotte Jackson","store_id": 4, "role": "Sales Associate", "salary": 52000, "start_date": "2022-08-01"},
    {"staff_id": 14, "name": "Lucas White",     "store_id": 4, "role": "Senior Associate",  "salary": 60000, "start_date": "2021-05-15"},
    {"staff_id": 15, "name": "Amelia Harris",   "store_id": 4, "role": "Sales Associate",  "salary": 50000, "start_date": "2023-03-01"},
    {"staff_id": 16, "name": "Mason Martin",    "store_id": 4, "role": "Team Leader",      "salary": 69000, "start_date": "2021-01-10"},
    {"staff_id": 17, "name": "Harper Garcia",   "store_id": 5, "role": "Sales Associate",  "salary": 50000, "start_date": "2022-10-01"},
    {"staff_id": 18, "name": "Ethan Martinez",  "store_id": 5, "role": "Senior Associate",  "salary": 59000, "start_date": "2021-09-20"},
    {"staff_id": 19, "name": "Evelyn Robinson", "store_id": 5, "role": "Sales Associate",  "salary": 49000, "start_date": "2023-05-01"},
    {"staff_id": 20, "name": "Alexander Clark", "store_id": 5, "role": "Team Leader",      "salary": 68000, "start_date": "2021-02-15"},
]

# ── Sydney suburbs for customers ─────────────────────────
SYDNEY_SUBURBS = [
    "Surry Hills", "Newtown", "Glebe", "Pyrmont", "Ultimo",
    "Parramatta", "Blacktown", "Penrith", "Liverpool", "Campbelltown",
    "Chatswood", "North Sydney", "Mosman", "Manly", "Dee Why",
    "Bondi", "Coogee", "Randwick", "Maroubra", "Cronulla",
    "Strathfield", "Burwood", "Hurstville", "Kogarah", "Miranda",
    "Castle Hill", "Baulkham Hills", "Rouse Hill", "Kellyville", "Bella Vista"
]

def get_seasonal_multiplier(date):
    month = date.month
    day   = date.weekday()

    # Monthly seasonality — Australian retail patterns
    monthly = {
        1: 0.75,  # Jan — post Christmas quiet
        2: 0.70,  # Feb — slowest month
        3: 0.85,  # Mar — autumn pick up
        4: 0.88,  # Apr — Easter boost
        5: 0.90,  # May — steady
        6: 1.10,  # Jun — EOFY sales
        7: 0.85,  # Jul — winter slow
        8: 0.88,  # Aug — slight pick up
        9: 0.92,  # Sep — spring
        10: 0.95, # Oct — pre Christmas build
        11: 1.15, # Nov — Black Friday
        12: 1.40, # Dec — Christmas peak
    }

    # Day of week multiplier
    daily = {
        0: 0.85,  # Monday
        1: 0.88,  # Tuesday
        2: 0.90,  # Wednesday
        3: 0.92,  # Thursday
        4: 1.10,  # Friday
        5: 1.25,  # Saturday — busiest
        6: 1.15,  # Sunday
    }

    return monthly[month] * daily[day]


def get_store_multiplier(store_id):
    # CBD highest volume, Penrith lowest
    multipliers = {1: 1.30, 2: 1.10, 3: 1.05, 4: 1.00, 5: 0.90}
    return multipliers[store_id]


def generate_customers(n=1000):
    print(f"Generating {n} customers...")
    customers = []
    for i in range(1, n + 1):
        gender    = random.choice(["M", "F"])
        first     = fake.first_name_male() if gender == "M" else fake.first_name_female()
        last      = fake.last_name()
        suburb    = random.choice(SYDNEY_SUBURBS)
        age_group = random.choices(
            ["18-25", "26-35", "36-45", "46-55", "55+"],
            weights=[15, 28, 25, 20, 12]
        )[0]
        segment = random.choices(
            ["Premium", "Regular", "Budget"],
            weights=[20, 55, 25]
        )[0]

        customers.append({
            "customer_id":    i,
            "first_name":     first,
            "last_name":      last,
            "email":          fake.email(),
            "phone":          fake.phone_number(),
            "suburb":         suburb,
            "age_group":      age_group,
            "segment":        segment,
            "loyalty_member": random.choices([True, False], weights=[65, 35])[0],
            "join_date":      fake.date_between(start_date="-3y", end_date="-1m"),
        })
    return pd.DataFrame(customers)


def generate_transactions(customers_df, n=12000):
    print(f"Generating {n} transactions...")
    transactions = []
    start_date   = datetime(2024, 1, 1)
    end_date     = datetime(2025, 12, 31)
    date_range   = (end_date - start_date).days

    product_ids  = [p["product_id"] for p in PRODUCTS]
    store_ids    = [s["store_id"]   for s in STORES]
    customer_ids = customers_df["customer_id"].tolist()

    # Premium customers buy more expensive items more often
    segment_map  = dict(zip(customers_df["customer_id"], customers_df["segment"]))

    for i in range(1, n + 1):
        sale_date   = start_date + timedelta(days=random.randint(0, date_range))
        store_id    = random.choices(store_ids, weights=[30, 22, 20, 17, 11])[0]
        customer_id = random.choice(customer_ids)
        segment     = segment_map[customer_id]

        # Product selection weighted by customer segment
        if segment == "Premium":
            product = random.choices(PRODUCTS, weights=[
                1 if p["price"] < 20 else 3 if p["price"] < 100 else 5
                for p in PRODUCTS
            ])[0]
        elif segment == "Budget":
            product = random.choices(PRODUCTS, weights=[
                5 if p["price"] < 20 else 2 if p["price"] < 100 else 1
                for p in PRODUCTS
            ])[0]
        else:
            product = random.choice(PRODUCTS)

        # Quantity — most purchases are 1-2 items
        qty = random.choices([1, 2, 3, 4, 5], weights=[55, 25, 12, 5, 3])[0]

        # Apply seasonal multiplier to simulate realistic patterns
        seasonal   = get_seasonal_multiplier(sale_date)
        store_mult = get_store_multiplier(store_id)

        # Skip transaction based on multiplier (simulates busy/quiet periods)
        if random.random() > (seasonal * store_mult * 0.6):
            continue

        # Staff assigned to this store
        store_staff = [s for s in STAFF if s["store_id"] == store_id]
        staff       = random.choice(store_staff)

        unit_price    = round(product["price"] * random.uniform(0.95, 1.05), 2)
        unit_cost     = product["cost"]
        total_revenue = round(unit_price * qty, 2)
        total_cost    = round(unit_cost  * qty, 2)
        profit        = round(total_revenue - total_cost, 2)
        margin_pct    = round((profit / total_revenue) * 100, 1) if total_revenue > 0 else 0

        # Payment method
        payment = random.choices(
            ["Card", "Cash", "Digital Wallet", "Loyalty Points"],
            weights=[55, 15, 25, 5]
        )[0]

        transactions.append({
            "transaction_id":  i,
            "sale_date":       sale_date.strftime("%Y-%m-%d"),
            "store_id":        store_id,
            "customer_id":     customer_id,
            "product_id":      product["product_id"],
            "staff_id":        staff["staff_id"],
            "quantity":        qty,
            "unit_price":      unit_price,
            "unit_cost":       unit_cost,
            "total_revenue":   total_revenue,
            "total_cost":      total_cost,
            "profit":          profit,
            "margin_pct":      margin_pct,
            "payment_method":  payment,
            "month":           sale_date.strftime("%Y-%m"),
            "day_of_week":     sale_date.strftime("%A"),
            "quarter":         f"Q{((sale_date.month - 1) // 3) + 1}",
            "year":            sale_date.year,
        })

    df = pd.DataFrame(transactions)
    df = df.reset_index(drop=True)
    df["transaction_id"] = range(1, len(df) + 1)
    print(f"Generated {len(df)} actual transactions after seasonal filtering")
    return df


def save_all_data(customers_df, transactions_df):
    os.makedirs("data", exist_ok=True)

    stores_df  = pd.DataFrame(STORES)
    products_df = pd.DataFrame(PRODUCTS)
    staff_df   = pd.DataFrame(STAFF)

    stores_df.to_csv("data/stores.csv",       index=False)
    products_df.to_csv("data/products.csv",   index=False)
    staff_df.to_csv("data/staff.csv",         index=False)
    customers_df.to_csv("data/customers.csv", index=False)
    transactions_df.to_csv("data/transactions.csv", index=False)

    print("\nAll data saved to data/ folder:")
    print(f"  Stores:       {len(stores_df)} rows")
    print(f"  Products:     {len(products_df)} rows")
    print(f"  Staff:        {len(staff_df)} rows")
    print(f"  Customers:    {len(customers_df)} rows")
    print(f"  Transactions: {len(transactions_df)} rows")
    print(f"  Total rows:   {len(stores_df)+len(products_df)+len(staff_df)+len(customers_df)+len(transactions_df)}")


def run():
    print("="*55)
    print("AusMart Sydney — Data Generator")
    print("="*55)

    customers_df    = generate_customers(1000)
    transactions_df = generate_transactions(customers_df, 12000)

    save_all_data(customers_df, transactions_df)

    print("\nSample transactions:")
    print(transactions_df[["transaction_id","sale_date","store_id",
                            "product_id","quantity","total_revenue","profit"]].head(10).to_string(index=False))
    print("="*55)
    print("Data generation complete!")
    print("="*55)


if __name__ == "__main__":
    run()