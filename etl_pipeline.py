import pandas as pd
import numpy as np
import mysql.connector
from datetime import datetime
from config import DB_CONFIG


def get_connection():
    return mysql.connector.connect(**DB_CONFIG)


# ══════════════════════════════════════════════════════
# EXTRACT — Read from CSV files (simulating Excel intake)
# ══════════════════════════════════════════════════════

def extract():
    print("="*55)
    print("STEP 1 — EXTRACT")
    print("="*55)

    data = {}

    files = {
        "stores":       "data/stores.csv",
        "products":     "data/products.csv",
        "staff":        "data/staff.csv",
        "customers":    "data/customers.csv",
        "transactions": "data/transactions.csv"
    }

    for name, path in files.items():
        df = pd.read_csv(path)
        data[name] = df
        print(f"Extracted {len(df):>6} rows from {path}")

    print(f"\nTotal rows extracted: {sum(len(v) for v in data.values())}")
    return data


# ══════════════════════════════════════════════════════
# TRANSFORM — Clean, validate and enrich each dataset
# ══════════════════════════════════════════════════════

def transform_stores(df):
    print("\n--- Transforming stores ---")
    before = len(df)

    # Clean text
    df["store_name"] = df["store_name"].str.strip()
    df["suburb"]     = df["suburb"].str.strip().str.title()
    df["region"]     = df["region"].str.strip().str.title()
    df["manager"]    = df["manager"].str.strip().str.title()

    # Validate target
    df = df[df["target_monthly"] > 0]

    # Enrich — add annual target
    df["target_annual"] = df["target_monthly"] * 12

    print(f"Stores: {before} → {len(df)} rows after cleaning")
    print(df[["store_name", "region", "target_monthly", "target_annual"]].to_string(index=False))
    return df


def transform_products(df):
    print("\n--- Transforming products ---")
    before = len(df)

    # Clean text
    df["name"]     = df["name"].str.strip()
    df["category"] = df["category"].str.strip().str.title()

    # Validate prices — cost must be less than price
    invalid = df[df["cost"] >= df["price"]]
    if len(invalid) > 0:
        print(f"WARNING: {len(invalid)} products where cost >= price — removing")
        df = df[df["cost"] < df["price"]]

    # Validate no negative prices
    df = df[(df["cost"] > 0) & (df["price"] > 0)]

    # Enrich — recalculate margin from cost and price
    df["margin_pct"]    = ((df["price"] - df["cost"]) / df["price"]).round(3)
    df["markup_amount"] = (df["price"] - df["cost"]).round(2)

    print(f"Products: {before} → {len(df)} rows after cleaning")
    print(df[["name", "category", "cost", "price", "margin_pct"]].to_string(index=False))
    return df


def transform_staff(df):
    print("\n--- Transforming staff ---")
    before = len(df)

    # Clean text
    df["name"] = df["name"].str.strip().str.title()
    df["role"] = df["role"].str.strip().str.title()

    # Fix date format
    df["start_date"] = pd.to_datetime(df["start_date"]).dt.strftime("%Y-%m-%d")

    # Validate salary — must be above minimum wage
    min_salary = 40000
    below_min  = df[df["salary"] < min_salary]
    if len(below_min) > 0:
        print(f"WARNING: {len(below_min)} staff below minimum salary threshold")

    # Enrich — calculate years of service
    today           = datetime.today()
    df["years_service"] = df["start_date"].apply(
        lambda x: round((today - datetime.strptime(x, "%Y-%m-%d")).days / 365.25, 1)
    )

    print(f"Staff: {before} → {len(df)} rows after cleaning")
    print(df[["name", "role", "salary", "years_service"]].to_string(index=False))
    return df


def transform_customers(df):
    print("\n--- Transforming customers ---")
    before = len(df)

    # Clean text
    df["first_name"] = df["first_name"].str.strip().str.title()
    df["last_name"]  = df["last_name"].str.strip().str.title()
    df["suburb"]     = df["suburb"].str.strip().str.title()
    df["segment"]    = df["segment"].str.strip().str.title()
    df["age_group"]  = df["age_group"].str.strip()

    # Validate email — must contain @ symbol
    before_email = len(df)
    df = df[df["email"].str.contains("@", na=False)]
    if before_email != len(df):
        print(f"Removed {before_email - len(df)} rows with invalid emails")

    # Fix date
    df["join_date"] = pd.to_datetime(df["join_date"]).dt.strftime("%Y-%m-%d")

    # Fix boolean
    df["loyalty_member"] = df["loyalty_member"].map(
        {True: 1, False: 0, "True": 1, "False": 0, 1: 1, 0: 0}
    ).fillna(0).astype(int)

    # Remove duplicates on email
    before_dedup = len(df)
    df = df.drop_duplicates(subset=["email"])
    if before_dedup != len(df):
        print(f"Removed {before_dedup - len(df)} duplicate emails")

    # Enrich — full name column
    df["full_name"] = df["first_name"] + " " + df["last_name"]

    print(f"Customers: {before} → {len(df)} rows after cleaning")
    print(f"Segments: {df['segment'].value_counts().to_dict()}")
    print(f"Loyalty members: {df['loyalty_member'].sum()} of {len(df)}")
    return df


def transform_transactions(df):
    print("\n--- Transforming transactions ---")
    before = len(df)

    # Fix date
    df["sale_date"] = pd.to_datetime(df["sale_date"]).dt.strftime("%Y-%m-%d")

    # Validate quantities
    before_qty = len(df)
    df = df[df["quantity"] > 0]
    if before_qty != len(df):
        print(f"Removed {before_qty - len(df)} rows with zero/negative quantity")

    # Validate prices
    before_price = len(df)
    df = df[df["unit_price"] > 0]
    if before_price != len(df):
        print(f"Removed {before_price - len(df)} rows with invalid prices")

    # Validate revenue makes sense
    before_rev = len(df)
    df = df[df["total_revenue"] > 0]
    if before_rev != len(df):
        print(f"Removed {before_rev - len(df)} rows with zero revenue")

    # Recalculate totals to ensure accuracy
    df["total_revenue"] = (df["unit_price"] * df["quantity"]).round(2)
    df["total_cost"]    = (df["unit_cost"]  * df["quantity"]).round(2)
    df["profit"]        = (df["total_revenue"] - df["total_cost"]).round(2)
    df["margin_pct"]    = ((df["profit"] / df["total_revenue"]) * 100).round(1)

    # Remove duplicates on transaction_id
    before_dedup = len(df)
    df = df.drop_duplicates(subset=["transaction_id"])
    if before_dedup != len(df):
        print(f"Removed {before_dedup - len(df)} duplicate transactions")

    # Validate dates are not in the future
    today           = datetime.today().strftime("%Y-%m-%d")
    before_date     = len(df)
    df              = df[df["sale_date"] <= today]
    if before_date != len(df):
        print(f"Removed {before_date - len(df)} future-dated transactions")

    print(f"Transactions: {before} → {len(df)} rows after cleaning")
    print(f"Date range:   {df['sale_date'].min()} to {df['sale_date'].max()}")
    print(f"Total revenue: ${df['total_revenue'].sum():,.2f}")
    print(f"Total profit:  ${df['profit'].sum():,.2f}")
    print(f"Avg margin:    {df['margin_pct'].mean():.1f}%")
    return df


def transform(data):
    print("\n" + "="*55)
    print("STEP 2 — TRANSFORM")
    print("="*55)

    clean = {}
    clean["stores"]       = transform_stores(data["stores"])
    clean["products"]     = transform_products(data["products"])
    clean["staff"]        = transform_staff(data["staff"])
    clean["customers"]    = transform_customers(data["customers"])
    clean["transactions"] = transform_transactions(data["transactions"])

    return clean


# ══════════════════════════════════════════════════════
# LOAD — Insert clean data into MySQL
# ══════════════════════════════════════════════════════

def load_to_mysql(df, table_name, id_col):
    conn   = get_connection()
    cursor = conn.cursor()

    cursor.execute("SET FOREIGN_KEY_CHECKS = 0")
    cursor.execute(f"TRUNCATE TABLE {table_name}")
    cursor.execute("SET FOREIGN_KEY_CHECKS = 1")

    # Only use columns that exist in MySQL table
    mysql_columns = {
        "stores":       ["store_id","store_name","suburb","region","manager","target_monthly"],
        "products":     ["product_id","name","category","cost","price","margin_pct"],
        "staff":        ["staff_id","name","store_id","role","salary","start_date"],
        "customers":    ["customer_id","first_name","last_name","email","phone",
                         "suburb","age_group","segment","loyalty_member","join_date"],
        "transactions": ["transaction_id","sale_date","store_id","customer_id",
                         "product_id","staff_id","quantity","unit_price","unit_cost",
                         "total_revenue","total_cost","profit","margin_pct",
                         "payment_method","month","day_of_week","quarter","year"]
    }

    cols         = mysql_columns[table_name]
    df_filtered  = df[cols].copy()
    placeholders = ", ".join(["%s"] * len(cols))
    col_str      = ", ".join(cols)

    inserted = 0
    skipped  = 0

    for _, row in df_filtered.iterrows():
        try:
            values = tuple(
                None if pd.isna(v) else
                int(v) if isinstance(v, (np.integer,)) else
                float(v) if isinstance(v, (np.floating,)) else
                bool(v) if isinstance(v, (np.bool_,)) else
                v
                for v in row
            )
            cursor.execute(
                f"INSERT INTO {table_name} ({col_str}) VALUES ({placeholders})",
                values
            )
            inserted += 1
        except Exception as e:
            skipped += 1

    conn.commit()
    cursor.close()
    conn.close()
    print(f"Loaded {inserted} rows into {table_name} (skipped {skipped})")


def load(clean_data):
    print("\n" + "="*55)
    print("STEP 3 — LOAD")
    print("="*55)

    # Load in order — parent tables before child tables
    load_to_mysql(clean_data["stores"],       "stores",       "store_id")
    load_to_mysql(clean_data["products"],     "products",     "product_id")
    load_to_mysql(clean_data["staff"],        "staff",        "staff_id")
    load_to_mysql(clean_data["customers"],    "customers",    "customer_id")
    load_to_mysql(clean_data["transactions"], "transactions", "transaction_id")


# ══════════════════════════════════════════════════════
# VERIFY — Confirm everything loaded correctly
# ══════════════════════════════════════════════════════

def verify():
    print("\n" + "="*55)
    print("STEP 4 — VERIFY")
    print("="*55)

    conn   = get_connection()
    cursor = conn.cursor()

    tables = ["stores", "products", "staff", "customers", "transactions"]
    for table in tables:
        cursor.execute(f"SELECT COUNT(*) FROM {table}")
        print(f"  {table:<15} {cursor.fetchone()[0]:>6} rows in MySQL")

    print("\nQuick revenue check:")
    cursor.execute("""
        SELECT
            COUNT(*)                      AS total_transactions,
            ROUND(SUM(total_revenue), 2)  AS total_revenue,
            ROUND(SUM(profit), 2)         AS total_profit,
            ROUND(AVG(margin_pct), 1)     AS avg_margin_pct
        FROM transactions
    """)
    row = cursor.fetchone()
    print(f"  Transactions: {row[0]:,}")
    print(f"  Total revenue: ${row[1]:,.2f}")
    print(f"  Total profit:  ${row[2]:,.2f}")
    print(f"  Avg margin:    {row[3]}%")

    cursor.close()
    conn.close()
    print("="*55)


# ══════════════════════════════════════════════════════
# RUN FULL PIPELINE
# ══════════════════════════════════════════════════════

if __name__ == "__main__":
    print("="*55)
    print("AusMart Sydney — ETL Pipeline")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*55)

    raw_data   = extract()
    clean_data = transform(raw_data)
    load(clean_data)
    verify()

    print(f"\nCompleted: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("ETL Pipeline finished successfully!")