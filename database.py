import mysql.connector
import pandas as pd
from config import DB_CONFIG


def get_connection():
    return mysql.connector.connect(**DB_CONFIG)


def create_tables():
    conn   = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS stores (
            store_id         INT PRIMARY KEY,
            store_name       VARCHAR(100) NOT NULL,
            suburb           VARCHAR(100),
            region           VARCHAR(50),
            manager          VARCHAR(100),
            target_monthly   DECIMAL(10,2)
        )
    """)
    print("Table stores ready")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS products (
            product_id  INT PRIMARY KEY,
            name        VARCHAR(200) NOT NULL,
            category    VARCHAR(50),
            cost        DECIMAL(10,2),
            price       DECIMAL(10,2),
            margin_pct  DECIMAL(5,3)
        )
    """)
    print("Table products ready")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS staff (
            staff_id    INT PRIMARY KEY,
            name        VARCHAR(100) NOT NULL,
            store_id    INT,
            role        VARCHAR(50),
            salary      DECIMAL(10,2),
            start_date  DATE,
            FOREIGN KEY (store_id) REFERENCES stores(store_id)
        )
    """)
    print("Table staff ready")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS customers (
            customer_id    INT PRIMARY KEY,
            first_name     VARCHAR(50),
            last_name      VARCHAR(50),
            email          VARCHAR(100),
            phone          VARCHAR(30),
            suburb         VARCHAR(100),
            age_group      VARCHAR(10),
            segment        VARCHAR(20),
            loyalty_member BOOLEAN,
            join_date      DATE
        )
    """)
    print("Table customers ready")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS transactions (
            transaction_id  INT PRIMARY KEY,
            sale_date       DATE NOT NULL,
            store_id        INT,
            customer_id     INT,
            product_id      INT,
            staff_id        INT,
            quantity        INT,
            unit_price      DECIMAL(10,2),
            unit_cost       DECIMAL(10,2),
            total_revenue   DECIMAL(10,2),
            total_cost      DECIMAL(10,2),
            profit          DECIMAL(10,2),
            margin_pct      DECIMAL(5,1),
            payment_method  VARCHAR(20),
            month           VARCHAR(7),
            day_of_week     VARCHAR(10),
            quarter         VARCHAR(2),
            year            INT,
            FOREIGN KEY (store_id)   REFERENCES stores(store_id),
            FOREIGN KEY (customer_id) REFERENCES customers(customer_id),
            FOREIGN KEY (product_id) REFERENCES products(product_id),
            FOREIGN KEY (staff_id)   REFERENCES staff(staff_id)
        )
    """)
    print("Table transactions ready")

    conn.commit()
    cursor.close()
    conn.close()
    print("\nAll 5 tables created in MySQL")


def load_table(filename, table_name, transform_fn=None):
    print(f"\nLoading {table_name}...")
    df = pd.read_csv(f"data/{filename}")

    if transform_fn:
        df = transform_fn(df)

    conn   = get_connection()
    cursor = conn.cursor()

    # Disable foreign key checks so TRUNCATE works
    cursor.execute("SET FOREIGN_KEY_CHECKS = 0")
    cursor.execute(f"TRUNCATE TABLE {table_name}")
    cursor.execute("SET FOREIGN_KEY_CHECKS = 1")

    cols         = ", ".join(df.columns)
    placeholders = ", ".join(["%s"] * len(df.columns))

    inserted = 0
    for _, row in df.iterrows():
        values = tuple(
            None if pd.isna(v) else bool(v) if isinstance(v, bool) else v
            for v in row
        )
        cursor.execute(
            f"INSERT INTO {table_name} ({cols}) VALUES ({placeholders})",
            values
        )
        inserted += 1

    conn.commit()
    cursor.close()
    conn.close()
    print(f"Loaded {inserted} rows into {table_name}")

def fix_staff(df):
    df["start_date"] = pd.to_datetime(df["start_date"]).dt.strftime("%Y-%m-%d")
    return df


def fix_customers(df):
    df["join_date"]      = pd.to_datetime(df["join_date"]).dt.strftime("%Y-%m-%d")
    df["loyalty_member"] = df["loyalty_member"].map({True: 1, False: 0, "True": 1, "False": 0})
    return df


def verify_database():
    conn   = get_connection()
    cursor = conn.cursor()

    print("\n" + "="*50)
    print("DATABASE VERIFICATION")
    print("="*50)

    tables = ["stores", "products", "staff", "customers", "transactions"]
    for table in tables:
        cursor.execute(f"SELECT COUNT(*) FROM {table}")
        count = cursor.fetchone()[0]
        print(f"  {table:<15} {count:>6} rows")

    print("\nRevenue summary by store:")
    cursor.execute("""
        SELECT s.store_name,
               COUNT(t.transaction_id)    AS transactions,
               ROUND(SUM(t.total_revenue),2) AS total_revenue,
               ROUND(SUM(t.profit),2)        AS total_profit,
               ROUND(AVG(t.total_revenue),2) AS avg_transaction
        FROM transactions t
        JOIN stores s ON t.store_id = s.store_id
        GROUP BY s.store_name
        ORDER BY total_revenue DESC
    """)
    rows = cursor.fetchall()
    print(f"  {'Store':<25} {'Txns':>6} {'Revenue':>12} {'Profit':>10} {'Avg Txn':>10}")
    print("  " + "-"*65)
    for row in rows:
        print(f"  {row[0]:<25} {row[1]:>6} ${row[2]:>10,.2f} ${row[3]:>8,.2f} ${row[4]:>8,.2f}")

    cursor.close()
    conn.close()
    print("="*50)


def get_stores():
    conn = get_connection()
    df   = pd.read_sql("SELECT * FROM stores", conn)
    conn.close()
    return df


def get_products():
    conn = get_connection()
    df   = pd.read_sql("SELECT * FROM products", conn)
    conn.close()
    return df


def get_transactions():
    conn = get_connection()
    df   = pd.read_sql("SELECT * FROM transactions", conn)
    conn.close()
    return df


if __name__ == "__main__":
    print("="*50)
    print("AusMart Sydney — Database Setup")
    print("="*50)

    create_tables()

    # Load in correct order — foreign keys require parent tables first
    load_table("stores.csv",       "stores")
    load_table("products.csv",     "products")
    load_table("staff.csv",        "staff",     fix_staff)
    load_table("customers.csv",    "customers", fix_customers)
    load_table("transactions.csv", "transactions")

    verify_database()
    print("\nMySQL database is ready!")