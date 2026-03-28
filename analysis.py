import pandas as pd
import mysql.connector
from config import DB_CONFIG


def get_connection():
    return mysql.connector.connect(**DB_CONFIG)


def run_query(sql, description):
    conn = get_connection()
    df   = pd.read_sql(sql, conn)
    conn.close()
    print(f"\n{'='*55}")
    print(f"{description}")
    print('='*55)
    print(df.to_string(index=False))
    return df


# ── Query 1: Revenue and profit by store ─────────────
def revenue_by_store():
    return run_query("""
        SELECT
            s.store_name,
            s.region,
            s.target_monthly,
            COUNT(t.transaction_id)          AS total_transactions,
            ROUND(SUM(t.total_revenue), 2)   AS total_revenue,
            ROUND(SUM(t.profit), 2)          AS total_profit,
            ROUND(AVG(t.total_revenue), 2)   AS avg_transaction_value,
            ROUND(AVG(t.margin_pct), 1)      AS avg_margin_pct,
            ROUND(SUM(t.total_revenue) / s.target_monthly / 24 * 100, 1) AS target_achieved_pct
        FROM transactions t
        JOIN stores s ON t.store_id = s.store_id
        GROUP BY s.store_id, s.store_name, s.region, s.target_monthly
        ORDER BY total_revenue DESC
    """, "Query 1 — Revenue and profit by store")


# ── Query 2: Top 10 products by revenue ──────────────
def top_products():
    return run_query("""
        SELECT
            p.name,
            p.category,
            p.price                           AS unit_price,
            COUNT(t.transaction_id)           AS times_sold,
            SUM(t.quantity)                   AS units_sold,
            ROUND(SUM(t.total_revenue), 2)    AS total_revenue,
            ROUND(SUM(t.profit), 2)           AS total_profit,
            ROUND(AVG(t.margin_pct), 1)       AS avg_margin_pct
        FROM transactions t
        JOIN products p ON t.product_id = p.product_id
        GROUP BY p.product_id, p.name, p.category, p.price
        ORDER BY total_revenue DESC
        LIMIT 10
    """, "Query 2 — Top 10 products by revenue")


# ── Query 3: Monthly revenue trend ───────────────────
def monthly_trend():
    return run_query("""
        SELECT
            month,
            year,
            COUNT(transaction_id)            AS transactions,
            ROUND(SUM(total_revenue), 2)     AS monthly_revenue,
            ROUND(SUM(profit), 2)            AS monthly_profit,
            ROUND(AVG(total_revenue), 2)     AS avg_transaction_value,
            ROUND(
                (SUM(total_revenue) - LAG(SUM(total_revenue))
                    OVER (ORDER BY year, month))
                / LAG(SUM(total_revenue))
                    OVER (ORDER BY year, month) * 100
            , 1) AS mom_growth_pct
        FROM transactions
        GROUP BY month, year
        ORDER BY year, month
    """, "Query 3 — Monthly revenue trend with MoM growth")


# ── Query 4: Customer segment analysis ───────────────
def customer_segments():
    return run_query("""
        SELECT
            c.segment,
            COUNT(DISTINCT t.customer_id)    AS unique_customers,
            COUNT(t.transaction_id)          AS total_transactions,
            ROUND(SUM(t.total_revenue), 2)   AS total_revenue,
            ROUND(AVG(t.total_revenue), 2)   AS avg_spend_per_transaction,
            ROUND(SUM(t.total_revenue) /
                COUNT(DISTINCT t.customer_id), 2) AS avg_revenue_per_customer,
            ROUND(AVG(t.margin_pct), 1)      AS avg_margin_pct
        FROM transactions t
        JOIN customers c ON t.customer_id = c.customer_id
        GROUP BY c.segment
        ORDER BY total_revenue DESC
    """, "Query 4 — Customer segment analysis")


# ── Query 5: Loyalty member vs non-member ────────────
def loyalty_analysis():
    return run_query("""
        SELECT
            CASE WHEN c.loyalty_member = 1
                 THEN 'Loyalty Member'
                 ELSE 'Non Member'
            END                              AS membership_status,
            COUNT(DISTINCT t.customer_id)   AS customers,
            COUNT(t.transaction_id)         AS transactions,
            ROUND(SUM(t.total_revenue), 2)  AS total_revenue,
            ROUND(AVG(t.total_revenue), 2)  AS avg_spend_per_visit,
            ROUND(SUM(t.total_revenue) /
                COUNT(DISTINCT t.customer_id), 2) AS revenue_per_customer
        FROM transactions t
        JOIN customers c ON t.customer_id = c.customer_id
        GROUP BY c.loyalty_member
        ORDER BY total_revenue DESC
    """, "Query 5 — Loyalty member vs non-member spend")


# ── Query 6: Sales by day of week ────────────────────
def day_of_week_analysis():
    return run_query("""
        SELECT
            day_of_week,
            COUNT(transaction_id)            AS transactions,
            ROUND(SUM(total_revenue), 2)     AS total_revenue,
            ROUND(AVG(total_revenue), 2)     AS avg_transaction_value,
            ROUND(SUM(profit), 2)            AS total_profit
        FROM transactions
        GROUP BY day_of_week
        ORDER BY FIELD(day_of_week,
            'Monday','Tuesday','Wednesday',
            'Thursday','Friday','Saturday','Sunday')
    """, "Query 6 — Sales by day of week")


# ── Query 7: Category performance ────────────────────
def category_performance():
    return run_query("""
        SELECT
            p.category,
            COUNT(t.transaction_id)          AS transactions,
            SUM(t.quantity)                  AS units_sold,
            ROUND(SUM(t.total_revenue), 2)   AS total_revenue,
            ROUND(SUM(t.profit), 2)          AS total_profit,
            ROUND(AVG(t.margin_pct), 1)      AS avg_margin_pct,
            ROUND(SUM(t.total_revenue) * 100.0 /
                (SELECT SUM(total_revenue) FROM transactions), 1
            )                                AS revenue_share_pct
        FROM transactions t
        JOIN products p ON t.product_id = p.product_id
        GROUP BY p.category
        ORDER BY total_revenue DESC
    """, "Query 7 — Category performance and revenue share")


# ── Query 8: Staff performance ───────────────────────
def staff_performance():
    return run_query("""
        SELECT
            st.name                          AS staff_name,
            st.role,
            s.store_name,
            COUNT(t.transaction_id)          AS transactions_handled,
            ROUND(SUM(t.total_revenue), 2)   AS revenue_generated,
            ROUND(AVG(t.total_revenue), 2)   AS avg_transaction_value,
            ROUND(SUM(t.profit), 2)          AS profit_generated
        FROM transactions t
        JOIN staff st   ON t.staff_id  = st.staff_id
        JOIN stores s   ON t.store_id  = s.store_id
        GROUP BY st.staff_id, st.name, st.role, s.store_name
        ORDER BY revenue_generated DESC
    """, "Query 8 — Staff performance ranking")


# ── Query 9: Payment method breakdown ────────────────
def payment_methods():
    return run_query("""
        SELECT
            payment_method,
            COUNT(transaction_id)            AS transactions,
            ROUND(SUM(total_revenue), 2)     AS total_revenue,
            ROUND(AVG(total_revenue), 2)     AS avg_transaction_value,
            ROUND(COUNT(transaction_id) * 100.0 /
                (SELECT COUNT(*) FROM transactions), 1
            )                                AS usage_pct
        FROM transactions
        GROUP BY payment_method
        ORDER BY transactions DESC
    """, "Query 9 — Payment method breakdown")


# ── Query 10: Quarterly performance ──────────────────
def quarterly_performance():
    return run_query("""
        SELECT
            year,
            quarter,
            COUNT(transaction_id)            AS transactions,
            ROUND(SUM(total_revenue), 2)     AS quarterly_revenue,
            ROUND(SUM(profit), 2)            AS quarterly_profit,
            ROUND(AVG(margin_pct), 1)        AS avg_margin_pct,
            ROUND(
                (SUM(total_revenue) - LAG(SUM(total_revenue))
                    OVER (ORDER BY year, quarter))
                / LAG(SUM(total_revenue))
                    OVER (ORDER BY year, quarter) * 100
            , 1)                             AS qoq_growth_pct
        FROM transactions
        GROUP BY year, quarter
        ORDER BY year, quarter
    """, "Query 10 — Quarterly performance with QoQ growth")


# ── Query 11: Top suburbs by customer revenue ────────
def top_suburbs():
    return run_query("""
        SELECT
            c.suburb,
            COUNT(DISTINCT t.customer_id)    AS unique_customers,
            COUNT(t.transaction_id)          AS transactions,
            ROUND(SUM(t.total_revenue), 2)   AS total_revenue,
            ROUND(AVG(t.total_revenue), 2)   AS avg_spend
        FROM transactions t
        JOIN customers c ON t.customer_id = c.customer_id
        GROUP BY c.suburb
        ORDER BY total_revenue DESC
        LIMIT 10
    """, "Query 11 — Top 10 Sydney suburbs by revenue")


# ── Query 12: Advanced — running total by store ──────
def running_total_by_store():
    return run_query("""
        SELECT
            month,
            store_id,
            ROUND(SUM(total_revenue), 2)     AS monthly_revenue,
            ROUND(
                SUM(SUM(total_revenue))
                OVER (PARTITION BY store_id ORDER BY month)
            , 2)                             AS running_total
        FROM transactions
        GROUP BY month, store_id
        ORDER BY store_id, month
        LIMIT 30
    """, "Query 12 — Running total revenue by store (window function)")


def run_all():
    print("="*55)
    print("AusMart Sydney — SQL Analysis")
    print("="*55)

    results = {
        "store_revenue":     revenue_by_store(),
        "top_products":      top_products(),
        "monthly_trend":     monthly_trend(),
        "segments":          customer_segments(),
        "loyalty":           loyalty_analysis(),
        "day_of_week":       day_of_week_analysis(),
        "categories":        category_performance(),
        "staff":             staff_performance(),
        "payments":          payment_methods(),
        "quarterly":         quarterly_performance(),
        "suburbs":           top_suburbs(),
        "running_total":     running_total_by_store(),
    }

    print("\n" + "="*55)
    print("All 12 SQL queries completed successfully!")
    print("Key findings:")
    print(f"  Top store:     {results['store_revenue'].iloc[0]['store_name']}")
    print(f"  Top product:   {results['top_products'].iloc[0]['name']}")
    print(f"  Top segment:   {results['segments'].iloc[0]['segment']}")
    print(f"  Best day:      {results['day_of_week'].sort_values('total_revenue', ascending=False).iloc[0]['day_of_week']}")
    print(f"  Top suburb:    {results['suburbs'].iloc[0]['suburb']}")
    print("="*55)

    return results


if __name__ == "__main__":
    run_all()