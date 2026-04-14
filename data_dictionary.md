# Data Dictionary — AusMart Sydney Retail BI Platform

## fact_sales

| Column | Data Type | Description |
|--------|-----------|-------------|
| transaction_id | INT | Unique transaction identifier |
| store_id | INT | Foreign key to dim_store |
| product_id | INT | Foreign key to dim_product |
| customer_id | INT | Foreign key to dim_customer |
| date_id | INT | Foreign key to dim_date |
| quantity | INT | Units purchased |
| unit_price | FLOAT | Price per unit at time of sale |
| total_amount | FLOAT | quantity x unit_price |

## dim_store

| Column | Data Type | Description |
|--------|-----------|-------------|
| store_id | INT | Unique store identifier |
| store_name | VARCHAR | Store name (e.g. AusMart Bondi) |
| region | VARCHAR | Sydney region (CBD, Western, etc.) |
| suburb | VARCHAR | Store suburb |

## dim_product

| Column | Data Type | Description |
|--------|-----------|-------------|
| product_id | INT | Unique product identifier |
| name | VARCHAR | Product name |
| category | VARCHAR | Product category |
| unit_price | FLOAT | Standard selling price |

## dim_date

| Column | Data Type | Description |
|--------|-----------|-------------|
| date_id | INT | Unique date identifier |
| date | DATE | Full date |
| month | VARCHAR | Month name |
| quarter | INT | Quarter number |
| year | INT | Year |
| day_of_week | VARCHAR | Day name |