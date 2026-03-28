# AusMart Sydney Sales Intelligence Platform

## Project Overview
End-to-end business intelligence platform analysing retail sales across 5 Sydney stores using Python, MySQL, Excel and Power BI.

## Tech Stack
- **Python** — data generation, ETL pipeline, SQL analysis, Excel reporting
- **MySQL** — relational database with 5 normalised tables
- **pandas** — data cleaning and transformation
- **openpyxl / xlsxwriter** — automated Excel report generation
- **Power BI** — 4-page interactive executive dashboard
- **Git** — version control

## Project Architecture
```
Excel Intake Form (input)
        ↓
Python ETL Pipeline (extract → transform → load)
        ↓
MySQL Database (aumart_sydney)
        ↓
Python SQL Analysis (12 advanced queries)
        ↓
Excel Insights Report (automated output)
        ↓
Power BI Dashboard (4 interactive pages)
```

## Key Business Insights
- **AusMart CBD** leads with $289,368 total revenue
- **Samsung 55 TV** is the top selling product
- **Saturday** is peak trading day across all stores
- **Loyalty members** spend 2x more than non-members
- **Castle Hill** is the top suburb by customer revenue
- **Grocery** dominates at 60% of total revenue

## Database Schema
- `stores` — 5 Sydney store locations
- `products` — 50 products across 5 categories
- `staff` — 20 staff members across stores
- `customers` — 1,000 Sydney customers
- `transactions` — 7,602 sales transactions (2024-2025)

## Files
| File | Purpose |
|---|---|
| `generate_data.py` | Generates realistic Sydney retail data |
| `database.py` | MySQL schema creation and data loading |
| `etl_pipeline.py` | Full ETL pipeline — Extract, Transform, Load |
| `analysis.py` | 12 advanced SQL queries with window functions |
| `excel_intake.py` | Professional Excel intake form generator |
| `report_generator.py` | Automated 6-sheet Excel insights report |
| `config.py` | Database credentials (excluded from GitHub) |

## Dashboard Pages
1. **Executive Overview** — KPI cards, revenue by store, monthly trend
2. **Store Performance** — revenue vs target, region analysis, staff ranking
3. **Customer Insights** — segments, loyalty analysis, top suburbs
4. **Product Analysis** — top 10 products, margins, payment methods

## How to Run
```bash
# Install dependencies
pip install pandas numpy mysql-connector-python openpyxl faker streamlit plotly xlsxwriter

# Generate data
python generate_data.py

# Set up MySQL database
python database.py

# Run ETL pipeline
python etl_pipeline.py

# Run SQL analysis
python analysis.py

# Generate Excel report
python report_generator.py
```

## Requirements
- Python 3.8+
- MySQL 8.0+
- Power BI Desktop (free)

## Author
Jeevanandh Mathivanan — Data Analyst
Sydney, Australia