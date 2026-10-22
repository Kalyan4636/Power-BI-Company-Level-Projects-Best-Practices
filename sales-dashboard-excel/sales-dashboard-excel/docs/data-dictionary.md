# Data dictionary

## Source columns (`data/sales_data.csv`)
| Column | Type | Notes |
|---|---|---|
| ID | text | 5-digit order id with leading zeros; unique |
| Order Date | date (ISO) | 2026-01-01 to 2026-12-30 |
| Region | text | Central / Northern / Southern / Western Europe |
| Country | text | 17 countries |
| Sales Channel | text | Corporate, Marketplace, Online Store, Retail Partners |
| Customer Type | text | Business, New, Returning |
| Product Category | text | Camping, Cycling, Fitness, Hiking, Travel |
| Product | text | 20 products, each in one category |
| Units Sold | int | |
| Unit Price | float | list price |
| Discount % | float | 0, 0.05, 0.10, 0.15, 0.20 |
| Unit Cost | float | |

## Derived columns (formulas on `Data`)
| Column | Formula |
|---|---|
| Revenue | `Units x Unit Price x (1 - Discount)` |
| COGS | `Units x Unit Cost` |
| Gross Profit | `Revenue - COGS` |
| Gross Margin % | `Gross Profit / Revenue` |
| Year / Quarter / MonthNo / Month | from Order Date |
| In Filter, Flag excl. * | see `architecture.md` |

## KPI definitions
| KPI | Definition |
|---|---|
| Total Revenue | sum of Revenue over filtered rows |
| Gross Profit / Margin | sum of GP; GP / Revenue |
| Units Sold | sum of Units |
| Orders | count of filtered rows (one row = one order) |
| Avg Order Value | Revenue / Orders |
| Avg Unit Price (net) | Revenue / Units, i.e. after discount |

Full-year reference values: revenue 645,223.18; gross profit 296,384.25; units 6,698; orders 2,707.

## Data quality checks performed
No blank cells, no duplicate IDs, one calendar year, every product maps to exactly one category, no row with unit price <= unit cost.
