# Build Guide: Sales & Profit Dashboard (about 45 minutes)

## Data headline (after cleaning)
- 14 test rows removed (Sales 100,000 / 1,000,000). Sales drops from 3.58M to 1.28M. Profit from -667K to 632K.
- Clean totals: Sales 1,278,119 | Cost 646,055 | Profit 632,064 | Margin 49.5% | Qty 15,604 | Orders 4,166.
- Yearly Sales: 2017 229K, 2018 290K, 2019 345K, 2020 414K. Steady growth of about 20%/yr.
- Use these numbers to check your cards. If they match, the model is right.

## Steps
1. Open Power BI Desktop > Transform data > paste `1_PowerQuery_Orders.m` (edit file path). Close & Apply.
2. File > Options > Current file > Data load: turn OFF "Auto date/time".
3. Add `Date` table and measures from `2_DAX_Measures.dax`. Mark Date as date table. Relate Date[Date] to Orders[Order Date].
4. Add hierarchies: Orders[Region > Country > State > City] and Orders[Category > Sub-Category > Product Name].
5. View > Themes > Browse for themes > `3_Theme.json`.
6. Canvas 1280x720 (Format page > Canvas settings). Build the pages below.

## Page 1: Executive Overview
| Zone | Visual | Fields |
|---|---|---|
| Top-left filters (row) | Slicers (dropdown): Year, Region, Segment, Category | Date[Year], Orders[...] |
| KPI row (5 cards) | Card + YoY text | Total Sales, Total Profit, Profit Margin %, Orders Count, Total Quantity; subtitle = YoY Arrow |
| Middle-left (wide) | Line + clustered column | X: Date[Year-Month]; Y: Selected Metric; title = Selected Title |
| Middle-right | Donut | Category, Total Sales |
| Bottom-left | Clustered bar | Region, Total Sales and Total Profit |
| Bottom-right | Bar, Top N filter = 5 | Sub-Category, Total Profit |

## Page 2: Geography
Map (Country, size = Total Sales, colour = Profit Margin %), bar chart of Country by Sales, matrix with Region > Country > State > City showing Sales, Profit, Margin %.

## Page 3: Products
Treemap (Category > Sub-Category, Sales), scatter (Sub-Category; X = Total Sales, Y = Profit Margin %, size = Quantity), table of Top 10 products (Product Rank <= 10) with Sales, Profit.

## Page 4: Customers & Shipping
Donut by Segment, bar of Top 10 Customer Name by Sales, column chart by Ship Mode (Sales, Avg Order Value), line of Customers by Year-Month.

## Page 5: Profitability
Waterfall (Sales, Cost, Profit), matrix Category x Region with Profit Margin % (conditional colour using Profit Color), table of loss-making orders (filter Profit < 0).

## Interactivity
- Sync Year, Region slicers across all pages (View > Sync slicers).
- Drill-through page on Country and Sub-Category.
- Add a Bookmark + button "Reset filters"; page navigator (Insert > Buttons > Navigator > Page navigator).
- Tooltip page: card of Sales, Profit, Margin %, mini line by Year-Month.

## Performance
Only 4,166 rows, so speed is not an issue. Keep the Auto date/time off, use measures instead of calculated columns, and keep to 8 visuals per page or fewer.

## Optional RLS
Modeling > Manage roles > role "North": `[Region] = "North"` on Orders (repeat for Central, South). Test with View as.

## Assumptions
Currency = EUR. Calendar year = fiscal year. The 14 test rows are junk, not real sales. Profit = Sales - Cost.
