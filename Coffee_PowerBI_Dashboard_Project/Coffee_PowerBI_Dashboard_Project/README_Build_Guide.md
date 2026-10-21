# Coffee Trends & Spend Analysis — Power BI Project Pack

## Dataset profile
- Rows: 3,547
- Columns after cleaning: 11
- Date range: 2024-03-01 to 2025-03-23
- Total spend: ₹112,245.58
- Average spend per transaction: ₹31.65
- Distinct coffee types: 8
- Exact duplicate rows in source: 0
- Missing values in source: 0

## Included files
- `Coffee_Cleaned.csv` — cleaned data with `money(Rs)` renamed to `Spend`; dates normalized.
- `Power_Query_M.txt` — Power Query M starter query. Update the file path before use.
- `DAX_Measures.txt` — date table, relationships guidance, base measures and time-intelligence measures.
- `Coffee_Insights_Theme.json` — importable Power BI theme.

## Build order
1. **Power Query cleaning:** Home → Transform data. Use the M query as a starter, update the CSV path, and verify the preview.
2. **Date table:** Create the `Calendar` table from the DAX file.
3. **Relationship:** `Calendar[Date]` one-to-many to `Coffee[Date]`; mark Calendar as the date table.
4. **Base DAX:** Add measures from `DAX_Measures.txt`.
5. **Time intelligence:** Add previous-month, MoM and YoY measures. Check the date range before interpreting partial periods.
6. **Visuals:** Monthly spend line chart; weekday spend columns; hourly spend and transaction combo chart; time-of-day distribution; weekday vs weekend comparison; coffee type ranking.
7. **Coffee-type tile slicer:** Add `Coffee[coffee_name]` to a slicer, select Tile/button style if available, then verify Edit interactions.
8. **Formatting and validation:** View → Themes → Browse for themes → choose the JSON file. Confirm all totals reconcile and slicers filter all intended visuals.

## Recommended report layout
- Canvas: 16:9.
- Top: title, date slicer, coffee-type tile slicer.
- KPI row: Total Spend, Total Transactions, Average Spend, Total Coffee Types.
- Middle: monthly trend and weekday spending.
- Bottom: spend by hour, time-of-day spend, weekday/weekend comparison.

## Important assumptions and cautions
- `Total Cups` equals transaction rows because the source has no quantity field. If one row does not equal one cup, add quantity data before describing this as actual cups sold.
- The source `Time` values appear as strings such as `15:50.5`; use `hour_of_day` for hourly analysis unless the original time encoding is confirmed.
- This pack contains the cleaned dataset, theme, Power Query, and DAX starter assets. A native `.pbix` report must be assembled and saved in Power BI Desktop; this environment cannot generate a genuine Power BI Desktop file directly.
- Review the Power Query path and field names against your local model before applying.
