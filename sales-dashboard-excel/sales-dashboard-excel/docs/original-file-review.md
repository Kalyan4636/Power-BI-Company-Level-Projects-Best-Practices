# Review of the original workbook (`Interactive_Excel_Sales_Dashboard.xlsx`)

The original uses native PivotTables, slicers, a timeline and a map chart. Findings from inspecting its XML:

**Working correctly**
- All four KPI cards are linked to cells (`Analyze!D3/D6/D9/D12`) holding `GETPIVOTDATA` formulas.
- Region / Sales Channel / Customer Type slicers and the Order Date timeline are connected to all 8 pivot tables.
- Revenue total reproduces exactly from Units x Price x (1 - Discount).

**Recommended fixes**
1. "Avg. Price" shows 238.35, which is *average revenue per order*. Net price per unit is 645,223 / 6,698 = **96.33**.
2. Contact button targets the placeholder `mailto:hello@gmail.com`.
3. Indian digit grouping (6,45,223) on a dashboard described as USD; use `$#,##0`.
4. Hard-coded "2026" title; derive it from the data.
5. Confirm a back-to-dashboard button exists on `Analyze` and `Data`.
6. Label the Top-5 chart as top 5 products (its grand total is the top-5 sum, not all units).
