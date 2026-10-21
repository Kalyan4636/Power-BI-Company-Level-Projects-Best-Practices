# DDC Coffee Sales - Power BI build guide

## 1. Data model (star schema)
FactSales (Date, StoreKey, ProductKey, ChannelKey, Orders, Revenue) at month x store x product x channel grain.
Relationships, all 1:* single direction: DimDate[Date]->FactSales[Date], DimStore, DimProduct, DimChannel on their keys.
Mark DimDate as the date table (Table tools > Mark as date table). Hide key columns and FactSales numeric columns; use measures only.

## 2. Steps
1. Load the 5 CSVs (see PowerQuery_M.txt), set types, Close & Apply.
2. Create the relationships above. Sort DimDate[MonthName] by MonthNum; sort Quarter ascending.
3. View > Themes > Browse for themes: ddc-theme.json. Page size: Format page > Canvas settings > 16:9 (1280x720).
4. Paste DAX_measures.dax into a _Measures table.
5. Slicers: DimDate[Year] (Tile, single select), DimDate[Quarter], DimStore[Store], DimChannel[Channel], DimProduct[Category] (Tile, multi-select with Ctrl). Turn off the slicer header, set item corner radius, fill amber when selected.
6. Visuals: 4 Card (new) KPIs; Area/Line (DimDate[MonthName] x [Revenue], [Revenue LY] dashed); Bar (DimProduct[Product] x [Revenue], colour by [Bar Colour] via fx); Donut (DimChannel[Channel] x [Revenue]); Matrix heatmap (Weekday x Hour, background colour rules) - needs hourly grain data.
7. View > Sync slicers if you add more pages. Format > Edit interactions: set slicers to Filter and visuals to Cross-highlight.
8. Bookmark "Reset" (all slicers cleared) and attach it to a button. Add alt text to every visual.
9. QA: Performance analyzer (all visuals < 120 ms), compare [Revenue] total with a SUM in SQL/Excel, test an empty selection.
10. Publish, then optional RLS: Manage roles > Store manager: [Store] = USERPRINCIPALNAME() mapping table.
