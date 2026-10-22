# Building the native (PivotTable + slicer) version in Excel

Use this when you need real PivotTables, slicers, a timeline and a filled map, which this repo's generator cannot create.

| Step | Actions |
|---|---|
| 1 Data | Select data, `Ctrl+T`, name `SalesDataTable`. Add Revenue, Gross Profit, Month columns. |
| 2 Pivots | New sheet `Analyze`. Insert > PivotTable x8: Sum Revenue; Sum Gross Profit; Sum Units; Avg Revenue; Month x (Revenue, GP); Country x Revenue; Category x Revenue; Product x Units (Value Filter > Top 10 > 5). |
| 3 KPIs | Rounded-rectangle shape, select it, type `=Analyze!D3` in the formula bar (cell holds `GETPIVOTDATA`). |
| 4 Charts | Click a pivot > PivotTable Analyze > PivotChart; hide field buttons; move onto the Dashboard sheet. |
| 5 Map | Click Country pivot > Insert > Maps > Filled Map. |
| 6 Slicers | Insert > Slicer (Region, Sales Channel, Customer Type). |
| 7 Timeline | Insert > Timeline > Order Date > Months. |
| 8 Link everything | For each slicer/timeline: Report Connections > tick **every** PivotTable. Test by selecting a value and confirming all visuals move. |
| 9 Navigation | Shapes with Right-click > Link > Place in This Document. |
| 10 Finish | Hide gridlines; shape properties "Don't move or size with cells"; Data > Refresh All. |
