# Click-to-filter macro

Turns a click on any chart bar, month column or map tile into a filter, so **every** visual updates. Click again to clear; click **ORDERS SHOWN** to reset.

Desktop Excel only (Windows/Mac). Not supported in Excel for the web or mobile.

## Install (2 minutes)
1. Open `dashboard/Sales_Performance_Dashboard.xlsx` and **Save As** `.xlsm`.
2. `Alt+F11` > **File > Import File...** and import `clsChart.cls` and `modCrossFilter.bas`.
3. Double-click **ThisWorkbook** and paste `ThisWorkbook.vba`.
4. Double-click the **Dashboard** sheet and paste `Sheet_Dashboard.vba`.
5. Save, reopen, click **Enable Content**.

## How it maps clicks
| Chart (top-left cell) | Sets filter |
|---|---|
| Monthly (B17) | From + To month |
| Category (L17) | Category |
| Top 5 products (L36) | Product |
| Top 10 countries (B55) | Country |
| Sales channel (L56) | Channel |
| Customer type (P56) | Customer type |
| Map tiles B37:K46 | Country |

Do not rename the `Dashboard` sheet or move charts to other anchor cells.

## Status
Written against the documented Excel object model; it could not be executed in the project's CI (no Excel). Please open an issue with your Excel version if something misbehaves.
