# Changelog
All notable changes to this project are documented here. Format: [Keep a Changelog](https://keepachangelog.com/), versioning: [SemVer](https://semver.org/).

## [1.0.0] - 2026-10-04
### Added
- Formula-driven dashboard workbook generated from `data/sales_data.csv` (2,707 orders, FY2026).
- 8 filters (Region, Country, Sales Channel, Customer Type, Category, Product, From/To month) wired to every visual.
- Power BI-style cross-filter highlighting: each chart ignores its own filter and greys out non-selected items.
- 6 KPI cards, highlights strip, 6 native charts, shaded tile map of Europe.
- Optional VBA module: click a bar / month / map tile to filter the whole dashboard.
- 4K PNG export pipeline (`scripts/render_png.py`).
- pytest suite comparing workbook results to an independent pandas calculation; GitHub Actions CI.
