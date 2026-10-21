# CRM Campaign Dashboard in Excel

A modern, SaaS-inspired CRM dashboard built in Microsoft Excel. It demonstrates how spreadsheet analytics and UI design can work together to create a clean, business-friendly reporting experience.

> **Data notice:** The workbook contains fictional sample CRM activity for demonstration purposes. Replace it with approved business data before using it for real reporting.

## Preview

Open `CRM_Interactive_Dashboard.xlsx` and select the **CRM Dashboard** worksheet.

The visual design includes a dark sidebar, pastel KPI cards, campaign activity charts, status summaries, recent activity, and campaign-level metrics.

## Features

- KPI cards for Calls, SMS, Emails, and Completion Rate
- Monthly channel activity chart
- Channel mix chart
- Status breakdown: Completed, Pending, and Failed
- Recent activity table
- Campaign activity summary
- Period and Channel dropdown filters
- Separate Raw Data and Analysis worksheets
- Modern, SaaS-inspired layout and color palette

## Repository structure

```text
crm-dashboard-excel-github/
├── CRM_Interactive_Dashboard.xlsx
├── README.md
├── LICENSE
├── .gitignore
├── data/
│   └── crm_sample_data.csv
└── docs/
    └── DATA_DICTIONARY.md
```

## Getting started

1. Download or clone this repository.
2. Open `CRM_Interactive_Dashboard.xlsx` in desktop Microsoft Excel.
3. Enable automatic calculation if prompted, and allow formulas to recalculate.
4. Go to the **CRM Dashboard** worksheet.
5. Change the **Period** and **Channel** dropdowns to explore the report.
6. Review **Raw Data** for the fictional source records and **Analysis** for chart-supporting calculations.

## Data fields

| Field | Description |
|---|---|
| Date | Date and time of the CRM interaction |
| Campaign | Campaign associated with the interaction |
| Channel | Call, SMS, or Email |
| Caller ID | Originating caller/contact identifier |
| Destination | Recipient or destination identifier |
| Status | Completed, Pending, or Failed |
| Response Time (min) | Response duration in minutes, where available |
| Agent | Assigned CRM agent |

## Customization

- Replace the rows in **Raw Data** with your own data while preserving the column headers.
- If you add more records than the sample workbook contains, extend the formula ranges and chart source ranges accordingly.
- Keep dates as real Excel date/time values and use consistent category values for `Channel` and `Status`.
- Review the calculations before using the dashboard for business decisions.
- For a production-ready version, consider converting source data into an Excel Table and using PivotTables, PivotCharts, and native slicers.

## Important limitations

- This is an Excel workbook, not a hosted web application.
- Dropdowns are used for filtering; native Excel slicers are not included.
- Some Excel formulas and chart behaviors may vary in third-party spreadsheet apps. Desktop Microsoft Excel is recommended.
- The dashboard uses fictional sample data and should not be treated as actual CRM performance.

## Suggested GitHub topics

`excel` `excel-dashboard` `crm` `data-analytics` `data-visualization` `business-intelligence` `dashboard-design`

## Author / attribution

Created as a portfolio-style CRM dashboard project. Customize this section with your name, LinkedIn profile, and portfolio links before publishing.
