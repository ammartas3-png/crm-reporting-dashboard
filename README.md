# CRM Reporting Dashboard

Interactive web GUI for generating CRM + PowerBI Excel reports.

## What it does

The web app lets you choose between two separate report programs:

- **Report Generator**: the original deployed CRM output workbook
  with combined data and pivots.
- **Bulk Country Reports (M-inhousemedia report etc.)**: a Main Report sheet
  plus one sheet per Country, with formula-linked country rows and centered
  pivots.

Both programs use the same uploaded inputs:

- PowerBI report file upload
- Pivot table name for Report Generator
- Up to four CRM file uploads in fixed browse rows
- Platform dropdown beside each uploaded CRM file
- Optional CRM sheet name shared by all CRM files
- Optional output filename. Report Generator defaults to `crm_powerbi_output.xlsx`;
  Bulk Country Reports defaults to `crm_country_report.xlsx`.

After submission, the Python serverless function generates the enriched `.xlsx`
workbook and returns it as a browser download.

## Local development

Install dependencies:

```bash
python3 -m pip install -r requirements.txt
```

Run the CLI version:

```bash
python3 program_a_report.py
python3 program_b_country_report.py
```

For the Vercel-style web app, install the Vercel CLI and run:

```bash
npx vercel dev
```

Then open the local URL printed by Vercel.

## Deploying to Vercel

This repository is Vercel-ready:

1. Import the GitHub repository into Vercel.
2. Keep the default framework preset as "Other".
3. Deploy. Vercel will install `requirements.txt`, serve `index.html`, and run
   `api/generate.py` as the Python serverless function.

## Required spreadsheet columns

CRM files must include:

```text
Customer Type, ID, Created, Name, Department, Status, Country, Campaign,
Sub-Campaign, Placement, Assigned to
```

The PowerBI report must include:

```text
Account No, Brand, Last 10 Comments, Voip Calls Attempts Cnt
```
