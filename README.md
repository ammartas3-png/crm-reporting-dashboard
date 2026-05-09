# CRM Reporting Dashboard

Interactive web GUI for generating the CRM + PowerBI Excel report.

## What it does

The web app keeps the same options as the original command-line script:

- PowerBI report file upload
- Optional PowerBI sheet name
- Pivot table name
- Up to four CRM file uploads in fixed browse rows
- Platform name beside each uploaded CRM file
- Optional CRM sheet name shared by all CRM files
- Optional output filename

After submission, the Python serverless function generates the enriched `.xlsx`
workbook and returns it as a browser download.

## Local development

Install dependencies:

```bash
python3 -m pip install -r requirements.txt
```

Run the CLI version:

```bash
python3 report_generator.py
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
