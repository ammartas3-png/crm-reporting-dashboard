# crm-reporting-dashboard
CRM Reporting Tool for FTD, CR, performance analytics, agents calls, check-in check out and country aff report

## CRM + PowerBI Excel merger

`crm_powerbi_merge.py` creates one output Excel workbook from:

- one PowerBI report Excel file
- one or more CRM Excel files

For every CRM file, the tool asks for the platform name and adds it as the first
output column named `Platform`. If you are running non-interactively, pass one
`--platform` value per CRM file in the same order as `--crm-files`.

### Install dependencies

```bash
python3 -m pip install -r requirements.txt
```

### Run

```bash
python3 crm_powerbi_merge.py \
  --powerbi-report "powerBI report.xlsx" \
  --crm-files "crm file 1.xlsx" "crm file 2.xlsx" \
  --output "crm output.xlsx"
```

Non-interactive example:

```bash
python3 crm_powerbi_merge.py \
  --powerbi-report "powerBI report.xlsx" \
  --crm-files "crm file 1.xlsx" "crm file 2.xlsx" \
  --platform "Brand A" \
  --platform "Brand B" \
  --output "crm output.xlsx"
```

The output columns are:

`Platform / Customer Type / ID / Created / Name / Department / Status / Country / Assigned to / Comments / Call Attempts`

CRM rows are matched to PowerBI rows by:

- CRM `ID` = PowerBI `Account No`
- entered platform = PowerBI `Brand name`

The `Comments` output comes from PowerBI `Last 10 Comments`, extracting the text
between the final `|` and `;` in each comment entry and writing it from bottom to
top. `Call Attempts` comes from PowerBI `Voip Calls Attempts Cnt`.

No comments are written for these CRM statuses:

- `DNC`
- `Invalid country`
- `Wrong number or email`
- `Duplicate`
- `no potential - no documents`
- `Test`
- `Under 18`
- `No language`

For `No Answer 1` through `No Answer 5`, the output comment is `NA VM`.
