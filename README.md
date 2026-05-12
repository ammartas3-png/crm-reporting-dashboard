# CRM Reporting Dashboard

Interactive web GUI for generating CRM + PowerBI Excel reports.

## What it does

The web app lets you choose between two separate report programs:

- **Program A - CRM Output Report**: the original deployed CRM output workbook
  with combined data and pivots.
- **Program B - Country Split Report**: a Main Report sheet plus one sheet per
  Country, with formula-linked country rows and centered pivots.

Both programs use the same uploaded inputs:

- PowerBI report file upload
- Optional PowerBI sheet name
- Pivot table name for Program A
- Up to four CRM file uploads in fixed browse rows
- Platform name beside each uploaded CRM file
- Optional CRM sheet name shared by all CRM files
- Optional output filename. Program A defaults to `crm_powerbi_output.xlsx`;
  Program B defaults to `crm_country_report.xlsx`.

After submission, the Python serverless function generates the enriched `.xlsx`
workbook and returns it as a browser download.

## Local development

Install dependencies:

```bash
python3 -m pip install -r requirements.txt
npm install
```

Run the CLI version:

```bash
python3 program_a_report.py
python3 program_b_country_report.py
```

For the Vercel-style web app and Telegram bot, install the Vercel CLI and run:

```bash
npx vercel dev
```

Then open the local URL printed by Vercel.

## Deploying to Vercel

This repository is Vercel-ready:

1. Import the GitHub repository into Vercel.
2. Deploy the Next.js app. Vercel will install `package.json` dependencies and
   serve the Telegram route at `/api/telegram`.
3. The existing Python workbook endpoint remains in `api/generate.py`.

## Telegram bot

The reporting bot webhook lives at `/api/telegram` using the Next.js App Router:

```text
app/api/telegram/route.js
lib/telegram.js
lib/googleSheets.js
lib/queryRouter.js
lib/calculations.js
lib/permissions.js
config/sheetsConfig.js
```

The bot receives Telegram messages, checks the Telegram user ID against
`ALLOWED_USERS`, reads Google Sheets rows, calculates simple metrics, and sends a
short answer back to Telegram.

Supported commands:

- `/start` - show example report questions.
- `/help` - show supported examples.

Supported MVP questions include:

- `How many FTD today?`
- `Germany total leads?`
- `Ahmet total calls?`
- `May Turkey leads count?`

Configure these environment variables in Vercel:

- `TELEGRAM_BOT_TOKEN` - BotFather token.
- `GOOGLE_SERVICE_ACCOUNT_EMAIL` - Google service account email.
- `GOOGLE_PRIVATE_KEY` - Google service account private key. Store multiline
  keys with escaped newlines (`\n`) if your secret manager requires it.
- `GOOGLE_SPREADSHEET_ID` - default spreadsheet ID.
- `ALLOWED_USERS` - comma-separated Telegram user IDs, for example
  `123456789,987654321`.

Optional tab/range overrides:

- `GOOGLE_LEADS_TAB`, `GOOGLE_LEADS_RANGE`
- `GOOGLE_FTD_TAB`, `GOOGLE_FTD_RANGE`
- `GOOGLE_TRANSACTION_TAB`, `GOOGLE_TRANSACTION_RANGE`

Keep the BotFather token out of the repository. Use it only from a secure shell
or secret manager when registering the webhook:

```bash
export TELEGRAM_BOT_TOKEN="your-bot-token"
export PUBLIC_APP_URL="https://your-next-app.vercel.app"

curl -X POST "https://api.telegram.org/bot${TELEGRAM_BOT_TOKEN}/setWebhook" \
  -H "Content-Type: application/json" \
  -d "{\"url\":\"${PUBLIC_APP_URL}/api/telegram\"}"
```

The initial tab configuration is in `config/sheetsConfig.js`. Update that file,
or the optional tab/range environment variables, when final Google Sheet and tab
names are available.

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
