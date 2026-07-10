# AGENTS.md

## Cursor Cloud specific instructions

This repo bundles two loosely-coupled products (see `README.md` for full details):

1. **CRM Report web GUI** — static `index.html`/`styles.css` served by the Python
   serverless handler in `api/generate.py`. Core Excel logic lives in
   `report_generator.py`, `program_a_report.py`, `program_b_country_report.py`,
   `lead_splitter.py`, `cr_maker.py`. The GUI posts `multipart/form-data` to
   `POST /api/generate`.
2. **Telegram reporting bot** — Next.js App Router route `app/api/telegram/route.js`
   (webhook `/api/telegram`) plus the landing page `app/page.js`.

### Running services in dev

- **Python GUI (port 8000):** `README.md` documents `npx vercel dev`, but that
  requires a Vercel login/link that is not available in this environment. Instead,
  run the handler directly with the stdlib HTTP server (it is a plain
  `BaseHTTPRequestHandler`, and `do_POST` ignores the path so `/api/generate`
  works):
  ```bash
  python3 -c "import sys; sys.path.insert(0,'api'); from http.server import HTTPServer; from generate import handler; HTTPServer(('0.0.0.0',8000), handler).serve_forever()"
  ```
  Then open http://localhost:8000/ — the Report, Lead Splitter, and CR tabs work
  fully offline with no credentials.
- **Next.js (port 3000):** `npm run dev`. Serves the landing page and the
  `/api/telegram` webhook. The webhook responds `{"ok":true,"ignored":true}` to
  unknown/empty payloads, so it can be smoke-tested without Telegram/Google creds.

### Testing

- JS: `npm test` (`node --test tests/js/*.test.js`).
- Python: `python3 -m unittest discover -s tests -p "test_*.py"`. Note: `pytest` is
  **not** a dependency, so use `unittest` (not `pytest`).

### Gotchas

- The Python CLI entrypoints (`python3 program_a_report.py`, etc.) are
  **interactive** — they prompt for file paths via stdin and will hang in a
  non-interactive shell. Use the web GUI (or call `build_output*` directly) to
  exercise report generation programmatically.
- External services are only needed for two sub-flows: the **Database check** tab
  (Google Sheets + n8n webhook) and the **Telegram bot** (Telegram Bot API +
  Google Sheets). Required env vars (`TELEGRAM_BOT_TOKEN`, `GOOGLE_PRIVATE_KEY`,
  `GOOGLE_SERVICE_ACCOUNT_EMAIL`, `GOOGLE_SPREADSHEET_ID`, `ALLOWED_USERS`) come
  from `.env.example`; copy to `.env.local` and fill in to test those flows.
- The report engine merges CRM rows to PowerBI rows on CRM `ID` == PowerBI
  `Account No`. Sample inputs for manual GUI testing must include the required
  columns listed in `README.md`.
