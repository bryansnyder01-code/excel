# SheetCraft AI

Describe a spreadsheet in plain English, preview it live, edit it by chat or by hand, and download a real `.xlsx` with working formulas and professional styling.

## How it works

```
 Chat prompt ──► FastAPI ──► Claude (structured output) ──► WorkbookSpec JSON
                                                              │
                         formula validator (auto-fix + repair loop with Claude)
                                                              │
               ┌──────────────────────────────────────────────┴───────────┐
               ▼                                                          ▼
     Frontend grid preview (live, editable)                 openpyxl builder ──► .xlsx download
```

One JSON contract, `WorkbookSpec` (`backend/app/schemas.py`), is shared by the LLM, the API, the grid and the builder, so the preview and the downloaded file never drift apart. Formulas stay formulas: the `.xlsx` recalculates in Excel, LibreOffice and Google Sheets.

## Features

| Feature | How to use it | Where it lives |
|---|---|---|
| Build from a description | Type in the chat | `llm_engine.py`, `ChatPanel.tsx` |
| Edit by chat or by hand | Chat, or click / type in the grid | `SpreadsheetGrid.tsx`, `store.ts` |
| **Upload a file** | Upload button, the welcome panel, or drop a file anywhere. `.xlsx`, `.csv`, `.tsv` | `importer.py`, `actions.ts` |
| **Charts** | Select labels + number columns, then **Chart**; or ask the AI. Real Excel charts in the download, live previews under the grid | `workbook_builder.py`, `ChartsPanel.tsx`, `MiniChart.tsx` |
| **Explain / fix a formula** | Select a cell, click **Explain** (or **Fix this error** on an error), apply the suggested fix in one click | `llm_engine.explain`, `FormulaBar.tsx` |
| **Save and reopen** | Automatic. **My workbooks** lists everything saved in this browser | `persist.ts`, `DocsDialog.tsx` |
| **Templates** | **Templates** button or the welcome panel. 7 starters: budget, invoice, loan schedule, inventory, timesheet, sales tracker, project timeline | `app/templates/*.json`, `scripts/make_templates.py` |
| **Dropdown lists** | Select cells, click **Dropdown**; or ask the AI. Real Excel data validation in the download; a picker in the grid | `workbook_builder.py`, `DropdownDialog.tsx` |
| **Version history** | **Versions** tab: every AI change, upload, template and summary is a version; save your own; restore any | `store.ts`, `HistoryPanel.tsx` |
| **Summary tabs** | **Summarise**: group by one column, optionally split by another, sum / average / count / min / max. Builds a tab of live SUMIFS-style formulas plus a chart | `pivot.py`, `PivotDialog.tsx` |
| **More export formats** | Arrow next to Download: Excel, PDF (all sheets), CSV (current sheet, calculated values), Open in Google Sheets | `pdf_export.py`, `csv.ts`, `google.ts` |
| **Copy and paste with Excel** | Ctrl+C / X / V, or right-click. Pastes ranges from Excel or Google Sheets; copies back out as tab-separated text. Inside the grid, formulas move like Excel's | `clipboard.ts`, `fill.ts` |
| **Fill handle** | Drag the small square at the selection's corner. Continues numbers, dates, months, weekdays, quarters and "Item 3"; copies formulas with moved references. Ctrl+D / Ctrl+R fill down / right | `fill.ts`, `store.fill` |
| **Resizable columns and panels** | Drag a column letter's right edge (double-click to fit). Drag the line between chat and grid. Widths are saved into the .xlsx | `SpreadsheetGrid.tsx`, `Workspace.tsx` |
| **Freeze panes** | **Freeze** in the toolbar: header row and/or the first columns. Saved as Excel freeze panes | `FormatToolbar.tsx`, `workbook_builder.py` |
| **Formula autocomplete** | Type `=SU…` in a cell or the formula bar: function suggestions (Tab to insert), then the argument you're on. The cells a formula refers to are outlined in colour | `FormulaInput.tsx`, `functions.ts` |
| **Find and replace** | Ctrl+F / Ctrl+H, or the magnifier. Match case, whole cell, all sheets | `FindBar.tsx` |
| **Keyboard shortcuts** | Ctrl+/ or the keyboard icon | `ShortcutsDialog.tsx` |
| **Preview AI changes** | Changes to an existing workbook show highlighted, with **Keep** / **Discard** above the grid (a changed tab gets a dot). Tick "Apply future changes without asking" to skip it | `ReviewBar.tsx`, `diff.ts`, `store.startReview` |
| **Live progress** | The chat lists each step while the AI works (planning, writing each sheet, checking formulas, fixing problems) | `/generate/stream`, `/edit/stream`, `api.streamWorkbook` |
| **Ask about your data** | Switch the chat to **Ask about the data**. Answers use the calculated numbers and never change the workbook; cell links jump to the evidence | `/ask`, `actions.askData` |
| **Conditional formatting** | **Highlight** button or right-click: greater/less than, between, equal, negative, text contains, colour scales, data bars, custom colours | `ConditionalDialog.tsx`, `cellStyle.ts` |
| **Decimals and currency** | Currency picker (12 currencies) and more/fewer decimals buttons; per cell, or the whole column when the header is selected | `format.ts`, `workbook_builder.py` |
| **Cell comments** | Right-click → Add comment, Shift+F2, or the comment button. Hover the red corner to read; saved as Excel comments | `NoteEditor.tsx` |
| **Dark mode** | Moon / sun button; follows the system setting until you choose | `theme.ts`, `globals.css` |
| **Phone layout** | Under 768px wide, Chat and Sheet become tabs at the bottom and the toolbar scrolls sideways | `Workspace.tsx` |
| **Accounts, sync and share links** | **Sign in** to sync every workbook to your account (two-way, newest wins). **Share** makes a link; whoever opens it gets their own copy | `accounts.py`, `store.py`, `account.ts`, `ShareDialog.tsx` |
| **Sort and filter** | The arrow on any header cell: sort A→Z / Z→A (formulas move with their rows), or tick which values to show. Filters are saved as Excel autofilters | `FilterMenu.tsx`, `store.sortBy`, `layout.ts` |
| **Merge, wrap and row heights** | **Merge** and **Wrap text** buttons in the toolbar; drag a row number's bottom edge to resize (double-click resets). Wrapped rows grow to fit | `store.toggleMerge`, `SpreadsheetGrid.tsx` |
| **More charts, and editing them** | Column, bar, line, area, pie, doughnut, scatter; stacked; combo (some series as lines); per-series colours; axis titles. The pencil on a chart edits it | `ChartDialog.tsx`, `MiniChart.tsx`, `workbook_builder.py` |
| **Check my workbook** | **Tools → Check my workbook**: typed-in totals, formulas that break their column's pattern, error cells, references to empty cells, plus an optional AI review of the logic. Each finding jumps to its cell; many have a one-click **Fix** | `checker.py`, `CheckPanel.tsx` |
| **Clean up data** | **Tools → Clean up data**: remove duplicates, trim spaces, change case, split a column, fix text dates, fill blanks, text to numbers. Shows what will change first; one undo step | `cleanup.ts`, `CleanupDialog.tsx` |
| **What-if and goal seek** | **Tools → What-if**: Best / Base / Worst scenarios side by side (saved in the workbook and written to a Scenarios sheet in the download), and goal seek ("make H5 equal 150,000 by changing B2") | `whatif.ts`, `WhatIfDialog.tsx` |
| **Big sheets** | Only the rows on screen are drawn, so 20,000-row uploads scroll, sort and filter smoothly | `layout.ts`, `SpreadsheetGrid.tsx` |
| **Edit together, live** | **Share → Create invite link**. People who open it and sign in edit the same workbook: changes and everyone's selection show up as they happen | `live.py`, `live.ts` |
| **Google sign-in and password reset** | "Continue with Google" when a Google Client ID is set; **Forgot your password?** emails a one-hour reset link | `accounts.py`, `mailer.py`, `AccountDialog.tsx` |
| **Save as template** | **Tools → Save as template** (optionally emptying the typed data). Your templates appear at the top of **Templates** | `my_templates.py`, `userTemplates.ts` |
| **AI usage limits** | Each AI request counts against a rolling 24 hours: 30 without an account, 200 signed in (configurable). The chat shows what's left when it runs low | `usage.py` |
| **AI in formulas, and AI fill** | Type `=AI("Categorise", A2)` in any cell; answers are cached in the workbook and saved as values in the download. **Tools → Fill a column with AI** previews, then fills a whole column from an instruction | `formula.ts`, `aiFormula.ts`, `/ai-eval`, `/ai-fill`, `AiFillDialog.tsx` |
| **Photo or PDF to a table** | **Tools → Import from a photo or PDF**: receipts, scanned tables, PDF statements become sheets | `/extract`, `ImportImageDialog.tsx` |
| **Live data** | **Tools → Live data**: a CSV link, a published Google Sheet or a JSON API fills a sheet and refreshes on a schedule. Private and local network addresses are refused | `/api/data/fetch`, `connections.ts`, `scheduler.py` |
| **Dashboards** | The **Dashboard** tab: KPI tiles, charts and notes you can drag and resize; **Suggest a dashboard** builds one from the workbook. Exported as a Dashboard sheet | `components/dashboard/*`, `dashboard.ts` |
| **View-only sharing and locked cells** | **Share** has a *Can edit* and a *Can view* link, and a role per person. **Tools → Protect cells** locks a sheet or ranges so only the owner can change them (also protected in the download) | `access.ts`, `ProtectDialog.tsx`, `accounts.py`, `live.py` |
| **Comment threads and @mentions** | **Tools → Comments**, or the purple corner on a cell. Reply, resolve, and type @ to mention someone (they get an email) | `comments.py`, `CommentsPanel.tsx` |
| **Who changed what** | **Tools → Who changed what**: every live edit to the selected cell, or recent activity across the workbook | `cell_history`, `CellHistoryPanel.tsx` |
| **Compare versions** | **Versions → Compare**, or **Tools → Compare versions**: any two versions side by side, changed cells highlighted | `CompareDialog.tsx` |
| **Command palette** | Ctrl+K or the ⌘ button: search every action, sheet, cell and template | `commands.ts`, `CommandPalette.tsx` |
| **Drag-and-drop summaries** | **Summarise** now takes fields dragged into Rows / Columns / Values, with a live preview | `PivotDialog.tsx` |
| **Scheduled email reports** | **Tools → Email reports**: send the workbook (Excel or PDF, with an optional AI summary) daily, weekly or monthly | `reports`, `scheduler.py`, `ReportsDialog.tsx` |
| **Plans and billing** | **Tools → Plan and usage**: today's AI use, and an upgrade to Pro (more AI requests) through Stripe | `billing.py`, `BillingDialog.tsx` |
| **Accessibility** | The grid is a proper ARIA grid (screen readers hear the cell and its value), dialogs trap and restore focus, a skip link, visible focus rings, and reduced motion is respected | `SpreadsheetGrid.tsx`, `Modal.tsx`, `globals.css` |

### Setup notes for the new features

* **PDF export** uses LibreOffice on the server (`soffice`). Without it, the PDF option is greyed out with a note.
* **Open in Google Sheets** needs a Google OAuth Client ID in `frontend/.env.local` as `NEXT_PUBLIC_GOOGLE_CLIENT_ID` (Google Cloud Console: create an OAuth client of type *Web application*, add the site URL as an authorised JavaScript origin, and enable the Google Drive API). It only asks for access to files it creates. Without a Client ID, the menu item is greyed out.
* **Saved workbooks** live in the browser (localStorage) and, when you're signed in, in your account too. The last 15 versions of each workbook are kept.
* **Accounts** are stored in SQLite at `backend/data/sheetcraft.db` (set `SHEETCRAFT_DATA_DIR` to move it; keep it on a persistent disk in production). Passwords are hashed with scrypt; sign-in tokens are stored only as hashes and expire after 30 days without use. Password reset emails need the `SHEETCRAFT_SMTP_*` settings; without them the reset link is written to the server log.
* **Google sign-in** needs the same OAuth Client ID in `SHEETCRAFT_GOOGLE_CLIENT_ID` (backend) and `NEXT_PUBLIC_GOOGLE_CLIENT_ID` (frontend).
* **Live editing** keeps rooms in the backend's memory by default, which works for one backend process. To run several, set `SHEETCRAFT_REDIS_URL` and every instance shares live changes through Redis. Inserting or deleting rows and columns while others edit keeps everyone's undo history.
* **Live data and email reports** run from a once-a-minute background job inside the backend (`SHEETCRAFT_BACKGROUND_JOBS`, on by default). Reports need the `SHEETCRAFT_SMTP_*` settings; without them the email is written to the server log.
* **Billing** is off until `SHEETCRAFT_STRIPE_SECRET_KEY`, `SHEETCRAFT_STRIPE_PRICE_PRO` and `SHEETCRAFT_STRIPE_WEBHOOK_SECRET` are set. Point a Stripe webhook at `/api/billing/webhook` (checkout and subscription events). Pro accounts get `SHEETCRAFT_AI_LIMIT_PRO` AI requests a day (2,000 by default).
* **Locked cells** are enforced by the server for live edits and saves, not only in the browser.
* **Share links** give a copy; **invite links** give live editing of the same workbook.
* **Uploads** keep values, formulas, number formats, bold/italic, explicit colours and dropdowns. Merged cells, wrap text, row heights and filters carry over too. Charts and pivot tables in uploaded files are not carried over (the chat says so). Up to 20,000 rows and 100 columns per sheet.

## Project structure

```
sheetcraft-ai/
├── README.md
├── screenshots/
├── backend/                           Python 3.11+, FastAPI
│   ├── requirements.txt
│   ├── .env.example                   ANTHROPIC_API_KEY, model, effort, CORS
│   ├── app/
│   │   ├── main.py                    FastAPI app, CORS, /health
│   │   ├── config.py                  settings from env (SHEETCRAFT_*)
│   │   ├── schemas.py                 WorkbookSpec (+ charts, dropdowns) and API models
│   │   ├── api/routes.py              generate, edit, explain, import, pivot, validate, export, templates
│   │   ├── templates/                 7 starter workbooks (JSON)
│   │   └── services/
│   │       ├── prompts.py             system prompts + message templates
│   │       ├── llm_engine.py          NL → WorkbookSpec, explain-a-cell, validation, repair loop
│   │       ├── formula_validator.py   formula / chart / dropdown checks + safe auto-corrections
│   │       ├── workbook_builder.py    WorkbookSpec → styled .xlsx with charts and dropdowns
│   │       ├── importer.py            .xlsx / .csv → WorkbookSpec
│   │       ├── pivot.py               summary tabs from SUMIFS-family formulas
│   │       └── pdf_export.py          .xlsx → .pdf via LibreOffice
│   ├── samples/                       3-tab budget (.json and built .xlsx)
│   ├── scripts/
│   │   ├── build_sample.py            JSON → .xlsx without calling the LLM
│   │   └── make_templates.py          regenerates app/templates/
│   └── tests/                         37 tests, LLM mocked at the HTTP layer
│
└── frontend/                          Next.js 15 (App Router) + Tailwind 4
    ├── package.json
    ├── .env.example                   NEXT_PUBLIC_API_URL, NEXT_PUBLIC_GOOGLE_CLIENT_ID
    ├── app/                           layout, page, global styles, icon
    ├── components/
    │   ├── Workspace.tsx              layout, header, export menu, autosave, drag-and-drop upload
    │   ├── chat/ChatPanel.tsx         chat, welcome (upload, templates, examples), restore links
    │   ├── chat/HistoryPanel.tsx      version history
    │   ├── charts/                    chart previews (dependency-free SVG)
    │   ├── dialogs/                   templates, my workbooks, chart, dropdown, summary
    │   ├── grid/                      grid, toolbar, formula bar (+ explain), sheet tabs
    │   └── ui/Modal.tsx
    ├── lib/
    │   ├── types.ts                   WorkbookSpec types (mirror of schemas.py)
    │   ├── api.ts                     backend client + downloads
    │   ├── actions.ts                 chat / upload / template / summary flows
    │   ├── store.ts                   workbook, selection, undo/redo, versions, chat
    │   ├── persist.ts                 saved workbooks (localStorage)
    │   ├── formula.ts                 live formula evaluation
    │   ├── refs.ts                    shifts formula references on row/column insert/delete
    │   ├── format.ts, cellStyle.ts    Excel-style display and styling
    │   ├── csv.ts, google.ts          CSV export, Google Sheets export
    │   └── ui.ts                      dialogs and toasts
    └── scripts/check-formulas.ts      formula engine + every template must evaluate cleanly
```

### Frontend notes

* **The grid is built in-house** rather than with Handsontable, because Handsontable needs a paid licence for commercial use and its formula engine (HyperFormula) is GPL. Live results come from `fast-formula-parser` plus `formula.js` (both MIT); formula.js fills in functions the parser leaves unimplemented (MAX, SUMIFS, AVERAGEIFS, COUNTIFS, PMT, ...), and XLOOKUP is implemented locally. The downloaded file always contains the real formulas. `npm run check:formulas` checks the engine and every template.
* **Chat and manual edits share one workbook.** The first prompt on a blank workbook calls `/generate`; every later prompt sends the current workbook, including your manual edits, to `/edit`.
* **Inserting or deleting rows and columns rewrites formulas, chart ranges and dropdown ranges** the way Excel does.
* **Shortcuts:** see Ctrl+/ in the app. `npm run check:grid` tests copy/paste parsing, the fill handle, formula moving, autocomplete, number formats and the change diff.

## Backend API

| Method | Path | Body | Returns |
|---|---|---|---|
| POST | `/api/workbooks/generate` | `{prompt}` | `{workbook, summary, issues}` |
| POST | `/api/workbooks/edit` | `{prompt, workbook}` | `{workbook, summary, issues}` (full updated workbook) |
| POST | `/api/workbooks/validate` | `WorkbookSpec` | `{workbook, issues}` after manual grid edits |
| POST | `/api/workbooks/export?format=xlsx\|pdf` | `WorkbookSpec` | file download |
| POST | `/api/workbooks/import` | multipart `file` (.xlsx / .csv) | `{workbook, summary, issues}` |
| POST | `/api/workbooks/pivot` | `{workbook, source_sheet, row_field, column_field?, value_field?, aggregation}` | `{workbook, summary, issues}` with the new tab appended |
| POST | `/api/workbooks/explain` | `{workbook, sheet, cell, computed?, question?}` | `{explanation, fixed_value?}` |
| GET | `/api/workbooks/capabilities` | | `{pdf}` |
| GET | `/api/templates`, `/api/templates/{id}` | | template list / one `WorkbookSpec` |
| POST | `/api/workbooks/generate/stream`, `/edit/stream` | as generate / edit | `text/event-stream`: `progress` events, then one `result` or `error` |
| POST | `/api/workbooks/ask` | `{workbook, question, values}` | `{answer, references: [{sheet, range}]}` |
| POST | `/api/auth/signup`, `/login` | `{email, password}` | `{token, user}` |
| POST / GET | `/api/auth/logout`, `/api/auth/me` | Bearer token | 204 / `{id, email}` |
| GET / PUT / DELETE | `/api/docs`, `/api/docs/{id}` | `{title, updated, data}` for PUT | your synced workbooks |
| POST / DELETE | `/api/docs/{id}/share` | | `{token}` / 204 |
| GET | `/api/shared/{token}` | | `{title, workbook}` (public, no chat or history) |
| GET | `/health` | | `{status, model}` |

`issues` lists every problem found; `fixed: true` means the engine already corrected it (for example, a division wrapped in `IFERROR`). The frontend can show the rest as chips in the chat.

## Run it

Backend (terminal 1):

```bash
cd backend
python -m pip install -r requirements.txt
cp .env.example .env            # add your ANTHROPIC_API_KEY
uvicorn app.main:app --reload   # http://localhost:8000/docs
```

Frontend (terminal 2), Node 20+:

```bash
cd frontend
npm install
cp .env.example .env.local      # only if the backend isn't on localhost:8000
npm run dev                     # http://localhost:3000
```

Without an API key you can still run the backend tests and build the sample:

```bash
cd backend
python -m pytest -q
python -m scripts.build_sample samples/small_business_budget.json budget.xlsx
```

## Desktop app (Windows and Mac)

The `desktop/` folder packages the whole app as one installable program with Electron: the backend is bundled inside and serves the screens, and the app asks for an Anthropic API key on first launch. Push the project to GitHub and run **Actions → Desktop app** to get a Windows `.exe` and a Mac `.dmg`. It's for one person on their own computer; live co-editing, invite links and emailed reports need the hosted version. Details, first-launch warnings for unsigned apps, and how to build it yourself: [desktop/README.md](desktop/README.md).

## Run it with Docker

```bash
ANTHROPIC_API_KEY=sk-ant-... docker compose up --build    # http://localhost:3000
```

Data (accounts, synced workbooks, templates) lives in the `sheetcraft-data` volume. Add `WITH_PDF=1` to include LibreOffice for PDF export.

## Deploying

Any host that runs containers works (Render, Railway, Fly.io, a VPS). Two services from the two Dockerfiles:

1. **Backend** (`backend/Dockerfile`, port 8000) with a persistent disk mounted at `/data`. Set `ANTHROPIC_API_KEY`, `SHEETCRAFT_APP_URL` (the frontend's URL), `SHEETCRAFT_CORS_ORIGINS` (a JSON list with the frontend's URL), and `SHEETCRAFT_TRUST_PROXY=true` behind the host's proxy. Optional: SMTP and Google settings. Run one instance (see live editing above).
2. **Frontend** (`frontend/Dockerfile`, port 3000). `NEXT_PUBLIC_API_URL` (the backend's public URL) and `NEXT_PUBLIC_GOOGLE_CLIENT_ID` are build arguments, because Next.js compiles them in.

The host must allow WebSockets to the backend for live editing (all of the above do).

## Tests and CI

`.github/workflows/ci.yml` runs on every push and pull request: the backend's pytest suite, and the frontend's typecheck, formula/grid/tools checks and production build. Locally: `cd backend && python -m pytest -q`, and `cd frontend && npm run typecheck && npm run check:grid && npm run check:tools`.

## LLM choice

The engine uses Anthropic's Claude Sonnet 5.5 (`claude-sonnet-5-5`), the current successor to the Claude 3.5 Sonnet named in the spec, with structured outputs so every response parses straight into `WorkbookSpec`. Model and effort level are set by `SHEETCRAFT_MODEL` and `SHEETCRAFT_EFFORT`. Server-side refusal fallback is enabled, so a rare safety decline is retried on another model inside the same call instead of failing.

## Error handling (step 4 of the plan)

`formula_validator.py` runs on every generated, edited or exported workbook:

| Problem | What happens |
|---|---|
| Reference to a sheet that doesn't exist | Sent back to Claude for a repair pass |
| Cell that references itself (circular) | Sent back to Claude |
| Unbalanced parentheses or quotes | Sent back to Claude |
| Division by a literal `0` | Sent back to Claude |
| Division by a cell that might be empty or zero | Auto-wrapped in `IFERROR(...,0)` |
| Invalid, duplicate or too-long sheet name | Renamed; formulas pointing at it are rewritten |
| Unknown function, reference past the data, row wider than columns | Reported as a warning |
