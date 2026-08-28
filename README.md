# ICEA LION Test Management Hub

Internal QA tooling for **ICEA LION** workstreams. It maps client Excel testcases into SimplifyQA format, compares regional workbooks, generates multi-module Execution Plans (EP), and runs the ICEA LION Daily FMS status reporter.

**First-time setup (unzip, Node.js, run):** see **[START-HERE.md](./START-HERE.md)**

---

## What it does

| Module | Purpose |
|--------|---------|
| **Map to SimplifyQA** | Convert a client `.xlsx` into a SimplifyQA-ready import file. Optionally fill **Pre-Requisite** from a Kenya mapper file. Module and Entity come from the UI (or `mapping.properties`), not from the client sheet. |
| **Compare, Map & Report** | Compare Client A vs Client B (for example Gen UG vs Life UG) and split into **Common**, **Unique A**, and **Unique B**. |
| **Map EP** | Fetch testcase **names** (not steps) from SimplifyQA Live API or a Summary Excel, filter by one or more modules and one entity, and write a standard Execution Plan workbook (one sheet per module). |
| **ICEA LION Reporter** | Generate FMS status reports, compare execution sheets, manage daily schedules, pick a live project from SimplifyQA, and upload a custom 4th template. |

Outputs land under `Generated Excel file/` and `output/` with companion logs under `Generated Excel file/logs/` and `logs/`.

---

## Requirements

- **Node.js 18+**
- Windows recommended (UI launcher uses `.cmd`; paths work on other OS via `npm start`)
- Input files must be **`.xlsx`** (max upload size 25 MB)
- For **Map EP Live API** and Reporter project list: `SIMPLIFYQA_BEARER_TOKEN` in `.env`

---

## Quick start

```bash
npm install
npm start
```

Then open: **http://localhost:3100**

> Do **not** use IDE Live Preview. Use the Express server above (or double-click `start-ui.cmd`).

Optional port override:

```powershell
$env:PORT=3200; npm start
```

---

## Usage (UI)

### Map to SimplifyQA

1. Open **Map to SimplifyQA**.
2. Upload or select a **Client** workbook (`.xlsx` only).
3. Optionally select a **Kenya / mapper** workbook for Pre-Requisite mapping.
4. If the client file has more than one sheet, pick the sheet when prompted (sheet names are not hardcoded).
5. Set **Module**, **Entity**, **Versions**, **Testcase type** (or load from `mapping.properties`). Entity is not selected by default; you can add a custom entity.
6. Run **Review** / **Generate**.
7. Download the Excel and log from the result panel. **Reset** clears inputs and results.

**Mapping rules:**

- Last step of each testcase: **Mandate Screenshot = Yes**.
- Missing sequences are corrected.
- Module / Entity on the output come from the UI / properties file, not from the client columns.
- Kenya is matched on **testcase name** (case-insensitive, extra spaces ignored). Empty Kenya Pre-Requisite is skipped, not treated as an error. **No partial name match.**
- Kenya is used only to fill Pre-Requisite. It does **not** add extra Kenya testcases. Output count stays the client count.

### Compare, Map & Report

1. Open **Compare, Map & Report**.
2. Upload **Excel A** and **Excel B** (optional Kenya for prereqs). Pick a sheet per file if a workbook has multiple sheets.
3. Set entity labels for Common / Unique A / Unique B (custom entity allowed; nothing selected by default).
4. Run compare / generate.
5. Download the 3-sheet SimplifyQA workbook + log.

**Matching rules (high level):**

- Same testcase **name** (soft text match allowed for typos/abbreviations).
- If both sides have a real TC ID, IDs must match.
- **Steps + expected** must match (soft compare) to go to **Common**.
- Name/ID match but different steps → listed on **both Unique** sheets (not Common).
- Common + Unique A + Unique B must not invent extra testcases.

### Map EP

1. Open **Map EP**.
2. Choose source: **Live SimplifyQA API** or **Summary Excel**.
3. For Live API, the bearer token is read from `.env` (`SIMPLIFYQA_BEARER_TOKEN`). You can change it in the UI only if needed.
4. Select **project**, **one or more modules**, and **one entity**. Assignee email is optional and usually left blank.
5. Generate. The workbook has **one sheet per selected module** (sheet name = module + entity) with testcase **names only**.
6. Use **Open Excel** from the result panel.

Selected modules are filtered strictly. Sibling or E2E modules are not pulled in unless you selected them. Counts should match SimplifyQA for that module + entity.

### ICEA LION Reporter

1. Open **ICEA LION Reporter**.
2. Choose **project** from the dropdown (live names from the token; a newly used project stays available next time).
3. Choose a **template** (built-in templates 1–3, or upload a custom 4th template from the UI).
4. Generate / schedule / compare execution sheets as in the original Reporter.
5. **View Excel** is not available in this module.

---

## Usage (CLI)

```bash
# Map one client file (uses mapping.properties / jobs.json)
npm run map -- "Your Client File.xlsx"

# Map all client files under Client doc/
npm run map:all

# List client files
node map_to_simplifyqa.js --list

# Run automated checks
npm test
```

---

## Project layout

```
├── server.js                   # Express UI + APIs (port 3100)
├── map_to_simplifyqa.js        # CLI mapper
├── mapping.properties.example  # Template for local mapping.properties
├── lib/
│   ├── mapper.js               # Client Excel → SimplifyQA rows
│   ├── compare.js              # Client A vs B compare
│   ├── epMapper.js             # Execution Plan mapping
│   ├── loadEnv.js             # .env + SimplifyQA token
│   ├── options.js             # Module / Entity dropdowns
│   └── reporter/              # ICEA LION Reporter module
├── public/                     # Browser UI
├── config/                     # Reporter application.properties (local)
├── Template/                  # Reporter Excel templates
├── scripts/                    # Version bump, free-port
├── test/                       # Automated checks
├── Client doc/                 # Local client inputs (not in git)
├── Kenya doc/                  # Kenya mapper sources (not in git)
├── Kenya orginial testcase/    # Extra Kenya sources (not in git)
├── Generated Excel file/       # Map / Compare / EP outputs + logs
└── output/                     # Reporter outputs
```

Local Excel folders, `.env`, and `mapping.properties` are **gitignored** so client data and tokens are not pushed to GitHub.

---

## Configuration

### `mapping.properties`

Local run settings (gitignored). On first start, if missing, the app copies from `mapping.properties.example`.

```powershell
Copy-Item mapping.properties.example mapping.properties
```

Example keys:

```properties
Module=Investment Management
Entity=Life UG
Versions=v1.0
TestcaseType=WEB

CompareModule=Accounts Payable
CompareEntityCommon=Life UG, Gen UG
CompareEntityUniqueA=Gen UG
CompareEntityUniqueB=Life UG
```

### Environment

Copy `.env.example` → `.env` (optional for Map/Compare; required for Live API / Reporter projects). The server auto-loads `.env` on startup and does **not** override variables already set in the shell.

| Variable | Default | Meaning |
|----------|---------|---------|
| `PORT` | `3100` | HTTP port for the UI |
| `HEALTH_POLL_MS` | `300000` (5 min) | How often the UI status bot calls `/api/health` |
| `SIMPLIFYQA_BEARER_TOKEN` | (empty) | SimplifyQA bearer token for Map EP Live API and Reporter project list |

```powershell
Copy-Item .env.example .env
# paste SIMPLIFYQA_BEARER_TOKEN=... inside .env
npm start
```

Never commit `.env`.

---

## Screenshots / demo (for reviewers)

When preparing a review pack, capture:

1. **Map to SimplifyQA** with Module & Entity filled
2. **Compare** with Client A + Client B selected
3. **Map EP** with more than one module selected
4. **ICEA LION Reporter** project + template dropdowns
5. Sample **output Excel** tabs and a short **log** excerpt

Place images under a local `docs/screenshots/` folder if you add them later (optional; not required to run the app).

---

## For managers / clients — review checklist

- [ ] `npm install` && `npm start` opens http://localhost:3100
- [ ] Map one known client file end-to-end; download Excel + log
- [ ] Compare Gen UG vs Life UG for one module; confirm Common vs Unique counts
- [ ] Map EP for one or more modules; sheet names match the selected modules; counts match SimplifyQA
- [ ] Reporter: project dropdown shows live names; generate using a built-in or uploaded template
- [ ] Confirm `npm test` passes in this environment
- [ ] Confirm no client Excel data or `.env` tokens were committed to the repo

---

## License

ISC — internal Simplify3x / ICEA LION delivery use.
