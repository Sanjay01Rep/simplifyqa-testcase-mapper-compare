# SimplifyQA Testcase Mapper & Compare

Internal QA tooling for **ICEA LION** Uganda workstreams. It reviews client Excel testcases, maps Kenya pre-requisites where available, and exports workbooks in **SimplifyQA** template format.

**Repository:** [Sanjay01Rep/simplifyqa-testcase-mapper-compare](https://github.com/Sanjay01Rep/simplifyqa-testcase-mapper-compare)

---

## What it does

| Module | Purpose |
|--------|---------|
| **Map to SimplifyQA** | Convert a client `.xlsx` into a SimplifyQA-ready workbook, optionally filling Pre-Requisite from Kenya source docs |
| **Compare clients** | Compare Client A vs Client B (e.g. Gen UG vs Life UG) and split into **Common**, **Unique A**, and **Unique B** sheets |

Outputs land under `Generated Excel file/` with companion `.log` files under `Generated Excel file/logs/`.

---

## Requirements

- **Node.js 18+**
- Windows recommended (UI launcher uses `.cmd`; paths work on other OS via `npm start`)
- Input files must be **`.xlsx`** (max upload size 25 MB)

---

## Quick start

```bash
npm install
npm start
```

Then open: **http://localhost:3100**

> Do **not** use IDE Live Preview. Use the Express server above (or double-click `start-ui.cmd`).

Optional port override:

```bash
# PowerShell
$env:PORT=3200; npm start
```

---

## Usage (UI)

### Map to SimplifyQA

1. Open **Map to SimplifyQA**.
2. Upload or select a **Client** workbook.
3. Optionally select a **Kenya** workbook for Pre-Requisite mapping.
4. Set **Module**, **Entity**, **Versions**, **Testcase type** (or load from `mapping.properties`).
5. Run **Review** / **Generate**.
6. Download the Excel and log from the result panel.

### Compare clients

1. Open **Compare clients**.
2. Upload **Client A** and **Client B** (optional Kenya for prereqs).
3. Set entity labels for Common / Unique A / Unique B sheets.
4. Run compare.
5. Download the 3-sheet SimplifyQA workbook + log.

**Matching rules (high level):**

- Same testcase **name** (soft text match allowed for typos/abbreviations).
- If both sides have a real TC ID, IDs must match.
- **Steps + expected** must match (soft compare) to go to **Common**.
- Name/ID match but different steps → listed on **both Unique** sheets (not Common).

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
├── server.js                 # Express UI + APIs (port 3100)
├── map_to_simplifyqa.js      # CLI mapper
├── mapping.properties.example # Committed template for local mapping.properties
├── lib/
│   ├── mapper.js             # Parse client Excel, build SimplifyQA rows
│   ├── compare.js            # Client A vs B compare logic
│   ├── loadEnv.js            # Optional .env loader
│   └── options.js            # Module / Entity dropdowns
├── public/                   # Browser UI (HTML / CSS / JS)
├── scripts/free-port.js      # Frees port 3100 before start
├── test/e2e.js               # Automated checks
├── Client doc/               # Local inputs (not in git)
├── Kenya doc/                # Local Kenya sources (not in git)
├── Kenya orginial testcase/  # Extra Kenya sources (not in git)
└── Generated Excel file/     # Outputs + logs (not in git)
```

Local Excel folders are **gitignored** so client data is not pushed to GitHub.

---

## Configuration

### `mapping.properties`

Local run settings (gitignored). On first start, if missing, the app copies from `mapping.properties.example`.

```bash
cp mapping.properties.example mapping.properties
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

Copy `.env.example` → `.env` (optional). The server auto-loads `.env` on startup and does **not** override variables already set in the shell.

| Variable | Default | Meaning |
|----------|---------|---------|
| `PORT`   | `3100`  | HTTP port for the UI |

```powershell
# Option A — .env file
Copy-Item .env.example .env
# edit PORT=3200 inside .env
npm start

# Option B — shell only
$env:PORT = "3200"
npm start
```

---

## Screenshots / demo (for reviewers)

When preparing a review pack, capture:

1. **Home / Map tab** with Module & Entity filled  
2. **Compare tab** with Client A + Client B selected  
3. Sample **output Excel** (Common / Unique A / Unique B sheet tabs)  
4. A short **log excerpt** showing Common count and `STEPS_DIFFER` warnings  

Place images under a local `docs/screenshots/` folder if you add them later (optional; not required to run the app).

---

## For managers / clients — review checklist

- [ ] `npm install` && `npm start` opens http://localhost:3100  
- [ ] Map one known client file end-to-end; download Excel + log  
- [ ] Compare Gen UG vs Life UG for one module; confirm Common vs Unique counts  
- [ ] Confirm `npm test` passes in this environment  
- [ ] Confirm no client Excel data was committed to the repo  

---

## License

ISC — internal Simplify3x / ICEA LION delivery use.
