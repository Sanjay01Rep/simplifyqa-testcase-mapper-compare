# ICEA LION Test Management Hub — How to unzip, set up, and run

Windows guide for managers and reviewers. You do **not** need GitHub.

This hub has four modules:

| Module | What it does |
|--------|----------------|
| **Map to SimplifyQA** | Client Excel → SimplifyQA import file (optional Kenya Pre-Requisites) |
| **Compare, Map & Report** | Excel A vs Excel B (for example Gen UG vs Life UG) → Common / Unique A / Unique B |
| **Map EP** | Build Execution Plans from SimplifyQA Live API or a Summary Excel (names only, one sheet per module) |
| **ICEA LION Reporter** | Daily FMS status reports, execution-sheet compare, schedules, and custom templates |

---

## What you need

| Item | Notes |
|------|--------|
| Windows 10 or 11 | This pack is built for Windows |
| **Node.js 18 or newer** | Includes `npm` |
| This ZIP | Unzip first, then install Node if you do not have it |
| SimplifyQA token (optional) | Needed only for **Map EP Live API** and Reporter project names. Put it in `.env` |

Check if Node is already installed. Open **PowerShell** or **Command Prompt** and run:

```powershell
node -v
npm -v
```

You should see versions such as `v18.x.x` or higher and an `npm` version. If the commands fail, install Node.js using the steps below.

---

## 1. Install Node.js 18+ (Windows)

### Official download (recommended)

1. Open: **https://nodejs.org/**
2. Click the **LTS** installer for Windows (this is **Node.js 18 or newer** and is the safest choice).
   - Direct Windows 64-bit LTS installer: **https://nodejs.org/en/download**
   - If you specifically want Node **18**: **https://nodejs.org/dist/latest-v18.x/**  
     Download `node-v18.x.x-x64.msi` (the highest 18.x `.msi` listed).

3. Run the `.msi` installer.
4. Keep the defaults. Make sure **“Add to PATH”** / **npm package manager** stays selected.
5. Click **Install**, then **Finish**.
6. **Close and reopen** PowerShell (PATH only updates in a new window).
7. Confirm:

```powershell
node -v
npm -v
```

`node -v` must be **v18.0.0 or higher**.

### If Windows blocks the installer

Use **Run as administrator**, or ask IT to allow Node.js. Company machines sometimes require admin rights.

---

## 2. Unzip this project

1. Find **`ICEA-Lion-Test-Management-Hub-share.zip`** (Desktop, email, or Teams). Older packs may still be named `ICEA-Lion-Testcase-Review-share.zip`.
2. Right-click the ZIP → **Extract All…**
3. Choose a simple folder, for example:

   `C:\Users\<you>\Desktop\ICEA LION Test Management Hub`

4. Click **Extract**.
5. Open that folder. You should see `package.json`, `server.js`, `START-HERE.md`, `Client doc`, and `Kenya doc`.

Do **not** run anything from inside the ZIP without extracting.

---

## 3. Install project dependencies

Open PowerShell **in the unzipped folder**.

**Easy way:** in File Explorer, click the address bar, type `powershell`, press Enter.

Then:

```powershell
npm install
```

**Do not skip this.** The ZIP does not include `node_modules`. Skipping it causes `Cannot find module 'express'`.

Wait until it finishes with no error, then go to step 4.

If `npm` is not recognized, Node is not on PATH — reopen PowerShell after installing Node, or reinstall Node.js.

---

## 4. Start the application

```powershell
npm start
```

You should see something like:

```
ICEA LION Test Management Hub UI  v1.3.0  http://localhost:3100
```

Leave this window **open**. Closing it stops the app.

### Optional: double-click launcher

You can instead double-click **`start-ui.cmd`**.

### Optional: SimplifyQA token (Map EP Live API and Reporter)

Copy `.env.example` to `.env` and paste your bearer token:

```powershell
Copy-Item .env.example .env
```

Then edit `.env` and set:

```
SIMPLIFYQA_BEARER_TOKEN=your-token-here
```

Restart `npm start` after saving. Never share or commit `.env`.

---

## 5. Open the UI

In Chrome or Edge go to:

**http://localhost:3100**

Do **not** use VS Code / Cursor **Live Preview**. That will not run the APIs.

The header **server bot** should show **Server up** (green).

Stop the app: in the PowerShell window press **Ctrl+C**.

---

## 6. What to do in the UI

### Map to SimplifyQA

1. Choose a client Excel (`Client doc` or upload). `.xlsx` only.
2. Optionally add a **Mapper file** from `Kenya doc` (Pre-Requisites only — it does not add extra testcases).
3. If the file has more than one sheet, pick the sheet when asked.
4. Select **Module** and **Entity** (Entity is not selected by default; you can type a custom entity).
5. **Review only** or **Generate Excel**.
6. Download the SimplifyQA workbook from the result panel. **Reset** clears the form and the results.

Outputs are saved under **`Generated Excel file/`**.

### Compare, Map & Report

1. **Upload files** — **Excel A** and **Excel B** (required). Mapper file is optional. Pick a sheet per file if needed. Click **Next: Configure**.
2. **Configure & generate** — choose **Module**, entities per sheet, then **Review** or **Generate Excel**.

Results:

- **Common** — same name/ID and same steps
- **Unique A / Unique B** — unmatched cases (empty unique sheets are **not** created)

### Map EP

1. Choose **Live SimplifyQA API** or **Summary Excel**.
2. Select **project**, **one or more modules**, and **one entity**. Assignee email is optional.
3. Generate. You get one workbook with **one sheet per module** (names only, not steps).
4. Use **Open Excel** from the result panel.

### ICEA LION Reporter

1. Pick a **project** from the dropdown (names come from SimplifyQA when a token is set).
2. Pick a **template** (built-in, or upload a custom 4th template).
3. Generate or schedule the FMS report. Compare execution sheets here if needed. **View Excel** is not used in this module.

---

## 7. Folders in this pack

| Folder / file | Meaning |
|---------------|---------|
| `Client doc/` | Real client `.xlsx` files (Gen UG / Life UG, etc.) |
| `Kenya doc/` | Mapper / Kenya source files for Pre-Requisites |
| `Kenya orginial testcase/` | Extra mapper sources |
| `Generated Excel file/` | Map / Compare / EP outputs (created when you generate) |
| `output/` | Reporter outputs |
| `Template/` | Reporter Excel templates |
| `mapping.properties` | Local module/entity defaults |
| `.env` | Local token and port (gitignored) |
| `start-ui.cmd` | Windows start script |

Use **`.xlsx` only** (max 25 MB per upload).

---

## 8. Common problems

| Problem | What to do |
|---------|------------|
| `node` / `npm` is not recognized | Install Node 18+, close all terminals, open a **new** PowerShell |
| `Cannot find module 'express'` | You skipped `npm install`. Run `npm install` in the unzipped folder, then `npm start` |
| Browser cannot open localhost:3100 | Make sure `npm start` is still running |
| UI loads but APIs fail / “wrong server” | Close Live Preview; use **http://localhost:3100** after `npm start` |
| Port 3100 already in use | Close other Node windows, then `npm start` again (it frees port 3100) |
| Green bot then red | The start window was closed — run `npm start` again |
| Excel will not upload | File must be `.xlsx`, not `.xls` or `.xlsm` |
| Map EP / Reporter asks for a token or returns 403 | Put `SIMPLIFYQA_BEARER_TOKEN` in `.env`, restart `npm start` |
| Multi-sheet file error | Pick the sheet from the alert / dropdown; sheet names are not hardcoded |

---

## 9. Optional checks

From the same unzipped folder:

```powershell
npm test
```

This runs automated checks. It is **not** required to use the UI.

---

## Need more detail?

See **`README.md`** in this folder for configuration, matching rules, Map EP / Reporter behaviour, and CLI mapping (`npm run map`).
