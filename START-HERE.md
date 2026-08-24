# ICEA LION Testcase Review — How to unzip, set up, and run

Windows guide for managers and reviewers. You do **not** need GitHub.

This tool maps client testcase Excels into SimplifyQA format and compares **Excel A vs Excel B** (for example Gen UG vs Life UG).

---

## What you need

| Item | Notes |
|------|--------|
| Windows 10 or 11 | This pack is built for Windows |
| **Node.js 18 or newer** | Includes `npm` |
| This ZIP | Unzip first, then install Node if you do not have it |

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

1. Find **`ICEA-Lion-Testcase-Review-share.zip`** (Desktop, email, or Teams).
2. Right-click the ZIP → **Extract All…**
3. Choose a simple folder, for example:

   `C:\Users\<you>\Desktop\ICEA Lion Testcase Review`

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
ICEA LION Testcase Review UI  v1.1.0  http://localhost:3100
```

Leave this window **open**. Closing it stops the app.

### Optional: double-click launcher

You can instead double-click **`start-ui.cmd`**.

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

1. Choose a client Excel (`Client doc` or upload).
2. Optionally add a **Mapper file** from `Kenya doc` (Pre-Requisites).
3. Select **Module** and **Entity** (one or more).
4. **Review only** or **Generate Excel**.
5. Download the SimplifyQA workbook from the result panel.

Outputs are saved under **`Generated Excel file/`** in this folder.

### Compare, Map & Report

Two steps:

1. **Upload files** — **Excel A** and **Excel B** (required). Mapper file is optional. Click **Next: Configure**.
2. **Configure & generate** — choose **Module**, entities per sheet, then **Review** or **Generate Excel**.

Results:

- **Common** — same name/ID and same steps  
- **Unique A / Unique B** — only if that side has unmatched cases (empty unique sheets are **not** created)

---

## 7. Folders in this pack

| Folder / file | Meaning |
|---------------|---------|
| `Client doc/` | Real client `.xlsx` files (Gen UG / Life UG, etc.) |
| `Kenya doc/` | Mapper / Kenya source files for Pre-Requisites |
| `Kenya orginial testcase/` | Extra mapper sources |
| `Generated Excel file/` | Created after you generate (not in the ZIP) |
| `mapping.properties` | Local module/entity defaults |
| `start-ui.cmd` | Windows start script |

Use **`.xlsx` only** (max 25 MB per upload).

---

## 8. Common problems

| Problem | What to do |
|---------|------------|
| `node` / `npm` is not recognized | Install Node 18+, close all terminals, open a **new** PowerShell |
| `Cannot find module 'express'` | You skipped `npm install`. Run `npm install` in the unzipped folder, then `npm start` |
| Browser cannot open localhost:3100 | Make sure `npm start` is still running |
| UI loads but Compare fails / “wrong server” | Close Live Preview; use **http://localhost:3100** after `npm start` |
| Port 3100 already in use | Close other Node windows, then `npm start` again (it frees port 3100) |
| Green bot then red | The start window was closed — run `npm start` again |
| Excel will not upload | File must be `.xlsx`, not `.xls` or `.xlsm` |

---

## 9. Optional checks

From the same unzipped folder:

```powershell
npm test
```

This runs automated checks. It is **not** required to use the UI.

---

## Need more detail?

See **`README.md`** in this folder for configuration, matching rules, and CLI mapping (`npm run map`).
