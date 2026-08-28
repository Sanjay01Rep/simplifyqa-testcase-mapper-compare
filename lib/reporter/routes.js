const fs = require("fs");
const path = require("path");
const os = require("os");
const express = require("express");
const ExcelJS = require("exceljs");
const multer = require("multer");
const {
  ROOT,
  PROPERTIES_PATH,
  OUTPUT_DIR,
  LOGS_DIR,
  TEMPLATE_META,
  ReporterError,
  loadProperties,
  formDefaultsFromProps,
  fetchSimplifyQaProjects,
  applyFormOverrides,
  executeReport,
  extractAlerts,
  discoverStatusSections,
  getTokenStatus,
  saveBearerToken,
} = require("./reporter");
const { getBearerToken } = require("../loadEnv");
const {
  readBranding,
  readScheduleConfig,
  readNotifyConfig,
} = require("./config-helpers");
const { notifyRunResult } = require("./notify");
const { createScheduler } = require("./scheduler");
const { listWorkbookSheets, writeCompareSheet } = require("./compare");
const { generateReportPdf } = require("./pdf-report");

const HISTORY_PATH = path.join(LOGS_DIR, "run-history.json");
const HISTORY_LIMIT = 12;

let running = false;
let runProgress = {
  active: false,
  stepId: null,
  label: "",
  steps: [],
  startedAt: null,
};
let scheduler = null;

function setProgress(step) {
  if (!step) return;
  runProgress = {
    ...runProgress,
    active: true,
    stepId: step.id,
    label: step.label || "",
    steps: [...(runProgress.steps || []), step].slice(-20),
  };
}

function readHistory() {
  try {
    if (!fs.existsSync(HISTORY_PATH)) return [];
    const raw = JSON.parse(fs.readFileSync(HISTORY_PATH, "utf8"));
    return Array.isArray(raw) ? raw : [];
  } catch {
    return [];
  }
}

function pushHistory(entry) {
  try {
    if (!fs.existsSync(LOGS_DIR)) fs.mkdirSync(LOGS_DIR, { recursive: true });
    const list = readHistory();
    list.unshift(entry);
    fs.writeFileSync(
      HISTORY_PATH,
      JSON.stringify(list.slice(0, HISTORY_LIMIT), null, 2),
      "utf8"
    );
  } catch (err) {
    console.warn("Could not write run history:", err.message);
  }
}

function currentProps() {
  return loadProperties(PROPERTIES_PATH);
}

function brandingFromDisk() {
  return readBranding(currentProps(), ROOT);
}

function relativeToRoot(absPath) {
  if (!absPath) return null;
  return path.relative(ROOT, absPath).split(path.sep).join("/");
}

function safeResolveUnder(baseDir, nameOrRel) {
  const cleaned = String(nameOrRel || "")
    .replace(/\\/g, "/")
    .replace(/^\/+/, "");
  if (!cleaned || cleaned.includes("..")) {
    const err = new Error("Invalid file path.");
    err.status = 400;
    throw err;
  }
  const abs = path.resolve(baseDir, cleaned);
  const base = path.resolve(baseDir);
  if (abs !== base && !abs.startsWith(base + path.sep)) {
    const err = new Error("File is outside the allowed folder.");
    err.status = 400;
    throw err;
  }
  return abs;
}

function previewCellText(value) {
  if (value == null || value === "") return "";
  if (typeof value === "number") {
    if (Number.isFinite(value) && value !== 0 && Math.abs(value) <= 1) {
      return `${Math.round(value * 1000) / 10}%`;
    }
    return String(value);
  }
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "object") {
    if (value.result != null && value.result !== "") {
      return previewCellText(value.result);
    }
    if (value.richText) return value.richText.map((p) => p.text).join("");
    if (value.text != null) return String(value.text);
    if (value.formula != null) return "";
    if (value.sharedFormula != null) return "";
    if (value.error) return String(value.error);
  }
  return String(value);
}

function isPercentNumFmt(numFmt) {
  return typeof numFmt === "string" && numFmt.includes("%");
}

async function buildWorkbookPreview(absPath, preferredSheet) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(absPath);
  const sheetNames = workbook.worksheets.map((ws) => ws.name);
  if (!sheetNames.length) {
    return { sheets: [], sheet: null, rows: [], maxCol: 0 };
  }

  let sheet =
    (preferredSheet &&
      workbook.worksheets.find((ws) => ws.name === preferredSheet)) ||
    null;
  if (!sheet) {
    sheet = workbook.worksheets[workbook.worksheets.length - 1];
  }

  let maxCol = 0;
  sheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (_cell, colNumber) => {
      if (colNumber > maxCol) maxCol = colNumber;
    });
  });
  maxCol = Math.min(Math.max(maxCol, 1), 20);

  const rows = [];
  const lastRow = Math.min(sheet.rowCount || 0, 120);
  for (let r = 1; r <= lastRow; r++) {
    const row = sheet.getRow(r);
    const cells = [];
    let rowHasContent = false;
    for (let c = 1; c <= maxCol; c++) {
      const cell = row.getCell(c);
      let text = previewCellText(cell.value);
      if (
        isPercentNumFmt(cell.numFmt) &&
        typeof cell.value === "object" &&
        typeof cell.value.result === "number"
      ) {
        text = `${Math.round(cell.value.result * 1000) / 10}%`;
      } else if (
        isPercentNumFmt(cell.numFmt) &&
        typeof cell.value === "number"
      ) {
        text = `${Math.round(cell.value * 1000) / 10}%`;
      }
      if (text) rowHasContent = true;
      cells.push({ text });
    }
    if (rowHasContent) rows.push({ row: r, cells });
  }

  return {
    sheets: sheetNames,
    sheet: sheet.name,
    rows,
    maxCol,
  };
}

function toFileUrl(absPath) {
  const resolved = path.resolve(absPath);
  if (resolved.startsWith("\\\\")) {
    const unc = resolved.replace(/\\/g, "/").replace(/^\/\//, "");
    return `file://${unc.split("/").map(encodeURIComponent).join("/")}`;
  }
  const normalized = resolved.replace(/\\/g, "/");
  return `file:///${normalized
    .split("/")
    .map((part, i) => (i === 0 ? part : encodeURIComponent(part)))
    .join("/")}`;
}

function resolveOpenPath(localAbsPath) {
  let shareRoot = "";
  try {
    if (fs.existsSync(PROPERTIES_PATH)) {
      const props = loadProperties(PROPERTIES_PATH);
      shareRoot = String(props.OUTPUT_SHARE_UNC || process.env.OUTPUT_SHARE_UNC || "").trim();
    } else {
      shareRoot = String(process.env.OUTPUT_SHARE_UNC || "").trim();
    }
  } catch {
    shareRoot = String(process.env.OUTPUT_SHARE_UNC || "").trim();
  }

  if (shareRoot) {
    const fileName = path.basename(localAbsPath);
    const unc = shareRoot.replace(/[/\\]+$/, "") + "\\" + fileName;
    return { openPath: unc, source: "unc" };
  }
  return { openPath: localAbsPath, source: "local" };
}

function buildExcelUri(absOrUncPath) {
  return `ms-excel:ofe|u|${encodeURIComponent(toFileUrl(absOrUncPath))}`;
}

function isLocalClient(req) {
  const raw = String(req.ip || req.socket.remoteAddress || "")
    .replace(/^::ffff:/, "")
    .trim();
  if (raw === "127.0.0.1" || raw === "::1" || raw === "localhost") return true;
  try {
    if (lanAddresses().includes(raw)) return true;
  } catch {
    /* ignore */
  }
  return false;
}

function lanAddresses() {
  const nets = os.networkInterfaces();
  const result = [];
  for (const entries of Object.values(nets)) {
    for (const net of entries || []) {
      if (net.family === "IPv4" && !net.internal) result.push(net.address);
    }
  }
  return result;
}

function escapeHtml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderExcelViewerHtml({ fileName, preview, downloadUrl }) {
  const sheets = preview.sheets || [];
  const current = preview.sheet || "";
  const tabs = sheets
    .map((name) => {
      const href = `/api/reporter/view/output?file=${encodeURIComponent(fileName)}&sheet=${encodeURIComponent(name)}`;
      const active = name === current ? " active" : "";
      return `<a class="sheet-tab${active}" href="${href}">${escapeHtml(name)}</a>`;
    })
    .join("");

  const maxCol = preview.maxCol || 0;
  let table = "<tr><th></th>";
  for (let c = 1; c <= maxCol; c++) {
    table += `<th>${String.fromCharCode(64 + Math.min(c, 26))}</th>`;
  }
  table += "</tr>";

  for (const row of preview.rows || []) {
    const texts = (row.cells || []).map((c) =>
      String(c.text || "").trim().toLowerCase()
    );
    const joined = texts.join(" ");
    let cls = "";
    if (texts.some((t) => t.includes("module-wise daily status"))) cls = "section-title";
    else if (texts.includes("module") || texts.includes("lead champion")) cls = "header-row";
    else if (
      texts.includes("total") ||
      joined.includes("execution rate") ||
      joined.includes("overall pass rate")
    ) {
      cls = "total-row";
    }

    table += `<tr class="${cls}"><td class="row-num">${row.row}</td>`;
    for (const cell of row.cells || []) {
      table += `<td>${escapeHtml(cell.text || "")}</td>`;
    }
    table += "</tr>";
  }

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${escapeHtml(fileName)} — ${escapeHtml(current)}</title>
  <style>
    :root { --ink:#132029; --muted:#4d6270; --brand:#0b5f6b; --line:#b7c7d2; --bg:#e8eef2; }
    body { margin:0; font-family:"Segoe UI", system-ui, sans-serif; color:var(--ink); background:var(--bg); }
    header { padding:1rem 1.25rem; background:#fff; border-bottom:1px solid var(--line); }
    h1 { margin:0; font-size:1.15rem; color:var(--brand); }
    .meta { margin:0.35rem 0 0; color:var(--muted); font-size:0.9rem; }
    .actions { margin-top:0.75rem; display:flex; flex-wrap:wrap; gap:0.75rem; }
    .actions a { color:var(--brand); font-weight:600; }
    .tabs { display:flex; flex-wrap:wrap; gap:0.4rem; padding:0.85rem 1.25rem 0; }
    .sheet-tab { text-decoration:none; border:1px solid var(--line); background:#fff; color:var(--ink); padding:0.4rem 0.7rem; font-size:0.86rem; font-weight:600; }
    .sheet-tab.active { background:var(--brand); border-color:var(--brand); color:#fff; }
    .wrap { margin:0.85rem 1.25rem 1.5rem; overflow:auto; max-height:calc(100vh - 10rem); border:1px solid var(--line); background:#fff; }
    table { border-collapse:collapse; width:max-content; min-width:100%; font-size:0.82rem; }
    th, td { border:1px solid #d5dee5; padding:0.35rem 0.55rem; white-space:nowrap; vertical-align:top; }
    th { background:#eef4f7; color:var(--muted); position:sticky; top:0; z-index:1; }
    td.row-num { background:#f5f8fa; color:var(--muted); text-align:right; font-family:ui-monospace,monospace; font-size:0.75rem; position:sticky; left:0; }
    tr.section-title td:not(.row-num) { background:#e7f3f4; font-weight:700; color:#084851; }
    tr.header-row td:not(.row-num) { background:#f3f7f9; font-weight:600; }
    tr.total-row td:not(.row-num) { background:#fff7ef; font-weight:600; }
  </style>
</head>
<body>
  <header>
    <h1>${escapeHtml(fileName)}</h1>
    <p class="meta">Sheet: ${escapeHtml(current)}</p>
    <div class="actions">
      <a href="${downloadUrl}">Download Excel report</a>
      <a href="/">Back to app</a>
    </div>
  </header>
  <div class="tabs">${tabs}</div>
  <div class="wrap"><table>${table}</table></div>
</body>
</html>`;
}

function initScheduler() {
  if (scheduler) return scheduler;
  scheduler = createScheduler({
    getConfig: () => readScheduleConfig(currentProps()),
    onFire: async () => {
      const result = await executeReport({
        onProgress: setProgress,
      });
      const alerts = result.alerts || extractAlerts(result.logLines || []);
      const branding = brandingFromDisk();
      const notifyConfig = readNotifyConfig(currentProps());
      await notifyRunResult({
        ok: true,
        notifyConfig,
        title: `[Scheduled] ${branding.title}`,
        summaryLines: [
          `Template: ${result.runtime && result.runtime.templateChoice}`,
          `Sheet: ${result.sheetName || ""}`,
          `Output: ${result.outputFile || ""}`,
          alerts.length ? `Alerts: ${alerts.length}` : "Alerts: none",
        ],
        excelPath: result.outputFile,
        pdfPath: null,
        logPath: result.logPath,
        log: console.log,
      });
      return { message: `Report generated: ${path.basename(result.outputFile)}` };
    },
    log: console.log,
  });
  scheduler.start();
  return scheduler;
}

function registerTemplateInProperties(propsPath, index, templateRelative, outputRelative) {
  let text = fs.existsSync(propsPath) ? fs.readFileSync(propsPath, "utf8") : "";
  const tKey = `TEMPLATE_${index}`;
  const oKey = `OUTPUT_FILE_${index}`;

  if (text.includes(`${tKey}=`)) {
    text = text.replace(new RegExp(`^${tKey}=.*$`, "m"), `${tKey}=${templateRelative}`);
  } else {
    const lastMatch = text.match(/OUTPUT_FILE_\d+=.*$/m);
    if (lastMatch) {
      const idx = text.indexOf(lastMatch[0]) + lastMatch[0].length;
      text = text.slice(0, idx) + `\n\n${tKey}=${templateRelative}\n${oKey}=${outputRelative}` + text.slice(idx);
    } else {
      text += `\n${tKey}=${templateRelative}\n${oKey}=${outputRelative}\n`;
    }
  }

  if (text.includes(`${oKey}=`)) {
    text = text.replace(new RegExp(`^${oKey}=.*$`, "m"), `${oKey}=${outputRelative}`);
  }

  if (text.includes("TEMPLATE_CHOICE=")) {
    text = text.replace(/^TEMPLATE_CHOICE=.*$/m, `TEMPLATE_CHOICE=${index}`);
  } else {
    text += `\nTEMPLATE_CHOICE=${index}\n`;
  }

  fs.writeFileSync(propsPath, text.replace(/\r?\n/g, "\n"), "utf8");
}

function createReporterRouter() {
  const router = express.Router();
  const templateUpload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 40 * 1024 * 1024 },
  });
  initScheduler();

  router.get("/health", (_req, res) => {
    res.json({ ok: true, running, progress: runProgress });
  });

  router.get("/run-progress", (_req, res) => {
    res.json({ ok: true, running, progress: runProgress });
  });

  router.get("/history", (_req, res) => {
    res.json({ ok: true, runs: readHistory() });
  });

  router.get("/branding", (_req, res) => {
    try {
      const branding = brandingFromDisk();
      res.json({
        ok: true,
        title: branding.title,
        tagline: branding.tagline,
        logoRightLabel: branding.logoRightLabel,
        logoLeft: {
          url: branding.logoLeft.url,
          relative: branding.logoLeft.relative,
          exists: branding.logoLeft.exists,
        },
        logoRight: {
          url: branding.logoRight.url,
          relative: branding.logoRight.relative,
          exists: branding.logoRight.exists,
        },
      });
    } catch (err) {
      res.status(400).json({ ok: false, message: err.message || String(err) });
    }
  });

  router.get("/schedule", (_req, res) => {
    res.json({
      ok: true,
      schedule: scheduler ? scheduler.status() : { enabled: false, times: [] },
    });
  });

  router.get("/sheets", async (req, res) => {
    try {
      let abs;
      if (req.query.file) {
        abs = safeResolveUnder(OUTPUT_DIR, req.query.file);
      } else {
        const props = currentProps();
        const choice = String(props.TEMPLATE_CHOICE || "1").trim();
        const outRel =
          props[`OUTPUT_FILE_${choice}`] || "output/FMS Status tracker.xlsx";
        abs = path.isAbsolute(outRel) ? outRel : path.join(ROOT, outRel);
      }
      if (!fs.existsSync(abs)) {
        return res.json({ ok: true, file: null, sheets: [] });
      }
      const sheets = await listWorkbookSheets(abs);
      res.json({
        ok: true,
        file: path.basename(abs),
        sheets,
      });
    } catch (err) {
      res
        .status(err.status || 400)
        .json({ ok: false, message: err.message || String(err) });
    }
  });

  router.post("/compare", async (req, res) => {
    try {
      const file = (req.body && req.body.file) || null;
      const previous = req.body && req.body.previous;
      const latest = req.body && req.body.latest;
      if (!previous || !latest) {
        return res.status(400).json({
          ok: false,
          message: "Select previous and latest sheet names to compare.",
        });
      }
      if (previous === latest) {
        return res.status(400).json({
          ok: false,
          message: "Previous and latest sheets must be different.",
        });
      }
      let abs;
      if (file) {
        abs = safeResolveUnder(OUTPUT_DIR, file);
      } else {
        const props = currentProps();
        const choice = String(props.TEMPLATE_CHOICE || "1").trim();
        const outRel = props[`OUTPUT_FILE_${choice}`] || `output/FMS Status tracker.xlsx`;
        abs = path.isAbsolute(outRel) ? outRel : path.join(ROOT, outRel);
      }
      if (!fs.existsSync(abs)) {
        return res.status(404).json({ ok: false, message: "Output workbook not found." });
      }
      const result = await writeCompareSheet(abs, previous, latest);
      res.json({
        ok: true,
        message: `Compare sheet created: ${result.sheetName}`,
        ...result,
        file: path.basename(abs),
        download: `/api/reporter/download/output?file=${encodeURIComponent(path.basename(abs))}`,
      });
    } catch (err) {
      res.status(400).json({ ok: false, message: err.message || String(err) });
    }
  });

  router.post("/pdf", async (req, res) => {
    try {
      const file = (req.body && req.body.file) || null;
      const sheetName = (req.body && req.body.sheetName) || null;
      let abs;
      if (file) {
        abs = safeResolveUnder(OUTPUT_DIR, file);
      } else {
        const props = currentProps();
        const choice = String(props.TEMPLATE_CHOICE || "1").trim();
        const outRel = props[`OUTPUT_FILE_${choice}`] || `output/FMS Status tracker.xlsx`;
        abs = path.isAbsolute(outRel) ? outRel : path.join(ROOT, outRel);
      }
      if (!fs.existsSync(abs)) {
        return res.status(404).json({ ok: false, message: "Output workbook not found." });
      }
      const branding = brandingFromDisk();
      const pdfName = path.basename(abs).replace(/\.xlsx$/i, "") + ".pdf";
      const pdfPath = path.join(OUTPUT_DIR, pdfName);
      await generateReportPdf({
        excelPath: abs,
        sheetName,
        outPath: pdfPath,
        branding,
      });
      res.json({
        ok: true,
        message: "PDF generated.",
        file: pdfName,
        download: `/api/reporter/download/pdf?file=${encodeURIComponent(pdfName)}`,
      });
    } catch (err) {
      res.status(400).json({ ok: false, message: err.message || String(err) });
    }
  });

  router.get("/download/pdf", (req, res) => {
    try {
      const abs = safeResolveUnder(OUTPUT_DIR, req.query.file);
      if (!fs.existsSync(abs)) {
        return res.status(404).json({ ok: false, message: "PDF file not found." });
      }
      res.download(abs, path.basename(abs));
    } catch (err) {
      res.status(err.status || 400).json({ ok: false, message: err.message });
    }
  });

  router.get("/download/output", (req, res) => {
    try {
      const abs = safeResolveUnder(OUTPUT_DIR, req.query.file);
      if (!fs.existsSync(abs)) {
        return res.status(404).json({ ok: false, message: "Output file not found." });
      }
      res.download(abs, path.basename(abs));
    } catch (err) {
      res.status(err.status || 400).json({ ok: false, message: err.message });
    }
  });

  router.get("/download/log", (req, res) => {
    try {
      const abs = safeResolveUnder(LOGS_DIR, req.query.file);
      if (!fs.existsSync(abs)) {
        return res.status(404).json({ ok: false, message: "Log file not found." });
      }
      res.download(abs, path.basename(abs));
    } catch (err) {
      res.status(err.status || 400).json({ ok: false, message: err.message });
    }
  });

  router.post("/upload-template", templateUpload.single("template"), async (req, res) => {
    try {
      const file = req.file;
      if (!file || !file.buffer) {
        return res.status(400).json({
          ok: false,
          message: "Please choose an Excel template (.xlsx) file to upload.",
        });
      }
      if (!file.originalname.toLowerCase().endsWith(".xlsx")) {
        return res.status(400).json({
          ok: false,
          message: "Only .xlsx template files are supported.",
        });
      }

      const templateDir = path.join(ROOT, "Template");
      if (!fs.existsSync(templateDir)) {
        fs.mkdirSync(templateDir, { recursive: true });
      }

      const safeFileName = path.basename(file.originalname);
      const targetPath = path.join(templateDir, safeFileName);
      fs.writeFileSync(targetPath, file.buffer);

      const stem = path.parse(safeFileName).name;
      const templateRelative = `Template/${safeFileName}`;
      const outputRelative = `output/${stem}.xlsx`;

      const props = loadProperties(PROPERTIES_PATH);

      // Check if template path is already registered under any index
      let targetIndex = null;
      for (let i = 1; i <= 50; i++) {
        if (props[`TEMPLATE_${i}`] === templateRelative) {
          targetIndex = i;
          break;
        }
      }

      // If not found, find first unused index
      if (!targetIndex) {
        for (let i = 1; i <= 50; i++) {
          if (!props[`TEMPLATE_${i}`] && !props[`OUTPUT_FILE_${i}`]) {
            targetIndex = i;
            break;
          }
        }
      }
      if (!targetIndex) targetIndex = 4;

      // Update application.properties
      registerTemplateInProperties(PROPERTIES_PATH, targetIndex, templateRelative, outputRelative);

      // Auto-inspect status sections in uploaded template
      try {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(file.buffer);
        const ws = wb.worksheets[0];
        if (ws && typeof discoverStatusSections === "function") {
          const sections = discoverStatusSections(ws);
          TEMPLATE_META[targetIndex] = {
            statusSections: sections.length,
            description: `${sections.length} execution-plan block(s).`,
          };
        }
      } catch {}

      const updatedProps = loadProperties(PROPERTIES_PATH);
      res.json({
        ok: true,
        message: `Template "${safeFileName}" uploaded successfully as Template ${targetIndex}.`,
        choice: String(targetIndex),
        templateName: safeFileName,
        form: formDefaultsFromProps(updatedProps),
      });
    } catch (err) {
      res.status(500).json({ ok: false, message: err.message || String(err) });
    }
  });

  router.get("/projects", async (_req, res) => {
    try {
      const token = getBearerToken(ROOT);
      const projects = await fetchSimplifyQaProjects(token);
      res.json({ ok: true, projects });
    } catch (err) {
      res.status(500).json({ ok: false, message: err.message || String(err), projects: [] });
    }
  });

  router.get("/form-defaults", async (_req, res) => {
    try {
      const props = loadProperties(PROPERTIES_PATH);
      const form = formDefaultsFromProps(props);
      const token = getBearerToken(ROOT);
      try {
        form.projects = await fetchSimplifyQaProjects(token);
      } catch {}
      res.json({ ok: true, form });
    } catch (err) {
      res.status(400).json({
        ok: false,
        message: err.message || String(err),
        code: err.code || "CONFIG_ERROR",
      });
    }
  });

  router.get("/properties", (_req, res) => {
    try {
      if (!fs.existsSync(PROPERTIES_PATH)) {
        return res.status(404).json({
          ok: false,
          message: "config/application.properties was not found.",
        });
      }
      const text = fs.readFileSync(PROPERTIES_PATH, "utf8");
      const props = loadProperties(PROPERTIES_PATH);
      res.json({
        ok: true,
        path: relativeToRoot(PROPERTIES_PATH),
        text,
        form: formDefaultsFromProps(props),
      });
    } catch (err) {
      res.status(400).json({
        ok: false,
        message: err.message || String(err),
        code: err.code || "CONFIG_ERROR",
      });
    }
  });

  router.put("/properties", (req, res) => {
    try {
      const text = req.body && typeof req.body.text === "string" ? req.body.text : null;
      if (text == null) {
        return res.status(400).json({
          ok: false,
          message: "Request body must include text (full properties file content).",
        });
      }
      fs.mkdirSync(path.dirname(PROPERTIES_PATH), { recursive: true });
      fs.writeFileSync(PROPERTIES_PATH, text.replace(/\r?\n/g, "\n"), "utf8");
      const props = loadProperties(PROPERTIES_PATH);
      res.json({
        ok: true,
        message: "application.properties saved.",
        form: formDefaultsFromProps(props),
      });
    } catch (err) {
      res.status(400).json({
        ok: false,
        message: err.message || String(err),
        code: err.code || "CONFIG_ERROR",
      });
    }
  });

  router.post("/run", async (req, res) => {
    if (running) {
      return res.status(409).json({
        ok: false,
        message: "A report is already running. Please wait for it to finish.",
      });
    }

    running = true;
    runProgress = {
      active: true,
      stepId: "prepare",
      label: "Starting…",
      steps: [],
      startedAt: new Date().toISOString(),
    };

    try {
      const form = (req.body && req.body.form) || {};
      const includePdf = form.includePdf === true;
      const result = await executeReport({
        form,
        onProgress: setProgress,
      });
      const outputName = result.outputFile
        ? path.basename(result.outputFile)
        : null;
      let preview = null;
      if (result.outputFile && fs.existsSync(result.outputFile)) {
        try {
          preview = await buildWorkbookPreview(
            result.outputFile,
            result.sheetName || null
          );
        } catch (previewErr) {
          preview = {
            error: previewErr.message || "Unable to build Excel preview.",
          };
        }
      }

      let excelOpen = null;
      if (result.outputFile && fs.existsSync(result.outputFile)) {
        const { openPath, source } = resolveOpenPath(result.outputFile);
        excelOpen = {
          localPath: result.outputFile,
          openPath,
          source,
          fileUrl: toFileUrl(openPath),
          excelUri: buildExcelUri(openPath),
          canLaunchLocal: isLocalClient(req),
          launchUrl: outputName
            ? `/api/launch-excel?file=${encodeURIComponent(path.join("output", outputName))}`
            : null,
        };
      }

      let pdfInfo = null;
      if (includePdf && result.outputFile && fs.existsSync(result.outputFile)) {
        try {
          const branding = brandingFromDisk();
          const pdfName =
            path.basename(result.outputFile).replace(/\.xlsx$/i, "") + ".pdf";
          const pdfPath = path.join(OUTPUT_DIR, pdfName);
          await generateReportPdf({
            excelPath: result.outputFile,
            sheetName: result.sheetName,
            outPath: pdfPath,
            branding,
          });
          pdfInfo = {
            file: pdfName,
            download: `/api/reporter/download/pdf?file=${encodeURIComponent(pdfName)}`,
            path: pdfPath,
          };
        } catch (pdfErr) {
          pdfInfo = { error: pdfErr.message || String(pdfErr) };
        }
      }

      const alerts = result.alerts || extractAlerts(result.logLines || []);
      const branding = brandingFromDisk();
      const props = currentProps();
      const notifyConfig = readNotifyConfig(props);
      await notifyRunResult({
        ok: true,
        notifyConfig,
        title: branding.title,
        summaryLines: [
          `Template: ${result.runtime && result.runtime.templateChoice}`,
          `Sheet: ${result.sheetName || ""}`,
          `Output: ${result.outputFile || ""}`,
          alerts.length ? `Alerts: ${alerts.length}` : "Alerts: none",
        ],
        excelPath: result.outputFile,
        pdfPath: pdfInfo && pdfInfo.path,
        logPath: result.logPath,
        log: console.log,
      });

      const payload = {
        ok: true,
        message: "Report generated successfully.",
        outputFile: relativeToRoot(result.outputFile),
        logFile: relativeToRoot(result.logPath),
        sheetName: result.sheetName,
        runtime: result.runtime,
        logLines: result.logLines,
        alerts,
        preview,
        excelOpen,
        pdf: pdfInfo,
        download: {
          output: outputName
            ? `/api/reporter/download/output?file=${encodeURIComponent(outputName)}`
            : null,
          log: result.logPath
            ? `/api/reporter/download/log?file=${encodeURIComponent(path.basename(result.logPath))}`
            : null,
          pdf: pdfInfo && pdfInfo.download ? pdfInfo.download : null,
        },
        openInExcel: excelOpen ? excelOpen.excelUri : null,
        view: {
          log: result.logPath
            ? `/api/reporter/view/log?file=${encodeURIComponent(path.basename(result.logPath))}`
            : null,
        },
      };

      pushHistory({
        at: new Date().toISOString(),
        ok: true,
        templateChoice: result.runtime && result.runtime.templateChoice,
        plans: (result.runtime && result.runtime.plans) || [],
        sheetName: result.sheetName,
        outputFile: payload.outputFile,
        logFile: payload.logFile,
        download: payload.download,
        alertCount: alerts.length,
      });

      res.json(payload);
    } catch (err) {
      const alerts = err.alerts || extractAlerts(err.logLines || []);
      const failPayload = {
        ok: false,
        message: err.message || String(err),
        code: err.code || "RUN_FAILED",
        details: err.details || "",
        outputFile: relativeToRoot(err.outputFile),
        logFile: relativeToRoot(err.logPath),
        logLines: err.logLines || [],
        alerts,
        download: {
          output: err.outputFile
            ? `/api/reporter/download/output?file=${encodeURIComponent(path.basename(err.outputFile))}`
            : null,
          log: err.logPath
            ? `/api/reporter/download/log?file=${encodeURIComponent(path.basename(err.logPath))}`
            : null,
        },
      };
      try {
        const branding = brandingFromDisk();
        const notifyConfig = readNotifyConfig(currentProps());
        await notifyRunResult({
          ok: false,
          notifyConfig,
          title: branding.title,
          summaryLines: [failPayload.message, failPayload.code].filter(Boolean),
          excelPath: null,
          pdfPath: null,
          logPath: err.logPath,
          log: console.log,
        });
      } catch {
        /* ignore notify errors */
      }
      pushHistory({
        at: new Date().toISOString(),
        ok: false,
        message: failPayload.message,
        code: failPayload.code,
        outputFile: failPayload.outputFile,
        logFile: failPayload.logFile,
        download: failPayload.download,
        alertCount: alerts.length,
      });
      res.status(400).json(failPayload);
    } finally {
      running = false;
      runProgress = { ...runProgress, active: false };
    }
  });

  router.get("/preview/output", async (req, res) => {
    try {
      const abs = safeResolveUnder(OUTPUT_DIR, req.query.file);
      if (!fs.existsSync(abs)) {
        return res.status(404).json({ ok: false, message: "Output file not found." });
      }
      const preview = await buildWorkbookPreview(abs, req.query.sheet || null);
      res.json({
        ok: true,
        file: path.basename(abs),
        ...preview,
      });
    } catch (err) {
      res.status(err.status || 400).json({
        ok: false,
        message: err.message || String(err),
      });
    }
  });

  router.get("/view/output", async (req, res) => {
    try {
      const abs = safeResolveUnder(OUTPUT_DIR, req.query.file);
      if (!fs.existsSync(abs)) {
        return res.status(404).send("Output workbook not found.");
      }
      const preview = await buildWorkbookPreview(abs, req.query.sheet || null);
      const fileName = path.basename(abs);
      const downloadUrl = `/api/reporter/download/output?file=${encodeURIComponent(fileName)}`;
      res.setHeader("Content-Type", "text/html; charset=utf-8");
      res.send(renderExcelViewerHtml({ fileName, preview, downloadUrl }));
    } catch (err) {
      res.status(err.status || 400).send(err.message || String(err));
    }
  });

  router.get("/view/log", (req, res) => {
    try {
      const abs = safeResolveUnder(LOGS_DIR, req.query.file);
      if (!fs.existsSync(abs)) {
        return res.status(404).send("Log file not found.");
      }
      const text = fs.readFileSync(abs, "utf8");
      res.setHeader("Content-Type", "text/plain; charset=utf-8");
      res.send(text);
    } catch (err) {
      res.status(err.status || 400).send(err.message || String(err));
    }
  });

  return router;
}

module.exports = {
  createReporterRouter,
  OUTPUT_DIR,
  LOGS_DIR,
};
