const fs = require("fs");
const path = require("path");
const express = require("express");
const multer = require("multer");
const {
  isXlsxFileName,
  listWorkbookSheets,
  nowStamp,
  ensureDir,
} = require("../mapper");
const {
  parseDsrWorkbook,
  parseSheetDateSuggestion,
  validateMeetingDate,
  splitCoverTitle,
  outputPptFileName,
  formatPct,
  DATE_RE,
} = require("./parseDsr");
const { fillWeeklyPptx } = require("./fillPptx");

const XLSX_ONLY_MSG = "Only .xlsx files are supported. Please choose a .xlsx workbook.";

function createPptReportRouter(options = {}) {
  const ROOT = options.root || path.resolve(__dirname, "../..");
  const DSR_DIR = path.join(ROOT, "Weekly DSR");
  const OUTPUT_DIR = path.join(ROOT, "output");
  const LOGS_DIR = path.join(ROOT, "logs");
  const TEMPLATE_PATH = path.join(
    ROOT,
    "Template",
    "PPT Template",
    "Weekly_Workstream_Meeting_16_Sept_2026_v1.pptx"
  );

  const upload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 80 * 1024 * 1024 },
  });

  const router = express.Router();

  function withMulter(mw) {
    return (req, res, next) => {
      mw(req, res, (err) => {
        if (err) {
          err.status = err.status || 400;
          return next(err);
        }
        next();
      });
    };
  }

  function listDsrFiles() {
    ensureDir(DSR_DIR);
    const out = [];
    const walk = (dir) => {
      let names = [];
      try {
        names = fs.readdirSync(dir);
      } catch {
        return;
      }
      for (const name of names) {
        if (name.startsWith("~$")) continue;
        const full = path.join(dir, name);
        let st;
        try {
          st = fs.statSync(full);
        } catch {
          continue;
        }
        if (st.isDirectory()) walk(full);
        else if (isXlsxFileName(name)) {
          out.push({
            name,
            relative: path.relative(DSR_DIR, full).split(path.sep).join("/"),
          });
        }
      }
    };
    walk(DSR_DIR);
    return out;
  }

  function safeUnder(dir, rel) {
    const abs = path.resolve(dir, String(rel || ""));
    const root = path.resolve(dir);
    if (abs !== root && !abs.startsWith(root + path.sep)) {
      const err = new Error("File must stay inside the Weekly DSR folder.");
      err.status = 400;
      throw err;
    }
    return abs;
  }

  async function resolveWorkbook(req) {
    const uploaded = req.file || null;
    const existing = String((req.body && req.body.existingFile) || "").trim();
    ensureDir(DSR_DIR);
    let filePath;
    let fileName;
    if (uploaded) {
      if (!isXlsxFileName(uploaded.originalname)) {
        const err = new Error(`Weekly DSR: ${XLSX_ONLY_MSG}`);
        err.status = 400;
        throw err;
      }
      fileName = path.basename(uploaded.originalname);
      filePath = path.join(DSR_DIR, fileName);
      fs.writeFileSync(filePath, uploaded.buffer);
    } else if (existing) {
      filePath = safeUnder(DSR_DIR, existing);
      if (!fs.existsSync(filePath)) {
        const err = new Error("That Weekly DSR file was not found. Upload it first.");
        err.status = 400;
        throw err;
      }
      fileName = path.basename(filePath);
    } else {
      const err = new Error("Upload a Weekly DSR .xlsx or pick one from Weekly DSR.");
      err.status = 400;
      throw err;
    }
    return { filePath, fileName };
  }

  function previewPayload(parsed) {
    const mapStatus = (s) => ({
      entity: s.entity,
      title: s.title,
      executionRate: formatPct(s.executionRate),
      overallPassRate: formatPct(s.overallPassRate),
      modules: (s.modules || []).map((m) => ({
        name: m.name,
        passed: m.passed,
        failed: m.failed,
        blocked: m.blocked,
        inProgress: m.inProgress,
        notExecuted: m.notExecuted,
        total: m.total,
        passRate: formatPct(m.passRate),
      })),
      totals: s.totals
        ? {
            ...s.totals,
            passRate: formatPct(s.totals.passRate),
          }
        : null,
    });
    const mapDef = (d) => ({
      entity: d.entity,
      title: d.title,
      closureRate: formatPct(d.closureRate),
      resolutionRate: formatPct(d.resolutionRate),
      modules: d.modules,
      totals: d.totals,
    });
    return {
      sheetName: parsed.sheetName,
      suggestedDate: parsed.suggestedDate,
      alerts: parsed.alerts || [],
      status: (parsed.status || []).map(mapStatus),
      defects: (parsed.defects || []).map(mapDef),
    };
  }

  router.get("/health", (_req, res) => {
    res.json({
      ok: true,
      templateFound: fs.existsSync(TEMPLATE_PATH),
      dateExample: "16 Sept 2026",
      datePattern: String(DATE_RE),
    });
  });

  router.get("/files", (_req, res) => {
    res.json({ ok: true, files: listDsrFiles() });
  });

  router.post("/upload", withMulter(upload.single("dsr")), (req, res, next) => {
    try {
      if (!req.file) {
        return res.status(400).json({ ok: false, message: "Choose a Weekly DSR .xlsx file." });
      }
      if (!isXlsxFileName(req.file.originalname)) {
        return res.status(400).json({ ok: false, message: XLSX_ONLY_MSG });
      }
      ensureDir(DSR_DIR);
      const dest = path.join(DSR_DIR, path.basename(req.file.originalname));
      fs.writeFileSync(dest, req.file.buffer);
      const sheets = listWorkbookSheets(req.file.buffer);
      res.json({
        ok: true,
        savedAs: path.basename(dest),
        relative: path.basename(dest),
        sheets,
        suggestedDate: sheets.length === 1 ? parseSheetDateSuggestion(sheets[0]) : "",
        files: listDsrFiles(),
        message: `Saved to Weekly DSR/${path.basename(dest)}`,
      });
    } catch (err) {
      next(err);
    }
  });

  router.post("/sheets", withMulter(upload.single("dsr")), async (req, res, next) => {
    try {
      const { filePath, fileName } = await resolveWorkbook(req);
      const sheets = listWorkbookSheets(filePath);
      res.json({
        ok: true,
        fileName,
        sheets,
        suggestedDates: Object.fromEntries(
          sheets.map((s) => [s, parseSheetDateSuggestion(s)])
        ),
      });
    } catch (err) {
      next(err);
    }
  });

  router.post("/preview", withMulter(upload.single("dsr")), async (req, res, next) => {
    try {
      const { filePath, fileName } = await resolveWorkbook(req);
      const sheetName = String((req.body && req.body.sheetName) || "").trim();
      if (!sheetName) {
        return res.status(400).json({ ok: false, message: "Choose which Excel sheet to use." });
      }
      const parsed = await parseDsrWorkbook(filePath, sheetName);
      res.json({
        ok: true,
        fileName,
        ...previewPayload(parsed),
      });
    } catch (err) {
      next(err);
    }
  });

  router.post("/generate", withMulter(upload.single("dsr")), async (req, res, next) => {
    try {
      const titleChoice = String((req.body && req.body.titleChoice) || "").trim();
      const titleCustom = String((req.body && req.body.titleCustom) || "").trim();
      const title =
        titleChoice === "custom"
          ? titleCustom
          : titleChoice || "WEEKLY WORKSTREAM MEETING";
      if (!title) {
        return res.status(400).json({
          ok: false,
          message: "Enter a custom meeting title, or pick one from the dropdown.",
        });
      }
      const dateCheck = validateMeetingDate((req.body && req.body.meetingDate) || "");
      if (!dateCheck.ok) {
        return res.status(400).json({ ok: false, message: dateCheck.message });
      }
      const { filePath, fileName } = await resolveWorkbook(req);
      const sheetName = String((req.body && req.body.sheetName) || "").trim();
      if (!sheetName) {
        return res.status(400).json({ ok: false, message: "Choose which Excel sheet to use." });
      }
      if (!fs.existsSync(TEMPLATE_PATH)) {
        return res.status(400).json({
          ok: false,
          message: "PPT template is missing from Template/PPT Template.",
        });
      }

      const parsed = await parseDsrWorkbook(filePath, sheetName);
      const outName = outputPptFileName(title, dateCheck.value);
      ensureDir(OUTPUT_DIR);
      const outputPath = path.join(OUTPUT_DIR, outName);
      const filled = await fillWeeklyPptx({
        templatePath: TEMPLATE_PATH,
        outputPath,
        parsed,
        title,
        dateText: dateCheck.value,
      });

      const logLines = [
        `Weekly PPT Report  ${nowStamp()}`,
        `Title: ${title}`,
        `Date: ${dateCheck.value}`,
        `Excel: ${fileName} / ${parsed.sheetName}`,
        `Output: output/${outName}`,
        `Slide 2 Kenya left as in the template.`,
        "",
        ...(parsed.alerts || []).map((a) => `ALERT: ${a}`),
        ...(filled.warnings || []).map((a) => `NOTE: ${a}`),
      ];
      ensureDir(LOGS_DIR);
      const logName = `Weekly PPT - ${nowStamp()}.log.txt`;
      const logPath = path.join(LOGS_DIR, logName);
      fs.writeFileSync(logPath, logLines.join("\n"), "utf8");

      res.json({
        ok: true,
        message: "Weekly PPT generated.",
        fileName: outName,
        outputFile: `output/${outName}`,
        logFile: `logs/${logName}`,
        download: {
          ppt: `/api/ppt/download?file=${encodeURIComponent(outName)}`,
          log: `/api/ppt/download-log?file=${encodeURIComponent(logName)}`,
        },
        alerts: [...(parsed.alerts || []), ...(filled.warnings || [])],
        preview: previewPayload(parsed),
        cover: splitCoverTitle(title),
        logLines,
      });
    } catch (err) {
      next(err);
    }
  });

  router.get("/download", (req, res) => {
    try {
      const name = path.basename(String((req.query && req.query.file) || ""));
      const abs = path.join(OUTPUT_DIR, name);
      if (!name || !fs.existsSync(abs)) {
        return res.status(404).json({ ok: false, message: "PPT file was not found." });
      }
      res.download(abs, name);
    } catch (err) {
      res.status(400).json({ ok: false, message: err.message || String(err) });
    }
  });

  router.get("/download-log", (req, res) => {
    try {
      const name = path.basename(String((req.query && req.query.file) || ""));
      const abs = path.join(LOGS_DIR, name);
      if (!name || !fs.existsSync(abs)) {
        return res.status(404).json({ ok: false, message: "Log file was not found." });
      }
      res.download(abs, name);
    } catch (err) {
      res.status(400).json({ ok: false, message: err.message || String(err) });
    }
  });

  return router;
}

module.exports = { createPptReportRouter };
