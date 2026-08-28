const fs = require("fs");
const http = require("http");
const path = require("path");
const os = require("os");
const { execSync } = require("child_process");
const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const {
  loadProperties,
  loadJobs,
  resolveJobConfig,
  isLockFile,
  nowStamp,
  safeBaseName,
  ensureDir,
  mapFromBuffers,
  writeWorkbook,
  writeLog,
  isXlsxFileName,
  listMapperHeaders,
  listWorkbookSheets,
} = require("./lib/mapper");
const { ENTITIES, MODULES } = require("./lib/options");
const { compareClientDocs, writeCompareLog } = require("./lib/compare");
const {
  EP_HEADER,
  SUPPORTED_DATE_FORMATS,
  validateAndFormatDate,
  sanitizeSheetName,
  extractTestcasesFromSummary,
  fetchTestcasesFromSimplifyQa,
  buildEpWorkbook,
} = require("./lib/epMapper");
const {
  loadProjectEnv,
  getBearerToken,
  getTokenStatus,
  saveBearerToken,
} = require("./lib/loadEnv");
const { createReporterRouter } = require("./lib/reporter/routes");

const ROOT = __dirname;
const explicitPort = process.env.PORT;
loadProjectEnv(ROOT);
if (explicitPort) process.env.PORT = explicitPort;

const PKG = require("./package.json");
const APP_VERSION = String(PKG.version || "0.0.0");
const APP_NAME = String(PKG.name || "simplifyqa-testcase-mapper-compare");

const CLIENT_DIR = path.join(ROOT, "Client doc");
const KENYA_DIR = path.join(ROOT, "Kenya doc");
const KENYA_ORIGINAL_DIR = path.join(ROOT, "Kenya orginial testcase");
const KENYA_DIRS = [KENYA_DIR, KENYA_ORIGINAL_DIR];
const EP_SAMPLE_DIR = path.join(ROOT, "Execution Plan Sample file");
const XLSX_ONLY_MSG = "Only .xlsx files are supported. Please choose a .xlsx workbook.";
const OUT_DIR = path.join(ROOT, "Generated Excel file");
const LOG_DIR = path.join(OUT_DIR, "logs");
const PROPS_PATH = path.join(ROOT, "mapping.properties");
const PROPS_EXAMPLE_PATH = path.join(ROOT, "mapping.properties.example");
const JOBS_PATH = path.join(ROOT, "jobs.json");
const HISTORY_PATH = path.join(LOG_DIR, "run-history.json");
const PORT = Number(process.env.PORT || 3100);
/** UI status-bot poll interval (ms). Default: 5 minutes. */
const HEALTH_POLL_MS = Math.max(
  5000,
  Number(process.env.HEALTH_POLL_MS || 5 * 60 * 1000) || 5 * 60 * 1000
);

/** First clone: create mapping.properties from the committed example (never overwrites). */
function ensureMappingProperties() {
  if (fs.existsSync(PROPS_PATH)) return;
  if (!fs.existsSync(PROPS_EXAMPLE_PATH)) return;
  fs.copyFileSync(PROPS_EXAMPLE_PATH, PROPS_PATH);
  console.log(`Created mapping.properties from mapping.properties.example`);
}
ensureMappingProperties();
const HISTORY_LIMIT = 8;

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 },
});

app.use(express.json({ limit: "4mb" }));

app.use((req, res, next) => {
  res.setHeader("X-ICEA-Lion", "testcase-review");
  const start = Date.now();
  res.on("finish", () => {
    const url = String(req.originalUrl || "");
    // Health is polled by the UI status bot — do not spam the console.
    if (!url.startsWith("/api") || url.startsWith("/api/health")) return;
    console.log(`[api] ${req.method} ${url} -> ${res.statusCode} (${Date.now() - start}ms)`);
  });
  next();
});

function withMulter(middleware) {
  return (req, res, next) => {
    middleware(req, res, (err) => {
      if (!err) return next();
      const message =
        err.code === "LIMIT_FILE_SIZE"
          ? "File is too large. Maximum size is 25 MB."
          : err.message || XLSX_ONLY_MSG;
      res.status(400).json({ ok: false, message });
    });
  };
}

function listXlsxUnder(dir, relativeRoot) {
  const out = [];
  if (!fs.existsSync(dir)) return out;
  function walk(current) {
    for (const name of fs.readdirSync(current)) {
      if (isLockFile(name) || name === "node_modules") continue;
      const full = path.join(current, name);
      let st;
      try {
        st = fs.statSync(full);
      } catch {
        continue;
      }
      if (st.isDirectory()) {
        walk(full);
      } else if (isXlsxFileName(name)) {
        out.push({
          name,
          relative: path.relative(relativeRoot, full).split(path.sep).join("/"),
          size: st.size,
        });
      }
    }
  }
  walk(dir);
  return out;
}

function listClientFiles() {
  return listXlsxUnder(CLIENT_DIR, CLIENT_DIR);
}

function listKenyaFiles() {
  const seen = new Set();
  const out = [];
  for (const dir of KENYA_DIRS) {
    for (const file of listXlsxUnder(dir, ROOT)) {
      if (seen.has(file.relative)) continue;
      seen.add(file.relative);
      out.push(file);
    }
  }
  return out;
}

function listEpSampleFiles() {
  if (!fs.existsSync(EP_SAMPLE_DIR)) return [];
  return listXlsxUnder(EP_SAMPLE_DIR, ROOT);
}

function resolveEpSampleRelative(relative) {
  const abs = safeResolveUnder(ROOT, relative);
  if (abs !== EP_SAMPLE_DIR && !abs.startsWith(EP_SAMPLE_DIR + path.sep)) {
    const err = new Error("Sample file must be inside Execution Plan Sample file folder.");
    err.status = 400;
    throw err;
  }
  return abs;
}

function readHistory() {
  try {
    if (!fs.existsSync(HISTORY_PATH)) return [];
    const raw = JSON.parse(fs.readFileSync(HISTORY_PATH, "utf8"));
    const list = Array.isArray(raw) ? raw : [];
    if (list.length > HISTORY_LIMIT) {
      const trimmed = list.slice(0, HISTORY_LIMIT);
      try {
        fs.writeFileSync(HISTORY_PATH, JSON.stringify(trimmed, null, 2), "utf8");
      } catch {
        /* keep serving the trimmed list even if write fails */
      }
      return trimmed;
    }
    return list;
  } catch {
    return [];
  }
}

function pushHistory(entry) {
  try {
    ensureDir(LOG_DIR);
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

function configFromRequest(req, clientFileName) {
  const jobs = loadJobs(JOBS_PATH);
  const props = loadProperties(PROPS_PATH);
  const resolved = resolveJobConfig(clientFileName, jobs, props);
  const body = req.body || {};
  const customModule = String(body.moduleCustom || "").trim();
  const customEntity = String(body.entityCustom || "").trim();
  const entityParts = customEntity
    ? customEntity.split(/[,;]/)
    : Array.isArray(body.entity)
      ? body.entity
      : String(body.entity || "").split(/[,;]/);
  const entityList = [
    ...new Set(entityParts.map((v) => String(v || "").trim()).filter(Boolean)),
  ];
  return {
    Module: customModule || String(body.module || "").trim(),
    Entity: entityList.join(", "),
    Versions: String(body.versions || resolved.Versions || "v1.0").trim(),
    TestcaseType: String(body.testcaseType || resolved.TestcaseType || "WEB").trim(),
  };
}

function saveUpload(dir, file) {
  ensureDir(dir);
  const dest = path.join(dir, path.basename(file.originalname));
  try {
    fs.writeFileSync(dest, file.buffer);
  } catch (err) {
    if (err && (err.code === "EBUSY" || err.code === "EPERM")) {
      return dest;
    }
    throw err;
  }
  return dest;
}

function assertExcelUpload(file, label) {
  if (!file) return;
  if (!isXlsxFileName(file.originalname)) {
    const err = new Error(`${label}: ${XLSX_ONLY_MSG}`);
    err.status = 400;
    throw err;
  }
}

function resolveKenyaRelative(relative) {
  const abs = safeResolveUnder(ROOT, relative);
  const allowed = KENYA_DIRS.some(
    (dir) => abs === dir || abs.startsWith(dir + path.sep)
  );
  if (!allowed) {
    const err = new Error("Mapper file must be inside Kenya doc or Kenya original folder.");
    err.status = 400;
    throw err;
  }
  return abs;
}

async function runMapping(req, { generate }) {
  const clientFile = (req.files && req.files.client && req.files.client[0]) || null;
  const kenyaFile = (req.files && req.files.kenya && req.files.kenya[0]) || null;
  const existingClient = String(req.body.existingClient || "").trim();
  const existingKenya = String(req.body.existingKenya || "").trim();

  let clientFileName;
  let clientBuffer;
  if (clientFile) {
    assertExcelUpload(clientFile, "Client document");
    clientFileName = clientFile.originalname;
    clientBuffer = clientFile.buffer;
    saveUpload(CLIENT_DIR, clientFile);
  } else if (existingClient) {
    const abs = safeResolveUnder(CLIENT_DIR, existingClient);
    clientFileName = path.basename(abs);
    clientBuffer = fs.readFileSync(abs);
  } else {
    const err = new Error("Upload a client document or pick one from Client doc.");
    err.status = 400;
    throw err;
  }

  let kenyaFileName = null;
  let kenyaBuffer = null;
  if (kenyaFile) {
    assertExcelUpload(kenyaFile, "Mapper file");
    kenyaFileName = kenyaFile.originalname;
    kenyaBuffer = kenyaFile.buffer;
    saveUpload(KENYA_DIR, kenyaFile);
  } else if (existingKenya) {
    const abs = resolveKenyaRelative(existingKenya);
    kenyaFileName = path.basename(abs);
    kenyaBuffer = fs.readFileSync(abs);
  }

  const mapperHeader = String((req.body && req.body.mapperHeader) || "").trim();
  const clientSheet = String((req.body && req.body.clientSheet) || "").trim();
  const workbookSheets = listWorkbookSheets(clientBuffer);
  if (workbookSheets.length >= 2 && !clientSheet) {
    const err = new Error(
      `This client file has ${workbookSheets.length} sheets (${workbookSheets.join(
        ", "
      )}). Choose which sheet to map.`
    );
    err.status = 400;
    throw err;
  }

  const config = configFromRequest(req, clientFileName);
  if (clientSheet) config.clientSheet = clientSheet;
  if (!config.Module) {
    const err = new Error("Select a Module (or type a custom module).");
    err.status = 400;
    throw err;
  }
  if (!config.Entity) {
    const err = new Error("Select an Entity, or type a custom entity.");
    err.status = 400;
    throw err;
  }

  const mapped = await mapFromBuffers({
    clientFileName,
    clientBuffer,
    kenyaFileName,
    kenyaBuffer,
    config,
    mapperHeader,
  });

  let outName = "";
  let logName = "";
  if (generate) {
    const stamp = nowStamp();
    const baseName = safeBaseName(clientFileName);
    outName = `${baseName}_SimplifyQA_${stamp}.xlsx`;
    logName = `${baseName}_mapping_${stamp}.log`;
    ensureDir(OUT_DIR);
    ensureDir(LOG_DIR);
    writeWorkbook(mapped.outRows, path.join(OUT_DIR, outName));
    mapped.summary.outName = outName;
    mapped.summary.logName = logName;
    writeLog(path.join(LOG_DIR, logName), mapped.summary);
    pushHistory({
      at: new Date().toISOString(),
      clientFile: clientFileName,
      kenyaFile: kenyaFileName,
      outName,
      logName,
      tcs: mapped.summary.mappedTcCount,
      steps: mapped.summary.mappedStepCount,
      issues: mapped.summary.issues.length,
      kenyaMatched: mapped.summary.kenyaMatched,
      pass: mapped.summary.tcMatch && mapped.summary.stepMatch && mapped.summary.seqOk,
    });
  }

  return {
    ok: true,
    generated: Boolean(generate),
    kenyaFile: kenyaFileName,
    summary: mapped.summary,
    preview: mapped.preview,
    download: generate
      ? {
          excel: `/api/download/excel?file=${encodeURIComponent(outName)}`,
          log: `/api/download/log?file=${encodeURIComponent(logName)}`,
        }
      : null,
  };
}

app.get("/api/health", (_req, res) => {
  res.json({
    ok: true,
    app: APP_NAME,
    version: APP_VERSION,
    healthPollMs: HEALTH_POLL_MS,
    port: PORT,
    pid: process.pid,
    routes: listedApiRoutes(),
  });
});

app.get("/api/config", (_req, res) => {
  const props = loadProperties(PROPS_PATH);
  const jobs = loadJobs(JOBS_PATH);
  res.json({
    ok: true,
    app: APP_NAME,
    version: APP_VERSION,
    healthPollMs: HEALTH_POLL_MS,
    props,
    jobs,
    clientFiles: listClientFiles(),
    kenyaFiles: listKenyaFiles(),
    epSampleFiles: listEpSampleFiles(),
    hasEnvToken: Boolean(getBearerToken(ROOT)),
    supportedDateFormats: SUPPORTED_DATE_FORMATS,
    xlsxOnly: true,
    xlsxMessage: XLSX_ONLY_MSG,
    modules: MODULES,
    entities: ENTITIES,
  });
});

app.get("/api/properties", (_req, res) => {
  const text = fs.existsSync(PROPS_PATH) ? fs.readFileSync(PROPS_PATH, "utf8") : "";
  res.json({ ok: true, text, props: loadProperties(PROPS_PATH) });
});

app.post("/api/properties", (req, res) => {
  try {
    const text = String((req.body && req.body.text) || "");
    fs.writeFileSync(PROPS_PATH, text, "utf8");
    res.json({ ok: true, props: loadProperties(PROPS_PATH) });
  } catch (err) {
    res.status(400).json({ ok: false, message: err.message || String(err) });
  }
});

app.get("/api/auth/status", (_req, res) => {
  try {
    res.json({ ok: true, ...getTokenStatus(ROOT) });
  } catch (err) {
    res.status(400).json({
      ok: false,
      message: err.message || String(err),
    });
  }
});

app.put("/api/auth/token", (req, res) => {
  try {
    const token = req.body && typeof req.body.token === "string" ? req.body.token : "";
    const status = saveBearerToken(token, ROOT);
    res.json({
      ok: true,
      message: "Bearer token saved to .env. Active for SimplifyQA Live calls (no restart needed).",
      ...status,
    });
  } catch (err) {
    res.status(err.status || 400).json({
      ok: false,
      message: err.message || String(err),
    });
  }
});

app.get("/api/history", (_req, res) => {
  res.json({ ok: true, runs: readHistory() });
});

app.post("/api/upload-kenya", withMulter(upload.single("kenya")), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ ok: false, message: "Choose a Kenya Excel file to upload." });
    }
    assertExcelUpload(req.file, "Kenya document");
    const dest = saveUpload(KENYA_DIR, req.file);
    const relative = path.relative(ROOT, dest).split(path.sep).join("/");
    res.json({
      ok: true,
      savedAs: path.basename(dest),
      relative,
      folder: "Kenya doc",
      kenyaFiles: listKenyaFiles(),
      message: `Saved to Kenya doc/${path.basename(dest)}`,
    });
  } catch (err) {
    res.status(err.status || 400).json({ ok: false, message: err.message || String(err) });
  }
});

app.post("/api/upload-client", withMulter(upload.single("client")), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ ok: false, message: "Choose a client Excel file to upload." });
    }
    assertExcelUpload(req.file, "Client document");
    const dest = saveUpload(CLIENT_DIR, req.file);
    const relative = path.relative(CLIENT_DIR, dest).split(path.sep).join("/");
    res.json({
      ok: true,
      savedAs: path.basename(dest),
      relative,
      folder: "Client doc",
      clientFiles: listClientFiles(),
      message: `Saved to Client doc/${path.basename(dest)}`,
    });
  } catch (err) {
    res.status(err.status || 400).json({ ok: false, message: err.message || String(err) });
  }
});

app.post(
  "/api/review",
  withMulter(
    upload.fields([
      { name: "client", maxCount: 1 },
      { name: "kenya", maxCount: 1 },
    ])
  ),
  async (req, res) => {
    try {
      const result = await runMapping(req, { generate: false });
      res.json(result);
    } catch (err) {
      res.status(err.status || 400).json({ ok: false, message: err.message || String(err) });
    }
  }
);

app.post(
  "/api/generate",
  withMulter(
    upload.fields([
      { name: "client", maxCount: 1 },
      { name: "kenya", maxCount: 1 },
    ])
  ),
  async (req, res) => {
    try {
      const result = await runMapping(req, { generate: true });
      res.json(result);
    } catch (err) {
      res.status(err.status || 400).json({ ok: false, message: err.message || String(err) });
    }
  }
);

function resolveCompareSide(req, side) {
  const uploadKey = side === "A" ? "clientA" : "clientB";
  const existingKey = side === "A" ? "existingClientA" : "existingClientB";
  const uploaded = (req.files && req.files[uploadKey] && req.files[uploadKey][0]) || null;
  const existing = String((req.body && req.body[existingKey]) || "").trim();
  const label = side === "A" ? "Excel A" : "Excel B";

  if (uploaded) {
    assertExcelUpload(uploaded, label);
    saveUpload(CLIENT_DIR, uploaded);
    return { fileName: uploaded.originalname, buffer: uploaded.buffer };
  }
  if (existing) {
    const abs = safeResolveUnder(CLIENT_DIR, existing);
    return { fileName: path.basename(abs), buffer: fs.readFileSync(abs) };
  }
  const err = new Error(`Choose ${label} (.xlsx) or pick an existing Client doc.`);
  err.status = 400;
  throw err;
}

function parseEntityField(raw, label) {
  const parts = Array.isArray(raw)
    ? raw
    : String(raw || "")
        .split(/[,;]/);
  const list = [...new Set(parts.map((v) => String(v || "").trim()).filter(Boolean))];
  if (!list.length) {
    const err = new Error(`Select or type Entity for ${label}.`);
    err.status = 400;
    throw err;
  }
  return list.join(", ");
}

async function runCompare(req, { generate }) {
  const sideA = resolveCompareSide(req, "A");
  const sideB = resolveCompareSide(req, "B");
  const kenyaFile = (req.files && req.files.kenya && req.files.kenya[0]) || null;
  const existingKenya = String((req.body && req.body.existingKenya) || "").trim();
  const body = req.body || {};
  const customModule = String(body.moduleCustom || "").trim();
  const moduleName = customModule || String(body.module || "").trim();
  const fallbackEntity = String(body.entity || "").trim();
  const entityCommon = parseEntityField(
    body.entityCustomCommon || body.entityCommon || fallbackEntity,
    "Common sheet"
  );
  const entityUniqueA = parseEntityField(
    body.entityCustomUniqueA || body.entityUniqueA || fallbackEntity,
    "unique Excel A sheet"
  );
  const entityUniqueB = parseEntityField(
    body.entityCustomUniqueB || body.entityUniqueB || fallbackEntity,
    "unique Excel B sheet"
  );
  const versions = String(body.versions || "v1.0").trim() || "v1.0";
  const testcaseType = String(body.testcaseType || "WEB").trim() || "WEB";

  if (!moduleName) {
    const err = new Error("Select a Module (or type a custom module).");
    err.status = 400;
    throw err;
  }

  let kenyaFileName = null;
  let kenyaBuffer = null;
  if (kenyaFile) {
    assertExcelUpload(kenyaFile, "Mapper file");
    kenyaFileName = kenyaFile.originalname;
    kenyaBuffer = kenyaFile.buffer;
    saveUpload(KENYA_DIR, kenyaFile);
  } else if (existingKenya) {
    const abs = resolveKenyaRelative(existingKenya);
    kenyaFileName = path.basename(abs);
    kenyaBuffer = fs.readFileSync(abs);
  }

  const mapperHeader = String((body.mapperHeader) || "").trim();
  const clientSheetA = String(body.clientSheetA || "").trim();
  const clientSheetB = String(body.clientSheetB || "").trim();
  const sheetsA = listWorkbookSheets(sideA.buffer);
  const sheetsB = listWorkbookSheets(sideB.buffer);
  if (sheetsA.length >= 2 && !clientSheetA) {
    const err = new Error(
      `Excel A has ${sheetsA.length} sheets (${sheetsA.join(", ")}). Choose which sheet to compare.`
    );
    err.status = 400;
    throw err;
  }
  if (sheetsB.length >= 2 && !clientSheetB) {
    const err = new Error(
      `Excel B has ${sheetsB.length} sheets (${sheetsB.join(", ")}). Choose which sheet to compare.`
    );
    err.status = 400;
    throw err;
  }

  const compared = await compareClientDocs({
    fileAName: sideA.fileName,
    bufferA: sideA.buffer,
    fileBName: sideB.fileName,
    bufferB: sideB.buffer,
    moduleName,
    entityCommon,
    entityUniqueA,
    entityUniqueB,
    versions,
    testcaseType,
    kenyaFileName,
    kenyaBuffer,
    mapperHeader,
    clientSheetA: clientSheetA || (sheetsA.length === 1 ? sheetsA[0] : ""),
    clientSheetB: clientSheetB || (sheetsB.length === 1 ? sheetsB[0] : ""),
  });

  let outName = "";
  let logName = "";
  if (generate) {
      const stamp = nowStamp();
    outName = `${safeBaseName(sideA.fileName)}_vs_${safeBaseName(sideB.fileName)}_Compare_${stamp}.xlsx`;
    logName = `${safeBaseName(sideA.fileName)}_vs_${safeBaseName(sideB.fileName)}_compare_${stamp}.log`;
    ensureDir(OUT_DIR);
    ensureDir(LOG_DIR);
    fs.writeFileSync(path.join(OUT_DIR, outName), compared.buffer);
    compared.summary.outName = outName;
    compared.summary.logName = logName;
    writeCompareLog(path.join(LOG_DIR, logName), compared.summary);
    pushHistory({
      at: new Date().toISOString(),
      clientFile: `${sideA.fileName} vs ${sideB.fileName}`,
      kenyaFile: kenyaFileName,
      outName,
      logName,
      tcs: compared.summary.mappedTcCount,
      steps: compared.summary.mappedStepCount,
      issues: (compared.summary.issues || []).length,
      kenyaMatched: compared.summary.kenyaMatched,
      pass: compared.summary.seqOk && compared.summary.tcMatch && compared.summary.stepMatch,
      kind: "compare",
    });
  }

  return {
    ok: true,
    generated: Boolean(generate),
    kenyaFile: kenyaFileName,
    summary: compared.summary,
    preview: compared.preview,
    download: generate
      ? {
          excel: `/api/download/excel?file=${encodeURIComponent(outName)}`,
          log: `/api/download/log?file=${encodeURIComponent(logName)}`,
        }
      : null,
  };
}

function compareUpload() {
  return withMulter(
    upload.fields([
      { name: "clientA", maxCount: 1 },
      { name: "clientB", maxCount: 1 },
      { name: "kenya", maxCount: 1 },
    ])
  );
}

app.post("/api/compare-review", compareUpload(), async (req, res) => {
  try {
    const result = await runCompare(req, { generate: false });
    res.json(result);
  } catch (err) {
    res.status(err.status || 400).json({ ok: false, message: err.message || String(err) });
  }
});

app.post("/api/compare", compareUpload(), async (req, res) => {
  try {
    const result = await runCompare(req, { generate: true });
    res.json(result);
  } catch (err) {
    res.status(err.status || 400).json({ ok: false, message: err.message || String(err) });
  }
});

app.post(
  "/api/client-sheets",
  withMulter(upload.single("client")),
  (req, res) => {
    try {
      const uploaded = req.file || null;
      const existingClient = String((req.body && req.body.existingClient) || "").trim();
      let fileName = "";
      let buffer = null;
      if (uploaded) {
        assertExcelUpload(uploaded, "Client document");
        fileName = uploaded.originalname;
        buffer = uploaded.buffer;
      } else if (existingClient) {
        const abs = safeResolveUnder(CLIENT_DIR, existingClient);
        fileName = path.basename(abs);
        buffer = fs.readFileSync(abs);
      } else {
        const err = new Error("Upload a client document or pick an existing Client doc.");
        err.status = 400;
        throw err;
      }
      const sheets = listWorkbookSheets(buffer);
      res.json({
        ok: true,
        fileName,
        sheets,
        needsChoice: sheets.length >= 2,
      });
    } catch (err) {
      res.status(err.status || 400).json({ ok: false, message: err.message || String(err) });
    }
  }
);

app.post(
  "/api/mapper-headers",
  withMulter(upload.single("kenya")),
  async (req, res) => {
    try {
      const uploaded = req.file || null;
      const existingKenya = String((req.body && req.body.existingKenya) || "").trim();
      let fileName = "";
      let buffer = null;
      if (uploaded) {
        assertExcelUpload(uploaded, "Mapper file");
        fileName = uploaded.originalname;
        buffer = uploaded.buffer;
      } else if (existingKenya) {
        const abs = resolveKenyaRelative(existingKenya);
        fileName = path.basename(abs);
        buffer = fs.readFileSync(abs);
      } else {
        const err = new Error("Upload a mapper file or pick an existing one.");
        err.status = 400;
        throw err;
      }
      const listed = listMapperHeaders(buffer, fileName);
      res.json({
        ok: listed.ok !== false || (listed.headers && listed.headers.length > 0),
        ...listed,
      });
    } catch (err) {
      res.status(err.status || 400).json({ ok: false, message: err.message || String(err) });
    }
  }
);

async function runEpMapping(req, { generate }) {
  const body = req.body || {};
  const mode = String(body.mode || "upload").toLowerCase();

  // Resolve modules: support modules array, comma-separated string, or customModule / module fallback
  let moduleList = [];
  const customModule = String(body.moduleCustom || "").trim();
  if (customModule) {
    moduleList = customModule.split(",").map((m) => m.trim()).filter(Boolean);
  } else if (Array.isArray(body.modules)) {
    moduleList = body.modules.map((m) => String(m).trim()).filter(Boolean);
  } else if (typeof body.modules === "string" && body.modules.trim()) {
    moduleList = body.modules.split(",").map((m) => m.trim()).filter(Boolean);
  } else if (body.module) {
    moduleList = [String(body.module).trim()].filter(Boolean);
  }

  const moduleName = moduleList.join(", ");
  const customEntity = String(body.entityCustom || "").trim();
  const entityName = customEntity || String(body.entity || "").trim();
  const version = String(body.version || "v1.0").trim() || "v1.0";
  const executionType = String(body.executionType || "Manual").trim() || "Manual";
  const assignedDate = String(body.assignedDate || "").trim();
  const dateFormat = String(body.dateFormat || "mm/dd/yyyy").trim();
  const assigneeEmail = String(body.assigneeEmail || "").trim();
  const projectId = Number(body.projectId) || 5;

  let testcases = [];
  let sourceLabel = "";

  if (assignedDate) {
    const dateValidation = validateAndFormatDate(assignedDate, dateFormat);
    if (!dateValidation.valid) {
      const err = new Error(dateValidation.error);
      err.status = 400;
      throw err;
    }
  }

  if (mode === "live") {
    const token = String(body.token || getBearerToken(ROOT) || "").trim();
    if (!token) {
      const err = new Error(
        "SimplifyQA Bearer token is required for live API export (provide token or set SIMPLIFYQA_BEARER_TOKEN in .env)."
      );
      err.status = 400;
      throw err;
    }
    const fetched = await fetchTestcasesFromSimplifyQa({
      token,
      projectId,
      moduleNames: moduleList,
      moduleName,
      entity: entityName,
    });
    testcases = fetched.testcases;
    sourceLabel = `SimplifyQA Live API (Project ${projectId})`;
  } else {
    const uploaded = req.file || (req.files && req.files.summary && req.files.summary[0]) || null;
    const existingSummary = String(body.existingSummary || "").trim();
    let buffer = null;
    let fileName = "";

    if (uploaded) {
      assertExcelUpload(uploaded, "Summary Excel");
      fileName = uploaded.originalname;
      buffer = uploaded.buffer;
      saveUpload(EP_SAMPLE_DIR, uploaded);
    } else if (existingSummary) {
      let abs;
      if (fs.existsSync(safeResolveUnder(ROOT, existingSummary))) {
        abs = safeResolveUnder(ROOT, existingSummary);
      } else {
        abs = resolveEpSampleRelative(existingSummary);
      }
      fileName = path.basename(abs);
      buffer = fs.readFileSync(abs);
    } else {
      const err = new Error(
        "Upload a SimplifyQA summary export file (.xlsx) or select an existing sample file, or switch to Live API mode."
      );
      err.status = 400;
      throw err;
    }

    const extracted = extractTestcasesFromSummary(buffer, {
      modules: moduleList,
      module: moduleName,
      entity: entityName,
      sheet: body.summarySheet,
    });
    testcases = extracted.testcases;
    sourceLabel = fileName;
  }

  if (!testcases.length) {
    const filterDesc = [moduleName ? `Module(s): ${moduleName}` : null, entityName ? `Entity: ${entityName}` : null]
      .filter(Boolean)
      .join(", ");
    const err = new Error(
      `No test cases found${filterDesc ? ` matching ${filterDesc}` : ""}. Ensure the export contains valid test case rows.`
    );
    err.status = 400;
    throw err;
  }

  const result = buildEpWorkbook({
    testcases,
    moduleNames: moduleList,
    moduleName,
    entityName,
    version,
    executionType,
    assignedDate,
    assigneeEmail,
    dateFormat,
  });

  if (generate) {
    const stamp = nowStamp();
    let modFilePart = moduleList.length === 1 ? moduleList[0] : (moduleList.length > 1 ? `${moduleList.length}_Modules` : "All_Modules");
    const cleanMod = safeBaseName(modFilePart || "Module");
    const cleanEnt = safeBaseName(entityName || "Entity");
    const outName = `${cleanMod}_${cleanEnt}_EP_${stamp}.xlsx`;
    const logName = `${cleanMod}_${cleanEnt}_ep_${stamp}.log`;
    ensureDir(OUT_DIR);
    ensureDir(LOG_DIR);
    fs.writeFileSync(path.join(OUT_DIR, outName), result.buffer);

    const logLines = [
      `ICEA LION Execution Plan (EP) Mapping Log`,
      `=========================================`,
      `Date/Time        : ${new Date().toISOString()}`,
      `Source           : ${sourceLabel}`,
      `Module(s)        : ${moduleName || "(all)"}`,
      `Entity           : ${entityName || "(all)"}`,
      `Output Sheet(s)  : ${(result.sheetNames || [result.sheetName]).join(", ")}`,
      `Total TCs Mapped : ${testcases.length}`,
      `Version Default  : ${version}`,
      `Execution Type   : ${executionType}`,
      `Assigned Date    : ${result.summary.assignedDate || "(none)"}`,
      `Assignee Email   : ${assigneeEmail || "(none)"}`,
      `Generated File   : ${outName}`,
      ``,
      `Sheets Breakdown:`,
      ...(result.sheetStats || []).map((s) => `  - Sheet "${s.sheetName}": ${s.testcaseCount} test case(s)`),
      ``,
      `Mapped Test Cases:`,
      ...testcases.map(
        (t, idx) => `  ${idx + 1}. [${t.id}] ${t.name} (Module: ${t.module || "-"}, Entity: ${t.entity || "-"})`
      ),
    ];
    fs.writeFileSync(path.join(LOG_DIR, logName), logLines.join("\n"), "utf8");

    pushHistory({
      at: new Date().toISOString(),
      clientFile: `${sourceLabel} → EP Plan`,
      kenyaFile: null,
      outName,
      logName,
      tcs: testcases.length,
      steps: testcases.length,
      issues: 0,
      kenyaMatched: 0,
      pass: true,
      kind: "ep",
    });

    return {
      ok: true,
      mode,
      download: {
        excel: `/api/download/excel?file=${encodeURIComponent(outName)}`,
        log: `/api/download/log?file=${encodeURIComponent(logName)}`,
      },
      summary: {
        ...result.summary,
        outName,
        logName,
        sourceLabel,
      },
      preview: result.preview,
    };
  }

  return {
    ok: true,
    mode,
    summary: {
      ...result.summary,
      sourceLabel,
    },
    preview: result.preview,
  };
}

app.get("/api/ep/sample-files", (_req, res) => {
  res.json({
    ok: true,
    sampleFiles: listEpSampleFiles(),
    supportedDateFormats: SUPPORTED_DATE_FORMATS,
    hasEnvToken: Boolean(getBearerToken(ROOT)),
  });
});

app.post("/api/ep/review", withMulter(upload.single("summary")), async (req, res) => {
  try {
    const result = await runEpMapping(req, { generate: false });
    res.json(result);
  } catch (err) {
    res.status(err.status || 400).json({ ok: false, message: err.message || String(err) });
  }
});

app.post("/api/ep/generate", withMulter(upload.single("summary")), async (req, res) => {
  try {
    const result = await runEpMapping(req, { generate: true });
    res.json(result);
  } catch (err) {
    res.status(err.status || 400).json({ ok: false, message: err.message || String(err) });
  }
});

app.get("/api/download/excel", (req, res) => {
  try {
    const abs = safeResolveUnder(OUT_DIR, req.query.file);
    if (!fs.existsSync(abs)) {
      return res.status(404).json({ ok: false, message: "Excel not found." });
    }
    res.download(abs, path.basename(abs));
  } catch (err) {
    res.status(err.status || 400).json({ ok: false, message: err.message || String(err) });
  }
});

app.get("/api/download/log", (req, res) => {
  try {
    const abs = safeResolveUnder(LOG_DIR, req.query.file);
    if (!fs.existsSync(abs)) {
      return res.status(404).json({ ok: false, message: "Log not found." });
    }
    res.download(abs, path.basename(abs));
  } catch (err) {
    res.status(err.status || 400).json({ ok: false, message: err.message || String(err) });
  }
});

app.get("/api/preview/output", (req, res) => {
  try {
    const abs = safeResolveUnder(OUT_DIR, req.query.file);
    if (!fs.existsSync(abs)) {
      return res.status(404).json({ ok: false, message: "File not found." });
    }
    const wb = XLSX.readFile(abs);
    const sheet = String(req.query.sheet || wb.SheetNames[0]);
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: "" });
    const maxCol = rows.reduce((m, r) => Math.max(m, r.length), 0);
    res.json({
      ok: true,
      sheet,
      sheets: wb.SheetNames,
      maxCol,
      rows: rows.slice(0, 120).map((cells, i) => ({
        row: i + 1,
        cells: Array.from({ length: maxCol }, (_, c) => ({
          text: String(cells[c] == null ? "" : cells[c]),
        })),
      })),
      truncated: rows.length > 120,
      totalRows: rows.length,
    });
  } catch (err) {
    res.status(400).json({ ok: false, message: err.message || String(err) });
  }
});

app.use("/api/reporter", createReporterRouter());

app.post("/api/launch-excel", (req, res) => {
  try {
    const rawFile = req.body && req.body.file;
    let abs;
    try {
      abs = safeResolveUnder(OUT_DIR, rawFile);
      if (!fs.existsSync(abs)) {
        abs = safeResolveUnder(ROOT, rawFile);
      }
    } catch {
      abs = safeResolveUnder(ROOT, rawFile);
    }
    if (!fs.existsSync(abs)) {
      return res.status(404).json({ ok: false, message: "Excel not found." });
    }
    const { execFile } = require("child_process");
    execFile("cmd.exe", ["/c", "start", "", abs], { windowsHide: false }, (err) => {
      if (err) {
        return res.status(500).json({ ok: false, message: err.message });
      }
      res.json({ ok: true, launched: true, message: "Opened in Excel." });
    });
  } catch (err) {
    res.status(err.status || 400).json({ ok: false, message: err.message || String(err) });
  }
});

function lanAddresses() {
  try {
    const nets = os.networkInterfaces();
    const result = [];
    for (const entries of Object.values(nets || {})) {
      for (const net of entries || []) {
        const family = net.family;
        if ((family === "IPv4" || family === 4) && !net.internal && net.address) {
          result.push(net.address);
        }
      }
    }
    return result;
  } catch (err) {
    console.warn("Could not list LAN addresses:", err.message);
    return [];
  }
}

app.use("/api", (req, res) => {
  res.status(404).json({
    ok: false,
    message: `Unknown API route ${req.method} ${req.originalUrl}. Start the app with npm start and open http://localhost:${PORT}.`,
  });
});

app.use(express.static(path.join(ROOT, "public")));
if (fs.existsSync(path.join(ROOT, "Logo"))) {
  app.use("/logo", express.static(path.join(ROOT, "Logo")));
}

app.use((err, req, res, next) => {
  if (res.headersSent) return next(err);
  const api = req.path && req.path.startsWith("/api");
  const message = err.message || String(err);
  if (api) {
    return res.status(err.status || 500).json({ ok: false, message });
  }
  res.status(err.status || 500).type("text/plain").send(message);
});

ensureDir(CLIENT_DIR);
ensureDir(KENYA_DIR);
ensureDir(KENYA_ORIGINAL_DIR);
ensureDir(OUT_DIR);
ensureDir(LOG_DIR);

function listedApiRoutes() {
  try {
    const stack = app.router && app.router.stack;
    if (!Array.isArray(stack)) return [];
    const out = [];
    for (const layer of stack) {
      if (!layer.route || !layer.route.path) continue;
      const methods = Object.keys(layer.route.methods || {}).filter((m) => layer.route.methods[m]);
      if (!methods.length) continue;
      out.push(`${methods.map((m) => m.toUpperCase()).join(",")} ${layer.route.path}`);
    }
    return out;
  } catch {
    return [];
  }
}

function portListeners(port) {
  try {
    return execSync(`netstat -ano | findstr :${port}`, { encoding: "utf8" }).trim();
  } catch {
    return "";
  }
}

function startServer(port = PORT) {
  const server = http.createServer(app);
  server.on("error", (err) => {
    if (err && err.code === "EADDRINUSE") {
      console.error(`ERROR: Port ${port} is already in use. This process did NOT start.`);
      console.error("Close the other Node/Live Preview window, then run npm start again.");
      const listeners = portListeners(port);
      if (listeners) {
        console.error("What is using the port:");
        console.error(listeners);
      }
    } else {
      console.error("Server error:", err);
    }
    process.exit(1);
  });
  server.on("close", () => {
    console.warn("HTTP server closed.");
  });
  server.listen(port, "0.0.0.0", () => {
    const addr = server.address();
    const bound = addr && typeof addr === "object" ? addr.port : port;
    console.log(`pid ${process.pid}`);
    console.log(`ICEA LION Test Management Hub UI  v${APP_VERSION}  http://localhost:${bound}`);
    for (const ip of lanAddresses()) {
      console.log(`  LAN  http://${ip}:${bound}`);
    }
    const routes = listedApiRoutes().filter((r) => r.includes("/api/"));
    console.log("API routes:");
    for (const route of routes) {
      console.log(`  ${route}`);
    }
    if (!routes.some((r) => r.includes("/api/compare"))) {
      console.error("ERROR: POST /api/compare was not registered.");
    }
    console.log("Keep this window open. Press Ctrl+C to stop.");
  });
  return server;
}

if (require.main === module) {
  process.on("uncaughtException", (err) => {
    console.error("uncaughtException:", err);
  });
  process.on("unhandledRejection", (err) => {
    console.error("unhandledRejection:", err);
  });
  if (process.stdin && typeof process.stdin.resume === "function") {
    try {
      process.stdin.resume();
    } catch {
      // ignore
    }
  }
  const server = startServer(PORT);
  const stop = () => {
    server.close(() => process.exit(0));
  };
  process.on("SIGINT", stop);
  process.on("SIGTERM", stop);
}

module.exports = { app, startServer, PORT, CLIENT_DIR, KENYA_DIR, ROOT };
