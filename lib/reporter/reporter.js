const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");
require("../loadEnv").loadProjectEnv(path.resolve(__dirname, "../../"), false);

const ROOT = path.resolve(__dirname, "../../");
const ORIGIN = "https://app.simplifyqa.ai";
const SWITCH_PROJECT_URL = `${ORIGIN}/pm/fields/8`;
const EXPORT_URL = `${ORIGIN}/tm/executionplan/export`;
const DEFECT_EXPORT_URL = `${ORIGIN}/pm/defect/export`;
const PROPERTIES_PATH = path.join(ROOT, "config", "application.properties");
const DOWNLOADS_DIR = path.join(ROOT, "downloads");
const OUTPUT_DIR = path.join(ROOT, "output");
const LOGS_DIR = path.join(ROOT, "logs");

const DEFAULT_TEMPLATES = {
  1: "Template/FMS Status tracker.xlsx",
  2: "Template/FMS 4 EP template.xlsx",
  3: "Template/E2E Report.xlsx",
};

/** UI-facing metadata for known templates */
const TEMPLATE_META = {
  1: {
    statusSections: 2,
    description: "2 execution-plan blocks (typically Sprint 1 General + Life Kenya).",
  },
  2: {
    statusSections: 4,
    description: "4 execution-plan blocks (Sprint 1 & 2 for General + Life Kenya).",
  },
  3: {
    statusSections: 2,
    description: "2 E2E blocks (General Kenya then Life Kenya). Extra modules are alerted, not added.",
  },
};

const MODULE_ALIASES = {
  Payables: "Accounts Payable",
  "Account Payables": "Accounts Payable",
  "Accounts Receivables": "Accounts Receivable",
  "Cash & Bank": "Cash & Bank Management",
  "Fixed Assets": "Fixed Assets Management",
  "Data Migration -Functional": "Data Migration TC's",
  "Data Migration -Technical": "Data Migration TC's",
  "Investment Receipting": "Investment Management",
};

/** Normalize SimplifyQA Entity values -> template entity labels */
const ENTITY_ALIASES = {
  "life ke": "Life Kenya",
  "life kenya": "Life Kenya",
  "gen ke/williamson": "General Kenya",
  "general kenya": "General Kenya",
  "ug general": "Uganda General",
  "uganda general": "Uganda General",
  "ug life": "Uganda Life",
  "uganda life": "Uganda Life",
  "tz general": "Tanzania",
  tanzania: "Tanzania",
};

/**
 * Defect State (SimplifyQA) -> template columns.
 * Pending = New + Reopened; Closed = Resolved + Closed; Deferred; Fixed.
 */
const DEFECT_STATE_MAP = {
  NEW: "pending",
  REOPENED: "pending",
  REOPEN: "pending",
  "RE-OPENED": "pending",
  RESOLVED: "closed",
  CLOSED: "closed",
  DEFERRED: "deferred",
  FIXED: "fixed",
};

const STATUS_MAP = {
  PASSED: "passed",
  PASS: "passed",
  FAILED: "failed",
  FAIL: "failed",
  BLOCKED: "blocked",
  "IN PROGRESS": "inProgress",
  IN_PROGRESS: "inProgress",
  INPROGRESS: "inProgress",
  "NOT EXECUTED": "notExecuted",
  NOT_EXECUTED: "notExecuted",
  NOTEXECUTED: "notExecuted",
  "": "notExecuted",
};

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function loadProperties(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new ReporterError(
      "CONFIG_MISSING",
      `Configuration file not found: ${filePath}. Please create config/application.properties.`
    );
  }
  const props = {};
  for (const line of fs.readFileSync(filePath, "utf8").split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#") || trimmed.startsWith(";")) continue;
    const eq = trimmed.indexOf("=");
    if (eq < 0) continue;
    props[trimmed.slice(0, eq).trim()] = trimmed.slice(eq + 1).trim();
  }
  return props;
}

const ENV_PATH = path.join(ROOT, ".env");
const TOKEN_ENV_KEY = "SIMPLIFYQA_BEARER_TOKEN";

function normalizeBearerToken(raw) {
  return String(raw || "")
    .trim()
    .replace(/^Bearer\s+/i, "");
}

function readJwtExpMs(token) {
  try {
    const part = String(token || "").split(".")[1];
    if (!part) return null;
    const json = Buffer.from(
      part.replace(/-/g, "+").replace(/_/g, "/"),
      "base64"
    ).toString("utf8");
    const payload = JSON.parse(json);
    if (!payload.exp) return null;
    return Number(payload.exp) * 1000;
  } catch {
    return null;
  }
}

function getTokenStatus() {
  const loadEnv = require("../loadEnv");
  return loadEnv.getTokenStatus(ROOT);
}

function saveBearerToken(rawToken) {
  const loadEnv = require("../loadEnv");
  return loadEnv.saveBearerToken(rawToken, ROOT);
}

function getBearerToken() {
  const loadEnv = require("../loadEnv");
  const raw = loadEnv.getBearerToken(ROOT) || "";
  if (
    !raw ||
    raw === "your_bearer_token_here" ||
    raw === "your_simplifyqa_bearer_token_here"
  ) {
    throw new ReporterError(
      "AUTH_MISSING",
      "Authentication token is missing. Paste a fresh token in the Auth panel (or set SIMPLIFYQA_BEARER_TOKEN in .env) and try again."
    );
  }
  const token = normalizeBearerToken(raw);
  if (isJwtExpired(token)) {
    throw new ReporterError(
      "AUTH_EXPIRED",
      "Your login session has expired. Paste a fresh bearer token in the Auth panel and try again."
    );
  }
  return token;
}

class ReporterError extends Error {
  constructor(code, message, details) {
    super(message);
    this.name = "ReporterError";
    this.code = code;
    this.details = details || "";
  }
}

function isJwtExpired(token) {
  const expMs = readJwtExpMs(token);
  if (expMs == null) return false;
  return Date.now() >= expMs;
}

function bodySnippet(text) {
  return String(text || "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 400);
}

function looksLikeNotFound(text) {
  return /not\s*found|does\s*not\s*exist|invalid\s*project|no\s*project|unknown\s*project|invalid\s*execution|no\s*execution\s*plan|plan\s*not/i.test(
    text || ""
  );
}

function looksLikeAuthError(text) {
  return /unauthorized|unauthorised|token\s*expired|jwt\s*expired|invalid\s*token|access\s*denied|authentication|forbidden/i.test(
    text || ""
  );
}

function classifyApiFailure(kind, status, body, ids) {
  const text = bodySnippet(body);
  const projectId = ids.projectId;
  const exePlanId = ids.exePlanId;

  if (status === 401 || status === 403 || looksLikeAuthError(text)) {
    return new ReporterError(
      "AUTH_EXPIRED",
      "Your login session has expired or the token is invalid. Paste a fresh bearer token in the Auth panel and try again.",
      `HTTP ${status}${text ? ` - ${text}` : ""}`
    );
  }

  if (kind === "project") {
    if (status === 404 || looksLikeNotFound(text) || status === 400) {
      return new ReporterError(
        "PROJECT_NOT_FOUND",
        `Project not found. Project ID ${projectId} does not exist or you do not have access to it. Please check PROJECT_ID in config/application.properties.`,
        `HTTP ${status}${text ? ` - ${text}` : ""}`
      );
    }
    return new ReporterError(
      "PROJECT_SWITCH_FAILED",
      `Unable to switch to project ${projectId}. Please verify the project ID and your access, then try again.`,
      `HTTP ${status}${text ? ` - ${text}` : ""}`
    );
  }

  if (kind === "export") {
    if (status === 404 || looksLikeNotFound(text)) {
      return new ReporterError(
        "PLAN_NOT_FOUND",
        `Execution plan not found. Plan ID ${exePlanId} is not present in project ${projectId}. Please check EXE_PLAN_ID values in config/application.properties.`,
        `HTTP ${status}${text ? ` - ${text}` : ""}`
      );
    }
    if (status === 400) {
      return new ReporterError(
        "PLAN_NOT_FOUND",
        `Execution plan not found or invalid. Plan ID ${exePlanId} could not be used for project ${projectId}. Please verify the plan ID belongs to this project.`,
        `HTTP ${status}${text ? ` - ${text}` : ""}`
      );
    }
    return new ReporterError(
      "EXTRACT_FAILED",
      `Unable to extract data for execution plan ${exePlanId} from project ${projectId}. Please try again, or check that the plan exists and your session is still valid.`,
      `HTTP ${status}${text ? ` - ${text}` : ""}`
    );
  }

  return new ReporterError(
    "API_ERROR",
    "Something went wrong while calling SimplifyQA. Please check the log for details and try again.",
    `HTTP ${status}${text ? ` - ${text}` : ""}`
  );
}

function parseArgs(argv) {
  const args = {
    localByPlan: {},
    dryRun: false,
    skipSwitch: false,
    includeDefects: true,
    localDefects: null,
  };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a.startsWith("--local:") && argv[i + 1]) {
      args.localByPlan[a.slice("--local:".length)] = argv[++i];
    } else if (a === "--dry-run") {
      args.dryRun = true;
    } else if (a === "--skip-switch") {
      args.skipSwitch = true;
    } else if (a === "--no-defects") {
      args.includeDefects = false;
    } else if (a === "--include-defects") {
      args.includeDefects = true;
    } else if (a === "--local-defects" && argv[i + 1]) {
      args.localDefects = argv[++i];
    }
  }
  return args;
}

function resolveModuleName(rawName) {
  let name = String(rawName || "").trim();
  if (!name) return null;
  // EP exports use E2E* module names that belong with the functional module.
  name = name.replace(/^E2E\s+/i, "").trim();
  return MODULE_ALIASES[name] || name;
}

/**
 * Defect module names must stay distinct from EP E2E modules.
 * Do not strip "E2E " — only apply naming aliases (Payables → Accounts Payable, etc.).
 */
function resolveDefectModuleName(rawName) {
  const name = String(rawName || "").trim();
  if (!name) return null;
  return MODULE_ALIASES[name] || name;
}

function normalizeEntityLabel(raw) {
  const key = String(raw || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
  if (!key) return null;
  return ENTITY_ALIASES[key] || String(raw).trim();
}

/** Split multi-entity values like "Life KE,Gen KE/Williamson" and normalize each. */
function parseDefectEntities(raw) {
  const text = String(raw || "").trim();
  if (!text) return [];
  const parts = text.split(",").map((p) => p.trim()).filter(Boolean);
  const out = [];
  const seen = new Set();
  for (const part of parts) {
    const label = normalizeEntityLabel(part);
    if (!label) continue;
    const k = label.toLowerCase();
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(label);
  }
  return out;
}

function mapDefectState(rawState) {
  const key = String(rawState || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
  if (!key) return null;
  return DEFECT_STATE_MAP[key] || null;
}

function entityFromDefectSectionTitle(title) {
  const t = String(title || "").toLowerCase();
  const checks = [
    ["uganda general", "Uganda General"],
    ["uganda life", "Uganda Life"],
    ["general kenya", "General Kenya"],
    ["life kenya", "Life Kenya"],
    ["tanzania", "Tanzania"],
  ];
  for (const [needle, label] of checks) {
    if (t.includes(needle)) return label;
  }
  return null;
}

function normalizeStatus(raw) {
  const key = String(raw || "").trim().toUpperCase();
  return STATUS_MAP[key] || "notExecuted";
}

function emptyCounts() {
  return {
    passed: 0,
    failed: 0,
    blocked: 0,
    inProgress: 0,
    notExecuted: 0,
  };
}

function rate(numerator, denominator) {
  if (!denominator) return 1;
  return numerator / denominator;
}

function summarizeModule(counts) {
  const total =
    counts.passed +
    counts.failed +
    counts.blocked +
    counts.inProgress +
    counts.notExecuted;
  return {
    ...counts,
    total,
    passRate: rate(counts.passed, total - counts.blocked),
  };
}

function cellText(value) {
  if (value == null) return "";
  if (typeof value === "object") {
    if (value.text) return String(value.text);
    if (value.richText) return value.richText.map((t) => t.text).join("");
    if (value.result != null) return String(value.result);
  }
  return String(value);
}

function resolvePathMaybe(p) {
  if (!p) return p;
  return path.isAbsolute(p) ? p : path.join(ROOT, p);
}

/**
 * Collect EXE_PLAN_ID_1..N consecutively until a gap.
 */
function collectNumbered(props, prefix) {
  const items = [];
  for (let i = 1; i <= 100; i++) {
    const key = `${prefix}${i}`;
    if (!(key in props)) break;
    const value = String(props[key] || "").trim();
    if (!value) break;
    items.push({ index: i, key, value });
  }
  return items;
}

function resolveTemplatePath(props) {
  const choice = String(props.TEMPLATE_CHOICE || "1").trim();
  const catalogKey = `TEMPLATE_${choice}`;
  const fromProps = props[catalogKey];
  const fallback = DEFAULT_TEMPLATES[choice];
  const relative = fromProps || fallback;
  if (!relative) {
    throw new ReporterError(
      "CONFIG_INVALID",
      `Unknown template choice "${choice}". Please set TEMPLATE_${choice}=... or use a valid TEMPLATE_CHOICE in config/application.properties.`
    );
  }

  const outputKey = `OUTPUT_FILE_${choice}`;
  const outputRelative = String(props[outputKey] || "").trim();
  if (!outputRelative) {
    throw new ReporterError(
      "CONFIG_INVALID",
      `Output file is missing for template choice ${choice}. Please set ${outputKey}=... in config/application.properties.`
    );
  }

  return {
    choice,
    templatePath: resolvePathMaybe(relative),
    relative,
    outputFile: resolvePathMaybe(outputRelative),
    outputRelative,
  };
}

function buildRuntime(props) {
  const projectId = Number(props.PROJECT_ID);
  if (!projectId) {
    throw new ReporterError(
      "CONFIG_INVALID",
      "Project ID is missing or invalid. Please set PROJECT_ID in config/application.properties."
    );
  }

  const planEntries = collectNumbered(props, "EXE_PLAN_ID_");
  if (!planEntries.length) {
    throw new ReporterError(
      "CONFIG_INVALID",
      "No execution plan IDs were provided. Please set at least EXE_PLAN_ID_1 in config/application.properties."
    );
  }

  const { choice, templatePath, relative, outputFile, outputRelative } =
    resolveTemplatePath(props);

  const plans = planEntries.map((p) => ({
    index: p.index,
    exePlanId: Number(p.value),
    sectionOverride: String(props[`SECTION_${p.index}`] || "").trim() || null,
  }));

  for (const plan of plans) {
    if (!plan.exePlanId || Number.isNaN(plan.exePlanId)) {
      throw new ReporterError(
        "CONFIG_INVALID",
        `Execution plan ID ${plan.index} is invalid. Please check EXE_PLAN_ID_${plan.index} in config/application.properties.`
      );
    }
  }

  return {
    projectId,
    templateChoice: choice,
    templatePath,
    templateRelative: relative,
    timezone: props.TIMEZONE || "Asia/Calcutta",
    offset: Number(
      props.OFFSET != null && props.OFFSET !== "" ? props.OFFSET : -330
    ),
    plans,
    outputFile,
    outputRelative,
  };
}

function discoverStatusSections(sheet) {
  const sections = [];
  for (let r = 1; r <= sheet.rowCount; r++) {
    for (let c = 1; c <= 3; c++) {
      const text = cellText(sheet.getRow(r).getCell(c).value).trim();
      if (/module-wise\s+daily\s+status/i.test(text)) {
        sections.push({ title: text, titleRow: r });
        break;
      }
    }
  }
  return sections;
}

function mapPlansToSections(plans, discoveredSections) {
  const mapped = [];
  const overflow = [];

  for (let i = 0; i < plans.length; i++) {
    const plan = plans[i];
    if (i >= discoveredSections.length) {
      overflow.push(plan);
      continue;
    }

    let sectionTitle = plan.sectionOverride;
    if (sectionTitle) {
      const found = discoveredSections.find(
        (s) => s.title.toLowerCase() === sectionTitle.toLowerCase()
      );
      if (!found) {
        throw new Error(
          `SECTION_${plan.index} not found in template: ${sectionTitle}`
        );
      }
      sectionTitle = found.title;
    } else {
      sectionTitle = discoveredSections[i].title;
    }

    mapped.push({
      ...plan,
      sectionTitle,
      sectionOrdinal: i + 1,
    });
  }

  return { mapped, overflow };
}

async function switchProject(token, projectId) {
  const url = `${SWITCH_PROJECT_URL}?projectId=${projectId}&fieldType=environment`;
  let res;
  try {
    res = await fetch(url, {
      method: "GET",
      headers: {
        accept: "application/json, text/plain, */*",
        authorization: `Bearer ${token}`,
        referer: `${ORIGIN}/user/home`,
        "user-agent":
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36",
      },
    });
  } catch (err) {
    throw new ReporterError(
      "NETWORK_ERROR",
      "Unable to reach SimplifyQA while switching project. Please check your internet connection and try again.",
      err.message
    );
  }

  const body = await res.text();
  if (!res.ok) {
    throw classifyApiFailure("project", res.status, body, { projectId });
  }

  // Some APIs return 200 with an error payload
  if (looksLikeNotFound(body) || /error|failed/i.test(body) && /project/i.test(body)) {
    if (looksLikeNotFound(body)) {
      throw classifyApiFailure("project", 404, body, { projectId });
    }
  }

  try {
    return body ? JSON.parse(body) : {};
  } catch {
    return {};
  }
}

async function downloadExport(token, projectId, exePlanId, timezone, offset) {
  let res;
  try {
    res = await fetch(EXPORT_URL, {
      method: "POST",
      headers: {
        accept: "*/*",
        authorization: `Bearer ${token}`,
        "content-type": "application/json",
        origin: ORIGIN,
        referer: `${ORIGIN}/user/execution-plan/details/${exePlanId}`,
        "user-agent":
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36",
      },
      body: JSON.stringify({
        data: { projectId, exePlanId },
        timezone,
        offset,
      }),
    });
  } catch (err) {
    throw new ReporterError(
      "NETWORK_ERROR",
      `Unable to reach SimplifyQA while downloading execution plan ${exePlanId}. Please check your internet connection and try again.`,
      err.message
    );
  }

  if (!res.ok) {
    const body = await res.text();
    throw classifyApiFailure("export", res.status, body, {
      projectId,
      exePlanId,
    });
  }

  const contentType = (res.headers.get("content-type") || "").toLowerCase();
  const buffer = Buffer.from(await res.arrayBuffer());

  // Auth/error pages sometimes come back as 200 JSON/HTML
  const asText = buffer.slice(0, 300).toString("utf8");
  if (
    contentType.includes("application/json") ||
    contentType.includes("text/html") ||
    looksLikeAuthError(asText) ||
    looksLikeNotFound(asText)
  ) {
    if (looksLikeAuthError(asText)) {
      throw classifyApiFailure("export", 401, asText, { projectId, exePlanId });
    }
    if (looksLikeNotFound(asText)) {
      throw classifyApiFailure("export", 404, asText, { projectId, exePlanId });
    }
    throw new ReporterError(
      "EXTRACT_FAILED",
      `Unable to extract data for execution plan ${exePlanId}. The server did not return a valid Excel file.`,
      bodySnippet(asText)
    );
  }

  // xlsx files are ZIP packages starting with PK
  if (buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
    throw new ReporterError(
      "EXTRACT_FAILED",
      `Unable to extract data for execution plan ${exePlanId}. The downloaded file is not a valid Excel export.`,
      `bytes=${buffer.length}`
    );
  }

  const disposition = res.headers.get("content-disposition") || "";
  const match = disposition.match(/filename\*?=(?:UTF-8''|")?([^\";]+)/i);
  const fileName = match
    ? decodeURIComponent(match[1].replace(/"/g, ""))
    : `EP-${exePlanId}_TrackingDetails.xlsx`;

  ensureDir(DOWNLOADS_DIR);
  const outPath = path.join(
    DOWNLOADS_DIR,
    `EP-${exePlanId}_${Date.now()}_${fileName}`
  );
  fs.writeFileSync(outPath, buffer);
  return outPath;
}

async function downloadDefectExport(token, projectId, timezone, offset) {
  let res;
  try {
    res = await fetch(DEFECT_EXPORT_URL, {
      method: "POST",
      headers: {
        accept: "*/*",
        authorization: `Bearer ${token}`,
        "content-type": "application/json",
        origin: ORIGIN,
        referer: `${ORIGIN}/user/defect`,
        "user-agent":
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36",
      },
      body: JSON.stringify({
        data: {
          searchFields: [
            { column: "id", sort: "dsc" },
            { column: "deleted", value: false, regEx: false },
            { column: "projectId", value: projectId, regEx: false },
          ],
          startIndex: 0,
          limit: 10000,
          projectId,
          attachFlag: false,
        },
        timezone,
        offset,
      }),
    });
  } catch (err) {
    throw new ReporterError(
      "NETWORK_ERROR",
      `Unable to reach SimplifyQA while downloading defects for project ${projectId}. Please check your internet connection and try again.`,
      err.message
    );
  }

  if (!res.ok) {
    const body = await res.text();
    throw classifyApiFailure("export", res.status, body, { projectId });
  }

  const contentType = (res.headers.get("content-type") || "").toLowerCase();
  const buffer = Buffer.from(await res.arrayBuffer());
  const asText = buffer.slice(0, 300).toString("utf8");
  if (
    contentType.includes("application/json") ||
    contentType.includes("text/html") ||
    looksLikeAuthError(asText) ||
    looksLikeNotFound(asText)
  ) {
    if (looksLikeAuthError(asText)) {
      throw classifyApiFailure("export", 401, asText, { projectId });
    }
    throw new ReporterError(
      "EXTRACT_FAILED",
      `Unable to extract defect data for project ${projectId}. The server did not return a valid Excel file.`,
      bodySnippet(asText)
    );
  }

  if (buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
    throw new ReporterError(
      "EXTRACT_FAILED",
      `Unable to extract defect data for project ${projectId}. The downloaded file is not a valid Excel export.`,
      `bytes=${buffer.length}`
    );
  }

  const disposition = res.headers.get("content-disposition") || "";
  const match = disposition.match(/filename\*?=(?:UTF-8''|")?([^\";]+)/i);
  const fileName = match
    ? decodeURIComponent(match[1].replace(/"/g, ""))
    : `Defects_${projectId}.xlsx`;

  ensureDir(DOWNLOADS_DIR);
  const outPath = path.join(
    DOWNLOADS_DIR,
    `Defects_${projectId}_${Date.now()}_${fileName}`
  );
  fs.writeFileSync(outPath, buffer);
  return outPath;
}

function emptyDefectCounts() {
  return { closed: 0, deferred: 0, fixed: 0, pending: 0 };
}

function pickDefectField(row, names) {
  for (const key of Object.keys(row || {})) {
    const norm = key.trim().toLowerCase();
    if (names.includes(norm)) return row[key];
  }
  return "";
}

/**
 * Load defect export and aggregate by normalized entity + module.
 * Multi-entity values count toward every listed entity.
 */
function aggregateDefectExport(filePath) {
  const wb = XLSX.readFile(filePath);
  if (!wb.SheetNames.length) {
    throw new ReporterError(
      "EXTRACT_FAILED",
      "Unable to extract defect data because the export workbook has no sheets."
    );
  }
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
    defval: "",
  });

  /** @type {Map<string, Map<string, {closed:number,deferred:number,fixed:number,pending:number}>>} */
  const byEntity = new Map();
  const unknownStates = new Map();
  let mappedRows = 0;
  let skippedUnknownState = 0;

  for (const row of rows) {
    const moduleRaw = pickDefectField(row, ["module"]);
    const entityRaw = pickDefectField(row, ["entity"]);
    const stateRaw = pickDefectField(row, ["state", "status"]);
    const moduleName = resolveDefectModuleName(moduleRaw);
    const entities = parseDefectEntities(entityRaw);
    const bucket = mapDefectState(stateRaw);

    if (!moduleName || !entities.length) continue;

    if (!bucket) {
      skippedUnknownState += 1;
      const label = String(stateRaw || "").trim() || "(blank)";
      unknownStates.set(label, (unknownStates.get(label) || 0) + 1);
      continue;
    }

    mappedRows += 1;
    for (const entity of entities) {
      const entityKey = entity.toLowerCase();
      if (!byEntity.has(entityKey)) byEntity.set(entityKey, new Map());
      const modMap = byEntity.get(entityKey);
      if (!modMap.has(moduleName)) modMap.set(moduleName, emptyDefectCounts());
      modMap.get(moduleName)[bucket] += 1;
    }
  }

  return {
    byEntity,
    mappedRows,
    skippedUnknownState,
    unknownStates: [...unknownStates.entries()].map(([state, count]) => ({
      state,
      count,
    })),
    exportRows: rows.length,
  };
}

function discoverDefectSections(sheet) {
  const sections = [];
  for (let r = 1; r <= sheet.rowCount; r++) {
    for (let c = 1; c <= 4; c++) {
      const text = cellText(sheet.getRow(r).getCell(c).value).trim();
      if (/defect\s+summary/i.test(text)) {
        sections.push({
          title: text,
          titleRow: r,
          entity: entityFromDefectSectionTitle(text),
        });
        break;
      }
    }
  }
  return sections;
}

function detectDefectLayout(sheet, titleRow) {
  const headerRow = titleRow + 1;
  const cols = {};
  sheet.getRow(headerRow).eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const t = cellText(cell.value).trim().toLowerCase();
    if (!t) return;
    if (t === "module") cols.module = colNumber;
    else if (t.startsWith("closed")) cols.closed = colNumber;
    else if (t.startsWith("deferred")) cols.deferred = colNumber;
    else if (t.startsWith("fixed")) cols.fixed = colNumber;
    else if (t.startsWith("pending")) cols.pending = colNumber;
    else if (t === "total") cols.total = colNumber;
  });

  if (
    !cols.module ||
    !cols.closed ||
    !cols.deferred ||
    !cols.fixed ||
    !cols.pending ||
    !cols.total
  ) {
    throw new Error(
      `Could not detect required defect columns at header row ${headerRow}`
    );
  }

  return { headerRow, firstDataRow: headerRow + 1, cols };
}

function buildDefectSectionSummary(sectionModules, entityCounts) {
  const rows = [];

  // Strict template order; only modules with at least one defect.
  for (const name of sectionModules) {
    const counts =
      (entityCounts && entityCounts.get(name)) || emptyDefectCounts();
    const total =
      counts.closed + counts.deferred + counts.fixed + counts.pending;
    if (total <= 0) continue;
    rows.push({ name, ...counts, total });
  }

  const totalSummary = rows.reduce(
    (acc, r) => {
      acc.closed += r.closed;
      acc.deferred += r.deferred;
      acc.fixed += r.fixed;
      acc.pending += r.pending;
      acc.total += r.total;
      return acc;
    },
    emptyDefectCounts()
  );
  totalSummary.total =
    totalSummary.closed +
    totalSummary.deferred +
    totalSummary.fixed +
    totalSummary.pending;

  const denom = totalSummary.total - totalSummary.deferred;
  const closureRate = denom > 0 ? totalSummary.closed / denom : 0;
  const resolutionRate =
    denom > 0 ? (totalSummary.closed + totalSummary.fixed) / denom : 0;

  return { rows, totalSummary, closureRate, resolutionRate };
}

function writeDefectSection(sheet, section, summary) {
  const titleRow = section.titleRow;
  const layout = detectDefectLayout(sheet, titleRow);
  const { cols, firstDataRow } = layout;
  let totalRowNum = findTotalRow(sheet, firstDataRow, cols.module);
  if (totalRowNum < 0) {
    throw new Error(`Total row not found for defect section: ${section.title}`);
  }

  const existingCount = totalRowNum - firstDataRow;
  const needed = summary.rows.length;
  const toInsert = needed - existingCount;

  if (toInsert > 0) {
    const styleSource = sheet.getRow(Math.max(firstDataRow, totalRowNum - 1));
    sheet.spliceRows(totalRowNum, 0, ...Array.from({ length: toInsert }, () => []));
    for (let i = 0; i < toInsert; i++) {
      copyRowStyle(styleSource, sheet.getRow(totalRowNum + i));
    }
    totalRowNum += toInsert;
  } else if (toInsert < 0) {
    sheet.spliceRows(firstDataRow + needed, -toInsert);
    totalRowNum = firstDataRow + needed;
  }

  // If no modules have defects, keep a single blank/zero structure: remove all data rows
  if (needed === 0) {
    // leave no module rows — total row stays at firstDataRow
    // formulas on total still valid with empty sum ranges avoided by writing zeros
  }

  const lastDataRow = needed > 0 ? firstDataRow + needed - 1 : firstDataRow - 1;
  const cClosed = colLetter(cols.closed);
  const cDef = colLetter(cols.deferred);
  const cFixed = colLetter(cols.fixed);
  const cPend = colLetter(cols.pending);
  const cTot = colLetter(cols.total);

  for (let i = 0; i < summary.rows.length; i++) {
    const r = firstDataRow + i;
    const row = sheet.getRow(r);
    const data = summary.rows[i];
    row.getCell(cols.module).value = data.name;
    setNumber(row.getCell(cols.closed), data.closed);
    setNumber(row.getCell(cols.deferred), data.deferred);
    setNumber(row.getCell(cols.fixed), data.fixed);
    setNumber(row.getCell(cols.pending), data.pending);
    setFormula(
      row.getCell(cols.total),
      `SUM(${cClosed}${r}:${cPend}${r})`,
      data.total
    );
    row.commit();
  }

  const totalRow = sheet.getRow(totalRowNum);
  totalRow.getCell(cols.module).value = "Total";
  if (needed > 0) {
    setFormula(
      totalRow.getCell(cols.closed),
      `SUM(${cClosed}${firstDataRow}:${cClosed}${lastDataRow})`,
      summary.totalSummary.closed
    );
    setFormula(
      totalRow.getCell(cols.deferred),
      `SUM(${cDef}${firstDataRow}:${cDef}${lastDataRow})`,
      summary.totalSummary.deferred
    );
    setFormula(
      totalRow.getCell(cols.fixed),
      `SUM(${cFixed}${firstDataRow}:${cFixed}${lastDataRow})`,
      summary.totalSummary.fixed
    );
    setFormula(
      totalRow.getCell(cols.pending),
      `SUM(${cPend}${firstDataRow}:${cPend}${lastDataRow})`,
      summary.totalSummary.pending
    );
    setFormula(
      totalRow.getCell(cols.total),
      `SUM(${cClosed}${totalRowNum}:${cPend}${totalRowNum})`,
      summary.totalSummary.total
    );
  } else {
    setNumber(totalRow.getCell(cols.closed), 0);
    setNumber(totalRow.getCell(cols.deferred), 0);
    setNumber(totalRow.getCell(cols.fixed), 0);
    setNumber(totalRow.getCell(cols.pending), 0);
    setNumber(totalRow.getCell(cols.total), 0);
  }
  totalRow.commit();

  for (let r = totalRowNum + 1; r <= totalRowNum + 4; r++) {
    if (rowHasLabel(sheet, r, "closure rate")) {
      writeDefectRateFormula(
        sheet,
        r,
        cols,
        `${cClosed}${totalRowNum}/(${cTot}${totalRowNum}-${cDef}${totalRowNum})`,
        summary.closureRate
      );
    }
    if (rowHasLabel(sheet, r, "resolution rate")) {
      writeDefectRateFormula(
        sheet,
        r,
        cols,
        `(${cClosed}${totalRowNum}+${cFixed}${totalRowNum})/(${cTot}${totalRowNum}-${cDef}${totalRowNum})`,
        summary.resolutionRate
      );
    }
  }

  // Restore title/rate merges broken by spliceRows (Template 2).
  ensureRowMerge(sheet, titleRow, cols.module, cols.total);
}

/** Write rate formula without leaving cells unmerged. */
function writeDefectRateFormula(sheet, rowNum, cols, formula, result) {
  setFormula(sheet.getRow(rowNum).getCell(cols.fixed), formula, result, "0%");
  sheet.getRow(rowNum).commit();
  ensureRowMerge(sheet, rowNum, cols.closed, cols.deferred);
  ensureRowMerge(sheet, rowNum, cols.fixed, cols.pending);
}

/** Remerge a horizontal range (safe if already merged or broken by spliceRows). */
function ensureRowMerge(sheet, rowNum, fromCol, toCol) {
  if (!fromCol || !toCol || toCol <= fromCol) return;
  try {
    sheet.unMergeCells(rowNum, fromCol, rowNum, toCol);
  } catch {
    // ignore
  }
  try {
    sheet.mergeCells(rowNum, fromCol, rowNum, toCol);
  } catch {
    // ignore
  }
}

function validateExportFile(filePath, exePlanId) {
  try {
    const wb = XLSX.readFile(filePath);
    if (!wb.SheetNames.length) {
      throw new Error("Workbook has no sheets");
    }
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
      header: 1,
      defval: "",
    });
    const hasPlanHeader = rows.some((r) =>
      String(r[0] || "")
        .toLowerCase()
        .includes("execution plan")
    );
    const hasAssigned = rows.some((r) =>
      String(r[0] || "")
        .toLowerCase()
        .includes("assigned date")
    );
    if (!hasPlanHeader && !hasAssigned) {
      throw new Error("Expected execution plan columns were not found");
    }
  } catch (err) {
    if (err instanceof ReporterError) throw err;
    throw new ReporterError(
      "EXTRACT_FAILED",
      `Unable to extract data from execution plan ${exePlanId}. The export file could not be read or is incomplete.`,
      err.message
    );
  }
}

function aggregateExport(filePath, templateModules) {
  const wb = XLSX.readFile(filePath);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  let headerIdx = rows.findIndex((r) =>
    String(r[0]).toLowerCase().includes("assigned date")
  );
  if (headerIdx < 0) headerIdx = 4;

  const dataRows = rows.slice(headerIdx + 1).filter(
    (r) => String(r[2] || "").trim() || String(r[6] || "").trim()
  );

  const stats = {};
  for (const row of dataRows) {
    const name = resolveModuleName(row[6]);
    if (!name) continue;
    if (!stats[name]) stats[name] = emptyCounts();
    stats[name][normalizeStatus(row[8])] += 1;
  }

  const templateNames = new Set(templateModules);
  const ordered = [];

  // Strict template order only — do not auto-add unknown modules to the sheet.
  for (const name of templateModules) {
    ordered.push({
      name,
      ...summarizeModule(stats[name] || emptyCounts()),
    });
  }

  const extraModules = Object.keys(stats)
    .filter((name) => !templateNames.has(name))
    .map((name) => ({
      name,
      ...summarizeModule(stats[name]),
    }));

  const totals = emptyCounts();
  for (const r of ordered) {
    totals.passed += r.passed;
    totals.failed += r.failed;
    totals.blocked += r.blocked;
    totals.inProgress += r.inProgress;
    totals.notExecuted += r.notExecuted;
  }
  const totalSummary = summarizeModule(totals);

  const extraTotal = extraModules.reduce((sum, r) => sum + r.total, 0);

  return {
    exportRows: dataRows.length,
    extraModules,
    extraTotal,
    rows: ordered,
    totalSummary,
    executionRate: rate(
      totalSummary.passed + totalSummary.failed,
      totalSummary.total - totalSummary.blocked
    ),
    overallPassRate: rate(
      totalSummary.passed,
      totalSummary.total - totalSummary.blocked
    ),
  };
}

function colLetter(n) {
  let s = "";
  let num = n;
  while (num > 0) {
    const m = (num - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    num = Math.floor((num - 1) / 26);
  }
  return s;
}

function findTitleRow(sheet, title) {
  const normalized = title.trim().toLowerCase();
  for (let r = 1; r <= sheet.rowCount; r++) {
    for (let c = 1; c <= 3; c++) {
      const text = cellText(sheet.getRow(r).getCell(c).value)
        .trim()
        .toLowerCase();
      if (text === normalized) return r;
    }
  }
  return -1;
}

function detectLayout(sheet, titleRow) {
  const headerRow = titleRow + 1;
  const cols = {};
  sheet.getRow(headerRow).eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const t = cellText(cell.value).trim().toLowerCase();
    if (!t) return;
    if (t.includes("lead champion")) cols.champion = colNumber;
    else if (t === "module") cols.module = colNumber;
    else if (t.startsWith("passed")) cols.passed = colNumber;
    else if (t.startsWith("failed")) cols.failed = colNumber;
    else if (t.startsWith("blocked")) cols.blocked = colNumber;
    else if (t.includes("progress")) cols.inProgress = colNumber;
    else if (t.includes("not executed")) cols.notExecuted = colNumber;
    else if (t === "total") cols.total = colNumber;
    else if (t.includes("pass rate")) cols.passRate = colNumber;
  });

  if (!cols.module || !cols.passed || !cols.total) {
    throw new Error(
      `Could not detect required columns at header row ${headerRow}`
    );
  }

  return { headerRow, firstDataRow: headerRow + 1, cols };
}

function findTotalRow(sheet, firstDataRow, moduleCol) {
  for (let r = firstDataRow; r < firstDataRow + 80; r++) {
    if (
      cellText(sheet.getRow(r).getCell(moduleCol).value).trim().toLowerCase() ===
      "total"
    ) {
      return r;
    }
  }
  return -1;
}

function readTemplateModules(sheet, firstDataRow, totalRowNum, moduleCol) {
  const modules = [];
  for (let r = firstDataRow; r < totalRowNum; r++) {
    const name = cellText(sheet.getRow(r).getCell(moduleCol).value).trim();
    if (name) modules.push(name);
  }
  return modules;
}

function setNumber(cell, value) {
  cell.value = Number(value) || 0;
}

function setFormula(cell, formula, result, numFmt) {
  cell.value = { formula, result: result == null ? 0 : result };
  if (numFmt) cell.numFmt = numFmt;
}

function copyRowStyle(fromRow, toRow) {
  toRow.height = fromRow.height;
  fromRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const target = toRow.getCell(colNumber);
    if (cell.style) target.style = Object.assign({}, cell.style);
  });
}

/** ExcelJS spliceRows breaks shared formulas; convert them to plain values first. */
function neutralizeSharedFormulas(sheet) {
  sheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const v = cell.value;
      if (!v || typeof v !== "object") return;
      if (v.sharedFormula != null || v.shareType === "shared" || v.ref) {
        if (typeof v.result !== "undefined") cell.value = v.result;
        else if (v.formula) cell.value = { formula: v.formula, result: 0 };
        else cell.value = 0;
      }
    });
  });
}

function rowHasLabel(sheet, rowNum, labelPrefix) {
  const row = sheet.getRow(rowNum);
  let found = false;
  row.eachCell({ includeEmpty: false }, (cell) => {
    const t = cellText(cell.value).trim().toLowerCase();
    if (t.startsWith(labelPrefix)) found = true;
  });
  return found;
}

function writeSection(sheet, sectionTitle, summary) {
  const titleRow = findTitleRow(sheet, sectionTitle);
  if (titleRow < 0) {
    throw new Error(`Section title not found in template: ${sectionTitle}`);
  }

  const layout = detectLayout(sheet, titleRow);
  const { cols, firstDataRow } = layout;
  let totalRowNum = findTotalRow(sheet, firstDataRow, cols.module);
  if (totalRowNum < 0) {
    throw new Error(`Total row not found for section: ${sectionTitle}`);
  }

  const existingCount = totalRowNum - firstDataRow;
  const needed = summary.rows.length;
  const toInsert = needed - existingCount;

  if (toInsert > 0) {
    const styleSource = sheet.getRow(Math.max(firstDataRow, totalRowNum - 1));
    sheet.spliceRows(totalRowNum, 0, ...Array.from({ length: toInsert }, () => []));
    for (let i = 0; i < toInsert; i++) {
      copyRowStyle(styleSource, sheet.getRow(totalRowNum + i));
    }
    totalRowNum += toInsert;
  } else if (toInsert < 0) {
    sheet.spliceRows(firstDataRow + needed, -toInsert);
    totalRowNum = firstDataRow + needed;
  }

  const lastDataRow = firstDataRow + needed - 1;
  const cPass = colLetter(cols.passed);
  const cFail = colLetter(cols.failed);
  const cBlock = colLetter(cols.blocked);
  const cIP = colLetter(cols.inProgress);
  const cNE = colLetter(cols.notExecuted);
  const cTot = colLetter(cols.total);

  for (let i = 0; i < summary.rows.length; i++) {
    const r = firstDataRow + i;
    const row = sheet.getRow(r);
    const data = summary.rows[i];

    row.getCell(cols.module).value = data.name;
    setNumber(row.getCell(cols.passed), data.passed);
    setNumber(row.getCell(cols.failed), data.failed);
    setNumber(row.getCell(cols.blocked), data.blocked);
    setNumber(row.getCell(cols.inProgress), data.inProgress);
    setNumber(row.getCell(cols.notExecuted), data.notExecuted);
    setFormula(
      row.getCell(cols.total),
      `SUM(${cPass}${r}:${cNE}${r})`,
      data.total
    );
    if (cols.passRate) {
      setFormula(
        row.getCell(cols.passRate),
        `${cPass}${r}/(${cTot}${r}-${cBlock}${r})`,
        data.passRate,
        "0%"
      );
    }
    row.commit();
  }

  const totalRow = sheet.getRow(totalRowNum);
  if (cols.champion) totalRow.getCell(cols.champion).value = "";
  totalRow.getCell(cols.module).value = "Total";
  setFormula(
    totalRow.getCell(cols.passed),
    `SUM(${cPass}${firstDataRow}:${cPass}${lastDataRow})`,
    summary.totalSummary.passed
  );
  setFormula(
    totalRow.getCell(cols.failed),
    `SUM(${cFail}${firstDataRow}:${cFail}${lastDataRow})`,
    summary.totalSummary.failed
  );
  setFormula(
    totalRow.getCell(cols.blocked),
    `SUM(${cBlock}${firstDataRow}:${cBlock}${lastDataRow})`,
    summary.totalSummary.blocked
  );
  setFormula(
    totalRow.getCell(cols.inProgress),
    `SUM(${cIP}${firstDataRow}:${cIP}${lastDataRow})`,
    summary.totalSummary.inProgress
  );
  setFormula(
    totalRow.getCell(cols.notExecuted),
    `SUM(${cNE}${firstDataRow}:${cNE}${lastDataRow})`,
    summary.totalSummary.notExecuted
  );
  setFormula(
    totalRow.getCell(cols.total),
    `SUM(${cPass}${totalRowNum}:${cNE}${totalRowNum})`,
    summary.totalSummary.total
  );
  if (cols.passRate) {
    setFormula(
      totalRow.getCell(cols.passRate),
      `${cPass}${totalRowNum}/(${cTot}${totalRowNum}-${cBlock}${totalRowNum})`,
      summary.totalSummary.passRate,
      "0%"
    );
  }
  totalRow.commit();

  for (let r = totalRowNum + 1; r <= totalRowNum + 4; r++) {
    if (rowHasLabel(sheet, r, "execution rate")) {
      setFormula(
        sheet.getRow(r).getCell(cols.inProgress),
        `(${cPass}${totalRowNum}+${cFail}${totalRowNum})/(${cTot}${totalRowNum}-${cBlock}${totalRowNum})`,
        summary.executionRate,
        "0%"
      );
      sheet.getRow(r).commit();
      ensureRowMerge(sheet, r, cols.passed, cols.blocked);
      ensureRowMerge(sheet, r, cols.inProgress, cols.passRate || cols.total);
    }
    if (rowHasLabel(sheet, r, "overall pass rate")) {
      setFormula(
        sheet.getRow(r).getCell(cols.inProgress),
        `${cPass}${totalRowNum}/(${cTot}${totalRowNum}-${cBlock}${totalRowNum})`,
        summary.overallPassRate,
        "0%"
      );
      sheet.getRow(r).commit();
      ensureRowMerge(sheet, r, cols.passed, cols.blocked);
      ensureRowMerge(sheet, r, cols.inProgress, cols.passRate || cols.total);
    }
  }

  const titleStart = cols.champion || cols.module;
  let titleEnd = cols.passRate || cols.total;
  sheet.getRow(titleRow).eachCell({ includeEmpty: false }, (_cell, c) => {
    if (c > titleEnd) titleEnd = c;
  });
  ensureRowMerge(sheet, titleRow, titleStart, titleEnd);
}

function outputStamp() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  let hours = now.getHours();
  const ampm = hours >= 12 ? "PM" : "AM";
  hours = hours % 12;
  if (hours === 0) hours = 12;
  return `${pad(now.getDate())}-${pad(now.getMonth() + 1)}-${now.getFullYear()} ${pad(hours)}-${pad(now.getMinutes())}-${pad(now.getSeconds())} ${ampm}`;
}

function dateSheetBase() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${pad(now.getDate())}-${pad(now.getMonth() + 1)}-${now.getFullYear()}`;
}

function sanitizeSheetName(name) {
  return String(name)
    .replace(/[:\\/?*\[\]]/g, "-")
    .trim()
    .slice(0, 31);
}

function uniqueSheetName(existingNames) {
  const existing = new Set(
    (existingNames || []).map((n) => String(n).toLowerCase())
  );
  const dateOnly = sanitizeSheetName(dateSheetBase());
  if (!existing.has(dateOnly.toLowerCase())) return dateOnly;

  // Same date already exists -> add 12-hour timestamp
  let name = sanitizeSheetName(outputStamp());
  let i = 2;
  while (existing.has(name.toLowerCase())) {
    const suffix = `-${i}`;
    name = sanitizeSheetName(outputStamp().slice(0, 31 - suffix.length) + suffix);
    i += 1;
  }
  return name;
}

function cloneCellValue(cell) {
  const value = cell.value;
  if (value == null || typeof value !== "object") return value;
  if (value instanceof Date) return new Date(value.getTime());

  // ExcelJS may expose the resolved formula even when the stored value is a shared formula.
  // Always copy as a standalone formula so appending sheets doesn't reuse a broken shared master.
  const formula = cell.formula || (value.formula != null ? String(value.formula) : null);
  if (formula) {
    const cloned = { formula: String(formula) };
    if (value.result !== undefined) cloned.result = value.result;
    else if (typeof cell.value === "number") cloned.result = cell.value;
    return cloned;
  }

  if (value.sharedFormula != null) {
    return value.result !== undefined ? value.result : null;
  }
  if (value.richText) {
    return { richText: value.richText.map((part) => Object.assign({}, part)) };
  }
  if (value.text != null || value.hyperlink != null) {
    return { text: value.text, hyperlink: value.hyperlink };
  }
  if (value.error) return { error: value.error };
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return value.result !== undefined ? value.result : null;
  }
}

function copyWorksheet(source, target) {
  target.properties = Object.assign({}, source.properties || {});
  if (source.views) target.views = JSON.parse(JSON.stringify(source.views));

  source.columns.forEach((col, idx) => {
    if (!col) return;
    const tCol = target.getColumn(idx + 1);
    if (col.width != null) tCol.width = col.width;
    if (col.hidden != null) tCol.hidden = col.hidden;
    if (col.style) tCol.style = JSON.parse(JSON.stringify(col.style));
  });

  source.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const tRow = target.getRow(rowNumber);
    if (row.height != null) tRow.height = row.height;
    if (row.hidden != null) tRow.hidden = row.hidden;
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const tCell = tRow.getCell(colNumber);
      tCell.value = cloneCellValue(cell);
      if (cell.style) {
        tCell.style = JSON.parse(JSON.stringify(cell.style));
      }
      if (cell.numFmt) tCell.numFmt = cell.numFmt;
    });
    tRow.commit();
  });

  const merges =
    (source.model && source.model.merges) ||
    Object.keys(source._merges || {}).map((k) => {
      const m = source._merges[k];
      if (!m) return null;
      return m.range || m;
    }).filter(Boolean);

  for (const merge of merges) {
    try {
      if (typeof merge === "string") target.mergeCells(merge);
      else if (merge && merge.top != null) {
        target.mergeCells(merge.top, merge.left, merge.bottom, merge.right);
      }
    } catch {
      // ignore invalid/duplicate merges
    }
  }
}

function activateWorksheet(workbook, sheet) {
  const idx = workbook.worksheets.findIndex((ws) => ws.id === sheet.id);
  const activeTab = idx >= 0 ? idx : workbook.worksheets.length - 1;
  // firstSheet scrolls the tab bar so the active sheet is visible (not stuck on sheet 1).
  workbook.views = [
    {
      x: 0,
      y: 0,
      width: 25000,
      height: 15000,
      firstSheet: activeTab,
      activeTab,
      visibility: "visible",
    },
  ];
}

async function saveWorkbookToMaster(filledWorkbook, masterPath, log) {
  ensureDir(path.dirname(masterPath));
  const sourceSheet = filledWorkbook.worksheets[0];
  if (!sourceSheet) {
    throw new ReporterError(
      "EXTRACT_FAILED",
      "Unable to save the report because the filled workbook has no sheet."
    );
  }

  if (!fs.existsSync(masterPath)) {
    const sheetName = uniqueSheetName([]);
    sourceSheet.name = sheetName;
    activateWorksheet(filledWorkbook, sourceSheet);
    await filledWorkbook.xlsx.writeFile(masterPath);
    log(`\nCreated workbook: ${masterPath}`);
    log(`Added sheet: ${sheetName}`);
    return { masterPath, sheetName, created: true };
  }

  // Round-trip the filled workbook so formulas are stored as normal formulas
  // (not fragile in-memory shared-formula refs) before copying into the master.
  const clean = new ExcelJS.Workbook();
  await clean.xlsx.load(await filledWorkbook.xlsx.writeBuffer());
  const cleanSheet = clean.worksheets[0];

  const master = new ExcelJS.Workbook();
  await master.xlsx.readFile(masterPath);
  const sheetName = uniqueSheetName(master.worksheets.map((ws) => ws.name));
  const target = master.addWorksheet(sheetName);
  copyWorksheet(cleanSheet, target);
  activateWorksheet(master, target);
  await master.xlsx.writeFile(masterPath);
  log(`\nUpdated workbook: ${masterPath}`);
  log(`Added sheet: ${sheetName}`);
  return { masterPath, sheetName, created: false };
}

function writeRunLog(logPath, lines) {
  fs.writeFileSync(logPath, lines.join("\n"), "utf8");
}

function clearPreviousLogs() {
  ensureDir(LOGS_DIR);
  for (const name of fs.readdirSync(LOGS_DIR)) {
    if (/\.log(\.txt)?$/i.test(name) || name.endsWith(".txt")) {
      fs.unlinkSync(path.join(LOGS_DIR, name));
    }
  }
}

/** Delete files in downloads/ older than maxAgeDays (default 8). */
function cleanupOldDownloads(log, maxAgeDays = 8) {
  ensureDir(DOWNLOADS_DIR);
  const cutoff = Date.now() - maxAgeDays * 24 * 60 * 60 * 1000;
  let removed = 0;
  for (const name of fs.readdirSync(DOWNLOADS_DIR)) {
    const full = path.join(DOWNLOADS_DIR, name);
    let stat;
    try {
      stat = fs.statSync(full);
    } catch {
      continue;
    }
    if (!stat.isFile()) continue;
    if (stat.mtimeMs > cutoff) continue;
    try {
      fs.unlinkSync(full);
      removed += 1;
    } catch (err) {
      if (log) log(`WARN: could not delete old download ${name}: ${err.message}`);
    }
  }
  if (log && removed) {
    log(
      `Cleaned downloads: removed ${removed} file(s) older than ${maxAgeDays} day(s).`
    );
  }
  return removed;
}

async function runReport(args, runtime, log, logLines, onProgress) {
  const progress = (id, label) => {
    if (typeof onProgress === "function") {
      try {
        onProgress({ id, label, at: new Date().toISOString() });
      } catch {
        /* ignore UI progress errors */
      }
    }
  };

  if (!fs.existsSync(runtime.templatePath)) {
    throw new ReporterError(
      "TEMPLATE_NOT_FOUND",
      `Template file not found: ${runtime.templateRelative || runtime.templatePath}. Please check TEMPLATE_CHOICE / TEMPLATE_n in config/application.properties.`
    );
  }

  progress("prepare", "Preparing template…");
  cleanupOldDownloads(log, 8);
  log(`Template choice: ${runtime.templateChoice} -> ${runtime.templateRelative}`);
  log(`Output file: ${runtime.outputRelative}`);
  log(`Project ID: ${runtime.projectId}`);
  log(`Plans configured: ${runtime.plans.length}`);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(runtime.templatePath);
  const sheet = workbook.worksheets[0];
  neutralizeSharedFormulas(sheet);
  const discovered = discoverStatusSections(sheet);

  if (!discovered.length) {
    throw new ReporterError(
      "TEMPLATE_INVALID",
      "No Module-wise Daily Status sections were found in the chosen template. Please pick a valid template."
    );
  }

  log(`Status sections in template: ${discovered.length}`);
  discovered.forEach((s, i) => log(`  [${i + 1}] ${s.title}`));

  const { mapped, overflow } = mapPlansToSections(runtime.plans, discovered);

  log("\nPlan -> section mapping:");
  for (const p of mapped) {
    log(
      `  EXE_PLAN_ID_${p.index}=${p.exePlanId} -> [${p.sectionOrdinal}] ${p.sectionTitle}`
    );
  }
  if (overflow.length) {
    log("\nOVERFLOW (more plans than template sections):");
    for (const p of overflow) {
      log(`  EXE_PLAN_ID_${p.index}=${p.exePlanId} (will NOT be filled)`);
    }
  }

  if (args.dryRun) {
    log("\nDry-run complete (no download / no Excel write).");
    clearPreviousLogs();
    const stamp = outputStamp();
    const logPath = path.join(LOGS_DIR, `FMS Status tracker - ${stamp}.log.txt`);
    writeRunLog(logPath, logLines);
    log(`Run log: ${logPath}`);
    if (overflow.length) {
      const err = new ReporterError(
        "PLAN_SECTION_OVERFLOW",
        `Too many execution plans for this template. The template has ${discovered.length} status section(s), but ${runtime.plans.length} plan(s) were provided.`
      );
      err.logPath = logPath;
      throw err;
    }
    return {
      dryRun: true,
      outputFile: null,
      logPath,
      sheetName: null,
    };
  }

  const needsApi = mapped.some((p) => !args.localByPlan[String(p.index)]);
  const token = needsApi ? getBearerToken() : null;

  if (token && !args.skipSwitch) {
    log(`\nSwitching project to projectId=${runtime.projectId}...`);
    await switchProject(token, runtime.projectId);
    log("Project switch OK");
  } else if (!needsApi) {
    log("\nLocal mode: skipping auth/project switch");
  }

  const fileJobs = mapped.map(async (plan) => {
    try {
      const local = args.localByPlan[String(plan.index)];
      if (local) {
        const filePath = resolvePathMaybe(local);
        if (!fs.existsSync(filePath)) {
          throw new ReporterError(
            "EXTRACT_FAILED",
            `Unable to extract data for execution plan ${plan.exePlanId}. Local file was not found: ${filePath}`
          );
        }
        validateExportFile(filePath, plan.exePlanId);
        return { plan, filePath };
      }
      const filePath = await downloadExport(
        token,
        runtime.projectId,
        plan.exePlanId,
        runtime.timezone,
        runtime.offset
      );
      validateExportFile(filePath, plan.exePlanId);
      return { plan, filePath };
    } catch (err) {
      if (err instanceof ReporterError) throw err;
      throw new ReporterError(
        "EXTRACT_FAILED",
        `Unable to extract data for execution plan ${plan.exePlanId}. ${err.message || "Unexpected error during download/read."}`,
        err.stack
      );
    }
  });

  log(`\nFetching ${mapped.length} execution plan export(s) in parallel...`);
  progress("download_ep", "Downloading execution plan exports…");
  const exports = await Promise.all(fileJobs);
  for (const item of exports) {
    log(`  Plan ${item.plan.index} (EP ${item.plan.exePlanId}): ${item.filePath}`);
  }

  progress("fill_status", "Filling status sections…");
  for (const item of exports) {
    const { plan, filePath } = item;
    try {
      const titleRow = findTitleRow(sheet, plan.sectionTitle);
      if (titleRow < 0) {
        throw new ReporterError(
          "TEMPLATE_INVALID",
          `Unable to fill the report because section "${plan.sectionTitle}" was not found in the template.`
        );
      }
      const layout = detectLayout(sheet, titleRow);
      const totalRowNum = findTotalRow(
        sheet,
        layout.firstDataRow,
        layout.cols.module
      );
      const templateModules = readTemplateModules(
        sheet,
        layout.firstDataRow,
        totalRowNum,
        layout.cols.module
      );
      const summary = aggregateExport(filePath, templateModules);
      writeSection(sheet, plan.sectionTitle, summary);

      const accounted = summary.totalSummary.total + summary.extraTotal;
      const ok = accounted === summary.exportRows;
      log(
        `\nPlan ${plan.index} EP-${plan.exePlanId}: export=${summary.exportRows} templateTotal=${summary.totalSummary.total} extra=${summary.extraTotal} tally=${ok ? "OK" : "MISMATCH"}`
      );
      if (summary.extraModules.length) {
        log(
          `\nALERT: Extra EP modules not in template (not added to sheet) for EP-${plan.exePlanId}:`
        );
        for (const r of summary.extraModules) {
          log(
            `  - ${r.name}: P=${r.passed} F=${r.failed} B=${r.blocked} T=${r.total}`
          );
        }
        log(
          "Add these modules to the template (in the order you want) if they should appear in the report."
        );
      }
      for (const r of summary.rows) {
        log(
          `  ${r.name}: P=${r.passed} F=${r.failed} B=${r.blocked} T=${r.total}`
        );
      }
      if (!ok) {
        throw new ReporterError(
          "TALLY_MISMATCH",
          `Unable to complete the report for execution plan ${plan.exePlanId}. Extracted counts do not match the export total (${summary.exportRows} vs template ${summary.totalSummary.total} + extra ${summary.extraTotal}).`
        );
      }
    } catch (err) {
      if (err instanceof ReporterError) throw err;
      throw new ReporterError(
        "EXTRACT_FAILED",
        `Unable to extract/fill data for execution plan ${plan.exePlanId}. ${err.message || "Unexpected error."}`,
        err.stack
      );
    }
  }

  // ---- Defects (optional; default on) ----
  const includeDefects = args.includeDefects !== false;
  if (includeDefects) {
    progress("defects", "Downloading and filling defects…");
    const defectSections = discoverDefectSections(sheet);
    if (!defectSections.length) {
      log("\nDefects: no Defect Summary sections found in template (skipped).");
    } else {
      log(`\nDefect sections in template: ${defectSections.length}`);
      defectSections.forEach((s, i) =>
        log(
          `  [${i + 1}] ${s.title}${s.entity ? ` (entity=${s.entity})` : " (entity unknown)"}`
        )
      );

      if (args.dryRun) {
        log("Dry-run: defect export/fill skipped.");
      } else {
        let defectPath = null;
        if (args.localDefects) {
          defectPath = resolvePathMaybe(args.localDefects);
          if (!fs.existsSync(defectPath)) {
            throw new ReporterError(
              "EXTRACT_FAILED",
              `Unable to extract defect data. Local file was not found: ${defectPath}`
            );
          }
          log(`Using local defect export: ${defectPath}`);
        } else {
          const defectToken = token || getBearerToken();
          log(`\nDownloading defect export for projectId=${runtime.projectId}...`);
          defectPath = await downloadDefectExport(
            defectToken,
            runtime.projectId,
            runtime.timezone,
            runtime.offset
          );
          log(`Defect export: ${defectPath}`);
        }

        const defectAgg = aggregateDefectExport(defectPath);
        log(
          `Defects mapped=${defectAgg.mappedRows} exportRows=${defectAgg.exportRows} unknownState=${defectAgg.skippedUnknownState}`
        );
        if (defectAgg.unknownStates.length) {
          log("\nALERT: Other defect statuses were found and ignored:");
          for (const u of defectAgg.unknownStates) {
            log(`  - ${u.state}: ${u.count}`);
          }
          log(
            "Defined mapping used: Pending=New+Reopened, Closed=Resolved+Closed, Deferred=Deferred, Fixed=Fixed."
          );
        }

        // Snapshot all template defect modules (before fills change rows).
        const allDefectTemplateModules = new Set();
        for (const section of defectSections) {
          try {
            const layout = detectDefectLayout(sheet, section.titleRow);
            const totalRowNum = findTotalRow(
              sheet,
              layout.firstDataRow,
              layout.cols.module
            );
            if (totalRowNum < 0) continue;
            for (const name of readTemplateModules(
              sheet,
              layout.firstDataRow,
              totalRowNum,
              layout.cols.module
            )) {
              allDefectTemplateModules.add(name);
            }
          } catch {
            // ignore sections we cannot read yet
          }
        }

        for (const [entityKey, modMap] of defectAgg.byEntity.entries()) {
          const extras = [];
          for (const [name, counts] of modMap.entries()) {
            if (allDefectTemplateModules.has(name)) continue;
            const total =
              counts.closed + counts.deferred + counts.fixed + counts.pending;
            if (total <= 0) continue;
            extras.push({ name, ...counts, total });
          }
          if (!extras.length) continue;
          log(
            `\nALERT: Extra defect modules not in template (not added to sheet) for [${entityKey}]:`
          );
          for (const r of extras) {
            log(
              `  - ${r.name}: C=${r.closed} D=${r.deferred} F=${r.fixed} P=${r.pending} T=${r.total}`
            );
          }
          log(
            "Add these modules to the template (in the order you want) if they should appear in the report."
          );
        }

        // Top-down by index; re-discover after each write so spliceRows cannot
        // leave lower sections with stale formula row numbers (Template 2).
        const sectionCount = defectSections.length;
        for (let i = 0; i < sectionCount; i++) {
          neutralizeSharedFormulas(sheet);
          const liveSections = discoverDefectSections(sheet);
          const section = liveSections[i];
          if (!section) {
            log(`\nSkipping defect section index ${i + 1}: no longer found after prior fills.`);
            continue;
          }
          if (!section.entity) {
            log(`\nSkipping defect section (cannot detect entity): ${section.title}`);
            continue;
          }
          const layout = detectDefectLayout(sheet, section.titleRow);
          const totalRowNum = findTotalRow(
            sheet,
            layout.firstDataRow,
            layout.cols.module
          );
          if (totalRowNum < 0) {
            throw new ReporterError(
              "TEMPLATE_INVALID",
              `Unable to fill defects because Total row was not found under "${section.title}".`
            );
          }
          const sectionModules = readTemplateModules(
            sheet,
            layout.firstDataRow,
            totalRowNum,
            layout.cols.module
          );
          const entityCounts =
            defectAgg.byEntity.get(section.entity.toLowerCase()) || new Map();
          const summary = buildDefectSectionSummary(sectionModules, entityCounts);
          writeDefectSection(sheet, section, summary);
          log(
            `\nDefects [${section.entity}] "${section.title}": modules=${summary.rows.length} total=${summary.totalSummary.total}`
          );
          for (const r of summary.rows) {
            log(
              `  ${r.name}: C=${r.closed} D=${r.deferred} F=${r.fixed} P=${r.pending} T=${r.total}`
            );
          }
        }
      }
    }
  } else {
    log("\nDefects: skipped (Include defects unchecked).");
  }

  ensureDir(OUTPUT_DIR);
  clearPreviousLogs();
  const stamp = outputStamp();
  const logPath = path.join(LOGS_DIR, `FMS Status tracker - ${stamp}.log.txt`);

  progress("save", "Saving Excel workbook…");
  const saved = await saveWorkbookToMaster(
    workbook,
    runtime.outputFile,
    log
  );
  progress("done", "Report ready");
  log(`\nFinal report: ${saved.masterPath} (sheet: ${saved.sheetName})`);

  if (overflow.length) {
    log(
      `\nERROR: ${overflow.length} plan(s) could not be filled (template has only ${discovered.length} status section(s)).`
    );
    log("Filled available sections, then failing the run.");
    writeRunLog(logPath, logLines);
    log(`Run log: ${logPath}`);
    const err = new ReporterError(
      "PLAN_SECTION_OVERFLOW",
      `Too many execution plans for this template. Filled ${mapped.length} section(s), but ${overflow.length} plan(s) could not be mapped.`
    );
    err.logPath = logPath;
    err.outputFile = saved.masterPath;
    err.sheetName = saved.sheetName;
    throw err;
  }

  writeRunLog(logPath, logLines);
  log(`Run log: ${logPath}`);
  return {
    dryRun: false,
    outputFile: saved.masterPath,
    logPath,
    sheetName: saved.sheetName,
  };
}

/**
 * Merge UI form values onto properties. If the form has no usable inputs,
 * return the file properties unchanged.
 */
function applyFormOverrides(baseProps, formInput) {
  const form = formInput || {};
  const projectId = String(form.projectId || "").trim();
  const templateChoice = String(form.templateChoice || "").trim();
  const planIds = Array.isArray(form.planIds)
    ? form.planIds.map((id) => String(id || "").trim()).filter(Boolean)
    : [];

  const hasFormInput = Boolean(projectId || templateChoice || planIds.length);
  if (!hasFormInput) return { ...baseProps };

  const props = { ...baseProps };
  if (projectId) props.PROJECT_ID = projectId;
  if (templateChoice) props.TEMPLATE_CHOICE = templateChoice;

  if (planIds.length) {
    for (let i = 1; i <= 100; i++) {
      delete props[`EXE_PLAN_ID_${i}`];
    }
    planIds.forEach((id, idx) => {
      props[`EXE_PLAN_ID_${idx + 1}`] = id;
    });
  }

  return props;
}

function listTemplateChoices(props) {
  const choices = [];
  for (let i = 1; i <= 50; i++) {
    const key = `TEMPLATE_${i}`;
    if (!(key in props) && !DEFAULT_TEMPLATES[i]) continue;
    if (!(key in props) && !props[`OUTPUT_FILE_${i}`] && !DEFAULT_TEMPLATES[i]) {
      continue;
    }
    const hasTemplate = Boolean(props[key] || DEFAULT_TEMPLATES[i]);
    if (!hasTemplate) continue;
    choices.push({
      choice: String(i),
      template: props[key] || DEFAULT_TEMPLATES[i] || "",
      outputFile: props[`OUTPUT_FILE_${i}`] || "",
      ...(TEMPLATE_META[i] || TEMPLATE_META[String(i)] || {
        statusSections: null,
        description: "Custom template",
      }),
    });
  }
  if (!choices.length) {
    choices.push(
      {
        choice: "1",
        template: DEFAULT_TEMPLATES[1],
        outputFile: "",
        ...TEMPLATE_META[1],
      },
      {
        choice: "2",
        template: DEFAULT_TEMPLATES[2],
        outputFile: "",
        ...TEMPLATE_META[2],
      },
      {
        choice: "3",
        template: DEFAULT_TEMPLATES[3],
        outputFile: "",
        ...TEMPLATE_META[3],
      }
    );
  }
  return choices;
}

async function fetchSimplifyQaProjects(token, origin = ORIGIN) {
  const defaultProjects = [
    { id: "2", name: "Financial Management System - Kenya" },
    { id: "5", name: "Financial Management System - Uganda" },
  ];

  if (!token) {
    return defaultProjects.map((p) => ({
      id: String(p.id),
      name: p.name,
      label: `${p.id} — ${p.name}`,
    }));
  }

  const cleanToken = token.startsWith("Bearer ") ? token : `Bearer ${token.trim()}`;
  const headers = {
    "Content-Type": "application/json",
    Accept: "application/json, text/plain, */*",
    Authorization: cleanToken,
    origin: origin,
    referer: `${origin}/user/home`,
    "user-agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36",
  };

  const projectMap = new Map();
  defaultProjects.forEach((p) => projectMap.set(String(p.id), p.name));

  try {
    const searchRes = await fetch(`${origin}/pm/project/search`, {
      method: "POST",
      headers,
      body: JSON.stringify({
        searchFields: [],
        selectFields: ["id", "code", "name", "projectName", "description", "status"],
        startIndex: 0,
        limit: 1000,
      }),
    });
    if (searchRes.ok) {
      const data = await searchRes.json();
      const records =
        (data && (data.data || data.records || data.result || (Array.isArray(data) ? data : []))) || [];
      if (Array.isArray(records)) {
        for (const r of records) {
          const id = String(r.id || r.projectId || "");
          const name = r.name || r.projectName || r.code;
          if (id && name) {
            projectMap.set(id, name);
          }
        }
      }
    }
  } catch {}

  const getEndpoints = ["/pm/project/user/all", "/pm/project/all", "/pm/project/my-projects", "/pm/user/projects"];
  for (const ep of getEndpoints) {
    try {
      const res = await fetch(`${origin}${ep}`, {
        method: "GET",
        headers,
      });
      if (res.ok) {
        const data = await res.json();
        const records =
          (data && (data.data || data.records || data.result || (Array.isArray(data) ? data : []))) || [];
        if (Array.isArray(records)) {
          for (const r of records) {
            const id = String(r.id || r.projectId || "");
            const name = r.name || r.projectName || r.code;
            if (id && name) {
              projectMap.set(id, name);
            }
          }
        }
      }
    } catch {}
  }

  return Array.from(projectMap.entries()).map(([id, name]) => ({
    id,
    name,
    label: `${id} — ${name}`,
  }));
}

function formDefaultsFromProps(props) {
  const planEntries = [];
  for (let i = 1; i <= 100; i++) {
    const key = `EXE_PLAN_ID_${i}`;
    if (!(key in props)) break;
    const value = String(props[key] || "").trim();
    if (!value) break;
    planEntries.push(value);
  }
  return {
    projectId: String(props.PROJECT_ID || "").trim(),
    templateChoice: String(props.TEMPLATE_CHOICE || "1").trim(),
    planIds: planEntries.length ? planEntries : [""],
    templates: listTemplateChoices(props),
  };
}

/**
 * Programmatic entry used by the local web UI / API.
 * @param {{ form?: object, dryRun?: boolean, skipSwitch?: boolean, localByPlan?: object, onProgress?: Function }} options
 */
async function executeReport(options = {}) {
  const form = options.form || {};
  const includeDefects =
    form.includeDefects === false || options.includeDefects === false
      ? false
      : true;

  const args = {
    localByPlan: options.localByPlan || {},
    dryRun: Boolean(options.dryRun),
    skipSwitch: Boolean(options.skipSwitch),
    includeDefects,
    localDefects: options.localDefects || null,
  };
  const logLines = [];
  const log = (msg) => {
    console.log(msg);
    logLines.push(String(msg));
  };

  const baseProps = loadProperties(PROPERTIES_PATH);
  const props = applyFormOverrides(baseProps, options.form);
  const runtime = buildRuntime(props);

  try {
    const result = await runReport(
      args,
      runtime,
      log,
      logLines,
      options.onProgress
    );
    return {
      ok: true,
      ...result,
      logLines,
      alerts: extractAlerts(logLines),
      runtime: {
        projectId: runtime.projectId,
        templateChoice: runtime.templateChoice,
        outputRelative: runtime.outputRelative,
        plans: runtime.plans.map((p) => ({
          index: p.index,
          exePlanId: p.exePlanId,
        })),
      },
    };
  } catch (err) {
    let logPath = err && err.logPath ? err.logPath : null;
    if (!logPath) {
      try {
        clearPreviousLogs();
        const stamp = outputStamp();
        logPath = path.join(LOGS_DIR, `FMS Status tracker - ${stamp}.log.txt`);
        const code = err && err.code ? err.code : "UNEXPECTED_ERROR";
        const details = err && err.details ? String(err.details) : "";
        logLines.push("");
        logLines.push("========== ERROR ==========");
        logLines.push(`Code: ${code}`);
        logLines.push(
          `Message: ${err && err.message ? err.message : String(err)}`
        );
        if (details) logLines.push(`Details: ${details}`);
        writeRunLog(logPath, logLines);
      } catch {
        /* ignore secondary log failure */
      }
    }
    const wrapped = err instanceof ReporterError ? err : new Error(err.message || String(err));
    if (!(err instanceof ReporterError)) {
      wrapped.code = "UNEXPECTED_ERROR";
    }
    wrapped.logPath = logPath;
    wrapped.outputFile = err && err.outputFile ? err.outputFile : null;
    wrapped.logLines = logLines;
    wrapped.alerts = extractAlerts(logLines);
    throw wrapped;
  }
}

/** Pull ALERT blocks from run log lines for the UI banner. */
function extractAlerts(logLines) {
  const alerts = [];
  let current = null;
  for (const raw of logLines || []) {
    const line = String(raw || "");
    const alertIdx = line.indexOf("ALERT:");
    if (alertIdx >= 0) {
      if (current) alerts.push(current);
      current = {
        title: line.slice(alertIdx + "ALERT:".length).trim() || "Alert",
        details: [],
      };
      continue;
    }
    if (!current) continue;
    const trimmed = line.trim();
    if (!trimmed) {
      alerts.push(current);
      current = null;
      continue;
    }
    // Only keep bullet detail lines and the short tip under an ALERT.
    if (/^-\s/.test(trimmed) || /^Add these modules/i.test(trimmed)) {
      current.details.push(trimmed);
      continue;
    }
    if (
      /^Defined mapping used/i.test(trimmed) ||
      /^Other defect statuses/i.test(trimmed)
    ) {
      current.details.push(trimmed);
      continue;
    }
    // Anything else (normal module tallies, plan logs) ends this alert block.
    alerts.push(current);
    current = null;
  }
  if (current) alerts.push(current);
  return alerts;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  try {
    const result = await executeReport({
      dryRun: args.dryRun,
      skipSwitch: args.skipSwitch,
      localByPlan: args.localByPlan,
      includeDefects: args.includeDefects,
      localDefects: args.localDefects,
    });
    if (result && result.outputFile) {
      console.log(`\nDone. Output: ${result.outputFile}`);
    }
  } catch (err) {
    console.error(`\nERROR: ${err && err.message ? err.message : String(err)}`);
    if (err && err.logPath) console.error(`Run log: ${err.logPath}`);
    process.exit(1);
  }
}

module.exports = {
  ROOT,
  PROPERTIES_PATH,
  OUTPUT_DIR,
  LOGS_DIR,
  TEMPLATE_META,
  ReporterError,
  loadProperties,
  buildRuntime,
  applyFormOverrides,
  formDefaultsFromProps,
  fetchSimplifyQaProjects,
  listTemplateChoices,
  executeReport,
  runReport,
  extractAlerts,
  discoverStatusSections,
  getTokenStatus,
  saveBearerToken,
};

if (require.main === module) {
  main().catch((err) => {
    console.error(err.message || err);
    process.exit(1);
  });
}
