const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");
const workstreamFill = require("./workstreamFill");
const DEFECT_FILTER_BY_PROJECT = require("./defectColumnFilters");
const sprint3Defects = require("./sprint3Defects");
const loadEnv = require("../loadEnv");
loadEnv.loadProjectEnv(path.resolve(__dirname, "../../"), false);

const ROOT = path.resolve(__dirname, "../../");
const ORIGIN = "https://app.simplifyqa.ai";
const SWITCH_PROJECT_URL = `${ORIGIN}/pm/fields/8`;
const EXPORT_URL = `${ORIGIN}/tm/executionplan/export`;
const DEFECT_EXPORT_URL = `${ORIGIN}/pm/defect/export`;
/** GET /pm/preference/{userId} — userId 71 is the captured Defects column-preference API. */
const DEFECT_PREFERENCE_USER_ID_DEFAULT = 71;
const PROPERTIES_PATH = path.join(ROOT, "config", "application.properties");
const DOWNLOADS_DIR = path.join(ROOT, "downloads");
const OUTPUT_DIR = path.join(ROOT, "output");
const LOGS_DIR = path.join(ROOT, "logs");

const DEFAULT_TEMPLATES = {
  1: "Template/FMS Status tracker.xlsx",
  2: "Template/FMS 4 EP template.xlsx",
  3: "Template/E2E Report.xlsx",
  4: "Template/SIT Template.xlsx",
};

const DEFAULT_OUTPUT_FILES = {
  1: "output/FMS Status tracker.xlsx",
  2: "output/FMS 4 EP tracker.xlsx",
  3: "output/E2E Report.xlsx",
  4: "output/SIT Status tracker.xlsx",
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
  4: {
    statusSections: 3,
    sitDualProject: true,
    description:
      "SIT: 3 blocks (General Uganda, Life Uganda, General Tanzania). Uses Uganda + Tanzania projects.",
  },
};

function isSitTemplateChoice(choice) {
  return String(choice || "").trim() === "4";
}

const MODULE_ALIASES = {
  Payables: "Accounts Payable",
  "Account Payables": "Accounts Payable",
  "Accounts Receivables": "Accounts Receivable",
  "Cash & Bank": "Cash & Bank Management",
  "Cash and Bank": "Cash & Bank Management",
  "Cash and Bank Management": "Cash & Bank Management",
  Intergations: "Integrations",
  "Intergations (Footprint)": "Integrations",
  "Intergations (ILMS)": "Integrations",
  "Integrations (Footprint)": "Integrations",
  "Integrations (ILMS)": "Integrations",
  "Fixed Assets": "Fixed Assets Management",
  "Data Migration -Functional": "Data Migration TC's",
  "Data Migration -Technical": "Data Migration TC's",
  "Investment Receipting": "Investment Management",
};

const MODULE_CANONICAL_BY_KEY = {
  integrations: "Integrations",
  "cash and bank": "Cash & Bank Management",
  "cash and bank management": "Cash & Bank Management",
};

const PASS_RATE_BANDS = {
  red: { argb: "FFFF0000", hex: "#FF0000" },
  amber: { argb: "FFFFC000", hex: "#FFC000" },
  green: { argb: "FF92D050", hex: "#92D050" },
};

/** Normalize SimplifyQA Entity values -> template entity labels */
const ENTITY_ALIASES = {
  "life ke": "Life Kenya",
  "life kenya": "Life Kenya",
  "gen ke": "General Kenya",
  "gen kenya": "General Kenya",
  williamson: "General Kenya",
  "gen ke/williamson": "General Kenya",
  "general kenya": "General Kenya",
  "ug general": "Uganda General",
  "uganda general": "Uganda General",
  "ug life": "Uganda Life",
  "uganda life": "Uganda Life",
  "tz general": "General Tanzania",
  tanzania: "General Tanzania",
  "tanzania general": "General Tanzania",
  "general tz": "General Tanzania",
  tz: "General Tanzania",
  "gen ug": "General Uganda",
  "general uganda": "General Uganda",
  "life ug": "Life Uganda",
  "life uganda": "Life Uganda",
  "gen tz": "General Tanzania",
  "general tanzania": "General Tanzania",
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

function moduleMatchKey(name, { stripE2E = true } = {}) {
  let s = String(name || "").trim();
  if (stripE2E) s = s.replace(/^e2e\s+/i, "");
  return s
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/\([^)]*\)/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .replace(/\bintergations\b/g, "integrations")
    .trim();
}

function lookupModuleAlias(name) {
  if (!name) return name;
  if (MODULE_ALIASES[name]) return MODULE_ALIASES[name];
  const lower = String(name).toLowerCase();
  for (const [k, v] of Object.entries(MODULE_ALIASES)) {
    if (k.toLowerCase() === lower) return v;
  }
  return name;
}

function matchCatalogModule(name, catalog, { stripE2E = true } = {}) {
  if (!name || !catalog || !catalog.length) return null;
  const key = moduleMatchKey(name, { stripE2E });
  if (!key) return null;
  return catalog.find((t) => moduleMatchKey(t, { stripE2E }) === key) || null;
}

function moduleInSet(name, nameSet, { stripE2E = true } = {}) {
  if (!name || !nameSet) return false;
  return moduleInList(name, [...nameSet], { stripE2E });
}

function moduleInList(name, list, { stripE2E = true } = {}) {
  if (!name || !list || !list.length) return false;
  const key = moduleMatchKey(name, { stripE2E });
  return list.some((t) => moduleMatchKey(t, { stripE2E }) === key);
}

function canonicalizeModuleName(rawName, { stripE2E = true, catalog = [] } = {}) {
  let name = String(rawName || "").trim();
  if (!name) return null;
  if (stripE2E) name = name.replace(/^E2E\s+/i, "").trim();
  name = lookupModuleAlias(name);
  const catalogHit = matchCatalogModule(name, catalog, { stripE2E });
  if (catalogHit) return catalogHit;
  const key = moduleMatchKey(name, { stripE2E });
  if (MODULE_CANONICAL_BY_KEY[key]) return MODULE_CANONICAL_BY_KEY[key];
  return name;
}

function resolveModuleName(rawName, catalog) {
  return canonicalizeModuleName(rawName, { stripE2E: true, catalog: catalog || [] });
}

/**
 * Defect module names must stay distinct from EP E2E modules.
 * Do not strip "E2E " — only apply naming aliases (Payables → Accounts Payable, etc.).
 */
function resolveDefectModuleName(rawName, catalog) {
  return canonicalizeModuleName(rawName, { stripE2E: false, catalog: catalog || [] });
}

function passRateBand(rate) {
  const pct = Math.round((Number(rate) || 0) * 100);
  if (pct <= 50) return PASS_RATE_BANDS.red;
  if (pct <= 80) return PASS_RATE_BANDS.amber;
  return PASS_RATE_BANDS.green;
}

function applyPassRateFill(cell, rate) {
  if (!cell) return;
  if (rate == null || !Number.isFinite(Number(rate))) return;
  const band = passRateBand(rate);
  const fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: band.argb },
    bgColor: { argb: band.argb },
  };
  const font = Object.assign({}, cell.font || {});
  font.color = { argb: "FF000000" };
  // Clone the style so band colours are not shared across Pass Rate cells.
  let style = {};
  try {
    style = JSON.parse(JSON.stringify(cell.style || {}));
  } catch {
    style = Object.assign({}, cell.style || {});
  }
  style.fill = fill;
  style.font = Object.assign({}, style.font || {}, font);
  cell.style = style;
  cell.fill = fill;
  cell.font = font;
}

function entityLookupKey(raw) {
  return String(raw || "")
    .trim()
    .toLowerCase()
    .replace(/\s*\/\s*/g, "/")
    .replace(/\s+/g, " ");
}

function normalizeEntityLabel(raw) {
  const key = entityLookupKey(raw);
  if (!key) return null;
  return ENTITY_ALIASES[key] || String(raw).trim();
}

/** Split multi-entity values like "Life KE,Gen KE/Williamson" and normalize each. */
function parseDefectEntities(raw) {
  const text = String(raw || "").trim();
  if (!text) return [];
  const out = [];
  const seen = new Set();
  const add = (label) => {
    if (!label) return;
    const k = String(label).toLowerCase();
    if (seen.has(k)) return;
    seen.add(k);
    out.push(label);
  };

  const wholeKey = entityLookupKey(text);
  if (ENTITY_ALIASES[wholeKey]) {
    add(ENTITY_ALIASES[wholeKey]);
    return out;
  }

  // Comma separates entities. Do not split on "/" — "Gen KE/Williamson" is one entity.
  for (const chunk of text.split(",").map((p) => p.trim()).filter(Boolean)) {
    add(normalizeEntityLabel(chunk));
  }
  return out;
}

function resolveDefectRowEntities(entityRaw, defaultEntity) {
  const entities = parseDefectEntities(entityRaw);
  if (entities.length) return entities;
  const fallback = String(defaultEntity || "").trim();
  return fallback ? [fallback] : [];
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
  const t = String(title || "").toLowerCase().replace(/\s+/g, " ");
  const checks = [
    ["life uganda", "Life Uganda"],
    ["general uganda", "General Uganda"],
    ["uganda life", "Uganda Life"],
    ["uganda general", "Uganda General"],
    ["general tanzania", "General Tanzania"],
    ["life kenya", "Life Kenya"],
    ["general kenya", "General Kenya"],
    ["tanzania", "General Tanzania"],
  ];
  for (const [needle, label] of checks) {
    if (t.includes(needle)) return label;
  }
  return null;
}

function entityLabelVariants(label) {
  const raw = String(label || "").trim();
  if (!raw) return [];
  const t = raw.toLowerCase();
  const out = new Set([raw, t]);
  if (t.includes("uganda") && t.includes("life")) {
    ["Life Uganda", "Uganda Life", "Life UG"].forEach((x) => out.add(x));
  } else if (t.includes("uganda")) {
    ["General Uganda", "Uganda General", "Gen UG"].forEach((x) => out.add(x));
  } else if (t.includes("kenya") && t.includes("life")) {
    ["Life Kenya", "Life KE"].forEach((x) => out.add(x));
  } else if (t.includes("kenya")) {
    ["General Kenya", "Gen KE", "Gen Kenya", "Gen KE/Williamson"].forEach((x) =>
      out.add(x)
    );
  } else if (t.includes("tanzania") || /\btz\b/.test(t)) {
    ["General Tanzania", "Tanzania", "Gen TZ", "TZ", "TZ General", "Tanzania General"].forEach(
      (x) => out.add(x)
    );
  }
  return [...out];
}

function entityCountsForSection(byEntity, sectionEntity) {
  const merged = new Map();
  if (!byEntity || !sectionEntity) return merged;
  const keys = new Set(entityLabelVariants(sectionEntity).map((s) => String(s).toLowerCase()));
  for (const [entKey, modMap] of byEntity.entries()) {
    if (!keys.has(String(entKey).toLowerCase())) continue;
    for (const [mod, counts] of modMap.entries()) {
      const cur = merged.get(mod) || emptyDefectCounts();
      cur.closed += counts.closed || 0;
      cur.deferred += counts.deferred || 0;
      cur.fixed += counts.fixed || 0;
      cur.pending += counts.pending || 0;
      merged.set(mod, cur);
    }
  }
  return merged;
}

function countsMatchingModule(modMap, templateName, { stripE2E = true } = {}) {
  const out = emptyDefectCounts();
  if (!modMap || !templateName) return out;
  const key = moduleMatchKey(templateName, { stripE2E });
  for (const [mod, counts] of modMap.entries()) {
    if (moduleMatchKey(mod, { stripE2E }) !== key) continue;
    out.closed += counts.closed || 0;
    out.deferred += counts.deferred || 0;
    out.fixed += counts.fixed || 0;
    out.pending += counts.pending || 0;
  }
  return out;
}

function extrasForDefectSection(extrasByEntity, sectionEntity) {
  if (!extrasByEntity || !sectionEntity) return [];
  const keys = new Set(
    entityLabelVariants(sectionEntity).map((s) => String(s).toLowerCase())
  );
  const out = [];
  const seen = new Set();
  for (const [entKey, extras] of extrasByEntity.entries()) {
    if (!keys.has(String(entKey).toLowerCase())) continue;
    for (const extra of extras || []) {
      if (!extra || !extra.name || seen.has(extra.name)) continue;
      seen.add(extra.name);
      out.push(extra);
    }
  }
  return out;
}

/** Modules with defects for this entity that are not rows in this Defect Summary block. */
function extraDefectModulesForSection(sectionModules, entityCounts) {
  const extras = [];
  if (!entityCounts || typeof entityCounts.entries !== "function") return extras;
  for (const [name, counts] of entityCounts.entries()) {
    const total =
      (counts.closed || 0) +
      (counts.deferred || 0) +
      (counts.fixed || 0) +
      (counts.pending || 0);
    if (total <= 0) continue;
    if (moduleInList(name, sectionModules || [], { stripE2E: false })) continue;
    extras.push({
      name,
      closed: counts.closed || 0,
      deferred: counts.deferred || 0,
      fixed: counts.fixed || 0,
      pending: counts.pending || 0,
      total,
    });
  }
  return extras;
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

function countStatusSectionsFromFile(filePath) {
  const inspected = workstreamFill.inspectTemplateFile(filePath);
  return inspected && inspected.statusSections ? inspected.statusSections : null;
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
  let relative = fromProps || fallback;
  if (!relative) {
    throw new ReporterError(
      "CONFIG_INVALID",
      `Unknown template choice "${choice}". Please set TEMPLATE_${choice}=... or use a valid TEMPLATE_CHOICE in config/application.properties.`
    );
  }

  const outputKey = `OUTPUT_FILE_${choice}`;
  const outputRelative = String(
    props[outputKey] || DEFAULT_OUTPUT_FILES[choice] || ""
  ).trim();
  if (!outputRelative) {
    throw new ReporterError(
      "CONFIG_INVALID",
      `Output file is missing for template choice ${choice}. Please set ${outputKey}=... in config/application.properties.`
    );
  }

  let templatePath = resolvePathMaybe(relative);
  if (!fs.existsSync(templatePath)) {
    const altRel = String(relative).replace(/\s\(\d+\)(\.xlsx)$/i, "$1");
    const altPath = resolvePathMaybe(altRel);
    if (altRel !== relative && fs.existsSync(altPath)) {
      relative = altRel;
      templatePath = altPath;
    }
  }

  return {
    choice,
    templatePath,
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

  const sit = isSitTemplateChoice(choice);
  const projectIdB = Number(props.PROJECT_ID_B);
  if (sit && !projectIdB) {
    throw new ReporterError(
      "CONFIG_INVALID",
      "SIT template needs a Tanzania project. Select Project B (Tanzania, id 6) as well as Uganda."
    );
  }

  const plans = planEntries.map((p) => ({
    index: p.index,
    exePlanId: Number(p.value),
    sectionOverride: String(props[`SECTION_${p.index}`] || "").trim() || null,
    projectId: sit ? (p.index <= 2 ? projectId : projectIdB) : projectId,
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
    projectIdB: sit ? projectIdB : null,
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

function qaJsonHeaders(token, referer) {
  return {
    accept: "application/json, text/plain, */*",
    authorization: `Bearer ${token}`,
    origin: ORIGIN,
    referer: referer || `${ORIGIN}/user/defect`,
    "content-type": "application/json",
    "user-agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36",
  };
}

function preferenceUserId(token) {
  const payload = loadEnv.readJwtPayload(token);
  const fromJwt = payload && (payload.userId || payload.uid);
  if (fromJwt) return Number(fromJwt);
  const fromEnv = Number(process.env.SIMPLIFYQA_USER_ID || "");
  if (fromEnv) return fromEnv;
  return DEFECT_PREFERENCE_USER_ID_DEFAULT;
}

function isEntityPreferenceColumn(col) {
  if (!col || typeof col !== "object") return false;
  const label = `${col.label || ""} ${col.originalLabel || ""} ${col.key || ""}`.toLowerCase();
  return /\bentit(?:y|ies)\b/.test(label);
}

function findEntityPreferenceIndex(columns) {
  if (!Array.isArray(columns)) return -1;
  return columns.findIndex(isEntityPreferenceColumn);
}

function defectEntityColumnTemplate({ id, projectId, key, label, seq }) {
  return {
    deleted: false,
    id,
    projectId,
    cellCssClass: "",
    controlType: "LIST",
    customFldType: "LIST",
    dateTimeFormat: "",
    displayKey: { field: "value" },
    enableColumnDrag: true,
    enableColumnSearch: true,
    enableInlineEdit: true,
    fieldType: "customFld",
    filter: true,
    hasMap: true,
    jsonArrayKey: "",
    key,
    keyname: "key",
    label,
    originalLabel: label,
    parentComponentId: 0,
    parentFieldId: "",
    parentId: 9,
    mandatory: true,
    mappedModule: "",
    maxLimit: 0,
    minLimit: 0,
    multi: true,
    readOnly: false,
    seq: seq || 24,
    showInViewMode: false,
    showProfileIcon: false,
    shownIn: ["TABLE", "EDIT", "CREATE"],
    sort: true,
    sortType: "dsc",
    type: "ARRAYKEYS",
    width: 235,
    isChecked: true,
    searchRow: "",
  };
}

const DEFECT_ENTITY_COLUMNS_BY_PROJECT = {
  2: defectEntityColumnTemplate({
    id: "27f2155c-d4a7-4ff9-96d2-8af880d5fec7",
    projectId: 2,
    key: "CUSTFIELD_96Entity_6_dhd40",
    label: "Entity",
    seq: 23,
  }),
  5: defectEntityColumnTemplate({
    id: "3f7d69be-1055-40ef-8689-7bae1df3cc66",
    projectId: 5,
    key: "CUSTFIELD_96ENTITY_6_7780l",
    label: "ENTITY",
    seq: 24,
  }),
  6: defectEntityColumnTemplate({
    id: "97c877bb-0274-4827-88e0-784c7cfc9f39",
    projectId: 6,
    key: "CUSTFIELD_96ENTITY_6_2hk5a",
    label: "ENTITY",
    seq: 24,
  }),
};

function defectEntityColumnForProject(projectId) {
  const pid = Number(projectId);
  const source =
    DEFECT_ENTITY_COLUMNS_BY_PROJECT[pid] || DEFECT_ENTITY_COLUMNS_BY_PROJECT[2];
  const copy = JSON.parse(JSON.stringify(source));
  if (Number.isFinite(pid) && pid > 0) copy.projectId = pid;
  return copy;
}

function normalizeEntityPreferenceColumn(col, projectId) {
  const copy = JSON.parse(JSON.stringify(col));
  copy.isChecked = true;
  copy.deleted = false;
  if (Number.isFinite(Number(projectId)) && Number(projectId) > 0) {
    copy.projectId = Number(projectId);
  }
  const shown = Array.isArray(copy.shownIn) ? copy.shownIn.slice() : [];
  if (!shown.includes("TABLE")) shown.push("TABLE");
  copy.shownIn = shown;
  return copy;
}

function defectFilterPreferenceBody(defectFilter) {
  return { defectFilter: Array.isArray(defectFilter) ? defectFilter : [] };
}

function mergeEntityIntoDefectFilter(defectFilter, projectId) {
  const filter = Array.isArray(defectFilter) ? defectFilter.slice() : [];
  const wanted = normalizeEntityPreferenceColumn(
    defectEntityColumnForProject(projectId),
    projectId
  );
  const idx = findEntityPreferenceIndex(filter);
  let changed = false;
  if (idx < 0) {
    wanted.seq = filter.length + 1;
    filter.push(wanted);
    changed = true;
  } else {
    const existing = filter[idx] || {};
    const wrongProject =
      Number(existing.projectId) > 0 &&
      Number(projectId) > 0 &&
      Number(existing.projectId) !== Number(projectId);
    const wrongKey = Boolean(wanted.key && existing.key && existing.key !== wanted.key);
    if (wrongProject || wrongKey) {
      wanted.seq = existing.seq || filter.length;
      filter[idx] = wanted;
      changed = true;
    } else if (
      existing.isChecked !== true ||
      existing.deleted === true ||
      !(Array.isArray(existing.shownIn) && existing.shownIn.includes("TABLE"))
    ) {
      filter[idx] = normalizeEntityPreferenceColumn(existing, projectId);
      changed = true;
    }
  }
  return { defectFilter: filter, changed };
}

function defectFilterMajorityProjectId(filter) {
  const counts = new Map();
  for (const col of filter || []) {
    const pid = Number(col && col.projectId);
    if (!Number.isFinite(pid) || pid <= 0) continue;
    counts.set(pid, (counts.get(pid) || 0) + 1);
  }
  let best = null;
  let n = 0;
  for (const [pid, count] of counts.entries()) {
    if (count > n) {
      best = pid;
      n = count;
    }
  }
  return best;
}

function preferenceHasProjectEntity(filter, projectId) {
  const wanted = defectEntityColumnForProject(projectId);
  const idx = findEntityPreferenceIndex(filter);
  if (idx < 0) return false;
  const col = filter[idx] || {};
  const shown = Array.isArray(col.shownIn) ? col.shownIn : [];
  return (
    col.key === wanted.key &&
    col.isChecked === true &&
    col.deleted !== true &&
    shown.includes("TABLE") &&
    (!Number(col.projectId) || Number(col.projectId) === Number(projectId))
  );
}

function capturedDefectFilterForProject(projectId) {
  const pid = Number(projectId);
  const source = DEFECT_FILTER_BY_PROJECT[pid];
  return source ? JSON.parse(JSON.stringify(source)) : null;
}

/**
 * Decide the defectFilter to PUT for this project.
 * If GET still has another project's columns, use that project's captured list
 * (Entity field keys are per-project; swapping only Entity onto the wrong list 400s).
 */
function resolveDefectFilterForProject(currentFilter, projectId) {
  const pid = Number(projectId);
  const current = Array.isArray(currentFilter) ? currentFilter : [];
  const majority = defectFilterMajorityProjectId(current);
  const captured = capturedDefectFilterForProject(pid);

  if (
    preferenceHasProjectEntity(current, pid) &&
    (majority == null || majority === pid)
  ) {
    return { defectFilter: current, changed: false, reason: "already-enabled" };
  }

  if (majority != null && majority !== pid && captured) {
    return {
      defectFilter: captured,
      changed: true,
      reason: "wrong-project-columns",
    };
  }

  const merged = mergeEntityIntoDefectFilter(current, pid);
  return {
    defectFilter: merged.defectFilter,
    changed: merged.changed,
    reason: merged.changed ? "merge-entity" : "already-enabled",
  };
}

function defectExportHasEntityColumn(columns) {
  return (columns || []).some((c) => /\bentit/i.test(String(c || "")));
}

function unwrapPreference(json) {
  if (!json || typeof json !== "object") return {};
  if (Array.isArray(json.defectFilter)) return json;
  if (json.data && typeof json.data === "object") return json.data;
  return json;
}

async function fetchUserPreference(token, userId) {
  const res = await fetch(`${ORIGIN}/pm/preference/${userId}`, {
    method: "GET",
    headers: qaJsonHeaders(token, `${ORIGIN}/user/defect`),
  });
  const text = await res.text();
  if (!res.ok) {
    throw new Error(`preference GET ${res.status}`);
  }
  return unwrapPreference(text ? JSON.parse(text) : {});
}

async function saveDefectFilterPreference(token, userId, defectFilter) {
  const body = defectFilterPreferenceBody(defectFilter);
  const res = await fetch(`${ORIGIN}/pm/preference/${userId}`, {
    method: "PUT",
    headers: qaJsonHeaders(token, `${ORIGIN}/user/defect`),
    body: JSON.stringify(body),
  });
  const text = await res.text();
  if (!res.ok) {
    const snippet = String(text || "")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 180);
    throw new Error(`preference PUT ${res.status}${snippet ? `: ${snippet}` : ""}`);
  }
  if (!text) return body;
  try {
    return JSON.parse(text);
  } catch {
    return body;
  }
}

/**
 * Enable this project's Entity column in user preference before export.
 * PUT body is only { defectFilter }. Do not skip just because the Excel
 * header is named ENTITY — that header can be another project's field.
 */
async function ensureDefectEntityColumnPreference(token, projectId, log) {
  const logger = typeof log === "function" ? log : () => {};
  const userId = preferenceUserId(token);
  const pid = Number(projectId);
  let pref;
  try {
    pref = await fetchUserPreference(token, userId);
  } catch (err) {
    logger(
      `Could not read defect column preference (/pm/preference/${userId}): ${err.message}`
    );
    return false;
  }

  const targetPid = pid || Number(pref.activeProjectId);
  const resolved = resolveDefectFilterForProject(pref.defectFilter, targetPid);
  const wanted = defectEntityColumnForProject(targetPid);
  if (!resolved.changed) {
    logger(
      `Defect Entity column preference already uses ${wanted.key} for project ${targetPid}.`
    );
    return true;
  }

  const entity = resolved.defectFilter[findEntityPreferenceIndex(resolved.defectFilter)];
  if (resolved.reason === "wrong-project-columns") {
    logger(
      `Defect column preference is still for another project; switching it to project ${targetPid} Entity ${entity && entity.key ? entity.key : wanted.key}.`
    );
  } else {
    logger(
      `Enabling defect Entity column ${entity && entity.key ? entity.key : wanted.key} for project ${targetPid}.`
    );
  }

  try {
    await saveDefectFilterPreference(token, userId, resolved.defectFilter);
    const verify = await fetchUserPreference(token, userId);
    if (preferenceHasProjectEntity(verify && verify.defectFilter, targetPid)) {
      logger("Defect Entity column preference saved.");
      return true;
    }
    logger(
      "ALERT: Could not confirm this project's Entity column preference after enabling it. Defect export may still omit Entity values."
    );
    return false;
  } catch (err) {
    logger(`Could not save defect Entity column preference: ${err.message}`);
    return false;
  }
}

/**
 * Sprint 3 Kenya: enable Sprint then Entity on this project's defectFilter.
 * PUT { defectFilter } only. Does not run for other templates.
 */
async function ensureSprint3DefectColumnPreference(token, projectId, log) {
  const logger = typeof log === "function" ? log : () => {};
  const userId = preferenceUserId(token);
  const pid = Number(projectId) || 2;
  let pref;
  try {
    pref = await fetchUserPreference(token, userId);
  } catch (err) {
    logger(
      `Could not read defect column preference (/pm/preference/${userId}): ${err.message}`
    );
    return false;
  }

  let filter = Array.isArray(pref.defectFilter) ? pref.defectFilter : [];
  const majority = defectFilterMajorityProjectId(filter);
  let changed = false;
  if (majority != null && majority !== pid) {
    filter = sprint3Defects.capturedKenyaDefectFilter();
    changed = true;
    logger(
      `Defect column preference is still for another project; switching it to Kenya Sprint + Entity columns for project ${pid}.`
    );
  }

  const sprintMerged = sprint3Defects.mergeSprintIntoDefectFilter(filter, pid);
  filter = sprintMerged.defectFilter;
  if (sprintMerged.changed) {
    changed = true;
    logger(
      `Enabling defect Sprint column (iterationId) for project ${pid}.`
    );
  }

  const entityMerged = mergeEntityIntoDefectFilter(filter, pid);
  filter = entityMerged.defectFilter;
  if (entityMerged.changed) {
    changed = true;
    logger(
      `Enabling defect Entity column ${defectEntityColumnForProject(pid).key} for project ${pid}.`
    );
  }

  if (!changed) {
    logger(
      `Sprint 3 defect columns already enabled (Sprint + Entity) for project ${pid}.`
    );
    return true;
  }

  try {
    await saveDefectFilterPreference(token, userId, filter);
    const verify = await fetchUserPreference(token, userId);
    const okSprint = sprint3Defects.preferenceHasSprintColumn(
      verify && verify.defectFilter,
      pid
    );
    const okEntity = preferenceHasProjectEntity(verify && verify.defectFilter, pid);
    if (okSprint && okEntity) {
      logger("Sprint 3 defect Sprint and Entity column preference saved.");
      return true;
    }
    logger(
      "ALERT: Could not confirm Sprint and Entity columns after enabling them. Defect export may omit Sprint or Entity."
    );
    return false;
  } catch (err) {
    logger(`Could not save Sprint 3 defect column preference: ${err.message}`);
    return false;
  }
}

async function downloadAndAggregateDefects(token, pid, timezone, offset, log, extraAggOptions) {
  const logger = typeof log === "function" ? log : () => {};
  const sprint3 = Boolean(extraAggOptions && extraAggOptions.sprint3);
  if (sprint3) {
    logger(`\nPreparing Sprint 3 defect Sprint + Entity columns for projectId=${pid}...`);
    await ensureSprint3DefectColumnPreference(token, pid, logger);
  } else {
    logger(`\nPreparing defect Entity column preference for projectId=${pid}...`);
    await ensureDefectEntityColumnPreference(token, pid, logger);
  }
  logger(`Downloading defect export for projectId=${pid}...`);
  const defectPath = await downloadDefectExport(token, pid, timezone, offset);
  logger(`Defect export: ${defectPath}`);
  return aggregateDefectExport(defectPath, extraAggOptions || {});
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
  const want = names.map((n) => String(n).trim().toLowerCase());
  let fallback = "";
  for (const key of Object.keys(row || {})) {
    const norm = key.trim().toLowerCase();
    if (!want.includes(norm)) continue;
    const val = row[key];
    if (String(val == null ? "" : val).trim()) return val;
    if (fallback === "") fallback = val;
  }
  return fallback;
}

/**
 * Load defect export and aggregate by normalized entity + module.
 * Multi-entity values count toward every listed entity.
 * Rows with a blank ENTITY use options.defaultEntity when provided (SIT TZ).
 */
function aggregateDefectExport(filePath, options = {}) {
  const defaultEntity = String(options.defaultEntity || "").trim() || null;
  const sprintFilter = String(options.sprintFilter || "").trim() || null;
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
  const columns = rows[0] ? Object.keys(rows[0]) : [];

  /** @type {Map<string, Map<string, {closed:number,deferred:number,fixed:number,pending:number}>>} */
  const byEntity = new Map();
  const unknownStates = new Map();
  let mappedRows = 0;
  let skippedUnknownState = 0;
  let skippedNoEntity = 0;
  let usedDefaultEntity = 0;
  let skippedSprint = 0;

  for (const row of rows) {
    const moduleRaw = pickDefectField(row, ["module"]);
    const entityRaw = pickDefectField(row, ["entity", "entities"]);
    const stateRaw = pickDefectField(row, ["state", "status"]);
    const sprintRaw = pickDefectField(row, ["sprint"]);
    const moduleName = resolveDefectModuleName(moduleRaw);
    const parsed = parseDefectEntities(entityRaw);
    const usedFallback = parsed.length === 0 && Boolean(defaultEntity);
    const entities = resolveDefectRowEntities(entityRaw, defaultEntity);
    const bucket = mapDefectState(stateRaw);

    if (!moduleName) continue;
    if (sprintFilter && !sprint3Defects.sprintValueMatches(sprintRaw, sprintFilter)) {
      skippedSprint += 1;
      continue;
    }
    if (!entities.length) {
      skippedNoEntity += 1;
      continue;
    }

    if (!bucket) {
      skippedUnknownState += 1;
      const label = String(stateRaw || "").trim() || "(blank)";
      unknownStates.set(label, (unknownStates.get(label) || 0) + 1);
      continue;
    }

    mappedRows += 1;
    if (usedFallback) usedDefaultEntity += 1;
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
    skippedNoEntity,
    usedDefaultEntity,
    skippedSprint,
    defaultEntity,
    unknownStates: [...unknownStates.entries()].map(([state, count]) => ({
      state,
      count,
    })),
    exportRows: rows.length,
    columns,
  };
}

function mergeDefectAggregates(aggs) {
  const byEntity = new Map();
  const unknownStates = new Map();
  let mappedRows = 0;
  let skippedUnknownState = 0;
  let skippedNoEntity = 0;
  let usedDefaultEntity = 0;
  let skippedSprint = 0;
  let exportRows = 0;
  const columns = [];
  const seenCol = new Set();
  for (const agg of aggs || []) {
    mappedRows += agg.mappedRows || 0;
    skippedUnknownState += agg.skippedUnknownState || 0;
    skippedNoEntity += agg.skippedNoEntity || 0;
    usedDefaultEntity += agg.usedDefaultEntity || 0;
    skippedSprint += agg.skippedSprint || 0;
    exportRows += agg.exportRows || 0;
    for (const col of agg.columns || []) {
      if (seenCol.has(col)) continue;
      seenCol.add(col);
      columns.push(col);
    }
    for (const u of agg.unknownStates || []) {
      unknownStates.set(u.state, (unknownStates.get(u.state) || 0) + (u.count || 0));
    }
    for (const [ent, modMap] of (agg.byEntity || new Map()).entries()) {
      if (!byEntity.has(ent)) byEntity.set(ent, new Map());
      const dest = byEntity.get(ent);
      for (const [mod, counts] of modMap.entries()) {
        if (!dest.has(mod)) dest.set(mod, emptyDefectCounts());
        const d = dest.get(mod);
        d.closed += counts.closed || 0;
        d.deferred += counts.deferred || 0;
        d.fixed += counts.fixed || 0;
        d.pending += counts.pending || 0;
      }
    }
  }
  return {
    byEntity,
    mappedRows,
    skippedUnknownState,
    skippedNoEntity,
    usedDefaultEntity,
    skippedSprint,
    unknownStates: [...unknownStates.entries()].map(([state, count]) => ({
      state,
      count,
    })),
    exportRows,
    columns,
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

function buildDefectSectionSummary(sectionModules, entityCounts, extraRows = []) {
  const rows = [];

  // Strict template order; only modules with at least one defect.
  for (const name of sectionModules) {
    const counts = countsMatchingModule(entityCounts, name, { stripE2E: false });
    const total =
      counts.closed + counts.deferred + counts.fixed + counts.pending;
    if (total <= 0) continue;
    rows.push({ name, ...counts, total });
  }

  const seen = new Set(rows.map((r) => moduleMatchKey(r.name, { stripE2E: false })));
  for (const extra of extraRows || []) {
    if (!extra || !extra.name) continue;
    const key = moduleMatchKey(extra.name, { stripE2E: false });
    if (seen.has(key)) continue;
    rows.push(extra);
    seen.add(key);
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

/** Template rate rows merge label D:F and value G:I (Total), never into Pass Rate. */
function restoreStatusRateRowMerges(sheet, rowNum, cols) {
  const left = cols.passed;
  const right = cols.passRate || cols.total;
  if (left && right && right > left) {
    try {
      sheet.unMergeCells(rowNum, left, rowNum, right);
    } catch {
      // ignore leftover G:J merges from older fills
    }
  }
  ensureRowMerge(sheet, rowNum, cols.passed, cols.blocked);
  ensureRowMerge(sheet, rowNum, cols.inProgress, cols.total);
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

function aggregateExport(filePath, templateModules, options = {}) {
  const includeExtraModules = Boolean(options.includeExtraModules);
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
    const name = resolveModuleName(row[6], templateModules);
    if (!name) continue;
    if (!stats[name]) stats[name] = emptyCounts();
    stats[name][normalizeStatus(row[8])] += 1;
  }

  const ordered = [];

  // Template order first; extras are appended only when the user opts in.
  for (const name of templateModules) {
    ordered.push({
      name,
      ...summarizeModule(stats[name] || emptyCounts()),
    });
  }

  const extraModules = Object.keys(stats)
    .filter((name) => !moduleInList(name, templateModules, { stripE2E: true }))
    .map((name) => ({
      name,
      ...summarizeModule(stats[name]),
    }));

  const sheetRows = includeExtraModules
    ? ordered.concat(extraModules)
    : ordered;

  const totals = emptyCounts();
  for (const r of sheetRows) {
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
    extraModulesIncluded: includeExtraModules && extraModules.length > 0,
    rows: sheetRows,
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
    if (cols.blocked) setNumber(row.getCell(cols.blocked), data.blocked);
    if (cols.inProgress) setNumber(row.getCell(cols.inProgress), data.inProgress);
    if (cols.notExecuted) setNumber(row.getCell(cols.notExecuted), data.notExecuted);
    setFormula(
      row.getCell(cols.total),
      `SUM(${cPass}${r}:${cNE}${r})`,
      data.total
    );
    if (cols.passRate) {
      const passCell = row.getCell(cols.passRate);
      setFormula(
        passCell,
        `${cPass}${r}/(${cTot}${r}-${cBlock}${r})`,
        data.passRate,
        "0%"
      );
      applyPassRateFill(passCell, data.passRate);
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
    const totalPassCell = totalRow.getCell(cols.passRate);
    setFormula(
      totalPassCell,
      `${cPass}${totalRowNum}/(${cTot}${totalRowNum}-${cBlock}${totalRowNum})`,
      summary.totalSummary.passRate,
      "0%"
    );
    applyPassRateFill(totalPassCell, summary.totalSummary.passRate);
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
      restoreStatusRateRowMerges(sheet, r, cols);
    }
    if (rowHasLabel(sheet, r, "overall pass rate")) {
      setFormula(
        sheet.getRow(r).getCell(cols.inProgress),
        `${cPass}${totalRowNum}/(${cTot}${totalRowNum}-${cBlock}${totalRowNum})`,
        summary.overallPassRate,
        "0%"
      );
      sheet.getRow(r).commit();
      restoreStatusRateRowMerges(sheet, r, cols);
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

function isFileLockError(err) {
  const code = String((err && err.code) || "").toUpperCase();
  const msg = String((err && (err.message || err)) || "").toLowerCase();
  return (
    code === "EBUSY" ||
    code === "EPERM" ||
    code === "EACCES" ||
    /ebusy|eperm|eacces|locked|being used by another|resource busy|access is denied/.test(
      msg
    )
  );
}

function throwIfOutputLocked(err, destPath) {
  if (!isFileLockError(err)) return;
  throw new ReporterError(
    "OUTPUT_LOCKED",
    `Unable to save the report because "${path.basename(destPath)}" is open in Excel. Close that file and generate again.`
  );
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
    try {
      await filledWorkbook.xlsx.writeFile(masterPath);
    } catch (err) {
      throwIfOutputLocked(err, masterPath);
      throw err;
    }
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
  try {
    await master.xlsx.readFile(masterPath);
  } catch (err) {
    throwIfOutputLocked(err, masterPath);
    throw err;
  }
  const sheetName = uniqueSheetName(master.worksheets.map((ws) => ws.name));
  const target = master.addWorksheet(sheetName);
  copyWorksheet(cleanSheet, target);
  activateWorksheet(master, target);
  try {
    await master.xlsx.writeFile(masterPath);
  } catch (err) {
    throwIfOutputLocked(err, masterPath);
    throw err;
  }
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

  const extraModulesFound = [];
  const sitRun = isSitTemplateChoice(runtime.templateChoice);
  const sprint3Run = sprint3Defects.isSprint3TemplatePath(runtime.templateRelative);

  progress("prepare", "Preparing template…");
  cleanupOldDownloads(log, 8);
  log(`Template choice: ${runtime.templateChoice} -> ${runtime.templateRelative}`);
  if (sprint3Run) {
    log("Template kind: Sprint 3. Defects are filtered by the Sprint value you entered.");
  }
  if (args.addMissingModules) {
    log("Missing template modules will be added to this generated report (template file is not changed).");
  }
  log(`Output file: ${runtime.outputRelative}`);
  log(`Project ID: ${runtime.projectId}${runtime.projectIdB ? ` + ${runtime.projectIdB}` : ""}`);
  log(`Plans configured: ${runtime.plans.length}`);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(runtime.templatePath);
  let sheet = workbook.worksheets[0];
  neutralizeSharedFormulas(sheet);
  let moduleSections = discoverStatusSections(sheet);
  if (!moduleSections.length) {
    for (const ws of workbook.worksheets) {
      const found = discoverStatusSections(ws);
      if (found.length) {
        sheet = ws;
        moduleSections = found;
        neutralizeSharedFormulas(sheet);
        break;
      }
    }
  }
  const workstreamHit = moduleSections.length
    ? { sheet, sections: [] }
    : workstreamFill.discoverWorkstreamStatusInWorkbook(
        workbook,
        runtime.templatePath
      );
  if (!moduleSections.length && workstreamHit.sheet && workstreamHit.sheet !== sheet) {
    sheet = workstreamHit.sheet;
    neutralizeSharedFormulas(sheet);
  }
  const workstreamSections = workstreamHit.sections || [];
  const templateKind = moduleSections.length
    ? "module-wise"
    : workstreamSections.length
      ? "workstream"
      : null;
  const discovered = moduleSections.length ? moduleSections : workstreamSections;

  if (!discovered.length) {
    throw new ReporterError(
      "TEMPLATE_INVALID",
      `No fillable status section was found in "${runtime.templateRelative}". Module-wise templates need a Module-wise Daily Status table. Workstream templates (like As and When Commission) need a Daily Status Report with Passed/Failed columns and no MODULE column.`
    );
  }

  if (templateKind === "workstream") {
    log(
      "Template kind: workstream. All testcases in the selected plan are counted into one status row. Module-wise mapper is not used."
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
      extraModules: [],
      extraModulesAdded: false,
    };
  }

  const needsApi = mapped.some((p) => !args.localByPlan[String(p.index)]);
  const token = needsApi ? getBearerToken() : null;

  const groupedByProject = new Map();
  for (const plan of mapped) {
    const pid = Number(plan.projectId || runtime.projectId);
    if (!groupedByProject.has(pid)) groupedByProject.set(pid, []);
    groupedByProject.get(pid).push(plan);
  }

  const exports = [];
  for (const [pid, group] of groupedByProject.entries()) {
    const groupNeedsApi = group.some((p) => !args.localByPlan[String(p.index)]);
    if (token && groupNeedsApi && !args.skipSwitch) {
      log(`\nSwitching project to projectId=${pid}...`);
      await switchProject(token, pid);
      log("Project switch OK");
    } else if (!groupNeedsApi) {
      log(`\nLocal mode: skipping auth/project switch for projectId=${pid}`);
    }

    log(`\nFetching ${group.length} execution plan export(s) for projectId=${pid}...`);
    progress("download_ep", "Downloading execution plan exports…");
    const fileJobs = group.map(async (plan) => {
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
          pid,
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
    const groupExports = await Promise.all(fileJobs);
    exports.push(...groupExports);
  }
  for (const item of exports) {
    log(`  Plan ${item.plan.index} (EP ${item.plan.exePlanId}): ${item.filePath}`);
  }

  progress("fill_status", "Filling status sections…");
  for (const item of exports) {
    const { plan, filePath } = item;
    try {
      if (templateKind === "workstream") {
        const summary = workstreamFill.aggregateWorkstreamExport(filePath);
        workstreamFill.writeWorkstreamSection(sheet, plan.sectionTitle, summary);
        const ok = summary.exportRows === summary.totalSummary.total;
        log(
          `\nPlan ${plan.index} EP-${plan.exePlanId}: workstream export=${summary.exportRows} filledTotal=${summary.totalSummary.total} tally=${ok ? "OK" : "MISMATCH"}`
        );
        log(
          `  All testcases: P=${summary.totalSummary.passed} F=${summary.totalSummary.failed} B=${summary.totalSummary.blocked} IP=${summary.totalSummary.inProgress} NE=${summary.totalSummary.notExecuted} T=${summary.totalSummary.total}`
        );
        if (!ok) {
          throw new ReporterError(
            "TALLY_MISMATCH",
            `Unable to complete the report for execution plan ${plan.exePlanId}. Extracted counts do not match the export total (${summary.exportRows} vs filled ${summary.totalSummary.total}).`
          );
        }
        continue;
      }
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
      const summary = aggregateExport(filePath, templateModules, {
        includeExtraModules: Boolean(args.addMissingModules),
      });
      writeSection(sheet, plan.sectionTitle, summary);

      const accounted =
        summary.totalSummary.total +
        (summary.extraModulesIncluded ? 0 : summary.extraTotal);
      const ok = accounted === summary.exportRows;
      log(
        `\nPlan ${plan.index} EP-${plan.exePlanId}: export=${summary.exportRows} templateTotal=${summary.totalSummary.total} extra=${summary.extraTotal} tally=${ok ? "OK" : "MISMATCH"}`
      );
      if (summary.extraModules.length) {
        extraModulesFound.push({
          source: "status",
          planIndex: plan.index,
          exePlanId: plan.exePlanId,
          sectionTitle: plan.sectionTitle,
          modules: summary.extraModules,
        });
        if (summary.extraModulesIncluded) {
          log(
            `\nIncluding extra EP modules in this report (output only; template file unchanged) for EP-${plan.exePlanId}:`
          );
          for (const r of summary.extraModules) {
            log(
              `  - ${r.name}: P=${r.passed} F=${r.failed} B=${r.blocked} T=${r.total}`
            );
          }
        } else {
          log(
            `\nALERT: Extra EP modules not in template (not added to sheet) for EP-${plan.exePlanId}:`
          );
          for (const r of summary.extraModules) {
            log(
              `  - ${r.name}: P=${r.passed} F=${r.failed} B=${r.blocked} T=${r.total}`
            );
          }
          log(
            "Add these modules from the run notes if they should appear in this report, or add them to the template for future runs."
          );
        }
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
  if (includeDefects && templateKind === "workstream") {
    progress("defects", "Downloading and filling defects…");
    const defectSections = workstreamFill.discoverWorkstreamDefectSections(sheet);
    if (!defectSections.length) {
      log("\nDefects: no workstream Defects section found in template (skipped).");
    } else if (args.dryRun) {
      log("Dry-run: defect export/fill skipped.");
    } else {
      log(`\nWorkstream defect sections in template: ${defectSections.length}`);
      defectSections.forEach((s, i) => log(`  [${i + 1}] ${s.title}`));

      const defectPaths = [];
      if (args.localDefects) {
        const defectPath = resolvePathMaybe(args.localDefects);
        if (!fs.existsSync(defectPath)) {
          throw new ReporterError(
            "EXTRACT_FAILED",
            `Unable to extract defect data. Local file was not found: ${defectPath}`
          );
        }
        log(`Using local defect export: ${defectPath}`);
        defectPaths.push(defectPath);
      } else {
        const defectToken = token || getBearerToken();
        const defectProjectIds = [
          ...new Set(
            mapped
              .map((p) => Number(p.projectId || runtime.projectId))
              .filter(Boolean)
          ),
        ];
        for (const pid of defectProjectIds) {
          if (defectToken && !args.skipSwitch) {
            log(`\nSwitching project to projectId=${pid} for defects...`);
            await switchProject(defectToken, pid);
            log("Project switch OK");
          }
          log(`\nPreparing defect Entity column preference for projectId=${pid}...`);
          await ensureDefectEntityColumnPreference(defectToken, pid, log);
          log(`Downloading defect export for projectId=${pid}...`);
          const defectPath = await downloadDefectExport(
            defectToken,
            pid,
            runtime.timezone,
            runtime.offset
          );
          log(`Defect export: ${defectPath}`);
          defectPaths.push(defectPath);
        }
      }

      const merged = emptyDefectCounts();
      let mappedRows = 0;
      let exportRows = 0;
      let skippedUnknownState = 0;
      const unknownStates = new Map();
      for (const defectPath of defectPaths) {
        const agg = workstreamFill.aggregateWorkstreamDefects(defectPath);
        mappedRows += agg.mappedRows;
        exportRows += agg.exportRows;
        skippedUnknownState += agg.skippedUnknownState;
        merged.closed += agg.counts.closed;
        merged.deferred += agg.counts.deferred;
        merged.fixed += agg.counts.fixed;
        merged.pending += agg.counts.pending;
        for (const u of agg.unknownStates || []) {
          unknownStates.set(u.state, (unknownStates.get(u.state) || 0) + (u.count || 0));
        }
      }
      merged.total = merged.closed + merged.deferred + merged.fixed + merged.pending;
      const denom = merged.total - merged.deferred;
      const combined = {
        counts: merged,
        mappedRows,
        exportRows,
        skippedUnknownState,
        unknownStates: [...unknownStates.entries()].map(([state, count]) => ({
          state,
          count,
        })),
        closureRate: denom > 0 ? merged.closed / denom : 0,
        resolutionRate: denom > 0 ? (merged.closed + merged.fixed) / denom : 0,
      };
      log(
        `Workstream defects mapped=${combined.mappedRows} exportRows=${combined.exportRows} unknownState=${combined.skippedUnknownState}`
      );
      log(
        `  Totals: C=${merged.closed} D=${merged.deferred} F=${merged.fixed} P=${merged.pending} T=${merged.total}`
      );
      if (combined.unknownStates.length) {
        log("\nALERT: Other defect statuses were found and ignored:");
        for (const u of combined.unknownStates) {
          log(`  - ${u.state}: ${u.count}`);
        }
      }
      workstreamFill.writeWorkstreamDefectSection(sheet, defectSections[0], combined);
    }
  } else if (includeDefects) {
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
        const defectSprint = String(args.defectSprint || "").trim();
        if (sprint3Run && includeDefects) {
          if (!defectSprint) {
            throw new ReporterError(
              "CONFIG_INVALID",
              "Sprint 3 needs a Sprint filter for defects (for example: Build Wave 1 Sprint 2)."
            );
          }
          log(`\nSprint 3 defect filter: ${defectSprint}`);
        }
        let defectAgg = null;
        if (args.localDefects) {
          const defectPath = resolvePathMaybe(args.localDefects);
          if (!fs.existsSync(defectPath)) {
            throw new ReporterError(
              "EXTRACT_FAILED",
              `Unable to extract defect data. Local file was not found: ${defectPath}`
            );
          }
          log(`Using local defect export: ${defectPath}`);
          defectAgg = aggregateDefectExport(defectPath, {
            sprintFilter: sprint3Run ? defectSprint : null,
          });
          if (defectAgg.skippedSprint) {
            log(
              `  ${defectAgg.skippedSprint} defect row(s) did not match Sprint filter "${defectSprint}", so they were skipped.`
            );
          }
        } else {
          const defectToken = token || getBearerToken();
          const defectProjectIds = [
            ...new Set(
              mapped
                .map((p) => Number(p.projectId || runtime.projectId))
                .filter(Boolean)
            ),
          ];
          if (sitRun) {
            log(
              `\nSIT: pulling defects from project(s) ${defectProjectIds.join(", ")} (Uganda + Tanzania).`
            );
          }
          const aggs = [];
          for (const pid of defectProjectIds) {
            if (defectToken && !args.skipSwitch) {
              log(`\nSwitching project to projectId=${pid} for defects...`);
              await switchProject(defectToken, pid);
              log("Project switch OK");
            }
            const defaultEntity =
              sitRun && runtime.projectIdB && Number(pid) === Number(runtime.projectIdB)
                ? "General Tanzania"
                : null;
            const agg = await downloadAndAggregateDefects(
              defectToken,
              pid,
              runtime.timezone,
              runtime.offset,
              log,
              {
                defaultEntity,
                sprint3: sprint3Run,
                sprintFilter: sprint3Run ? defectSprint : null,
              }
            );
            if (agg.skippedSprint) {
              log(
                `  ${agg.skippedSprint} defect row(s) did not match Sprint filter "${defectSprint}", so they were skipped.`
              );
            }
            if (agg.usedDefaultEntity) {
              log(
                `  ${agg.usedDefaultEntity} defect row(s) had a blank ENTITY; counted under ${defaultEntity}.`
              );
            }
            if (agg.skippedNoEntity) {
              log(
                `  ALERT: ${agg.skippedNoEntity} defect row(s) had no ENTITY and no project default, so they were skipped.`
              );
              if (agg.columns && agg.columns.length) {
                log(`  Defect export columns: ${agg.columns.join(", ")}`);
              }
            }
            aggs.push(agg);
          }
          defectAgg = mergeDefectAggregates(aggs);
        }
        log(
          `Defects mapped=${defectAgg.mappedRows} exportRows=${defectAgg.exportRows} unknownState=${defectAgg.skippedUnknownState}`
        );
        if (defectAgg.byEntity && defectAgg.byEntity.size) {
          log(
            `Defect entities found: ${[...defectAgg.byEntity.keys()].join(", ")}`
          );
        }
        if (defectAgg.unknownStates.length) {
          log("\nALERT: Other defect statuses were found and ignored:");
          for (const u of defectAgg.unknownStates) {
            log(`  - ${u.state}: ${u.count}`);
          }
          log(
            "Defined mapping used: Pending=New+Reopened, Closed=Resolved+Closed, Deferred=Deferred, Fixed=Fixed."
          );
        }

        // Fill each Defect Summary from this section's module rows plus any
        // modules that have defects for this entity but are not in this block.
        // Do not use a workbook-wide module set — another section listing
        // Procurement must not hide Life Uganda's Procurement defects.
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
          const entityCounts = entityCountsForSection(
            defectAgg.byEntity,
            section.entity
          );
          const extraRows = extraDefectModulesForSection(
            sectionModules,
            entityCounts
          );
          const summary = buildDefectSectionSummary(
            sectionModules,
            entityCounts,
            extraRows
          );
          if (extraRows.length) {
            log(
              `\nIncluding extra defect modules for [${section.entity}] (output only; template file unchanged):`
            );
            for (const r of extraRows) {
              log(
                `  - ${r.name}: C=${r.closed} D=${r.deferred} F=${r.fixed} P=${r.pending} T=${r.total}`
              );
            }
          }
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
    extraModules: extraModulesFound,
    extraModulesAdded: Boolean(args.addMissingModules) && extraModulesFound.length > 0,
  };
}

/**
 * Merge UI form values onto properties. If the form has no usable inputs,
 * return the file properties unchanged.
 */
function applyFormOverrides(baseProps, formInput) {
  const form = formInput || {};
  const projectId = String(form.projectId || "").trim();
  const projectIdB = String(form.projectIdB || "").trim();
  const templateChoice = String(form.templateChoice || "").trim();
  const planIds = Array.isArray(form.planIds)
    ? form.planIds.map((id) => String(id || "").trim()).filter(Boolean)
    : [];

  const hasFormInput = Boolean(projectId || projectIdB || templateChoice || planIds.length);
  if (!hasFormInput) return { ...baseProps };

  const props = { ...baseProps };
  if (projectId) props.PROJECT_ID = projectId;
  if (projectIdB) props.PROJECT_ID_B = projectIdB;
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
    const rel = props[key] || DEFAULT_TEMPLATES[i] || "";
    const abs = rel ? resolvePathMaybe(rel) : "";
    const inspected = abs ? workstreamFill.inspectTemplateFile(abs) : null;
    const meta = TEMPLATE_META[i] || TEMPLATE_META[String(i)] || {};
    const statusSections = meta.statusSections || (inspected && inspected.statusSections) || null;
    const kind = meta.kind || (inspected && inspected.kind) || null;
    let description = meta.description;
    if (!description && kind === "workstream") {
      description =
        "1 workstream block. All testcases in the selected plan are counted into one status row (not module-wise).";
    } else if (!description && statusSections) {
      description = `${statusSections} execution-plan block(s).`;
    } else if (!description && i >= 5) {
      description = "Custom template. No Module-wise Daily Status block found.";
    } else if (!description) {
      description = "Custom template";
    }
    choices.push({
      choice: String(i),
      template: rel,
      outputFile: props[`OUTPUT_FILE_${i}`] || DEFAULT_OUTPUT_FILES[i] || "",
      statusSections,
      kind,
      description,
      sitDualProject: Boolean(meta.sitDualProject),
      sprint3Defects: sprint3Defects.isSprint3TemplatePath(rel),
      deletable: i >= 5,
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
      },
      {
        choice: "4",
        template: DEFAULT_TEMPLATES[4],
        outputFile: DEFAULT_OUTPUT_FILES[4],
        ...TEMPLATE_META[4],
      }
    );
  }
  return choices;
}

async function fetchSimplifyQaProjects(token, origin = ORIGIN) {
  const defaultProjects = [
    { id: "2", name: "Financial Management System - Kenya" },
    { id: "5", name: "Financial Management System - Uganda" },
    { id: "6", name: "Financial Management System - Tanzania" },
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

function jsonRecords(data) {
  if (!data) return [];
  if (Array.isArray(data)) return data;
  if (Array.isArray(data.data)) return data.data;
  if (Array.isArray(data.records)) return data.records;
  if (Array.isArray(data.result)) return data.result;
  if (data.data && Array.isArray(data.data.content)) return data.data.content;
  if (data.data && Array.isArray(data.data.records)) return data.data.records;
  return [];
}

function unwrapPlanPayload(data) {
  if (!data || typeof data !== "object") return null;
  if (Array.isArray(data.data) && data.data[0]) return data.data[0];
  if (data.data && typeof data.data === "object" && !Array.isArray(data.data)) {
    return data.data;
  }
  if (data.result && typeof data.result === "object" && !Array.isArray(data.result)) {
    return data.result;
  }
  return data;
}

function planRecordId(row) {
  return String(
    (row && (row.id || row.exePlanId || row.executionPlanId || row.planId || row.code)) || ""
  ).trim();
}

function planRecordProjectId(row) {
  if (!row || typeof row !== "object") return "";
  const nested = row.project && typeof row.project === "object" ? row.project : null;
  return String(
    row.projectId ||
      row.projectID ||
      row.project_id ||
      (nested && (nested.id || nested.projectId || nested.projectID)) ||
      ""
  ).trim();
}

function planRecordName(row) {
  const name = String(
    (row &&
      (row.name ||
        row.executionPlanName ||
        row.planName ||
        row.title ||
        row.description ||
        row.code)) ||
      ""
  ).trim();
  return name;
}

function uniquePlanIds(ids) {
  const seen = new Set();
  const out = [];
  for (const raw of ids || []) {
    const id = String(raw || "").trim();
    if (!id || id === "custom" || seen.has(id)) continue;
    seen.add(id);
    out.push(id);
  }
  return out;
}

function qaAuthHeaders(token, origin, referer) {
  const cleanToken = token.startsWith("Bearer ") ? token : `Bearer ${String(token).trim()}`;
  return {
    Accept: "application/json, text/plain, */*",
    Authorization: cleanToken,
    origin,
    referer,
    "user-agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/151.0.0.0 Safari/537.36",
  };
}

/**
 * GET /tm/executionplan/{id} — this is the API that returns the EP name
 * (e.g. 7 → EP_Sprint2_BuildWave01_ Gen KE/Williamson).
 */
async function fetchSimplifyQaExecutionPlanById(token, planId, origin = ORIGIN) {
  const id = String(planId || "").trim();
  if (!token || !id) return null;
  try {
    const res = await fetch(`${origin}/tm/executionplan/${encodeURIComponent(id)}`, {
      method: "GET",
      headers: qaAuthHeaders(token, origin, `${origin}/user/execution-plan/details/${id}`),
    });
    if (!res.ok) return null;
    const data = await res.json();
    const row = unwrapPlanPayload(data);
    const resolvedId = planRecordId(row) || id;
    const name = planRecordName(row);
    if (!resolvedId) return null;
    return {
      id: resolvedId,
      name: name || `Plan ${resolvedId}`,
      label: `${resolvedId} — ${name || `Plan ${resolvedId}`}`,
      projectId: planRecordProjectId(row),
    };
  } catch {
    return null;
  }
}

/**
 * List execution plans for a project via POST /tm/executionplan/search,
 * then fill any missing names with GET /tm/executionplan/{id}.
 */
async function fetchSimplifyQaExecutionPlans(token, projectId, planIds = [], origin = ORIGIN) {
  if (!token) return [];
  const planMap = new Map();
  const project = String(projectId || "").trim();

  if (project && project !== "custom") {
    try {
      const res = await fetch(`${origin}/tm/executionplan/search`, {
        method: "POST",
        headers: {
          ...qaAuthHeaders(token, origin, `${origin}/user/execution-plan`),
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          searchFields: [
            { column: "id", sort: "dsc" },
            { column: "deleted", value: false, regEx: false },
            { column: "projectId", value: Number(project) || project },
          ],
          startIndex: 0,
          limit: 1000,
        }),
      });
      if (res.ok) {
        const data = await res.json();
        for (const row of jsonRecords(data)) {
          const id = planRecordId(row);
          const name = planRecordName(row);
          if (!id) continue;
          planMap.set(id, {
            id,
            name: name || `Plan ${id}`,
            label: `${id} — ${name || `Plan ${id}`}`,
            projectId: planRecordProjectId(row) || project,
          });
        }
      }
    } catch {}
  }

  const extraIds = uniquePlanIds(planIds).filter((id) => {
    const existing = planMap.get(id);
    return !existing || !existing.name || existing.name === `Plan ${id}`;
  });
  if (extraIds.length) {
    const extras = await Promise.all(
      extraIds.map((id) => fetchSimplifyQaExecutionPlanById(token, id, origin))
    );
    for (const plan of extras) {
      if (!plan || !plan.id) continue;
      const id = String(plan.id);
      const existing = planMap.get(id);
      if (existing) {
        planMap.set(id, {
          ...existing,
          ...plan,
          projectId: existing.projectId || plan.projectId || project,
        });
        continue;
      }
      // Never inject a plan from another project into this dropdown.
      if (project && plan.projectId && String(plan.projectId) !== String(project)) {
        continue;
      }
      if (project && !plan.projectId) continue;
      planMap.set(id, plan);
    }
  }

  return Array.from(planMap.values());
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
    projectIdB: String(props.PROJECT_ID_B || "").trim(),
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
    addMissingModules:
      form.addMissingModules === true || options.addMissingModules === true,
    defectSprint: String(form.defectSprint || options.defectSprint || "").trim(),
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
        projectIdB: runtime.projectIdB || null,
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
    if (
      /^-\s/.test(trimmed) ||
      /^Add these modules/i.test(trimmed)
    ) {
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
  fetchSimplifyQaExecutionPlans,
  listTemplateChoices,
  executeReport,
  runReport,
  extractAlerts,
  discoverStatusSections,
  getTokenStatus,
  saveBearerToken,
  resolveModuleName,
  resolveDefectModuleName,
  moduleMatchKey,
  passRateBand,
  applyPassRateFill,
  restoreStatusRateRowMerges,
  isFileLockError,
  defectExportHasEntityColumn,
  isEntityPreferenceColumn,
  findEntityPreferenceIndex,
  defectEntityColumnForProject,
  defectFilterPreferenceBody,
  mergeEntityIntoDefectFilter,
  resolveDefectFilterForProject,
  capturedDefectFilterForProject,
  preferenceHasProjectEntity,
  resolveDefectRowEntities,
  aggregateDefectExport,
  mergeDefectAggregates,
  entityCountsForSection,
  extraDefectModulesForSection,
  buildDefectSectionSummary,
};

if (require.main === module) {
  main().catch((err) => {
    console.error(err.message || err);
    process.exit(1);
  });
}
