const ExcelJS = require("exceljs");

const MONTHS = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sept",
  "Oct",
  "Nov",
  "Dec",
];

const DATE_RE = /^(?:[1-9]|[12]\d|3[01]) Sept \d{4}$/;

const STATUS_ORDER = ["General Uganda", "Life Uganda", "General Tanzania"];
const DEFECT_ORDER = STATUS_ORDER;

function cellRaw(cell) {
  if (!cell) return null;
  const v = cell.value;
  if (v == null || v === "") return null;
  if (typeof v === "object") {
    if (v.richText) return v.richText.map((t) => t.text).join("");
    if (v.formula != null || v.sharedFormula != null) {
      return v.result != null ? v.result : null;
    }
    if (v.result != null) return v.result;
    if (v.text != null) return v.text;
    if (v instanceof Date) return v;
  }
  return v;
}

function cellText(cell) {
  const v = cellRaw(cell);
  if (v == null) return "";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v).replace(/\s+/g, " ").trim();
}

function cellNumber(cell) {
  const v = cellRaw(cell);
  if (v == null || v === "") return 0;
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const n = Number(String(v).replace(/[% ,]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

/** Excel percent cells are 0–1; whole percents like 34 stay 34. */
function cellRate(cell) {
  const v = cellRaw(cell);
  if (v == null || v === "") return 0;
  if (typeof v === "number" && Number.isFinite(v)) {
    return v <= 1 ? v : v / 100;
  }
  const s = String(v).trim();
  if (s.endsWith("%")) {
    const n = Number(s.replace("%", "").trim());
    return Number.isFinite(n) ? n / 100 : 0;
  }
  const n = Number(s);
  if (!Number.isFinite(n)) return 0;
  return n <= 1 ? n : n / 100;
}

function formatPct(rate) {
  const n = Math.round((Number(rate) || 0) * 100);
  return `${n}%`;
}

function pctInt(rate) {
  return Math.round((Number(rate) || 0) * 100);
}

function entityFromTitle(title) {
  const t = String(title || "").toLowerCase();
  if (t.includes("life") && t.includes("uganda")) return "Life Uganda";
  if (t.includes("tanzania") || /\btz\b/.test(t)) return "General Tanzania";
  if (t.includes("uganda")) return "General Uganda";
  return null;
}

function isHeaderStatus(row) {
  const joined = row.map((c) => String(c || "").toLowerCase()).join(" ");
  return joined.includes("module") && joined.includes("passed") && joined.includes("pass");
}

function isHeaderDefect(row) {
  const joined = row.map((c) => String(c || "").toLowerCase()).join(" ");
  return joined.includes("module") && joined.includes("closed") && joined.includes("pending");
}

function rowValues(ws, r, cols) {
  const out = [];
  for (let c = 1; c <= cols; c++) out.push(cellText(ws.getRow(r).getCell(c)));
  return out;
}

function findCol(headerRow, ...needles) {
  const lower = headerRow.map((h) => String(h || "").toLowerCase().trim());
  for (const needle of needles) {
    const i = lower.findIndex((h) => h === needle || h.startsWith(needle));
    if (i >= 0) return i;
  }
  return -1;
}

function nearlyEqual(a, b, eps = 0.015) {
  return Math.abs(Number(a) - Number(b)) <= eps;
}

function parseSheetDateSuggestion(sheetName) {
  const m = String(sheetName || "").trim().match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (!m) return "";
  const day = Number(m[1]);
  const month = Number(m[2]);
  const year = Number(m[3]);
  if (month < 1 || month > 12 || day < 1 || day > 31) return "";
  return `${day} ${MONTHS[month - 1]} ${year}`;
}

function validateMeetingDate(value) {
  const text = String(value || "").trim();
  if (!DATE_RE.test(text)) {
    return {
      ok: false,
      message: 'Date must look like 16 Sept 2026 (day, then "Sept", then year).',
    };
  }
  const day = Number(text.split(" ")[0]);
  if (day < 1 || day > 31) {
    return { ok: false, message: "Enter a valid day between 1 and 31." };
  }
  return { ok: true, value: text };
}

function titleCaseWords(s) {
  return String(s || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(" ");
}

function splitCoverTitle(mainTitle) {
  const raw = String(mainTitle || "").trim() || "WEEKLY WORKSTREAM MEETING";
  const headline = raw.toUpperCase();
  const pretty = titleCaseWords(raw);
  const parts = pretty.split(/\s+/);
  let line1 = pretty;
  let line2 = "";
  if (parts.length >= 2 && parts[parts.length - 1].toLowerCase() === "meeting") {
    line2 = parts.pop();
    line1 = parts.join(" ");
  }
  return { headline, line1, line2: line2 || " " };
}

function outputPptFileName(title, dateText) {
  const t = String(title || "Weekly Workstream Meeting")
    .replace(/[<>:"/\\|?*]/g, " ")
    .trim()
    .replace(/\s+/g, "_");
  const d = String(dateText || "")
    .trim()
    .replace(/\s+/g, "_");
  return `${t}_${d || "report"}.pptx`;
}

async function readRows(filePath, sheetName) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws =
    (sheetName && wb.getWorksheet(sheetName)) ||
    wb.worksheets[0];
  if (!ws) throw new Error("The workbook has no sheets.");
  const cols = Math.max(ws.columnCount || 0, 10);
  const rows = [];
  const meta = [];
  ws.eachRow({ includeEmpty: true }, (row, r) => {
    const texts = [];
    const nums = [];
    const rates = [];
    const formulas = [];
    for (let c = 1; c <= cols; c++) {
      const cell = row.getCell(c);
      texts.push(cellText(cell));
      nums.push(cellNumber(cell));
      rates.push(cellRate(cell));
      const v = cell.value;
      const formula =
        v && typeof v === "object" ? v.formula || v.sharedFormula || "" : "";
      const hasResult =
        v && typeof v === "object" && (v.formula || v.sharedFormula)
          ? v.result != null && v.result !== ""
          : true;
      formulas.push({ formula: String(formula || ""), hasResult });
    }
    rows[r] = texts;
    meta[r] = { nums, rates, formulas };
  });
  return { ws, sheetName: ws.name, rows, meta, cols };
}

function parseStatusBlock(rows, meta, startRow, maxRow) {
  const title = (rows[startRow] || []).join(" ").trim();
  const entity = entityFromTitle(title);
  let headerRow = -1;
  for (let r = startRow; r <= Math.min(startRow + 3, maxRow); r++) {
    if (isHeaderStatus(rows[r] || [])) {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) return null;
  const header = rows[headerRow] || [];
  const cMod = findCol(header, "module");
  const cPass = findCol(header, "passed");
  const cFail = findCol(header, "failed");
  const cBlock = findCol(header, "blocked");
  const cIp = findCol(header, "in progress");
  const cNe = findCol(header, "not executed");
  const cTot = findCol(header, "total");
  const cRate = findCol(header, "pass rate");
  if (cMod < 0 || cPass < 0) return null;

  const modules = [];
  let totals = null;
  let executionRate = 0;
  let overallPassRate = 0;
  let end = headerRow;

  for (let r = headerRow + 1; r <= maxRow; r++) {
    const texts = rows[r] || [];
    const joined = texts.join(" ").toLowerCase();
    const name = texts[cMod] || "";
    if (/defect\s+summary/i.test(joined) || /sit\s*\(qa\)/i.test(joined) && /module-wise|general |life /i.test(joined) && r > headerRow + 1) {
      break;
    }
    if (/execution rate/i.test(joined)) {
      executionRate = firstRate(meta[r], texts);
      end = r;
      continue;
    }
    if (/overall pass rate/i.test(joined)) {
      overallPassRate = firstRate(meta[r], texts);
      end = r;
      break;
    }
    if (!name) continue;
    if (/^total$/i.test(name.trim())) {
      totals = moduleFromRow(name, texts, meta[r], {
        cPass, cFail, cBlock, cIp, cNe, cTot, cRate,
      });
      end = r;
      continue;
    }
    if (/lead champion/i.test(name) || /^module$/i.test(name)) continue;
    modules.push(
      moduleFromRow(name, texts, meta[r], {
        cPass, cFail, cBlock, cIp, cNe, cTot, cRate,
      })
    );
    end = r;
  }

  if (!totals && modules.length) {
    totals = sumModules(modules);
  }

  return {
    kind: "status",
    title,
    entity,
    modules,
    totals,
    executionRate,
    overallPassRate,
    startRow,
    endRow: end,
  };
}

function firstRate(metaRow, texts) {
  for (const t of texts || []) {
    if (String(t).includes("%")) {
      const n = Number(String(t).replace("%", "").trim());
      if (Number.isFinite(n)) return n / 100;
    }
  }
  if (metaRow && metaRow.rates) {
    for (const r of metaRow.rates) {
      if (r > 0 && r <= 1) return r;
    }
  }
  return 0;
}

function moduleFromRow(name, texts, metaRow, cols) {
  const nums = (metaRow && metaRow.nums) || [];
  const rates = (metaRow && metaRow.rates) || [];
  const n = (i) => (i >= 0 ? Number(nums[i]) || 0 : 0);
  const passed = n(cols.cPass);
  const failed = n(cols.cFail);
  const blocked = n(cols.cBlock);
  const inProgress = n(cols.cIp);
  const notExecuted = n(cols.cNe);
  let total = n(cols.cTot);
  if (!total) total = passed + failed + blocked + inProgress + notExecuted;
  let passRate = cols.cRate >= 0 ? rates[cols.cRate] || 0 : 0;
  if (!passRate) {
    const denom = total - blocked;
    passRate = denom > 0 ? passed / denom : 0;
  }
  return {
    name: String(name || "").trim(),
    passed,
    failed,
    blocked,
    inProgress,
    notExecuted,
    total,
    passRate,
  };
}

function sumModules(modules) {
  const t = {
    name: "Total",
    passed: 0,
    failed: 0,
    blocked: 0,
    inProgress: 0,
    notExecuted: 0,
    total: 0,
    passRate: 0,
  };
  for (const m of modules) {
    t.passed += m.passed;
    t.failed += m.failed;
    t.blocked += m.blocked;
    t.inProgress += m.inProgress;
    t.notExecuted += m.notExecuted;
    t.total += m.total;
  }
  const denom = t.total - t.blocked;
  t.passRate = denom > 0 ? t.passed / denom : 0;
  return t;
}

function parseDefectBlock(rows, meta, startRow, maxRow) {
  const title = (rows[startRow] || []).join(" ").trim();
  const entity = entityFromTitle(title);
  let headerRow = -1;
  for (let r = startRow; r <= Math.min(startRow + 3, maxRow); r++) {
    if (isHeaderDefect(rows[r] || [])) {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) return null;
  const header = rows[headerRow] || [];
  const cMod = findCol(header, "module");
  const cClosed = findCol(header, "closed");
  const cDef = findCol(header, "deferred");
  const cFix = findCol(header, "fixed");
  const cPend = findCol(header, "pending");
  const cTot = findCol(header, "total");
  const modules = [];
  let totals = null;
  let closureRate = 0;
  let resolutionRate = 0;
  let end = headerRow;

  for (let r = headerRow + 1; r <= maxRow; r++) {
    const texts = rows[r] || [];
    const joined = texts.join(" ").toLowerCase();
    const name = cMod >= 0 ? texts[cMod] : texts[2] || texts[1] || "";
    if (/defect\s+summary/i.test(joined) && r > headerRow + 1) break;
    if (/sit\s*\(qa\)/i.test(joined) && /general |life /i.test(joined)) break;
    if (/closure rate/i.test(joined)) {
      closureRate = firstRate(meta[r], texts);
      end = r;
      continue;
    }
    if (/resolution rate/i.test(joined)) {
      resolutionRate = firstRate(meta[r], texts);
      end = r;
      break;
    }
    if (!name) continue;
    if (/^total$/i.test(name.trim())) {
      totals = defectFromRow(name, texts, meta[r], {
        cClosed, cDef, cFix, cPend, cTot,
      });
      end = r;
      continue;
    }
    if (/^module$/i.test(name)) continue;
    modules.push(
      defectFromRow(name, texts, meta[r], { cClosed, cDef, cFix, cPend, cTot })
    );
    end = r;
  }
  if (!totals && modules.length) totals = sumDefects(modules);
  return {
    kind: "defects",
    title,
    entity,
    modules,
    totals,
    closureRate,
    resolutionRate,
    startRow,
    endRow: end,
  };
}

function defectFromRow(name, texts, metaRow, cols) {
  const nums = (metaRow && metaRow.nums) || [];
  const n = (i) => (i >= 0 ? Number(nums[i]) || 0 : 0);
  const closed = n(cols.cClosed);
  const deferred = n(cols.cDef);
  const fixed = n(cols.cFix);
  const pending = n(cols.cPend);
  let total = n(cols.cTot);
  if (!total) total = closed + deferred + fixed + pending;
  return { name: String(name || "").trim(), closed, deferred, fixed, pending, total };
}

function sumDefects(modules) {
  const t = {
    name: "Total",
    closed: 0,
    deferred: 0,
    fixed: 0,
    pending: 0,
    total: 0,
  };
  for (const m of modules) {
    t.closed += m.closed;
    t.deferred += m.deferred;
    t.fixed += m.fixed;
    t.pending += m.pending;
    t.total += m.total;
  }
  return t;
}

function auditStatus(section, alerts, formulaNotes) {
  const { modules, totals } = section;
  for (const m of modules) {
    const sum = m.passed + m.failed + m.blocked + m.inProgress + m.notExecuted;
    if (sum !== m.total) {
      alerts.push(
        `${section.entity || section.title}: ${m.name} Total ${m.total} does not match Passed+Failed+Blocked+In Progress+Not Executed (${sum}). Using the Excel displayed values.`
      );
    }
    const denom = m.total - m.blocked;
    const expected = denom > 0 ? m.passed / denom : 0;
    if (!nearlyEqual(expected, m.passRate)) {
      alerts.push(
        `${section.entity || section.title}: ${m.name} Pass Rate ${formatPct(m.passRate)} does not match Passed / (Total − Blocked) (${formatPct(expected)}). Using the Excel displayed values.`
      );
    }
  }
  if (totals && modules.length) {
    const summed = sumModules(modules);
    for (const k of ["passed", "failed", "blocked", "inProgress", "notExecuted", "total"]) {
      if (summed[k] !== totals[k]) {
        alerts.push(
          `${section.entity || section.title}: Total ${k} ${totals[k]} does not match the sum of module rows (${summed[k]}). Using the Excel displayed values.`
        );
        break;
      }
    }
  }
  if (totals) {
    const denom = totals.total - totals.blocked;
    const exec = denom > 0 ? (totals.passed + totals.failed) / denom : 0;
    const pass = denom > 0 ? totals.passed / denom : 0;
    if (!nearlyEqual(exec, section.executionRate)) {
      alerts.push(
        `${section.entity || section.title}: Execution Rate ${formatPct(section.executionRate)} does not match (Passed+Failed)/(Total−Blocked) (${formatPct(exec)}). Using the Excel displayed values.`
      );
    }
    if (!nearlyEqual(pass, section.overallPassRate)) {
      alerts.push(
        `${section.entity || section.title}: Overall Pass Rate ${formatPct(section.overallPassRate)} does not match Passed/(Total−Blocked) (${formatPct(pass)}). Using the Excel displayed values.`
      );
    }
  }
  for (const note of formulaNotes || []) alerts.push(note);
}

function auditDefects(section, alerts) {
  const { modules, totals } = section;
  for (const m of modules) {
    const sum = m.closed + m.deferred + m.fixed + m.pending;
    if (sum !== m.total) {
      alerts.push(
        `${section.entity || section.title} defects: ${m.name} Total ${m.total} does not match Closed+Deferred+Fixed+Pending (${sum}). Using the Excel displayed values.`
      );
    }
  }
  if (totals) {
    const denom = totals.total - totals.deferred;
    const closure = denom > 0 ? totals.closed / denom : 0;
    const resolution = denom > 0 ? (totals.closed + totals.fixed) / denom : 0;
    if (!nearlyEqual(closure, section.closureRate)) {
      alerts.push(
        `${section.entity || section.title}: Closure Rate ${formatPct(section.closureRate)} does not match Closed/(Total−Deferred) (${formatPct(closure)}). Using the Excel displayed values.`
      );
    }
    if (!nearlyEqual(resolution, section.resolutionRate)) {
      alerts.push(
        `${section.entity || section.title}: Resolution Rate ${formatPct(section.resolutionRate)} does not match (Closed+Fixed)/(Total−Deferred) (${formatPct(resolution)}). Using the Excel displayed values.`
      );
    }
  }
}

function collectFormulaNotes(meta, maxRow) {
  const notes = [];
  let missing = 0;
  for (let r = 1; r <= maxRow; r++) {
    const row = meta[r];
    if (!row) continue;
    for (const f of row.formulas || []) {
      if (f.formula && !f.hasResult) missing += 1;
    }
  }
  if (missing) {
    notes.push(
      `${missing} formula cell(s) have no cached result. Values shown in Excel will still be used.`
    );
  }
  return notes;
}

async function parseDsrWorkbook(filePath, sheetName) {
  const { rows, meta, sheetName: resolved } = await readRows(filePath, sheetName);
  const maxRow = rows.length - 1;
  const status = [];
  const defects = [];
  let r = 1;
  while (r <= maxRow) {
    const joined = (rows[r] || []).join(" ");
    if (/defect\s+summary/i.test(joined)) {
      const block = parseDefectBlock(rows, meta, r, maxRow);
      if (block) {
        defects.push(block);
        r = block.endRow + 1;
        continue;
      }
    }
    if (/sit\s*\(qa\)/i.test(joined) && !/defect/i.test(joined)) {
      const block = parseStatusBlock(rows, meta, r, maxRow);
      if (block) {
        status.push(block);
        r = block.endRow + 1;
        continue;
      }
    }
    r += 1;
  }

  const alerts = [];
  const formulaNotes = collectFormulaNotes(meta, maxRow);
  alerts.push(...formulaNotes);
  for (const s of status) auditStatus(s, alerts, []);
  for (const d of defects) auditDefects(d, alerts);

  const byEntityStatus = new Map();
  for (const s of status) {
    if (s.entity) byEntityStatus.set(s.entity, s);
  }
  const byEntityDefects = new Map();
  for (const d of defects) {
    if (d.entity) byEntityDefects.set(d.entity, d);
  }

  return {
    sheetName: resolved,
    suggestedDate: parseSheetDateSuggestion(resolved),
    status,
    defects,
    statusByEntity: Object.fromEntries(byEntityStatus),
    defectsByEntity: Object.fromEntries(byEntityDefects),
    alerts,
    statusOrder: STATUS_ORDER,
    defectOrder: DEFECT_ORDER,
  };
}

module.exports = {
  DATE_RE,
  STATUS_ORDER,
  DEFECT_ORDER,
  MONTHS,
  parseDsrWorkbook,
  parseSheetDateSuggestion,
  validateMeetingDate,
  splitCoverTitle,
  outputPptFileName,
  formatPct,
  pctInt,
  titleCaseWords,
};
