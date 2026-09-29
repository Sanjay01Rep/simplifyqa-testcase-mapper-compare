/**
 * SIT Entity Summary — fills the block to the right of Daily Status
 * in Template/SIT Template.xlsx (columns M–V, rows 2–8).
 *
 *   M2:V2  headers
 *   M3:V4  General Uganda (main + delta)
 *   M5:V6  Life Uganda
 *   M7:V8  General Tanzania
 *
 * Merged: Entity, Total, Not Executed, Status across each pair of rows.
 * Deltas and "was N%" come from a previous dated sheet when compare is on.
 */

const ENTITY_HEADER = "Entity";
const COMPARE_SHEET_PREFIX = /^compare /i;

const COL = {
  entity: 13,
  total: 14,
  passed: 15,
  failed: 16,
  blocked: 17,
  inProgress: 18,
  notExecuted: 19,
  execution: 20,
  pass: 21,
  status: 22,
};

const HEADER_ROW = 2;

const ROW_POSITIONS = [
  { label: "General Uganda", mainRow: 3, deltaRow: 4 },
  { label: "Life Uganda", mainRow: 5, deltaRow: 6 },
  { label: "General Tanzania", mainRow: 7, deltaRow: 8 },
];

const BAND = {
  red: "FFFF0000",
  amber: "FFFFC000",
  green: "FF92D050",
};

/** Pastel fills from the SIT Entity Summary template (Passed / Failed / Blocked). */
const METRIC_FILL = {
  passed: "FFE2F0D9",
  failed: "FFFBE5E5",
  blocked: "FFFFF2CC",
};

const FONT = {
  good: "FF2E7D32",
  bad: "FFC62828",
  muted: "FF888888",
  body: "FF222222",
};

/** Same grey as Daily Status Lead Champion (theme 2, tint -0.25). */
const ENTITY_HEADER_FILL = {
  type: "pattern",
  pattern: "solid",
  fgColor: { theme: 2, tint: -0.25 },
};

const ENTITY_HEADER_FONT = {
  bold: true,
  size: 12,
  name: "Calibri",
  color: { argb: "FF000000" },
};

/** Excel character widths so In Progress / Not Executed / Execution % are not clipped. */
const ENTITY_COL_MIN_WIDTH = {
  inProgress: 13,
  notExecuted: 14,
  execution: 13,
};

function cellText(value) {
  if (value == null) return "";
  if (typeof value === "object") {
    if (value.richText) return value.richText.map((t) => t.text).join("");
    if (value.text != null) return String(value.text);
    if (value.result != null && typeof value.result !== "object") return String(value.result);
  }
  return String(value);
}

function numberOf(value) {
  if (value == null || value === "") return 0;
  if (typeof value === "number") return value;
  if (typeof value === "object" && value.result != null && typeof value.result === "number") {
    return value.result;
  }
  const n = Number(String(value).replace(/%/g, "").trim());
  return Number.isFinite(n) ? n : 0;
}

/** Stored as Excel % fraction (0.42). Whole numbers 42 mean 42%. */
function percentOf(value) {
  const n = numberOf(value);
  if (Math.abs(n) > 1) return n / 100;
  return n;
}

function rate(numerator, denominator) {
  return denominator > 0 ? numerator / denominator : 0;
}

/** Same 0–50 Red, 51–80 Amber, 81–100 Green as Pass Rate cells. */
function classifyStatus(passRate) {
  const pct = Math.round((Number(passRate) || 0) * 100);
  if (pct <= 50) return { label: "Red", argb: BAND.red };
  if (pct <= 80) return { label: "Amber", argb: BAND.amber };
  return { label: "Green", argb: BAND.green };
}

function deriveEntityLabel(sectionTitle) {
  return String(sectionTitle || "").split("|")[0].trim();
}

function buildEntityRows(sitStatusSummaries) {
  return (sitStatusSummaries || []).map(({ sectionTitle, summary }) => {
    const t = summary.totalSummary;
    const executionRate =
      summary.executionRate != null
        ? summary.executionRate
        : rate(t.passed + t.failed, t.total - t.blocked);
    const passRate =
      summary.overallPassRate != null
        ? summary.overallPassRate
        : rate(t.passed, t.total - t.blocked);
    return {
      entityLabel: deriveEntityLabel(sectionTitle),
      total: t.total,
      passed: t.passed,
      failed: t.failed,
      blocked: t.blocked,
      inProgress: t.inProgress,
      notExecuted: t.notExecuted,
      executionRate,
      passRate,
    };
  });
}

function isDatedRunSheet(ws) {
  if (!ws || !ws.name) return false;
  if (COMPARE_SHEET_PREFIX.test(ws.name)) return false;
  if (String(ws.name).trim().toLowerCase() === "entity summary") return false;
  return true;
}

function findPreviousDatedSheet(existingWorksheets) {
  const list = existingWorksheets || [];
  for (let i = list.length - 1; i >= 0; i--) {
    if (!isDatedRunSheet(list[i])) continue;
    if (readEntitySummary(list[i])) return list[i];
  }
  return null;
}

function resolveCompareSheet(existingWorksheets, preferredName) {
  const list = existingWorksheets || [];
  const want = String(preferredName || "").trim().toLowerCase();
  if (want) {
    const named = list.find((ws) => String(ws.name || "").trim().toLowerCase() === want);
    if (named && readEntitySummary(named)) return named;
  }
  return findPreviousDatedSheet(list);
}

/** Some older sheets shifted Entity Summary to N after the Daily Status title overlap. */
function detectEntityStartCol(sheet) {
  if (!sheet) return 0;
  const candidates = [COL.entity, COL.entity + 1, COL.entity - 1, COL.entity + 2];
  for (const start of candidates) {
    if (start < 1) continue;
    const header = cellText(sheet.getRow(HEADER_ROW).getCell(start).value).trim();
    if (header === ENTITY_HEADER) return start;
  }
  for (const start of candidates) {
    if (start < 1) continue;
    const found = ROW_POSITIONS.some((pos) => {
      const label = cellText(sheet.getRow(pos.mainRow).getCell(start).value).trim();
      return label === pos.label;
    });
    if (found) return start;
  }
  return 0;
}

function colsFromStart(start) {
  const d = start - COL.entity;
  return {
    entity: COL.entity + d,
    total: COL.total + d,
    passed: COL.passed + d,
    failed: COL.failed + d,
    blocked: COL.blocked + d,
    inProgress: COL.inProgress + d,
    notExecuted: COL.notExecuted + d,
    execution: COL.execution + d,
    pass: COL.pass + d,
    status: COL.status + d,
  };
}

function sheetHasEntitySummary(sheet) {
  return detectEntityStartCol(sheet) > 0;
}

function readEntitySummaryBlock(sheet) {
  const start = detectEntityStartCol(sheet);
  if (!start) return null;
  const c = colsFromStart(start);
  const byLabel = {};
  for (const pos of ROW_POSITIONS) {
    const row = sheet.getRow(pos.mainRow);
    const label = cellText(row.getCell(c.entity).value).trim() || pos.label;
    const total = numberOf(row.getCell(c.total).value);
    const passed = numberOf(row.getCell(c.passed).value);
    if (!label || (total === 0 && passed === 0)) continue;
    byLabel[label] = {
      total,
      passed,
      failed: numberOf(row.getCell(c.failed).value),
      blocked: numberOf(row.getCell(c.blocked).value),
      inProgress: numberOf(row.getCell(c.inProgress).value),
      notExecuted: numberOf(row.getCell(c.notExecuted).value),
      executionRate: percentOf(row.getCell(c.execution).value),
      passRate: percentOf(row.getCell(c.pass).value),
    };
  }
  return Object.keys(byLabel).length ? byLabel : null;
}

/**
 * Older dated sheets have no Entity Summary block. Read the Daily Status
 * Total row for each SIT section so compare still works.
 */
function readDailyStatusTotals(sheet) {
  if (!sheet) return null;
  const byLabel = {};
  const last = Math.min(sheet.rowCount || 0, 120);
  for (let r = 1; r <= last; r++) {
    let title = "";
    for (let c = 1; c <= 3; c++) {
      const text = cellText(sheet.getRow(r).getCell(c).value).trim();
      if (/module-wise\s+daily\s+status/i.test(text)) {
        title = text;
        break;
      }
    }
    if (!title) continue;
    const label = deriveEntityLabel(title);
    if (!label || byLabel[label]) continue;
    for (let i = r + 2; i <= Math.min(r + 40, last); i++) {
      const moduleCell = cellText(sheet.getRow(i).getCell(3).value).trim().toLowerCase();
      if (moduleCell !== "total") continue;
      const row = sheet.getRow(i);
      const passed = numberOf(row.getCell(4).value);
      const failed = numberOf(row.getCell(5).value);
      const blocked = numberOf(row.getCell(6).value);
      const inProgress = numberOf(row.getCell(7).value);
      const notExecuted = numberOf(row.getCell(8).value);
      const total = numberOf(row.getCell(9).value) || passed + failed + blocked + inProgress + notExecuted;
      byLabel[label] = {
        total,
        passed,
        failed,
        blocked,
        inProgress,
        notExecuted,
        executionRate: rate(passed + failed, total - blocked),
        passRate: percentOf(row.getCell(10).value) || rate(passed, total - blocked),
      };
      break;
    }
  }
  return Object.keys(byLabel).length ? byLabel : null;
}

function readEntitySummary(sheet) {
  return readEntitySummaryBlock(sheet) || readDailyStatusTotals(sheet);
}

function applyFill(cell, argb) {
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb }, bgColor: { indexed: 64 } };
}

function cloneStylePart(value) {
  if (!value || typeof value !== "object") return value;
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return value;
  }
}

function applyMetricFill(cell, metric) {
  const argb = METRIC_FILL[metric];
  if (argb) applyFill(cell, argb);
}

function arrowFor(delta) {
  if (delta > 0) return "▲";
  if (delta < 0) return "▼";
  return "–";
}

function deltaFontArgb(metric, delta) {
  if (delta === 0) return FONT.muted;
  const up = delta > 0;
  if (metric === "passed") return up ? FONT.good : FONT.bad;
  if (metric === "failed" || metric === "blocked" || metric === "notExecuted") {
    return up ? FONT.bad : FONT.good;
  }
  if (metric === "inProgress") return up ? FONT.bad : FONT.good;
  return FONT.muted;
}

function writeDeltaCell(cell, latest, prev, metric) {
  const delta = Number(latest || 0) - Number(prev || 0);
  cell.value = delta === 0 ? "–" : `${arrowFor(delta)} ${Math.abs(delta)}`;
  cell.font = { size: 8, color: { argb: deltaFontArgb(metric, delta) } };
  cell.alignment = { horizontal: "center", wrapText: true };
  applyMetricFill(cell, metric);
}

function clearDeltaRow(deltaRow) {
  [COL.passed, COL.failed, COL.blocked, COL.inProgress, COL.execution, COL.pass].forEach((c) => {
    const cell = deltaRow.getCell(c);
    cell.value = null;
  });
}

const HEADERS = [
  "Entity",
  "Total",
  "Passed",
  "Failed",
  "Blocked",
  "In Progress",
  "Not Executed",
  "Execution %",
  "Pass %",
  "Status",
];

function writeEntityHeaders(sheet) {
  if (!sheet) return;
  const headerRow = sheet.getRow(HEADER_ROW);
  HEADERS.forEach((label, idx) => {
    const cell = headerRow.getCell(COL.entity + idx);
    cell.value = label;
    cell.font = { ...ENTITY_HEADER_FONT };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: false };
    cell.fill = cloneStylePart(ENTITY_HEADER_FILL);
  });
  applyEntityColumnWidths(sheet);
  headerRow.commit();
}

function applyEntityColumnWidths(sheet) {
  if (!sheet) return;
  const start = detectEntityStartCol(sheet) || COL.entity;
  const d = start - COL.entity;
  [
    ["inProgress", COL.inProgress],
    ["notExecuted", COL.notExecuted],
    ["execution", COL.execution],
  ].forEach(([key, base]) => {
    const col = sheet.getColumn(base + d);
    const min = ENTITY_COL_MIN_WIDTH[key];
    const current = Number(col.width) || 0;
    if (current < min) col.width = min;
  });
}

function writeEntityRows(sheet, entityRows, previousByLabel) {
  if (!sheet) return;
  writeEntityHeaders(sheet);
  for (const data of entityRows || []) {
    const pos = ROW_POSITIONS.find((p) => p.label === data.entityLabel);
    if (!pos) continue;

    const mainRow = sheet.getRow(pos.mainRow);
    mainRow.getCell(COL.entity).value = data.entityLabel;
    mainRow.getCell(COL.total).value = data.total;
    mainRow.getCell(COL.passed).value = data.passed;
    applyMetricFill(mainRow.getCell(COL.passed), "passed");
    mainRow.getCell(COL.failed).value = data.failed;
    applyMetricFill(mainRow.getCell(COL.failed), "failed");
    mainRow.getCell(COL.blocked).value = data.blocked;
    applyMetricFill(mainRow.getCell(COL.blocked), "blocked");
    mainRow.getCell(COL.inProgress).value = data.inProgress;
    mainRow.getCell(COL.notExecuted).value = data.notExecuted;
    const execCell = mainRow.getCell(COL.execution);
    execCell.value = data.executionRate;
    execCell.numFmt = "0%";
    const passCell = mainRow.getCell(COL.pass);
    passCell.value = data.passRate;
    passCell.numFmt = "0%";
    const status = classifyStatus(data.passRate);
    const statusCell = mainRow.getCell(COL.status);
    statusCell.value = status.label;
    applyFill(statusCell, status.argb);
    statusCell.font = { bold: true, size: 9, color: { argb: "FF000000" } };
    statusCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    mainRow.commit();

    const deltaRow = sheet.getRow(pos.deltaRow);
    clearDeltaRow(deltaRow);
    const prev = previousByLabel && previousByLabel[data.entityLabel];
    if (prev) {
      writeDeltaCell(deltaRow.getCell(COL.passed), data.passed, prev.passed, "passed");
      writeDeltaCell(deltaRow.getCell(COL.failed), data.failed, prev.failed, "failed");
      writeDeltaCell(deltaRow.getCell(COL.blocked), data.blocked, prev.blocked, "blocked");
      writeDeltaCell(
        deltaRow.getCell(COL.inProgress),
        data.inProgress,
        prev.inProgress,
        "inProgress"
      );
      const wasExec = deltaRow.getCell(COL.execution);
      wasExec.value = `was ${Math.round((prev.executionRate || 0) * 100)}%`;
      wasExec.font = { size: 8, color: { argb: FONT.muted } };
      wasExec.alignment = { horizontal: "center", wrapText: true };
      const wasPass = deltaRow.getCell(COL.pass);
      wasPass.value = `was ${Math.round((prev.passRate || 0) * 100)}%`;
      wasPass.font = { size: 8, color: { argb: FONT.muted } };
      wasPass.alignment = { horizontal: "center", wrapText: true };
    }
    deltaRow.commit();
  }
}

module.exports = {
  ENTITY_HEADER,
  COL,
  HEADER_ROW,
  ROW_POSITIONS,
  deriveEntityLabel,
  buildEntityRows,
  findPreviousDatedSheet,
  resolveCompareSheet,
  sheetHasEntitySummary,
  detectEntityStartCol,
  readEntitySummary,
  readDailyStatusTotals,
  writeEntityHeaders,
  writeEntityRows,
  classifyStatus,
  isDatedRunSheet,
};
