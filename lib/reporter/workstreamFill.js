/**
 * Workstream (non module-wise) reporter fill.
 * Used for templates like As and When Commission: one Daily Status Report row
 * and one Defects totals row. Counts every testcase / defect. Does not call
 * the module-wise mapper (detectLayout / aggregateExport / writeSection).
 */
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

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

const PASS_RATE_BANDS = {
  red: { argb: "FFFF0000" },
  amber: { argb: "FFFFC000" },
  green: { argb: "FF92D050" },
};

function cellText(value) {
  if (value == null) return "";
  if (typeof value === "object") {
    if (value.text) return String(value.text);
    if (value.richText) return value.richText.map((t) => t.text).join("");
    if (value.result != null) return String(value.result);
  }
  return String(value);
}

function normalizeStatus(raw) {
  const key = String(raw || "").trim().toUpperCase();
  return STATUS_MAP[key] || "notExecuted";
}

function mapDefectState(rawState) {
  const key = String(rawState || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
  if (!key) return null;
  return DEFECT_STATE_MAP[key] || null;
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

function emptyDefectCounts() {
  return { closed: 0, deferred: 0, fixed: 0, pending: 0 };
}

function rate(numerator, denominator) {
  if (!denominator) return 1;
  return numerator / denominator;
}

function summarizeCounts(counts) {
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

function setNumber(cell, value) {
  cell.value = Number(value) || 0;
}

function setFormula(cell, formula, result, numFmt) {
  cell.value = { formula, result: result == null ? 0 : result };
  if (numFmt) cell.numFmt = numFmt;
}

function passRateBand(rateValue) {
  const pct = Math.round((Number(rateValue) || 0) * 100);
  if (pct <= 50) return PASS_RATE_BANDS.red;
  if (pct <= 80) return PASS_RATE_BANDS.amber;
  return PASS_RATE_BANDS.green;
}

function applyPassRateFill(cell, rateValue) {
  if (!cell) return;
  if (rateValue == null || !Number.isFinite(Number(rateValue))) return;
  const band = passRateBand(rateValue);
  const fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: band.argb },
    bgColor: { argb: band.argb },
  };
  const font = Object.assign({}, cell.font || {});
  font.color = { argb: "FF000000" };
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

function rowHasLabel(sheet, rowNum, labelPrefix) {
  const row = sheet.getRow(rowNum);
  let found = false;
  row.eachCell({ includeEmpty: false }, (cell) => {
    const t = cellText(cell.value).trim().toLowerCase();
    if (t.startsWith(labelPrefix)) found = true;
  });
  return found;
}

function isModuleWiseTitle(text) {
  return /module-wise\s+daily\s+status/i.test(text);
}

function rowJoinedText(sheet, r, maxCol = 20) {
  const row = sheet.getRow(r);
  const last = Math.max(Number(row.cellCount) || 0, maxCol);
  const parts = [];
  for (let c = 1; c <= last; c++) {
    const t = cellText(row.getCell(c).value).trim();
    if (t) parts.push(t);
  }
  return parts.join(" ").replace(/\s+/g, " ").trim();
}

function isWorkstreamStatusTitle(text) {
  const t = String(text || "").trim();
  if (!t || isModuleWiseTitle(t)) return false;
  return /daily\s+status\s+report/i.test(t);
}

function looksLikeWorkstreamFile(filePath) {
  return /as\s*and\s*when/i.test(path.basename(filePath || ""));
}

function isWorkstreamDefectTitle(text) {
  const t = String(text || "").trim();
  if (!t) return false;
  if (/defect\s+summary/i.test(t)) return false;
  return /\bdefects?\b/i.test(t);
}

function detectWorkstreamLayout(sheet, titleRow) {
  const headerRow = titleRow + 1;
  const cols = {};
  sheet.getRow(headerRow).eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const t = cellText(cell.value).trim().toLowerCase();
    if (!t) return;
    if (t.includes("lead champion") || t === "champion") cols.champion = colNumber;
    else if (t === "module") cols.module = colNumber;
    else if (t.startsWith("passed")) cols.passed = colNumber;
    else if (t.startsWith("failed")) cols.failed = colNumber;
    else if (t.startsWith("blocked")) cols.blocked = colNumber;
    else if (t.includes("progress")) cols.inProgress = colNumber;
    else if (t.includes("not executed")) cols.notExecuted = colNumber;
    else if (t === "total") cols.total = colNumber;
    else if (t.includes("pass rate")) cols.passRate = colNumber;
  });
  if (cols.module) return null;
  if (!cols.passed || !cols.failed || !cols.total) return null;
  return { headerRow, firstDataRow: headerRow + 1, cols };
}

function detectWorkstreamDefectLayout(sheet, titleRow) {
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
  if (cols.module) return null;
  if (!cols.closed || !cols.deferred || !cols.fixed || !cols.pending || !cols.total) {
    return null;
  }
  return { headerRow, firstDataRow: headerRow + 1, cols };
}

function discoverWorkstreamStatusSections(sheet) {
  const sections = [];
  if (!sheet) return sections;
  const lastRow = Math.min(Number(sheet.rowCount) || 0, 80);
  for (let r = 1; r <= lastRow; r++) {
    let title = "";
    for (let c = 1; c <= 12; c++) {
      const text = cellText(sheet.getRow(r).getCell(c).value).trim();
      if (isWorkstreamStatusTitle(text)) {
        title = text;
        break;
      }
    }
    if (!title && isWorkstreamStatusTitle(rowJoinedText(sheet, r))) {
      title = rowJoinedText(sheet, r);
    }
    if (!title) continue;
    const layout = detectWorkstreamLayout(sheet, r);
    if (!layout) continue;
    sections.push({ title, titleRow: r, layout });
  }
  if (sections.length) return sections;

  for (let titleRow = 1; titleRow <= lastRow; titleRow++) {
    if (isModuleWiseTitle(rowJoinedText(sheet, titleRow))) continue;
    const layout = detectWorkstreamLayout(sheet, titleRow);
    if (!layout) continue;
    const title = rowJoinedText(sheet, titleRow) || "Daily Status Report";
    sections.push({ title, titleRow, layout });
    break;
  }
  return sections;
}

function discoverWorkstreamStatusInWorkbook(workbook, filePath) {
  const sheets = (workbook && workbook.worksheets) || [];
  for (const sheet of sheets) {
    const sections = discoverWorkstreamStatusSections(sheet);
    if (sections.length) return { sheet, sections };
  }
  if (looksLikeWorkstreamFile(filePath) && sheets.length) {
    const sections = discoverWorkstreamStatusSections(sheets[0]);
    if (sections.length) return { sheet: sheets[0], sections };
  }
  return { sheet: sheets[0] || null, sections: [] };
}

function discoverWorkstreamDefectSections(sheet) {
  const sections = [];
  if (!sheet) return sections;
  const lastRow = Math.min(Number(sheet.rowCount) || 0, 80);
  for (let r = 1; r <= lastRow; r++) {
    let title = "";
    for (let c = 1; c <= 12; c++) {
      const text = cellText(sheet.getRow(r).getCell(c).value).trim();
      if (isWorkstreamDefectTitle(text)) {
        title = text;
        break;
      }
    }
    if (!title && isWorkstreamDefectTitle(rowJoinedText(sheet, r))) {
      title = rowJoinedText(sheet, r);
    }
    if (!title) continue;
    const layout = detectWorkstreamDefectLayout(sheet, r);
    if (!layout) continue;
    sections.push({ title, titleRow: r, layout });
  }
  return sections;
}

function inspectTemplateFile(filePath) {
  try {
    if (!filePath || !fs.existsSync(filePath)) return null;
    const wb = XLSX.readFile(filePath);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    if (!sheet) return null;
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    let moduleWise = 0;
    let workstream = 0;
    for (let i = 0; i < rows.length; i++) {
      const joined = (rows[i] || []).map((c) => String(c || "")).join(" ");
      if (isModuleWiseTitle(joined)) {
        moduleWise += 1;
        continue;
      }
      if (!isWorkstreamStatusTitle(joined)) continue;
      const header = (rows[i + 1] || []).map((c) => String(c || "").trim().toLowerCase());
      const hasModule = header.some((t) => t === "module");
      const hasPassed = header.some((t) => t.startsWith("passed"));
      const hasFailed = header.some((t) => t.startsWith("failed"));
      if (!hasModule && hasPassed && hasFailed) workstream += 1;
    }
    if (moduleWise) return { kind: "module-wise", statusSections: moduleWise };
    if (workstream) return { kind: "workstream", statusSections: workstream };
    return null;
  } catch {
    return null;
  }
}

function pickField(row, names) {
  const want = names.map((n) => String(n).trim().toLowerCase());
  for (const key of Object.keys(row || {})) {
    if (want.includes(key.trim().toLowerCase())) return row[key];
  }
  return "";
}

function aggregateWorkstreamExport(filePath) {
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

  const counts = emptyCounts();
  for (const row of dataRows) {
    counts[normalizeStatus(row[8])] += 1;
  }
  const totalSummary = summarizeCounts(counts);
  return {
    exportRows: dataRows.length,
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

function aggregateWorkstreamDefects(filePath) {
  const wb = XLSX.readFile(filePath);
  if (!wb.SheetNames.length) {
    throw new Error("Unable to extract defect data because the export workbook has no sheets.");
  }
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
  const counts = emptyDefectCounts();
  const unknownStates = new Map();
  let mappedRows = 0;
  let skippedUnknownState = 0;

  for (const row of rows) {
    const stateRaw = pickField(row, ["state", "status"]);
    const bucket = mapDefectState(stateRaw);
    if (!bucket) {
      skippedUnknownState += 1;
      const label = String(stateRaw || "").trim() || "(blank)";
      unknownStates.set(label, (unknownStates.get(label) || 0) + 1);
      continue;
    }
    mappedRows += 1;
    counts[bucket] += 1;
  }

  const total =
    counts.closed + counts.deferred + counts.fixed + counts.pending;
  const denom = total - counts.deferred;
  return {
    counts: { ...counts, total },
    mappedRows,
    skippedUnknownState,
    exportRows: rows.length,
    unknownStates: [...unknownStates.entries()].map(([state, count]) => ({
      state,
      count,
    })),
    closureRate: denom > 0 ? counts.closed / denom : 0,
    resolutionRate: denom > 0 ? (counts.closed + counts.fixed) / denom : 0,
  };
}

function findTitleRow(sheet, title) {
  const normalized = String(title || "").trim().toLowerCase();
  for (let r = 1; r <= sheet.rowCount; r++) {
    for (let c = 1; c <= 12; c++) {
      const text = cellText(sheet.getRow(r).getCell(c).value).trim().toLowerCase();
      if (text === normalized) return r;
    }
    if (rowJoinedText(sheet, r).toLowerCase() === normalized) return r;
  }
  return -1;
}

function findWorkstreamTotalRow(sheet, dataRow, cols) {
  for (let r = dataRow + 1; r <= dataRow + 5; r++) {
    if (rowHasLabel(sheet, r, "execution rate") || rowHasLabel(sheet, r, "overall pass rate")) {
      return r - 1 > dataRow ? r - 1 : -1;
    }
    const champCol = cols.champion || cols.passed - 1;
    const label = champCol
      ? cellText(sheet.getRow(r).getCell(champCol).value).trim().toLowerCase()
      : "";
    if (label === "total") return r;
  }
  return dataRow + 1;
}

function writeWorkstreamSection(sheet, sectionTitle, summary) {
  const titleRow = findTitleRow(sheet, sectionTitle);
  if (titleRow < 0) {
    throw new Error(`Workstream section title not found in template: ${sectionTitle}`);
  }
  const layout = detectWorkstreamLayout(sheet, titleRow);
  if (!layout) {
    throw new Error(`Workstream status columns were not found under: ${sectionTitle}`);
  }
  const { cols, firstDataRow } = layout;
  const dataRowNum = firstDataRow;
  const totalRowNum = findWorkstreamTotalRow(sheet, dataRowNum, cols);
  const data = summary.totalSummary;
  const cPass = colLetter(cols.passed);
  const cFail = colLetter(cols.failed);
  const cBlock = colLetter(cols.blocked || cols.passed);
  const cNE = colLetter(cols.notExecuted || cols.total);
  const cTot = colLetter(cols.total);

  const dataRow = sheet.getRow(dataRowNum);
  setNumber(dataRow.getCell(cols.passed), data.passed);
  setNumber(dataRow.getCell(cols.failed), data.failed);
  if (cols.blocked) setNumber(dataRow.getCell(cols.blocked), data.blocked);
  if (cols.inProgress) setNumber(dataRow.getCell(cols.inProgress), data.inProgress);
  if (cols.notExecuted) setNumber(dataRow.getCell(cols.notExecuted), data.notExecuted);
  setFormula(
    dataRow.getCell(cols.total),
    cols.notExecuted
      ? `SUM(${cPass}${dataRowNum}:${cNE}${dataRowNum})`
      : `SUM(${cPass}${dataRowNum}:${cFail}${dataRowNum})`,
    data.total
  );
  if (cols.passRate) {
    const passCell = dataRow.getCell(cols.passRate);
    const formula = cols.blocked
      ? `${cPass}${dataRowNum}/(${cTot}${dataRowNum}-${cBlock}${dataRowNum})`
      : `${cPass}${dataRowNum}/${cTot}${dataRowNum}`;
    setFormula(passCell, formula, data.passRate, "0%");
    applyPassRateFill(passCell, data.passRate);
  }
  dataRow.commit();

  if (totalRowNum > dataRowNum) {
    const totalRow = sheet.getRow(totalRowNum);
    const champCol = cols.champion;
    if (champCol && !cellText(totalRow.getCell(champCol).value).trim()) {
      totalRow.getCell(champCol).value = "Total";
    }
    const countCols = [
      ["passed", cols.passed, data.passed],
      ["failed", cols.failed, data.failed],
      ["blocked", cols.blocked, data.blocked],
      ["inProgress", cols.inProgress, data.inProgress],
      ["notExecuted", cols.notExecuted, data.notExecuted],
    ];
    for (const [, col, value] of countCols) {
      if (!col) continue;
      const letter = colLetter(col);
      setFormula(
        totalRow.getCell(col),
        `SUM(${letter}${dataRowNum}:${letter}${dataRowNum})`,
        value
      );
    }
    setFormula(
      totalRow.getCell(cols.total),
      cols.notExecuted
        ? `SUM(${cPass}${totalRowNum}:${cNE}${totalRowNum})`
        : `SUM(${cPass}${totalRowNum}:${colLetter(cols.failed)}${totalRowNum})`,
      data.total
    );
    if (cols.passRate) {
      const formula = cols.blocked
        ? `${cPass}${totalRowNum}/(${cTot}${totalRowNum}-${cBlock}${totalRowNum})`
        : `${cPass}${totalRowNum}/${cTot}${totalRowNum}`;
      setFormula(totalRow.getCell(cols.passRate), formula, data.passRate, "0%");
      applyPassRateFill(totalRow.getCell(cols.passRate), data.passRate);
    }
    totalRow.commit();
  }

  const rateRowRef = totalRowNum > dataRowNum ? totalRowNum : dataRowNum;
  const execFormula = cols.blocked
    ? `(${cPass}${rateRowRef}+${cFail}${rateRowRef})/(${cTot}${rateRowRef}-${cBlock}${rateRowRef})`
    : `(${cPass}${rateRowRef}+${cFail}${rateRowRef})/${cTot}${rateRowRef}`;
  const overallFormula = cols.blocked
    ? `${cPass}${rateRowRef}/(${cTot}${rateRowRef}-${cBlock}${rateRowRef})`
    : `${cPass}${rateRowRef}/${cTot}${rateRowRef}`;
  const rateValueCol = cols.inProgress || cols.passed;

  for (let rr = dataRowNum + 1; rr <= dataRowNum + 8; rr++) {
    if (rowHasLabel(sheet, rr, "execution rate")) {
      setFormula(
        sheet.getRow(rr).getCell(rateValueCol),
        execFormula,
        summary.executionRate,
        "0%"
      );
      sheet.getRow(rr).commit();
    }
    if (rowHasLabel(sheet, rr, "overall pass rate")) {
      setFormula(
        sheet.getRow(rr).getCell(rateValueCol),
        overallFormula,
        summary.overallPassRate,
        "0%"
      );
      sheet.getRow(rr).commit();
    }
  }
}

function writeWorkstreamDefectSection(sheet, section, summary) {
  const titleRow = section.titleRow;
  const layout = section.layout || detectWorkstreamDefectLayout(sheet, titleRow);
  if (!layout) {
    throw new Error(`Workstream defect columns were not found under: ${section.title}`);
  }
  const { cols, firstDataRow } = layout;
  const r = firstDataRow;
  const row = sheet.getRow(r);
  const data = summary.counts;
  const cClosed = colLetter(cols.closed);
  const cPend = colLetter(cols.pending);
  const cTot = colLetter(cols.total);
  const cDef = colLetter(cols.deferred);
  const cFixed = colLetter(cols.fixed);

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

  for (let rr = firstDataRow + 1; rr <= firstDataRow + 5; rr++) {
    if (rowHasLabel(sheet, rr, "closure rate")) {
      setFormula(
        sheet.getRow(rr).getCell(cols.pending || cols.closed),
        `(${cDef}${r}+${cFixed}${r})/${cTot}${r}`,
        summary.counts.total
          ? (summary.counts.deferred + summary.counts.fixed) / summary.counts.total
          : 0,
        "0%"
      );
      sheet.getRow(rr).commit();
    }
    if (rowHasLabel(sheet, rr, "resolution rate")) {
      setFormula(
        sheet.getRow(rr).getCell(cols.pending || cols.closed),
        `(${cClosed}${r}+${cDef}${r}+${cFixed}${r})/${cTot}${r}`,
        summary.counts.total
          ? (summary.counts.closed + summary.counts.deferred + summary.counts.fixed) /
            summary.counts.total
          : 0,
        "0%"
      );
      sheet.getRow(rr).commit();
    }
  }
}

module.exports = {
  inspectTemplateFile,
  looksLikeWorkstreamFile,
  discoverWorkstreamStatusSections,
  discoverWorkstreamStatusInWorkbook,
  discoverWorkstreamDefectSections,
  aggregateWorkstreamExport,
  aggregateWorkstreamDefects,
  writeWorkstreamSection,
  writeWorkstreamDefectSection,
  isWorkstreamStatusTitle,
};
