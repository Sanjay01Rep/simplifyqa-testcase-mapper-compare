const ExcelJS = require("exceljs");

function cellText(value) {
  if (value == null) return "";
  if (typeof value === "object") {
    if (value.text) return String(value.text);
    if (value.richText) return value.richText.map((t) => t.text).join("");
    if (value.result != null) return String(value.result);
  }
  return String(value);
}

function cellNumber(value) {
  if (value == null || value === "") return 0;
  if (typeof value === "number") return value;
  if (typeof value === "object" && value.result != null) {
    return Number(value.result) || 0;
  }
  const n = Number(String(value).replace(/%/g, "").trim());
  return Number.isFinite(n) ? n : 0;
}

async function listWorkbookSheets(absPath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(absPath);
  return wb.worksheets.map((ws, idx) => ({
    name: ws.name,
    index: idx,
    id: ws.id,
  }));
}

/**
 * Extract status module rows keyed by section title.
 * Returns { sections: [{ title, modules: { name: {passed,failed,...} } }] }
 */
function extractStatusSections(sheet) {
  const sections = [];
  for (let r = 1; r <= sheet.rowCount; r++) {
    let title = "";
    for (let c = 1; c <= 4; c++) {
      const t = cellText(sheet.getRow(r).getCell(c).value).trim();
      if (/module-wise\s+daily\s+status/i.test(t)) {
        title = t;
        break;
      }
    }
    if (!title) continue;

    // Find header row
    let headerRow = -1;
    let cols = {};
    for (let hr = r + 1; hr <= r + 5; hr++) {
      const row = sheet.getRow(hr);
      const map = {};
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const t = cellText(cell.value).trim().toLowerCase();
        if (!t) return;
        if (t === "module" || t === "modules") map.module = colNumber;
        else if (t.startsWith("passed")) map.passed = colNumber;
        else if (t.startsWith("failed")) map.failed = colNumber;
        else if (t.startsWith("blocked")) map.blocked = colNumber;
        else if (t.includes("progress")) map.inProgress = colNumber;
        else if (t.includes("not executed")) map.notExecuted = colNumber;
        else if (t === "total") map.total = colNumber;
      });
      if (map.module && map.passed) {
        headerRow = hr;
        cols = map;
        break;
      }
    }
    if (headerRow < 0) continue;

    const modules = {};
    for (let dr = headerRow + 1; dr < headerRow + 80; dr++) {
      const name = cellText(sheet.getRow(dr).getCell(cols.module).value).trim();
      if (!name) continue;
      if (name.toLowerCase() === "total") break;
      modules[name] = {
        passed: cellNumber(sheet.getRow(dr).getCell(cols.passed).value),
        failed: cols.failed
          ? cellNumber(sheet.getRow(dr).getCell(cols.failed).value)
          : 0,
        blocked: cols.blocked
          ? cellNumber(sheet.getRow(dr).getCell(cols.blocked).value)
          : 0,
        inProgress: cols.inProgress
          ? cellNumber(sheet.getRow(dr).getCell(cols.inProgress).value)
          : 0,
        notExecuted: cols.notExecuted
          ? cellNumber(sheet.getRow(dr).getCell(cols.notExecuted).value)
          : 0,
        total: cols.total
          ? cellNumber(sheet.getRow(dr).getCell(cols.total).value)
          : 0,
      };
    }
    sections.push({ title, modules });
  }
  return sections;
}

function extractDefectSections(sheet) {
  const sections = [];
  for (let r = 1; r <= sheet.rowCount; r++) {
    let title = "";
    for (let c = 1; c <= 4; c++) {
      const t = cellText(sheet.getRow(r).getCell(c).value).trim();
      if (/defect\s+summary/i.test(t)) {
        title = t;
        break;
      }
    }
    if (!title) continue;

    let headerRow = -1;
    let cols = {};
    for (let hr = r + 1; hr <= r + 4; hr++) {
      const map = {};
      sheet.getRow(hr).eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const t = cellText(cell.value).trim().toLowerCase();
        if (!t) return;
        if (t === "module") map.module = colNumber;
        else if (t.startsWith("closed")) map.closed = colNumber;
        else if (t.startsWith("deferred")) map.deferred = colNumber;
        else if (t.startsWith("fixed")) map.fixed = colNumber;
        else if (t.startsWith("pending")) map.pending = colNumber;
        else if (t === "total") map.total = colNumber;
      });
      if (map.module && map.closed) {
        headerRow = hr;
        cols = map;
        break;
      }
    }
    if (headerRow < 0) continue;

    const modules = {};
    for (let dr = headerRow + 1; dr < headerRow + 80; dr++) {
      const name = cellText(sheet.getRow(dr).getCell(cols.module).value).trim();
      if (!name) continue;
      if (name.toLowerCase() === "total") break;
      modules[name] = {
        closed: cellNumber(sheet.getRow(dr).getCell(cols.closed).value),
        deferred: cols.deferred
          ? cellNumber(sheet.getRow(dr).getCell(cols.deferred).value)
          : 0,
        fixed: cols.fixed
          ? cellNumber(sheet.getRow(dr).getCell(cols.fixed).value)
          : 0,
        pending: cols.pending
          ? cellNumber(sheet.getRow(dr).getCell(cols.pending).value)
          : 0,
        total: cols.total
          ? cellNumber(sheet.getRow(dr).getCell(cols.total).value)
          : 0,
      };
    }
    sections.push({ title, modules });
  }
  return sections;
}

const FILL = {
  improved: "C6EFCE",
  worsened: "FFC7CE",
  changed: "FFE699",
};

function deltaFill(metric, prev, latest) {
  if (prev === latest) return null;
  const up = latest > prev;
  if (metric === "passed" || metric === "closed" || metric === "fixed") {
    return up ? FILL.improved : FILL.worsened;
  }
  if (
    metric === "failed" ||
    metric === "blocked" ||
    metric === "pending" ||
    metric === "notExecuted"
  ) {
    return up ? FILL.worsened : FILL.improved;
  }
  return FILL.changed;
}

function applyFill(cell, hex) {
  if (!hex) return;
  cell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: `FF${hex}` },
  };
}

/**
 * Write a compare sheet into the workbook for previousSheet vs latestSheet.
 */
async function writeCompareSheet(absPath, previousName, latestName) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(absPath);
  const prev = wb.worksheets.find((s) => s.name === previousName);
  const latest = wb.worksheets.find((s) => s.name === latestName);
  if (!prev || !latest) {
    throw new Error("One or both selected sheets were not found in the workbook.");
  }

  const now = new Date();
  const stamp = now.toLocaleString();
  const sheetNameBase = `Compare ${now
    .toISOString()
    .slice(0, 16)
    .replace("T", " ")
    .replace(/:/g, "-")}`;
  let sheetName = sheetNameBase.slice(0, 31);
  let i = 2;
  while (wb.worksheets.some((s) => s.name === sheetName)) {
    const suffix = `-${i}`;
    sheetName = (sheetNameBase.slice(0, 31 - suffix.length) + suffix).slice(0, 31);
    i += 1;
  }

  const out = wb.addWorksheet(sheetName);
  let row = 1;
  out.getCell(row, 1).value = "Report Compare";
  out.getCell(row, 1).font = { bold: true, size: 14 };
  row += 1;
  out.getCell(row, 1).value = `Generated: ${stamp}`;
  row += 1;
  out.getCell(row, 1).value = `Previous: ${previousName}`;
  row += 1;
  out.getCell(row, 1).value = `Latest: ${latestName}`;
  row += 1;
  out.getCell(row, 1).value =
    "Highlight: green = improved, red = worsened, amber = other change. Values show Latest (Δ vs Previous).";
  row += 2;

  const prevStatus = extractStatusSections(prev);
  const latestStatus = extractStatusSections(latest);
  const statusTitles = [
    ...new Set([
      ...prevStatus.map((s) => s.title),
      ...latestStatus.map((s) => s.title),
    ]),
  ];

  const statusMetrics = [
    "passed",
    "failed",
    "blocked",
    "inProgress",
    "notExecuted",
    "total",
  ];

  for (const title of statusTitles) {
    out.getCell(row, 1).value = title;
    out.getCell(row, 1).font = { bold: true };
    row += 1;
    const headers = ["Module", ...statusMetrics.map((m) => `${m} (latest / Δ)`)];
    headers.forEach((h, idx) => {
      const cell = out.getCell(row, idx + 1);
      cell.value = h;
      cell.font = { bold: true };
    });
    row += 1;

    const a = prevStatus.find((s) => s.title === title)?.modules || {};
    const b = latestStatus.find((s) => s.title === title)?.modules || {};
    const names = [...new Set([...Object.keys(a), ...Object.keys(b)])].sort(
      (x, y) => x.localeCompare(y)
    );
    for (const name of names) {
      out.getCell(row, 1).value = name;
      statusMetrics.forEach((metric, idx) => {
        const prevV = (a[name] && a[name][metric]) || 0;
        const latestV = (b[name] && b[name][metric]) || 0;
        const delta = latestV - prevV;
        const cell = out.getCell(row, idx + 2);
        cell.value =
          delta === 0 ? latestV : `${latestV} (${delta > 0 ? "+" : ""}${delta})`;
        applyFill(cell, deltaFill(metric, prevV, latestV));
      });
      row += 1;
    }
    row += 1;
  }

  const prevDef = extractDefectSections(prev);
  const latestDef = extractDefectSections(latest);
  const defTitles = [
    ...new Set([
      ...prevDef.map((s) => s.title),
      ...latestDef.map((s) => s.title),
    ]),
  ];
  const defMetrics = ["closed", "deferred", "fixed", "pending", "total"];

  for (const title of defTitles) {
    out.getCell(row, 1).value = title;
    out.getCell(row, 1).font = { bold: true };
    row += 1;
    const headers = ["Module", ...defMetrics.map((m) => `${m} (latest / Δ)`)];
    headers.forEach((h, idx) => {
      const cell = out.getCell(row, idx + 1);
      cell.value = h;
      cell.font = { bold: true };
    });
    row += 1;

    const a = prevDef.find((s) => s.title === title)?.modules || {};
    const b = latestDef.find((s) => s.title === title)?.modules || {};
    const names = [...new Set([...Object.keys(a), ...Object.keys(b)])].sort(
      (x, y) => x.localeCompare(y)
    );
    for (const name of names) {
      out.getCell(row, 1).value = name;
      defMetrics.forEach((metric, idx) => {
        const prevV = (a[name] && a[name][metric]) || 0;
        const latestV = (b[name] && b[name][metric]) || 0;
        const delta = latestV - prevV;
        const cell = out.getCell(row, idx + 2);
        cell.value =
          delta === 0 ? latestV : `${latestV} (${delta > 0 ? "+" : ""}${delta})`;
        applyFill(cell, deltaFill(metric, prevV, latestV));
      });
      row += 1;
    }
    row += 1;
  }

  out.getColumn(1).width = 28;
  for (let c = 2; c <= 8; c++) out.getColumn(c).width = 16;

  // Activate compare sheet
  const idx = wb.worksheets.findIndex((ws) => ws.id === out.id);
  wb.views = [
    {
      x: 0,
      y: 0,
      width: 25000,
      height: 15000,
      firstSheet: idx,
      activeTab: idx,
      visibility: "visible",
    },
  ];

  await wb.xlsx.writeFile(absPath);
  return { sheetName, generatedAt: stamp, previousName, latestName };
}

module.exports = {
  listWorkbookSheets,
  extractStatusSections,
  extractDefectSections,
  writeCompareSheet,
};
