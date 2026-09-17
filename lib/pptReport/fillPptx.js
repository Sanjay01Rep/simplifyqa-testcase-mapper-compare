const fs = require("fs");
const path = require("path");
const JSZip = require("jszip");
const { formatPct, pctInt, splitCoverTitle } = require("./parseDsr");

const PASS_BANDS = {
  red: { fill: "E53935", text: "000000" },
  amber: { fill: "FFC107", text: "000000" },
  green: { fill: "8BC34A", text: "000000" },
};
const PASS_TEXT = "000000";

function passBand(rate) {
  const pct = pctInt(rate);
  if (pct <= 50) return PASS_BANDS.red;
  if (pct <= 80) return PASS_BANDS.amber;
  return PASS_BANDS.green;
}

function escapeXml(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function splitTables(xml) {
  const out = [];
  const re = /<a:tbl\b[\s\S]*?<\/a:tbl>/g;
  let m;
  while ((m = re.exec(xml))) {
    out.push({ xml: m[0], start: m.index, end: m.index + m[0].length });
  }
  return out;
}

function tableRows(tbl) {
  const out = [];
  const re = /<a:tr\b[\s\S]*?<\/a:tr>/g;
  let m;
  while ((m = re.exec(tbl))) {
    out.push({ xml: m[0], start: m.index, end: m.index + m[0].length });
  }
  return out;
}

function rowCells(tr) {
  const out = [];
  const re = /<a:tc\b[\s\S]*?<\/a:tc>/g;
  let m;
  while ((m = re.exec(tr))) out.push(m[0]);
  return out;
}

function setRowHeight(trXml, h) {
  if (!Number.isFinite(h)) return trXml;
  if (/<a:tr\b[^>]*\bh="/.test(trXml)) {
    return trXml.replace(/(<a:tr\b[^>]*\bh=")(\d+)/, `$1${Math.round(h)}`);
  }
  return trXml.replace(/<a:tr\b/, `<a:tr h="${Math.round(h)}"`);
}

function rowHeight(trXml) {
  const m = trXml.match(/<a:tr\b[^>]*\bh="(\d+)"/);
  return m ? Number(m[1]) : 219456;
}

function setTcText(tcXml, text) {
  let first = true;
  return tcXml.replace(/<a:t(?: [^>]*)?>[\s\S]*?<\/a:t>/g, (full) => {
    const open = full.match(/^<a:t(?: [^>]*)?>/)[0];
    if (first) {
      first = false;
      return `${open}${escapeXml(text)}</a:t>`;
    }
    return `${open}</a:t>`;
  });
}

function setTcPrCellFill(tcPr, hex) {
  const fillXml = `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
  const open = tcPr.match(/^<a:tcPr\b[^>]*>/);
  if (!open) return tcPr;
  const inner = tcPr.slice(open[0].length).replace(/<\/a:tcPr>\s*$/, "");
  const cleaned = inner.replace(
    /<(a:ln[LRTB])\b[\s\S]*?<\/\1>|<a:solidFill>[\s\S]*?<\/a:solidFill>|<a:noFill\s*\/>|<a:noFill>[\s\S]*?<\/a:noFill>/g,
    (m) => (/^<a:ln/.test(m) ? m : "")
  );
  return `${open[0]}${cleaned}${fillXml}</a:tcPr>`;
}

function setRunTextBlack(xml) {
  let out = xml.replace(/<a:highlight>[\s\S]*?<\/a:highlight>/g, "");
  out = out.replace(/<a:rPr\b[\s\S]*?<\/a:rPr>/g, (rpr) => {
    let next = rpr.replace(/<a:highlight>[\s\S]*?<\/a:highlight>/g, "");
    if (/<a:solidFill>[\s\S]*?<\/a:solidFill>/.test(next)) {
      return next.replace(
        /<a:solidFill>[\s\S]*?<\/a:solidFill>/,
        `<a:solidFill><a:srgbClr val="${PASS_TEXT}"/></a:solidFill>`
      );
    }
    return next.replace(
      /<a:rPr([^>/]*)(\/?)>/,
      (mm, attrs, slash) => {
        const fill = `<a:solidFill><a:srgbClr val="${PASS_TEXT}"/></a:solidFill>`;
        if (slash) return `<a:rPr${attrs}>${fill}</a:rPr>`;
        return `<a:rPr${attrs}>${fill}`;
      }
    );
  });
  return out;
}

function setPassRateCell(tcXml, rate) {
  const band = passBand(rate);
  let out = setTcText(tcXml, formatPct(rate));
  if (/<a:tcPr\b[\s\S]*?<\/a:tcPr>/.test(out)) {
    out = out.replace(/<a:tcPr\b[\s\S]*?<\/a:tcPr>/, (tcPr) => setTcPrCellFill(tcPr, band.fill));
  } else if (/<a:tcPr\b[^>]*\/>/.test(out)) {
    out = out.replace(
      /<a:tcPr([^>]*)\/>/,
      `<a:tcPr$1><a:solidFill><a:srgbClr val="${band.fill}"/></a:solidFill></a:tcPr>`
    );
  }
  return setRunTextBlack(out);
}

function chartRatesFromParsed(parsed) {
  const order = parsed.statusOrder || [];
  return order.map((entity) => {
    const section = parsed.statusByEntity && parsed.statusByEntity[entity];
    return {
      exec: section ? pctInt(section.executionRate) : 0,
      pass: section ? pctInt(section.overallPassRate) : 0,
    };
  });
}

function fillNumCache(cacheXml, values) {
  return cacheXml.replace(
    /<c:pt idx="(\d+)">\s*<c:v>[\s\S]*?<\/c:v>\s*<\/c:pt>/g,
    (pt, idx) => {
      const v = values[Number(idx)];
      if (v == null || !Number.isFinite(v)) return pt;
      return `<c:pt idx="${idx}"><c:v>${v}</c:v></c:pt>`;
    }
  );
}

function fillExecPassChart(xml, rates) {
  const exec = rates.map((r) => r.exec);
  const pass = rates.map((r) => r.pass);
  let series = 0;
  let out = xml.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, (ser) => {
    const values = series === 0 ? exec : pass;
    series += 1;
    let next = ser.replace(/<c:dLbl>[\s\S]*?<\/c:dLbl>/g, "");
    next = next.replace(/<c:numCache>[\s\S]*?<\/c:numCache>/, (cache) =>
      fillNumCache(cache, values)
    );
    return next;
  });
  const peak = Math.max(0, ...exec, ...pass);
  const axisMax = Math.max(50, Math.ceil(peak / 10) * 10);
  out = out.replace(
    /(<c:valAx>[\s\S]*?<c:scaling>[\s\S]*?<c:max val=")[^"]+/,
    `$1${axisMax}`
  );
  return out;
}

async function fillChartWorkbook(zip, rates) {
  const embedPath = "ppt/embeddings/Workbook1.xlsx";
  const file = zip.file(embedPath);
  if (!file) return false;
  const inner = await JSZip.loadAsync(await file.async("nodebuffer"));
  const sheetPath = "xl/worksheets/sheet1.xml";
  const sheetFile = inner.file(sheetPath);
  if (!sheetFile) return false;
  let sheet = await sheetFile.async("string");
  const pairs = [
    ["B2", rates[0] && rates[0].exec],
    ["C2", rates[0] && rates[0].pass],
    ["B3", rates[1] && rates[1].exec],
    ["C3", rates[1] && rates[1].pass],
    ["B4", rates[2] && rates[2].exec],
    ["C4", rates[2] && rates[2].pass],
  ];
  for (const [ref, val] of pairs) {
    if (val == null || !Number.isFinite(val)) continue;
    const re = new RegExp(`(<c r="${ref}"[^>]*>\\s*<v>)[^<]+`);
    if (re.test(sheet)) {
      sheet = sheet.replace(re, `$1${val}`);
    }
  }
  inner.file(sheetPath, sheet);
  zip.file(
    embedPath,
    await inner.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })
  );
  return true;
}

function rebuildTable(tblXml, dataRows, totalRow, { passRateCol, warnings, label }) {
  const rows = tableRows(tblXml);
  if (rows.length < 2) return tblXml;
  const header = rows[0].xml;
  const templateData = rows[1].xml;
  const templateTotal = rows[rows.length - 1].xml;
  const origH = rows.reduce((s, r) => s + rowHeight(r.xml), 0);
  const origModules = Math.max(0, rows.length - 2);
  const nData = Math.max(dataRows.length, 1);
  const nRows = 1 + nData + 1;
  let rowH = rowHeight(templateData);
  const natural = nRows * rowH;
  if (nData > origModules && natural > origH) {
    const next = Math.max(130000, Math.floor(origH / nRows));
    if (next < rowH) {
      rowH = next;
      warnings.push(
        `${label}: added ${nData - origModules} module row(s) and reduced row height so columns/footer do not overlap.`
      );
    }
  }

  const built = [setRowHeight(header, rowH)];
  for (const data of dataRows) {
    built.push(fillStatusOrDefectRow(templateData, data, { passRateCol, rowH }));
  }
  built.push(
    fillStatusOrDefectRow(templateTotal, totalRow, { passRateCol, rowH })
  );

  const innerStart = rows[0].start;
  const innerEnd = rows[rows.length - 1].end;
  return tblXml.slice(0, innerStart) + built.join("") + tblXml.slice(innerEnd);
}

function fillStatusOrDefectRow(trXml, data, { passRateCol, rowH }) {
  const cells = rowCells(trXml);
  const values = data.values || [];
  const nextCells = cells.map((tc, i) => {
    const text = values[i] != null ? String(values[i]) : "";
    if (passRateCol && i === passRateCol && data.passRate != null) {
      return setPassRateCell(tc, data.passRate);
    }
    return setTcText(tc, text);
  });
  let out = trXml;
  // Replace cells from the end so indices stay valid if we used string replace by content.
  const found = [...trXml.matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)];
  if (found.length !== nextCells.length) return setRowHeight(trXml, rowH);
  let rebuilt = trXml.slice(0, found[0].index);
  for (let i = 0; i < found.length; i++) {
    rebuilt += nextCells[i];
  }
  rebuilt += trXml.slice(found[found.length - 1].index + found[found.length - 1][0].length);
  return setRowHeight(rebuilt, rowH);
}

function statusRowPayload(mod) {
  return {
    passRate: mod.passRate,
    values: [
      mod.name,
      String(mod.passed ?? 0),
      String(mod.failed ?? 0),
      String(mod.blocked ?? 0),
      String(mod.inProgress ?? 0),
      String(mod.notExecuted ?? 0),
      String(mod.total ?? 0),
      formatPct(mod.passRate),
    ],
  };
}

function defectRowPayload(mod) {
  return {
    values: [
      mod.name,
      String(mod.closed ?? 0),
      String(mod.deferred ?? 0),
      String(mod.fixed ?? 0),
      String(mod.pending ?? 0),
      String(mod.total ?? 0),
    ],
  };
}

function replaceExactRun(xml, exact, next) {
  const needle = `<a:t>${escapeXml(exact)}</a:t>`;
  const idx = xml.indexOf(needle);
  if (idx < 0) {
    const loose = xml.replace(
      new RegExp(`<a:t>${escapeXml(exact)}</a:t>`),
      `<a:t>${escapeXml(next)}</a:t>`
    );
    return loose;
  }
  return xml.slice(0, idx) + `<a:t>${escapeXml(next)}</a:t>` + xml.slice(idx + needle.length);
}

function fillCoverSlide(xml, { headline, line1, line2, dateText }) {
  let out = xml;
  out = replaceExactRun(out, "WEEKLY WORKSTREAM MEETING", headline);
  out = replaceExactRun(out, "Weekly Workstream", line1);
  out = replaceExactRun(out, "Meeting", line2);
  out = out.replace(
    /<a:t>16 Sept 2026<\/a:t>/,
    `<a:t>${escapeXml(dateText)}</a:t>`
  );
  return out;
}

function fillRatePairs(xml, pairs, execLabel, passLabel) {
  const parts = xml.split(/(<a:t(?: [^>]*)?>[\s\S]*?<\/a:t>)/);
  let mode = null;
  let idx = -1;
  let filled = false;
  return parts
    .map((part) => {
      const m = part.match(/^<a:t([^>]*)>([\s\S]*)<\/a:t>$/);
      if (!m) return part;
      const attrs = m[1];
      const text = m[2]
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">");
      if (execLabel.test(text)) {
        idx += 1;
        mode = "exec";
        filled = false;
        return part;
      }
      if (passLabel.test(text)) {
        mode = "pass";
        filled = false;
        return part;
      }
      const pair = idx >= 0 ? pairs[idx] : null;
      if (!pair) return part;
      if (mode === "exec") {
        if (/^MODULE$/i.test(text.trim())) {
          mode = null;
          return part;
        }
        if (!filled) {
          filled = true;
          return `<a:t${attrs}>${escapeXml(pair.exec)}</a:t>`;
        }
        if (/^[\d.%\s]*$/.test(text)) return `<a:t${attrs}></a:t>`;
        mode = null;
        return part;
      }
      if (mode === "pass") {
        if (/^MODULE$/i.test(text.trim()) || /^Module$/i.test(text.trim())) {
          mode = null;
          return part;
        }
        if (!filled) {
          filled = true;
          return `<a:t${attrs}>${escapeXml(pair.pass)}</a:t>`;
        }
        if (/^[\d.%\s]*$/.test(text)) return `<a:t${attrs}></a:t>`;
        mode = null;
      }
      return part;
    })
    .join("");
}

function extraModuleNames(section, originalNames) {
  const orig = new Set((originalNames || []).map((n) => n.toLowerCase()));
  return (section.modules || [])
    .map((m) => m.name)
    .filter((n) => n && !orig.has(n.toLowerCase()));
}

function originalModuleNames(tblXml) {
  const rows = tableRows(tblXml);
  const names = [];
  for (let i = 1; i < rows.length - 1; i++) {
    const cells = rowCells(rows[i].xml);
    const t = (cells[0].match(/<a:t(?: [^>]*)?>([\s\S]*?)<\/a:t>/) || [])[1] || "";
    names.push(
      t.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").trim()
    );
  }
  return names;
}

function fillStatusSlide(xml, parsed, warnings) {
  const tables = splitTables(xml);
  const order = parsed.statusOrder || [];
  let out = xml;
  // Work back-to-front so indexes stay valid.
  for (let i = tables.length - 1; i >= 0; i--) {
    const entity = order[i];
    const section = (parsed.statusByEntity && parsed.statusByEntity[entity]) || parsed.status[i];
    if (!section) {
      warnings.push(`Slide 3 table ${i + 1}: no Excel status block for ${entity || "this section"}; left as in the template.`);
      continue;
    }
    const origNames = originalModuleNames(tables[i].xml);
    const extras = extraModuleNames(section, origNames);
    if (extras.length) {
      warnings.push(
        `${entity}: added extra module row(s) from Excel: ${extras.join(", ")}.`
      );
    }
    const dataRows = (section.modules || []).map(statusRowPayload);
    const totalRow = statusRowPayload(section.totals || { name: "Total", passRate: 0 });
    const nextTbl = rebuildTable(tables[i].xml, dataRows, totalRow, {
      passRateCol: 7,
      warnings,
      label: entity || `Status table ${i + 1}`,
    });
    out = out.slice(0, tables[i].start) + nextTbl + out.slice(tables[i].end);
  }

  const pairs = order.map((entity) => {
    const section = parsed.statusByEntity && parsed.statusByEntity[entity];
    return {
      exec: formatPct(section ? section.executionRate : 0),
      pass: formatPct(section ? section.overallPassRate : 0),
    };
  });
  out = fillRatePairs(out, pairs, /^Exec\.\s*$/, /Pass\s*$/);
  return out;
}

function fillDefectSlide(xml, parsed, warnings) {
  const tables = splitTables(xml);
  const order = parsed.defectOrder || [];
  let out = xml;
  for (let i = tables.length - 1; i >= 0; i--) {
    const entity = order[i];
    const section = (parsed.defectsByEntity && parsed.defectsByEntity[entity]) || parsed.defects[i];
    if (!section) {
      warnings.push(`Slide 4 table ${i + 1}: no Excel defect block for ${entity || "this section"}; left as in the template.`);
      continue;
    }
    const origNames = originalModuleNames(tables[i].xml);
    const extras = extraModuleNames(section, origNames);
    if (extras.length) {
      warnings.push(
        `${entity} defects: added extra module row(s) from Excel: ${extras.join(", ")}.`
      );
    }
    const dataRows = (section.modules || []).map(defectRowPayload);
    const totalRow = defectRowPayload(
      section.totals || { name: "Total", closed: 0, deferred: 0, fixed: 0, pending: 0, total: 0 }
    );
    const nextTbl = rebuildTable(tables[i].xml, dataRows, totalRow, {
      passRateCol: null,
      warnings,
      label: `${entity || "Defects"} defects`,
    });
    out = out.slice(0, tables[i].start) + nextTbl + out.slice(tables[i].end);
  }
  const pairs = order.map((entity) => {
    const section = parsed.defectsByEntity && parsed.defectsByEntity[entity];
    return {
      exec: formatPct(section ? section.closureRate : 0),
      pass: formatPct(section ? section.resolutionRate : 0),
    };
  });
  out = fillRatePairs(
    out,
    pairs,
    /Closure Rate/,
    /Resolution Rate/
  );
  return out;
}

async function fillWeeklyPptx({
  templatePath,
  outputPath,
  parsed,
  title,
  dateText,
}) {
  if (!fs.existsSync(templatePath)) {
    throw new Error(`PPT template was not found: ${templatePath}`);
  }
  const buf = fs.readFileSync(templatePath);
  const zip = await JSZip.loadAsync(buf);
  const warnings = [];
  const cover = splitCoverTitle(title);

  const s1 = "ppt/slides/slide1.xml";
  const s3 = "ppt/slides/slide3.xml";
  const s4 = "ppt/slides/slide4.xml";
  if (!zip.file(s1) || !zip.file(s3) || !zip.file(s4)) {
    throw new Error("PPT template is missing expected slides (1, 3, 4).");
  }

  const slide1 = await zip.file(s1).async("string");
  zip.file(s1, fillCoverSlide(slide1, { ...cover, dateText }));

  const slide3 = await zip.file(s3).async("string");
  zip.file(s3, fillStatusSlide(slide3, parsed, warnings));

  const slide4 = await zip.file(s4).async("string");
  zip.file(s4, fillDefectSlide(slide4, parsed, warnings));

  const rates = chartRatesFromParsed(parsed);
  const chartPath = "ppt/charts/chart1.xml";
  if (zip.file(chartPath)) {
    const chartXml = await zip.file(chartPath).async("string");
    zip.file(chartPath, fillExecPassChart(chartXml, rates));
    const embedded = await fillChartWorkbook(zip, rates);
    if (!embedded) {
      warnings.push(
        "Slide 3 chart values were written to the chart cache, but the embedded workbook was not found."
      );
    }
  } else {
    warnings.push(
      "Slide 3 chart (Execution vs Overall Pass Rate) was not found in the template; tables were still updated."
    );
  }

  const outBuf = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, outBuf);
  return { warnings, cover };
}

module.exports = {
  fillWeeklyPptx,
  passBand,
  PASS_BANDS,
  setPassRateCell,
  fillExecPassChart,
  chartRatesFromParsed,
};
