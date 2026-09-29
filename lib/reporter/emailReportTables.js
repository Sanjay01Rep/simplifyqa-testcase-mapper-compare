/**
 * Builds Entity Summary, module-wise Daily Status, and Defect tables
 * from a generated SIT sheet using the Outlook HTML from the sample .msg
 * (Excel paste: fixed widths, header fills, rowspan champions).
 */

const ExcelJS = require("exceljs");
const entitySummary = require("./entitySummary");

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
  if (typeof value === "object" && typeof value.result === "number") return value.result;
  const n = Number(String(value).replace(/%/g, "").trim());
  return Number.isFinite(n) ? n : 0;
}

function percentText(value) {
  const n = numberOf(value);
  const frac = Math.abs(n) > 1 ? n / 100 : n;
  return `${Math.round(frac * 100)}%`;
}

function escapeHtml(value) {
  if (value == null) return "";
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function pickSheet(workbook, sheetName) {
  const want = String(sheetName || "").trim();
  if (want) {
    const named = workbook.worksheets.find((ws) => String(ws.name) === want);
    if (named) return named;
  }
  for (let i = workbook.worksheets.length - 1; i >= 0; i--) {
    if (entitySummary.isDatedRunSheet(workbook.worksheets[i])) return workbook.worksheets[i];
  }
  return workbook.worksheets[workbook.worksheets.length - 1] || null;
}

function para(htmlInner) {
  return (
    `<div style="font-family: Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">${htmlInner}</div>` +
    `<div style="font-family: Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);"><br></div>`
  );
}

const ENTITY = {
  tableWidth: "819.24px",
  widths: [
    "71.175px",
    "66.3125px",
    "68.8375px",
    "67.6625px",
    "82.0375px",
    "109.425px",
    "97.2125px",
    "94.7px",
    "63.925px",
    "97.15px",
  ],
};

const DAILY = {
  tableWidth: "614.42pt",
  titleWidth: "614.55pt",
  widths: [
    "137.6px",
    "200.8px",
    "45.6pt",
    "53.4pt",
    "43.8pt",
    "59.4pt",
    "71.4pt",
    "39.15pt",
    "64px",
  ],
  headerBg: [
    "rgb(174, 170, 170)",
    "rgb(255, 255, 255)",
    "rgb(0, 176, 80)",
    "rgb(255, 0, 0)",
    "rgb(244, 176, 132)",
    "rgb(173, 173, 173)",
    "rgb(255, 217, 102)",
    "rgb(255, 255, 255)",
    "rgb(146, 208, 80)",
  ],
  totalBg: [
    "rgb(255, 255, 255)",
    null,
    "rgb(0, 176, 80)",
    "rgb(255, 0, 0)",
    "rgb(244, 176, 132)",
    "rgb(173, 173, 173)",
    "rgb(255, 217, 102)",
    "rgb(91, 155, 213)",
    null,
  ],
};

const DEFECT = {
  tableWidth: "424.23pt",
  titleWidth: "424.2pt",
  widths: ["200.8px", "60.8px", "71.2px", "58.4px", "79.2px", "95.2px"],
  headerBg: [
    "rgb(255, 255, 255)",
    "rgb(112, 173, 71)",
    "rgb(244, 176, 132)",
    "rgb(255, 255, 0)",
    "rgb(255, 0, 0)",
    "rgb(255, 255, 255)",
  ],
};

const FILL = {
  headerGrey: "rgb(174, 170, 170)",
  passed: "rgb(226, 240, 217)",
  failed: "rgb(251, 229, 229)",
  blocked: "rgb(255, 242, 204)",
  module: "rgb(202, 237, 251)",
  dailyTitle: "rgb(72, 116, 203)",
  defectTitle: "rgb(255, 230, 153)",
  execRate: "rgb(255, 192, 0)",
  overallPass: "rgb(112, 173, 71)",
  closure: "rgb(112, 173, 71)",
  resolution: "rgb(248, 203, 173)",
  white: "rgb(255, 255, 255)",
};

function statusFill(label) {
  const key = String(label || "").toLowerCase();
  if (key === "red") return "rgb(255, 0, 0)";
  if (key === "amber") return "rgb(255, 192, 0)";
  if (key === "green") return "rgb(146, 208, 80)";
  return FILL.white;
}

function passBandFill(pctText) {
  const n = parseInt(String(pctText || "").replace(/%/g, ""), 10);
  if (!Number.isFinite(n)) return null;
  return statusFill(entitySummary.classifyStatus(n / 100).label);
}

function deltaColor(text) {
  const t = String(text || "");
  if (/^was\s/i.test(t.trim())) return "rgb(136, 136, 136)";
  if (/▼/.test(t)) return "rgb(198, 40, 40)";
  if (/▲/.test(t)) return "rgb(46, 125, 50)";
  return "rgb(34, 34, 34)";
}

function borderCss(kind) {
  if (kind === "none") {
    return "border-width: medium; border-style: none; border-color: currentcolor;";
  }
  if (kind === "headerBar") {
    return "border-width: medium 0.5pt 0.5pt; border-style: none solid solid; border-color: currentcolor rgb(0, 0, 0) rgb(0, 0, 0);";
  }
  if (kind === "titleBar") {
    return "border-width: medium medium medium 0.5pt; border-style: none none none solid; border-color: currentcolor currentcolor currentcolor rgb(0, 0, 0);";
  }
  return "border-width: 0.5pt; border-style: solid; border-color: rgb(0, 0, 0);";
}

function outlookTd(text, opts = {}) {
  const rowspan = opts.rowspan && opts.rowspan > 1 ? ` rowspan="${opts.rowspan}"` : "";
  const colspan = opts.colspan && opts.colspan > 1 ? ` colspan="${opts.colspan}"` : "";
  const vAlign = opts.vAlign || "middle";
  const color = opts.color || "rgb(0, 0, 0)";
  const nowrap = opts.nowrap ? " white-space: nowrap;" : "";
  const width = opts.width ? ` width: ${opts.width};` : "";
  const height = opts.height ? ` height: ${opts.height};` : "";
  const bg = opts.bg != null ? ` background-color: ${opts.bg};` : "";
  const override = opts.headerOverride
    ? ` data-editing-info="{&quot;bgColorOverride&quot;:true}"`
    : "";
  const emptyOpen = opts.border === "none" && (text == null || text === "");
  const align = emptyOpen ? "" : " text-align: center;";
  const inner = emptyOpen
    ? ""
    : opts.bold
      ? `<span style="font-weight: 700;">${escapeHtml(text)}</span>`
      : escapeHtml(text);
  const innerHtml = emptyOpen
    ? ""
    : `<div style="direction: ltr; text-align: center;${nowrap} font-family: Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">${inner}</div>`;
  return (
    `<td${rowspan}${colspan} style="direction: ltr;${align}${nowrap} ${borderCss(
      opts.border
    )}${bg} padding-top: 1px; padding-right: 1px; padding-left: 1px; vertical-align: ${vAlign}; color: ${color};${width}${height} box-sizing: border-box;"${override}>` +
    innerHtml +
    `</td>`
  );
}

function outlookTable(width, innerHtml) {
  return (
    `<table style="direction: ltr; width: ${width}; box-sizing: border-box; border-collapse: collapse; border-spacing: 0px;" cellspacing="0" cellpadding="0" border="0">` +
    `<tbody>${innerHtml}</tbody></table>` +
    `<div style="font-family: Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);"><br></div>`
  );
}

function textTable(headers, rows) {
  const lines = [headers.join("\t")];
  for (const row of rows) {
    lines.push(
      row
        .map((cell) => {
          const c = cell && typeof cell === "object" ? cell : { text: cell };
          return c.text == null ? "" : String(c.text);
        })
        .join("\t")
    );
  }
  return lines.join("\n");
}

function parseEntitySummary(sheet) {
  const start = entitySummary.detectEntityStartCol
    ? entitySummary.detectEntityStartCol(sheet)
    : entitySummary.COL.entity;
  const col0 = start || entitySummary.COL.entity;
  const d = col0 - entitySummary.COL.entity;
  const col = (key) => entitySummary.COL[key] + d;
  const headers = [
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
  const w = ENTITY.widths;
  const textRows = [];
  const bodyParts = [];
  const headerRow = `<tr>${headers
    .map((h, i) =>
      outlookTd(h, {
        bg: FILL.headerGrey,
        bold: true,
        width: w[i],
        height: i === 0 ? "15.6pt" : "15px",
        headerOverride: true,
      })
    )
    .join("")}</tr>`;
  for (const pos of entitySummary.ROW_POSITIONS) {
    const main = sheet.getRow(pos.mainRow);
    const delta = sheet.getRow(pos.deltaRow);
    const label = cellText(main.getCell(col("entity")).value).trim() || pos.label;
    const total = numberOf(main.getCell(col("total")).value);
    const passed = numberOf(main.getCell(col("passed")).value);
    if (!label || (total === 0 && passed === 0)) continue;
    const passFrac =
      Math.abs(numberOf(main.getCell(col("pass")).value)) > 1
        ? numberOf(main.getCell(col("pass")).value) / 100
        : numberOf(main.getCell(col("pass")).value);
    const status =
      cellText(main.getCell(col("status")).value).trim() ||
      entitySummary.classifyStatus(passFrac).label;
    const failed = numberOf(main.getCell(col("failed")).value);
    const blocked = numberOf(main.getCell(col("blocked")).value);
    const inProgress = numberOf(main.getCell(col("inProgress")).value);
    const notExecuted = numberOf(main.getCell(col("notExecuted")).value);
    const execPct = percentText(main.getCell(col("execution")).value);
    const passPct = percentText(main.getCell(col("pass")).value);
    const dPassed = cellText(delta.getCell(col("passed")).value).trim();
    const dFailed = cellText(delta.getCell(col("failed")).value).trim();
    const dBlocked = cellText(delta.getCell(col("blocked")).value).trim();
    const dInProg = cellText(delta.getCell(col("inProgress")).value).trim();
    const dExec = cellText(delta.getCell(col("execution")).value).trim();
    const dPass = cellText(delta.getCell(col("pass")).value).trim();
    const hasDelta = Boolean(dPassed || dFailed || dBlocked || dInProg || dExec || dPass);
    const span = hasDelta ? 2 : 1;
    const bodyColor = "rgb(34, 34, 34)";
    const mainTr = `<tr>${[
      outlookTd(label, {
        rowspan: span,
        bold: true,
        bg: FILL.white,
        color: bodyColor,
        width: w[0],
        height: span === 2 ? "31.2pt" : "15.6pt",
      }),
      outlookTd(total, {
        rowspan: span,
        bg: FILL.white,
        color: bodyColor,
        width: w[1],
        height: span === 2 ? "30px" : "15px",
      }),
      outlookTd(passed, {
        bg: FILL.passed,
        vAlign: "bottom",
        color: bodyColor,
        width: w[2],
        height: "15px",
      }),
      outlookTd(failed, {
        bg: FILL.failed,
        vAlign: "bottom",
        color: bodyColor,
        width: w[3],
        height: "15px",
      }),
      outlookTd(blocked, {
        bg: FILL.blocked,
        vAlign: "bottom",
        color: bodyColor,
        width: w[4],
        height: "15px",
      }),
      outlookTd(inProgress, {
        bg: FILL.white,
        vAlign: "bottom",
        color: bodyColor,
        width: w[5],
        height: "15px",
      }),
      outlookTd(notExecuted, {
        rowspan: span,
        bg: FILL.white,
        color: bodyColor,
        width: w[6],
        height: span === 2 ? "30px" : "15px",
      }),
      outlookTd(execPct, {
        bg: FILL.white,
        vAlign: "bottom",
        color: bodyColor,
        width: w[7],
        height: "15px",
      }),
      outlookTd(passPct, {
        bg: FILL.white,
        vAlign: "bottom",
        color: bodyColor,
        width: w[8],
        height: "15px",
      }),
      outlookTd(status, {
        rowspan: span,
        bg: statusFill(status),
        bold: true,
        color: bodyColor,
        width: w[9],
        height: span === 2 ? "30px" : "15px",
      }),
    ].join("")}</tr>`;
    bodyParts.push(mainTr);
    textRows.push([
      { text: label },
      { text: total },
      { text: passed },
      { text: failed },
      { text: blocked },
      { text: inProgress },
      { text: notExecuted },
      { text: execPct },
      { text: passPct },
      { text: status },
    ]);
    if (hasDelta) {
      bodyParts.push(
        `<tr>${[
          outlookTd(dPassed, {
            bg: FILL.passed,
            vAlign: "bottom",
            color: deltaColor(dPassed),
            width: w[2],
            height: "15px",
          }),
          outlookTd(dFailed, {
            bg: FILL.failed,
            vAlign: "bottom",
            color: deltaColor(dFailed),
            width: w[3],
            height: "15px",
          }),
          outlookTd(dBlocked, {
            bg: FILL.blocked,
            vAlign: "bottom",
            color: deltaColor(dBlocked),
            width: w[4],
            height: "15px",
          }),
          outlookTd(dInProg, {
            bg: FILL.white,
            vAlign: "bottom",
            color: deltaColor(dInProg),
            width: w[5],
            height: "15px",
          }),
          outlookTd(dExec, {
            bg: FILL.white,
            vAlign: "bottom",
            color: deltaColor(dExec),
            width: w[7],
            height: "15px",
          }),
          outlookTd(dPass, {
            bg: FILL.white,
            vAlign: "bottom",
            color: deltaColor(dPass),
            width: w[8],
            height: "15px",
          }),
        ].join("")}</tr>`
      );
      textRows.push([
        { text: "" },
        { text: "" },
        { text: dPassed },
        { text: dFailed },
        { text: dBlocked },
        { text: dInProg },
        { text: "" },
        { text: dExec },
        { text: dPass },
        { text: "" },
      ]);
    }
  }
  if (!bodyParts.length) return { html: "", text: "" };
  return {
    html: outlookTable(ENTITY.tableWidth, headerRow + bodyParts.join("")),
    text: `Entity Summary\n${textTable(headers, textRows)}\n`,
  };
}

function groupConsecutive(rows, key) {
  const groups = [];
  for (const row of rows) {
    const last = groups[groups.length - 1];
    if (last && last.key && row[key] === last.key) last.rows.push(row);
    else groups.push({ key: row[key], rows: [row] });
  }
  return groups;
}

function parseDailyStatusSections(sheet) {
  const blocks = [];
  const last = Math.min(sheet.rowCount || 0, 200);
  const w = DAILY.widths;
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
    const headerRow = r + 1;
    const rows = [];
    let totalRow = null;
    let execRate = "";
    let passRate = "";
    let lastChampion = "";
    for (let i = headerRow + 1; i <= Math.min(headerRow + 50, last); i++) {
      const moduleName = cellText(sheet.getRow(i).getCell(3).value).trim();
      const row = sheet.getRow(i);
      if (/^execution rate/i.test(cellText(row.getCell(4).value))) {
        execRate = percentText(row.getCell(7).value || row.getCell(9).value);
        continue;
      }
      if (/^overall pass rate/i.test(cellText(row.getCell(4).value))) {
        passRate = percentText(row.getCell(7).value || row.getCell(9).value);
        break;
      }
      if (!moduleName) continue;
      if (moduleName.toLowerCase() === "total") {
        totalRow = {
          champion: "",
          module: "Total",
          passed: numberOf(row.getCell(4).value),
          failed: numberOf(row.getCell(5).value),
          blocked: numberOf(row.getCell(6).value),
          inProgress: numberOf(row.getCell(7).value),
          notExecuted: numberOf(row.getCell(8).value),
          total: numberOf(row.getCell(9).value),
          passRate: percentText(row.getCell(10).value),
        };
        continue;
      }
      if (/^lead champion$/i.test(moduleName) || /^module$/i.test(moduleName)) continue;
      let champion = cellText(row.getCell(2).value).trim();
      if (champion) lastChampion = champion;
      else champion = lastChampion;
      rows.push({
        champion,
        module: moduleName,
        passed: numberOf(row.getCell(4).value),
        failed: numberOf(row.getCell(5).value),
        blocked: numberOf(row.getCell(6).value),
        inProgress: numberOf(row.getCell(7).value),
        notExecuted: numberOf(row.getCell(8).value),
        total: numberOf(row.getCell(9).value),
        passRate: percentText(row.getCell(10).value),
      });
    }
    const headers = [
      "Lead Champion",
      "MODULE",
      "Passed",
      "Failed",
      "Blocked",
      "In Progress",
      "Not Executed",
      "Total",
      "Pass Rate",
    ];
    const titleHtml = `<tr>${outlookTd(title, {
      colspan: 9,
      bg: FILL.dailyTitle,
      bold: true,
      nowrap: true,
      border: "titleBar",
      width: DAILY.titleWidth,
      height: "15.6pt",
    })}</tr>`;
    const headerHtml = `<tr>${headers
      .map((h, i) =>
        outlookTd(h, {
          bg: DAILY.headerBg[i],
          bold: true,
          nowrap: i === 0 || i === 1 || i === 8,
          border: "headerBar",
          width: w[i],
          height: i === 0 ? "15.6pt" : "15px",
        })
      )
      .join("")}</tr>`;
    const metricCells = (row) =>
      [
        outlookTd(row.module, {
          bg: FILL.module,
          bold: true,
          nowrap: true,
          width: w[1],
          height: "15px",
        }),
        outlookTd(row.passed, { nowrap: true, width: w[2], height: "15px" }),
        outlookTd(row.failed, { nowrap: true, width: w[3], height: "15px" }),
        outlookTd(row.blocked, { nowrap: true, width: w[4], height: "15px" }),
        outlookTd(row.inProgress, { nowrap: true, width: w[5], height: "15px" }),
        outlookTd(row.notExecuted, { nowrap: true, width: w[6], height: "15px" }),
        outlookTd(row.total, { bold: true, nowrap: true, width: w[7], height: "15px" }),
        outlookTd(row.passRate, {
          bg: passBandFill(row.passRate),
          bold: true,
          nowrap: true,
          width: w[8],
          height: "15px",
        }),
      ].join("");
    const bodyParts = [];
    for (const group of groupConsecutive(rows, "champion")) {
      const span = group.rows.length;
      group.rows.forEach((row, idx) => {
        const champ =
          idx === 0
            ? outlookTd(row.champion, {
                rowspan: span,
                bold: true,
                nowrap: true,
                width: w[0],
                height: span > 1 ? `${span * 15.6}pt` : "15.6pt",
              })
            : "";
        bodyParts.push(`<tr>${champ}${metricCells(row)}</tr>`);
      });
    }
    const textRows = rows.map((row) => [
      { text: row.champion },
      { text: row.module },
      { text: row.passed },
      { text: row.failed },
      { text: row.blocked },
      { text: row.inProgress },
      { text: row.notExecuted },
      { text: row.total },
      { text: row.passRate },
    ]);
    if (totalRow) {
      bodyParts.push(
        `<tr>${[
          outlookTd("", { bg: FILL.white, nowrap: true, width: w[0], height: "15.6pt" }),
          outlookTd("Total", { bold: true, nowrap: true, width: w[1], height: "15px" }),
          outlookTd(totalRow.passed, { bg: DAILY.totalBg[2], bold: true, width: w[2], height: "15px" }),
          outlookTd(totalRow.failed, { bg: DAILY.totalBg[3], bold: true, width: w[3], height: "15px" }),
          outlookTd(totalRow.blocked, { bg: DAILY.totalBg[4], bold: true, width: w[4], height: "15px" }),
          outlookTd(totalRow.inProgress, { bg: DAILY.totalBg[5], bold: true, width: w[5], height: "15px" }),
          outlookTd(totalRow.notExecuted, { bg: DAILY.totalBg[6], bold: true, width: w[6], height: "15px" }),
          outlookTd(totalRow.total, { bg: DAILY.totalBg[7], bold: true, nowrap: true, width: w[7], height: "15px" }),
          outlookTd(totalRow.passRate, {
            bg: passBandFill(totalRow.passRate),
            bold: true,
            nowrap: true,
            width: w[8],
            height: "15px",
          }),
        ].join("")}</tr>`
      );
      textRows.push([
        { text: "" },
        { text: "Total" },
        { text: totalRow.passed },
        { text: totalRow.failed },
        { text: totalRow.blocked },
        { text: totalRow.inProgress },
        { text: totalRow.notExecuted },
        { text: totalRow.total },
        { text: totalRow.passRate },
      ]);
    }
    if (execRate) {
      bodyParts.push(
        `<tr>${[
          outlookTd("", { border: "none", nowrap: true, width: w[0], height: "15.6pt" }),
          outlookTd("", { border: "none", nowrap: true, width: w[1], height: "15px" }),
          outlookTd("Execution Rate", {
            colspan: 3,
            bg: FILL.execRate,
            bold: true,
            nowrap: true,
            width: w[2],
            height: "15px",
          }),
          outlookTd(execRate, { colspan: 3, bold: true, nowrap: true, width: w[5], height: "15px" }),
          outlookTd("", { border: "none", nowrap: true, width: w[8], height: "15px" }),
        ].join("")}</tr>`
      );
    }
    if (passRate) {
      bodyParts.push(
        `<tr>${[
          outlookTd("", { border: "none", nowrap: true, width: w[0], height: "15.6pt" }),
          outlookTd("", { border: "none", nowrap: true, width: w[1], height: "15px" }),
          outlookTd("Overall Pass Rate", {
            colspan: 3,
            bg: FILL.overallPass,
            bold: true,
            nowrap: true,
            width: w[2],
            height: "15px",
          }),
          outlookTd(passRate, { colspan: 3, bold: true, nowrap: true, width: w[5], height: "15px" }),
          outlookTd("", { border: "none", nowrap: true, width: w[8], height: "15px" }),
        ].join("")}</tr>`
      );
    }
    const rateNote = [
      execRate ? `Execution Rate: ${execRate}` : "",
      passRate ? `Overall Pass Rate: ${passRate}` : "",
    ]
      .filter(Boolean)
      .join(" · ");
    blocks.push({
      html: outlookTable(DAILY.tableWidth, titleHtml + headerHtml + bodyParts.join("")),
      text: `${title}\n${textTable(headers, textRows)}${rateNote ? `\n${rateNote}` : ""}\n`,
    });
  }
  return blocks;
}

function parseDefectSections(sheet) {
  const blocks = [];
  const last = Math.min(sheet.rowCount || 0, 200);
  const w = DEFECT.widths;
  for (let r = 1; r <= last; r++) {
    let title = "";
    for (let c = 1; c <= 8; c++) {
      const text = cellText(sheet.getRow(r).getCell(c).value).trim();
      if (/defect summary/i.test(text)) {
        title = text;
        break;
      }
    }
    if (!title) continue;
    const headerRow = r + 1;
    const rows = [];
    let totalRow = null;
    let closure = "";
    let resolution = "";
    for (let i = headerRow + 1; i <= Math.min(headerRow + 40, last); i++) {
      const moduleName = cellText(sheet.getRow(i).getCell(3).value).trim();
      const row = sheet.getRow(i);
      const label4 = cellText(row.getCell(4).value).trim();
      if (/^closure rate/i.test(label4)) {
        closure = percentText(row.getCell(6).value || row.getCell(7).value);
        continue;
      }
      if (/^resolution rate/i.test(label4)) {
        resolution = percentText(row.getCell(6).value || row.getCell(7).value);
        break;
      }
      if (!moduleName || /^module$/i.test(moduleName)) continue;
      const rec = {
        module: moduleName,
        closed: numberOf(row.getCell(4).value),
        deferred: numberOf(row.getCell(5).value),
        fixed: numberOf(row.getCell(6).value),
        pending: numberOf(row.getCell(7).value),
        total: numberOf(row.getCell(8).value),
      };
      if (moduleName.toLowerCase() === "total") {
        totalRow = rec;
        continue;
      }
      rows.push(rec);
    }
    const headers = ["Module", "Closed", "Deferred", "Fixed", "Pending", "Total"];
    const titleHtml = `<tr>${outlookTd(title, {
      colspan: 6,
      bg: FILL.defectTitle,
      bold: true,
      nowrap: true,
      width: DEFECT.titleWidth,
      height: "15.6pt",
    })}</tr>`;
    const headerHtml = `<tr>${headers
      .map((h, i) =>
        outlookTd(h, {
          bg: DEFECT.headerBg[i],
          bold: true,
          nowrap: true,
          width: w[i],
          height: i === 0 ? "15.6pt" : "15px",
        })
      )
      .join("")}</tr>`;
    const bodyParts = rows.map((row) => {
      return `<tr>${[
        outlookTd(row.module, {
          bg: FILL.module,
          bold: true,
          nowrap: true,
          width: w[0],
          height: "15.6pt",
        }),
        outlookTd(row.closed, { nowrap: true, width: w[1], height: "15px" }),
        outlookTd(row.deferred, { nowrap: true, width: w[2], height: "15px" }),
        outlookTd(row.fixed, { nowrap: true, width: w[3], height: "15px" }),
        outlookTd(row.pending, { nowrap: true, width: w[4], height: "15px" }),
        outlookTd(row.total, { bold: true, nowrap: true, width: w[5], height: "15px" }),
      ].join("")}</tr>`;
    });
    const tableRows = rows.map((row) => [
      { text: row.module },
      { text: row.closed },
      { text: row.deferred },
      { text: row.fixed },
      { text: row.pending },
      { text: row.total },
    ]);
    if (totalRow) {
      bodyParts.push(
        `<tr>${[
          outlookTd("Total", { bg: FILL.white, bold: true, nowrap: true, width: w[0], height: "15.6pt" }),
          outlookTd(totalRow.closed, { bold: true, nowrap: true, width: w[1], height: "15px" }),
          outlookTd(totalRow.deferred, { bold: true, nowrap: true, width: w[2], height: "15px" }),
          outlookTd(totalRow.fixed, { bold: true, nowrap: true, width: w[3], height: "15px" }),
          outlookTd(totalRow.pending, { bold: true, nowrap: true, width: w[4], height: "15px" }),
          outlookTd(totalRow.total, { bold: true, nowrap: true, width: w[5], height: "15px" }),
        ].join("")}</tr>`
      );
      tableRows.push([
        { text: "Total" },
        { text: totalRow.closed },
        { text: totalRow.deferred },
        { text: totalRow.fixed },
        { text: totalRow.pending },
        { text: totalRow.total },
      ]);
    }
    if (closure) {
      bodyParts.push(
        `<tr>${[
          outlookTd("", { border: "none", nowrap: true, width: w[0], height: "15.6pt" }),
          outlookTd("Closure Rate", {
            colspan: 2,
            bg: FILL.closure,
            bold: true,
            nowrap: true,
            width: w[1],
            height: "15px",
          }),
          outlookTd(closure, { colspan: 2, bold: true, nowrap: true, width: w[3], height: "15px" }),
          outlookTd("", { border: "none", nowrap: true, width: w[5], height: "15px" }),
        ].join("")}</tr>`
      );
    }
    if (resolution) {
      bodyParts.push(
        `<tr>${[
          outlookTd("", { border: "none", nowrap: true, width: w[0], height: "15.6pt" }),
          outlookTd("Resolution Rate", {
            colspan: 2,
            bg: FILL.resolution,
            bold: true,
            nowrap: true,
            width: w[1],
            height: "15px",
          }),
          outlookTd(resolution, { colspan: 2, bold: true, nowrap: true, width: w[3], height: "15px" }),
          outlookTd("", { border: "none", nowrap: true, width: w[5], height: "15px" }),
        ].join("")}</tr>`
      );
    }
    const rateNote = [
      closure ? `Closure Rate: ${closure}` : "",
      resolution ? `Resolution Rate: ${resolution}` : "",
    ]
      .filter(Boolean)
      .join(" · ");
    blocks.push({
      html: outlookTable(DEFECT.tableWidth, titleHtml + headerHtml + bodyParts.join("")),
      text: `${title}\n${textTable(headers, tableRows)}${rateNote ? `\n${rateNote}` : ""}\n`,
    });
  }
  return blocks;
}

async function buildReportEmailTables(excelPath, sheetName) {
  const empty = { html: "", text: "" };
  if (!excelPath) return empty;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(excelPath);
  const sheet = pickSheet(wb, sheetName);
  if (!sheet) return empty;
  const entity = parseEntitySummary(sheet);
  const daily = parseDailyStatusSections(sheet);
  const defects = parseDefectSections(sheet);
  return {
    html: [
      entity.html,
      daily.length ? para("The module-wise details are provided below:") : "",
      ...daily.map((b) => b.html),
      ...defects.map((b) => b.html),
    ].join(""),
    text: [entity.text, ...daily.map((b) => b.text), ...defects.map((b) => b.text)]
      .filter(Boolean)
      .join("\n"),
  };
}

module.exports = {
  buildReportEmailTables,
  parseEntitySummary,
  parseDailyStatusSections,
  parseDefectSections,
};
