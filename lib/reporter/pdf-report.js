const fs = require("fs");
const path = require("path");
const PDFDocument = require("pdfkit");
const ExcelJS = require("exceljs");
const {
  extractStatusSections,
  extractDefectSections,
} = require("./compare");

function rate(n, d) {
  if (!d) return 1;
  return n / d;
}

function pct(n) {
  return `${Math.round(n * 1000) / 10}%`;
}

function drawTable(doc, startX, startY, headers, rows, colWidths, opts = {}) {
  const rowH = opts.rowH || 14;
  const fontSize = opts.fontSize || 7;
  let x = startX;
  let y = startY;

  doc.fontSize(fontSize).font("Helvetica-Bold");
  headers.forEach((h, i) => {
    doc.rect(x, y, colWidths[i], rowH).stroke("#888");
    if (opts.headerFills && opts.headerFills[i]) {
      doc.save();
      doc.rect(x, y, colWidths[i], rowH).fill(opts.headerFills[i]);
      doc.restore();
      doc.fillColor("#111").strokeColor("#888");
      doc.rect(x, y, colWidths[i], rowH).stroke();
    }
    doc.fillColor("#111").text(String(h), x + 2, y + 3, {
      width: colWidths[i] - 4,
      align: "center",
    });
    x += colWidths[i];
  });
  y += rowH;

  doc.font("Helvetica");
  for (const row of rows) {
    if (y > doc.page.height - 50) {
      doc.addPage({ layout: "landscape", size: "A4" });
      y = 40;
    }
    x = startX;
    const isTotal = String(row[0] || "").toLowerCase() === "total";
    row.forEach((val, i) => {
      if (isTotal) {
        doc.save();
        doc.rect(x, y, colWidths[i], rowH).fill("#FFF3CD");
        doc.restore();
      }
      doc.rect(x, y, colWidths[i], rowH).stroke("#bbb");
      doc
        .fillColor("#111")
        .font(isTotal ? "Helvetica-Bold" : "Helvetica")
        .text(String(val), x + 2, y + 3, {
          width: colWidths[i] - 4,
          align: i === 0 ? "left" : "center",
        });
      x += colWidths[i];
    });
    y += rowH;
  }
  return y;
}

async function generateReportPdf({
  excelPath,
  sheetName,
  outPath,
  branding,
}) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(excelPath);
  const sheet =
    (sheetName && wb.worksheets.find((s) => s.name === sheetName)) ||
    wb.worksheets[wb.worksheets.length - 1];
  if (!sheet) throw new Error("No sheet found for PDF export.");

  const statusSections = extractStatusSections(sheet);
  const defectSections = extractDefectSections(sheet);
  const title = (branding && branding.title) || "ICEA LION Reporter";

  await new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      layout: "landscape",
      size: "A4",
      margin: 28,
    });
    const stream = fs.createWriteStream(outPath);
    doc.pipe(stream);
    stream.on("finish", resolve);
    stream.on("error", reject);

    const pageW = doc.page.width;
    const left = 28;

    // Header logos + title
    const logoH = 36;
    try {
      if (branding?.logoRight?.exists) {
        doc.image(branding.logoRight.absolute, left, 22, { height: logoH });
      }
    } catch {
      /* ignore */
    }
    try {
      if (branding?.logoLeft?.exists) {
        doc.image(branding.logoLeft.absolute, pageW - 28 - 110, 18, {
          height: logoH + 4,
        });
      }
    } catch {
      /* ignore */
    }

    doc
      .font("Helvetica-Bold")
      .fontSize(14)
      .fillColor("#003f6e")
      .text(title, left + 120, 28, {
        width: pageW - 280,
        align: "center",
      });
    doc
      .moveTo(left + 120, 48)
      .lineTo(pageW - 150, 48)
      .strokeColor("#c62828")
      .lineWidth(2)
      .stroke();

    doc
      .font("Helvetica")
      .fontSize(8)
      .fillColor("#555")
      .text(`Sheet: ${sheet.name} · Generated: ${new Date().toLocaleString()}`, left, 56, {
        width: pageW - 56,
        align: "center",
      });

    let y = 72;
    const statusWidths = [110, 48, 48, 48, 52, 58, 42, 48];
    const statusHeaders = [
      "MODULE",
      "Passed",
      "Failed",
      "Blocked",
      "In Progress",
      "Not Executed",
      "Total",
      "Pass Rate",
    ];
    const statusFills = [
      "#E8EEF2",
      "#C6EFCE",
      "#FFC7CE",
      "#FCE4D6",
      "#D9D9D9",
      "#FFE699",
      "#BDD7EE",
      "#C6EFCE",
    ];

    for (const section of statusSections) {
      if (y > doc.page.height - 120) {
        doc.addPage({ layout: "landscape", size: "A4" });
        y = 40;
      }
      doc
        .font("Helvetica-Bold")
        .fontSize(9)
        .fillColor("#003f6e")
        .text(section.title, left, y);
      y += 14;

      const names = Object.keys(section.modules);
      const rows = names.map((name) => {
        const m = section.modules[name];
        const passRate = rate(m.passed, m.total - m.blocked);
        return [
          name,
          m.passed,
          m.failed,
          m.blocked,
          m.inProgress,
          m.notExecuted,
          m.total,
          pct(passRate),
        ];
      });
      const totals = names.reduce(
        (acc, name) => {
          const m = section.modules[name];
          acc.passed += m.passed;
          acc.failed += m.failed;
          acc.blocked += m.blocked;
          acc.inProgress += m.inProgress;
          acc.notExecuted += m.notExecuted;
          acc.total += m.total;
          return acc;
        },
        {
          passed: 0,
          failed: 0,
          blocked: 0,
          inProgress: 0,
          notExecuted: 0,
          total: 0,
        }
      );
      rows.push([
        "Total",
        totals.passed,
        totals.failed,
        totals.blocked,
        totals.inProgress,
        totals.notExecuted,
        totals.total,
        pct(rate(totals.passed, totals.total - totals.blocked)),
      ]);

      y = drawTable(doc, left, y, statusHeaders, rows, statusWidths, {
        headerFills: statusFills,
      });
      const execRate = rate(
        totals.passed + totals.failed,
        totals.total - totals.blocked
      );
      const overall = rate(totals.passed, totals.total - totals.blocked);
      doc
        .font("Helvetica")
        .fontSize(8)
        .fillColor("#111")
        .text(
          `Execution Rate: ${pct(execRate)}   |   Overall Pass Rate: ${pct(overall)}`,
          left,
          y + 4
        );
      y += 22;
    }

    const defWidths = [120, 55, 55, 55, 55, 50];
    const defHeaders = ["Module", "Closed", "Deferred", "Fixed", "Pending", "Total"];
    const defFills = ["#E8EEF2", "#C6EFCE", "#FCE4D6", "#BDD7EE", "#FFC7CE", "#5B9BD5"];

    for (const section of defectSections) {
      if (y > doc.page.height - 120) {
        doc.addPage({ layout: "landscape", size: "A4" });
        y = 40;
      }
      doc
        .font("Helvetica-Bold")
        .fontSize(9)
        .fillColor("#003f6e")
        .text(section.title, left, y);
      y += 14;

      const names = Object.keys(section.modules);
      const rows = names.map((name) => {
        const m = section.modules[name];
        return [name, m.closed, m.deferred, m.fixed, m.pending, m.total];
      });
      const totals = names.reduce(
        (acc, name) => {
          const m = section.modules[name];
          acc.closed += m.closed;
          acc.deferred += m.deferred;
          acc.fixed += m.fixed;
          acc.pending += m.pending;
          acc.total += m.total;
          return acc;
        },
        { closed: 0, deferred: 0, fixed: 0, pending: 0, total: 0 }
      );
      rows.push([
        "Total",
        totals.closed,
        totals.deferred,
        totals.fixed,
        totals.pending,
        totals.total,
      ]);
      y = drawTable(doc, left, y, defHeaders, rows, defWidths, {
        headerFills: defFills,
      });
      const denom = totals.total - totals.deferred;
      const closure = rate(totals.closed, denom);
      const resolution = rate(totals.closed + totals.fixed, denom);
      doc
        .fontSize(8)
        .text(
          `Closure Rate: ${pct(closure)}   |   Resolution Rate: ${pct(resolution)}`,
          left,
          y + 4
        );
      y += 22;
    }

    doc.end();
  });

  return outPath;
}

module.exports = { generateReportPdf };
