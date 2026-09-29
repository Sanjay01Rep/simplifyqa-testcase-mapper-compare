/**
 * Per-template status-email chains (SIT first; others can be added in
 * config/reporter-email-chains.json). Does not store To/CC — Outlook Reply All
 * keeps the live distribution list.
 */

const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, "../..");
const CHAINS_PATH = path.join(ROOT, "config", "reporter-email-chains.json");

const MONTHS = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function para(htmlInner) {
  return (
    `<div style="font-family: Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);">${htmlInner}</div>` +
    `<div style="font-family: Calibri, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);"><br></div>`
  );
}

function formatMailDate(sheetOrDate) {
  const raw = String(sheetOrDate || "").trim();
  if (!raw) return "";
  const m = raw.match(/\b(\d{1,2})-(\d{1,2})-(\d{4})\b/);
  if (m) {
    const day = String(Number(m[1])).padStart(2, "0");
    const month = MONTHS[Number(m[2]) - 1];
    if (month) return `${day}-${month}-${m[3]}`;
  }
  const iso = raw.match(/\b(\d{4})-(\d{2})-(\d{2})\b/);
  if (iso) {
    const month = MONTHS[Number(iso[2]) - 1];
    if (month) return `${iso[3]}-${month}-${iso[1]}`;
  }
  return raw;
}

function todayMailDate() {
  const now = new Date();
  const day = String(now.getDate()).padStart(2, "0");
  return `${day}-${MONTHS[now.getMonth()]}-${now.getFullYear()}`;
}

function loadEmailChains() {
  try {
    if (!fs.existsSync(CHAINS_PATH)) return {};
    const raw = JSON.parse(fs.readFileSync(CHAINS_PATH, "utf8"));
    return raw && typeof raw === "object" ? raw : {};
  } catch {
    return {};
  }
}

function getEmailChain(templateChoice) {
  const choice = String(templateChoice || "").trim();
  if (!choice) return null;
  const chain = loadEmailChains()[choice];
  if (!chain || chain.enabled === false) return null;
  return {
    choice,
    label: String(chain.label || `Template ${choice}`).trim(),
    search: String(chain.search || "").trim(),
    subjectTemplate: String(chain.subjectTemplate || "").trim(),
    greeting: String(chain.greeting || "Dear All,").trim() || "Dear All,",
    introTemplate: String(chain.introTemplate || "").trim(),
    closing:
      String(chain.closing || "").trim() ||
      "Please let us know if any additional details are required.\n\nThank you for your continued support.",
    attachExcel: chain.attachExcel !== false,
  };
}

function fillTemplate(template, vars) {
  return String(template || "").replace(/\{(\w+)\}/g, (_, key) =>
    vars[key] != null ? String(vars[key]) : ""
  );
}

function introToHtml(chain, intro, asOf, vs) {
  if (String(chain.choice) === "4") {
    return (
      `Please find below the <b>SIT (QA) Daily Status Report</b> as of <b>${escapeHtml(asOf)}</b>` +
      ` for <b>Gen UG, Life UG and Gen TZ</b> (changes vs ${escapeHtml(vs)}).`
    );
  }
  return escapeHtml(intro);
}

function buildStatusEmail({ chain, asOfDate, compareDate, tables }) {
  const asOf = asOfDate || todayMailDate();
  const vs = compareDate || asOf;
  const subject = fillTemplate(chain.subjectTemplate, { asOfDate: asOf, compareDate: vs });
  const intro = fillTemplate(chain.introTemplate, { asOfDate: asOf, compareDate: vs });
  const closing = chain.closing || "";
  const tableText = tables && tables.text ? `\n${tables.text}\n` : "\n";
  const tableHtml = tables && tables.html ? tables.html : "";
  const text = `${chain.greeting}\n\n${intro}\n${tableText}${closing}\n`;
  const html = [
    para(escapeHtml(chain.greeting)),
    para(introToHtml(chain, intro, asOf, vs)),
    tableHtml,
    closing
      .split(/\n\n/)
      .map((p) => para(escapeHtml(p).replace(/\n/g, "<br/>")))
      .join(""),
  ].join("");
  const wrapped = `<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head><body dir="ltr">${html}</body></html>`;
  return { subject, html: wrapped, text, search: chain.search, asOfDate: asOf, compareDate: vs };
}

module.exports = {
  CHAINS_PATH,
  formatMailDate,
  todayMailDate,
  loadEmailChains,
  getEmailChain,
  buildStatusEmail,
};
