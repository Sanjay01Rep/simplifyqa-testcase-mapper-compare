const nodemailer = require("nodemailer");
const fs = require("fs");
const path = require("path");

function smtpConfigured() {
  return Boolean(
    String(process.env.SMTP_HOST || "").trim() &&
      String(process.env.SMTP_FROM || "").trim()
  );
}

function createTransport() {
  const host = String(process.env.SMTP_HOST || "").trim();
  if (!host) return null;
  const port = Number(process.env.SMTP_PORT || 587);
  const secure =
    String(process.env.SMTP_SECURE || "").toLowerCase() === "true" ||
    port === 465;
  const user = String(process.env.SMTP_USER || "").trim();
  const pass = String(process.env.SMTP_PASS || "").trim();
  const options = {
    host,
    port,
    secure,
  };
  if (user || pass) {
    options.auth = { user, pass };
  }
  return nodemailer.createTransport(options);
}

async function sendEmailNotify({
  to,
  subject,
  text,
  attachments,
  log,
}) {
  if (!to || !to.length) return { sent: false, reason: "no recipients" };
  if (!smtpConfigured()) {
    if (log) log("Notify email skipped: SMTP_HOST/SMTP_FROM not set in .env");
    return { sent: false, reason: "smtp not configured" };
  }
  const transport = createTransport();
  if (!transport) return { sent: false, reason: "smtp not configured" };

  const from = String(process.env.SMTP_FROM || "").trim();
  await transport.sendMail({
    from,
    to: to.join(", "),
    subject,
    text,
    attachments: attachments || [],
  });
  if (log) log(`Notify email sent to: ${to.join(", ")}`);
  return { sent: true };
}

async function sendTeamsNotify({ webhookUrl, title, text, log }) {
  if (!webhookUrl) return { sent: false, reason: "no webhook" };
  const body = {
    "@type": "MessageCard",
    "@context": "http://schema.org/extensions",
    summary: title,
    themeColor: "005696",
    title,
    text: String(text || "").replace(/\n/g, "<br/>"),
  };
  const res = await fetch(webhookUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const msg = `Teams webhook failed: HTTP ${res.status}`;
    if (log) log(msg);
    return { sent: false, reason: msg };
  }
  if (log) log("Notify Teams webhook sent.");
  return { sent: true };
}

async function notifyRunResult({
  ok,
  notifyConfig,
  title,
  summaryLines,
  excelPath,
  pdfPath,
  logPath,
  log,
}) {
  if (!notifyConfig) return;
  if (ok && !notifyConfig.onSuccess) return;
  if (!ok && !notifyConfig.onFailure) return;

  const subject = ok
    ? `[OK] ${title}`
    : `[FAILED] ${title}`;
  const text = [
    subject,
    "",
    ...(summaryLines || []),
    "",
    logPath ? `Log: ${logPath}` : "",
    excelPath ? `Excel: ${excelPath}` : "",
    pdfPath ? `PDF: ${pdfPath}` : "",
  ]
    .filter(Boolean)
    .join("\n");

  const attachments = [];
  if (ok && excelPath && fs.existsSync(excelPath)) {
    attachments.push({
      filename: path.basename(excelPath),
      path: excelPath,
    });
  }
  if (ok && pdfPath && fs.existsSync(pdfPath)) {
    attachments.push({
      filename: path.basename(pdfPath),
      path: pdfPath,
    });
  }
  if (!ok && logPath && fs.existsSync(logPath)) {
    attachments.push({
      filename: path.basename(logPath),
      path: logPath,
    });
  }

  const results = [];
  try {
    results.push(
      await sendEmailNotify({
        to: notifyConfig.emails,
        subject,
        text,
        attachments,
        log,
      })
    );
  } catch (err) {
    if (log) log(`Notify email error: ${err.message}`);
    results.push({ sent: false, reason: err.message });
  }

  try {
    results.push(
      await sendTeamsNotify({
        webhookUrl: notifyConfig.teamsWebhook,
        title: subject,
        text,
        log,
      })
    );
  } catch (err) {
    if (log) log(`Notify Teams error: ${err.message}`);
    results.push({ sent: false, reason: err.message });
  }

  return results;
}

module.exports = {
  notifyRunResult,
  smtpConfigured,
};
