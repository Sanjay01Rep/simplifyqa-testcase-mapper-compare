const path = require("path");
const fs = require("fs");

function readBranding(props, rootDir = path.join(__dirname, "../..")) {
  const title =
    String(props.APP_TITLE || "").trim() || "ICEA LION Reporter";
  const tagline =
    String(props.APP_TAGLINE || "").trim() ||
    "Generate FMS status Excel reports from SimplifyQA";
  const leftRel =
    String(props.LOGO_LEFT || "").trim() || "Logo/ICEA Lion.png";
  const rightRel =
    String(props.LOGO_RIGHT || "").trim() || "Logo/Simplify-icon.png";
  const rightLabel =
    String(props.LOGO_RIGHT_LABEL || "").trim() || "Simplify3x";

  const leftAbs = path.isAbsolute(leftRel) ? leftRel : path.join(rootDir, leftRel);
  const rightAbs = path.isAbsolute(rightRel) ? rightRel : path.join(rootDir, rightRel);

  return {
    title,
    tagline,
    logoRightLabel: rightLabel,
    logoLeft: {
      relative: leftRel.replace(/\\/g, "/"),
      absolute: leftAbs,
      exists: fs.existsSync(leftAbs),
      url: `/logo/ICEA%20Lion.png`,
    },
    logoRight: {
      relative: rightRel.replace(/\\/g, "/"),
      absolute: rightAbs,
      exists: fs.existsSync(rightAbs),
      url: `/logo/Simplify-icon.png`,
    },
  };
}

function parseScheduleTimes(raw) {
  const text = String(raw || "").trim();
  if (!text) return [];
  const out = [];
  const seen = new Set();
  for (const part of text.split(/[,;]+/)) {
    const t = part.trim();
    if (!t) continue;
    const m = /^(\d{1,2}):(\d{2})$/.exec(t);
    if (!m) continue;
    let hh = Number(m[1]);
    let mm = Number(m[2]);
    if (hh < 0 || hh > 23 || mm < 0 || mm > 59) continue;
    const key = `${String(hh).padStart(2, "0")}:${String(mm).padStart(2, "0")}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(key);
  }
  return out.sort();
}

function readScheduleConfig(props) {
  const enabled =
    String(props.SCHEDULE_ENABLED || "false").trim().toLowerCase() ===
    "true";
  const times = parseScheduleTimes(props.SCHEDULE_TIMES || "");
  return { enabled, times };
}

function readNotifyConfig(props) {
  const emails = String(props.NOTIFY_EMAIL || "")
    .split(/[,;]+/)
    .map((e) => e.trim())
    .filter(Boolean);
  const teamsWebhook = String(props.TEAMS_WEBHOOK_URL || "").trim();
  const on = String(props.NOTIFY_ON || "both").trim().toLowerCase();
  return {
    emails,
    teamsWebhook,
    onSuccess: on === "both" || on === "success",
    onFailure: on === "both" || on === "failure",
  };
}

module.exports = {
  readBranding,
  parseScheduleTimes,
  readScheduleConfig,
  readNotifyConfig,
};
