/**
 * Delete generated and downloaded files older than MAX_AGE_DAYS.
 * Leaves the output folder alone. Keeps run-history.json.
 */

const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const MAX_AGE_DAYS = 6;
const KEEP_NAMES = new Set(["run-history.json"]);

function listFiles(dir, recursive) {
  const out = [];
  if (!dir || !fs.existsSync(dir)) return out;
  let names;
  try {
    names = fs.readdirSync(dir);
  } catch {
    return out;
  }
  for (const name of names) {
    const full = path.join(dir, name);
    let st;
    try {
      st = fs.statSync(full);
    } catch {
      continue;
    }
    if (st.isDirectory()) {
      if (recursive) out.push(...listFiles(full, true));
      continue;
    }
    if (st.isFile()) out.push({ full, name, mtimeMs: st.mtimeMs });
  }
  return out;
}

function cleanupDir(dir, { cutoff, protect, recursive, log }) {
  let removed = 0;
  for (const file of listFiles(dir, recursive)) {
    if (protect.has(file.name)) continue;
    if (file.mtimeMs >= cutoff) continue;
    try {
      fs.unlinkSync(file.full);
      removed += 1;
    } catch (err) {
      if (log) log(`WARN: could not delete old file ${file.name}: ${err.message}`);
    }
  }
  return removed;
}

function cleanupOldGeneratedAndDownloads({
  root = ROOT,
  maxAgeDays = MAX_AGE_DAYS,
  now = Date.now(),
  log,
} = {}) {
  const cutoff = now - maxAgeDays * 24 * 60 * 60 * 1000;
  const dirs = [
    { dir: path.join(root, "downloads"), recursive: false, label: "downloads" },
    { dir: path.join(root, "Generated Excel file"), recursive: true, label: "generated" },
    { dir: path.join(root, "logs"), recursive: false, label: "logs" },
  ];
  const counts = {};
  let total = 0;
  for (const item of dirs) {
    const n = cleanupDir(item.dir, {
      cutoff,
      protect: KEEP_NAMES,
      recursive: item.recursive,
      log,
    });
    counts[item.label] = n;
    total += n;
  }
  if (log && total) {
    log(
      `Cleaned files older than ${maxAgeDays} day(s): ${total} removed` +
        ` (downloads ${counts.downloads}, generated ${counts.generated}, logs ${counts.logs}).`
    );
  }
  return { removed: total, counts, maxAgeDays };
}

module.exports = {
  MAX_AGE_DAYS,
  cleanupOldGeneratedAndDownloads,
};
