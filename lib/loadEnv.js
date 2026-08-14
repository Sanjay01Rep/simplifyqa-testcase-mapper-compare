/**
 * Minimal .env loader (no dependency).
 * Does not override variables already set in the process environment.
 */
const fs = require("fs");
const path = require("path");

function stripQuotes(value) {
  const v = String(value || "").trim();
  if (
    (v.startsWith('"') && v.endsWith('"')) ||
    (v.startsWith("'") && v.endsWith("'"))
  ) {
    return v.slice(1, -1);
  }
  return v;
}

function loadEnvFile(filePath) {
  const resolved = path.resolve(filePath);
  if (!fs.existsSync(resolved)) return { loaded: false, path: resolved, keys: [] };
  const text = fs.readFileSync(resolved, "utf8");
  const keys = [];
  for (const raw of text.split(/\r?\n/)) {
    const line = raw.trim();
    if (!line || line.startsWith("#")) continue;
    const eq = line.indexOf("=");
    if (eq <= 0) continue;
    const key = line.slice(0, eq).trim();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) continue;
    if (Object.prototype.hasOwnProperty.call(process.env, key)) continue;
    process.env[key] = stripQuotes(line.slice(eq + 1));
    keys.push(key);
  }
  return { loaded: true, path: resolved, keys };
}

function loadProjectEnv(rootDir) {
  return loadEnvFile(path.join(rootDir || process.cwd(), ".env"));
}

module.exports = { loadEnvFile, loadProjectEnv };
