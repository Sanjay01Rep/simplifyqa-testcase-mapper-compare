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

function loadEnvFile(filePath, override = false) {
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
    if (!override && Object.prototype.hasOwnProperty.call(process.env, key)) continue;
    process.env[key] = stripQuotes(line.slice(eq + 1));
    keys.push(key);
  }
  return { loaded: true, path: resolved, keys };
}

function loadProjectEnv(rootDir, override = false) {
  return loadEnvFile(path.join(rootDir || process.cwd(), ".env"), override);
}

const TOKEN_ENV_KEY = "SIMPLIFYQA_BEARER_TOKEN";

function normalizeBearerToken(raw) {
  return String(raw || "")
    .trim()
    .replace(/^Bearer\s+/i, "");
}

function getBearerToken(rootDir) {
  const envPath = path.join(rootDir || process.cwd(), ".env");
  if (fs.existsSync(envPath)) {
    try {
      const text = fs.readFileSync(envPath, "utf8");
      for (const raw of text.split(/\r?\n/)) {
        const line = raw.trim();
        if (!line || line.startsWith("#")) continue;
        const eq = line.indexOf("=");
        if (eq <= 0) continue;
        const key = line.slice(0, eq).trim();
        if (key === TOKEN_ENV_KEY) {
          const val = stripQuotes(line.slice(eq + 1));
          if (val) {
            process.env[TOKEN_ENV_KEY] = val;
            return normalizeBearerToken(val);
          }
        }
      }
    } catch {}
  }
  return normalizeBearerToken(process.env[TOKEN_ENV_KEY]);
}

function readJwtExpMs(token) {
  try {
    const part = String(token || "").split(".")[1];
    if (!part) return null;
    const json = Buffer.from(part.replace(/-/g, "+").replace(/_/g, "/"), "base64").toString("utf8");
    const payload = JSON.parse(json);
    if (!payload.exp) return null;
    return Number(payload.exp) * 1000;
  } catch {
    return null;
  }
}

function formatDuration(ms) {
  const totalSec = Math.max(0, Math.floor(ms / 1000));
  const hours = Math.floor(totalSec / 3600);
  const mins = Math.floor((totalSec % 3600) / 60);
  if (hours >= 48) {
    const days = Math.floor(hours / 24);
    return `${days} day${days === 1 ? "" : "s"}`;
  }
  if (hours > 0) return `${hours}h ${mins}m`;
  if (mins > 0) return `${mins} minute${mins === 1 ? "" : "s"}`;
  return `${totalSec} second${totalSec === 1 ? "" : "s"}`;
}

/** UI-safe token status (never includes the secret). */
function getTokenStatus(rootDir) {
  const token = getBearerToken(rootDir);
  const placeholder =
    !token ||
    token === "your_bearer_token_here" ||
    token === "your_simplifyqa_bearer_token_here";
  if (placeholder) {
    return {
      present: false,
      expired: false,
      expiresAt: null,
      expiresInMs: null,
      message: "No token set. Paste a SimplifyQA bearer token below.",
    };
  }
  const expiresAtMs = readJwtExpMs(token);
  if (expiresAtMs == null) {
    return {
      present: true,
      expired: false,
      expiresAt: null,
      expiresInMs: null,
      message: "Token is set. Expiry time is not available for this token format.",
    };
  }
  const expiresInMs = expiresAtMs - Date.now();
  const expired = expiresInMs <= 0;
  return {
    present: true,
    expired,
    expiresAt: new Date(expiresAtMs).toISOString(),
    expiresInMs,
    message: expired
      ? "Token expired. Paste a fresh bearer token below."
      : `Token OK · expires in about ${formatDuration(expiresInMs)}.`,
  };
}

/**
 * Write SIMPLIFYQA_BEARER_TOKEN into .env and update process.env
 * so live calls work immediately without restarting the server.
 */
function saveBearerToken(rawToken, rootDir) {
  const token = normalizeBearerToken(rawToken);
  if (!token || token === "your_bearer_token_here") {
    const err = new Error("Paste a valid SimplifyQA bearer token before saving.");
    err.status = 400;
    throw err;
  }

  const envPath = path.join(rootDir || process.cwd(), ".env");
  let text = "";
  if (fs.existsSync(envPath)) {
    text = fs.readFileSync(envPath, "utf8");
  }

  const line = `${TOKEN_ENV_KEY}=${token}`;
  const re = new RegExp(`^\\s*${TOKEN_ENV_KEY}\\s*=.*$`, "m");
  if (re.test(text)) {
    text = text.replace(re, line);
  } else {
    const trimmed = text.replace(/\s+$/, "");
    text = trimmed ? `${trimmed}\n${line}\n` : `${line}\n`;
  }

  fs.writeFileSync(envPath, text, "utf8");
  process.env[TOKEN_ENV_KEY] = token;
  return getTokenStatus(rootDir);
}

module.exports = {
  loadEnvFile,
  loadProjectEnv,
  normalizeBearerToken,
  getBearerToken,
  readJwtExpMs,
  getTokenStatus,
  saveBearerToken,
};
