/**
 * Auto-bump the header badge (package.json version) for product fixes
 * and enhancements. Docs, tests, and config-only edits do not bump.
 *
 * At most one bump until that version is committed, so npm start while
 * iterating does not keep incrementing. A later enhancement can upgrade
 * a pending patch (1.1.1) to a minor (1.2.0).
 */
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");

const ROOT = path.join(__dirname, "..");
const PKG_PATH = path.join(ROOT, "package.json");
const LOCK_PATH = path.join(ROOT, "package-lock.json");
const INDEX_PATH = path.join(ROOT, "public", "index.html");
const START_HERE_PATH = path.join(ROOT, "START-HERE.md");

function normalize(rel) {
  return String(rel || "")
    .replace(/\\/g, "/")
    .replace(/^\.\//, "");
}

function isSkipPath(rel) {
  const n = normalize(rel);
  if (!n) return true;
  if (n.endsWith(".md")) return true;
  if (n.startsWith("docs/")) return true;
  if (n.startsWith(".cursor/")) return true;
  if (n.startsWith("test/")) return true;
  if (n.startsWith("Generated Excel file/")) return true;
  if (n.startsWith("Client doc/") || n.startsWith("Kenya ")) return true;
  if (
    n === "mapping.properties" ||
    n === "mapping.properties.example" ||
    n === "jobs.json" ||
    n === ".env" ||
    n === ".env.example" ||
    n === ".gitignore"
  ) {
    return true;
  }
  if (n === "package.json" || n === "package-lock.json") return true;
  if (n === "scripts/sync-app-version.js") return true;
  return false;
}

function isProductPath(rel) {
  if (isSkipPath(rel)) return false;
  const n = normalize(rel);
  return (
    n.startsWith("lib/") ||
    n.startsWith("public/") ||
    n === "server.js" ||
    n === "map_to_simplifyqa.js" ||
    n === "scripts/free-port.js"
  );
}

function parseSemver(value) {
  const m = String(value || "0.0.0").trim().match(/^v?(\d+)\.(\d+)\.(\d+)/);
  if (!m) return { major: 0, minor: 0, patch: 0, raw: "0.0.0" };
  return { major: Number(m[1]), minor: Number(m[2]), patch: Number(m[3]), raw: `${m[1]}.${m[2]}.${m[3]}` };
}

function formatSemver(v) {
  return `${v.major}.${v.minor}.${v.patch}`;
}

function classifyChange(paths, extra = {}) {
  const added = new Set((extra.addedPaths || []).map(normalize));
  const product = [...new Set((paths || []).map(normalize))].filter(isProductPath);
  if (!product.length) return "none";
  const hasNewProduct = product.some((p) => added.has(p));
  if (hasNewProduct || extra.enhancementHint) return "minor";
  return "patch";
}

function nextVersion(current, kind, headVersion) {
  const cur = parseSemver(current);
  const head = parseSemver(headVersion || current);
  if (kind !== "patch" && kind !== "minor") return cur.raw;
  const alreadyBumped = cur.raw !== head.raw;
  if (!alreadyBumped) {
    if (kind === "minor") return formatSemver({ major: head.major, minor: head.minor + 1, patch: 0 });
    return formatSemver({ major: head.major, minor: head.minor, patch: head.patch + 1 });
  }
  if (kind === "minor" && cur.major === head.major && cur.minor === head.minor && cur.patch > head.patch) {
    return formatSemver({ major: head.major, minor: head.minor + 1, patch: 0 });
  }
  return cur.raw;
}

function git(cmd) {
  try {
    return execSync(cmd, { cwd: ROOT, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"] }).trim();
  } catch {
    return "";
  }
}

function headPackageVersion() {
  const raw = git("git show HEAD:package.json");
  const fallback = parseSemver(JSON.parse(fs.readFileSync(PKG_PATH, "utf8")).version).raw;
  if (!raw) return fallback;
  try {
    return parseSemver(JSON.parse(raw).version).raw;
  } catch {
    return fallback;
  }
}

function workingTreeChanges() {
  const out = git("git status --porcelain -uall");
  const paths = [];
  const added = [];
  for (const line of out.split(/\r?\n/)) {
    if (line.length < 4) continue;
    const code = line.slice(0, 2);
    let file = line.slice(3).trim();
    if (file.includes(" -> ")) file = file.split(" -> ").pop().trim();
    if (file.startsWith('"') && file.endsWith('"')) {
      try {
        file = JSON.parse(file);
      } catch {
        file = file.slice(1, -1);
      }
    }
    const n = normalize(file);
    if (!n) continue;
    paths.push(n);
    if (/[A?]/.test(code)) added.push(n);
  }
  return { paths, added };
}

function writeVersion(version) {
  const pkg = JSON.parse(fs.readFileSync(PKG_PATH, "utf8"));
  pkg.version = version;
  fs.writeFileSync(PKG_PATH, `${JSON.stringify(pkg, null, 2)}\n`);

  if (fs.existsSync(LOCK_PATH)) {
    let lock = fs.readFileSync(LOCK_PATH, "utf8");
    let n = 0;
    lock = lock.replace(/"version": "\d+\.\d+\.\d+"/g, (m) => {
      n += 1;
      return n <= 2 ? `"version": "${version}"` : m;
    });
    fs.writeFileSync(LOCK_PATH, lock);
  }

  if (fs.existsSync(INDEX_PATH)) {
    const html = fs.readFileSync(INDEX_PATH, "utf8");
    let next = html.replace(
      /(id="appVersion"[^>]*>)v?\d+\.\d+\.\d+/,
      `$1v${version}`
    );
    next = next.replace(
      /(\/(?:app\.js|styles\.css)\?v=)\d+\.\d+\.\d+/g,
      `$1${version}`
    );
    // Ensure cache-bust query exists even if the file had a bare script src.
    next = next.replace(
      /(<script\s+src="\/app\.js)(")/,
      `$1?v=${version}$2`
    );
    next = next.replace(
      /(<link\s+rel="stylesheet"\s+href="\/styles\.css)(")/,
      `$1?v=${version}$2`
    );
    fs.writeFileSync(INDEX_PATH, next);
  }

  if (fs.existsSync(START_HERE_PATH)) {
    const md = fs.readFileSync(START_HERE_PATH, "utf8");
    fs.writeFileSync(
      START_HERE_PATH,
      md.replace(
        /ICEA LION Testcase Review UI  v\d+\.\d+\.\d+/,
        `ICEA LION Testcase Review UI  v${version}`
      )
    );
  }
}

function syncAppVersion({ silent } = {}) {
  if (String(process.env.SKIP_VERSION_SYNC || "") === "1") {
    return { bumped: false, version: parseSemver(JSON.parse(fs.readFileSync(PKG_PATH, "utf8")).version).raw, kind: "skipped" };
  }
  const pkg = JSON.parse(fs.readFileSync(PKG_PATH, "utf8"));
  const current = parseSemver(pkg.version).raw;
  const head = headPackageVersion();
  const { paths, added } = workingTreeChanges();
  const kind = classifyChange(paths, { addedPaths: added });
  const next = nextVersion(current, kind, head);
  if (next === current) {
    if (!silent) console.log(`App version v${current}`);
    return { bumped: false, version: current, kind };
  }
  writeVersion(next);
  const why = kind === "minor" ? "enhancement" : "fix";
  if (!silent) console.log(`App version v${current} → v${next} (${why})`);
  return { bumped: true, version: next, kind, from: current };
}

if (require.main === module) {
  syncAppVersion({ silent: process.argv.includes("--silent") });
}

module.exports = {
  classifyChange,
  nextVersion,
  isProductPath,
  isSkipPath,
  parseSemver,
  syncAppVersion,
};
