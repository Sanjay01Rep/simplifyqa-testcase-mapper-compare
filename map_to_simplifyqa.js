/**
 * SimplifyQA Client Testcase Mapper (CLI)
 *
 * Usage:
 *   node map_to_simplifyqa.js --all
 *   node map_to_simplifyqa.js "Investment Managment.xlsx"
 *   node map_to_simplifyqa.js --list
 */

const fs = require("fs");
const path = require("path");
const {
  loadProperties,
  loadJobs,
  resolveJobConfig,
  isLockFile,
  todayStamp,
  safeBaseName,
  ensureDir,
  mapFromBuffers,
  writeWorkbook,
  writeLog,
} = require("./lib/mapper");

const base = __dirname;
const CLIENT_DIR = path.join(base, "Client doc");
const KENYA_DIR = path.join(base, "Kenya doc");
const KENYA_ORIGINAL_DIR = path.join(base, "Kenya orginial testcase");
const OUT_DIR = path.join(base, "Generated Excel file");
const LOG_DIR = path.join(OUT_DIR, "logs");
const JOBS_PATH = path.join(base, "jobs.json");
const PROPS_PATH = path.join(base, "mapping.properties");

function listClientFiles() {
  if (!fs.existsSync(CLIENT_DIR)) return [];
  return fs
    .readdirSync(CLIENT_DIR)
    .filter((f) => /\.xlsx$/i.test(f) && !isLockFile(f));
}

function listKenyaXlsx() {
  const dirs = [KENYA_DIR, KENYA_ORIGINAL_DIR];
  const files = [];
  for (const dir of dirs) {
    if (!fs.existsSync(dir)) continue;
    for (const f of fs.readdirSync(dir)) {
      if (!isLockFile(f) && /\.xlsx$/i.test(f)) {
        files.push({ dir, name: f });
      }
    }
  }
  return files;
}

function findKenyaFor(clientFileName) {
  const baseName = safeBaseName(clientFileName).toLowerCase();
  const files = listKenyaXlsx();
  const hit =
    files.find((f) => safeBaseName(f.name).toLowerCase() === baseName) ||
    files.find((f) => safeBaseName(f.name).toLowerCase().includes(baseName)) ||
    files.find((f) => baseName.includes(safeBaseName(f.name).toLowerCase()));
  return hit || null;
}

function printSummary(summary) {
  console.log(`\n=== ${summary.clientFile} ===`);
  console.log(`Module/Entity : ${summary.module} / ${summary.entity}`);
  console.log(`Output        : Generated Excel file/${summary.outName}`);
  console.log(`Log           : Generated Excel file/logs/${summary.logName}`);
  console.log(
    `TCs           : ${summary.mappedTcCount} [${summary.tcMatch ? "PASS" : "FAIL"}]`
  );
  console.log(
    `Steps         : ${summary.mappedStepCount} [${summary.stepMatch ? "PASS" : "FAIL"}]`
  );
  console.log(`Kenya prereq  : ${summary.kenyaMatched}/${summary.kenyaCount}`);
  console.log(`Issues logged : ${summary.issues.length}`);
  const fixes = summary.issues.filter((i) => i.severity === "FIX").length;
  const warns = summary.issues.filter((i) => i.severity === "WARN").length;
  const errors = summary.issues.filter((i) => i.severity === "ERROR").length;
  console.log(`  FIX=${fixes} WARN=${warns} ERROR=${errors}`);
}

async function processFile(fileName, jobsConfig, propsFallback) {
  const config = resolveJobConfig(fileName, jobsConfig, propsFallback);
  if (!config.Module || !config.Entity) {
    throw new Error("Module/Entity missing in jobs.json or mapping.properties.");
  }

  const clientPath = path.join(CLIENT_DIR, fileName);
  const clientBuffer = fs.readFileSync(clientPath);
  const kenyaHit = findKenyaFor(fileName);
  const kenyaBuffer = kenyaHit
    ? fs.readFileSync(path.join(kenyaHit.dir, kenyaHit.name))
    : null;

  const mapped = await mapFromBuffers({
    clientFileName: fileName,
    clientBuffer,
    kenyaFileName: kenyaHit ? kenyaHit.name : null,
    kenyaBuffer,
    config,
  });

  const stamp = todayStamp();
  const baseName = safeBaseName(fileName);
  const outName = `${baseName}_SimplifyQA_${stamp}.xlsx`;
  const logName = `${baseName}_mapping_${stamp}.log`;
  ensureDir(OUT_DIR);
  ensureDir(LOG_DIR);
  writeWorkbook(mapped.outRows, path.join(OUT_DIR, outName));
  mapped.summary.outName = outName;
  mapped.summary.logName = logName;
  writeLog(path.join(LOG_DIR, logName), mapped.summary);
  printSummary(mapped.summary);
  return mapped.summary;
}

async function main() {
  const args = process.argv.slice(2);
  const jobsConfig = loadJobs(JOBS_PATH);
  const propsFallback = loadProperties(PROPS_PATH);

  if (args.includes("--list")) {
    console.log("Client files:");
    listClientFiles().forEach((f) => {
      const cfg = resolveJobConfig(f, jobsConfig, propsFallback);
      console.log(
        ` - ${f} | Module=${cfg.Module || "(missing)"} Entity=${cfg.Entity || "(missing)"} | jobs.json=${cfg.fromJobs ? "yes" : "no"}`
      );
    });
    return;
  }

  let targets = [];
  if (args.includes("--all") || args.length === 0) {
    const fromJobs = (jobsConfig.jobs || []).map((j) => j.file);
    const onDisk = listClientFiles();
    const ordered = [...fromJobs.filter((f) => onDisk.includes(f))];
    for (const f of onDisk) {
      if (!ordered.includes(f)) ordered.push(f);
    }
    targets = ordered;
  } else {
    targets = args.filter((a) => !a.startsWith("--"));
  }

  if (!targets.length) {
    console.error("No client files to process. Use --list, --all, or pass a filename.");
    process.exit(1);
  }

  let failed = 0;
  for (const fileName of targets) {
    try {
      await processFile(fileName, jobsConfig, propsFallback);
    } catch (err) {
      failed++;
      console.error(`\nFAIL ${fileName}: ${err.message}`);
    }
  }

  console.log(`\nDone. ${targets.length - failed}/${targets.length} succeeded.`);
  if (failed) process.exitCode = 1;
}

if (require.main === module) {
  main();
}

module.exports = { processFile, listClientFiles, findKenyaFor };
