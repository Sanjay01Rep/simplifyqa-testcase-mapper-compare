const fs = require("fs");
const XLSX = require("xlsx");
const {
  parseClientWorkbook,
  buildOutputRows,
  exactNameKey,
  safeBaseName,
  previewFromRows,
  parseKenyaSource,
  applyKenyaPrereqs,
  isPlaceholderId,
} = require("./mapper");

function excelSheetName(fileName, used) {
  let base = safeBaseName(fileName || "Sheet")
    .replace(/[\[\]:*?/\\]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 31);
  if (!base) base = "Sheet";
  let name = base;
  let n = 2;
  while (used.has(name.toLowerCase())) {
    const suffix = ` ${n}`;
    name = (base.slice(0, 31 - suffix.length) + suffix).trim();
    n += 1;
  }
  used.add(name.toLowerCase());
  return name;
}

/** Soft key for step compare: ignore punctuation and common abbreviations (no/number, &/and). */
function compareTextKey(value) {
  let s = String(value || "").toLowerCase();
  s = s.replace(/&/g, " and ");
  s = s.replace(/\ball they way\b/g, "all the way");
  s = s.replace(/\b(no|num)\b/g, "number");
  s = s.replace(/\bph\b/g, "phone");
  s = s.replace(/\bvai\b/g, "via");
  s = s.replace(/\binoice\b/g, "invoice");
  s = s.replace(/\bdop\b/g, "drop");
  s = s.replace(/\bsuccesful\b/g, "successful");
  s = s.replace(/\breconcilaition\b/g, "reconciliation");
  s = s.replace(/\breconcilation\b/g, "reconciliation");
  s = s.replace(/\bm[\s-]?pesa\b/g, "mobile money");
  s = s.replace(/\bmobile\s*\/\s*bank\b/g, "mobile money bank");
  s = s.replace(/\bafter fully approval\b/g, "after full approval");
  s = s.replace(/([a-z])\s+(\d)/g, "$1$2");
  s = s.replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();
  return s;
}

function levenshtein(a, b) {
  const s = String(a || "");
  const t = String(b || "");
  if (s === t) return 0;
  if (!s.length) return t.length;
  if (!t.length) return s.length;
  const rows = s.length + 1;
  const cols = t.length + 1;
  const prev = new Array(cols);
  const cur = new Array(cols);
  for (let j = 0; j < cols; j++) prev[j] = j;
  for (let i = 1; i < rows; i++) {
    cur[0] = i;
    for (let j = 1; j < cols; j++) {
      const cost = s[i - 1] === t[j - 1] ? 0 : 1;
      cur[j] = Math.min(cur[j - 1] + 1, prev[j] + 1, prev[j - 1] + cost);
    }
    for (let j = 0; j < cols; j++) prev[j] = cur[j];
  }
  return prev[cols - 1];
}

function isAdjacentTransposition(a, b) {
  if (a.length !== b.length || a.length < 3) return false;
  let i = 0;
  while (i < a.length && a[i] === b[i]) i += 1;
  if (i >= a.length - 1) return false;
  if (a[i] === b[i + 1] && a[i + 1] === b[i] && a.slice(i + 2) === b.slice(i + 2)) return true;
  return false;
}

function stemWord(w) {
  if (w.length >= 5 && w.endsWith("ly")) return w.slice(0, -2);
  if (w.length >= 5 && w.endsWith("ing")) return w.slice(0, -3);
  if (w.length >= 5 && w.endsWith("ed")) return w.slice(0, -2);
  if (w.length >= 5 && w.endsWith("s") && !w.endsWith("ss")) return w.slice(0, -1);
  return w;
}

function wordsClose(a, b) {
  if (a === b) return true;
  if (stemWord(a) === stemWord(b) && stemWord(a).length >= 4) return true;
  if (isAdjacentTransposition(a, b)) return true;
  // Keep short tokens exact so KRA vs URA (and similar) stay different.
  if (Math.min(a.length, b.length) < 5) return false;
  const maxDist = Math.max(1, Math.floor(Math.min(a.length, b.length) * 0.25));
  return levenshtein(a, b) <= maxDist;
}

function sameWordBag(a, b) {
  const wa = a.split(" ").filter(Boolean).sort();
  const wb = b.split(" ").filter(Boolean).sort();
  if (wa.length !== wb.length || !wa.length) return false;
  return wa.every((w, i) => w === wb[i]);
}

function extraSuccessOnly(ka, kb) {
  const [shorter, longer] = ka.length <= kb.length ? [ka, kb] : [kb, ka];
  if (!shorter || !longer.startsWith(shorter)) return false;
  const rest = longer.slice(shorter.length).trim();
  return /^(successful|successfully|success|ok|done)$/.test(rest);
}

/** True when texts match after soft normalize, allowing minor misspellings. */
function softTextsEqual(a, b) {
  const ka = compareTextKey(a);
  const kb = compareTextKey(b);
  if (ka === kb) return true;
  if (!ka || !kb) return false;
  if (sameWordBag(ka, kb)) return true;
  if (extraSuccessOnly(ka, kb)) return true;
  const wa = ka.split(" ").filter(Boolean);
  const wb = kb.split(" ").filter(Boolean);
  if (wa.length === wb.length) {
    return wa.every((w, i) => wordsClose(w, wb[i]));
  }
  // Different word counts: allow only tiny overall edit distance on the full string.
  const maxLen = Math.max(ka.length, kb.length);
  if (maxLen < 12) return false;
  return levenshtein(ka, kb) <= Math.max(2, Math.floor(maxLen * 0.08));
}

function stepsEqual(a, b) {
  if (!a || !b || a.steps.length !== b.steps.length) return false;
  return a.steps.every((step, i) => {
    const other = b.steps[i];
    return softTextsEqual(step.stepDesc, other.stepDesc) && softTextsEqual(step.expected, other.expected);
  });
}

function realIdKey(tc) {
  const id = String(tc.clientId || "").trim();
  if (!id || isPlaceholderId(id)) return "";
  return exactNameKey(id);
}

function canPair(a, b) {
  const sameName =
    exactNameKey(a.name) === exactNameKey(b.name) || softTextsEqual(a.name, b.name);
  if (!sameName) return false;
  const idA = realIdKey(a);
  const idB = realIdKey(b);
  if (idA && idB) return idA === idB;
  return true;
}

function compareTestcases(listA, listB) {
  const usedB = new Set();
  const common = [];
  const unmatchedA = [];
  const unmatchedB = [];
  const nameMatchStepMismatch = [];

  listA.forEach((tcA, iA) => {
    const idxB = listB.findIndex((tcB, iB) => !usedB.has(iB) && canPair(tcA, tcB));
    if (idxB < 0) {
      unmatchedA.push(tcA);
      return;
    }
    usedB.add(idxB);
    const tcB = listB[idxB];
    if (stepsEqual(tcA, tcB)) {
      common.push({ fromA: tcA, fromB: tcB });
    } else {
      unmatchedA.push(tcA);
      unmatchedB.push(tcB);
      nameMatchStepMismatch.push({
        name: tcA.name,
        id: tcA.clientId || tcB.clientId || "",
        stepsA: tcA.steps.length,
        stepsB: tcB.steps.length,
      });
    }
  });

  listB.forEach((tcB, iB) => {
    if (!usedB.has(iB)) unmatchedB.push(tcB);
  });

  return { common, unmatchedA, unmatchedB, nameMatchStepMismatch };
}

function parseForCompare(buffer, fileName) {
  return parseClientWorkbook(buffer, fileName, {
    Module: "Compare",
    Entity: "Compare",
    Versions: "v1.0",
    TestcaseType: "WEB",
  });
}

function tagIssues(issues, label) {
  return (issues || []).map((issue) => ({
    ...issue,
    message: `[${label}] ${issue.message}`,
  }));
}

function tcBrief(tc) {
  return {
    name: tc.name,
    id: tc.clientId || "",
    steps: tc.steps.length,
    prerequisites: tc.prerequisites || "",
  };
}

function joinEntities(value) {
  if (Array.isArray(value)) {
    return [...new Set(value.map((v) => String(v || "").trim()).filter(Boolean))].join(", ");
  }
  return String(value || "")
    .split(/[,;]/)
    .map((v) => v.trim())
    .filter(Boolean)
    .filter((v, i, arr) => arr.indexOf(v) === i)
    .join(", ");
}

async function compareClientDocs({
  fileAName,
  bufferA,
  fileBName,
  bufferB,
  moduleName,
  entity,
  entityCommon,
  entityUniqueA,
  entityUniqueB,
  versions,
  testcaseType,
  kenyaFileName,
  kenyaBuffer,
}) {
  const parsedA = parseForCompare(bufferA, fileAName);
  const parsedB = parseForCompare(bufferB, fileBName);
  const result = compareTestcases(parsedA.testcases, parsedB.testcases);

  const fallbackEntity = joinEntities(entity);
  const commonEntity = joinEntities(entityCommon) || fallbackEntity;
  const uniqueAEntity = joinEntities(entityUniqueA) || fallbackEntity;
  const uniqueBEntity = joinEntities(entityUniqueB) || fallbackEntity;
  const baseConfig = {
    Module: moduleName,
    Versions: versions || "v1.0",
    TestcaseType: testcaseType || "WEB",
  };
  const commonTcs = result.common.map((pair) => pair.fromA);

  const issues = [
    ...tagIssues(parsedA.issues, fileAName),
    ...tagIssues(parsedB.issues, fileBName),
  ];
  if (result.common.length + result.unmatchedA.length !== parsedA.testcases.length) {
    issues.push({
      type: "COUNT_INVARIANT_A",
      severity: "ERROR",
      message: `Count mismatch File A: common (${result.common.length}) + unique A (${result.unmatchedA.length}) = ${
        result.common.length + result.unmatchedA.length
      }, but File A parsed ${parsedA.testcases.length}.`,
    });
  }
  if (result.common.length + result.unmatchedB.length !== parsedB.testcases.length) {
    issues.push({
      type: "COUNT_INVARIANT_B",
      severity: "ERROR",
      message: `Count mismatch File B: common (${result.common.length}) + unique B (${result.unmatchedB.length}) = ${
        result.common.length + result.unmatchedB.length
      }, but File B parsed ${parsedB.testcases.length}.`,
    });
  }
  result.nameMatchStepMismatch.forEach((row) => {
    issues.push({
      type: "STEPS_DIFFER",
      severity: "WARN",
      message: `Name/ID matched but steps differ (A ${row.stepsA} vs B ${row.stepsB}). Both unmatched.`,
      testcase: row.name,
    });
  });

  let kenyaStats = { matched: 0, kenyaCount: 0 };
  if (kenyaBuffer && kenyaFileName) {
    const kenya = await parseKenyaSource(kenyaFileName, kenyaBuffer);
    const applied = applyKenyaPrereqs(
      [...commonTcs, ...result.unmatchedA, ...result.unmatchedB],
      kenya
    );
    kenyaStats = { matched: applied.matched, kenyaCount: applied.kenyaCount };
    issues.push(...applied.issues);
  }

  const commonRows = buildOutputRows(commonTcs, { ...baseConfig, Entity: commonEntity });
  const rowsA = buildOutputRows(result.unmatchedA, { ...baseConfig, Entity: uniqueAEntity });
  const rowsB = buildOutputRows(result.unmatchedB, { ...baseConfig, Entity: uniqueBEntity });

  const used = new Set(["common"]);
  const sheetCommon = "Common";
  const sheetA = excelSheetName(fileAName, used);
  const sheetB = excelSheetName(fileBName, used);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(commonRows), sheetCommon);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rowsA), sheetA);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rowsB), sheetB);

  const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  const mappedTcCount = commonTcs.length + result.unmatchedA.length + result.unmatchedB.length;
  const mappedStepCount =
    commonTcs.reduce((n, tc) => n + tc.steps.length, 0) +
    result.unmatchedA.reduce((n, tc) => n + tc.steps.length, 0) +
    result.unmatchedB.reduce((n, tc) => n + tc.steps.length, 0);

  const summary = {
    fileA: fileAName,
    fileB: fileBName,
    kenyaFile: kenyaFileName || "",
    module: moduleName,
    entity: commonEntity,
    entityCommon: commonEntity,
    entityUniqueA: uniqueAEntity,
    entityUniqueB: uniqueBEntity,
    versions: baseConfig.Versions,
    testcaseType: baseConfig.TestcaseType,
    sheets: {
      common: sheetCommon,
      a: sheetA,
      b: sheetB,
    },
    counts: {
      a: parsedA.testcases.length,
      b: parsedB.testcases.length,
      common: result.common.length,
      unmatchedA: result.unmatchedA.length,
      unmatchedB: result.unmatchedB.length,
      nameMatchStepMismatch: result.nameMatchStepMismatch.length,
      aCheck: result.common.length + result.unmatchedA.length,
      bCheck: result.common.length + result.unmatchedB.length,
    },
    mappedTcCount,
    mappedStepCount,
    tcMatch: parsedA.tcMatch && parsedB.tcMatch,
    stepMatch: parsedA.stepMatch && parsedB.stepMatch,
    seqOk: parsedA.seqOk && parsedB.seqOk,
    kenyaMatched: kenyaStats.matched,
    kenyaCount: kenyaStats.kenyaCount,
    issues,
    nameMatchStepMismatch: result.nameMatchStepMismatch,
    commonNames: result.common.map((p) => tcBrief(p.fromA)),
    unmatchedA: result.unmatchedA.map(tcBrief),
    unmatchedB: result.unmatchedB.map(tcBrief),
  };

  return {
    buffer,
    summary,
    preview: previewFromRows(commonRows, 80),
  };
}

function writeCompareLog(logPath, summary) {
  const lines = [];
  lines.push("=".repeat(72));
  lines.push("SimplifyQA Client Compare Log");
  lines.push(`Generated at : ${new Date().toISOString()}`);
  lines.push(`File A       : ${summary.fileA}`);
  lines.push(`File B       : ${summary.fileB}`);
  lines.push(`Kenya file   : ${summary.kenyaFile || "(none)"}`);
  lines.push(`Module       : ${summary.module}`);
  lines.push(`Entity Common: ${summary.entityCommon || summary.entity || ""}`);
  lines.push(`Entity Unique A : ${summary.entityUniqueA || ""}`);
  lines.push(`Entity Unique B : ${summary.entityUniqueB || ""}`);
  lines.push(`Versions     : ${summary.versions}`);
  lines.push(`TestcaseType : ${summary.testcaseType}`);
  lines.push(`Output Excel : ${summary.outName || ""}`);
  lines.push(`Kenya prereq : ${summary.kenyaMatched || 0}/${summary.kenyaCount || 0} matched`);
  lines.push("=".repeat(72));
  lines.push("");
  lines.push("--- COUNTS ---");
  lines.push(`File A testcases : ${summary.counts.a}`);
  lines.push(`File B testcases : ${summary.counts.b}`);
  lines.push(`Common (same name/id AND same steps) : ${summary.counts.common}`);
  lines.push(`Unmatched A (${summary.sheets.a}) : ${summary.counts.unmatchedA}`);
  lines.push(`Unmatched B (${summary.sheets.b}) : ${summary.counts.unmatchedB}`);
  lines.push(
    `Check A : common + unmatched A = ${summary.counts.common + summary.counts.unmatchedA} (file A ${summary.counts.a})`
  );
  lines.push(
    `Check B : common + unmatched B = ${summary.counts.common + summary.counts.unmatchedB} (file B ${summary.counts.b})`
  );
  lines.push(
    `Name matched but steps differ (both unmatched) : ${summary.counts.nameMatchStepMismatch}`
  );
  lines.push("");
  lines.push("--- COMMON ---");
  if (!summary.commonNames.length) {
    lines.push("(none)");
  } else {
    summary.commonNames.forEach((tc, i) => {
      lines.push(`${i + 1}. ${tc.id || "(no-id)"} | ${tc.name} | steps=${tc.steps}`);
    });
  }
  lines.push("");
  lines.push("--- NAME MATCH / STEPS DIFFER ---");
  if (!summary.nameMatchStepMismatch.length) {
    lines.push("(none)");
  } else {
    summary.nameMatchStepMismatch.forEach((row, i) => {
      lines.push(
        `${i + 1}. ${row.id || "(no-id)"} | ${row.name} | A steps=${row.stepsA} B steps=${row.stepsB}`
      );
    });
  }
  lines.push("");
  lines.push(`--- UNMATCHED A (${summary.sheets.a}) ---`);
  if (!summary.unmatchedA.length) {
    lines.push("(none)");
  } else {
    summary.unmatchedA.forEach((tc, i) => {
      lines.push(`${i + 1}. ${tc.id || "(no-id)"} | ${tc.name} | steps=${tc.steps}`);
    });
  }
  lines.push("");
  lines.push(`--- UNMATCHED B (${summary.sheets.b}) ---`);
  if (!summary.unmatchedB.length) {
    lines.push("(none)");
  } else {
    summary.unmatchedB.forEach((tc, i) => {
      lines.push(`${i + 1}. ${tc.id || "(no-id)"} | ${tc.name} | steps=${tc.steps}`);
    });
  }
  lines.push("");
  lines.push("--- ISSUES ---");
  if (!(summary.issues || []).length) {
    lines.push("(none)");
  } else {
    summary.issues.forEach((issue, i) => {
      lines.push(
        `${i + 1}. [${issue.severity}] ${issue.type} | ${issue.testcase || ""} | ${issue.message}`
      );
    });
  }
  lines.push("");
  fs.writeFileSync(logPath, lines.join("\n"), "utf8");
  return logPath;
}

module.exports = {
  compareClientDocs,
  compareTestcases,
  canPair,
  stepsEqual,
  compareTextKey,
  softTextsEqual,
  excelSheetName,
  writeCompareLog,
  joinEntities,
};
