const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const SIMPLIFYQA_HEADER = [
  "Name*",
  "Versions*",
  "Pre-Requisites",
  "Description",
  "Module*",
  "Labels",
  "Testcase Type*",
  "User Story",
  "Defects",
  "Seq*",
  "Step Description*",
  "Expected Result",
  "Parameter",
  "Mandate Screenshot (Type Yes for required steps)",
  "Entity",
];

function cell(row, idx) {
  if (row == null || row[idx] == null) return "";
  return String(row[idx]).trim();
}

function normalizeText(value) {
  return String(value || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .trim();
}

function todayStamp() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function nowStamp() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  const hh = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  const ss = String(d.getSeconds()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}_${hh}-${mi}-${ss}`;
}

function safeBaseName(fileName) {
  return path
    .basename(fileName, path.extname(fileName))
    .replace(/[<>:"/\\|?*]/g, "_");
}

function isLockFile(name) {
  return String(name || "").startsWith("~$");
}

function loadProperties(filePath) {
  if (!fs.existsSync(filePath)) return {};
  const text = fs.readFileSync(filePath, "utf8");
  const props = {};
  for (const line of text.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const idx = trimmed.indexOf("=");
    if (idx === -1) continue;
    props[trimmed.slice(0, idx).trim()] = trimmed.slice(idx + 1).trim();
  }
  return props;
}

function stringifyProperties(props) {
  const keys = Object.keys(props);
  const lines = [
    "# SimplifyQA Testcase Mapping Configuration",
    "# Update Module and Entity as needed.",
    "",
  ];
  for (const key of keys) {
    if (key.startsWith("#")) continue;
    lines.push(`${key}=${props[key]}`);
  }
  return lines.join("\n") + "\n";
}

function loadJobs(filePath) {
  if (!fs.existsSync(filePath)) return { defaults: {}, jobs: [] };
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function resolveJobConfig(fileName, jobsConfig, propsFallback) {
  const defaults = (jobsConfig && jobsConfig.defaults) || {};
  const job = ((jobsConfig && jobsConfig.jobs) || []).find(
    (j) => j.file.toLowerCase() === String(fileName || "").toLowerCase()
  );

  return {
    Module: (job && job.Module) || defaults.Module || propsFallback.Module || "",
    Entity: (job && job.Entity) || defaults.Entity || propsFallback.Entity || "",
    Versions:
      (job && job.Versions) || defaults.Versions || propsFallback.Versions || "v1.0",
    TestcaseType:
      (job && job.TestcaseType) ||
      defaults.TestcaseType ||
      propsFallback.TestcaseType ||
      "WEB",
    fromJobs: Boolean(job),
  };
}

function normalizeName(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[_/\\]+/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isPlaceholderId(value) {
  const v = String(value || "").trim().toLowerCase();
  return v === "new" || v === "na" || v === "n/a" || v === "-" || v === "new test";
}

function isHeaderEchoRow(id, name, seqRaw, stepDesc) {
  const i = exactNameKey(id);
  const n = exactNameKey(name);
  const seq = exactNameKey(seqRaw);
  const step = exactNameKey(stepDesc);
  if (i === "id" && (n === "name" || n === "testcase name" || n === "test case name" || !n)) return true;
  if (n === "name" && (seq === "seq" || seq === "sequence" || step === "step description")) return true;
  return false;
}

function isSectionMarkerRow(id, name, seqRaw, stepDesc, expected, description) {
  if (name || stepDesc || expected || description || String(seqRaw || "").trim() !== "") return false;
  const i = exactNameKey(id);
  if (!i || /\d/.test(String(id || ""))) return false;
  return i.length <= 12;
}

const JUNK_TC_LABELS = new Set([
  "ug",
  "name",
  "id",
  "seq",
  "sequence",
  "module",
  "entity",
  "description",
  "step description",
  "expected result",
  "expected",
]);

function isJunkLabel(value) {
  return JUNK_TC_LABELS.has(exactNameKey(value));
}

function isJunkTestcase(tc) {
  const name = exactNameKey(tc.name);
  const id = exactNameKey(tc.clientId);
  // Exclude header/section leftovers like Name, UG, ID regardless of steps.
  if (isJunkLabel(name) || isJunkLabel(id)) return true;
  if (!name && isJunkLabel(id)) return true;
  return false;
}

function stepContentKey(step) {
  return `${exactNameKey(step.stepDesc)}|${exactNameKey(step.expected)}`;
}

function tcIdentityKey(tc) {
  const name = exactNameKey(tc.name);
  const id = String(tc.clientId || "").trim();
  if (id && !isPlaceholderId(id)) return `${name}|${exactNameKey(id)}`;
  return `${name}|`;
}

function mergeDuplicateTestcases(testcases, issues) {
  const kept = [];
  const byKey = new Map();
  for (const tc of testcases) {
    if (isJunkTestcase(tc)) {
      issues.push({
        type: "SKIPPED_JUNK_TC",
        severity: "INFO",
        message: `Skipped header/section row "${tc.clientId || ""} ${tc.name}".`,
        testcase: tc.name,
      });
      continue;
    }
    const key = tcIdentityKey(tc);
    if (!exactNameKey(tc.name)) {
      kept.push(tc);
      continue;
    }
    const prev = byKey.get(key);
    if (!prev) {
      byKey.set(key, tc);
      kept.push(tc);
      continue;
    }
    const prevHasRealId = Boolean(prev.clientId && !isPlaceholderId(prev.clientId));
    const nextHasRealId = Boolean(tc.clientId && !isPlaceholderId(tc.clientId));
    if (!prevHasRealId && nextHasRealId) prev.clientId = tc.clientId;
    if (!prev.description && tc.description) prev.description = tc.description;
    if (!prev.prerequisites && tc.prerequisites) prev.prerequisites = tc.prerequisites;
    const have = new Set(prev.steps.map(stepContentKey));
    let added = 0;
    for (const step of tc.steps) {
      const k = stepContentKey(step);
      if (k === "|" || have.has(k)) continue;
      prev.steps.push(step);
      have.add(k);
      added += 1;
    }
    issues.push({
      type: "DUP_TC_MERGED",
      severity: "INFO",
      message: `Merged duplicate testcase "${tc.name}" (${tc.clientId || "no-id"}) from a second sheet${
        added ? `, added ${added} extra step(s)` : ""
      }.`,
      testcase: tc.name,
    });
  }
  return kept;
}

function detectHeaderMap(headerRow) {
  const normalize = (v) => String(v || "").trim().toLowerCase();
  const map = {};
  headerRow.forEach((h, i) => {
    const key = normalize(h);
    if (key === "id") map.id = i;
    else if (key === "name" || key === "testcase name" || key === "test case name" || key === "test case")
      map.name = i;
    else if (key === "module") map.module = i;
    else if (key === "entity") map.entity = i;
    else if (key === "description") map.description = i;
    else if (key === "seq" || key === "sequence") map.seq = i;
    else if (key === "step description" || key === "stepdescription" || key === "steps")
      map.step = i;
    else if (key === "expected result" || key === "expectedresult" || key === "expected")
      map.expected = i;
    else if (
      key.includes("pre-requisite") ||
      key.includes("prerequisite") ||
      key.includes("pre requisite") ||
      key.includes("pre-condition") ||
      key.includes("precondition")
    ) {
      map.prereq = i;
    }
  });
  return map;
}

function workbookFromSource(source) {
  if (Buffer.isBuffer(source) || source instanceof Uint8Array) {
    return XLSX.read(source, { type: "buffer" });
  }
  if (typeof source === "string") {
    return XLSX.readFile(source);
  }
  throw new Error("Unsupported workbook source.");
}

function listWorkbookSheets(source) {
  const wb = workbookFromSource(source);
  return (wb.SheetNames || []).map((n) => String(n || "").trim()).filter(Boolean);
}

function resolveWorkbookSheets(allSheetNames, requestedNames) {
  const byTrimmed = new Map();
  for (const raw of allSheetNames || []) {
    const key = String(raw || "").trim();
    if (key && !byTrimmed.has(key)) byTrimmed.set(key, raw);
  }
  const selected = [];
  const missing = [];
  for (const name of requestedNames || []) {
    const key = String(name || "").trim();
    if (!key) continue;
    if (byTrimmed.has(key)) selected.push(byTrimmed.get(key));
    else missing.push(key);
  }
  return { selected, missing };
}

function parseClientWorkbook(source, clientFileName, config) {
  const issues = [];
  if (!config.Module || !config.Entity) {
    throw new Error("Module and Entity are required.");
  }

  const clientWb = workbookFromSource(source);
  const allSheetNames = clientWb.SheetNames || [];
  if (!allSheetNames.length) {
    throw new Error(`Empty workbook: ${clientFileName}`);
  }
  const requested = Array.isArray(config.clientSheets)
    ? config.clientSheets.map((s) => String(s || "").trim()).filter(Boolean)
    : String(config.clientSheet || "").trim()
      ? [String(config.clientSheet).trim()]
      : [];
  let selectedSheets;
  if (requested.length) {
    const resolved = resolveWorkbookSheets(allSheetNames, requested);
    selectedSheets = resolved.selected;
    if (!selectedSheets.length) {
      const available = allSheetNames.map((n) => String(n || "").trim()).filter(Boolean);
      throw new Error(
        `Sheet "${requested.join(", ")}" was not found. Available: ${available.join(", ")}.`
      );
    }
    if (allSheetNames.length > 1) {
      issues.push({
        type: "SHEET_SELECTED",
        severity: "INFO",
        message: `Mapped from sheet "${selectedSheets
          .map((n) => String(n).trim())
          .join(", ")}" (${allSheetNames.length} sheets in file).`,
      });
    }
  } else {
    selectedSheets = allSheetNames.length === 2 ? allSheetNames.slice(0, 2) : [allSheetNames[0]];
    if (selectedSheets.length === 2) {
      issues.push({
        type: "SHEETS_AUTO_SELECTED",
        severity: "INFO",
        message: `Workbook has 2 sheets. Auto-selected both sheets for mapping: ${selectedSheets
          .map((n) => String(n).trim())
          .join(", ")}.`,
      });
    }
  }
  const testcases = [];
  let namedTcRows = 0;
  let inferredTcCount = 0;
  let clientStepLikeCount = 0;
  let hadAnyRows = false;

  for (const sheetName of selectedSheets) {
    const clientRows = XLSX.utils.sheet_to_json(clientWb.Sheets[sheetName], {
      header: 1,
      defval: "",
    });
    if (clientRows.length === 0) continue;
    hadAnyRows = true;

    const headerMap = detectHeaderMap(clientRows[0]);
    const requiredKeys = ["name", "seq", "step", "expected"];
    const missingHeaders = requiredKeys.filter((k) => headerMap[k] === undefined);
    if (missingHeaders.length) {
      issues.push({
        type: "HEADER",
        severity: "WARN",
        message: `Sheet "${sheetName}": Could not auto-detect columns: ${missingHeaders.join(
          ", "
        )}. Falling back to positional layout (Name=1, Desc=4, Seq=5, Step=6, Expected=7).`,
      });
      Object.assign(headerMap, {
        id: headerMap.id ?? 0,
        name: headerMap.name ?? 1,
        module: headerMap.module ?? 2,
        entity: headerMap.entity ?? 3,
        description: headerMap.description ?? 4,
        seq: headerMap.seq ?? 5,
        step: headerMap.step ?? 6,
        expected: headerMap.expected ?? 7,
      });
    }

    const dataRows = clientRows.slice(1).filter((r) =>
      r.some((c) => String(c || "").trim() !== "")
    );
    namedTcRows += dataRows.filter((r) => cell(r, headerMap.name) || cell(r, headerMap.id ?? 0)).length;
    clientStepLikeCount += dataRows.filter((r) => {
      const stepDesc = normalizeText(cell(r, headerMap.step));
      const expected = normalizeText(cell(r, headerMap.expected));
      const seqRaw = cell(r, headerMap.seq);
      return Boolean(stepDesc || expected || seqRaw);
    }).length;

    let current = null;
    for (const row of dataRows) {
      const id = cell(row, headerMap.id ?? 0);
      const hasRealId = Boolean(id && !isPlaceholderId(id));
      const name = cell(row, headerMap.name);
      const description = normalizeText(cell(row, headerMap.description));
      const seqRaw = cell(row, headerMap.seq);
      let stepDesc = normalizeText(cell(row, headerMap.step));
      let expected = normalizeText(cell(row, headerMap.expected));
      const inlinePrereq =
        headerMap.prereq != null ? normalizeText(cell(row, headerMap.prereq)) : "";

      if (isHeaderEchoRow(id, name, seqRaw, stepDesc)) {
        issues.push({
          type: "SKIPPED_HEADER_ROW",
          severity: "INFO",
          message: `Skipped repeated header row in sheet "${sheetName}".`,
        });
        continue;
      }
      if (isSectionMarkerRow(id, name, seqRaw, stepDesc, expected, description)) {
        continue;
      }
      if ((name || id) && (isJunkLabel(name) || (!name && isJunkLabel(id)))) {
        issues.push({
          type: "SKIPPED_JUNK_TC",
          severity: "INFO",
          message: `Skipped junk testcase label "${id || ""} ${name || ""}" in sheet "${sheetName}".`,
        });
        continue;
      }

      let inferredNewFromSeqRestart = false;
      const seqNumPreview = seqRaw === "" ? null : Number(seqRaw);
      if (!id && !name && current && current.steps.length > 0 && seqNumPreview === 1) {
        inferredNewFromSeqRestart = true;
      }

      const sameNameAsCurrent =
        current && name && exactNameKey(name) === exactNameKey(current.name);
      const sameIdAsCurrent =
        current &&
        hasRealId &&
        current.clientId &&
        !isPlaceholderId(current.clientId) &&
        exactNameKey(id) === exactNameKey(current.clientId);
      const startsNew = Boolean(
        inferredNewFromSeqRestart ||
          ((hasRealId || name) && !sameNameAsCurrent && !sameIdAsCurrent)
      );

      if (startsNew) {
        if (inferredNewFromSeqRestart) {
          issues.push({
            type: "INFERRED_NEW_TC",
            severity: "FIX",
            message: `Inferred new testcase from seq restart at 1 with empty Name/ID in "${sheetName}". Using Description as Name: "${description || "(empty)"}".`,
            testcase: description || "Untitled",
          });
        }
        if ((hasRealId || name) && !name) {
          issues.push({
            type: "MISSING_NAME",
            severity: "WARN",
            message: `New testcase started with ID "${id}" but Name is empty in sheet "${sheetName}".`,
            clientId: id,
          });
        }
        const resolvedName =
          name || (inferredNewFromSeqRestart ? description : "") || id || "Untitled";
        current = {
          clientId: id,
          name: resolvedName,
          description: inferredNewFromSeqRestart ? "" : description,
          prerequisites: inlinePrereq,
          steps: [],
          seqFixes: [],
        };
        if (!id) {
          issues.push({
            type: "MISSING_ID",
            severity: "INFO",
            message: `Testcase "${current.name}" has no Client ID (ignored in template).`,
            testcase: current.name,
          });
        }
        if (!description && !inferredNewFromSeqRestart) {
          issues.push({
            type: "EMPTY_DESCRIPTION",
            severity: "INFO",
            message: `Testcase "${current.name}" has empty Description.`,
            testcase: current.name,
          });
        }
        testcases.push(current);
      } else if (id && isPlaceholderId(id) && current) {
        issues.push({
          type: "PLACEHOLDER_ID_IGNORED",
          severity: "INFO",
          message: `Ignored placeholder ID "${id}" in sheet "${sheetName}" and continued testcase "${current.name}".`,
          testcase: current.name,
        });
      } else if (current && inlinePrereq && !current.prerequisites) {
        current.prerequisites = inlinePrereq;
      }

      if (!current) {
        issues.push({
          type: "ORPHAN_ROW",
          severity: "WARN",
          message: `Row content found before any testcase header in "${sheetName}". Creating Untitled TC.`,
        });
        current = {
          clientId: "",
          name: "Untitled",
          description: "",
          prerequisites: inlinePrereq,
          steps: [],
          seqFixes: [],
        };
        testcases.push(current);
      }

      const hadEmptyStep = !stepDesc && Boolean(expected);
      if (hadEmptyStep) {
        stepDesc = expected;
        issues.push({
          type: "EMPTY_STEP_FILLED",
          severity: "FIX",
          message: "Empty Step Description filled from Expected Result.",
          testcase: current.name,
          expectedPreview: expected.slice(0, 120),
        });
      }

      if (!stepDesc && !expected && !seqRaw) {
        issues.push({
          type: "SKIPPED_EMPTY",
          severity: "INFO",
          message: `Skipped empty content row under "${current.name}".`,
          testcase: current.name,
        });
        continue;
      }

      const hadEmptyExpected = !expected && Boolean(stepDesc);
      if (hadEmptyExpected) {
        expected = stepDesc;
        issues.push({
          type: "EMPTY_EXPECTED_FILLED",
          severity: "FIX",
          message: "Empty Expected Result filled from Step Description.",
          testcase: current.name,
          stepPreview: stepDesc.slice(0, 120),
        });
      }

      let originalSeq = null;
      if (seqRaw === "") {
        issues.push({
          type: "MISSING_SEQ",
          severity: "FIX",
          message: `Missing sequence number under "${current.name}" — will auto-assign.`,
          testcase: current.name,
        });
      } else {
        const n = Number(seqRaw);
        if (Number.isNaN(n)) {
          issues.push({
            type: "INVALID_SEQ",
            severity: "FIX",
            message: `Invalid sequence "${seqRaw}" under "${current.name}" — will auto-assign.`,
            testcase: current.name,
          });
        } else {
          originalSeq = n;
        }
      }

      current.steps.push({
        originalSeq,
        stepDesc,
        expected,
        filledFromExpected: hadEmptyStep,
        filledFromStep: hadEmptyExpected,
      });
    }
  }

  if (!hadAnyRows) {
    throw new Error(`Empty workbook: ${clientFileName}`);
  }

  const merged = mergeDuplicateTestcases(testcases, issues);
  testcases.length = 0;
  testcases.push(...merged);

  for (const tc of testcases) {
    if (tc.steps.length === 0) {
      issues.push({
        type: "NO_STEPS",
        severity: "ERROR",
        message: `Testcase "${tc.name}" has no steps.`,
        testcase: tc.name,
      });
      continue;
    }

    const originals = tc.steps.map((s) => s.originalSeq);
    const present = originals.filter((s) => s !== null);
    const uniqueSorted = [...new Set(present)].sort((a, b) => a - b);
    if (uniqueSorted.length > 0) {
      const min = uniqueSorted[0];
      const max = uniqueSorted[uniqueSorted.length - 1];
      const missing = [];
      for (let s = min; s <= max; s++) {
        if (!uniqueSorted.includes(s)) missing.push(s);
      }
      if (missing.length) {
        issues.push({
          type: "SEQ_GAP",
          severity: "FIX",
          message: `Sequence gap(s) in "${tc.name}": missing ${missing.join(", ")} (original range ${min}-${max}). Renumbered 1..${tc.steps.length}.`,
          testcase: tc.name,
          missing,
        });
      }
      if (min !== 1) {
        issues.push({
          type: "SEQ_NOT_START_1",
          severity: "FIX",
          message: `Sequences in "${tc.name}" start at ${min} instead of 1. Renumbered.`,
          testcase: tc.name,
        });
      }
      const seen = new Set();
      for (const s of present) {
        if (seen.has(s)) {
          issues.push({
            type: "SEQ_DUPLICATE",
            severity: "FIX",
            message: `Duplicate original sequence ${s} in "${tc.name}". Renumbered.`,
            testcase: tc.name,
          });
        }
        seen.add(s);
      }
    }

    tc.steps.forEach((s, i) => {
      const newSeq = i + 1;
      if (s.originalSeq !== null && s.originalSeq !== newSeq) {
        tc.seqFixes.push({ from: s.originalSeq, to: newSeq });
      }
      s.seq = newSeq;
    });

    if (tc.seqFixes.length) {
      issues.push({
        type: "SEQ_RENUMBERED",
        severity: "FIX",
        message: `Renumbered sequences for "${tc.name}": ${tc.seqFixes
          .map((f) => `${f.from}->${f.to}`)
          .join(", ")}`,
        testcase: tc.name,
      });
    }
  }

  const mappedTcCount = testcases.length;
  const mappedStepCount = testcases.reduce((n, tc) => n + tc.steps.length, 0);
  inferredTcCount = issues.filter((i) => i.type === "INFERRED_NEW_TC").length;
  const uniqueIdentityCount = new Set(testcases.map(tcIdentityKey)).size;
  const clientTcCount = uniqueIdentityCount;

  return {
    clientFile: clientFileName,
    sheetName: selectedSheets.map((n) => String(n).trim()).join(", "),
    sheetNames: selectedSheets.map((n) => String(n).trim()),
    testcases,
    issues,
    namedTcRows,
    inferredTcCount,
    clientTcCount,
    mappedTcCount,
    clientStepLikeCount,
    mappedStepCount,
    tcMatch: clientTcCount === mappedTcCount,
    stepMatch: clientStepLikeCount === mappedStepCount,
    seqOk: testcases.every((tc) => tc.steps.every((s, i) => s.seq === i + 1)),
    mandateOk: testcases.every((tc) => tc.steps.length > 0),
  };
}

function isNameLikeHeader(key) {
  const k = String(key || "").toLowerCase();
  return (
    k === "name" ||
    k === "testcase" ||
    k === "test case" ||
    k.includes("test case name") ||
    k.includes("testcase name") ||
    k === "scenario" ||
    k === "title"
  );
}

function isPrereqLikeHeader(key) {
  const k = String(key || "").toLowerCase();
  return (
    k.includes("pre-requisite") ||
    k.includes("prerequisite") ||
    k.includes("pre requisite") ||
    k.includes("pre-condition") ||
    k.includes("precondition") ||
    k.includes("pre condition")
  );
}

function headerCellText(value) {
  return String(value || "").trim();
}

function findHeaderColumnIndex(headerRow, wanted) {
  const want = headerCellText(wanted).toLowerCase();
  if (!want) return -1;
  return (headerRow || []).findIndex((h) => headerCellText(h).toLowerCase() === want);
}

/** Read headers from a mapper Excel for the UI dropdown. */
function listMapperHeaders(source, fileName) {
  const wb = workbookFromSource(source);
  const issues = [];
  for (const sheet of wb.SheetNames || []) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], {
      header: 1,
      defval: "",
    });
    if (!rows.length) continue;

    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      const row = rows[i] || [];
      const headers = [];
      const seen = new Set();
      row.forEach((cellValue, idx) => {
        const name = headerCellText(cellValue);
        if (!name) return;
        const key = name.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        headers.push({ index: idx, name });
      });
      if (headers.length < 2) continue;

      const keys = headers.map((h) => h.name.toLowerCase());
      const hasName = keys.some(isNameLikeHeader) || detectHeaderMap(row).name != null;
      if (!hasName && !keys.some(isPrereqLikeHeader)) continue;

      const prereqHeader = headers.find((h) => isPrereqLikeHeader(h.name));
      return {
        ok: true,
        fileName: fileName || "",
        sheet,
        headerRow: i,
        headers: headers.map((h) => h.name),
        defaultHeader: prereqHeader ? prereqHeader.name : "",
        defaultLabel: prereqHeader
          ? prereqHeader.name
          : "Pre-Requisite (auto-detect)",
      };
    }
  }

  issues.push({
    type: "MAPPER_HEADER",
    severity: "WARN",
    message: `No usable header row found in mapper file "${fileName || ""}".`,
  });
  return {
    ok: false,
    fileName: fileName || "",
    sheet: "",
    headerRow: -1,
    headers: [],
    defaultHeader: "",
    defaultLabel: "Pre-Requisite (auto-detect)",
    issues,
  };
}

function resolveMapperHeaderRow(rows, prereqHeader) {
  let headerIdx = 0;
  let map = detectHeaderMap(rows[0] || []);
  const wanted = headerCellText(prereqHeader);

  const tryRow = (i) => {
    const trial = detectHeaderMap(rows[i] || []);
    const keys = (rows[i] || []).map((c) => String(c || "").trim().toLowerCase());
    if (trial.name == null) {
      const ni = keys.findIndex(isNameLikeHeader);
      if (ni >= 0) trial.name = ni;
    }
    if (wanted) {
      const pi = findHeaderColumnIndex(rows[i] || [], wanted);
      if (pi >= 0) trial.prereq = pi;
    } else if (trial.prereq == null) {
      const pi = keys.findIndex(isPrereqLikeHeader);
      if (pi >= 0) trial.prereq = pi;
    }
    return trial;
  };

  if (map.name == null || map.prereq == null || wanted) {
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      const trial = tryRow(i);
      if (trial.name != null && trial.prereq != null) {
        headerIdx = i;
        map = trial;
        break;
      }
    }
  } else {
    map = tryRow(0);
  }

  return { headerIdx, map };
}

function parseKenyaExcel(source, fileName, options = {}) {
  const wb = workbookFromSource(source);
  const entries = [];
  const issues = [];
  const prereqHeader = headerCellText(options.prereqHeader);

  for (const sheet of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], {
      header: 1,
      defval: "",
    });
    if (!rows.length) continue;

    const { headerIdx, map } = resolveMapperHeaderRow(rows, prereqHeader);

    if (map.name == null || map.prereq == null) {
      issues.push({
        type: "KENYA_HEADER",
        severity: "WARN",
        message: prereqHeader
          ? `Mapper sheet "${sheet}" has no Name column or selected header "${prereqHeader}".`
          : `Mapper sheet "${sheet}" has no Name + Pre-Requisite columns detected.`,
      });
      continue;
    }

    for (const row of rows.slice(headerIdx + 1)) {
      const name = cell(row, map.name);
      const prereq = normalizeText(cell(row, map.prereq));
      // Empty pre-requisite is ignored (not an issue)
      if (!name || !prereq) continue;
      entries.push({ name, prereq, sheet });
    }
  }

  if (!entries.length) {
    issues.push({
      type: "KENYA_EMPTY",
      severity: "INFO",
      message: prereqHeader
        ? `No filled Name + "${prereqHeader}" pairs in mapper file "${fileName}". Empty values are ignored.`
        : `No filled Name + Pre-Requisite pairs in mapper file "${fileName}". Empty pre-requisites are ignored.`,
    });
  }

  return {
    entries,
    issues,
    sourceType: "excel",
    fileName,
    prereqHeader: prereqHeader || "",
  };
}

function isXlsxFileName(fileName) {
  const base = path.basename(fileName || "");
  return /\.xlsx$/i.test(base) && !base.startsWith("~$");
}

function assertXlsx(fileName, label) {
  if (!isXlsxFileName(fileName)) {
    throw new Error(`${label} must be a .xlsx file. Other formats are not supported.`);
  }
}

function loadTemplateHeader() {
  const samplePath = path.join(__dirname, "..", "Sample File.xlsx");
  if (!fs.existsSync(samplePath)) return SIMPLIFYQA_HEADER.slice();
  const wb = XLSX.readFile(samplePath);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
    header: 1,
    defval: "",
  });
  const header = (rows[0] || []).map((c) => String(c || "").trim()).filter(Boolean);
  return header.length ? header : SIMPLIFYQA_HEADER.slice();
}

async function parseKenyaSource(fileName, buffer, options = {}) {
  assertXlsx(fileName, "Mapper file");
  return parseKenyaExcel(buffer, fileName, options);
}

function exactNameKey(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function findPrereqForName(tcName, kenya) {
  const target = exactNameKey(tcName);
  if (!target) return { prereq: "", how: null };
  const exact = kenya.entries.find((e) => exactNameKey(e.name) === target);
  if (exact && exact.prereq) return { prereq: exact.prereq, how: "exact" };
  return { prereq: "", how: null };
}

function applyKenyaPrereqs(testcases, kenya) {
  const issues = [...(kenya.issues || [])];
  let matched = 0;
  const usedKenya = new Set();

  for (const tc of testcases) {
    const found = findPrereqForName(tc.name, kenya);
    if (!found.prereq) continue; // empty / no exact match → ignore
    tc.prerequisites = found.prereq;
    tc.prereqSource = found.how;
    matched += 1;
    usedKenya.add(exactNameKey(tc.name));
    issues.push({
      type: "PREREQ_MATCHED",
      severity: "FIX",
      message: `Pre-Requisite mapped from mapper file (exact name) for "${tc.name}".`,
      testcase: tc.name,
    });
  }

  for (const entry of kenya.entries) {
    const key = exactNameKey(entry.name);
    if (!usedKenya.has(key)) {
      issues.push({
        type: "PREREQ_ORPHAN_KENYA",
        severity: "INFO",
        message: `Mapper Pre-Requisite not used (no exact client testcase name match): "${entry.name}".`,
        testcase: entry.name,
      });
    }
  }

  return { issues, matched, kenyaCount: kenya.entries.length };
}

function blankCell(value) {
  if (value == null) return undefined;
  if (typeof value === "string" && value.trim() === "") return undefined;
  return value;
}

function aoaToSheetOmitEmpty(aoa) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (const addr of Object.keys(ws)) {
    if (addr[0] === "!") continue;
    const cell = ws[addr];
    if (!cell || cell.v == null || cell.v === "") delete ws[addr];
  }
  return ws;
}

function buildOutputRows(testcases, config) {
  const outRows = [loadTemplateHeader()];
  for (const tc of testcases) {
    const lastIdx = tc.steps.length - 1;
    tc.steps.forEach((step, idx) => {
      const isFirst = idx === 0;
      const isLast = idx === lastIdx;
      outRows.push([
        blankCell(isFirst ? tc.name : undefined),
        blankCell(isFirst ? config.Versions : undefined),
        blankCell(isFirst ? tc.prerequisites : undefined),
        blankCell(isFirst ? tc.description : undefined),
        blankCell(isFirst ? config.Module : undefined),
        undefined,
        blankCell(isFirst ? config.TestcaseType : undefined),
        undefined,
        undefined,
        blankCell(step.seq),
        blankCell(step.stepDesc),
        blankCell(step.expected),
        undefined,
        blankCell(isLast ? "Yes" : undefined),
        blankCell(isFirst ? config.Entity : undefined),
      ]);
    });
  }
  return outRows;
}

function writeWorkbook(outRows, outPath) {
  const ws = aoaToSheetOmitEmpty(outRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, outPath);
  return outPath;
}

function workbookBuffer(outRows) {
  const ws = aoaToSheetOmitEmpty(outRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

function issueCounts(issues) {
  return issues.reduce((acc, i) => {
    acc[i.type] = (acc[i.type] || 0) + 1;
    return acc;
  }, {});
}

function buildSummary({ parsed, config, kenyaStats, outName, logName, issues }) {
  return {
    clientFile: parsed.clientFile,
    sheetName: parsed.sheetName,
    module: config.Module,
    entity: config.Entity,
    versions: config.Versions,
    testcaseType: config.TestcaseType,
    outName: outName || "",
    logName: logName || "",
    namedTcRows: parsed.namedTcRows,
    inferredTcCount: parsed.inferredTcCount,
    clientTcCount: parsed.clientTcCount,
    mappedTcCount: parsed.mappedTcCount,
    clientStepLikeCount: parsed.clientStepLikeCount,
    mappedStepCount: parsed.mappedStepCount,
    tcMatch: parsed.tcMatch,
    stepMatch: parsed.stepMatch,
    seqOk: parsed.seqOk,
    mandateOk: parsed.mandateOk,
    kenyaMatched: kenyaStats ? kenyaStats.matched : 0,
    kenyaCount: kenyaStats ? kenyaStats.kenyaCount : 0,
    issueCounts: issueCounts(issues),
    issues,
    perTc: parsed.testcases.map((tc) => ({
      name: tc.name,
      clientId: tc.clientId,
      steps: tc.steps.length,
      prerequisites: tc.prerequisites || "",
      prereqSource: tc.prereqSource || "",
    })),
  };
}

function writeLog(logPath, summary) {
  const lines = [];
  lines.push("=".repeat(72));
  lines.push("SimplifyQA Mapping Log");
  lines.push(`Generated at : ${new Date().toISOString()}`);
  lines.push(`Client file  : ${summary.clientFile}`);
  lines.push(`Sheet        : ${summary.sheetName}`);
  lines.push(`Module       : ${summary.module}`);
  lines.push(`Entity       : ${summary.entity}`);
  lines.push(`Versions     : ${summary.versions}`);
  lines.push(`TestcaseType : ${summary.testcaseType}`);
  lines.push(`Output Excel : ${summary.outName}`);
  lines.push(`Mapper prereq: ${summary.kenyaMatched || 0}/${summary.kenyaCount || 0} matched`);
  lines.push("=".repeat(72));
  lines.push("");
  lines.push("--- REVALIDATION ---");
  lines.push(`Named TC rows    : ${summary.namedTcRows}`);
  lines.push(`Inferred TCs     : ${summary.inferredTcCount}`);
  lines.push(`Client testcases : ${summary.clientTcCount}`);
  lines.push(`Mapped testcases : ${summary.mappedTcCount}  [${summary.tcMatch ? "PASS" : "FAIL"}]`);
  lines.push(`Client step rows : ${summary.clientStepLikeCount}`);
  lines.push(`Mapped step rows : ${summary.mappedStepCount}  [${summary.stepMatch ? "PASS" : "FAIL"}]`);
  lines.push(`Sequence 1..N    : [${summary.seqOk ? "PASS" : "FAIL"}]`);
  lines.push(`Mandate last Yes : [${summary.mandateOk ? "PASS" : "FAIL"}]`);
  lines.push("");
  lines.push("--- ISSUE SUMMARY ---");
  const types = Object.keys(summary.issueCounts || {});
  if (!types.length) lines.push("No issues found.");
  else for (const t of types.sort()) lines.push(`  ${t}: ${summary.issueCounts[t]}`);
  lines.push("");
  lines.push("--- ISSUE DETAILS ---");
  if (!summary.issues.length) lines.push("(none)");
  else {
    summary.issues.forEach((iss, idx) => {
      lines.push(
        `${idx + 1}. [${iss.severity}] ${iss.type}` +
          (iss.testcase ? ` | TC: ${iss.testcase}` : "") +
          ` | ${iss.message}`
      );
    });
  }
  lines.push("");
  lines.push("--- PER TESTCASE ---");
  summary.perTc.forEach((tc, i) => {
    lines.push(
      `${i + 1}. ${tc.clientId || "(no-id)"} | ${tc.name} | steps=${tc.steps}` +
        (tc.prerequisites ? " | prereq=YES" : " | prereq=NO")
    );
  });
  lines.push("");
  lines.push("--- END ---");
  fs.writeFileSync(logPath, lines.join("\n"), "utf8");
}

function previewFromRows(outRows, maxRows = 80) {
  const rows = outRows.slice(0, maxRows).map((cells, i) => ({
    row: i + 1,
    cells: cells.map((text) => ({ text: String(text == null ? "" : text) })),
  }));
  return {
    sheet: "Sheet1",
    sheets: ["Sheet1"],
    maxCol: (outRows[0] && outRows[0].length) || SIMPLIFYQA_HEADER.length,
    rows,
    truncated: outRows.length > maxRows,
    totalRows: outRows.length,
  };
}

async function mapFromBuffers({
  clientFileName,
  clientBuffer,
  kenyaFileName,
  kenyaBuffer,
  config,
  mapperHeader,
}) {
  const parsed = parseClientWorkbook(clientBuffer, clientFileName, config);
  let kenyaStats = null;
  let issues = [...parsed.issues];

  if (kenyaBuffer && kenyaFileName) {
    const kenya = await parseKenyaSource(kenyaFileName, kenyaBuffer, {
      prereqHeader: mapperHeader,
    });
    const applied = applyKenyaPrereqs(parsed.testcases, kenya);
    kenyaStats = { matched: applied.matched, kenyaCount: applied.kenyaCount };
    issues = [...issues, ...applied.issues];
  }

  const outRows = buildOutputRows(parsed.testcases, config);
  const summary = buildSummary({
    parsed,
    config,
    kenyaStats,
    issues,
  });

  return {
    parsed,
    issues,
    summary,
    outRows,
    preview: previewFromRows(outRows),
  };
}

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

module.exports = {
  SIMPLIFYQA_HEADER,
  cell,
  normalizeText,
  normalizeName,
  todayStamp,
  nowStamp,
  safeBaseName,
  isLockFile,
  loadProperties,
  stringifyProperties,
  loadJobs,
  resolveJobConfig,
  detectHeaderMap,
  parseClientWorkbook,
  parseKenyaSource,
  applyKenyaPrereqs,
  blankCell,
  aoaToSheetOmitEmpty,
  buildOutputRows,
  writeWorkbook,
  workbookBuffer,
  writeLog,
  buildSummary,
  previewFromRows,
  mapFromBuffers,
  ensureDir,
  isXlsxFileName,
  assertXlsx,
  loadTemplateHeader,
  exactNameKey,
  isPlaceholderId,
  isJunkTestcase,
  listMapperHeaders,
  listWorkbookSheets,
};
