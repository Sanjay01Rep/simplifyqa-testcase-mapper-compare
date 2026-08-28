const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const EP_HEADER = [
  "Testcase ID*",
  "Testcase Name",
  "Testcase Version*",
  "Execution Type*",
  "Assigned Date",
  "Assignee Email ID*",
];

const SUPPORTED_DATE_FORMATS = [
  { id: "mm/dd/yyyy", label: "mm/dd/yyyy (e.g. 08/20/1996, 12/09/2025)" },
  { id: "dd/mm/yyyy", label: "dd/mm/yyyy (e.g. 02/12/2023, 20/08/1996)" },
  { id: "m/d/yyyy", label: "m/d/yyyy (e.g. 4/18/2024, 8/20/1996)" },
  { id: "yyyy/mm/dd", label: "yyyy/mm/dd (e.g. 2023/12/31, 2025/09/12)" },
];

/**
 * Validate and format an assigned date string against the 4 supported formats.
 * Accepts user-typed strings or ISO YYYY-MM-DD from HTML5 date picker.
 */
function validateAndFormatDate(rawDate, preferredFormat = "mm/dd/yyyy") {
  const text = String(rawDate || "").trim();
  if (!text) return { valid: true, value: "", format: preferredFormat };

  // 1. ISO format: YYYY-MM-DD (from input type="date")
  const isoMatch = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
  if (isoMatch) {
    const y = parseInt(isoMatch[1], 10);
    const m = parseInt(isoMatch[2], 10);
    const d = parseInt(isoMatch[3], 10);
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return {
        valid: true,
        value: formatDateParts(y, m, d, preferredFormat),
        format: preferredFormat,
      };
    }
  }

  // 2. Format 4: YYYY/MM/DD or YYYY/M/D
  const ymdMatch = /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/.exec(text);
  if (ymdMatch) {
    const y = parseInt(ymdMatch[1], 10);
    const m = parseInt(ymdMatch[2], 10);
    const d = parseInt(ymdMatch[3], 10);
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return {
        valid: true,
        value: formatDateParts(y, m, d, preferredFormat || "yyyy/mm/dd"),
        format: "yyyy/mm/dd",
      };
    }
  }

  // 3. Formats with slashes: e.g. MM/DD/YYYY, DD/MM/YYYY, M/D/YYYY
  const slashMatch = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
  if (slashMatch) {
    const p1 = parseInt(slashMatch[1], 10);
    const p2 = parseInt(slashMatch[2], 10);
    const y = parseInt(slashMatch[3], 10);

    let m, d, detectedFormat;
    if (preferredFormat === "dd/mm/yyyy") {
      d = p1;
      m = p2;
      detectedFormat = "dd/mm/yyyy";
    } else {
      // Default / mm/dd/yyyy or m/d/yyyy
      m = p1;
      d = p2;
      detectedFormat = slashMatch[1].length === 1 || slashMatch[2].length === 1 ? "m/d/yyyy" : "mm/dd/yyyy";
    }

    // Heuristic: if first part > 12, it must be day
    if (p1 > 12 && p2 <= 12) {
      d = p1;
      m = p2;
      detectedFormat = "dd/mm/yyyy";
    }

    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return {
        valid: true,
        value: formatDateParts(y, m, d, preferredFormat || detectedFormat),
        format: detectedFormat,
      };
    }
  }

  return {
    valid: false,
    value: text,
    error: `Invalid date format "${text}". Supported formats: 1. mm/dd/yyyy (08/20/1996), 2. dd/mm/yyyy (02/12/2023), 3. m/d/yyyy (4/18/2024), 4. yyyy/mm/dd (2023/12/31).`,
  };
}

function padZero(n) {
  return n < 10 ? `0${n}` : String(n);
}

function formatDateParts(year, month, day, fmt) {
  switch (fmt) {
    case "dd/mm/yyyy":
      return `${padZero(day)}/${padZero(month)}/${year}`;
    case "m/d/yyyy":
      return `${month}/${day}/${year}`;
    case "yyyy/mm/dd":
      return `${year}/${padZero(month)}/${padZero(day)}`;
    case "mm/dd/yyyy":
    default:
      return `${padZero(month)}/${padZero(day)}/${year}`;
  }
}

/**
 * Sanitize Excel Sheet name (max 31 characters, remove : \ / ? * [ ] )
 */
function sanitizeSheetName(name, fallback = "Sheet1") {
  const clean = String(name || "")
    .replace(/[\\/*?[\]:]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!clean) return fallback;
  return clean.slice(0, 31);
}

function normalizeTextMatch(str) {
  return String(str || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Clean SimplifyQA internal custom field formatting (e.g. CUSTFIELD_96Life UG_7_d225p -> Life UG)
 */
function cleanSimplifyQaEntity(raw) {
  if (!raw) return "";
  let s = String(raw).trim();
  s = s.replace(/^CUSTFIELD_\d+/i, "");
  s = s.replace(/_\d+_[a-zA-Z0-9]+$/i, "");
  return s.trim();
}

/**
 * Fetch list of modules for a given project in SimplifyQA
 */
async function fetchProjectModules(token, projectId = 5, origin = "https://app.simplifyqa.ai") {
  try {
    const url = `${origin}/pm/module/search`;
    const cleanToken = token.startsWith("Bearer ") ? token : `Bearer ${token.trim()}`;
    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json, text/plain, */*",
        Authorization: cleanToken,
      },
      body: JSON.stringify({
        searchFields: [
          { column: "projectId", value: Number(projectId) || 5, regEx: false },
          { column: "deleted", value: false },
        ],
        selectFields: ["id", "code", "name"],
        startIndex: 0,
        limit: 10000,
      }),
    });
    if (!res.ok) return [];
    const data = await res.json();
    return (data && (data.data || data.records || data.result || (Array.isArray(data) ? data : []))) || [];
  } catch {
    return [];
  }
}

/**
 * Extract test cases from an exported SimplifyQA Summary Excel workbook.
 * Supports multiple modules: options.modules can be an array of module names or comma-separated string,
 * or options.module for single module.
 */
function extractTestcasesFromSummary(bufferOrWorkbook, options = {}) {
  let wb;
  if (Buffer.isBuffer(bufferOrWorkbook) || typeof bufferOrWorkbook === "string") {
    wb = XLSX.read(bufferOrWorkbook, { type: Buffer.isBuffer(bufferOrWorkbook) ? "buffer" : "file" });
  } else {
    wb = bufferOrWorkbook;
  }

  const sheetNames = wb.SheetNames || [];
  if (!sheetNames.length) {
    throw new Error("The uploaded summary Excel has no sheets.");
  }

  const chosenSheet = options.sheet || sheetNames[0];
  const ws = wb.Sheets[chosenSheet];
  if (!ws) {
    throw new Error(`Sheet "${chosenSheet}" was not found in workbook.`);
  }

  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!rows.length) {
    throw new Error("Summary Excel is empty.");
  }

  // Detect column indices from header
  const header = (rows[0] || []).map((h) => String(h || "").trim().toLowerCase());
  const idIdx = header.findIndex((h) => h === "id" || h === "testcase id" || h === "testcase id*");
  const nameIdx = header.findIndex((h) => h === "name" || h === "testcase name" || h === "test case name");
  const moduleIdx = header.findIndex((h) => h === "module" || h === "module*");
  const execTypeIdx = header.findIndex((h) => h === "execution type" || h === "execution type*");
  const entityIdx = header.findIndex((h) => h === "entity" || h === "entity*");
  const versionIdx = header.findIndex(
    (h) => h === "selected version(/s)" || h === "version" || h === "versions*" || h === "testcase version*"
  );

  const colId = idIdx >= 0 ? idIdx : 0;
  const colName = nameIdx >= 0 ? nameIdx : 1;
  const colModule = moduleIdx >= 0 ? moduleIdx : 2;
  const colExec = execTypeIdx >= 0 ? execTypeIdx : 3;
  const colEntity = entityIdx >= 0 ? entityIdx : 6;
  const colVersion = versionIdx >= 0 ? versionIdx : 10;

  // Prepare list of filter modules
  let rawModules = options.modules || options.module || [];
  if (typeof rawModules === "string") {
    rawModules = rawModules.split(",").map((m) => m.trim()).filter(Boolean);
  } else if (!Array.isArray(rawModules)) {
    rawModules = [rawModules].filter(Boolean);
  }
  const filterModules = rawModules.map((m) => String(m).trim()).filter(Boolean);

  const filterEntity = String(options.entity || "").trim();

  const allTestcases = [];
  const filteredTestcases = [];
  const seenIds = new Set();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row.length) continue;

    const rawId = String(row[colId] || "").trim();
    const rawName = String(row[colName] || "").trim();
    if (!rawId || !rawName) continue; // Skip step rows

    if (seenIds.has(rawId)) continue; // Avoid duplicates
    seenIds.add(rawId);

    const tcModule = String(row[colModule] || "").trim();
    const tcExec = String(row[colExec] || "").trim();
    const tcEntity = String(row[colEntity] || "").trim();
    const tcVer = String(row[colVersion] || "").trim();

    const tc = {
      id: rawId,
      name: rawName,
      module: tcModule,
      executionType: tcExec,
      entity: tcEntity,
      version: tcVer,
      rowNumber: i + 1,
    };

    allTestcases.push(tc);

    // Apply module filter if supplied (exact normalized match for selected modules)
    if (filterModules.length > 0) {
      if (!tcModule) continue;
      const rowNorm = normalizeTextMatch(tcModule);
      const matchesAnyModule = filterModules.some((fm) => {
        return rowNorm === normalizeTextMatch(fm);
      });
      if (!matchesAnyModule) {
        continue;
      }
    }

    // Apply entity filter if supplied (clean + match on entity parts)
    if (filterEntity) {
      if (!tcEntity) continue;
      const filterNorm = normalizeTextMatch(filterEntity);
      const entityParts = tcEntity.split(/[,;/|]/).map((e) => cleanSimplifyQaEntity(e.trim())).filter(Boolean);
      const matches = entityParts.some((p) => {
        const pNorm = normalizeTextMatch(p);
        return pNorm === filterNorm || pNorm.includes(filterNorm) || normalizeTextMatch(tcEntity).includes(filterNorm);
      });
      if (!matches) {
        continue;
      }
    }

    filteredTestcases.push(tc);
  }

  return {
    testcases: filteredTestcases,
    allTestcases,
    totalInExport: allTestcases.length,
    filteredCount: filteredTestcases.length,
    sheetName: chosenSheet,
    allSheets: sheetNames,
  };
}

/**
 * Fetch test cases from SimplifyQA API using Bearer token and search filter.
 * Supports multiple modules: moduleNames / moduleIds or single moduleName / moduleId.
 */
async function fetchTestcasesFromSimplifyQa({
  token,
  projectId = 5,
  moduleNames,
  moduleName,
  moduleIds,
  moduleId,
  entity,
  origin = "https://app.simplifyqa.ai",
  limit = 10000,
}) {
  if (!token) {
    throw new Error("SimplifyQA Bearer token is required for live API export.");
  }

  const cleanToken = token.startsWith("Bearer ") ? token : `Bearer ${token.trim()}`;

  // Prepare list of filter modules
  let rawModules = moduleNames || moduleName || [];
  if (typeof rawModules === "string") {
    rawModules = rawModules.split(",").map((m) => m.trim()).filter(Boolean);
  } else if (!Array.isArray(rawModules)) {
    rawModules = [rawModules].filter(Boolean);
  }
  const filterModules = rawModules.map((m) => String(m).trim()).filter(Boolean);

  const filterEntity = String(entity || "").trim();

  // 1. Fetch modules for this project to resolve IDs and names
  const projectModules = await fetchProjectModules(cleanToken, projectId, origin);
  const moduleMap = {};
  const targetModuleIds = [];

  let rawIds = moduleIds || moduleId || [];
  if (typeof rawIds === "string") {
    rawIds = rawIds.split(",").map((id) => Number(id.trim())).filter(Boolean);
  } else if (!Array.isArray(rawIds)) {
    rawIds = [Number(rawIds)].filter(Boolean);
  }
  targetModuleIds.push(...rawIds);

  for (const m of projectModules) {
    if (m && m.id) {
      moduleMap[m.id] = m.name || m.code || `Module ${m.id}`;
      if (filterModules.length > 0) {
        const mNorm = normalizeTextMatch(m.name);
        const matchesAny = filterModules.some((fm) => {
          return mNorm === normalizeTextMatch(fm);
        });
        if (matchesAny && !targetModuleIds.includes(m.id)) {
          targetModuleIds.push(m.id);
        }
      }
    }
  }

  const searchUrl = `${origin}/tm/testcase/search`;
  const searchFields = [
    { column: "id", sort: "dsc" },
    { column: "projectId", value: Number(projectId) || 5 },
    { column: "deleted", value: false, regEx: false },
    { column: "obsolete", value: false, regEx: false },
  ];

  if (targetModuleIds.length > 0) {
    searchFields.push({ column: "moduleId", value: targetModuleIds, regEx: false, condition: "and" });
  }

  const payload = {
    searchFields,
    selectFields: ["code", "name", "moduleId", "moduleName", "executionType", "customFields", "versions"],
    startIndex: 0,
    limit: Number(limit) || 10000,
  };

  const res = await fetch(searchUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json, text/plain, */*",
      Authorization: cleanToken,
    },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    const errText = await res.text().catch(() => "");
    if (res.status === 401 || res.status === 403 || /INVALIDTOKEN|Not authorised|unauthorized/i.test(errText)) {
      throw new Error(
        `SimplifyQA API Authentication Error (${res.status}): Invalid or unauthorized bearer token. Please paste your valid SimplifyQA Bearer token in the Auth panel on the right (or in .env).`
      );
    }
    throw new Error(`SimplifyQA API error (${res.status}): ${errText || res.statusText}`);
  }

  const data = await res.json();
  const records = (data && (data.data || data.records || data.result || (Array.isArray(data) ? data : []))) || [];

  const testcases = [];
  for (const r of records) {
    const id = r.code ? String(r.code).trim() : (r.id ? `TC-${r.id}` : "");
    const name = String(r.name || "").trim();
    if (!id || !name) continue;

    // Resolve module name from moduleMap or record properties
    const tcModule =
      (r.moduleId && moduleMap[r.moduleId]) ||
      String(r.moduleName || (r.module && r.module.name) || "").trim();

    // Module strict filter (exact normalized match)
    if (filterModules.length > 0) {
      if (!tcModule) continue;
      const rowNorm = normalizeTextMatch(tcModule);
      const matchesAnyModule = filterModules.some((fm) => {
        return rowNorm === normalizeTextMatch(fm);
      });
      if (!matchesAnyModule) {
        continue;
      }
    }

    // Extract and clean strings from customFields
    const rawCustomFieldStrings = [];
    if (r.customFields && typeof r.customFields === "object") {
      for (const v of Object.values(r.customFields)) {
        const valArr = Array.isArray(v) ? v : [v];
        for (const item of valArr) {
          if (typeof item === "string" || typeof item === "number") {
            rawCustomFieldStrings.push(String(item));
          } else if (item && typeof item === "object") {
            if (item.value) rawCustomFieldStrings.push(String(item.value));
            if (item.key) rawCustomFieldStrings.push(String(item.key));
            if (item.name) rawCustomFieldStrings.push(String(item.name));
          }
        }
      }
    }

    const cleanedEntities = rawCustomFieldStrings
      .map(cleanSimplifyQaEntity)
      .filter(Boolean);
    const tcEntity = Array.from(new Set(cleanedEntities)).join(", ");

    // Entity filter (supports clean entity match or normalized substring match for SimplifyQA API)
    if (filterEntity) {
      const filterNorm = normalizeTextMatch(filterEntity);
      const matches = rawCustomFieldStrings.some((s) => {
        const cleaned = cleanSimplifyQaEntity(s);
        const sParts = String(cleaned || s).split(/[,;/|]/).map((p) => p.trim()).filter(Boolean);
        return sParts.some((p) => {
          const pNorm = normalizeTextMatch(p);
          return pNorm === filterNorm || pNorm.includes(filterNorm) || normalizeTextMatch(s).includes(filterNorm);
        });
      });
      if (!matches) {
        continue;
      }
    }

    testcases.push({
      id,
      name,
      module: tcModule || (filterModules.length === 1 ? filterModules[0] : "") || "",
      executionType: r.executionType || "Manual",
      entity: tcEntity || entity || "",
      version: r.versions && r.versions.length ? r.versions[0] : "v1.0",
    });
  }

  return {
    testcases,
    totalFetched: records.length,
    filteredCount: testcases.length,
  };
}

/**
 * Build the Execution Plan Excel workbook buffer matching Demo EP Plan template.
 * When testcases span multiple modules, creates a separate sheet for each module.
 */
function buildEpWorkbook({
  testcases,
  moduleNames = [],
  moduleName = "",
  entityName = "",
  version = "v1.0",
  executionType = "Manual",
  assignedDate = "",
  assigneeEmail = "",
  dateFormat = "mm/dd/yyyy",
}) {
  if (!Array.isArray(testcases) || !testcases.length) {
    throw new Error("No test cases to export to Execution Plan.");
  }

  // Format date
  const dateResult = validateAndFormatDate(assignedDate, dateFormat);
  if (assignedDate && !dateResult.valid) {
    throw new Error(dateResult.error);
  }
  const finalDate = dateResult.value || "";

  // Normalize execution type: MANUAL -> Manual
  const normExecType = (raw) => {
    if (executionType && executionType.trim()) return executionType.trim();
    const str = String(raw || "Manual").trim();
    if (/^manual$/i.test(str)) return "Manual";
    if (/^automated$/i.test(str)) return "Automated";
    return str || "Manual";
  };

  const normVersion = (raw) => {
    if (version && version.trim()) return version.trim();
    return String(raw || "v1.0").trim() || "v1.0";
  };

  // Group testcases by module
  // If module is present on testcase, group by tc.module; otherwise use moduleName or "Execution Plan"
  let moduleList = Array.isArray(moduleNames) && moduleNames.length ? moduleNames : (moduleName ? [moduleName] : []);
  if (typeof moduleList === "string") {
    moduleList = moduleList.split(",").map((m) => m.trim()).filter(Boolean);
  }

  const moduleGroups = new Map();
  for (const tc of testcases) {
    let mod = tc.module || (moduleList.length === 1 ? moduleList[0] : "General");
    // Normalize module display name against moduleList if possible
    if (moduleList.length) {
      const match = moduleList.find((m) => normalizeTextMatch(m) === normalizeTextMatch(mod));
      if (match) mod = match;
    }
    if (!moduleGroups.has(mod)) {
      moduleGroups.set(mod, []);
    }
    moduleGroups.get(mod).push(tc);
  }

  const wb = XLSX.utils.book_new();
  const createdSheets = [];
  const sheetStats = [];
  let firstSheetPreview = null;

  const usedSheetNames = new Set();
  const getUniqueSheetName = (name) => {
    let base = sanitizeSheetName(name, "Sheet1");
    let candidate = base;
    let count = 1;
    while (usedSheetNames.has(candidate.toLowerCase())) {
      candidate = sanitizeSheetName(`${base.slice(0, 27)}_${count}`, `Sheet${count}`);
      count++;
    }
    usedSheetNames.add(candidate.toLowerCase());
    return candidate;
  };

  for (const [modName, tcs] of moduleGroups.entries()) {
    const sName = getUniqueSheetName(modName);
    createdSheets.push(sName);

    const rows = [EP_HEADER];
    for (const tc of tcs) {
      rows.push([
        tc.id,
        tc.name,
        normVersion(tc.version),
        normExecType(tc.executionType),
        finalDate,
        String(assigneeEmail || "").trim(),
      ]);
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [
      { wch: 18 }, // Testcase ID*
      { wch: 45 }, // Testcase Name
      { wch: 18 }, // Testcase Version*
      { wch: 16 }, // Execution Type*
      { wch: 16 }, // Assigned Date
      { wch: 25 }, // Assignee Email ID*
    ];

    XLSX.utils.book_append_sheet(wb, ws, sName);

    sheetStats.push({
      sheetName: sName,
      moduleName: modName,
      testcaseCount: tcs.length,
    });

    if (!firstSheetPreview) {
      const previewRows = rows.slice(0, 100).map((r, idx) => ({
        row: idx + 1,
        cells: r.map((c) => ({ text: String(c == null ? "" : c) })),
      }));
      firstSheetPreview = {
        sheet: sName,
        sheets: [], // will fill below
        rows: previewRows,
        totalRows: rows.length,
        maxCol: EP_HEADER.length,
      };
    }
  }

  if (firstSheetPreview) {
    firstSheetPreview.sheets = createdSheets;
  }

  const buffer = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  const primarySheetName = createdSheets[0] || "Sheet1";
  const displayModuleName = moduleList.length > 0 ? moduleList.join(", ") : (moduleGroups.keys().next().value || "");

  return {
    buffer,
    workbook: wb,
    sheetName: primarySheetName,
    sheetNames: createdSheets,
    sheetStats,
    rowCount: testcases.length,
    preview: firstSheetPreview,
    summary: {
      sheetName: primarySheetName,
      sheetNames: createdSheets,
      sheetStats,
      testcaseCount: testcases.length,
      moduleName: displayModuleName,
      modules: moduleList,
      entityName,
      version: normVersion(""),
      executionType: normExecType(""),
      assignedDate: finalDate,
      assigneeEmail: String(assigneeEmail || "").trim(),
    },
  };
}

module.exports = {
  EP_HEADER,
  SUPPORTED_DATE_FORMATS,
  validateAndFormatDate,
  sanitizeSheetName,
  extractTestcasesFromSummary,
  fetchTestcasesFromSimplifyQa,
  buildEpWorkbook,
};
