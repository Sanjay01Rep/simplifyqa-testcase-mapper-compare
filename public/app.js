const clientFileEl = document.getElementById("clientFile");
const kenyaFileEl = document.getElementById("kenyaFile");
const existingClientEl = document.getElementById("existingClient");
const clientSheetEl = document.getElementById("clientSheet");
const clientSheetFieldEl = document.getElementById("clientSheetField");
const clientSheetHintEl = document.getElementById("clientSheetHint");
const existingKenyaEl = document.getElementById("existingKenya");
const mapperHeaderEl = document.getElementById("mapperHeader");
const moduleEl = document.getElementById("module");
const moduleCustomEl = document.getElementById("moduleCustom");
const entityCustomEl = document.getElementById("entityCustom");
const entityChecksEl = document.getElementById("entityChecks");
const serverBannerEl = document.getElementById("serverBanner");
const versionsEl = document.getElementById("versions");
const testcaseTypeEl = document.getElementById("testcaseType");
const propertiesTextEl = document.getElementById("propertiesText");
const runStatusEl = document.getElementById("runStatus");
const propsStatusEl = document.getElementById("propsStatus");
const resultBoxEl = document.getElementById("resultBox");
const resultMessageEl = document.getElementById("resultMessage");
const resultLinksEl = document.getElementById("resultLinks");
const logOutputEl = document.getElementById("logOutput");
const runBtn = document.getElementById("runBtn");
const reviewBtn = document.getElementById("reviewBtn");
const previewPanelEl = document.getElementById("previewPanel");
const previewMetaEl = document.getElementById("previewMeta");
const previewTableBody = document.querySelector("#previewTable tbody");
const progressStepsEl = document.getElementById("progressSteps");
const validateBannerEl = document.getElementById("validateBanner");
const historyListEl = document.getElementById("historyList");
const issueListEl = document.getElementById("issueList");
const issueChipsEl = document.getElementById("issueChips");
const issueSearchEl = document.getElementById("issueSearch");
const statsRowEl = document.getElementById("statsRow");
const runForm = document.getElementById("runForm");
const uploadNoticeEl = document.getElementById("uploadNotice");

const XLSX_ONLY_MSG = "Only .xlsx files are supported. Please choose a .xlsx workbook.";
const STEP_ORDER = ["upload", "review", "prereq", "generate", "done"];

function setStatus(el, message, kind) {
  if (!el) return;
  el.textContent = message || "";
  el.classList.remove("ok", "bad");
  if (kind) el.classList.add(kind);
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function fillSelect(select, files, placeholder) {
  if (!select) return;
  const current = select.value;
  select.innerHTML = "";
  const blank = document.createElement("option");
  blank.value = "";
  blank.textContent = placeholder;
  select.appendChild(blank);
  for (const f of files || []) {
    const opt = document.createElement("option");
    opt.value = f.relative || f.name;
    opt.textContent = f.relative && f.relative !== f.name ? f.relative : f.name;
    select.appendChild(opt);
  }
  if ([...select.options].some((o) => o.value === current)) {
    select.value = current;
  }
}

function showUploadNotice(message, kind) {
  if (!uploadNoticeEl) return;
  if (!message) {
    uploadNoticeEl.classList.add("hidden");
    uploadNoticeEl.textContent = "";
    return;
  }
  uploadNoticeEl.classList.remove("hidden", "ok", "warn", "bad");
  uploadNoticeEl.classList.add(kind || "bad");
  uploadNoticeEl.innerHTML = `<strong>${escapeHtml(message)}</strong>`;
}

function isXlsxFile(file) {
  if (!file) return false;
  const name = String(file.name || "").toLowerCase();
  return name.endsWith(".xlsx") && !name.startsWith("~$");
}

async function parseApiJson(res) {
  const text = await res.text();
  const trimmed = (text || "").trim();
  if (!trimmed || trimmed.startsWith("<")) {
    const cannotPost = /Cannot POST\s+(\/\S+)/i.exec(trimmed);
    if (cannotPost) {
      throw new Error(
        `Wrong server answered ${cannotPost[1]} (HTML 404). Close Live Preview and any old Node on port 3100, run npm start in a project terminal, then open http://localhost:3100.`
      );
    }
    throw new Error(
      "Cannot reach the mapping API. In a project terminal run npm start, then open http://localhost:3100 (not Live Preview)."
    );
  }
  try {
    return JSON.parse(trimmed);
  } catch {
    throw new Error("Unexpected server response. Restart npm start and try again.");
  }
}

async function apiFetch(url, options) {
  let res;
  try {
    res = await fetch(url, options);
  } catch (err) {
    throw new Error(
      "Failed to reach the server (Failed to fetch). Run npm start and use http://localhost:3100."
    );
  }
  return parseApiJson(res);
}

function selectedModule() {
  const custom = (moduleCustomEl && moduleCustomEl.value.trim()) || "";
  if (custom) return custom;
  return (moduleEl && moduleEl.value.trim()) || "";
}

function parseEntityList(value) {
  return String(value || "")
    .split(/[,;]/)
    .map((v) => v.trim())
    .filter(Boolean);
}

function selectedEntity() {
  const custom = (entityCustomEl && entityCustomEl.value.trim()) || "";
  if (custom) return parseEntityList(custom).join(", ");
  if (entityChecksEl) return checkedValues(entityChecksEl).join(", ");
  return "";
}

function selectedEntities() {
  const custom = (entityCustomEl && entityCustomEl.value.trim()) || "";
  if (custom) return parseEntityList(custom);
  if (entityChecksEl) return checkedValues(entityChecksEl);
  return [];
}

function cmpSelectedModule() {
  const custom = (cmpModuleCustomEl && cmpModuleCustomEl.value.trim()) || "";
  if (custom) return custom;
  return (cmpModuleEl && cmpModuleEl.value.trim()) || "";
}

function sheetEntities(checkEl, customEl) {
  const custom = (customEl && customEl.value.trim()) || "";
  if (custom) return parseEntityList(custom);
  return checkedValues(checkEl);
}

function fillChoices(select, values, placeholder) {
  if (!select) return;
  const current = select.value;
  select.innerHTML = "";
  const blank = document.createElement("option");
  blank.value = "";
  blank.textContent = placeholder;
  select.appendChild(blank);
  for (const value of values || []) {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value;
    select.appendChild(opt);
  }
  if ([...select.options].some((o) => o.value === current)) {
    select.value = current;
  } else {
    select.value = "";
  }
}

const FALLBACK_MODULES = [
  "General Ledger",
  "Cash & Bank Management",
  "Investment Management",
  "Investment Management>Investment Receipting",
  "Payroll",
  "Payroll>FA Payroll",
  "Salary Review",
  "Salary Review>Bonus",
  "Budgeting",
  "Accounts Receivable",
  "Procurement",
  "Procurement>Inventory",
  "Credit Control",
  "Accounts Payable",
  "Inventory Management",
  "Fixed Assets Management",
  "Financial Reporting",
  "Priority Integrations",
  "Lease",
  "Reinsurance",
  "Tax Management",
  "Expense Management",
  "Consolidation",
  "Power BI Reporting",
  "Integrations",
  "Data Migration",
  "Data Migration>Data Migration -Functional",
  "Data Migration>Data Migration -Technical",
  "E2E Testcases",
  "E2E Testcases>E2E Credit Control",
  "E2E Testcases>E2E General Ledger",
  "E2E Testcases>E2E Cash & Bank",
  "E2E Testcases>E2E Account Payables",
  "E2E Testcases>E2E Fixed Assets",
  "E2E Testcases>E2E Procurement",
  "E2E Testcases>E2E Budgeting",
  "E2E Testcases>E2E Expense Management",
  "E2E Testcases>E2E Accounts Receivables",
  "E2E Testcases>E2E Tax Management",
  "E2E Testcases>E2E Investment Management",
  "E2E Testcases>E2E Financial Reporting",
];
const FALLBACK_ENTITIES = ["Life UG", "Gen UG", "Gen TZ"];

function showServerBanner(ok, message) {
  if (!serverBannerEl) return;
  if (!message) {
    serverBannerEl.classList.add("hidden");
    return;
  }
  serverBannerEl.classList.remove("hidden", "ok", "warn", "bad");
  serverBannerEl.classList.add(ok ? "ok" : "bad");
  serverBannerEl.innerHTML = `<strong>${escapeHtml(message)}</strong>`;
}

function setProgress(activeId, generated) {
  progressStepsEl.hidden = false;
  const reached = STEP_ORDER.indexOf(activeId);
  for (const li of progressStepsEl.querySelectorAll("li")) {
    const id = li.getAttribute("data-step");
    const idx = STEP_ORDER.indexOf(id);
    li.classList.remove("active", "done");
    if (id === "prereq" && !generated && activeId === "review") {
      // keep pending
    }
    if (idx < reached) li.classList.add("done");
    else if (idx === reached) li.classList.add("active");
  }
}

function renderStats(summary) {
  statsRowEl.classList.remove("hidden");
  statsRowEl.innerHTML = [
    ["Testcases", summary.mappedTcCount, summary.tcMatch ? "PASS" : "FAIL"],
    ["Steps", summary.mappedStepCount, summary.stepMatch ? "PASS" : "FAIL"],
    ["Sequence", summary.seqOk ? "OK" : "FIX", summary.seqOk ? "PASS" : "FAIL"],
    [
      "Mapper pre-reqs",
      `${summary.kenyaMatched || 0}/${summary.kenyaCount || 0}`,
      summary.kenyaCount ? "mapped" : "none",
    ],
  ]
    .map(
      ([label, value, hint]) =>
        `<div class="stat"><span>${escapeHtml(label)}</span><strong>${escapeHtml(
          String(value)
        )}</strong><span>${escapeHtml(hint)}</span></div>`
    )
    .join("");
}

let lastIssues = [];
let issueFilter = "ALL";
let issueQuery = "";

function renderIssues(issues, keepFilter) {
  if (Array.isArray(issues)) lastIssues = issues.slice();
  if (!keepFilter) {
    issueFilter = "ALL";
    issueQuery = "";
    if (issueSearchEl) issueSearchEl.value = "";
  }
  const list = lastIssues || [];
  const counts = { FIX: 0, WARN: 0, ERROR: 0, INFO: 0 };
  for (const iss of list) counts[iss.severity] = (counts[iss.severity] || 0) + 1;
  const chips = Object.entries(counts).filter(([, n]) => n);
  const chipHtml = [];
  if (list.length) {
    chipHtml.push(
      `<button type="button" class="chip ALL${issueFilter === "ALL" ? " active" : ""}" data-filter="ALL">ALL ${list.length}</button>`
    );
  }
  for (const [k, n] of chips) {
    chipHtml.push(
      `<button type="button" class="chip ${escapeHtml(k)}${issueFilter === k ? " active" : ""}" data-filter="${escapeHtml(
        k
      )}">${escapeHtml(k)} ${n}</button>`
    );
  }
  issueChipsEl.innerHTML = chipHtml.join("");
  issueChipsEl.querySelectorAll("button.chip").forEach((btn) => {
    btn.addEventListener("click", () => {
      const next = btn.getAttribute("data-filter") || "ALL";
      issueFilter = issueFilter === next && next !== "ALL" ? "ALL" : next;
      renderIssues(null, true);
    });
  });

  if (issueSearchEl) issueSearchEl.hidden = !list.length;

  if (!list.length) {
    issueListEl.innerHTML = `<li class="muted">Review a client file to see sequencing, missing steps, expected results, and pre-req matches.</li>`;
    return;
  }

  const q = issueQuery.trim().toLowerCase();
  const filtered = list.filter((iss) => {
    if (issueFilter !== "ALL" && iss.severity !== issueFilter) return false;
    if (!q) return true;
    const hay = `${iss.severity || ""} ${iss.type || ""} ${iss.testcase || ""} ${iss.message || ""}`.toLowerCase();
    return hay.includes(q);
  });
  if (!filtered.length) {
    const why = q
      ? `No matches for “${escapeHtml(issueQuery.trim())}”.`
      : `No ${escapeHtml(issueFilter)} items.`;
    issueListEl.innerHTML = `<li class="muted">${why}</li>`;
    return;
  }

  issueListEl.innerHTML = filtered
    .slice(0, 120)
    .map(
      (iss) => `<li class="${escapeHtml(iss.severity)}">
        <strong>[${escapeHtml(iss.severity)}]</strong> ${escapeHtml(iss.type)}
        <div class="meta">${escapeHtml(iss.testcase || "")} ${escapeHtml(iss.message)}</div>
      </li>`
    )
    .join("");
}

function renderPreview(preview, label) {
  if (!preview || !preview.rows) {
    previewPanelEl.classList.add("hidden");
    return;
  }
  previewPanelEl.classList.remove("hidden");
  previewMetaEl.textContent = `${label || "Mapped preview"} · ${preview.totalRows || preview.rows.length} rows`;
  previewTableBody.innerHTML = "";

  preview.rows.forEach((row, idx) => {
    const tr = document.createElement("tr");
    if (idx === 0) tr.className = "header-row";
    const num = document.createElement("td");
    num.className = "row-num";
    num.textContent = String(row.row);
    tr.appendChild(num);
    for (const cell of row.cells || []) {
      const td = document.createElement("td");
      td.textContent = cell.text || "";
      td.title = cell.text || "";
      tr.appendChild(td);
    }
    previewTableBody.appendChild(tr);
  });
}

function renderBanner(summary) {
  const warns = (summary.issues || []).filter((i) => i.severity === "WARN").length;
  const errors = (summary.issues || []).filter((i) => i.severity === "ERROR").length;
  const fixes = (summary.issues || []).filter((i) => i.severity === "FIX").length;
  validateBannerEl.classList.remove("hidden", "ok", "warn", "bad");
  if (errors) {
    validateBannerEl.classList.add("bad");
    validateBannerEl.innerHTML = `<strong>Review found ${errors} error(s)</strong> · ${warns} warning(s) · ${fixes} auto-fix(es). Check Validation before importing.`;
  } else if (warns) {
    validateBannerEl.classList.add("warn");
    validateBannerEl.innerHTML = `<strong>${warns} warning(s)</strong> · ${fixes} auto-fix(es) applied (sequence, empty steps, mapper pre-reqs).`;
  } else {
    validateBannerEl.classList.add("ok");
    validateBannerEl.innerHTML = `<strong>Review passed.</strong> ${fixes} auto-fix(es) applied. Ready to generate SimplifyQA Excel.`;
  }
}

function formDataFromUi() {
  const fd = new FormData();
  if (clientFileEl.files[0]) fd.append("client", clientFileEl.files[0]);
  if (kenyaFileEl.files[0]) fd.append("kenya", kenyaFileEl.files[0]);
  fd.append("existingClient", existingClientEl.value || "");
  fd.append("clientSheet", (clientSheetEl && clientSheetEl.value) || "");
  fd.append("existingKenya", existingKenyaEl.value || "");
  fd.append("mapperHeader", (mapperHeaderEl && mapperHeaderEl.value) || "");
  fd.append("module", selectedModule());
  fd.append("moduleCustom", (moduleCustomEl && moduleCustomEl.value.trim()) || "");
  fd.append("entityCustom", (entityCustomEl && entityCustomEl.value.trim()) || "");
  selectedEntities().forEach((v) => fd.append("entity", v));
  fd.append("versions", versionsEl.value.trim() || "v1.0");
  fd.append("testcaseType", testcaseTypeEl.value.trim() || "WEB");
  return fd;
}

function showResult(data, generated) {
  resultBoxEl.classList.remove("hidden");
  const s = data.summary || {};
  const usedSheets = Array.isArray(s.sheetNames) ? s.sheetNames : [];
  const sheetNote =
    usedSheets.length === 1
      ? ` Mapped from sheet "${usedSheets[0]}".`
      : usedSheets.length > 1
        ? ` Mapped from sheets ${usedSheets.join(", ")}.`
        : "";
  resultMessageEl.textContent = generated
    ? `Generated ${s.outName} · ${s.mappedTcCount} TCs · ${s.mappedStepCount} steps · ${s.issues.length} log items.${sheetNote}`
    : `Reviewed ${s.mappedTcCount} TCs / ${s.mappedStepCount} steps. Generate Excel when ready.${sheetNote}`;

  resultLinksEl.innerHTML = "";
  if (data.download && data.download.excel) {
    resultLinksEl.innerHTML = `
      <a href="${data.download.excel}">Download Excel</a>
      <a href="${data.download.log}">Download log</a>
      <button type="button" class="linkish" id="openExcelBtn">Open in Excel</button>
    `;
    const openBtn = document.getElementById("openExcelBtn");
    if (openBtn) {
      openBtn.addEventListener("click", async () => {
        await fetch("/api/launch-excel", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ file: s.outName }),
        });
      });
    }
  }

  logOutputEl.textContent = (s.issues || [])
    .map((i, n) => `${n + 1}. [${i.severity}] ${i.type} | ${i.testcase || ""} | ${i.message}`)
    .join("\n");
}

async function run(path, generated) {
  if (!selectedModule()) {
    setStatus(runStatusEl, "Select a Module, or type a custom module.", "bad");
    return;
  }
  if (!selectedEntities().length) {
    setStatus(runStatusEl, "Select an Entity, or type a custom entity.", "bad");
    return;
  }
  if (!clientFileEl.files[0] && !existingClientEl.value) {
    setStatus(runStatusEl, "Choose a client document.", "bad");
    return;
  }
  if (clientSheetFieldEl && !clientSheetFieldEl.classList.contains("hidden") && !(clientSheetEl && clientSheetEl.value)) {
    const msg = "This client file has more than one sheet. Choose which sheet to map.";
    window.alert(msg);
    setStatus(runStatusEl, msg, "bad");
    showUploadNotice(msg, "bad");
    return;
  }

  runBtn.disabled = true;
  reviewBtn.disabled = true;
  setProgress("upload", generated);
  setStatus(runStatusEl, generated ? "Generating…" : "Reviewing…");

  try {
    setProgress("review", generated);
    const data = await apiFetch(path, { method: "POST", body: formDataFromUi() });
    if (!data.ok) throw new Error(data.message || "Request failed");

    if ((data.summary.kenyaCount || 0) > 0 || data.kenyaFile) setProgress("prereq", generated);
    if (generated) setProgress("generate", generated);
    setProgress("done", generated);

    renderStats(data.summary);
    renderIssues(data.summary.issues);
    renderBanner(data.summary);
    renderPreview(data.preview, data.summary.outName || data.summary.clientFile);
    showResult(data, generated);
    setStatus(
      runStatusEl,
      generated ? "Excel generated." : "Review complete.",
      "ok"
    );
    if (generated) {
      loadHistory();
      if (previewPanelEl) {
        previewPanelEl.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    }
  } catch (err) {
    setStatus(runStatusEl, err.message || String(err), "bad");
  } finally {
    runBtn.disabled = false;
    reviewBtn.disabled = false;
  }
}

async function loadConfig() {
  fillChoices(moduleEl, FALLBACK_MODULES, "-- Select module --");
  fillMapEntityChecks(FALLBACK_ENTITIES);
  const data = await apiFetch("/api/config");
  if (!data.ok) return;
  showServerBanner(true, "");
  const versionEl = document.getElementById("appVersion");
  if (versionEl && data.version) {
    versionEl.textContent = `v${String(data.version).replace(/^v/i, "")}`;
  }
  fillSelect(existingClientEl, data.clientFiles, "Or pick an existing Client doc");
  fillSelect(existingKenyaEl, data.kenyaFiles, "Or pick an existing mapper file");
  fillChoices(moduleEl, data.modules || FALLBACK_MODULES, "-- Select module --");
  fillMapEntityChecks(data.entities || FALLBACK_ENTITIES);
  if (cmpExistingAEl) fillSelect(cmpExistingAEl, data.clientFiles, "Or pick an existing Client doc");
  if (cmpExistingBEl) fillSelect(cmpExistingBEl, data.clientFiles, "Or pick an existing Client doc");
  if (cmpExistingKenyaEl) fillSelect(cmpExistingKenyaEl, data.kenyaFiles, "Or pick an existing mapper file");
  if (cmpModuleEl) fillChoices(cmpModuleEl, data.modules || FALLBACK_MODULES, "-- Select module --");
  fillEntityChecks(data.entities || FALLBACK_ENTITIES);
  applyCompareProps(data.props || {});
  clearMapEntitySelection();
  if (epExistingSummaryEl) fillSelect(epExistingSummaryEl, data.epSampleFiles || [], "Or pick existing sample summary file");
  if (epModuleEl) fillChoices(epModuleEl, data.modules || FALLBACK_MODULES, "-- All modules (or pick) --");
  if (epEntityEl) fillChoices(epEntityEl, data.entities || FALLBACK_ENTITIES, "-- All entities (or pick) --");
  updateEpSheetNamePreview();
  const props = data.props || {};
  // Map Entity is never auto-selected — user must pick (or type custom / Load from properties).
  if (props.Versions) versionsEl.value = props.Versions;
  if (props.TestcaseType) testcaseTypeEl.value = props.TestcaseType;
  if (props.Versions && cmpVersionsEl) cmpVersionsEl.value = props.Versions;
  if (props.TestcaseType && cmpTypeEl) cmpTypeEl.value = props.TestcaseType;
  if (data.healthPollMs) startHealthPoll(data.healthPollMs);
}

async function loadProperties(options = {}) {
  setStatus(propsStatusEl, "Loading…");
  const data = await apiFetch("/api/properties");
  if (!data.ok) {
    setStatus(propsStatusEl, data.message || "Failed to load properties.", "bad");
    return;
  }
  propertiesTextEl.value = data.text || "";
  const props = data.props || {};
  versionsEl.value = props.Versions || versionsEl.value || "v1.0";
  testcaseTypeEl.value = props.TestcaseType || testcaseTypeEl.value || "WEB";
  if (cmpVersionsEl) cmpVersionsEl.value = props.Versions || cmpVersionsEl.value || "v1.0";
  if (cmpTypeEl) cmpTypeEl.value = props.TestcaseType || cmpTypeEl.value || "WEB";
  // Map / Compare Entity stay unchecked unless explicitly requested.
  if (options.applyMapEntity) applyMapEntityFromProps(props);
  else clearMapEntitySelection();
  applyCompareProps(props, { applyCompareEntity: Boolean(options.applyCompareEntity) });
  if (options.applyEpProps) {
    if (epModuleEl && props.Module) epModuleEl.value = props.Module;
    if (epEntityEl && props.Entity) epEntityEl.value = props.Entity;
    if (epVersionEl && props.Versions) epVersionEl.value = props.Versions;
    updateEpSheetNamePreview();
  }
  setStatus(propsStatusEl, "Loaded mapping.properties.", "ok");
}

function clearMapEntitySelection() {
  if (entityChecksEl) setCheckedValues(entityChecksEl, []);
  if (entityCustomEl) entityCustomEl.value = "";
}

function resetMapForm() {
  if (clientFileEl) clientFileEl.value = "";
  if (kenyaFileEl) kenyaFileEl.value = "";
  if (existingClientEl) existingClientEl.value = "";
  if (existingKenyaEl) existingKenyaEl.value = "";
  if (mapperHeaderEl) mapperHeaderEl.value = "";
  if (moduleEl) moduleEl.value = "";
  if (moduleCustomEl) moduleCustomEl.value = "";
  clearMapEntitySelection();
  if (versionsEl) versionsEl.value = "v1.0";
  if (testcaseTypeEl) testcaseTypeEl.value = "WEB";
  hideClientSheetPicker();
  showUploadNotice("");
  showServerBanner(true, "");
  clearMapResults();
  setStatus(runStatusEl, "Form reset.", "ok");
}

function clearMapResults() {
  if (resultBoxEl) resultBoxEl.classList.add("hidden");
  if (resultMessageEl) resultMessageEl.textContent = "";
  if (resultLinksEl) resultLinksEl.innerHTML = "";
  if (logOutputEl) logOutputEl.textContent = "";
  if (previewPanelEl) previewPanelEl.classList.add("hidden");
  if (previewMetaEl) previewMetaEl.textContent = "";
  if (previewTableBody) previewTableBody.innerHTML = "";
  if (statsRowEl) {
    statsRowEl.classList.add("hidden");
    statsRowEl.innerHTML = "";
  }
  if (progressStepsEl) progressStepsEl.hidden = true;
  if (validateBannerEl) {
    validateBannerEl.classList.add("hidden");
    validateBannerEl.textContent = "";
  }
  lastIssues = [];
  issueFilter = "ALL";
  issueQuery = "";
  if (issueSearchEl) {
    issueSearchEl.value = "";
    issueSearchEl.hidden = true;
  }
  if (issueListEl) {
    issueListEl.innerHTML =
      '<li class="muted">Review a client file to see sequencing, missing steps, expected results, and pre-req matches.</li>';
  }
  if (issueChipsEl) issueChipsEl.innerHTML = "";
}

function resetCompareForm() {
  if (cmpFileAEl) cmpFileAEl.value = "";
  if (cmpFileBEl) cmpFileBEl.value = "";
  if (cmpKenyaFileEl) cmpKenyaFileEl.value = "";
  if (cmpExistingAEl) cmpExistingAEl.value = "";
  if (cmpExistingBEl) cmpExistingBEl.value = "";
  if (cmpExistingKenyaEl) cmpExistingKenyaEl.value = "";
  if (cmpMapperHeaderEl) cmpMapperHeaderEl.value = "";
  if (cmpModuleEl) cmpModuleEl.value = "";
  if (cmpModuleCustomEl) cmpModuleCustomEl.value = "";
  setCheckedValues(cmpEntityCommonEl, []);
  setCheckedValues(cmpEntityUniqueAEl, []);
  setCheckedValues(cmpEntityUniqueBEl, []);
  if (cmpEntityCustomCommonEl) cmpEntityCustomCommonEl.value = "";
  if (cmpEntityCustomUniqueAEl) cmpEntityCustomUniqueAEl.value = "";
  if (cmpEntityCustomUniqueBEl) cmpEntityCustomUniqueBEl.value = "";
  if (cmpVersionsEl) cmpVersionsEl.value = "v1.0";
  if (cmpTypeEl) cmpTypeEl.value = "WEB";
  hideCmpSheetPicker("A");
  hideCmpSheetPicker("B");
  showCmpNotice("");
  clearCompareResults();
  setCompareStep(1);
  setStatus(cmpStatusEl, "Form reset.", "ok");
}

function clearCompareResults() {
  const cmpResult = document.getElementById("cmpResult");
  const cmpResultMessage = document.getElementById("cmpResultMessage");
  const cmpResultLinks = document.getElementById("cmpResultLinks");
  const cmpStats = document.getElementById("cmpStats");
  const cmpProgress = document.getElementById("cmpProgressSteps");
  const cmpPreview = document.getElementById("cmpPreviewPanel");
  const cmpPreviewMeta = document.getElementById("cmpPreviewMeta");
  const cmpPreviewBody = document.querySelector("#cmpPreviewTable tbody");
  if (cmpResult) cmpResult.classList.add("hidden");
  if (cmpResultMessage) cmpResultMessage.textContent = "";
  if (cmpResultLinks) cmpResultLinks.innerHTML = "";
  if (cmpStats) {
    cmpStats.classList.add("hidden");
    cmpStats.innerHTML = "";
  }
  if (cmpProgress) cmpProgress.hidden = true;
  if (cmpPreview) cmpPreview.classList.add("hidden");
  if (cmpPreviewMeta) cmpPreviewMeta.textContent = "";
  if (cmpPreviewBody) cmpPreviewBody.innerHTML = "";
  if (validateBannerEl) {
    validateBannerEl.classList.add("hidden");
    validateBannerEl.textContent = "";
  }
  lastIssues = [];
  issueFilter = "ALL";
  issueQuery = "";
  if (issueSearchEl) {
    issueSearchEl.value = "";
    issueSearchEl.hidden = true;
  }
  if (issueListEl) {
    issueListEl.innerHTML =
      '<li class="muted">Review a client file to see sequencing, missing steps, expected results, and pre-req matches.</li>';
  }
  if (issueChipsEl) issueChipsEl.innerHTML = "";
}

async function saveProperties() {
  setStatus(propsStatusEl, "Saving…");
  const data = await apiFetch("/api/properties", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text: propertiesTextEl.value }),
  });
    if (!data.ok) {
      setStatus(propsStatusEl, data.message || "Save failed.", "bad");
    return;
  }
  setStatus(propsStatusEl, "Saved mapping.properties.", "ok");
}

function onXlsxChosen(fileInput, label) {
  const file = fileInput.files[0];
  if (!file) return;
  if (!isXlsxFile(file)) {
    showUploadNotice(XLSX_ONLY_MSG, "bad");
    setStatus(runStatusEl, XLSX_ONLY_MSG, "bad");
    fileInput.value = "";
    return;
  }
  showUploadNotice("");
  setStatus(runStatusEl, "");
}

function assignFileToInput(input, file) {
  const dt = new DataTransfer();
  dt.items.add(file);
  input.files = dt.files;
  input.dispatchEvent(new Event("change", { bubbles: true }));
}

function bindDropzone(zone, input, onInvalid) {
  if (!zone || !input) return;
  zone.addEventListener("dragenter", (e) => {
    e.preventDefault();
    zone.classList.add("dragover");
  });
  zone.addEventListener("dragover", (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
    zone.classList.add("dragover");
  });
  zone.addEventListener("dragleave", (e) => {
    if (!zone.contains(e.relatedTarget)) zone.classList.remove("dragover");
  });
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("dragover");
    const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (!file) return;
    if (!isXlsxFile(file)) {
      if (onInvalid) onInvalid();
      else {
        showUploadNotice(XLSX_ONLY_MSG, "bad");
        setStatus(runStatusEl, XLSX_ONLY_MSG, "bad");
      }
      return;
    }
    assignFileToInput(input, file);
  });
}

async function loadHistory() {
  const res = await fetch("/api/health").catch(() => null);
  if (!res) {
    historyListEl.innerHTML = `<li class="muted">Start the UI with npm start, then open http://localhost:3100.</li>`;
    return;
  }
  const healthText = await res.text();
  let health = null;
  try {
    health = JSON.parse(healthText);
  } catch {
    health = null;
  }
  const icea = res.headers.get("X-ICEA-Lion");
  if (!health || health.ok !== true || icea !== "testcase-review") {
    historyListEl.innerHTML = `<li class="muted">Start the UI with npm start, then open http://localhost:3100 (not Live Preview).</li>`;
    showServerBanner(
      false,
      "Wrong server on this URL. Close Live Preview, run npm start in a project terminal, then open http://localhost:3100."
    );
    return;
  }
  if (Array.isArray(health.routes) && !health.routes.some((r) => String(r).includes("/api/compare"))) {
    showServerBanner(false, "This Node process is missing POST /api/compare. Stop it and run npm start again.");
  }
  const hist = await fetch("/api/history");
  const data = await parseApiJson(hist);
  const runs = (data.runs || []).slice(0, 8);
  if (!runs.length) {
    historyListEl.innerHTML = `<li class="muted">No runs yet.</li>`;
    return;
  }
  historyListEl.innerHTML = runs
    .map((run) => {
      const when = run.at ? new Date(run.at).toLocaleString() : "";
      const ok = run.pass ? "ok-dot" : "bad-dot";
      const links = run.outName
        ? `<div class="links">
            <a href="/api/download/excel?file=${encodeURIComponent(run.outName)}">Excel</a>
            ${
              run.logName
                ? `<a href="/api/download/log?file=${encodeURIComponent(run.logName)}">Log</a>`
                : ""
            }
          </div>`
        : "";
      return `<li>
        <div><span class="${ok}">●</span> ${escapeHtml(run.clientFile || run.outName)}</div>
        <div class="meta">${escapeHtml(when)} · ${run.tcs || 0} TCs · ${run.steps || 0} steps · ${run.issues || 0} issues</div>
        ${links}
      </li>`;
    })
    .join("");
}

runForm.addEventListener("submit", (e) => {
  e.preventDefault();
  run("/api/generate", true);
});
runBtn.addEventListener("click", () => run("/api/generate", true));
reviewBtn.addEventListener("click", () => run("/api/review", false));
document.getElementById("reloadFormBtn").addEventListener("click", () =>
  loadProperties({ applyMapEntity: true, applyCompareEntity: true })
);
document.getElementById("reloadPropsBtn").addEventListener("click", () =>
  loadProperties({ applyMapEntity: true, applyCompareEntity: true })
);
const resetMapBtn = document.getElementById("resetMapBtn");
if (resetMapBtn) resetMapBtn.addEventListener("click", resetMapForm);
document.getElementById("savePropsBtn").addEventListener("click", saveProperties);
document.getElementById("reloadHistoryBtn").addEventListener("click", loadHistory);
if (issueSearchEl) {
  issueSearchEl.addEventListener("input", () => {
    issueQuery = issueSearchEl.value || "";
    renderIssues(null, true);
  });
}
clientFileEl.addEventListener("change", () => {
  onXlsxChosen(clientFileEl, "Client document");
  if (clientFileEl.files[0] && existingClientEl) existingClientEl.value = "";
  refreshClientSheets();
});
if (existingClientEl) {
  existingClientEl.addEventListener("change", () => {
    if (existingClientEl.value && clientFileEl) clientFileEl.value = "";
    refreshClientSheets();
  });
}
kenyaFileEl.addEventListener("change", () => {
  onXlsxChosen(kenyaFileEl, "Mapper file");
  refreshMapperHeaders("map");
});
if (existingKenyaEl) {
  existingKenyaEl.addEventListener("change", () => refreshMapperHeaders("map"));
}
bindDropzone(document.getElementById("clientDrop"), clientFileEl);
bindDropzone(document.getElementById("kenyaDrop"), kenyaFileEl);
["dragover", "drop"].forEach((evt) => {
  document.addEventListener(evt, (e) => {
    if (e.dataTransfer && [...(e.dataTransfer.types || [])].includes("Files")) {
      e.preventDefault();
    }
  });
});

fillChoices(moduleEl, FALLBACK_MODULES, "-- Select module --");
fillMapEntityChecks(FALLBACK_ENTITIES);

const viewMapEl = document.getElementById("viewMap");
const viewCompareEl = document.getElementById("viewCompare");
const tabMap = document.getElementById("tabMap");
const tabCompare = document.getElementById("tabCompare");
const cmpFileAEl = document.getElementById("cmpFileA");
const cmpFileBEl = document.getElementById("cmpFileB");
const cmpExistingAEl = document.getElementById("cmpExistingA");
const cmpExistingBEl = document.getElementById("cmpExistingB");
const cmpSheetAEl = document.getElementById("cmpSheetA");
const cmpSheetBEl = document.getElementById("cmpSheetB");
const cmpSheetFieldAEl = document.getElementById("cmpSheetFieldA");
const cmpSheetFieldBEl = document.getElementById("cmpSheetFieldB");
const cmpSheetHintAEl = document.getElementById("cmpSheetHintA");
const cmpSheetHintBEl = document.getElementById("cmpSheetHintB");
const cmpKenyaFileEl = document.getElementById("cmpKenyaFile");
const cmpExistingKenyaEl = document.getElementById("cmpExistingKenya");
const cmpMapperHeaderEl = document.getElementById("cmpMapperHeader");
const cmpModuleEl = document.getElementById("cmpModule");
const cmpModuleCustomEl = document.getElementById("cmpModuleCustom");
const cmpEntityCommonEl = document.getElementById("cmpEntityCommon");
const cmpEntityUniqueAEl = document.getElementById("cmpEntityUniqueA");
const cmpEntityUniqueBEl = document.getElementById("cmpEntityUniqueB");
const cmpEntityCustomCommonEl = document.getElementById("cmpEntityCustomCommon");
const cmpEntityCustomUniqueAEl = document.getElementById("cmpEntityCustomUniqueA");
const cmpEntityCustomUniqueBEl = document.getElementById("cmpEntityCustomUniqueB");
const cmpVersionsEl = document.getElementById("cmpVersions");
const cmpTypeEl = document.getElementById("cmpType");
const cmpBtn = document.getElementById("cmpBtn");
const cmpReviewBtn = document.getElementById("cmpReviewBtn");
const cmpProgressStepsEl = document.getElementById("cmpProgressSteps");
const cmpStatusEl = document.getElementById("cmpStatus");
const cmpStatsEl = document.getElementById("cmpStats");
const cmpResultEl = document.getElementById("cmpResult");
const cmpResultMessageEl = document.getElementById("cmpResultMessage");
const cmpResultLinksEl = document.getElementById("cmpResultLinks");
const cmpLogOutputEl = document.getElementById("cmpLogOutput");
const cmpNoticeEl = document.getElementById("cmpNotice");
const cmpPreviewPanelEl = document.getElementById("cmpPreviewPanel");
const cmpPreviewMetaEl = document.getElementById("cmpPreviewMeta");
const cmpPreviewTableBody = document.querySelector("#cmpPreviewTable tbody");

const authStatusHintEl = document.getElementById("authStatusHint");
const authReadyActionsEl = document.getElementById("authReadyActions");
const authEditBlockEl = document.getElementById("authEditBlock");
const authTokenLabelEl = document.getElementById("authTokenLabel");
const authTokenInputEl = document.getElementById("authTokenInput");
const saveAuthBtn = document.getElementById("saveAuthBtn");
const cancelAuthBtn = document.getElementById("cancelAuthBtn");
const changeAuthBtn = document.getElementById("changeAuthBtn");
const authStatusEl = document.getElementById("authStatus");

const viewEpEl = document.getElementById("viewEp");
const tabEp = document.getElementById("tabEp");
const resetEpBtn = document.getElementById("resetEpBtn");
const epReloadPropsBtn = document.getElementById("epReloadPropsBtn");
const epForm = document.getElementById("epForm");
const epNoticeEl = document.getElementById("epNotice");
const epSourceUploadRadio = document.getElementById("epSourceUploadRadio");
const epSourceLiveRadio = document.getElementById("epSourceLiveRadio");
const epUploadSection = document.getElementById("epUploadSection");
const epLiveSection = document.getElementById("epLiveSection");
const epSummaryFileEl = document.getElementById("epSummaryFile");
const epExistingSummaryEl = document.getElementById("epExistingSummary");
const epSheetFieldEl = document.getElementById("epSheetField");
const epSheetSelectEl = document.getElementById("epSheetSelect");
const epProjectIdEl = document.getElementById("epProjectId");
const epModuleEl = document.getElementById("epModule");
const epModuleCustomEl = document.getElementById("epModuleCustom");
const epEntityEl = document.getElementById("epEntity");
const epEntityCustomEl = document.getElementById("epEntityCustom");
const epOutputSheetPreviewEl = document.getElementById("epOutputSheetPreview");
const epVersionEl = document.getElementById("epVersion");
const epExecutionTypeEl = document.getElementById("epExecutionType");
const epAssignedDateEl = document.getElementById("epAssignedDate");
const epDatePickerEl = document.getElementById("epDatePicker");
const epDateFormatEl = document.getElementById("epDateFormat");
const epAssigneeEmailEl = document.getElementById("epAssigneeEmail");
const epReviewBtn = document.getElementById("epReviewBtn");
const epGenerateBtn = document.getElementById("epGenerateBtn");
const epStatusEl = document.getElementById("epStatus");
const epStatsEl = document.getElementById("epStats");
const epResultEl = document.getElementById("epResult");
const epResultMessageEl = document.getElementById("epResultMessage");
const epResultLinksEl = document.getElementById("epResultLinks");

function setView(which) {
  const isMap = which === "map";
  const isCompare = which === "compare";
  const isEp = which === "ep";

  if (viewMapEl) viewMapEl.classList.toggle("hidden", !isMap);
  if (viewCompareEl) viewCompareEl.classList.toggle("hidden", !isCompare);
  if (viewEpEl) viewEpEl.classList.toggle("hidden", !isEp);

  if (previewPanelEl) {
    const hasMapPreview = previewTableBody && previewTableBody.children.length;
    previewPanelEl.classList.toggle("hidden", isCompare || (!isMap && !isEp && !hasMapPreview));
  }
  if (cmpPreviewPanelEl) {
    const hasCmpPreview = cmpPreviewTableBody && cmpPreviewTableBody.children.length;
    cmpPreviewPanelEl.classList.toggle("hidden", !isCompare || !hasCmpPreview);
  }
  if (tabMap) {
    tabMap.classList.toggle("active", isMap);
    tabMap.setAttribute("aria-selected", isMap ? "true" : "false");
  }
  if (tabCompare) {
    tabCompare.classList.toggle("active", isCompare);
    tabCompare.setAttribute("aria-selected", isCompare ? "true" : "false");
  }
  if (tabEp) {
    tabEp.classList.toggle("active", isEp);
    tabEp.setAttribute("aria-selected", isEp ? "true" : "false");
  }
  if (isCompare) setCompareStep(1);
}

function hasCompareExcelA() {
  return Boolean(
    (cmpFileAEl && cmpFileAEl.files && cmpFileAEl.files[0]) ||
      (cmpExistingAEl && cmpExistingAEl.value)
  );
}

function hasCompareExcelB() {
  return Boolean(
    (cmpFileBEl && cmpFileBEl.files && cmpFileBEl.files[0]) ||
      (cmpExistingBEl && cmpExistingBEl.value)
  );
}

function setCompareStep(step) {
  const n = Number(step) === 2 ? 2 : 1;
  const step1 = document.getElementById("cmpStep1");
  const step2 = document.getElementById("cmpStep2");
  const tab1 = document.getElementById("cmpWizardTab1");
  const tab2 = document.getElementById("cmpWizardTab2");
  if (step1) step1.classList.toggle("hidden", n !== 1);
  if (step2) step2.classList.toggle("hidden", n !== 2);
  if (tab1) {
    tab1.classList.toggle("active", n === 1);
    tab1.classList.toggle("done", n > 1);
  }
  if (tab2) {
    tab2.classList.toggle("active", n === 2);
    tab2.classList.toggle("done", false);
  }
  if (n === 1) {
    const s1 = document.getElementById("cmpStep1Status");
    if (s1) {
      s1.textContent = "";
      s1.className = "status";
    }
  }
}

function goCompareNext() {
  const statusEl = document.getElementById("cmpStep1Status");
  if (!hasCompareExcelA()) {
    showCmpNotice("Choose Excel A before continuing.", "bad");
    if (statusEl) setStatus(statusEl, "Choose Excel A.", "bad");
    return;
  }
  if (!hasCompareExcelB()) {
    showCmpNotice("Choose Excel B before continuing.", "bad");
    if (statusEl) setStatus(statusEl, "Choose Excel B.", "bad");
    return;
  }
  if (cmpSheetFieldAEl && !cmpSheetFieldAEl.classList.contains("hidden") && !(cmpSheetAEl && cmpSheetAEl.value)) {
    showCmpNotice("Choose which sheet to use for Excel A.", "bad");
    if (statusEl) setStatus(statusEl, "Choose Excel A sheet.", "bad");
    return;
  }
  if (cmpSheetFieldBEl && !cmpSheetFieldBEl.classList.contains("hidden") && !(cmpSheetBEl && cmpSheetBEl.value)) {
    showCmpNotice("Choose which sheet to use for Excel B.", "bad");
    if (statusEl) setStatus(statusEl, "Choose Excel B sheet.", "bad");
    return;
  }
  showCmpNotice("");
  if (statusEl) setStatus(statusEl, "");
  setCompareStep(2);
}

function goCompareBack() {
  setCompareStep(1);
}

function fillMapEntityChecks(values) {
  if (!entityChecksEl) return;
  const list = values && values.length ? values : FALLBACK_ENTITIES;
  entityChecksEl.innerHTML = list
    .map(
      (value) =>
        `<label class="check-opt"><input type="checkbox" value="${escapeHtml(
          value
        )}"> ${escapeHtml(value)}</label>`
    )
    .join("");
  // Never auto-check Map Entity — user must choose.
  setCheckedValues(entityChecksEl, []);
}

function applyMapEntityFromProps(props) {
  if (!props) return;
  if (entityChecksEl) {
    const fromProps = parseEntityProp(props.Entity);
    if (fromProps.length) setCheckedValues(entityChecksEl, fromProps);
  }
  if (entityCustomEl && props.EntityCustom != null) {
    entityCustomEl.value = String(props.EntityCustom || "");
  }
}

function fillEntityChecks(values) {
  const list = values && values.length ? values : FALLBACK_ENTITIES;
  // Never auto-check Compare entities — user must choose (or Load from properties).
  [cmpEntityCommonEl, cmpEntityUniqueAEl, cmpEntityUniqueBEl].forEach((box) => {
    if (!box) return;
    const selected = checkedValues(box);
    box.innerHTML = list
      .map(
        (value) =>
          `<label class="check-opt"><input type="checkbox" value="${escapeHtml(
            value
          )}"> ${escapeHtml(value)}</label>`
      )
      .join("");
    setCheckedValues(box, selected);
  });
}

function checkedValues(box) {
  if (!box) return [];
  return [...box.querySelectorAll("input[type=checkbox]:checked")].map((el) => el.value);
}

function setCheckedValues(box, values) {
  if (!box) return;
  const set = new Set((values || []).map((v) => String(v).trim()).filter(Boolean));
  box.querySelectorAll("input[type=checkbox]").forEach((el) => {
    el.checked = set.has(el.value);
  });
}

function parseEntityProp(value) {
  return String(value || "")
    .split(/[,;]/)
    .map((v) => v.trim())
    .filter(Boolean);
}

function applyCompareProps(props, options = {}) {
  if (!props) return;
  // Module is never auto-selected — user must choose (or type a custom module).
  if (props.Versions && cmpVersionsEl) cmpVersionsEl.value = props.Versions;
  if (props.TestcaseType && cmpTypeEl) cmpTypeEl.value = props.TestcaseType;
  // Entity stays unchecked unless explicitly loading from properties.
  if (options.applyCompareEntity) {
    setCheckedValues(cmpEntityCommonEl, parseEntityProp(props.CompareEntityCommon));
    setCheckedValues(cmpEntityUniqueAEl, parseEntityProp(props.CompareEntityUniqueA));
    setCheckedValues(cmpEntityUniqueBEl, parseEntityProp(props.CompareEntityUniqueB));
    if (cmpEntityCustomCommonEl) {
      cmpEntityCustomCommonEl.value = String(props.CompareEntityCustomCommon || "");
    }
    if (cmpEntityCustomUniqueAEl) {
      cmpEntityCustomUniqueAEl.value = String(props.CompareEntityCustomUniqueA || "");
    }
    if (cmpEntityCustomUniqueBEl) {
      cmpEntityCustomUniqueBEl.value = String(props.CompareEntityCustomUniqueB || "");
    }
  }
  if (cmpExistingKenyaEl && props.CompareKenya) {
    const kenyaVal = String(props.CompareKenya).trim();
    if ([...cmpExistingKenyaEl.options].some((o) => o.value === kenyaVal)) {
      cmpExistingKenyaEl.value = kenyaVal;
    }
  }
  if (cmpModuleEl) cmpModuleEl.value = "";
  if (cmpModuleCustomEl) cmpModuleCustomEl.value = "";
}

function upsertPropLine(text, key, value) {
  const lines = String(text || "").split(/\r?\n/);
  let found = false;
  const next = lines.map((line) => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#") || !trimmed.includes("=")) return line;
    const idx = trimmed.indexOf("=");
    if (trimmed.slice(0, idx).trim() !== key) return line;
    found = true;
    return `${key}=${value}`;
  });
  if (!found) {
    if (next.length && next[next.length - 1] !== "") next.push("");
    next.push(`${key}=${value}`);
  }
  return next.join("\n");
}

async function saveCompareMapping() {
  const common = sheetEntities(cmpEntityCommonEl, cmpEntityCustomCommonEl);
  const uniqueA = sheetEntities(cmpEntityUniqueAEl, cmpEntityCustomUniqueAEl);
  const uniqueB = sheetEntities(cmpEntityUniqueBEl, cmpEntityCustomUniqueBEl);
  if (!common.length || !uniqueA.length || !uniqueB.length) {
    setStatus(cmpStatusEl, "Select or type Entity for Common, Unique A, and Unique B before saving.", "bad");
    return;
  }
  const loaded = await apiFetch("/api/properties");
  let text = loaded.text || "";
  text = upsertPropLine(text, "CompareModule", cmpSelectedModule());
  text = upsertPropLine(text, "CompareModuleCustom", (cmpModuleCustomEl && cmpModuleCustomEl.value.trim()) || "");
  text = upsertPropLine(text, "CompareEntityCommon", checkedValues(cmpEntityCommonEl).join(", "));
  text = upsertPropLine(text, "CompareEntityUniqueA", checkedValues(cmpEntityUniqueAEl).join(", "));
  text = upsertPropLine(text, "CompareEntityUniqueB", checkedValues(cmpEntityUniqueBEl).join(", "));
  text = upsertPropLine(
    text,
    "CompareEntityCustomCommon",
    (cmpEntityCustomCommonEl && cmpEntityCustomCommonEl.value.trim()) || ""
  );
  text = upsertPropLine(
    text,
    "CompareEntityCustomUniqueA",
    (cmpEntityCustomUniqueAEl && cmpEntityCustomUniqueAEl.value.trim()) || ""
  );
  text = upsertPropLine(
    text,
    "CompareEntityCustomUniqueB",
    (cmpEntityCustomUniqueBEl && cmpEntityCustomUniqueBEl.value.trim()) || ""
  );
  text = upsertPropLine(
    text,
    "CompareKenya",
    (cmpExistingKenyaEl && cmpExistingKenyaEl.value) || ""
  );
  text = upsertPropLine(text, "Versions", (cmpVersionsEl && cmpVersionsEl.value.trim()) || "v1.0");
  text = upsertPropLine(text, "TestcaseType", (cmpTypeEl && cmpTypeEl.value.trim()) || "WEB");
  const saved = await apiFetch("/api/properties", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text }),
  });
  if (!saved.ok) {
    setStatus(cmpStatusEl, saved.message || "Could not save mapping.", "bad");
    return;
  }
  if (propertiesTextEl) propertiesTextEl.value = text;
  setStatus(cmpStatusEl, "Compare mapping saved to mapping.properties.", "ok");
}

function showCmpNotice(message, kind) {
  if (!cmpNoticeEl) return;
  if (!message) {
    cmpNoticeEl.classList.add("hidden");
    cmpNoticeEl.textContent = "";
    return;
  }
  cmpNoticeEl.classList.remove("hidden", "ok", "warn", "bad");
  cmpNoticeEl.classList.add(kind || "bad");
  cmpNoticeEl.innerHTML = `<strong>${escapeHtml(message)}</strong>`;
}

function onCmpXlsxChosen(fileInput) {
  const file = fileInput.files[0];
  if (!file) return;
  if (!isXlsxFile(file)) {
    showCmpNotice(XLSX_ONLY_MSG, "bad");
    setStatus(cmpStatusEl, XLSX_ONLY_MSG, "bad");
    fileInput.value = "";
    return;
  }
  showCmpNotice("");
  setStatus(cmpStatusEl, "");
}

function setCmpProgress(activeId, generated) {
  if (!cmpProgressStepsEl) return;
  cmpProgressStepsEl.hidden = false;
  const reached = STEP_ORDER.indexOf(activeId);
  for (const li of cmpProgressStepsEl.querySelectorAll("li")) {
    const id = li.getAttribute("data-step");
    const idx = STEP_ORDER.indexOf(id);
    li.classList.remove("active", "done");
    if (idx < reached) li.classList.add("done");
    else if (idx === reached) li.classList.add("active");
  }
}

function renderCmpStats(summary) {
  const c = (summary && summary.counts) || {};
  const aSum = (c.common || 0) + (c.unmatchedA || 0);
  const bSum = (c.common || 0) + (c.unmatchedB || 0);
  const aOk = aSum === (c.a || 0);
  const bOk = bSum === (c.b || 0);
  cmpStatsEl.classList.remove("hidden");
  cmpStatsEl.innerHTML = [
    ["Excel A", c.a || 0, summary.fileA || ""],
    ["Excel B", c.b || 0, summary.fileB || ""],
    ["Common", c.common || 0, "same steps"],
    ["Unmatched A", c.unmatchedA || 0, c.unmatchedA ? summary.sheets && summary.sheets.a : "no unique — sheet omitted"],
    ["Unmatched B", c.unmatchedB || 0, c.unmatchedB ? summary.sheets && summary.sheets.b : "no unique — sheet omitted"],
    ["A total check", `${c.common || 0}+${c.unmatchedA || 0}=${aSum}`, aOk ? "PASS" : "FAIL"],
    ["B total check", `${c.common || 0}+${c.unmatchedB || 0}=${bSum}`, bOk ? "PASS" : "FAIL"],
    ["Steps differ", c.nameMatchStepMismatch || 0, "both unmatched"],
    ["Sequence", summary.seqOk ? "OK" : "FIX", summary.seqOk ? "PASS" : "FAIL"],
    [
      "Mapper pre-reqs",
      `${summary.kenyaMatched || 0}/${summary.kenyaCount || 0}`,
      summary.kenyaCount ? "mapped" : "none",
    ],
  ]
    .map(
      ([label, value, hint]) =>
        `<div class="stat"><span>${escapeHtml(label)}</span><strong>${escapeHtml(
          String(value)
        )}</strong><span>${escapeHtml(hint || "")}</span></div>`
    )
    .join("");
}

function renderCmpPreview(preview, label) {
  if (!cmpPreviewPanelEl || !cmpPreviewTableBody) return;
  if (!preview || !preview.rows) {
    cmpPreviewPanelEl.classList.add("hidden");
    return;
  }
  cmpPreviewPanelEl.classList.remove("hidden");
  cmpPreviewMetaEl.textContent = `${label || "Common"} · ${preview.totalRows || preview.rows.length} rows`;
  cmpPreviewTableBody.innerHTML = "";
  preview.rows.forEach((row, idx) => {
    const tr = document.createElement("tr");
    if (idx === 0) tr.className = "header-row";
    const num = document.createElement("td");
    num.className = "row-num";
    num.textContent = String(row.row);
    tr.appendChild(num);
    for (const cell of row.cells || []) {
      const td = document.createElement("td");
      td.textContent = cell.text || "";
      td.title = cell.text || "";
      tr.appendChild(td);
    }
    cmpPreviewTableBody.appendChild(tr);
  });
}

function compareFormData() {
  const fd = new FormData();
  if (cmpFileAEl.files[0]) fd.append("clientA", cmpFileAEl.files[0]);
  if (cmpFileBEl.files[0]) fd.append("clientB", cmpFileBEl.files[0]);
  if (cmpKenyaFileEl && cmpKenyaFileEl.files[0]) fd.append("kenya", cmpKenyaFileEl.files[0]);
  fd.append("existingClientA", cmpExistingAEl.value || "");
  fd.append("existingClientB", cmpExistingBEl.value || "");
  fd.append("existingKenya", (cmpExistingKenyaEl && cmpExistingKenyaEl.value) || "");
  fd.append("clientSheetA", (cmpSheetAEl && cmpSheetAEl.value) || "");
  fd.append("clientSheetB", (cmpSheetBEl && cmpSheetBEl.value) || "");
  fd.append("mapperHeader", (cmpMapperHeaderEl && cmpMapperHeaderEl.value) || "");
  fd.append("module", cmpSelectedModule());
  fd.append("moduleCustom", (cmpModuleCustomEl && cmpModuleCustomEl.value.trim()) || "");
  sheetEntities(cmpEntityCommonEl, cmpEntityCustomCommonEl).forEach((v) => fd.append("entityCommon", v));
  sheetEntities(cmpEntityUniqueAEl, cmpEntityCustomUniqueAEl).forEach((v) => fd.append("entityUniqueA", v));
  sheetEntities(cmpEntityUniqueBEl, cmpEntityCustomUniqueBEl).forEach((v) => fd.append("entityUniqueB", v));
  fd.append(
    "entityCustomCommon",
    (cmpEntityCustomCommonEl && cmpEntityCustomCommonEl.value.trim()) || ""
  );
  fd.append(
    "entityCustomUniqueA",
    (cmpEntityCustomUniqueAEl && cmpEntityCustomUniqueAEl.value.trim()) || ""
  );
  fd.append(
    "entityCustomUniqueB",
    (cmpEntityCustomUniqueBEl && cmpEntityCustomUniqueBEl.value.trim()) || ""
  );
  fd.append("versions", (cmpVersionsEl && cmpVersionsEl.value.trim()) || "v1.0");
  fd.append("testcaseType", (cmpTypeEl && cmpTypeEl.value.trim()) || "WEB");
  return fd;
}

function showCmpResult(data, generated) {
  cmpResultEl.classList.remove("hidden");
  const s = data.summary || {};
  const c = s.counts || {};
  cmpResultMessageEl.textContent = generated
    ? `Generated ${s.outName} · Common ${c.common} · unmatched A ${c.unmatchedA} · unmatched B ${c.unmatchedB} · Mapper ${s.kenyaMatched || 0}/${s.kenyaCount || 0}.`
    : `Reviewed Common ${c.common} · unmatched A ${c.unmatchedA} · unmatched B ${c.unmatchedB}. Generate Excel when ready.`;
  cmpResultLinksEl.innerHTML = "";
  if (data.download && data.download.excel) {
    cmpResultLinksEl.innerHTML = `
      <a href="${data.download.excel}">Download Excel</a>
      <a href="${data.download.log}">Download log</a>
      <button type="button" class="linkish" id="cmpOpenExcelBtn">Open in Excel</button>
    `;
    const openBtn = document.getElementById("cmpOpenExcelBtn");
    if (openBtn) {
      openBtn.addEventListener("click", async () => {
        await fetch("/api/launch-excel", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ file: s.outName }),
        });
      });
    }
  }
  cmpLogOutputEl.textContent = (s.issues || [])
    .map((i, n) => `${n + 1}. [${i.severity}] ${i.type} | ${i.testcase || ""} | ${i.message}`)
    .join("\n");
}

async function runCompare(path, generated) {
  try {
  if (!cmpSelectedModule()) {
    setStatus(cmpStatusEl, "Select a Module, or type a custom module.", "bad");
    showCmpNotice("Select a Module, or type a custom module.", "bad");
    return;
  }
  if (!sheetEntities(cmpEntityCommonEl, cmpEntityCustomCommonEl).length) {
    setStatus(cmpStatusEl, "Select or type Entity for the Common sheet.", "bad");
    showCmpNotice("Select or type Entity for the Common sheet.", "bad");
    return;
  }
  if (!sheetEntities(cmpEntityUniqueAEl, cmpEntityCustomUniqueAEl).length) {
    setStatus(cmpStatusEl, "Select or type Entity for unique Excel A sheet.", "bad");
    showCmpNotice("Select or type Entity for unique Excel A sheet.", "bad");
    return;
  }
  if (!sheetEntities(cmpEntityUniqueBEl, cmpEntityCustomUniqueBEl).length) {
    setStatus(cmpStatusEl, "Select or type Entity for unique Excel B sheet.", "bad");
    showCmpNotice("Select or type Entity for unique Excel B sheet.", "bad");
    return;
  }
  if (!cmpFileAEl || (!cmpFileAEl.files[0] && !(cmpExistingAEl && cmpExistingAEl.value))) {
    setStatus(cmpStatusEl, "Choose Excel A.", "bad");
    showCmpNotice("Choose Excel A.", "bad");
    return;
  }
  if (!cmpFileBEl || (!cmpFileBEl.files[0] && !(cmpExistingBEl && cmpExistingBEl.value))) {
    setStatus(cmpStatusEl, "Choose Excel B.", "bad");
    showCmpNotice("Choose Excel B.", "bad");
    return;
  }
  if (cmpSheetFieldAEl && !cmpSheetFieldAEl.classList.contains("hidden") && !(cmpSheetAEl && cmpSheetAEl.value)) {
    setStatus(cmpStatusEl, "Choose which sheet to use for Excel A.", "bad");
    showCmpNotice("Choose which sheet to use for Excel A.", "bad");
    return;
  }
  if (cmpSheetFieldBEl && !cmpSheetFieldBEl.classList.contains("hidden") && !(cmpSheetBEl && cmpSheetBEl.value)) {
    setStatus(cmpStatusEl, "Choose which sheet to use for Excel B.", "bad");
    showCmpNotice("Choose which sheet to use for Excel B.", "bad");
    return;
  }

  showCmpNotice("");
  if (cmpBtn) cmpBtn.disabled = true;
  if (cmpReviewBtn) cmpReviewBtn.disabled = true;
  setCmpProgress("upload", generated);
  setStatus(cmpStatusEl, generated ? "Generating…" : "Reviewing…");
    setCmpProgress("review", generated);
    const data = await apiFetch(path, { method: "POST", body: compareFormData() });
    if (!data.ok) throw new Error(data.message || "Compare failed");
    if ((data.summary.kenyaCount || 0) > 0 || data.kenyaFile) setCmpProgress("prereq", generated);
    if (generated) setCmpProgress("generate", generated);
    setCmpProgress("done", generated);

    renderCmpStats(data.summary);
    renderIssues(data.summary.issues);
    renderBanner(data.summary);
    renderCmpPreview(data.preview, (data.summary && data.summary.outName) || "Common");
    showCmpResult(data, generated);
    setStatus(cmpStatusEl, generated ? "Excel generated." : "Review complete.", "ok");
    if (generated) {
      loadHistory();
      if (cmpPreviewPanelEl) {
        cmpPreviewPanelEl.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    }
  } catch (err) {
    setStatus(cmpStatusEl, err.message || String(err), "bad");
    showCmpNotice(err.message || String(err), "bad");
  } finally {
    if (cmpBtn) cmpBtn.disabled = false;
    if (cmpReviewBtn) cmpReviewBtn.disabled = false;
  }
}

if (tabMap) tabMap.addEventListener("click", () => setView("map"));
if (tabCompare) tabCompare.addEventListener("click", () => setView("compare"));
if (tabEp) tabEp.addEventListener("click", () => setView("ep"));

function showEpNotice(message, kind) {
  if (!epNoticeEl) return;
  if (!message) {
    epNoticeEl.classList.add("hidden");
    epNoticeEl.textContent = "";
    return;
  }
  epNoticeEl.classList.remove("hidden", "ok", "warn", "bad");
  epNoticeEl.classList.add(kind || "bad");
  epNoticeEl.innerHTML = `<strong>${escapeHtml(message)}</strong>`;
}

function updateEpSheetNamePreview() {
  if (!epOutputSheetPreviewEl) return;
  const m = (epModuleCustomEl && epModuleCustomEl.value.trim()) || (epModuleEl && epModuleEl.value.trim()) || "";
  const e = (epEntityCustomEl && epEntityCustomEl.value.trim()) || (epEntityEl && epEntityEl.value.trim()) || "";
  let name = "Sheet1";
  if (m && e) name = `${m} ${e}`;
  else if (e) name = e;
  else if (m) name = m;
  epOutputSheetPreviewEl.textContent = name.slice(0, 31);
}

function formatIsoDate(isoStr, fmt) {
  if (!isoStr) return "";
  const parts = isoStr.split("-");
  if (parts.length !== 3) return isoStr;
  const y = parts[0];
  const m = parts[1];
  const d = parts[2];
  switch (fmt) {
    case "dd/mm/yyyy":
      return `${d}/${m}/${y}`;
    case "m/d/yyyy":
      return `${parseInt(m, 10)}/${parseInt(d, 10)}/${y}`;
    case "yyyy/mm/dd":
      return `${y}/${m}/${d}`;
    case "mm/dd/yyyy":
    default:
      return `${m}/${d}/${y}`;
  }
}

function clearEpResults() {
  if (epResultEl) epResultEl.classList.add("hidden");
  if (epResultMessageEl) epResultMessageEl.textContent = "";
  if (epResultLinksEl) epResultLinksEl.innerHTML = "";
  if (epStatsEl) {
    epStatsEl.classList.add("hidden");
    epStatsEl.innerHTML = "";
  }
}

function resetEpForm() {
  if (epSummaryFileEl) epSummaryFileEl.value = "";
  if (epExistingSummaryEl) epExistingSummaryEl.value = "";
  if (epSheetFieldEl) epSheetFieldEl.classList.add("hidden");
  if (epSheetSelectEl) epSheetSelectEl.innerHTML = '<option value="">-- Choose sheet --</option>';
  if (epProjectIdEl) epProjectIdEl.value = "5";
  if (epModuleEl) epModuleEl.value = "";
  if (epModuleCustomEl) epModuleCustomEl.value = "";
  if (epEntityEl) epEntityEl.value = "";
  if (epEntityCustomEl) epEntityCustomEl.value = "";
  if (epVersionEl) epVersionEl.value = "v1.0";
  if (epExecutionTypeEl) epExecutionTypeEl.value = "Manual";
  if (epAssignedDateEl) epAssignedDateEl.value = "";
  if (epDatePickerEl) epDatePickerEl.value = "";
  if (epDateFormatEl) epDateFormatEl.value = "mm/dd/yyyy";
  if (epAssigneeEmailEl) epAssigneeEmailEl.value = "";
  if (epSourceUploadRadio) epSourceUploadRadio.checked = true;
  if (epUploadSection) epUploadSection.classList.remove("hidden");
  if (epLiveSection) epLiveSection.classList.add("hidden");
  showEpNotice("");
  clearEpResults();
  updateEpSheetNamePreview();
  setStatus(epStatusEl, "Form reset.", "ok");
}

function epFormData() {
  const fd = new FormData();
  const mode = epSourceLiveRadio && epSourceLiveRadio.checked ? "live" : "upload";
  fd.append("mode", mode);

  if (mode === "upload") {
    if (epSummaryFileEl && epSummaryFileEl.files && epSummaryFileEl.files[0]) {
      fd.append("summary", epSummaryFileEl.files[0]);
    }
    if (epExistingSummaryEl && epExistingSummaryEl.value) {
      fd.append("existingSummary", epExistingSummaryEl.value);
    }
    if (epSheetSelectEl && epSheetSelectEl.value) {
      fd.append("summarySheet", epSheetSelectEl.value);
    }
  } else {
    fd.append("projectId", (epProjectIdEl && epProjectIdEl.value) || "5");
  }

  if (epModuleEl && epModuleEl.value) fd.append("module", epModuleEl.value);
  if (epModuleCustomEl && epModuleCustomEl.value) fd.append("moduleCustom", epModuleCustomEl.value.trim());
  if (epEntityEl && epEntityEl.value) fd.append("entity", epEntityEl.value);
  if (epEntityCustomEl && epEntityCustomEl.value) fd.append("entityCustom", epEntityCustomEl.value.trim());
  if (epVersionEl && epVersionEl.value) fd.append("version", epVersionEl.value.trim());
  if (epExecutionTypeEl && epExecutionTypeEl.value) fd.append("executionType", epExecutionTypeEl.value);
  if (epAssignedDateEl && epAssignedDateEl.value) fd.append("assignedDate", epAssignedDateEl.value.trim());
  if (epDateFormatEl && epDateFormatEl.value) fd.append("dateFormat", epDateFormatEl.value);
  if (epAssigneeEmailEl && epAssigneeEmailEl.value) fd.append("assigneeEmail", epAssigneeEmailEl.value.trim());

  return fd;
}

async function runEp(url, isGenerate) {
  showEpNotice("");
  setStatus(epStatusEl, isGenerate ? "Generating Execution Plan Excel…" : "Reviewing test cases…");
  clearEpResults();
  if (epGenerateBtn) epGenerateBtn.disabled = true;
  if (epReviewBtn) epReviewBtn.disabled = true;

  try {
    const fd = epFormData();
    const data = await apiFetch(url, { method: "POST", body: fd });
    if (!data.ok) {
      showEpNotice(data.message || "Execution Plan mapping failed.", "bad");
      setStatus(epStatusEl, data.message || "Failed.", "bad");
      return;
    }

    setStatus(epStatusEl, isGenerate ? "Execution Plan Excel generated." : "Review completed.", "ok");

    if (data.summary) {
      if (epStatsEl) {
        epStatsEl.classList.remove("hidden");
        epStatsEl.innerHTML = `
          <div class="stat"><span class="val">${data.summary.testcaseCount || 0}</span><span class="lbl">Test Cases</span></div>
          <div class="stat"><span class="val">${escapeHtml(data.summary.sheetName || "-")}</span><span class="lbl">Sheet Name</span></div>
          <div class="stat"><span class="val">${escapeHtml(data.summary.version || "v1.0")}</span><span class="lbl">Version</span></div>
          <div class="stat"><span class="val">${escapeHtml(data.summary.executionType || "Manual")}</span><span class="lbl">Type</span></div>
          ${data.summary.assignedDate ? `<div class="stat"><span class="val">${escapeHtml(data.summary.assignedDate)}</span><span class="lbl">Date</span></div>` : ""}
        `;
      }
    }

    if (data.preview) {
      renderPreview(data.preview, `Execution Plan: ${data.summary.sheetName}`);
      if (previewPanelEl) {
        previewPanelEl.classList.remove("hidden");
        previewPanelEl.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    }

    if (isGenerate && data.download) {
      if (epResultEl) epResultEl.classList.remove("hidden");
      if (epResultMessageEl) {
        epResultMessageEl.textContent = `Generated ${data.summary.testcaseCount} test case(s) into Execution Plan sheet "${data.summary.sheetName}".`;
      }
      if (epResultLinksEl) {
        let links = "";
        if (data.download.excel) {
          links += `<a href="${data.download.excel}">Download EP Excel</a>`;
        }
        if (data.download.log) {
          links += `<a href="${data.download.log}">Download log</a>`;
        }
        links += `<button type="button" class="linkish" id="epOpenExcelBtn">Open in Excel</button>`;
        epResultLinksEl.innerHTML = links;

        const openBtn = document.getElementById("epOpenExcelBtn");
        if (openBtn) {
          openBtn.addEventListener("click", async () => {
            await fetch("/api/launch-excel", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ file: data.summary.outName }),
            });
          });
        }
      }
      loadHistory();
    }
  } catch (err) {
    showEpNotice(err.message || "Failed to process Execution Plan.", "bad");
    setStatus(epStatusEl, err.message || "Failed.", "bad");
  } finally {
    if (epGenerateBtn) epGenerateBtn.disabled = false;
    if (epReviewBtn) epReviewBtn.disabled = false;
  }
}

// EP Event Listeners
if (epSourceUploadRadio && epSourceLiveRadio) {
  epSourceUploadRadio.addEventListener("change", () => {
    if (epUploadSection) epUploadSection.classList.remove("hidden");
    if (epLiveSection) epLiveSection.classList.add("hidden");
  });
  epSourceLiveRadio.addEventListener("change", () => {
    if (epUploadSection) epUploadSection.classList.add("hidden");
    if (epLiveSection) epLiveSection.classList.remove("hidden");
  });
}

if (epSummaryFileEl) {
  epSummaryFileEl.addEventListener("change", () => {
    onXlsxChosen(epSummaryFileEl, "Summary Excel");
    if (epSummaryFileEl.files[0] && epExistingSummaryEl) epExistingSummaryEl.value = "";
  });
}
if (epExistingSummaryEl) {
  epExistingSummaryEl.addEventListener("change", () => {
    if (epExistingSummaryEl.value && epSummaryFileEl) epSummaryFileEl.value = "";
  });
}

if (epDatePickerEl && epAssignedDateEl && epDateFormatEl) {
  epDatePickerEl.addEventListener("change", () => {
    if (epDatePickerEl.value) {
      epAssignedDateEl.value = formatIsoDate(epDatePickerEl.value, epDateFormatEl.value);
    }
  });
  epDateFormatEl.addEventListener("change", () => {
    if (epDatePickerEl.value) {
      epAssignedDateEl.value = formatIsoDate(epDatePickerEl.value, epDateFormatEl.value);
    }
  });
}

[epModuleEl, epModuleCustomEl, epEntityEl, epEntityCustomEl].forEach((el) => {
  if (el) {
    el.addEventListener("input", updateEpSheetNamePreview);
    el.addEventListener("change", updateEpSheetNamePreview);
  }
});

if (resetEpBtn) resetEpBtn.addEventListener("click", resetEpForm);
if (epReloadPropsBtn) {
  epReloadPropsBtn.addEventListener("click", () => loadProperties({ applyEpProps: true }));
}

if (epForm) {
  epForm.addEventListener("submit", (e) => {
    e.preventDefault();
    runEp("/api/ep/generate", true);
  });
}
if (epGenerateBtn) {
  epGenerateBtn.addEventListener("click", (e) => {
    e.preventDefault();
    runEp("/api/ep/generate", true);
  });
}
if (epReviewBtn) {
  epReviewBtn.addEventListener("click", (e) => {
    e.preventDefault();
    runEp("/api/ep/review", false);
  });
}

bindDropzone(document.getElementById("epSummaryDrop"), epSummaryFileEl);
const cmpNextBtn = document.getElementById("cmpNextBtn");
const cmpBackBtn = document.getElementById("cmpBackBtn");
const cmpWizardTab1 = document.getElementById("cmpWizardTab1");
const cmpWizardTab2 = document.getElementById("cmpWizardTab2");
if (cmpNextBtn) cmpNextBtn.addEventListener("click", (e) => {
  e.preventDefault();
  goCompareNext();
});
if (cmpBackBtn) cmpBackBtn.addEventListener("click", (e) => {
  e.preventDefault();
  goCompareBack();
});
if (cmpWizardTab1) {
  cmpWizardTab1.addEventListener("click", () => setCompareStep(1));
}
if (cmpWizardTab2) {
  cmpWizardTab2.addEventListener("click", () => {
    if (hasCompareExcelA() && hasCompareExcelB()) setCompareStep(2);
    else goCompareNext();
  });
}
if (cmpBtn) {
  cmpBtn.addEventListener("click", (e) => {
    e.preventDefault();
    runCompare("/api/compare", true);
  });
}
if (cmpReviewBtn) {
  cmpReviewBtn.addEventListener("click", (e) => {
    e.preventDefault();
    runCompare("/api/compare-review", false);
  });
}
if (document.getElementById("compareForm")) {
  document.getElementById("compareForm").addEventListener("submit", (e) => {
    e.preventDefault();
    const step2 = document.getElementById("cmpStep2");
    if (step2 && step2.classList.contains("hidden")) {
      goCompareNext();
      return;
    }
    runCompare("/api/compare", true);
  });
}
setCompareStep(1);
if (document.getElementById("cmpReloadFormBtn")) {
  document.getElementById("cmpReloadFormBtn").addEventListener("click", () =>
    loadProperties({ applyMapEntity: true, applyCompareEntity: true })
  );
}
const resetCompareBtn = document.getElementById("resetCompareBtn");
if (resetCompareBtn) resetCompareBtn.addEventListener("click", resetCompareForm);
if (document.getElementById("cmpSaveMapBtn")) {
  document.getElementById("cmpSaveMapBtn").addEventListener("click", () => {
    saveCompareMapping().catch((err) => setStatus(cmpStatusEl, err.message || String(err), "bad"));
  });
}
if (cmpFileAEl) {
  cmpFileAEl.addEventListener("change", () => {
    onCmpXlsxChosen(cmpFileAEl);
    if (cmpFileAEl.files[0] && cmpExistingAEl) cmpExistingAEl.value = "";
    refreshCmpClientSheets("A");
  });
}
if (cmpFileBEl) {
  cmpFileBEl.addEventListener("change", () => {
    onCmpXlsxChosen(cmpFileBEl);
    if (cmpFileBEl.files[0] && cmpExistingBEl) cmpExistingBEl.value = "";
    refreshCmpClientSheets("B");
  });
}
if (cmpExistingAEl) {
  cmpExistingAEl.addEventListener("change", () => {
    if (cmpExistingAEl.value && cmpFileAEl) cmpFileAEl.value = "";
    refreshCmpClientSheets("A");
  });
}
if (cmpExistingBEl) {
  cmpExistingBEl.addEventListener("change", () => {
    if (cmpExistingBEl.value && cmpFileBEl) cmpFileBEl.value = "";
    refreshCmpClientSheets("B");
  });
}
if (cmpKenyaFileEl) {
  cmpKenyaFileEl.addEventListener("change", () => {
    onCmpXlsxChosen(cmpKenyaFileEl);
    refreshMapperHeaders("compare");
  });
}
if (cmpExistingKenyaEl) {
  cmpExistingKenyaEl.addEventListener("change", () => refreshMapperHeaders("compare"));
}
bindDropzone(document.getElementById("cmpDropA"), cmpFileAEl, () => {
  showCmpNotice(XLSX_ONLY_MSG, "bad");
  setStatus(cmpStatusEl, XLSX_ONLY_MSG, "bad");
});
bindDropzone(document.getElementById("cmpDropB"), cmpFileBEl, () => {
  showCmpNotice(XLSX_ONLY_MSG, "bad");
  setStatus(cmpStatusEl, XLSX_ONLY_MSG, "bad");
});
bindDropzone(document.getElementById("cmpKenyaDrop"), cmpKenyaFileEl, () => {
  showCmpNotice(XLSX_ONLY_MSG, "bad");
  setStatus(cmpStatusEl, XLSX_ONLY_MSG, "bad");
});
fillChoices(cmpModuleEl, FALLBACK_MODULES, "-- Select module --");
fillEntityChecks(FALLBACK_ENTITIES);

function resetMapperHeaderSelect(selectEl) {
  if (!selectEl) return;
  selectEl.innerHTML = `<option value="">Pre-Requisite (default)</option>`;
}

function hideClientSheetPicker() {
  if (clientSheetEl) {
    clientSheetEl.innerHTML = `<option value="">-- Choose sheet --</option>`;
    clientSheetEl.value = "";
  }
  if (clientSheetFieldEl) clientSheetFieldEl.classList.add("hidden");
  if (clientSheetHintEl) clientSheetHintEl.classList.add("hidden");
}

function hideCmpSheetPicker(side) {
  const sheetEl = side === "B" ? cmpSheetBEl : cmpSheetAEl;
  const fieldEl = side === "B" ? cmpSheetFieldBEl : cmpSheetFieldAEl;
  const hintEl = side === "B" ? cmpSheetHintBEl : cmpSheetHintAEl;
  if (sheetEl) {
    sheetEl.innerHTML = `<option value="">-- Choose sheet --</option>`;
    sheetEl.value = "";
  }
  if (fieldEl) fieldEl.classList.add("hidden");
  if (hintEl) hintEl.classList.add("hidden");
}

async function refreshCmpClientSheets(side) {
  const label = side === "B" ? "Excel B" : "Excel A";
  const fileEl = side === "B" ? cmpFileBEl : cmpFileAEl;
  const existingEl = side === "B" ? cmpExistingBEl : cmpExistingAEl;
  const sheetEl = side === "B" ? cmpSheetBEl : cmpSheetAEl;
  const fieldEl = side === "B" ? cmpSheetFieldBEl : cmpSheetFieldAEl;
  const hintEl = side === "B" ? cmpSheetHintBEl : cmpSheetHintAEl;
  if (!sheetEl || !fieldEl) return;

  const hasFile = fileEl && fileEl.files && fileEl.files[0];
  const existing = existingEl && existingEl.value;
  if (!hasFile && !existing) {
    hideCmpSheetPicker(side);
    return;
  }

  const fd = new FormData();
  if (hasFile) fd.append("client", fileEl.files[0]);
  else fd.append("existingClient", existing);

  try {
    const data = await apiFetch("/api/client-sheets", { method: "POST", body: fd });
    const sheets = (data && data.sheets) || [];
    sheetEl.innerHTML = `<option value="">-- Choose sheet --</option>`;
    for (const name of sheets) {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      sheetEl.appendChild(opt);
    }
    if (sheets.length >= 2) {
      fieldEl.classList.remove("hidden");
      if (hintEl) hintEl.classList.remove("hidden");
      window.alert(
        `${label} has ${sheets.length} sheets: ${sheets.join(", ")}.\n\nChoose which sheet to compare.`
      );
      showCmpNotice(`${label} has ${sheets.length} sheets. Choose the sheet to compare.`, "warn");
      const statusEl = document.getElementById("cmpStep1Status");
      if (statusEl) setStatus(statusEl, `Choose ${label} sheet.`, "bad");
    } else {
      if (sheets.length === 1) sheetEl.value = sheets[0];
      fieldEl.classList.add("hidden");
      if (hintEl) hintEl.classList.add("hidden");
    }
  } catch (err) {
    hideCmpSheetPicker(side);
    showCmpNotice(err.message || `Could not read ${label} sheets.`, "bad");
  }
}

async function refreshClientSheets() {
  if (!clientSheetEl || !clientSheetFieldEl) return;
  const hasFile = clientFileEl && clientFileEl.files && clientFileEl.files[0];
  const existing = existingClientEl && existingClientEl.value;
  if (!hasFile && !existing) {
    hideClientSheetPicker();
    return;
  }

  const fd = new FormData();
  if (hasFile) fd.append("client", clientFileEl.files[0]);
  else fd.append("existingClient", existing);

  try {
    const data = await apiFetch("/api/client-sheets", { method: "POST", body: fd });
    const sheets = (data && data.sheets) || [];
    clientSheetEl.innerHTML = `<option value="">-- Choose sheet --</option>`;
    for (const name of sheets) {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      clientSheetEl.appendChild(opt);
    }
    if (sheets.length >= 2) {
      clientSheetFieldEl.classList.remove("hidden");
      if (clientSheetHintEl) clientSheetHintEl.classList.remove("hidden");
      const msg = `This client file has ${sheets.length} sheets: ${sheets.join(
        ", "
      )}.\n\nChoose which sheet to map.`;
      window.alert(msg);
      showUploadNotice(`This file has ${sheets.length} sheets. Choose the sheet to map.`, "warn");
      setStatus(runStatusEl, "Choose a client sheet.", "bad");
    } else {
      if (sheets.length === 1) clientSheetEl.value = sheets[0];
      clientSheetFieldEl.classList.add("hidden");
      if (clientSheetHintEl) clientSheetHintEl.classList.add("hidden");
    }
  } catch (err) {
    hideClientSheetPicker();
    showUploadNotice(err.message || "Could not read client sheets.", "bad");
  }
}

function fillMapperHeaderSelect(selectEl, data) {
  if (!selectEl) return;
  const previous = selectEl.value;
  resetMapperHeaderSelect(selectEl);
  const headers = (data && data.headers) || [];
  for (const name of headers) {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    selectEl.appendChild(opt);
  }
  if (previous && [...selectEl.options].some((o) => o.value === previous)) {
    selectEl.value = previous;
  } else {
    // Keep empty = default Pre-Requisite auto-detect (even if a prereq-like header was found)
    selectEl.value = "";
  }
}

async function refreshMapperHeaders(which) {
  const isCompare = which === "compare";
  const fileInput = isCompare ? cmpKenyaFileEl : kenyaFileEl;
  const existingEl = isCompare ? cmpExistingKenyaEl : existingKenyaEl;
  const headerEl = isCompare ? cmpMapperHeaderEl : mapperHeaderEl;
  if (!headerEl) return;

  const hasFile = fileInput && fileInput.files && fileInput.files[0];
  const existing = existingEl && existingEl.value;
  if (!hasFile && !existing) {
    resetMapperHeaderSelect(headerEl);
    return;
  }

  const fd = new FormData();
  if (hasFile) fd.append("kenya", fileInput.files[0]);
  if (existing) fd.append("existingKenya", existing);
  try {
    const data = await apiFetch("/api/mapper-headers", { method: "POST", body: fd });
    if (!data.ok && !(data.headers && data.headers.length)) {
      resetMapperHeaderSelect(headerEl);
      return;
    }
    fillMapperHeaderSelect(headerEl, data);
  } catch {
    resetMapperHeaderSelect(headerEl);
  }
}

function setServerBotState(state, detail) {
  const bot = document.getElementById("serverBot");
  const text = document.getElementById("serverBotText");
  if (!bot || !text) return;
  bot.classList.remove("up", "down", "unknown");
  bot.classList.add(state);
  if (state === "up") {
    text.textContent = "Server up";
    bot.title = detail || "Server is reachable";
  } else if (state === "down") {
    text.textContent = "Server down";
    bot.title = detail || "Server is not reachable. Run npm start or start-ui.cmd";
  } else {
    text.textContent = "Checking…";
    bot.title = "Checking server status";
  }
}

async function checkServerStatus() {
  if (typeof document !== "undefined" && document.hidden) return false;
  try {
    const res = await fetch("/api/health", { cache: "no-store" });
    if (!res.ok) {
      setServerBotState("down", `HTTP ${res.status}`);
      return false;
    }
    const icea = res.headers.get("X-ICEA-Lion");
    const data = await res.json().catch(() => null);
    if (!data || data.ok !== true || icea !== "testcase-review") {
      setServerBotState("down", "Wrong server on this URL (use npm start, not Live Preview)");
      return false;
    }
    setServerBotState("up", `v${data.version || "?"} · pid ${data.pid || "?"}`);
    return true;
  } catch {
    setServerBotState("down", "Cannot reach server. Run npm start or double-click start-ui.cmd");
    return false;
  }
}

setServerBotState("unknown");
checkServerStatus();

let healthPollTimer = null;
function startHealthPoll(ms) {
  if (healthPollTimer) clearInterval(healthPollTimer);
  const interval = Math.max(5000, Number(ms) || 5 * 60 * 1000);
  healthPollTimer = setInterval(checkServerStatus, interval);
}

// Default 5 minutes; loadConfig may override from server HEALTH_POLL_MS.
startHealthPoll(5 * 60 * 1000);
document.addEventListener("visibilitychange", () => {
  if (!document.hidden) checkServerStatus();
});

function setAuthEditVisible(visible, { allowCancel = true } = {}) {
  if (authEditBlockEl) authEditBlockEl.classList.toggle("hidden", !visible);
  if (authReadyActionsEl) authReadyActionsEl.classList.toggle("hidden", visible);
  if (cancelAuthBtn) cancelAuthBtn.classList.toggle("hidden", !allowCancel);
}

function applyAuthStatus(data, { justSaved = false } = {}) {
  if (!authStatusHintEl) return;
  authStatusHintEl.classList.remove("ok", "bad", "warn");

  if (!data.present) {
    authStatusHintEl.textContent = "No token set — paste your SimplifyQA bearer token below.";
    authStatusHintEl.classList.add("bad");
    if (authTokenLabelEl) authTokenLabelEl.textContent = "SimplifyQA bearer token";
    if (saveAuthBtn) saveAuthBtn.textContent = "Save token";
    setAuthEditVisible(true, { allowCancel: false });
    setStatus(authStatusEl, "");
    return;
  }

  if (data.expired) {
    authStatusHintEl.textContent = "Token expired — paste a fresh bearer token below.";
    authStatusHintEl.classList.add("bad");
    if (authTokenLabelEl) authTokenLabelEl.textContent = "New SimplifyQA bearer token";
    if (saveAuthBtn) saveAuthBtn.textContent = "Update token";
    setAuthEditVisible(true, { allowCancel: false });
    setStatus(authStatusEl, "");
    return;
  }

  authStatusHintEl.textContent = justSaved
    ? "Token saved and ready."
    : data.expiresInMs != null
      ? data.message
      : "Token ready.";

  if (data.expiresInMs != null && data.expiresInMs < 60 * 60 * 1000) {
    authStatusHintEl.classList.add("warn");
  } else {
    authStatusHintEl.classList.add("ok");
  }
  if (authTokenLabelEl) authTokenLabelEl.textContent = "New SimplifyQA bearer token";
  if (saveAuthBtn) saveAuthBtn.textContent = "Update token";
  setAuthEditVisible(false);
  setStatus(
    authStatusEl,
    justSaved ? "Applied immediately — no restart needed." : "",
    justSaved ? "ok" : undefined
  );
}

async function loadAuthStatus() {
  if (!authStatusHintEl) return;
  authStatusHintEl.classList.remove("ok", "bad", "warn");
  authStatusHintEl.textContent = "Checking token…";
  try {
    const data = await apiFetch("/api/auth/status");
    if (!data.ok) {
      authStatusHintEl.textContent = data.message || "Could not read token status.";
      authStatusHintEl.classList.add("bad");
      setAuthEditVisible(true, { allowCancel: false });
      return;
    }
    applyAuthStatus(data);
  } catch (err) {
    authStatusHintEl.textContent = err.message || "Could not read token status.";
    authStatusHintEl.classList.add("bad");
    setAuthEditVisible(true, { allowCancel: false });
  }
}

if (changeAuthBtn) {
  changeAuthBtn.addEventListener("click", () => {
    if (authTokenInputEl) authTokenInputEl.value = "";
    setStatus(authStatusEl, "");
    setAuthEditVisible(true, { allowCancel: true });
  });
}

if (cancelAuthBtn) {
  cancelAuthBtn.addEventListener("click", () => {
    if (authTokenInputEl) authTokenInputEl.value = "";
    setStatus(authStatusEl, "");
    setAuthEditVisible(false);
  });
}

if (saveAuthBtn) {
  saveAuthBtn.addEventListener("click", async () => {
    const token = (authTokenInputEl && authTokenInputEl.value.trim()) || "";
    if (!token) {
      setStatus(authStatusEl, "Paste a token first.", "bad");
      return;
    }
    setStatus(authStatusEl, "Saving…");
    saveAuthBtn.disabled = true;
    try {
      const res = await fetch("/api/auth/token", {
        method: "PUT",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token }),
      });
      const data = await res.json();
      if (!data.ok) {
        setStatus(authStatusEl, data.message || "Save failed.", "bad");
        return;
      }
      if (authTokenInputEl) authTokenInputEl.value = "";
      applyAuthStatus(data, { justSaved: true });
    } catch (err) {
      setStatus(authStatusEl, err.message || String(err), "bad");
    } finally {
      saveAuthBtn.disabled = false;
    }
  });
}

loadConfig().catch((err) => {
  showServerBanner(false, err.message || String(err));
  setStatus(runStatusEl, err.message || String(err), "bad");
  setServerBotState("down", err.message || String(err));
});
loadProperties().catch((err) => setStatus(propsStatusEl, err.message || String(err), "bad"));
loadHistory().catch(() => {});
loadAuthStatus().catch(() => {});
