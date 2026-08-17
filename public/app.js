const clientFileEl = document.getElementById("clientFile");
const kenyaFileEl = document.getElementById("kenyaFile");
const existingClientEl = document.getElementById("existingClient");
const existingKenyaEl = document.getElementById("existingKenya");
const mapperHeaderEl = document.getElementById("mapperHeader");
const moduleEl = document.getElementById("module");
const moduleCustomEl = document.getElementById("moduleCustom");
const entityEl = document.getElementById("entity");
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

function selectedEntity() {
  if (entityChecksEl) {
    return checkedValues(entityChecksEl).join(", ");
  }
  return (entityEl && entityEl.value.trim()) || "";
}

function selectedEntities() {
  if (entityChecksEl) return checkedValues(entityChecksEl);
  const one = selectedEntity();
  return one ? [one] : [];
}

function cmpSelectedModule() {
  const custom = (cmpModuleCustomEl && cmpModuleCustomEl.value.trim()) || "";
  if (custom) return custom;
  return (cmpModuleEl && cmpModuleEl.value.trim()) || "";
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

function renderIssues(issues) {
  const list = issues || [];
  const counts = { FIX: 0, WARN: 0, ERROR: 0, INFO: 0 };
  for (const iss of list) counts[iss.severity] = (counts[iss.severity] || 0) + 1;
  issueChipsEl.innerHTML = Object.entries(counts)
    .filter(([, n]) => n)
    .map(([k, n]) => `<span class="chip ${k}">${k} ${n}</span>`)
    .join("");

  if (!list.length) {
    issueListEl.innerHTML = `<li class="muted">No issues found.</li>`;
    return;
  }

  issueListEl.innerHTML = list
    .slice(0, 80)
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
  fd.append("existingKenya", existingKenyaEl.value || "");
  fd.append("mapperHeader", (mapperHeaderEl && mapperHeaderEl.value) || "");
  fd.append("module", selectedModule());
  fd.append("moduleCustom", (moduleCustomEl && moduleCustomEl.value.trim()) || "");
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
    usedSheets.length === 2
      ? ` 2-sheet file detected; mapped from both sheets (${usedSheets.join(", ")}).`
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
    setStatus(runStatusEl, "Select at least one Entity (Life UG, Gen UG, or Gen TZ).", "bad");
    return;
  }
  if (!clientFileEl.files[0] && !existingClientEl.value) {
    setStatus(runStatusEl, "Choose a client document.", "bad");
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
  const props = data.props || {};
  applyMapEntityFromProps(props);
  if (props.Versions) versionsEl.value = props.Versions;
  if (props.TestcaseType) testcaseTypeEl.value = props.TestcaseType;
  if (props.Versions && cmpVersionsEl) cmpVersionsEl.value = props.Versions;
  if (props.TestcaseType && cmpTypeEl) cmpTypeEl.value = props.TestcaseType;
  if (data.healthPollMs) startHealthPoll(data.healthPollMs);
}

async function loadProperties() {
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
  applyMapEntityFromProps(props);
  applyCompareProps(props);
  setStatus(propsStatusEl, "Loaded mapping.properties.", "ok");
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
  const runs = data.runs || [];
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
document.getElementById("reloadFormBtn").addEventListener("click", loadProperties);
document.getElementById("reloadPropsBtn").addEventListener("click", loadProperties);
document.getElementById("savePropsBtn").addEventListener("click", saveProperties);
document.getElementById("reloadHistoryBtn").addEventListener("click", loadHistory);
clientFileEl.addEventListener("change", () => onXlsxChosen(clientFileEl, "Client document"));
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
const cmpKenyaFileEl = document.getElementById("cmpKenyaFile");
const cmpExistingKenyaEl = document.getElementById("cmpExistingKenya");
const cmpMapperHeaderEl = document.getElementById("cmpMapperHeader");
const cmpModuleEl = document.getElementById("cmpModule");
const cmpModuleCustomEl = document.getElementById("cmpModuleCustom");
const cmpEntityCommonEl = document.getElementById("cmpEntityCommon");
const cmpEntityUniqueAEl = document.getElementById("cmpEntityUniqueA");
const cmpEntityUniqueBEl = document.getElementById("cmpEntityUniqueB");
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

function setView(which) {
  const compare = which === "compare";
  if (viewMapEl) viewMapEl.classList.toggle("hidden", compare);
  if (viewCompareEl) viewCompareEl.classList.toggle("hidden", !compare);
  if (previewPanelEl) {
    const hasMapPreview = previewTableBody && previewTableBody.children.length;
    previewPanelEl.classList.toggle("hidden", compare || !hasMapPreview);
  }
  if (cmpPreviewPanelEl) {
    const hasCmpPreview = cmpPreviewTableBody && cmpPreviewTableBody.children.length;
    cmpPreviewPanelEl.classList.toggle("hidden", !compare || !hasCmpPreview);
  }
  if (tabMap) {
    tabMap.classList.toggle("active", !compare);
    tabMap.setAttribute("aria-selected", compare ? "false" : "true");
  }
  if (tabCompare) {
    tabCompare.classList.toggle("active", compare);
    tabCompare.setAttribute("aria-selected", compare ? "true" : "false");
  }
  if (compare) setCompareStep(1);
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
  const selected = checkedValues(entityChecksEl);
  entityChecksEl.innerHTML = list
    .map(
      (value) =>
        `<label class="check-opt"><input type="checkbox" value="${escapeHtml(
          value
        )}"> ${escapeHtml(value)}</label>`
    )
    .join("");
  setCheckedValues(entityChecksEl, selected.length ? selected : ["Life UG"]);
}

function applyMapEntityFromProps(props) {
  if (!props || !entityChecksEl) return;
  const fromProps = parseEntityProp(props.Entity);
  if (fromProps.length) setCheckedValues(entityChecksEl, fromProps);
}

function fillEntityChecks(values) {
  const list = values && values.length ? values : FALLBACK_ENTITIES;
  const defaults = {
    common: ["Gen UG", "Life UG"],
    uniqueA: ["Gen UG"],
    uniqueB: ["Life UG"],
  };
  [
    [cmpEntityCommonEl, defaults.common],
    [cmpEntityUniqueAEl, defaults.uniqueA],
    [cmpEntityUniqueBEl, defaults.uniqueB],
  ].forEach(([box, fallback]) => {
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
    setCheckedValues(box, selected.length ? selected : fallback);
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

function applyCompareProps(props) {
  if (!props) return;
  // Module is never auto-selected — user must choose (or type a custom module).
  if (props.Versions && cmpVersionsEl) cmpVersionsEl.value = props.Versions;
  if (props.TestcaseType && cmpTypeEl) cmpTypeEl.value = props.TestcaseType;
  setCheckedValues(cmpEntityCommonEl, parseEntityProp(props.CompareEntityCommon));
  setCheckedValues(cmpEntityUniqueAEl, parseEntityProp(props.CompareEntityUniqueA));
  setCheckedValues(cmpEntityUniqueBEl, parseEntityProp(props.CompareEntityUniqueB));
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
  const common = checkedValues(cmpEntityCommonEl);
  const uniqueA = checkedValues(cmpEntityUniqueAEl);
  const uniqueB = checkedValues(cmpEntityUniqueBEl);
  if (!common.length || !uniqueA.length || !uniqueB.length) {
    setStatus(cmpStatusEl, "Select Entity for Common, Unique A, and Unique B before saving.", "bad");
    return;
  }
  const loaded = await apiFetch("/api/properties");
  let text = loaded.text || "";
  text = upsertPropLine(text, "CompareModule", cmpSelectedModule());
  text = upsertPropLine(text, "CompareModuleCustom", (cmpModuleCustomEl && cmpModuleCustomEl.value.trim()) || "");
  text = upsertPropLine(text, "CompareEntityCommon", common.join(", "));
  text = upsertPropLine(text, "CompareEntityUniqueA", uniqueA.join(", "));
  text = upsertPropLine(text, "CompareEntityUniqueB", uniqueB.join(", "));
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
  fd.append("mapperHeader", (cmpMapperHeaderEl && cmpMapperHeaderEl.value) || "");
  fd.append("module", cmpSelectedModule());
  fd.append("moduleCustom", (cmpModuleCustomEl && cmpModuleCustomEl.value.trim()) || "");
  checkedValues(cmpEntityCommonEl).forEach((v) => fd.append("entityCommon", v));
  checkedValues(cmpEntityUniqueAEl).forEach((v) => fd.append("entityUniqueA", v));
  checkedValues(cmpEntityUniqueBEl).forEach((v) => fd.append("entityUniqueB", v));
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
  if (!checkedValues(cmpEntityCommonEl).length) {
    setStatus(cmpStatusEl, "Select at least one Entity for the Common sheet.", "bad");
    showCmpNotice("Select at least one Entity for the Common sheet.", "bad");
    return;
  }
  if (!checkedValues(cmpEntityUniqueAEl).length) {
    setStatus(cmpStatusEl, "Select Entity for unique Excel A sheet.", "bad");
    showCmpNotice("Select Entity for unique Excel A sheet.", "bad");
    return;
  }
  if (!checkedValues(cmpEntityUniqueBEl).length) {
    setStatus(cmpStatusEl, "Select Entity for unique Excel B sheet.", "bad");
    showCmpNotice("Select Entity for unique Excel B sheet.", "bad");
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
  document.getElementById("cmpReloadFormBtn").addEventListener("click", loadProperties);
}
if (document.getElementById("cmpSaveMapBtn")) {
  document.getElementById("cmpSaveMapBtn").addEventListener("click", () => {
    saveCompareMapping().catch((err) => setStatus(cmpStatusEl, err.message || String(err), "bad"));
  });
}
if (cmpFileAEl) cmpFileAEl.addEventListener("change", () => onCmpXlsxChosen(cmpFileAEl));
if (cmpFileBEl) cmpFileBEl.addEventListener("change", () => onCmpXlsxChosen(cmpFileBEl));
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

loadConfig().catch((err) => {
  showServerBanner(false, err.message || String(err));
  setStatus(runStatusEl, err.message || String(err), "bad");
  setServerBotState("down", err.message || String(err));
});
loadProperties().catch((err) => setStatus(propsStatusEl, err.message || String(err), "bad"));
loadHistory().catch(() => {});
