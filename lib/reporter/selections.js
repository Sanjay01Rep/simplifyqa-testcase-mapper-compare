/**
 * Per-template Reporter form memory.
 * Saved after a successful generate; used to refill the UI and scheduled runs.
 */

const fs = require("fs");
const path = require("path");

const ROOT = path.join(__dirname, "../..");
const DEFAULT_PATH = path.join(ROOT, "config", "reporter-selections.json");
const PROPERTIES_PATH = path.join(ROOT, "config", "application.properties");

function emptyStore() {
  return { lastTemplate: "", templates: {} };
}

function normalizePlanIds(planIds) {
  if (!Array.isArray(planIds)) return [];
  return planIds.map((id) => String(id || "").trim()).filter(Boolean);
}

function normalizeSelection(form) {
  const src = form && typeof form === "object" ? form : {};
  return {
    projectId: String(src.projectId || "").trim(),
    projectIdB: String(src.projectIdB || "").trim(),
    templateChoice: String(src.templateChoice || "").trim(),
    planIds: normalizePlanIds(src.planIds),
    includeDefects: src.includeDefects !== false,
    includePdf: src.includePdf === true,
    defectSprint: String(src.defectSprint || "").trim(),
    compareEntitySummary: src.compareEntitySummary !== false,
    compareSheet: String(src.compareSheet || "").trim(),
    draftOutlookEmail: src.draftOutlookEmail !== false,
  };
}

function readStore(filePath = DEFAULT_PATH) {
  try {
    if (!fs.existsSync(filePath)) return emptyStore();
    const raw = JSON.parse(fs.readFileSync(filePath, "utf8"));
    const templates = {};
    const src = raw && raw.templates && typeof raw.templates === "object" ? raw.templates : {};
    for (const [key, value] of Object.entries(src)) {
      const choice = String(key || "").trim();
      if (!choice) continue;
      templates[choice] = normalizeSelection({ ...value, templateChoice: choice });
    }
    return {
      lastTemplate: String((raw && raw.lastTemplate) || "").trim(),
      templates,
    };
  } catch {
    return emptyStore();
  }
}

function writeStore(store, filePath = DEFAULT_PATH) {
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  const payload = {
    lastTemplate: String((store && store.lastTemplate) || "").trim(),
    templates: (store && store.templates) || {},
  };
  fs.writeFileSync(filePath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
}

function getSelection(templateChoice, filePath = DEFAULT_PATH) {
  const choice = String(templateChoice || "").trim();
  if (!choice) return null;
  const saved = readStore(filePath).templates[choice];
  return saved || null;
}

function saveSelection(form, filePath = DEFAULT_PATH) {
  const selection = normalizeSelection(form);
  if (!selection.templateChoice) return selection;
  const store = readStore(filePath);
  store.templates[selection.templateChoice] = selection;
  store.lastTemplate = selection.templateChoice;
  writeStore(store, filePath);
  return selection;
}

function upsertPropertyLine(text, key, value) {
  const line = `${key}=${value == null ? "" : String(value)}`;
  const re = new RegExp(`^${key}=.*$`, "m");
  if (re.test(text)) return text.replace(re, line);
  return `${String(text || "").replace(/\s*$/, "")}\n${line}\n`;
}

function syncPropertiesFromSelection(form, propsPath = PROPERTIES_PATH) {
  const selection = normalizeSelection(form);
  if (!selection.templateChoice || !fs.existsSync(propsPath)) return;
  let text = fs.readFileSync(propsPath, "utf8");
  text = upsertPropertyLine(text, "TEMPLATE_CHOICE", selection.templateChoice);
  if (selection.projectId) text = upsertPropertyLine(text, "PROJECT_ID", selection.projectId);
  if (selection.projectIdB) text = upsertPropertyLine(text, "PROJECT_ID_B", selection.projectIdB);

  const existing = [];
  const re = /^EXE_PLAN_ID_(\d+)=/gm;
  let match;
  while ((match = re.exec(text))) existing.push(Number(match[1]));
  const max = Math.max(selection.planIds.length, ...existing, 0);
  for (let i = 1; i <= max; i++) {
    text = upsertPropertyLine(text, `EXE_PLAN_ID_${i}`, selection.planIds[i - 1] || "");
  }
  fs.writeFileSync(propsPath, text.replace(/\r?\n/g, "\n"), "utf8");
}

function rememberSuccessfulRun(form, filePath = DEFAULT_PATH, propsPath = PROPERTIES_PATH) {
  const saved = saveSelection(form, filePath);
  try {
    syncPropertiesFromSelection(saved, propsPath);
  } catch {
    /* properties stay as they were; JSON is still saved */
  }
  return saved;
}

function hasIncomingPlans(form) {
  return Array.isArray(form.planIds) && form.planIds.some((id) => String(id || "").trim());
}

function mergeFormWithSavedSelection(incoming, templateChoiceFromProps, filePath = DEFAULT_PATH) {
  const form = incoming && typeof incoming === "object" ? { ...incoming } : {};
  const choice = String(form.templateChoice || templateChoiceFromProps || "").trim();
  const saved = getSelection(choice, filePath);
  if (!saved) {
    if (choice && !form.templateChoice) form.templateChoice = choice;
    return form;
  }
  const merged = { ...saved, ...form };
  merged.templateChoice = form.templateChoice || saved.templateChoice || choice;
  if (!hasIncomingPlans(form)) merged.planIds = saved.planIds;
  if (!String(form.projectId || "").trim()) merged.projectId = saved.projectId;
  if (!String(form.projectIdB || "").trim() && saved.projectIdB) {
    merged.projectIdB = saved.projectIdB;
  }
  if (form.includeDefects === undefined) merged.includeDefects = saved.includeDefects;
  if (form.includePdf === undefined) merged.includePdf = saved.includePdf;
  if (form.defectSprint === undefined) merged.defectSprint = saved.defectSprint;
  if (form.compareEntitySummary === undefined) {
    merged.compareEntitySummary = saved.compareEntitySummary;
  }
  if (form.compareSheet === undefined) merged.compareSheet = saved.compareSheet;
  if (form.draftOutlookEmail === undefined) merged.draftOutlookEmail = saved.draftOutlookEmail;
  return merged;
}

function applySavedOntoForm(form, saved) {
  if (!saved) return form;
  return {
    ...form,
    ...saved,
    templates: form.templates,
    projects: form.projects,
    plans: form.plans,
    templateChoice: saved.templateChoice || form.templateChoice,
    planIds:
      saved.planIds && saved.planIds.length ? saved.planIds : form.planIds,
  };
}

module.exports = {
  DEFAULT_PATH,
  emptyStore,
  normalizeSelection,
  readStore,
  writeStore,
  getSelection,
  saveSelection,
  syncPropertiesFromSelection,
  rememberSuccessfulRun,
  mergeFormWithSavedSelection,
  applySavedOntoForm,
};
