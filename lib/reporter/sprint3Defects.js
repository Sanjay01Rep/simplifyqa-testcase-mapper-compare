/**
 * Sprint 3 Kenya defect extras: Sprint column preference + sprint filter.
 * Other templates must not call this module.
 */
const CAPTURED_KENYA_DEFECT_FILTER = require("./sprint3ColumnFilter");

function isSprint3TemplatePath(rel) {
  return /sprint\s*3/i.test(String(rel || "").replace(/\\/g, "/"));
}

function normalizeSprintText(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function extractSprintNeedles(input) {
  const raw = String(input || "").trim();
  const needles = [];
  if (!raw) return { raw: "", needles };
  const re = /\bsprint\s*(\d+)\b/gi;
  let m;
  while ((m = re.exec(raw))) {
    needles.push(`sprint ${m[1]}`);
  }
  return { raw, needles };
}

function sprintValueMatches(cellValue, userInput) {
  const cell = normalizeSprintText(cellValue);
  if (!cell) return false;
  const { raw, needles } = extractSprintNeedles(userInput);
  if (!raw) return true;
  if (needles.length) {
    return needles.some((needle) => {
      const n = needle.replace(/\s+/g, "\\s*");
      return new RegExp(`\\b${n}\\b`, "i").test(cell);
    });
  }
  const rawNorm = normalizeSprintText(raw);
  return cell === rawNorm || cell.includes(rawNorm);
}

function isSprintPreferenceColumn(col) {
  if (!col || typeof col !== "object") return false;
  if (String(col.key || "") === "iterationId") return true;
  const label = `${col.label || ""} ${col.originalLabel || ""}`.toLowerCase();
  return /\bsprint\b/.test(label);
}

function findSprintPreferenceIndex(columns) {
  if (!Array.isArray(columns)) return -1;
  return columns.findIndex(isSprintPreferenceColumn);
}

function kenyaSprintColumn(projectId) {
  const source =
    (CAPTURED_KENYA_DEFECT_FILTER || []).find(isSprintPreferenceColumn) || {
      originalLabel: "Sprint",
      label: "Sprint",
      key: "iterationId",
      controlType: "RELEASE-SPRINT",
      hasMap: true,
      displayKey: { field: "name" },
      parentId: 9,
      seq: 10,
      fieldType: "defaultFld",
      mandatory: true,
      deleted: false,
      readOnly: false,
      shownIn: ["TABLE", "EDIT", "CREATE", "IMPORT"],
      filter: true,
      sort: true,
      sortType: "dsc",
      width: 235,
      keyname: "id",
      type: "KEY",
      isDefault: true,
      id: "e73fce4f-005c-480b-9a47-bce672cbb4bf",
      projectId: 2,
      isChecked: true,
      searchRow: "",
    };
  const copy = JSON.parse(JSON.stringify(source));
  copy.isChecked = true;
  copy.deleted = false;
  const pid = Number(projectId);
  if (Number.isFinite(pid) && pid > 0) copy.projectId = pid;
  const shown = Array.isArray(copy.shownIn) ? copy.shownIn.slice() : [];
  if (!shown.includes("TABLE")) shown.push("TABLE");
  copy.shownIn = shown;
  return copy;
}

function capturedKenyaDefectFilter() {
  return JSON.parse(JSON.stringify(CAPTURED_KENYA_DEFECT_FILTER || []));
}

function preferenceHasSprintColumn(filter, projectId) {
  const idx = findSprintPreferenceIndex(filter);
  if (idx < 0) return false;
  const col = filter[idx] || {};
  const shown = Array.isArray(col.shownIn) ? col.shownIn : [];
  return (
    col.isChecked === true &&
    col.deleted !== true &&
    shown.includes("TABLE") &&
    (!Number(col.projectId) ||
      !Number(projectId) ||
      Number(col.projectId) === Number(projectId))
  );
}

function mergeSprintIntoDefectFilter(defectFilter, projectId) {
  const filter = Array.isArray(defectFilter) ? defectFilter.slice() : [];
  const wanted = kenyaSprintColumn(projectId);
  const idx = findSprintPreferenceIndex(filter);
  let changed = false;
  if (idx < 0) {
    wanted.seq = filter.length + 1;
    filter.push(wanted);
    changed = true;
  } else {
    const existing = filter[idx] || {};
    if (
      existing.isChecked !== true ||
      existing.deleted === true ||
      !(Array.isArray(existing.shownIn) && existing.shownIn.includes("TABLE"))
    ) {
      filter[idx] = Object.assign({}, existing, {
        isChecked: true,
        deleted: false,
        shownIn: wanted.shownIn,
      });
      changed = true;
    }
  }
  return { defectFilter: filter, changed };
}

module.exports = {
  isSprint3TemplatePath,
  extractSprintNeedles,
  sprintValueMatches,
  isSprintPreferenceColumn,
  findSprintPreferenceIndex,
  kenyaSprintColumn,
  capturedKenyaDefectFilter,
  preferenceHasSprintColumn,
  mergeSprintIntoDefectFilter,
};
