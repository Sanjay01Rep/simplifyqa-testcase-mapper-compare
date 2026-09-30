/**
 * Pass Rate / Execution Rate denominators.
 * Checked: Total − Blocked. Unchecked (default): Total.
 * Execution Rate numerator is always Passed + Failed.
 */

function excludeBlockedEnabled(value) {
  return value === true || value === "true";
}

function rate(numerator, denominator) {
  if (!denominator) return 1;
  return numerator / denominator;
}

function rateDenom(total, blocked, excludeBlocked) {
  const t = Number(total) || 0;
  const b = Number(blocked) || 0;
  return excludeBlockedEnabled(excludeBlocked) ? t - b : t;
}

function passRate(passed, total, blocked, excludeBlocked) {
  return rate(Number(passed) || 0, rateDenom(total, blocked, excludeBlocked));
}

function executionRate(passed, failed, total, blocked, excludeBlocked) {
  return rate(
    (Number(passed) || 0) + (Number(failed) || 0),
    rateDenom(total, blocked, excludeBlocked)
  );
}

function excelPassRateFormula(cPass, cTot, cBlock, row, excludeBlocked) {
  if (excludeBlockedEnabled(excludeBlocked) && cBlock) {
    return `${cPass}${row}/(${cTot}${row}-${cBlock}${row})`;
  }
  return `${cPass}${row}/${cTot}${row}`;
}

function excelExecutionRateFormula(cPass, cFail, cTot, cBlock, row, excludeBlocked) {
  const num = `(${cPass}${row}+${cFail}${row})`;
  if (excludeBlockedEnabled(excludeBlocked) && cBlock) {
    return `${num}/(${cTot}${row}-${cBlock}${row})`;
  }
  return `${num}/${cTot}${row}`;
}

module.exports = {
  excludeBlockedEnabled,
  rate,
  rateDenom,
  passRate,
  executionRate,
  excelPassRateFormula,
  excelExecutionRateFormula,
};
