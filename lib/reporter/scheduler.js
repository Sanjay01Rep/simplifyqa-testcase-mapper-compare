/**
 * Simple daily multi-time scheduler (local machine clock).
 * Fires at most once per scheduled HH:MM per calendar day.
 */

function parseHm(hm) {
  const [h, m] = String(hm).split(":").map(Number);
  return { h, m };
}

function nextRunFrom(times, fromDate = new Date()) {
  if (!times || !times.length) return null;
  const base = new Date(fromDate);
  for (let dayOffset = 0; dayOffset <= 1; dayOffset++) {
    const day = new Date(base);
    day.setDate(base.getDate() + dayOffset);
    for (const t of times) {
      const { h, m } = parseHm(t);
      const candidate = new Date(
        day.getFullYear(),
        day.getMonth(),
        day.getDate(),
        h,
        m,
        0,
        0
      );
      if (candidate > fromDate) return candidate;
    }
  }
  return null;
}

function createScheduler({ getConfig, onFire, log }) {
  let timer = null;
  let lastFiredKey = "";
  let lastRunAt = null;
  let lastRunOk = null;
  let lastRunMessage = "";

  function status() {
    const cfg = getConfig();
    const next = cfg.enabled ? nextRunFrom(cfg.times) : null;
    return {
      enabled: Boolean(cfg.enabled),
      times: cfg.times || [],
      nextRunAt: next ? next.toISOString() : null,
      nextRunLocal: next ? next.toLocaleString() : null,
      lastRunAt,
      lastRunOk,
      lastRunMessage,
    };
  }

  async function tick() {
    const cfg = getConfig();
    if (!cfg.enabled || !cfg.times.length) return;
    const now = new Date();
    const key = `${now.getFullYear()}-${now.getMonth()}-${now.getDate()} ${String(
      now.getHours()
    ).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;
    if (!cfg.times.includes(key.slice(-5))) return;
    if (lastFiredKey === key) return;
    lastFiredKey = key;
    if (log) log(`Scheduler firing at ${key}`);
    try {
      const result = await onFire();
      lastRunAt = new Date().toISOString();
      lastRunOk = true;
      lastRunMessage = (result && result.message) || "Scheduled run completed.";
    } catch (err) {
      lastRunAt = new Date().toISOString();
      lastRunOk = false;
      lastRunMessage = err.message || String(err);
      if (log) log(`Scheduler run failed: ${lastRunMessage}`);
    }
  }

  function start() {
    if (timer) return;
    timer = setInterval(() => {
      tick().catch(() => {});
    }, 20 * 1000);
    if (log) log("Scheduler started (checks every 20s).");
  }

  function stop() {
    if (timer) clearInterval(timer);
    timer = null;
  }

  return { start, stop, status, tick };
}

module.exports = { createScheduler, nextRunFrom };
