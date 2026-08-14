const { execSync } = require("child_process");

const port = Number(process.env.PORT || 3100);

function listeningPids(portNum) {
  try {
    const out = execSync("netstat -ano", { encoding: "utf8" });
    const pids = new Set();
    for (const line of out.split(/\r?\n/)) {
      if (!line.includes(`:${portNum}`) || !/LISTENING/i.test(line)) continue;
      const parts = line.trim().split(/\s+/);
      const pid = parts[parts.length - 1];
      if (/^\d+$/.test(pid) && pid !== "0") pids.add(pid);
    }
    return [...pids];
  } catch {
    return [];
  }
}

const pids = listeningPids(port);
if (!pids.length) {
  console.log(`Port ${port} is free.`);
  process.exit(0);
}

for (const pid of pids) {
  try {
    console.log(`Port ${port} in use by PID ${pid} — stopping it...`);
    execSync(`taskkill /PID ${pid} /F`, { stdio: "inherit" });
  } catch (err) {
    console.error(`Could not stop PID ${pid}:`, err.message || err);
    process.exit(1);
  }
}

setTimeout(() => {
  const still = listeningPids(port);
  if (still.length) {
    console.error(`Port ${port} still busy (PIDs: ${still.join(", ")}).`);
    process.exit(1);
  }
  console.log(`Port ${port} is free.`);
}, 400);
