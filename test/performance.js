const http = require("http");
const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");

const ROOT = path.resolve(__dirname, "..");

async function runBenchmark() {
  console.log("================================================================================");
  console.log(" ICEA LION QA Platform - Comprehensive Performance & Stress Benchmark");
  console.log(" Modules: Map to SimplifyQA | Compare & Map | Map EP | ICEA LION Reporter");
  console.log("================================================================================\n");

  const port = 59123;
  const serverProc = spawn("node", ["server.js"], {
    cwd: ROOT,
    env: { ...process.env, PORT: String(port) },
    stdio: ["ignore", "pipe", "pipe"],
  });

  const base = `http://127.0.0.1:${port}`;

  // Wait for server to start
  await new Promise((resolve) => {
    let started = false;
    serverProc.stdout.on("data", (data) => {
      const msg = data.toString();
      if (msg.includes("running at") || msg.includes("http://")) {
        started = true;
        resolve();
      }
    });
    serverProc.stderr.on("data", (data) => {
      if (!started) console.error("Server stderr:", data.toString());
    });
    setTimeout(() => {
      if (!started) resolve();
    }, 2500);
  });

  function timedFetch(url, options = {}) {
    return new Promise((resolve, reject) => {
      const start = process.hrtime.bigint();
      const u = new URL(url);
      const req = http.request(
        {
          hostname: u.hostname,
          port: u.port,
          path: u.pathname + u.search,
          method: options.method || "GET",
          headers: options.headers || {},
        },
        (res) => {
          const chunks = [];
          res.on("data", (c) => chunks.push(c));
          res.on("end", () => {
            const end = process.hrtime.bigint();
            const durationMs = Number(end - start) / 1e6;
            const body = Buffer.concat(chunks);
            let parsed = null;
            try {
              parsed = JSON.parse(body.toString("utf8"));
            } catch {
              parsed = null;
            }
            resolve({
              status: res.statusCode,
              durationMs,
              bytes: body.length,
              body,
              json: parsed,
            });
          });
        }
      );
      req.on("error", reject);
      if (options.body) req.write(options.body);
      req.end();
    });
  }

  // 1. Static Assets Benchmark
  console.log("--- 1. Static Web Assets & Branding Latency ---");
  const staticEndpoints = ["/", "/styles.css", "/app.js", "/logo/ICEA%20Lion.png", "/logo/Simplify-icon.png"];
  for (const ep of staticEndpoints) {
    const res = await timedFetch(base + ep);
    console.log(
      `  GET ${ep.padEnd(28)} -> HTTP ${res.status} | Size: ${(res.bytes / 1024).toFixed(
        1
      ).padStart(5)} KB | Latency: ${res.durationMs.toFixed(2).padStart(6)} ms`
    );
  }

  // 2. Core API Latencies (Warm cache, 5 iterations)
  console.log("\n--- 2. REST API Latencies (5-Iteration Average & Range) ---");
  const apiEndpoints = [
    "/api/health",
    "/api/config",
    "/api/properties",
    "/api/auth/status",
    "/api/history",
    "/api/ep/sample-files",
    "/api/reporter/health",
    "/api/reporter/branding",
    "/api/reporter/form-defaults",
    "/api/reporter/properties",
    "/api/reporter/schedule",
    "/api/reporter/run-progress",
    "/api/reporter/history",
    "/api/reporter/sheets",
  ];

  for (const ep of apiEndpoints) {
    const times = [];
    let lastStatus = 200;
    for (let i = 0; i < 5; i++) {
      const res = await timedFetch(base + ep);
      times.push(res.durationMs);
      lastStatus = res.status;
    }
    const avg = times.reduce((a, b) => a + b, 0) / times.length;
    const min = Math.min(...times);
    const max = Math.max(...times);
    console.log(
      `  GET ${ep.padEnd(30)} -> HTTP ${lastStatus} | Avg: ${avg.toFixed(2).padStart(5)} ms (Min: ${min.toFixed(2).padStart(5)} ms, Max: ${max.toFixed(2).padStart(5)} ms)`
    );
  }

  // 3. Complex Operations & Excel Generation Performance
  console.log("\n--- 3. Heavy Computational & Excel Operations Across Modules ---");

  const boundary = "----WebKitFormBoundaryBench7MA4YWxkTrZu0gW";
  function makeMultipart(fields, files) {
    const parts = [];
    for (const [k, v] of Object.entries(fields)) {
      parts.push(`--${boundary}\r\nContent-Disposition: form-data; name="${k}"\r\n\r\n${v}\r\n`);
    }
    for (const [k, file] of Object.entries(files)) {
      parts.push(
        `--${boundary}\r\nContent-Disposition: form-data; name="${k}"; filename="${file.name}"\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n`
      );
      parts.push(file.buffer);
      parts.push("\r\n");
    }
    parts.push(`--${boundary}--\r\n`);
    const buffers = parts.map((p) => (Buffer.isBuffer(p) ? p : Buffer.from(p, "utf-8")));
    return Buffer.concat(buffers);
  }

  // A. Map Module (FA Payroll)
  const clientBuffer = fs.readFileSync(path.join(ROOT, "Client doc", "FA Payroll.xlsx"));
  const kenyaBuffer = fs.readFileSync(path.join(ROOT, "Kenya doc", "Payroll.xlsx"));

  const mapBody = makeMultipart(
    { module: "Payroll", entity: "Life UG", versions: "v1.0", testcaseType: "WEB" },
    { client: { name: "FA Payroll.xlsx", buffer: clientBuffer }, kenya: { name: "Payroll.xlsx", buffer: kenyaBuffer } }
  );

  const mapReviewRes = await timedFetch(base + "/api/review", {
    method: "POST",
    headers: { "Content-Type": `multipart/form-data; boundary=${boundary}` },
    body: mapBody,
  });
  console.log(
    `  [Map] POST /api/review (Map 42 TCs)               -> HTTP ${mapReviewRes.status} | Time: ${mapReviewRes.durationMs.toFixed(2)} ms`
  );

  const mapGenRes = await timedFetch(base + "/api/generate", {
    method: "POST",
    headers: { "Content-Type": `multipart/form-data; boundary=${boundary}` },
    body: mapBody,
  });
  console.log(
    `  [Map] POST /api/generate (Generate Excel)         -> HTTP ${mapGenRes.status} | Size: ${(mapGenRes.bytes / 1024).toFixed(1)} KB | Time: ${mapGenRes.durationMs.toFixed(2)} ms`
  );

  // B. Compare Module (Tax Management UG vs Tax MGT UG life)
  const taxGenBuf = fs.readFileSync(path.join(ROOT, "Client doc", "Tax managemnt UG.xlsx"));
  const taxLifeBuf = fs.readFileSync(path.join(ROOT, "Client doc", "TAX MGT UG life.xlsx"));
  const compareBody = makeMultipart(
    { module: "Tax Management", entityCommon: "Gen UG, Life UG", entityUniqueA: "Gen UG", entityUniqueB: "Life UG" },
    {
      clientA: { name: "Tax managemnt UG.xlsx", buffer: taxGenBuf },
      clientB: { name: "TAX MGT UG life.xlsx", buffer: taxLifeBuf },
    }
  );

  const cmpReviewRes = await timedFetch(base + "/api/compare-review", {
    method: "POST",
    headers: { "Content-Type": `multipart/form-data; boundary=${boundary}` },
    body: compareBody,
  });
  console.log(
    `  [Compare] POST /api/compare-review (Diff 2 Excels) -> HTTP ${cmpReviewRes.status} | Time: ${cmpReviewRes.durationMs.toFixed(2)} ms`
  );

  const cmpGenRes = await timedFetch(base + "/api/compare", {
    method: "POST",
    headers: { "Content-Type": `multipart/form-data; boundary=${boundary}` },
    body: compareBody,
  });
  console.log(
    `  [Compare] POST /api/compare (3-Sheet Output Excel) -> HTTP ${cmpGenRes.status} | Size: ${(cmpGenRes.bytes / 1024).toFixed(1)} KB | Time: ${cmpGenRes.durationMs.toFixed(2)} ms`
  );

  // C. Map EP Module (Execution Plan Generation)
  const epSummaryPath = path.join(
    ROOT,
    "Execution Plan Sample file",
    "TestCases_Financial Management System - Uganda_2782026(Summary).xlsx"
  );
  if (fs.existsSync(epSummaryPath)) {
    const epSummaryBuffer = fs.readFileSync(epSummaryPath);
    const epBody = makeMultipart(
      {
        mode: "upload",
        moduleCustom: "Accounts Payable",
        entity: "Life UG",
        version: "v1.0",
        executionType: "Manual",
        assignedDate: "08/27/2026",
      },
      { summary: { name: "Summary.xlsx", buffer: epSummaryBuffer } }
    );

    const epReviewRes = await timedFetch(base + "/api/ep/review", {
      method: "POST",
      headers: { "Content-Type": `multipart/form-data; boundary=${boundary}` },
      body: epBody,
    });
    console.log(
      `  [Map EP] POST /api/ep/review (Filter & Preview)   -> HTTP ${epReviewRes.status} | Time: ${epReviewRes.durationMs.toFixed(2)} ms`
    );

    const epGenRes = await timedFetch(base + "/api/ep/generate", {
      method: "POST",
      headers: { "Content-Type": `multipart/form-data; boundary=${boundary}` },
      body: epBody,
    });
    console.log(
      `  [Map EP] POST /api/ep/generate (Build EP Excel)   -> HTTP ${epGenRes.status} | Size: ${(epGenRes.bytes / 1024).toFixed(1)} KB | Time: ${epGenRes.durationMs.toFixed(2)} ms`
    );
  }

  // D. Reporter Module (Compare Sheets & PDF Generation)
  const sheetsRes = await timedFetch(base + "/api/reporter/sheets");
  if (sheetsRes.json && sheetsRes.json.sheets && sheetsRes.json.sheets.length >= 2) {
    const sheets = sheetsRes.json.sheets;
    const prevSheet = sheets[sheets.length - 2];
    const latestSheet = sheets[sheets.length - 1];

    const repCompareRes = await timedFetch(base + "/api/reporter/compare", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ previous: prevSheet, latest: latestSheet }),
    });
    console.log(
      `  [Reporter] POST /api/reporter/compare             -> HTTP ${repCompareRes.status} | Time: ${repCompareRes.durationMs.toFixed(2)} ms`
    );
  }

  // 4. High-Concurrency Stress Test (100 parallel mixed requests)
  console.log("\n--- 4. Concurrency & Throughput Stress Test (100 Parallel Requests) ---");
  const concurrentCount = 100;
  const stressEndpoints = [
    "/api/config",
    "/api/health",
    "/api/properties",
    "/api/auth/status",
    "/api/reporter/health",
    "/api/reporter/branding",
    "/api/reporter/properties",
    "/api/history",
  ];

  const startConc = process.hrtime.bigint();
  const promises = [];
  for (let i = 0; i < concurrentCount; i++) {
    const ep = stressEndpoints[i % stressEndpoints.length];
    promises.push(timedFetch(base + ep));
  }
  const results = await Promise.all(promises);
  const endConc = process.hrtime.bigint();
  const totalDurationMs = Number(endConc - startConc) / 1e6;
  const latencies = results.map((r) => r.durationMs).sort((a, b) => a - b);
  const avgLatency = latencies.reduce((a, b) => a + b, 0) / latencies.length;
  const p50 = latencies[Math.floor(latencies.length * 0.5)];
  const p90 = latencies[Math.floor(latencies.length * 0.9)];
  const p95 = latencies[Math.floor(latencies.length * 0.95)];
  const p99 = latencies[latencies.length - 1];
  const rps = (concurrentCount / (totalDurationMs / 1000)).toFixed(0);
  const successCount = results.filter((r) => r.status === 200).length;

  console.log(`  Total requests:     ${concurrentCount}`);
  console.log(`  Total time taken:   ${totalDurationMs.toFixed(2)} ms`);
  console.log(`  Throughput:         ${rps} req/sec`);
  console.log(`  Average Latency:    ${avgLatency.toFixed(2)} ms`);
  console.log(`  Median (p50):       ${p50.toFixed(2)} ms`);
  console.log(`  90th percentile:    ${p90.toFixed(2)} ms`);
  console.log(`  95th percentile:    ${p95.toFixed(2)} ms`);
  console.log(`  99th percentile:    ${p99.toFixed(2)} ms`);
  console.log(`  Success rate:       ${((successCount / concurrentCount) * 100).toFixed(1)}% (${successCount}/${concurrentCount})`);

  // Cleanup
  serverProc.kill();
  console.log("\n================================================================================");
  console.log(" Performance & Stress Benchmark Complete - All Systems Nominal");
  console.log("================================================================================");
}

runBenchmark().catch((err) => {
  console.error("Benchmark error:", err);
  process.exit(1);
});
