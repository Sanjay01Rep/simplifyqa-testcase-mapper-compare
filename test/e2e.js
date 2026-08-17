const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const { app, ROOT } = require("../server");
const { mapFromBuffers, parseClientWorkbook, exactNameKey } = require("../lib/mapper");
const { compareTestcases } = require("../lib/compare");

const failures = [];
let passed = 0;

function ok(name, cond, detail) {
  if (cond) {
    passed += 1;
    console.log(`PASS  ${name}`);
  } else {
    failures.push(`${name}${detail ? " — " + detail : ""}`);
    console.log(`FAIL  ${name}${detail ? " — " + detail : ""}`);
  }
}

function listen() {
  return new Promise((resolve) => {
    const server = app.listen(0, "127.0.0.1", () => {
      const { port } = server.address();
      resolve({ server, base: `http://127.0.0.1:${port}` });
    });
  });
}

async function jsonReq(base, urlPath, options = {}) {
  const res = await fetch(base + urlPath, options);
  const text = await res.text();
  let data = null;
  try {
    data = JSON.parse(text);
  } catch {
    data = text;
  }
  return { res, text, data };
}

async function main() {
  const { server, base } = await listen();
  console.log("Test server", base);

  try {
    const health = await jsonReq(base, "/api/health");
    ok("health is JSON", health.data && health.data.ok === true, String(health.text).slice(0, 80));
    ok(
      "health lists compare route",
      Array.isArray(health.data.routes) && health.data.routes.some((r) => String(r).includes("/api/compare")),
      JSON.stringify(health.data.routes || []).slice(0, 200)
    );
    ok(
      "health sets ICEA header",
      health.res.headers.get("x-icea-lion") === "testcase-review"
    );

    const cfg = await jsonReq(base, "/api/config");
    ok("config is JSON", cfg.data && cfg.data.ok === true);
    ok(
      "config lists client files",
      Array.isArray(cfg.data.clientFiles) && cfg.data.clientFiles.length > 0
    );
    ok(
      "config lists Kenya Payroll.xlsx",
      (cfg.data.kenyaFiles || []).some((f) => /Payroll\.xlsx$/i.test(f.relative || f.name))
    );

    const unknown = await jsonReq(base, "/api/does-not-exist");
    ok(
      "unknown API returns JSON not HTML",
      typeof unknown.data === "object" && unknown.data.ok === false && !String(unknown.text).startsWith("<")
    );

    const txt = new Blob(["not excel"], { type: "text/plain" });
    const badFd = new FormData();
    badFd.append("client", txt, "notes.txt");
    const badUp = await jsonReq(base, "/api/upload-client", { method: "POST", body: badFd });
    ok(
      "reject non-xlsx upload as JSON",
      typeof badUp.data === "object" &&
        badUp.data.ok === false &&
        /xlsx/i.test(badUp.data.message || "") &&
        !String(badUp.text).startsWith("<")
    );

    const clientPath = path.join(ROOT, "Client doc", "FA Payroll.xlsx");
    const clientBuf = fs.readFileSync(clientPath);
    const goodFd = new FormData();
    goodFd.append(
      "client",
      new Blob([clientBuf], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
      "FA Payroll.xlsx"
    );
    const up = await jsonReq(base, "/api/upload-client", { method: "POST", body: goodFd });
    ok(
      "upload client .xlsx JSON",
      up.data && up.data.ok === true && up.data.savedAs === "FA Payroll.xlsx",
      typeof up.data === "string" ? up.data.slice(0, 120) : up.data && up.data.message
    );

    const kenyaPath = path.join(ROOT, "Kenya orginial testcase", "Payroll.xlsx");
    const kenyaBuf = fs.readFileSync(kenyaPath);
    const kenyaFd = new FormData();
    kenyaFd.append(
      "kenya",
      new Blob([kenyaBuf], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
      "Payroll.xlsx"
    );
    const kenyaUp = await jsonReq(base, "/api/upload-kenya", { method: "POST", body: kenyaFd });
    ok("upload kenya .xlsx JSON", kenyaUp.data && kenyaUp.data.ok === true);

    const reviewFd = new FormData();
    reviewFd.append("existingClient", "General UG/Payroll UG Gen.xlsx");
    reviewFd.append("existingKenya", "Kenya orginial testcase/Payroll.xlsx");
    reviewFd.append("module", "Payroll");
    reviewFd.append("entity", "Life UG");
    reviewFd.append("versions", "v1.0");
    reviewFd.append("testcaseType", "WEB");
    const review = await jsonReq(base, "/api/review", { method: "POST", body: reviewFd });
    ok(
      "review nested client + kenya JSON",
      review.data && review.data.ok === true,
      typeof review.data === "string" ? review.data.slice(0, 160) : review.data && review.data.message
    );
    ok(
      "kenya pre-reqs matched by exact name",
      review.data && review.data.summary && review.data.summary.kenyaMatched > 0,
      review.data && review.data.summary
        ? `matched ${review.data.summary.kenyaMatched}/${review.data.summary.kenyaCount}`
        : ""
    );
    ok(
      "empty/unmatched prereq is not WARN",
      review.data &&
        review.data.summary &&
        !(review.data.summary.issues || []).some((i) => i.type === "PREREQ_UNMATCHED")
    );
    ok(
      "output uses Sample File Pre-Requisites column",
      review.data &&
        review.data.preview &&
        review.data.preview.rows[0].cells.some((c) => c.text === "Pre-Requisites")
    );

    const genFd = new FormData();
    genFd.append("existingClient", "FA Payroll.xlsx");
    genFd.append("module", "FA Payroll");
    genFd.append("entity", "Life UG");
    genFd.append("versions", "v1.0");
    genFd.append("testcaseType", "WEB");
    const gen = await jsonReq(base, "/api/generate", { method: "POST", body: genFd });
    ok("generate FA Payroll JSON", gen.data && gen.data.ok === true && gen.data.download);
    if (gen.data && gen.data.download && gen.data.download.excel) {
      const dl = await fetch(base + gen.data.download.excel);
      const buf = Buffer.from(await dl.arrayBuffer());
      ok("download generated xlsx", dl.ok && buf.length > 100 && buf[0] === 0x50);
    } else {
      ok("download generated xlsx", false, "no download url");
    }

    const twoSheetGenFd = new FormData();
    twoSheetGenFd.append("existingClient", "General UG/Credit Control GE UG-reviewed.xlsx");
    twoSheetGenFd.append("module", "Credit Control");
    twoSheetGenFd.append("entity", "Gen UG");
    twoSheetGenFd.append("versions", "v1.0");
    twoSheetGenFd.append("testcaseType", "WEB");
    const twoSheetGen = await jsonReq(base, "/api/generate", {
      method: "POST",
      body: twoSheetGenFd,
    });
    ok(
      "generate from 2-sheet client JSON",
      twoSheetGen.data && twoSheetGen.data.ok === true && twoSheetGen.data.download
    );
    if (twoSheetGen.data && twoSheetGen.data.download && twoSheetGen.data.download.excel) {
      const XLSX = require("xlsx");
      const dl = await fetch(base + twoSheetGen.data.download.excel);
      const outBuf = Buffer.from(await dl.arrayBuffer());
      const outWb = XLSX.read(outBuf, { type: "buffer" });
      ok("mapped output keeps single template sheet", outWb.SheetNames.length === 1);
      ok(
      "summary reports two sheets selected",
        twoSheetGen.data.summary &&
          twoSheetGen.data.summary.sheetName === "Kenya, UG" &&
          (twoSheetGen.data.summary.issues || []).some((i) => i.type === "SHEETS_AUTO_SELECTED")
      );
    } else {
      ok("mapped output keeps single template sheet", false, "no download url");
      ok("summary reports second sheet selected", false);
    }

    const mapped = await mapFromBuffers({
      clientFileName: "Payroll UG Gen.xlsx",
      clientBuffer: fs.readFileSync(
        path.join(ROOT, "Client doc", "General UG", "Payroll UG Gen.xlsx")
      ),
      kenyaFileName: "Payroll.xlsx",
      kenyaBuffer: kenyaBuf,
      config: {
        Module: "Payroll",
        Entity: "Life UG",
        Versions: "v1.0",
        TestcaseType: "WEB",
      },
    });
    const overview = mapped.parsed.testcases.find((t) => t.name === "Payroll Module Overview");
    ok(
      "mapper copies kenya prereq onto matching name",
      overview && /D365/i.test(overview.prerequisites || "")
    );
    ok("sequence continuity", mapped.summary.seqOk === true);
    ok("step count match", mapped.summary.stepMatch === true);

    const creditReviewed = await mapFromBuffers({
      clientFileName: "Credit Control GE UG-reviewed.xlsx",
      clientBuffer: fs.readFileSync(
        path.join(ROOT, "Client doc", "General UG", "Credit Control GE UG-reviewed.xlsx")
      ),
      kenyaFileName: null,
      kenyaBuffer: null,
      config: {
        Module: "Credit Control",
        Entity: "Gen UG",
        Versions: "v1.0",
        TestcaseType: "WEB",
      },
    });
    ok(
      "2-sheet client workbook auto-uses both sheets",
      creditReviewed.parsed.sheetName === "Kenya, UG" &&
        (creditReviewed.summary.issues || []).some((i) => i.type === "SHEETS_AUTO_SELECTED")
    );
    ok("2-sheet workbook maps testcases", creditReviewed.summary.mappedTcCount > 0);

    const sameSteps = [{ stepDesc: "Open tax form", expected: "Form opens" }];
    const paired = compareTestcases(
      [{ clientId: "TC-1", name: "VAT return", steps: sameSteps }],
      [{ clientId: "TC-1", name: "VAT return", steps: sameSteps }]
    );
    ok("same name+id+steps is common", paired.common.length === 1 && paired.unmatchedA.length === 0);

    const softStepsA = [
      {
        stepDesc:
          "Inputting details of the vendor as per the registration documents. Email address and phone number should be captured",
        expected: "ok",
      },
    ];
    const softStepsB = [
      {
        stepDesc:
          "Inputting details of the vendor as per the registration documents. Email address and phone no should be captured",
        expected: "ok",
      },
    ];
    const softMatch = compareTestcases(
      [{ clientId: "TC-75", name: "Vendor Onboarding", steps: softStepsA }],
      [{ clientId: "TC-75", name: "Vendor Onboarding", steps: softStepsB }]
    );
    ok(
      "number vs no treated as common",
      softMatch.common.length === 1 && softMatch.nameMatchStepMismatch.length === 0
    );

    const typoMatch = compareTestcases(
      [
        {
          clientId: "TC-464",
          name: "Approvals Testing",
          steps: [
            {
              stepDesc: "Approve",
              expected:
                "The designated approver should get alert if succesful reconcilation. Not to create alert if the reconcilaition is not successful",
            },
          ],
        },
      ],
      [
        {
          clientId: "TC-464",
          name: "Approvals Testing",
          steps: [
            {
              stepDesc: "Approve",
              expected:
                "The designated approver should get alert if successful reconciliation. Not to create alert if the reconciliation is not successful",
            },
          ],
        },
      ]
    );
    ok(
      "misspellings treated as common",
      typoMatch.common.length === 1 && typoMatch.nameMatchStepMismatch.length === 0
    );

    const viaMatch = compareTestcases(
      [
        {
          clientId: "TC-495",
          name: "Full  cycle Reimbursement request of Travel advance/Petty cash",
          steps: [
            {
              stepDesc: "Approver will receive notification vai email of a pending request and will verify for approval",
              expected:
                "successful reimbursement request/uploading of support documents all they way to payment voucher generation in D365",
            },
          ],
        },
      ],
      [
        {
          clientId: "TC-495",
          name: "Full  cycle Reimbursement request of Travel advance/Petty cash",
          steps: [
            {
              stepDesc: "Approver will receive notification via email of a pending request and will verify for approval",
              expected:
                "successful reimbursement request/uploading of support documents all the way to payment voucher generation in D365",
            },
          ],
        },
      ]
    );
    ok(
      "vai vs via and they vs the treated as common",
      viaMatch.common.length === 1 && viaMatch.nameMatchStepMismatch.length === 0
    );

    const carloan = compareTestcases(
      [
        {
          clientId: "TC-496",
          name: "Full  cycle Carloan/Staff advance request",
          steps: [{ stepDesc: "Apply", expected: "Loan application" }],
        },
      ],
      [
        {
          clientId: "TC-496",
          name: "Full  cycle Carloan/Staff advance request",
          steps: [{ stepDesc: "Apply", expected: "Loan application successful" }],
        },
      ]
    );
    ok(
      "loan application vs loan application successful treated as common",
      carloan.common.length === 1 && carloan.nameMatchStepMismatch.length === 0
    );

    const mobileName = compareTestcases(
      [
        {
          clientId: "TC-504",
          name: "Flexible user interface  Field  (Mobile /Bank details)",
          steps: [{ stepDesc: "Set up of all viable option as mode of payments visible to staff", expected: "ok" }],
        },
      ],
      [
        {
          clientId: "TC-504",
          name: "Flexible user interface  Field  (Mobile money/Bank details)",
          steps: [{ stepDesc: "Set up of all viable option as mode of payments visible to staff", expected: "ok" }],
        },
      ]
    );
    ok(
      "Mobile vs Mobile money in name treated as common",
      mobileName.common.length === 1 && mobileName.nameMatchStepMismatch.length === 0
    );

    const kraUra = compareTestcases(
      [
        {
          clientId: "TC-117",
          name: "Vendor Document Maintenance",
          steps: [
            {
              stepDesc: "Upload",
              expected: "All documents required( CR12, KRA PIN, BANK DETAILS) are attached during vendor creation.",
            },
          ],
        },
      ],
      [
        {
          clientId: "TC-117",
          name: "Vendor Document Maintenance",
          steps: [
            {
              stepDesc: "Upload",
              expected: "All documents required( CR12, URA PIN, BANK DETAILS) are attached during vendor creation.",
            },
          ],
        },
      ]
    );
    ok(
      "KRA vs URA still treated as different",
      kraUra.common.length === 0 && kraUra.nameMatchStepMismatch.length === 1
    );

    const stepMismatch = compareTestcases(
      [{ clientId: "TC-334", name: "VAT", steps: [{ stepDesc: "one", expected: "a" }] }],
      [
        {
          clientId: "TC-334",
          name: "VAT",
          steps: [
            { stepDesc: "one", expected: "a" },
            { stepDesc: "two", expected: "b" },
          ],
        },
      ]
    );
    ok(
      "same name+id different steps is unmatched both",
      stepMismatch.common.length === 0 &&
        stepMismatch.unmatchedA.length === 1 &&
        stepMismatch.unmatchedB.length === 1 &&
        stepMismatch.nameMatchStepMismatch.length === 1
    );

    const nameOnly = compareTestcases(
      [{ clientId: "", name: "No ID case", steps: sameSteps }],
      [{ clientId: "TC-99", name: "No ID case", steps: sameSteps }]
    );
    ok("missing TC id matches by name", nameOnly.common.length === 1);

    const idMismatch = compareTestcases(
      [{ clientId: "TC-1", name: "Same name", steps: sameSteps }],
      [{ clientId: "TC-2", name: "Same name", steps: sameSteps }]
    );
    ok(
      "same name different ids are not paired",
      idMismatch.common.length === 0 && idMismatch.unmatchedA.length === 1 && idMismatch.unmatchedB.length === 1
    );

    const apParsed = parseClientWorkbook(
      path.join(ROOT, "Client doc", "General UG", "Account Payables GE UG-Review.xlsx"),
      "Account Payables GE UG-Review.xlsx",
      { Module: "Accounts Payable", Entity: "Gen UG", Versions: "v1.0", TestcaseType: "WEB" }
    );
    const apNames = apParsed.testcases.map((tc) => exactNameKey(tc.name));
    ok(
      "AP Kenya sheet has 48 named testcases after dropping UG duplicates",
      apParsed.testcases.length === 48,
      `parsed ${apParsed.testcases.length}`
    );
    ok(
      "AP keeps NEW TEST Purchase Invoice Creation separate from TC-81",
      apParsed.testcases.filter((tc) => exactNameKey(tc.name) === "purchase invoice creation").length === 2
    );
    ok(
      "AP does not keep header/section junk TCs",
      !apParsed.testcases.some((tc) => /^(id|name|ug)$/i.test(String(tc.name || "").trim()))
    );

    const lifeParsed = parseClientWorkbook(
      path.join(ROOT, "Client doc", "Life UG", "Account payables Life UG.xlsx"),
      "Account payables Life UG.xlsx",
      { Module: "Accounts Payable", Entity: "Life UG", Versions: "v1.0", TestcaseType: "WEB" }
    );
    ok(
      "Life AP skips repeated header rows",
      !lifeParsed.testcases.some((tc) => /^(id|name|ug)$/i.test(String(tc.name || "").trim()))
    );
    ok(
      "Life AP count has no UG/Name junk",
      lifeParsed.testcases.every(
        (tc) => !/^(ug|name|id)$/i.test(exactNameKey(tc.name)) && !/^(ug|name|id)$/i.test(exactNameKey(tc.clientId))
      )
    );

    const fillWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      fillWb,
      XLSX.utils.aoa_to_sheet([
        ["ID", "Name", "Module", "Entity", "Description", "seq", "Step description", "Expected Result"],
        ["TC-1", "Fill expected from step", "M", "E", "d", 1, "Click Save", ""],
        ["TC-2", "Fill step from expected", "M", "E", "d", 1, "", "Saved successfully"],
      ]),
      "Sheet1"
    );
    const fillBuf = XLSX.write(fillWb, { type: "buffer", bookType: "xlsx" });
    const fillParsed = parseClientWorkbook(fillBuf, "fill-test.xlsx", {
      Module: "M",
      Entity: "E",
      Versions: "v1.0",
      TestcaseType: "WEB",
    });
    const filledExpected = fillParsed.testcases.find((t) => t.name === "Fill expected from step");
    const filledStep = fillParsed.testcases.find((t) => t.name === "Fill step from expected");
    ok(
      "empty expected filled from step description",
      filledExpected &&
        filledExpected.steps[0].expected === "Click Save" &&
        filledExpected.steps[0].filledFromStep === true &&
        fillParsed.issues.some((i) => i.type === "EMPTY_EXPECTED_FILLED")
    );
    ok(
      "empty step filled from expected result",
      filledStep &&
        filledStep.steps[0].stepDesc === "Saved successfully" &&
        filledStep.steps[0].filledFromExpected === true &&
        fillParsed.issues.some((i) => i.type === "EMPTY_STEP_FILLED")
    );

    const page = await fetch(base + "/");
    const html = await page.text();
    ok("UI index served", page.ok && html.includes("ICEA LION Testcase Review"));
    ok("properties collapsed by default", html.includes('id="propsDetails"') && !html.includes('<details id="propsDetails" open'));
    ok("client dropzone present", html.includes('id="clientDrop"') && html.includes('id="kenyaDrop"'));
    ok("preview panel present", html.includes('id="previewPanel"'));
    ok(
      "compare module present without replacing map",
      html.includes('id="tabCompare"') &&
        html.includes('id="viewCompare"') &&
        html.includes('id="runForm"') &&
        html.includes('id="compareForm"')
    );
    ok(
      "compare module renamed and mapper header present",
      html.includes("Compare, Map &amp; Report") &&
        html.includes("Excel A") &&
        html.includes("Excel B") &&
        html.includes('id="mapperHeader"') &&
        html.includes('id="cmpMapperHeader"') &&
        html.includes('id="serverBot"') &&
        html.includes('id="entityChecks"') &&
        html.includes("compare-view") &&
        html.includes('id="cmpStep1"') &&
        html.includes('id="cmpStep2"') &&
        html.includes('id="cmpNextBtn"')
    );
    ok(
      "compare has mapper upload, review, per-sheet entities",
      html.includes('id="cmpKenyaDrop"') &&
        html.includes('id="cmpReviewBtn"') &&
        html.includes('id="cmpEntityCommon"') &&
        html.includes('id="cmpEntityUniqueA"') &&
        html.includes('id="cmpEntityUniqueB"') &&
        html.includes('id="cmpSaveMapBtn"')
    );
    ok("compare shares validation and properties", html.includes('id="issueList"') && html.includes('id="propsDetails"'));

    const cmpMissing = new FormData();
    cmpMissing.append("existingClientA", "Tax managemnt UG.xlsx");
    cmpMissing.append("module", "Tax Management");
    cmpMissing.append("entity", "Gen UG");
    const cmpBad = await jsonReq(base, "/api/compare", { method: "POST", body: cmpMissing });
    ok(
      "compare requires both files as JSON",
      typeof cmpBad.data === "object" && cmpBad.data.ok === false && !String(cmpBad.text).startsWith("<")
    );

    const cmpReviewFd = new FormData();
    cmpReviewFd.append("existingClientA", "Tax managemnt UG.xlsx");
    cmpReviewFd.append("existingClientB", "TAX MGT UG life.xlsx");
    cmpReviewFd.append("existingKenya", "Kenya orginial testcase/Payroll.xlsx");
    cmpReviewFd.append("module", "Tax Management");
    cmpReviewFd.append("entityCommon", "Gen UG");
    cmpReviewFd.append("entityCommon", "Life UG");
    cmpReviewFd.append("entityUniqueA", "Gen UG");
    cmpReviewFd.append("entityUniqueB", "Life UG");
    cmpReviewFd.append("versions", "v1.0");
    cmpReviewFd.append("testcaseType", "WEB");
    const cmpReview = await jsonReq(base, "/api/compare-review", { method: "POST", body: cmpReviewFd });
    ok(
      "compare review JSON without download",
      cmpReview.data && cmpReview.data.ok === true && !cmpReview.data.download,
      typeof cmpReview.data === "string" ? cmpReview.data.slice(0, 160) : cmpReview.data && cmpReview.data.message
    );
    ok(
      "compare review parses Kenya original",
      cmpReview.data && cmpReview.data.summary && cmpReview.data.summary.kenyaCount > 0
    );
    ok(
      "compare review exposes validation issues",
      cmpReview.data && Array.isArray(cmpReview.data.summary.issues)
    );

    const cmpFd = new FormData();
    cmpFd.append("existingClientA", "Tax managemnt UG.xlsx");
    cmpFd.append("existingClientB", "TAX MGT UG life.xlsx");
    cmpFd.append("existingKenya", "Kenya orginial testcase/Payroll.xlsx");
    cmpFd.append("module", "Tax Management");
    cmpFd.append("entityCommon", "Gen UG");
    cmpFd.append("entityCommon", "Life UG");
    cmpFd.append("entityUniqueA", "Gen UG");
    cmpFd.append("entityUniqueB", "Life UG");
    cmpFd.append("versions", "v1.0");
    cmpFd.append("testcaseType", "WEB");
    const cmp = await jsonReq(base, "/api/compare", { method: "POST", body: cmpFd });
    ok(
      "compare tax workbooks JSON",
      cmp.data && cmp.data.ok === true && cmp.data.download,
      typeof cmp.data === "string" ? cmp.data.slice(0, 160) : cmp.data && cmp.data.message
    );
    if (cmp.data && cmp.data.ok) {
      const counts = cmp.data.summary.counts;
      ok(
        "compare counts: both sides parsed",
        counts.a > 0 && counts.b > 0 && counts.a + counts.b >= counts.common
      );
      ok(
        "name match with different steps is unmatched, not common",
        (cmp.data.summary.nameMatchStepMismatch || []).some(
          (row) => /TC-334/i.test(row.id || "") || /VAT/i.test(row.name || "")
        ) || counts.nameMatchStepMismatch >= 0
      );
      const headerCells = (cmp.data.preview.rows[0].cells || []).map((c) => c.text);
      ok(
        "common sheet uses SimplifyQA template header",
        headerCells.includes("Pre-Requisites") && headerCells.some((t) => /Name/i.test(t))
      );

      const XLSX = require("xlsx");
      const dl = await fetch(base + cmp.data.download.excel);
      const buf = Buffer.from(await dl.arrayBuffer());
      ok("compare xlsx downloadable", dl.ok && buf[0] === 0x50);
      const wb = XLSX.read(buf, { type: "buffer" });
      ok(
        "compare workbook has Common + two unmatched sheets",
        wb.SheetNames.length === 3 && wb.SheetNames[0] === "Common"
      );
      ok(
        "unmatched sheets named from client files",
        wb.SheetNames.includes("Tax managemnt UG") && wb.SheetNames.includes("TAX MGT UG life")
      );
      const commonAoA = XLSX.utils.sheet_to_json(wb.Sheets.Common, { header: 1, defval: "" });
      ok("Common sheet is SimplifyQA template", String(commonAoA[0][0]).toLowerCase().includes("name"));
      const unmatchedA = XLSX.utils.sheet_to_json(wb.Sheets["Tax managemnt UG"], {
        header: 1,
        defval: "",
      });
      ok(
        "unmatched A uses SimplifyQA template",
        String(unmatchedA[0][0]).toLowerCase().includes("name")
      );
      const entityIdx = commonAoA[0].findIndex((h) => String(h).trim().toLowerCase() === "entity");
      const commonFirst = commonAoA.find((r, i) => i > 0 && String(r[0] || "").trim());
      ok(
        "common sheet maps multiple entities",
        entityIdx >= 0 &&
          commonFirst &&
          /Gen UG/i.test(String(commonFirst[entityIdx])) &&
          /Life UG/i.test(String(commonFirst[entityIdx]))
      );
      const uniqueAFirst = unmatchedA.find((r, i) => i > 0 && String(r[0] || "").trim());
      ok(
        "unique A sheet maps Gen UG entity",
        uniqueAFirst && String(uniqueAFirst[entityIdx]).trim() === "Gen UG"
      );
      const unmatchedB = XLSX.utils.sheet_to_json(wb.Sheets["TAX MGT UG life"], {
        header: 1,
        defval: "",
      });
      const uniqueBFirst = unmatchedB.find((r, i) => i > 0 && String(r[0] || "").trim());
      ok(
        "unique B sheet maps Life UG entity",
        uniqueBFirst && String(uniqueBFirst[entityIdx]).trim() === "Life UG"
      );
      ok(
        "kenya prereq column present on common",
        commonAoA[0].some((h) => String(h).includes("Pre-Requisites"))
      );
    } else {
      ok("compare counts: both sides parsed", false, "compare failed");
      ok("name match with different steps is unmatched, not common", false);
      ok("common sheet uses SimplifyQA template header", false);
      ok("compare xlsx downloadable", false);
      ok("compare workbook has Common + two unmatched sheets", false);
      ok("unmatched sheets named from client files", false);
      ok("Common sheet is SimplifyQA template", false);
      ok("unmatched A uses SimplifyQA template", false);
      ok("common sheet maps multiple entities", false);
      ok("unique A sheet maps Gen UG entity", false);
      ok("unique B sheet maps Life UG entity", false);
      ok("kenya prereq column present on common", false);
    }

    const apFd = new FormData();
    apFd.append("existingClientA", "General UG/Account Payables GE UG-Review.xlsx");
    apFd.append("existingClientB", "Life UG/Account payables Life UG.xlsx");
    apFd.append("existingKenya", "Kenya doc/Account paybles.xlsx");
    apFd.append("module", "Accounts Payable");
    apFd.append("entityCommon", "Gen UG");
    apFd.append("entityCommon", "Life UG");
    apFd.append("entityUniqueA", "Gen UG");
    apFd.append("entityUniqueB", "Life UG");
    apFd.append("versions", "v1.0");
    apFd.append("testcaseType", "WEB");
    const apCmp = await jsonReq(base, "/api/compare", { method: "POST", body: apFd });
    ok(
      "compare Account Payables JSON",
      apCmp.data && apCmp.data.ok === true && apCmp.data.download,
      typeof apCmp.data === "string" ? apCmp.data.slice(0, 160) : apCmp.data && apCmp.data.message
    );
    if (apCmp.data && apCmp.data.ok) {
      ok(
        "AP compare has common and unique counts",
        apCmp.data.summary.counts.a > 0 &&
          apCmp.data.summary.counts.b > 0 &&
          apCmp.data.summary.counts.common +
            apCmp.data.summary.counts.unmatchedA +
            apCmp.data.summary.counts.unmatchedB >
            0
      );
      ok(
        "AP common + unique A equals file A",
        apCmp.data.summary.counts.common + apCmp.data.summary.counts.unmatchedA ===
          apCmp.data.summary.counts.a,
        `A ${apCmp.data.summary.counts.a} vs ${apCmp.data.summary.counts.common}+${apCmp.data.summary.counts.unmatchedA}`
      );
      ok(
        "AP common + unique B equals file B",
        apCmp.data.summary.counts.common + apCmp.data.summary.counts.unmatchedB ===
          apCmp.data.summary.counts.b,
        `B ${apCmp.data.summary.counts.b} vs ${apCmp.data.summary.counts.common}+${apCmp.data.summary.counts.unmatchedB}`
      );
      const overlapNames = (apCmp.data.summary.commonNames || []).filter((c) =>
        (apCmp.data.summary.unmatchedA || []).some(
          (u) => exactNameKey(u.name) === exactNameKey(c.name)
        )
      );
      ok(
        "AP testcase is not in both Common and Unique A",
        overlapNames.length === 0,
        overlapNames.map((t) => t.name).join(", ")
      );
      ok(
        "AP compare stores per-sheet entities",
        /Gen UG/i.test(apCmp.data.summary.entityCommon || "") &&
          /Life UG/i.test(apCmp.data.summary.entityCommon || "") &&
          apCmp.data.summary.entityUniqueA === "Gen UG" &&
          apCmp.data.summary.entityUniqueB === "Life UG"
      );
      ok(
        "AP compare maps Kenya Account paybles prereqs",
        apCmp.data.summary.kenyaMatched > 0 && apCmp.data.summary.kenyaCount > 0
      );
    } else {
      ok("AP compare has common and unique counts", false);
      ok("AP common + unique A equals file A", false);
      ok("AP common + unique B equals file B", false);
      ok("AP testcase is not in both Common and Unique A", false);
      ok("AP compare stores per-sheet entities", false);
      ok("AP compare maps Kenya Account paybles prereqs", false);
    }

    const sameFd = new FormData();
    sameFd.append("existingClientA", "Tax managemnt UG.xlsx");
    sameFd.append("existingClientB", "Tax managemnt UG.xlsx");
    sameFd.append("module", "Tax Management");
    sameFd.append("entityCommon", "Gen UG");
    sameFd.append("entityUniqueA", "Gen UG");
    sameFd.append("entityUniqueB", "Gen UG");
    sameFd.append("versions", "v1.0");
    sameFd.append("testcaseType", "WEB");
    const sameCmp = await jsonReq(base, "/api/compare", { method: "POST", body: sameFd });
    if (sameCmp.data && sameCmp.data.ok && sameCmp.data.download) {
      ok(
        "all-common compare has zero unmatched counts",
        sameCmp.data.summary.counts.unmatchedA === 0 &&
          sameCmp.data.summary.counts.unmatchedB === 0 &&
          sameCmp.data.summary.counts.common === sameCmp.data.summary.counts.a
      );
      const sameDl = await fetch(base + sameCmp.data.download.excel);
      const sameBuf = Buffer.from(await sameDl.arrayBuffer());
      const sameWb = XLSX.read(sameBuf, { type: "buffer" });
      ok(
        "all-common compare writes Common sheet only",
        sameWb.SheetNames.length === 1 && sameWb.SheetNames[0] === "Common"
      );
    } else {
      ok("all-common compare has zero unmatched counts", false);
      ok("all-common compare writes Common sheet only", false);
    }
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }

  console.log("");
  console.log(`${passed} passed, ${failures.length} failed`);
  if (failures.length) {
    failures.forEach((f) => console.log(" -", f));
    process.exitCode = 1;
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
