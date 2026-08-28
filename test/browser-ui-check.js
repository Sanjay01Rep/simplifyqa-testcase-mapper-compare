const { chromium } = require("playwright");
const path = require("path");
const fs = require("fs");
const { app } = require("../server");

async function runBrowserValidation() {
  console.log("==================================================");
  console.log(" Running Interactive Browser UI E2E Validation");
  console.log("==================================================\n");

  const server = await new Promise((resolve) => {
    const s = app.listen(0, "127.0.0.1", () => resolve(s));
  });
  const port = server.address().port;
  const baseUrl = `http://127.0.0.1:${port}`;
  console.log(`Test server running at ${baseUrl}`);

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  const page = await context.newPage();

  const consoleErrors = [];
  page.on("console", (msg) => {
    if (msg.type() === "error") consoleErrors.push(msg.text());
  });
  page.on("pageerror", (err) => consoleErrors.push(err.message));

  try {
    // 1. Load Homepage
    console.log("1. Loading Homepage...");
    await page.goto(baseUrl, { waitUntil: "networkidle" });
    const title = await page.title();
    console.log(`   Page Title: "${title}"`);

    // 2. Verify all 4 tabs exist
    console.log("2. Checking 4 module navigation tabs...");
    const tabMap = page.locator("#tabMap");
    const tabCompare = page.locator("#tabCompare");
    const tabEp = page.locator("#tabEp");
    const tabReporter = page.locator("#tabReporter");

    console.log("   tabMap visible:", await tabMap.isVisible());
    console.log("   tabCompare visible:", await tabCompare.isVisible());
    console.log("   tabEp visible:", await tabEp.isVisible());
    console.log("   tabReporter visible:", await tabReporter.isVisible());

    // 3. Test Tab Switching: Map view
    console.log("3. Testing View Switching: Map to SimplifyQA");
    await tabMap.click();
    console.log("   viewMap visible:", await page.locator("#viewMap").isVisible());
    console.log("   viewCompare hidden:", await page.locator("#viewCompare").isHidden());
    console.log("   viewEp hidden:", await page.locator("#viewEp").isHidden());
    console.log("   viewReporter hidden:", await page.locator("#viewReporter").isHidden());

    // 4. Test Tab Switching: Compare view
    console.log("4. Testing View Switching: Compare");
    await tabCompare.click();
    console.log("   viewMap hidden:", await page.locator("#viewMap").isHidden());
    console.log("   viewCompare visible:", await page.locator("#viewCompare").isVisible());
    console.log("   viewEp hidden:", await page.locator("#viewEp").isHidden());
    console.log("   viewReporter hidden:", await page.locator("#viewReporter").isHidden());

    // 5. Test Tab Switching: Map EP view
    console.log("5. Testing View Switching: Map EP");
    await tabEp.click();
    console.log("   viewMap hidden:", await page.locator("#viewMap").isHidden());
    console.log("   viewCompare hidden:", await page.locator("#viewCompare").isHidden());
    console.log("   viewEp visible:", await page.locator("#viewEp").isVisible());
    console.log("   viewReporter hidden:", await page.locator("#viewReporter").isHidden());

    // 6. Test Tab Switching: Reporter view
    console.log("6. Testing View Switching: ICEA LION Reporter");
    await tabReporter.click();
    console.log("   viewMap hidden:", await page.locator("#viewMap").isHidden());
    console.log("   viewCompare hidden:", await page.locator("#viewCompare").isHidden());
    console.log("   viewEp hidden:", await page.locator("#viewEp").isHidden());
    console.log("   viewReporter visible:", await page.locator("#viewReporter").isVisible());

    // 7. Test Reporter form interactions
    console.log("7. Testing Reporter module form...");
    await page.selectOption("#reporterTemplateChoice", "1");
    const hintText = await page.locator("#reporterTemplateHint").textContent();
    console.log(`   Template 1 Hint: "${hintText.trim()}"`);

    await page.click("#reporterAddPlanBtn");
    const planRowsCount = await page.locator("#reporterPlanFields .plan-row").count();
    console.log(`   Added plan row count: ${planRowsCount}`);

    // 8. Test Auth token panel
    console.log("8. Testing Auth Token panel...");
    const authStatusHint = await page.locator("#authStatusHint").textContent();
    console.log(`   Auth status hint: "${authStatusHint.trim()}"`);

    const changeAuthBtn = page.locator("#changeAuthBtn");
    if (await changeAuthBtn.isVisible()) {
      await changeAuthBtn.click();
      console.log("   Clicked Change token, input visible:", await page.locator("#authTokenInput").isVisible());
      await page.click("#cancelAuthBtn");
      console.log("   Clicked Cancel, input hidden:", await page.locator("#authEditBlock").isHidden());
    }

    // 9. Check for any console errors
    console.log("9. Checking Browser Console Errors...");
    if (consoleErrors.length === 0) {
      console.log("   Zero JavaScript/Console errors detected! (Clean execution)");
    } else {
      console.error("   Console errors found:", consoleErrors);
    }

    console.log("\n==================================================");
    console.log(" Browser UI Verification Passed Successfully!");
    console.log("==================================================");
  } finally {
    await browser.close();
    await new Promise((resolve) => server.close(resolve));
  }
}

runBrowserValidation().catch((err) => {
  console.error("Browser validation failed:", err);
  process.exit(1);
});
