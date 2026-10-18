// Fresh-profile extension test: no site-access settings or permission requests.
const { chromium } = require(process.env.PLAYWRIGHT_PATH || 'playwright');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const assert = require('node:assert/strict');
(async () => {
  const root = path.resolve(process.env.EXTENSION_ROOT || path.join(__dirname, '..'));
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'e3-site-scope-'));
  let context;
  try {
    context = await chromium.launchPersistentContext(profile, {
      channel: 'chromium', headless: true,
      args: [`--disable-extensions-except=${root}`, `--load-extension=${root}`]
    });
    await context.route('https://**/*', route => route.fulfill({
      contentType: 'text/html', body: '<!doctype html><html><head></head><body>Offline fixture</body></html>'
    }));
    const worker = context.serviceWorkers()[0] || await context.waitForEvent('serviceworker');
    const allowed = ['https://e3.nycu.edu.tw/', 'https://e3p.nycu.edu.tw/', 'https://timetable.nycu.edu.tw/', 'https://portal.nycu.edu.tw/'];
    const denied = ['https://github.com/', 'https://www.nycu.edu.tw/', 'https://portal.nycu.edu.tw.evil.test/', 'https://timetable.nycu.edu.tw.evil.test/'];
    await worker.evaluate(async () => {
      await chrome.storage.local.set({ lastSeenVersion: chrome.runtime.getManifest().version,
        assignments: [], courses: [], announcements: [], lastSyncTime: Date.now() });
    });
    for (const url of [...allowed, ...denied]) {
      const granted = await worker.evaluate(origin => chrome.permissions.contains({ origins: [origin + '*'] }), url);
      assert.equal(granted, allowed.includes(url), `Default host access: ${url}`);
      const page = await context.newPage();
      await page.goto(url);
      if (allowed.includes(url)) {
        await page.locator('.e3-helper-sidebar-toggle').waitFor();
        await page.locator('.e3-helper-sidebar-toggle').click();
        await page.locator('.e3-helper-sidebar.expanded').waitFor();
      } else {
        await page.waitForTimeout(300);
        assert.equal(await page.locator('.e3-helper-sidebar-toggle, .e3-helper-sidebar').count(), 0, url);
      }
      await page.close();
    }
    console.log('Fresh install passed: E3 + timetable + portal granted and sidebar visible by default; unrelated sites denied and no sidebar.');
  } finally {
    await context?.close();
    fs.rmSync(profile, { recursive: true, force: true });
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
