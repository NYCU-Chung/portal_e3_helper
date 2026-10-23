// Offline Safari-capability fixture in Chromium; real Safari device testing is separate.
const { chromium } = require(process.env.PLAYWRIGHT_PATH || 'playwright');
const assert = require('node:assert/strict');
const path = require('node:path');
(async () => {
  const browser = await chromium.launch({ headless: true });
  try {
    for (const language of ['en', 'zh-TW']) {
      const context = await browser.newContext({ viewport: { width: 768, height: 1024 }, hasTouch: true });
      const page = await context.newPage();
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      await context.route('**/*', route => route.fulfill({ contentType: 'text/html', body: '<!doctype html><title>Offline E3 fixture</title><body><a href="https://e3p.nycu.edu.tw/pluginfile.php/1/a.pdf">First.pdf</a><a href="https://e3p.nycu.edu.tw/pluginfile.php/2/b.pdf">&lt;img src=x onerror=alert(1)&gt;.pdf</a></body>' }));
      await page.addInitScript(language => {
        const data = { interfaceLanguage: language, assignments: [], courses: [], announcements: [], messages: [], lastSyncTime: Date.now(), lastSeenVersion: '2.3.0' };
        window.chrome = {
          i18n: { getUILanguage: () => language },
          runtime: { id: 'fixture', getManifest: () => ({ version: '2.3.0', permissions: ['storage'] }), getURL: value => value,
            onMessage: { addListener() {} }, sendMessage: (_, callback) => { callback?.({ success: true }); return Promise.resolve({ success: true }); } },
          storage: { onChanged: { addListener() {} }, local: {
            get: async (_, callback) => { callback?.(data); return data; },
            set: async update => Object.assign(data, update)
          } }
        };
      }, language);
      await page.goto('https://e3p.nycu.edu.tw/');
      await page.addScriptTag({ path: path.join(__dirname, '../i18n.js') });
      await page.addScriptTag({ path: path.join(__dirname, '../exam-detection.js') });
      await page.addScriptTag({ path: path.join(__dirname, '../content.js') });
      await page.locator('.e3-helper-sidebar-toggle').tap();
      await page.locator('[data-tab="downloads"]').tap();
      await page.locator('#e3-helper-scan-current').tap();
      await page.locator('.e3-helper-pdf-checkbox').nth(1).waitFor();
      await page.locator('#e3-helper-select-all').tap();
      await page.locator('#e3-helper-download-separate').tap();
      const list = page.locator('#e3-helper-safari-download-links');
      assert.equal(await list.locator('a').count(), 2);
      assert.equal(await list.locator('img').count(), 0, 'filenames must remain text');
      assert.equal(context.pages().length, 1, 'batch click must not open tabs');
      assert.equal(await page.locator('#e3-helper-download-separate').isEnabled(), true);
      assert.ok((await page.locator('.e3-helper-download-status').innerText()).includes(language === 'en' ? 'Download links are ready' : '下載連結已備妥'));
      const [popup] = await Promise.all([context.waitForEvent('page'), list.locator('a').nth(1).tap()]);
      await popup.waitForLoadState();
      assert.equal(popup.url(), 'https://e3p.nycu.edu.tw/pluginfile.php/2/b.pdf');
      await popup.close();
      await page.locator('.e3-helper-pdf-checkbox').nth(1).uncheck();
      await page.locator('#e3-helper-download-separate').tap();
      assert.equal(await list.locator('a').count(), 1);
      await page.locator('#e3-helper-scan-current').tap();
      assert.equal(await list.count(), 0, 'a new scan removes old prepared links');
      assert.deepEqual(errors, []);
      if (process.env.SCREENSHOT_PATH && language === 'en') await page.screenshot({ path: process.env.SCREENSHOT_PATH });
      await context.close();
    }
    console.log('Safari download fixture passed: bilingual links, real click navigation, safe filenames, selection replacement, no automatic tabs.');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exit(1); });
