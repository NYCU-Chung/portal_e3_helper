// Optional browser smoke test: install Playwright or set PLAYWRIGHT_PATH.
// Set CHROME_PATH to use an existing Chrome executable.
const { chromium } = require(process.env.PLAYWRIGHT_PATH || 'playwright');
const assert = require('node:assert/strict');
const path = require('node:path');
(async () => {
  const browser = await chromium.launch({ headless: true, ...(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {}) });
  try {
    const page = await browser.newPage({ viewport: { width: 1280, height: 900 }, hasTouch: true });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.route('http://e3-test.local/**', route => route.fulfill({ contentType: 'text/html', body: '<!doctype html><html><head></head><body></body></html>' }));
    await page.addInitScript(() => {
      const changes = [];
      const seed = { interfaceLanguage: 'en', assignments: [{ eventId: 'manual-1', name: '作業中文原文', course: '課程中文原文', deadline: Date.now() + 86400000, manualStatus: 'pending', isManual: true }], courses: [], announcements: [], messages: [{ id: 'msg-1', type: 'message', title: '', courseName: 'Fixture course', author: 'Fixture sender', timestamp: Date.now(), url: 'https://e3p.nycu.edu.tw/mail?m=1' }], lastSyncTime: Date.now(), lastSeenVersion: '2.2.0' };
      const read = () => JSON.parse(localStorage.getItem('fixture') || JSON.stringify(seed));
      window.chrome = {
        i18n: { getUILanguage: () => 'zh-TW' },
        runtime: { id: 'test', getManifest: () => ({ version: '2.2.0' }), getURL: p => p, onMessage: { addListener() {} }, sendMessage: (message, callback) => { if (message.action === 'openNotificationSettings') { window.openedNotificationSettings = true; callback({ success: true }); return undefined; } const result = { success: true }; callback?.(result); return Promise.resolve(result); } },
        storage: {
          onChanged: { addListener: listener => changes.push(listener) },
          local: {
            get: async (keys, callback) => { const data = read(); callback?.(data); return data; },
            set: async update => {
              const data = read();
              localStorage.setItem('fixture', JSON.stringify({ ...data, ...update }));
              const change = Object.fromEntries(Object.entries(update).map(([key, value]) => [key, { oldValue: data[key], newValue: value }]));
              changes.forEach(listener => listener(change, 'local'));
            }
          }
        }
      };
    });
    await page.goto('http://e3-test.local/');
    const inject = async () => {
      await page.addScriptTag({ path: path.join(__dirname, '../i18n.js') });
      await page.addScriptTag({ path: path.join(__dirname, '../content.js') });
      await page.waitForSelector('.e3-helper-sidebar-toggle');
      await page.locator('.e3-helper-sidebar-toggle').click();
    };
    // A callback-only failed preference read must not block actual sidebar startup.
    await page.evaluate(() => {
      const get = chrome.storage.local.get;
      chrome.storage.local.get = (keys, callback) => {
        if (keys.length === 1 && keys[0] === 'interfaceLanguage') {
          chrome.runtime.lastError = { message: 'Transient preference failure' };
          callback();
          delete chrome.runtime.lastError;
          return;
        }
        return get(keys, callback);
      };
    });
    await inject();
    assert.equal(await page.locator('[data-tab="assignments"]').innerText(), '作業');
    assert.deepEqual(errors, []);
    // Isolate the next scenario from notifications created during default-language startup.
    await page.evaluate(() => localStorage.removeItem('fixture'));
    await page.reload();
    await inject();
    const checkTabLayout = async () => {
      for (const width of [280, 350, 480, 800]) {
        const layout = await page.evaluate(width => {
          const sidebar = document.querySelector('.e3-helper-sidebar');
          sidebar.style.width = `${width}px`;
          const tabs = sidebar.querySelector('.e3-helper-tabs');
          const bounds = tabs.getBoundingClientRect();
          return {
            overflow: tabs.scrollWidth > tabs.clientWidth,
            visible: [...tabs.children].every(tab => {
              const rect = tab.getBoundingClientRect();
              return rect.left >= bounds.left && rect.right <= bounds.right &&
                rect.top >= bounds.top && rect.bottom <= bounds.bottom;
            })
          };
        }, width);
        assert.equal(layout.overflow, false, `Tabs overflow at ${width}px`);
        assert.equal(layout.visible, true, `Tab is clipped at ${width}px`);
      }
    };
    await checkTabLayout();
    assert.deepEqual((await page.locator('.e3-helper-tab').allTextContents()).map(text => text.replace(/\d+$/, '')), ['Assignments', 'Courses', 'Downloads', 'Announcements', 'Alerts', 'Help']);
    assert.ok((await page.locator('[data-content="assignments"]').innerText()).includes('作業中文原文'));
    assert.ok((await page.locator('[data-content="assignments"]').innerText()).includes('Mark as submitted'));
    for (const tab of ['grades', 'downloads', 'announcements', 'notifications', 'help']) {
      await page.locator(`[data-tab="${tab}"]`).click();
      await page.waitForTimeout(100);
      const text = await page.locator(`[data-content="${tab}"]`).innerText();
      assert.ok(text.trim(), `${tab} should render`);
      assert.ok(!/[\u3400-\u9fff]/.test(text.replaceAll('作業中文原文', '').replaceAll('課程中文原文', '')), `${tab} contains untranslated UI: ${text}`);
    }
    assert.ok((await page.locator('[data-content="announcements"]').innerText()).includes('(No subject)'));
    await page.locator('#e3-helper-more-btn').click();
    await page.locator('#e3-helper-settings-btn').click();
    assert.equal(await page.locator('#e3-helper-language').inputValue(), 'en');
    await page.locator('#e3-helper-notification-settings').click();
    assert.equal(await page.evaluate(() => window.openedNotificationSettings), true);
    assert.ok((await page.locator('#e3-helper-settings-modal').innerText()).includes('Saving a new language refreshes this page.'));
    await page.locator('#e3-helper-enable-ai').check();
    const settings = await page.locator('#e3-helper-settings-modal').innerText();
    assert.ok(!/[\u3400-\u9fff]/.test(settings.replace('繁體中文', '')), settings);
    await page.locator('#e3-helper-language').selectOption('zh-TW');
    await Promise.all([page.waitForEvent('framenavigated'), page.locator('#e3-helper-save-settings').click()]);
    await page.waitForLoadState('load');
    assert.equal(await page.evaluate(() => JSON.parse(localStorage.getItem('fixture')).interfaceLanguage), 'zh-TW');
    await inject();
    await checkTabLayout();
    assert.equal(await page.locator('[data-tab="assignments"]').innerText(), '作業');
    assert.ok((await page.locator('[data-content="assignments"]').innerText()).includes('標記為已繳交'));
    await page.locator('[data-tab="announcements"]').click();
    assert.ok((await page.locator('[data-content="announcements"]').innerText()).includes('(無主旨)'));
    assert.deepEqual(errors, []);
    console.log('Passed: six tabs, untranslated UI checks, language settings/save/reload and navigation widths.');
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
