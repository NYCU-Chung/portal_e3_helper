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
      const seed = { interfaceLanguage: 'en', assignments: [{ eventId: 'manual-1', name: '作業中文原文', course: '課程中文原文', deadline: Date.now() + 86400000, manualStatus: 'pending', isManual: true }], courses: [], announcements: [], lastSyncTime: Date.now(), lastSeenVersion: '2.2.0' };
      const read = () => JSON.parse(localStorage.getItem('fixture') || JSON.stringify(seed));
      window.chrome = {
        i18n: { getUILanguage: () => 'zh-TW' },
        runtime: { id: 'test', getManifest: () => ({ version: '2.2.0' }), getURL: p => p, onMessage: { addListener() {} }, sendMessage: (message, callback) => { const result = { success: true }; callback?.(result); return Promise.resolve(result); } },
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
    await page.locator('#e3-helper-more-btn').click();
    await page.locator('#e3-helper-settings-btn').click();
    assert.equal(await page.locator('#e3-helper-language').inputValue(), 'en');
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
    await page.setViewportSize({ width: 768, height: 1024 });
    await page.locator('#e3-helper-close-btn').tap();
    await page.locator('.e3-helper-sidebar-toggle').tap();
    assert.equal(await page.locator('.e3-helper-sidebar').evaluate(el => el.classList.contains('expanded')), true);
    await page.waitForFunction(() => getComputedStyle(document.querySelector('.e3-helper-sidebar')).transform === 'matrix(1, 0, 0, 1, 0, 0)');
    const session = await page.context().newCDPSession(page);
    const dragTouch = async (start, end, cancel = false) => {
      await session.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x: start.x, y: start.y }] });
      await session.send('Input.dispatchTouchEvent', { type: 'touchMove', touchPoints: [{ x: end.x, y: end.y }] });
      await session.send('Input.dispatchTouchEvent', { type: cancel ? 'touchCancel' : 'touchEnd', touchPoints: [] });
    };
    const handle = await page.locator('.e3-helper-resize-handle').boundingBox();
    const oldWidth = await page.locator('.e3-helper-sidebar').evaluate(el => el.offsetWidth);
    await dragTouch({ x: handle.x + 3, y: 300 }, { x: handle.x + 103, y: 300 });
    assert.ok(Math.abs(await page.locator('.e3-helper-sidebar').evaluate(el => el.offsetWidth) - (oldWidth - 100)) <= 1, 'Touch resizing follows the drag within pixel rounding');
    await page.locator('#e3-helper-close-btn').tap();
    await page.waitForFunction(() => !document.querySelector('.e3-helper-sidebar').getAnimations().length);
    const toggle = await page.locator('.e3-helper-sidebar-toggle').boundingBox();
    await dragTouch({ x: toggle.x + 10, y: toggle.y + 10 }, { x: toggle.x + 10, y: toggle.y + 110 });
    assert.equal(await page.locator('.e3-helper-sidebar').evaluate(el => el.classList.contains('expanded')), false, 'Dragging must not open the sidebar');
    assert.ok(await page.evaluate(() => localStorage.getItem('e3-helper-toggle-top')), 'Touch drag saves the entrance position');
    const movedToggle = await page.locator('.e3-helper-sidebar-toggle').boundingBox();
    await dragTouch({ x: movedToggle.x + 10, y: movedToggle.y + 10 }, { x: movedToggle.x + 10, y: movedToggle.y + 30 }, true);
    await page.locator('.e3-helper-sidebar-toggle').tap();
    assert.equal(await page.locator('.e3-helper-sidebar').evaluate(el => el.classList.contains('expanded')), true, 'Cancelled gestures must not break the next tap');
    await page.waitForFunction(() => getComputedStyle(document.querySelector('.e3-helper-sidebar')).transform === 'matrix(1, 0, 0, 1, 0, 0)');
    await page.setViewportSize({ width: 320, height: 900 });
    const sidebarBounds = await page.locator('.e3-helper-sidebar').boundingBox();
    assert.ok(sidebarBounds.width <= 320 && sidebarBounds.x >= 0, 'Sidebar fits a narrow split view');
    assert.deepEqual(errors, []);
    console.log('Passed: six tabs, language/settings, touch tap/resize/drag/cancel and narrow split view.');
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
