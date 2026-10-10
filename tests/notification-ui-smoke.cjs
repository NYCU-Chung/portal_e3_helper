const { chromium } = require(process.env.PLAYWRIGHT_PATH || 'playwright');
const assert = require('node:assert/strict');
const path = require('node:path');
const fs = require('node:fs');
(async () => {
  const browser = await chromium.launch({ headless: true, ...(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {}) });
  try {
    const page = await browser.newPage({ viewport: { width: 1100, height: 1100 } });
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.route('https://extension.test/**', route => {
      const name = new URL(route.request().url()).pathname.slice(1) || 'notification-settings.html';
      route.fulfill({ contentType: name.endsWith('.js') ? 'text/javascript' : name.endsWith('.css') ? 'text/css' : 'text/html', body: fs.readFileSync(path.join(__dirname, '..', name), 'utf8') });
    });
    await page.addInitScript(({ safari, language, native, callbackOnly }) => {
      window.fixture = { interfaceLanguage: language };
      window.chrome = {
        i18n: { getUILanguage: () => language }, runtime: { getManifest: () => ({ permissions: native ? ['nativeMessaging'] : [] }), getURL: p => native ? `safari-web-extension://test/${p}` : p, sendNativeMessage: async (app, message) => { window.nativeMessage = message; if (message.action === 'showNotification') window.notification = message; return { success: true, permission: 'granted' }; } },
        storage: { local: { get: async () => window.fixture, set: async value => Object.assign(window.fixture, value) }, onChanged: { addListener() {} } },
        notifications: safari ? undefined : { getPermissionLevel: async () => 'granted', create: async (...args) => { window.notification = args; } }
      };
      if (callbackOnly) {
        for (const target of [chrome.storage.local, chrome.notifications]) {
          if (!target) continue;
          for (const [key, method] of Object.entries(target)) {
            if (typeof method !== 'function') continue;
            target[key] = (...args) => {
              const callback = args.pop();
              if (typeof callback !== 'function') throw new Error(`Missing callback: ${key}`);
              method(...args).then(callback);
              return undefined;
            };
          }
        }
      }
    }, { callbackOnly: Boolean(process.env.CHROME_93_FIXTURE), safari: Boolean(process.env.SAFARI_FIXTURE || process.env.SAFARI_NATIVE), native: Boolean(process.env.SAFARI_NATIVE), language: process.env.UI_LANGUAGE || 'en' });
    await page.goto('https://extension.test/');
    await page.locator('.reminder-row').first().waitFor();
    assert.deepEqual(await page.locator('#reminder-list select').evaluateAll(inputs => inputs.map(input => input.value)), ['24', '1']);
    assert.equal(await page.locator('#email').count(), 0);
    if (process.env.SAFARI_FIXTURE) {
      assert.equal(await page.locator('#desktop').isDisabled(), true);
      assert.equal(await page.locator('#desktop').isChecked(), false);
      assert.equal(await page.locator('#test').isDisabled(), true);
      assert.equal(await page.locator('#support-notice').isVisible(), true);
      assert.equal(await page.locator('#preferences').evaluate(element => element.disabled), true);
    } else {
      if (!process.env.UI_LANGUAGE) assert.ok(!/[\u3400-\u9fff]/.test(await page.locator('main').innerText()));
      await page.locator('input[value=daily]').check();
      await page.locator('#time').fill('18:30');
      await page.locator('#reminder-list select').first().selectOption('48');
      await page.locator('#reminder-list select').nth(1).selectOption('3');
      await page.locator('#save').click();
      await page.waitForFunction(() => window.fixture.notificationSettings?.time === '18:30');
      assert.deepEqual(await page.evaluate(() => window.fixture.notificationSettings.reminders), [48, 3]);
      await page.locator('#add-reminder').click();
      assert.equal(await page.locator('.reminder-row').count(), 3);
      await page.locator('.remove-button').last().click();
      assert.equal(await page.locator('.reminder-row').count(), 2);
      await page.locator('#test').click();
      await page.waitForFunction(() => window.notification);
      await page.locator('#desktop').uncheck();
      assert.equal(await page.locator('#preferences').evaluate(element => element.disabled), true);
      await page.locator('#desktop').check();
      await page.locator('#save').click();
      await page.screenshot({ path: `/tmp/e3-notifications-desktop-${process.env.UI_LANGUAGE || 'en'}.png`, fullPage: true });
      await page.setViewportSize({ width: 375, height: 900 });
      assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
      await page.screenshot({ path: `/tmp/e3-notifications-mobile-${process.env.UI_LANGUAGE || 'en'}.png`, fullPage: true });
      await page.locator('.remove-button').first().click();
      await page.locator('.remove-button').first().click();
      await page.locator('#save').click();
      await page.waitForFunction(() => window.fixture.notificationSettings?.reminders.length === 0);
    }
    assert.deepEqual(errors, []);
    console.log('Passed: notification dropdowns, timing, save feedback, supported browser states and responsive layout.');
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
