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
      window.fixture = { interfaceLanguage: language, notificationSettings: { desktop: true, updates: false },
        notificationDeliveryError: '桌面通知失敗，請檢查瀏覽器與系統通知設定。',
        notificationLog: [
          { id: 'assignment-1', category: 'assignment', title: 'Old <img src=x onerror="window.injected=1">', message: 'Line one\nLine two', url: 'https://e3p.nycu.edu.tw/mod/assign/view.php?id=1', time: 1790000000000 },
          { id: 'participant-2', category: 'participant', title: 'Newest', message: 'Members', url: 'https://evil.example/', time: 1790000600000 }
        ] };
      window.storageListeners = [];
      window.chrome = {
        i18n: { getUILanguage: () => language }, runtime: { getManifest: () => ({ permissions: native ? ['nativeMessaging'] : [] }), getURL: p => native ? `safari-web-extension://test/${p}` : p, sendNativeMessage: async (app, message) => { window.nativeMessage = message; if (message.action === 'showNotification') window.notification = message; return { success: true, permission: 'granted' }; } },
        storage: { local: { get: async () => window.fixture, set: async value => Object.assign(window.fixture, value) }, onChanged: { addListener: fn => window.storageListeners.push(fn) } },
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
      // Per-event switches: the legacy `updates: false` turns off announcements and messages only.
      assert.deepEqual(await page.locator('#event-list input').evaluateAll(inputs => inputs.map(input => [input.dataset.event, input.checked])),
        [['assignment', true], ['grading', true], ['announcement', false], ['message', false], ['participant', true]]);
      await page.waitForFunction(() => !document.getElementById('permission').hidden);
      assert.equal(await page.locator('#permission').getAttribute('data-state'), 'granted');
      // Delivery errors stay visible after edits and clear when the engine clears them.
      assert.equal(await page.locator('#delivery-error').isVisible(), true);
      // History: newest first, course text never parsed as HTML, links only to E3.
      assert.deepEqual(await page.locator('#history-list .history-title').allInnerTexts(), ['Newest', 'Old <img src=x onerror="window.injected=1">']);
      assert.equal(await page.locator('#history-list img').count(), 0);
      assert.equal(await page.evaluate(() => window.injected), undefined);
      assert.deepEqual(await page.locator('#history-list .history-title').evaluateAll(items => items.map(item => item.tagName)), ['STRONG', 'A']);
      // A time chosen in one row cannot be picked again in another.
      assert.equal(await page.locator('#reminder-list select').first().locator('option[value="1"]').isDisabled(), true);
      assert.equal(await page.locator('#reminder-list select').first().locator('option[value="24"]').isDisabled(), false);
      assert.equal(await page.evaluate(() => { const event = new Event('beforeunload', { cancelable: true }); dispatchEvent(event); return event.defaultPrevented; }), false);
      await page.locator('#event-list input[data-event=message]').check();
      await page.locator('#event-list input[data-event=participant]').uncheck();
      assert.equal(await page.evaluate(() => { const event = new Event('beforeunload', { cancelable: true }); dispatchEvent(event); return event.defaultPrevented; }), true);
      assert.equal(await page.locator('#delivery-error').isVisible(), true);
      assert.match(await page.locator('#summary').innerText(), process.env.UI_LANGUAGE === 'zh-TW' ? /不通知：公告、課程成員變動/ : /Not notifying: Announcements, Course membership changes/);
      await page.locator('#save').click();
      await page.waitForFunction(() => window.fixture.notificationSettings?.events);
      assert.deepEqual(await page.evaluate(() => window.fixture.notificationSettings.events), { assignment: true, grading: true, announcement: false, message: true, participant: false });
      assert.equal(await page.evaluate(() => 'updates' in window.fixture.notificationSettings), false);
      assert.equal(await page.evaluate(() => { const event = new Event('beforeunload', { cancelable: true }); dispatchEvent(event); return event.defaultPrevented; }), false);
      await page.evaluate(() => {
        const set = E3NotificationAPI.set;
        E3NotificationAPI.set = async value => {
          window.pendingSettings = value.notificationSettings;
          await new Promise(resolve => { window.finishSettingsSave = resolve; });
          await set(value);
          E3NotificationAPI.set = set;
        };
      });
      await page.locator('#event-list input[data-event=grading]').uncheck();
      await page.locator('#save').click();
      await page.waitForFunction(() => window.finishSettingsSave);
      await page.locator('#event-list input[data-event=grading]').check();
      await page.evaluate(() => window.finishSettingsSave());
      await page.waitForFunction(() => document.getElementById('save').textContent === E3HelperI18n.text('儲存設定'));
      assert.equal(await page.evaluate(() => window.fixture.notificationSettings.events.grading), false);
      assert.equal(await page.locator('#save').isEnabled(), true);
      assert.equal(await page.locator('#status').getAttribute('data-state'), 'pending');
      assert.equal(await page.evaluate(() => { const event = new Event('beforeunload', { cancelable: true }); dispatchEvent(event); return event.defaultPrevented; }), true);
      await page.locator('#save').click();
      await page.waitForFunction(() => window.fixture.notificationSettings.events.grading === true);
      await page.evaluate(() => window.storageListeners.forEach(fn => fn({ notificationDeliveryError: { newValue: '' }, notificationLog: { newValue: [] } }, 'local')));
      assert.equal(await page.locator('#delivery-error').isVisible(), false);
      assert.equal(await page.locator('#history-empty').isVisible(), true);
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
