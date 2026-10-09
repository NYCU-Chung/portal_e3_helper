// Exercise the packaged extension with real Chrome storage/alarms/notification APIs.
const { chromium } = require(process.env.PLAYWRIGHT_PATH || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
(async () => {
  const extension = process.env.EXTENSION_PATH || path.resolve(__dirname, '..');
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'e3-notifications-integration-'));
  let context;
  const launch = () => chromium.launchPersistentContext(profile, {
    headless: true, channel: 'chromium',
    // Prevent installation/startup sync from reaching E3 or any external service.
    args: [`--disable-extensions-except=${extension}`, `--load-extension=${extension}`, '--host-resolver-rules=MAP * ~NOTFOUND']
  });
  try {
    context = await launch();
    let worker = context.serviceWorkers()[0] || await context.waitForEvent('serviceworker');
    const id = new URL(worker.url()).host;
    const page = await context.newPage();
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await worker.evaluate(async () => {
      await E3HelperI18n.ready;
      await chrome.storage.local.set({ interfaceLanguage: 'en', notificationSettings: { desktop: true, updates: true, mode: 'instant', time: '20:00', reminders: [] } });
      await E3Notifications.tick();
    });
    await page.goto(`chrome-extension://${id}/notification-settings.html`);
    await page.locator('#desktop').waitFor();
    assert.equal(await page.title(), 'Notification settings · E3 Helper');
    assert.equal(await page.locator('#desktop').isChecked(), true);
    await page.locator('#test').click();
    await page.waitForFunction(() => document.getElementById('status').dataset.state === 'success');
    assert.ok(await worker.evaluate(async () => (await chrome.notifications.getAll())['e3-notification-test']));
    const realAPIs = await worker.evaluate(async () => {
      const alarms = await chrome.alarms.getAll();
      await E3Notifications.enqueue('integration-immediate', { type: 'basic', iconUrl: chrome.runtime.getURL('128.png'), title: 'E3 integration test', message: 'Immediate delivery' }, 'https://e3p.nycu.edu.tw/');
      await E3Notifications.enqueue('integration-immediate', { type: 'basic', title: 'Duplicate' });
      return { alarm: alarms.find(alarm => alarm.name === 'deliverE3Notifications'), storage: await chrome.storage.local.get(['notificationQueue', 'notificationHistory']), active: await chrome.notifications.getAll() };
    });
    assert.equal(realAPIs.alarm.periodInMinutes, 1);
    assert.ok(realAPIs.active['e3-alert-integration-immediate']);
    assert.equal(realAPIs.storage.notificationHistory.filter(id => id === 'integration-immediate').length, 1);
    assert.equal(realAPIs.storage.notificationQueue.length, 0);

    // Real storage events: initial snapshot is silent; later additions notify.
    const updates = await worker.evaluate(async () => {
      await chrome.storage.local.set({ announcements: [{ id: 'baseline', title: 'Original course text' }] });
      await E3Notifications.tick();
      const initial = await chrome.notifications.getAll();
      await chrome.storage.local.set({ announcements: [{ id: 'baseline', title: 'Original course text' }, { id: 'new', title: '課程原文', url: 'https://e3p.nycu.edu.tw/' }] });
      await E3Notifications.tick();
      return { initial, active: await chrome.notifications.getAll() };
    });
    assert.equal(updates.initial['e3-alert-announcements-baseline'], undefined);
    assert.ok(updates.active['e3-alert-announcements-new']);

    // Save via the actual options page and observe the worker using the setting.
    await page.locator('#desktop').uncheck();
    await page.locator('#save').click();
    await page.waitForFunction(() => document.getElementById('status').dataset.state === 'success');
    const disabled = await worker.evaluate(async () => {
      await E3Notifications.enqueue('integration-off', { type: 'basic', title: 'Disabled', message: 'Must not deliver' });
      return { settings: (await chrome.storage.local.get('notificationSettings')).notificationSettings, active: await chrome.notifications.getAll() };
    });
    assert.equal(disabled.settings.desktop, false);
    assert.equal(disabled.active['e3-alert-integration-off'], undefined);
    // Pause delivery after link persistence, then save disable through the real
    // options page while its storage-change handler waits behind this tick.
    await page.locator('#desktop').check();
    await page.locator('#save').click();
    await page.waitForFunction(() => document.getElementById('status').dataset.state === 'success');
    await worker.evaluate(async () => {
      await E3Notifications.tick();
      const now = Date.now();
      await chrome.storage.local.set({ notificationQueue: [{
        id: 'integration-disable-race', due: now - 1, created: now - 2, delivered: {}, attempts: 0,
        options: { type: 'basic', iconUrl: chrome.runtime.getURL('128.png'), title: 'Disable race', message: 'Must not deliver' }
      }] });
      const set = E3NotificationAPI.set;
      let entered;
      const paused = new Promise(resolve => { entered = resolve; });
      const blocked = new Promise(resolve => { globalThis.releaseDelivery = resolve; });
      E3NotificationAPI.set = async value => {
        await set(value);
        if (value.notificationLinks && !value.notificationQueue) { entered(); await blocked; }
      };
      globalThis.pendingDelivery = E3Notifications.tick().finally(() => { E3NotificationAPI.set = set; });
      await paused;
    });
    await page.locator('#desktop').uncheck();
    await page.locator('#save').click();
    await page.waitForFunction(() => document.getElementById('status').dataset.state === 'success');
    const raced = await worker.evaluate(async () => {
      releaseDelivery();
      await pendingDelivery;
      await E3Notifications.tick();
      delete globalThis.releaseDelivery;
      delete globalThis.pendingDelivery;
      return { active: await chrome.notifications.getAll(), storage: await chrome.storage.local.get('notificationQueue') };
    });
    assert.equal(raced.active['e3-alert-integration-disable-race'], undefined);
    assert.equal(raced.storage.notificationQueue.length, 0);
    await page.locator('#desktop').check();
    await page.locator('input[value=daily]').check();
    await page.locator('#time').fill('23:59');
    await page.locator('#save').click();
    await page.waitForFunction(() => document.getElementById('status').dataset.state === 'success');
    const delayed = await worker.evaluate(async () => {
      await E3Notifications.enqueue('integration-delayed', { type: 'basic', iconUrl: chrome.runtime.getURL('128.png'), title: 'Delayed', message: 'Persistent daily delivery' });
      return { storage: await chrome.storage.local.get('notificationQueue'), active: await chrome.notifications.getAll() };
    });
    assert.ok(delayed.storage.notificationQueue.some(item => item.id === 'integration-delayed' && item.due > Date.now()));
    assert.equal(delayed.active['e3-alert-integration-delayed'], undefined);
    assert.deepEqual(errors, []);

    // Restart the entire browser profile; both queued work and history survive.
    await context.close();
    context = await launch();
    worker = context.serviceWorkers()[0] || await context.waitForEvent('serviceworker');
    const persisted = await worker.evaluate(async () => {
      await E3HelperI18n.ready;
      await E3Notifications.tick();
      return chrome.storage.local.get(['notificationQueue', 'notificationHistory']);
    });
    assert.ok(persisted.notificationHistory.includes('integration-immediate'));
    assert.ok(persisted.notificationQueue.some(item => item.id === 'integration-delayed'));
    const delivered = await worker.evaluate(async () => {
      const { notificationQueue } = await chrome.storage.local.get('notificationQueue');
      notificationQueue.forEach(item => { item.due = Date.now() - 1; });
      await chrome.storage.local.set({ notificationQueue });
      await E3Notifications.tick();
      return { storage: await chrome.storage.local.get(['notificationQueue', 'notificationHistory']), active: await chrome.notifications.getAll() };
    });
    assert.ok(delivered.active['e3-alert-integration-delayed']);
    assert.equal(delivered.storage.notificationQueue.length, 0);

    // Invalid real notification options must produce an error and a retry queue.
    const failure = await worker.evaluate(async () => {
      await chrome.storage.local.set({ notificationSettings: { desktop: true, updates: true, mode: 'instant', reminders: [] } });
      await E3Notifications.enqueue('integration-retry', { type: 'basic', iconUrl: chrome.runtime.getURL('128.png'), title: 'Retry' });
      return chrome.storage.local.get(['notificationQueue', 'notificationHistory', 'notificationDeliveryError']);
    });
    assert.ok(failure.notificationDeliveryError);
    assert.ok(failure.notificationQueue.some(item => item.id === 'integration-retry'));
    assert.ok(!failure.notificationHistory.includes('integration-retry'));
    const retried = await worker.evaluate(async () => {
      const { notificationQueue } = await chrome.storage.local.get('notificationQueue');
      notificationQueue.forEach(item => { item.options.message = 'Recovered delivery'; item.due = Date.now() - 1; });
      await chrome.storage.local.set({ notificationQueue });
      await E3Notifications.tick();
      return { storage: await chrome.storage.local.get(['notificationQueue', 'notificationDeliveryError']), active: await chrome.notifications.getAll() };
    });
    assert.ok(retried.active['e3-alert-integration-retry']);
    assert.equal(retried.storage.notificationQueue.length, 0);
    assert.equal(retried.storage.notificationDeliveryError, '');
    const reminders = await worker.evaluate(async () => {
      const deadline = Date.now() + 1800000;
      await chrome.storage.local.set({
        notificationSettings: { desktop: true, updates: true, mode: 'instant', reminders: [24, 1] },
        assignments: [
          { eventId: 'live', name: 'Active deadline', deadline },
          { eventId: 'submitted', name: 'Already submitted', deadline, manualStatus: 'submitted' },
          { eventId: 'expired', name: 'Expired', deadline: Date.now() - 1 }
        ]
      });
      await E3Notifications.tick();
      return Object.keys(await chrome.notifications.getAll());
    });
    assert.equal(reminders.filter(id => id.startsWith('e3-alert-deadline-live-')).length, 1);
    assert.ok(reminders.some(id => /^e3-alert-deadline-live-\d+-1$/.test(id)));
    assert.ok(!reminders.some(id => /deadline-(submitted|expired)-/.test(id)));

    // Fire the real Chrome alarm to verify background.js dispatches to the engine.
    await worker.evaluate(async () => {
      await chrome.storage.local.set({ notificationSettings: { desktop: true, updates: true, mode: 'daily', time: '23:59', reminders: [] } });
      await E3Notifications.enqueue('integration-alarm', { type: 'basic', iconUrl: chrome.runtime.getURL('128.png'), title: 'Alarm', message: 'Real alarm delivery' });
      const { notificationQueue } = await chrome.storage.local.get('notificationQueue');
      notificationQueue.forEach(item => { item.due = Date.now() - 1; });
      await chrome.storage.local.set({ notificationQueue });
      await chrome.alarms.create('deliverE3Notifications', { when: Date.now() + 1000 });
    });
    const alarmPage = await context.newPage();
    await alarmPage.goto(`chrome-extension://${id}/notification-settings.html`);
    await alarmPage.waitForFunction(async () => Boolean((await chrome.notifications.getAll())['e3-alert-integration-alarm']), null, { timeout: 15000 });
    console.log('Passed: packaged MV3 worker, real notification APIs, options save/test, storage update delivery, alarm registration, saved-disable delivery race, daily scheduling, browser restart persistence, failed-delivery retry, missed-threshold collapse, deadline filtering and actual alarm dispatch.');
  } finally {
    if (context) {
      for (const worker of context.serviceWorkers()) {
        await worker.evaluate(async () => {
          for (const id of Object.keys(await chrome.notifications.getAll())) await chrome.notifications.clear(id);
        }).catch(() => {});
      }
      await context.close();
    }
    fs.rmSync(profile, { recursive: true, force: true });
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
