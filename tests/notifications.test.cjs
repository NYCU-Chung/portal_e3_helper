const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const source = fs.readFileSync(require('node:path').join(__dirname, '../notification-engine.js'), 'utf8');
function load(seed = {}, now = new Date(2026, 9, 1, 10).getTime(), desktopSupported = true, callbackOnly = false) {
  const data = structuredClone(seed), desktop = [], requests = [];
  let listener, clicked, failDesktop = false;
  const opened = [], cleared = [];
  class Clock extends Date { constructor(...args) { super(...(args.length ? args : [now])); } static now() { return now; } }
  const context = vm.createContext({ Date: Clock, URL, AbortSignal, console, E3HelperI18n: { ready: Promise.resolve(), text: s => s, language: 'en' },
    fetch: async (...args) => { requests.push(args); throw new Error('Unexpected network request'); },
    chrome: { runtime: { getURL: p => p },
      storage: { local: { get: async () => structuredClone(data), set: async value => Object.assign(data, structuredClone(value)) }, onChanged: { addListener: fn => { listener = fn; } } },
      notifications: desktopSupported ? { create: async (...args) => { if (failDesktop) throw new Error('blocked'); desktop.push(args); }, clear: async id => { cleared.push(id); }, onClicked: { addListener(fn) { clicked = fn; } } } : undefined } });
  context.chrome.tabs = { create: async options => opened.push(options.url) };
  if (callbackOnly) {
    for (const target of [context.chrome.storage.local, context.chrome.notifications, context.chrome.tabs]) {
      if (!target) continue;
      for (const [key, method] of Object.entries(target)) {
        if (typeof method !== 'function') continue;
        target[key] = (...args) => {
          const callback = args.pop();
          assert.equal(typeof callback, 'function');
          Promise.resolve().then(() => method(...args)).then(callback, error => {
            context.chrome.runtime.lastError = { message: error.message };
            callback();
            delete context.chrome.runtime.lastError;
          });
          return undefined;
        };
      }
    }
  }
  vm.runInContext(fs.readFileSync(require('node:path').join(__dirname, '../notification-api.js'), 'utf8'), context);
  vm.runInContext(fs.readFileSync(require('node:path').join(__dirname, '../desktop-notifications.js'), 'utf8'), context);
  vm.runInContext(source, context);
  return { engine: context.E3Notifications, data, desktop, requests, advance: ms => { now += ms; }, failDesktop: value => { failDesktop = value; }, listener, clicked, opened, cleared };
}
const options = { type: 'basic', title: 'Test', message: 'Content' };
test('immediate desktop delivery is deduplicated across worker restarts', async () => {
  const app = load();
  await app.engine.enqueue('one', options, 'https://e3p.nycu.edu.tw/test');
  await app.engine.enqueue('one', options);
  assert.equal(app.desktop.length, 1);
  const restarted = load(app.data);
  await restarted.engine.enqueue('one', options);
  assert.equal(restarted.desktop.length, 0);
});
test('daily delivery waits until the configured local time', async () => {
  const app = load({ notificationSettings: { mode: 'daily', time: '20:00', reminders: [] } });
  await app.engine.enqueue('daily', options);
  assert.equal(app.desktop.length, 0);
  app.advance(10 * 3600000);
  await app.engine.tick();
  assert.equal(app.desktop.length, 1);
  assert.equal(app.data.notificationQueue.length, 0);
});
test('submitted, removed, changed and expired deadlines are skipped', async () => {
  const now = new Date(2026, 9, 1, 10).getTime();
  for (const kind of ['submitted', 'removed', 'changed', 'expired']) {
    const app = load({ assignments: [{ eventId: 'a', name: 'Task', deadline: now + 3600000 }], notificationSettings: { mode: 'daily', time: '10:30', reminders: [1] } }, now);
    await app.engine.tick();
    assert.equal(app.data.notificationQueue.length, 1);
    if (kind === 'submitted') app.data.assignmentStatuses = { a: 'submitted' };
    if (kind === 'removed') app.data.assignments = [];
    if (kind === 'changed') app.data.assignments[0].deadline += 86400000;
    app.advance(kind === 'expired' ? 7200000 : 1800000);
    await app.engine.tick();
    assert.equal(app.desktop.length, 0, kind);
  }
});
test('desktop failure retries and disabled desktop sends nothing', async () => {
  const app = load({ notificationSettings: { reminders: [] } });
  app.failDesktop(true);
  await app.engine.enqueue('retry', options);
  assert.ok(app.data.notificationDeliveryError);
  app.advance(120000); app.failDesktop(false);
  await app.engine.tick();
  assert.equal(app.desktop.length, 1);
  assert.equal(app.data.notificationQueue.length, 0);
  const disabled = load({ notificationSettings: { desktop: false } });
  await disabled.engine.enqueue('off', options);
  assert.equal(disabled.desktop.length, 0);
});
test('retired channel credentials are removed and never used for network requests', async () => {
  const app = load({ notificationSettings: { desktop: true, email: true, emailTo: 'student@example.com', webhook: 'https://relay.example/send', token: 'secret', reminders: [] } });
  await app.engine.enqueue('clean', options);
  assert.equal(app.requests.length, 0);
  assert.deepEqual(Object.keys(app.data.notificationSettings).sort(), ['desktop', 'mode', 'reminders', 'time', 'updates']);
});
test('announcement baseline is silent, later additions notify even after an empty snapshot', async () => {
  const app = load({ notificationSettings: { reminders: [] } });
  app.listener({ announcements: { newValue: [{ id: 'historic', title: 'Old' }] } }, 'local');
  await app.engine.tick();
  assert.equal(app.desktop.length, 0);
  app.listener({ announcements: { oldValue: [], newValue: [{ id: 'new', title: 'New' }] } }, 'local');
  await app.engine.tick();
  assert.equal(app.desktop.length, 1);
});
test('turning off deadline reminders cancels previously queued reminders', async () => {
  const now = new Date(2026, 9, 1, 10).getTime();
  const app = load({ assignments: [{ eventId: 'a', name: 'Task', deadline: now + 3600000 }], notificationSettings: { mode: 'daily', time: '10:30', reminders: [1] } }, now);
  await app.engine.tick();
  app.data.notificationSettings.reminders = [];
  app.advance(1800000);
  await app.engine.tick();
  assert.equal(app.desktop.length, 0);
  assert.equal(app.data.notificationQueue.length, 0);
});

test('Safari without notifications API never retries unsupported desktop delivery', async () => {
  const app = load({ notificationSettings: { desktop: true, reminders: [] } }, Date.now(), false);
  await app.engine.enqueue('unsupported', options);
  assert.equal(app.desktop.length, 0);
  assert.equal((app.data.notificationQueue || []).length, 0);
});


test('Chrome 93 callback delivery failures retry, persist and open trusted click links', async () => {
  const app = load({ notificationSettings: { reminders: [] } }, Date.now(), true, true);
  app.failDesktop(true);
  await app.engine.enqueue('callback', options, 'https://e3p.nycu.edu.tw/source');
  assert.equal(app.data.notificationQueue.length, 1);
  assert.equal(app.data.notificationHistory.length, 0);
  app.advance(120000);
  app.failDesktop(false);
  await app.engine.tick();
  assert.equal(app.desktop.length, 1);
  assert.equal(app.data.notificationQueue.length, 0);
  assert.ok(app.data.notificationHistory.includes('callback'));
  app.clicked('e3-alert-callback');
  await new Promise(resolve => setImmediate(resolve));
  assert.deepEqual(app.opened, ['https://e3p.nycu.edu.tw/source']);
  assert.deepEqual(app.cleared, ['e3-alert-callback']);
  app.data.notificationLinks['e3-alert-untrusted'] = 'https://example.com/';
  app.clicked('e3-alert-untrusted');
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(app.opened.length, 1);
  const restarted = load(app.data, Date.now(), true, true);
  await restarted.engine.enqueue('callback', options);
  assert.equal(restarted.desktop.length, 0);
});
