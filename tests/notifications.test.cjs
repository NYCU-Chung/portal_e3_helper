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
  return { engine: context.E3Notifications, data, desktop, requests, advance: ms => { now += ms; }, failDesktop: value => { failDesktop = value; }, listener, clicked, opened, cleared, context };
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


function queued(id, due) {
  return { id, options, url: 'https://e3p.nycu.edu.tw/source', created: due - 1, due, delivered: {}, attempts: 0 };
}
test('unrelated saves preserve overdue daily alerts and future retry times', async () => {
  const now = new Date(2026, 9, 8, 10).getTime();
  const settings = { desktop: true, updates: true, mode: 'daily', time: '09:00', reminders: [] };
  const app = load({ notificationSettings: settings, notificationQueue: [queued('overdue', now - 3600000), queued('retry', now + 120000)] }, now);
  app.data.notificationSettings = { ...settings, updates: false };
  app.listener({ notificationSettings: { oldValue: settings, newValue: app.data.notificationSettings } }, 'local');
  await app.engine.tick();
  assert.equal(app.desktop.length, 1);
  assert.equal(app.data.notificationQueue[0].due, now + 120000);
});
test('timing changes reschedule future alerts while preserving already-due work', async () => {
  const now = new Date(2026, 9, 8, 10).getTime();
  const settings = { desktop: true, mode: 'daily', time: '09:00', reminders: [] };
  const app = load({ notificationSettings: settings, notificationQueue: [queued('overdue', now - 1), queued('future', now + 86400000)] }, now);
  app.data.notificationSettings = { ...settings, time: '12:00' };
  app.listener({ notificationSettings: { oldValue: settings, newValue: app.data.notificationSettings } }, 'local');
  await app.engine.tick();
  assert.equal(app.desktop.length, 1);
  assert.equal(app.data.notificationQueue[0].due, now + 7200000);
});
test('queue capacity failures reject explicitly and retain the warning until space is available', async () => {
  const now = new Date(2026, 9, 8, 10).getTime();
  const app = load({ notificationSettings: { mode: 'daily', time: '20:00', reminders: [] }, notificationQueue: Array.from({ length: 200 }, (_, i) => queued(`queued-${i}`, now + 3600000)) }, now);
  await assert.rejects(app.engine.enqueue('overflow', options), /通知佇列已滿/);
  assert.ok(app.data.notificationDeliveryError);
  await app.engine.tick();
  assert.ok(app.data.notificationDeliveryError);
  assert.equal(app.data.notificationQueue.length, 200);
  app.advance(3600000);
  await app.engine.tick();
  assert.equal(app.data.notificationDeliveryError, '');
  await app.engine.enqueue('overflow', options);
  assert.ok(app.data.notificationQueue.some(item => item.id === 'overflow'));
});
test('a click on the first visible alert opens E3 while later delivery is pending', async () => {
  const now = Date.now();
  const app = load({ notificationSettings: { reminders: [] }, notificationQueue: [queued('first', now - 1), queued('second', now - 1)] }, now);
  let release, started;
  const blocked = new Promise(resolve => { release = resolve; });
  const entered = new Promise(resolve => { started = resolve; });
  app.context.chrome.notifications.create = async id => { if (id === 'e3-alert-second') { started(); await blocked; } };
  const processing = app.engine.tick();
  await entered;
  try {
    app.clicked('e3-alert-first');
    await new Promise(resolve => setImmediate(resolve));
    assert.deepEqual(app.opened, ['https://e3p.nycu.edu.tw/source']);
    assert.deepEqual(app.cleared, ['e3-alert-first']);
  } finally { release(); await processing; }
});

test('overflowed storage updates survive restart and enqueue when capacity returns', async () => {
  const now = new Date(2026, 9, 8, 10).getTime();
  const app = load({ notificationSettings: { mode: 'daily', time: '20:00', reminders: [] }, notificationQueue: Array.from({ length: 200 }, (_, i) => queued(`queued-${i}`, now + 3600000)) }, now);
  app.data.messages = [{ id: 'new', title: 'Original subject', url: 'https://e3p.nycu.edu.tw/source' }];
  app.listener({ messages: { oldValue: [], newValue: app.data.messages } }, 'local');
  await app.engine.tick();
  assert.deepEqual(app.data.notificationPendingUpdates, [{ source: 'messages', id: 'new' }]);
  assert.ok(app.data.notificationDeliveryError);
  const restarted = load(app.data, now + 3600000);
  await restarted.engine.tick();
  await restarted.engine.tick();
  assert.equal(restarted.data.notificationPendingUpdates.length, 0);
  assert.equal(restarted.data.notificationDeliveryError, '');
  assert.ok(restarted.data.notificationQueue.some(item => item.id === 'messages-new'));
});

test('assignment and grading sidebar entries remain available when desktop queue is full', async () => {
  const background = fs.readFileSync(require('node:path').join(__dirname, '../background.js'), 'utf8');
  for (const name of ['sendAssignmentNotification', 'sendGradingNotification']) {
    const start = background.indexOf(`async function ${name}(`);
    const fn = background.slice(start, background.indexOf('\n}', start) + 2);
    const app = load({ notificationSettings: { reminders: [] } });
    app.context.console = { log() {}, warn() {}, error() {} };
    app.context.uiText = s => s;
    app.context.ui = (parts, ...values) => parts.reduce((result, part, index) => result + part + (values[index] ?? ''), '');
    app.context.E3Notifications.enqueue = async () => { throw new Error('Queue full'); };
    vm.runInContext(fn, app.context);
    await app.context[name]({ eventId: 'new', name: 'Original assignment', course: 'Original course', deadline: Date.now() + 3600000, url: 'https://e3p.nycu.edu.tw/source' });
    assert.equal(app.data.notifications?.length, 1, name);
  }
});

test('waking with 30 minutes remaining sends only the current deadline threshold', async () => {
  const now = new Date(2026, 9, 9, 10).getTime();
  const app = load({ assignments: [{ eventId: 'a', name: 'Task', deadline: now + 1800000 }], notificationSettings: { reminders: [24, 1] } }, now);
  await app.engine.tick();
  assert.deepEqual(app.desktop.map(([id]) => id), [`e3-alert-deadline-a-${now + 1800000}-1`]);
  await app.engine.tick();
  const restarted = load(app.data, now);
  await restarted.engine.tick();
  assert.equal(app.desktop.length, 1);
  assert.equal(restarted.desktop.length, 0);
});

test('a saved disable during link persistence prevents the pending visible alert', async () => {
  const now = Date.now();
  const settings = { desktop: true, reminders: [] };
  const app = load({ notificationSettings: settings, notificationQueue: [queued('first', now - 1)] }, now);
  let release, started;
  const blocked = new Promise(resolve => { release = resolve; });
  const entered = new Promise(resolve => { started = resolve; });
  const set = app.context.chrome.storage.local.set;
  app.context.chrome.storage.local.set = async value => {
    await set(value);
    if (value.notificationLinks && !value.notificationQueue) { started(); await blocked; }
  };
  const processing = app.engine.tick();
  await entered;
  try {
    const disabled = { ...settings, desktop: false };
    await app.context.E3NotificationAPI.set({ notificationSettings: disabled });
    app.listener({ notificationSettings: { oldValue: settings, newValue: disabled } }, 'local');
  } finally { release(); await processing; }
  await app.engine.tick();
  assert.equal(app.desktop.length, 0);
  assert.equal(app.data.notificationQueue.length, 0);
});

test('normal threshold progression still sends the 24-hour and then the 1-hour reminder', async () => {
  const now = Date.now();
  const deadline = now + 25 * 3600000;
  const app = load({ assignments: [{ eventId: 'a', name: 'Task', deadline }], notificationSettings: { reminders: [1, 24, 24] } }, now);
  await app.engine.tick();
  assert.equal(app.desktop.length, 0);
  app.advance(3600000);
  await app.engine.tick();
  assert.deepEqual(app.desktop.map(([id]) => id), [`e3-alert-deadline-a-${deadline}-24`]);
  app.advance(23 * 3600000);
  await app.engine.tick();
  assert.deepEqual(app.desktop.map(([id]) => id), [`e3-alert-deadline-a-${deadline}-24`, `e3-alert-deadline-a-${deadline}-1`]);
});

test('a queued or retrying historical deadline is superseded by the current threshold', async () => {
  const now = Date.now();
  const deadline = now + 1800000;
  for (const attempts of [0, 2]) {
    const old = { ...queued(`deadline-a-${deadline}-24`, now - 1), attempts, deadline: { eventId: 'a', timestamp: deadline, hours: 24 } };
    const current = { ...queued(`deadline-a-${deadline}-1`, now - 1), deadline: { eventId: 'a', timestamp: deadline, hours: 1 } };
    const app = load({ assignments: [{ eventId: 'a', name: 'Task', deadline }], notificationSettings: { mode: 'daily', reminders: [24, 1] }, notificationQueue: [old, current] }, now);
    await app.engine.tick();
    assert.deepEqual(app.desktop.map(([id]) => id), [`e3-alert-deadline-a-${deadline}-1`]);
    assert.equal(app.data.notificationQueue.length, 0);
  }
});

test('a disable saved during the first delivery prevents subsequent batch alerts on callback-only Chrome', async () => {
  const now = Date.now();
  const settings = { desktop: true, reminders: [] };
  const app = load({ notificationSettings: settings, notificationQueue: [queued('first', now - 1), queued('second', now - 1)] }, now, true, true);
  let release, started;
  const blocked = new Promise(resolve => { release = resolve; });
  const entered = new Promise(resolve => { started = resolve; });
  const create = app.context.chrome.notifications.create;
  app.context.chrome.notifications.create = (id, options, callback) => {
    if (id === 'e3-alert-first') {
      started();
      blocked.then(() => create(id, options, callback));
    } else create(id, options, callback);
  };
  const processing = app.engine.tick();
  await entered;
  try {
    const disabled = { ...settings, desktop: false };
    await app.context.E3NotificationAPI.set({ notificationSettings: disabled });
    app.listener({ notificationSettings: { oldValue: settings, newValue: disabled } }, 'local');
  } finally { release(); await processing; }
  await app.engine.tick();
  assert.deepEqual(app.desktop.map(([id]) => id), ['e3-alert-first']);
  assert.equal(app.data.notificationQueue.length, 0);
});

test('delivery revalidates saved update and reminder preferences after link persistence', async () => {
  const now = Date.now();
  const deadline = now + 1800000;
  for (const kind of ['updates', 'reminders']) {
    const settings = { desktop: true, updates: true, reminders: [1] };
    const item = kind === 'updates' ? queued('messages-new', now - 1) : {
      ...queued(`deadline-a-${deadline}-1`, now - 1), deadline: { eventId: 'a', timestamp: deadline, hours: 1 }
    };
    const app = load({ notificationSettings: settings, assignments: [{ eventId: 'a', name: 'Task', deadline }], notificationQueue: [item] }, now);
    const set = app.context.chrome.storage.local.set;
    app.context.chrome.storage.local.set = async value => {
      await set(value);
      if (value.notificationLinks && !value.notificationQueue) {
        await app.context.E3NotificationAPI.set({ notificationSettings: { ...settings, [kind]: kind === 'updates' ? false : [] } });
      }
    };
    await app.engine.tick();
    assert.equal(app.desktop.some(([id]) => id === `e3-alert-${item.id}`), false, kind);
    assert.equal(app.data.notificationQueue.length, 0);
  }
});
