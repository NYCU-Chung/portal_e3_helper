const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../background.js'), 'utf8');
const start = source.indexOf('async function syncAnnouncementsAndMessagesSilently()');
const manual = source.indexOf('async function loadAnnouncementsAndMessagesInBackground()', start);
const end = source.indexOf('\n}', manual) + 2;
function load(tabs = []) {
  const created = [], messages = [], removed = [];
  let timeout, badges = 0;
  const context = vm.createContext({
    console: { log() {}, error() {} }, uiText: text => text,
    setTimeout(fn, delay) { if (delay === 90000) timeout = fn; return 1; }, clearTimeout() {},
    updateBadgeFromStorage() { badges++; },
    chrome: {
      tabs: {
        query: async () => tabs,
        create: async options => { created.push(options); return { id: 42 }; },
        remove: async id => { removed.push(id); },
        sendMessage(id, request, callback) { messages.push({ id, request }); callback({ success: true }); },
        onUpdated: { addListener() {}, removeListener() {} }
      },
      runtime: {}, storage: { local: { set: async () => {}, remove: async () => {} } }
    }
  });
  vm.runInContext(source.slice(start, end), context);
  return { context, created, messages, removed, expire: () => timeout(), badges: () => badges };
}
test('scheduled announcement sync with no E3 tab does not create a tab', async () => {
  const app = load();
  const pending = app.context.syncAnnouncementsAndMessagesSilently();
  await new Promise(setImmediate);
  assert.equal(app.created.length, 0);
  const result = await pending;
  assert.equal(result.success, false);
  assert.equal(result.reason, 'no_e3_tab');
  assert.equal(app.messages.length, 0);
});
test('scheduled announcement sync reuses an existing E3 tab and updates badge', async () => {
  const app = load([{ id: 7 }]);
  const result = await app.context.syncAnnouncementsAndMessagesSilently();
  assert.equal(result.success, true);
  assert.equal(app.created.length, 0);
  assert.equal(app.messages[0].id, 7);
  assert.equal(app.messages[0].request.action, 'loadAnnouncementsAndMessagesInTab');
  assert.equal(app.badges(), 1);
});
test('manual announcement loading can still create a background tab and cleans up on timeout', async () => {
  const app = load();
  const pending = app.context.loadAnnouncementsAndMessagesInBackground();
  const rejected = assert.rejects(pending, /載入超時/);
  await new Promise(setImmediate);
  assert.equal(app.created.length, 1);
  assert.equal(app.created[0].url, 'https://e3p.nycu.edu.tw/');
  assert.equal(app.created[0].active, false);
  app.expire();
  await rejected;
  assert.deepEqual(app.removed, [42]);
});
function loadParticipantLock(seed = {}) {
  const lockStart = source.indexOf('let participantCheckQueue');
  const lockEnd = source.indexOf('\n}', lockStart) + 2;
  const data = structuredClone(seed);
  let now = 1_000_000;
  const context = vm.createContext({ Date: { now: () => now },
    chrome: { storage: { local: { get: async () => structuredClone(data), set: async value => Object.assign(data, structuredClone(value)) } } } });
  vm.runInContext(`${source.slice(lockStart, lockEnd)}\nglobalThis.update = updateParticipantCheck;`, context);
  return { update: context.update, data, advance: ms => { now += ms; } };
}
test('concurrent membership checks from several tabs are granted once', async () => {
  const lock = loadParticipantLock();
  const results = await Promise.all([1, 2, 3, 4, 5].map(() => lock.update('claim')));
  assert.equal(typeof results[0], 'string');
  assert.deepEqual(results.slice(1), [false, false, false, false]);
  assert.equal(await lock.update('force'), false, 'a manual check never overlaps a running one');
  await lock.update('release', results[0]);
  assert.equal(await lock.update('claim'), false, 'scheduled checks wait five minutes after the last start');
  const manual = await lock.update('force');
  assert.equal(typeof manual, 'string', 'a manual check may run right after the previous one finished');
  await lock.update('release', manual);
  lock.advance(5 * 60000 + 1);
  assert.equal(typeof await lock.update('claim'), 'string');
});
test('an unreleased membership check stops blocking after ten minutes', async () => {
  const lock = loadParticipantLock();
  assert.equal(typeof await lock.update('claim'), 'string');
  lock.advance(9 * 60000);
  assert.equal(await lock.update('force'), false);
  lock.advance(60000 + 1);
  assert.equal(typeof await lock.update('claim'), 'string');
});
test('expired membership owners cannot release or publish over a replacement check', async () => {
  const lock = loadParticipantLock();
  const first = await lock.update('claim');
  lock.advance(10 * 60000 + 1);
  const second = await lock.update('claim');
  assert.notEqual(first, second);
  assert.equal(await lock.update('release', first), false);
  assert.equal(await lock.update('commit', first, { course: { count: 1 } }), false);
  assert.equal(lock.data.participantCounts, undefined);
  assert.equal(await lock.update('force'), false);
  assert.equal(await lock.update('commit', second, { course: { count: 2 } }), true);
  assert.equal(lock.data.participantCounts.course.count, 2);
  assert.equal(await lock.update('release', second), true);
});
test('a delayed tab discards its counts and alerts after another tab acquires the lease', async () => {
  const lock = loadParticipantLock({ courses: [{ id: 'course', fullname: 'Course' }], participantCounts: { course: { count: 1 } } });
  const content = fs.readFileSync(path.join(__dirname, '../content.js'), 'utf8');
  const extract = name => {
    const start = content.search(new RegExp(`(?:async )?function ${name}\\(`));
    return content.slice(start, content.indexOf('\n}', start) + 2);
  };
  const notifications = [];
  const context = vm.createContext({ console: { log() {}, error() {} }, setTimeout: resolve => resolve(),
    chrome: { runtime: { sendMessage: (request, callback) => {
      const action = request.action === 'releaseParticipantCheck' ? 'release' : request.action === 'commitParticipantCheck' ? 'commit' : 'claim';
      lock.update(action, request.token, request.counts).then(result => callback({ granted: Boolean(result), token: typeof result === 'string' ? result : undefined }));
    } }, storage: { local: { get: async () => structuredClone(lock.data), set: async value => Object.assign(lock.data, value) } } },
    fetchCourseParticipants: async () => {
      lock.advance(10 * 60000 + 1);
      const owner = await lock.update('claim');
      await lock.update('commit', owner, { course: { count: 3 } });
      return { count: 2 };
    }, saveParticipantChangeNotifications: async changes => notifications.push(...changes), updateNotificationBadge() {}, updateLastCheckTimeDisplay() {} });
  vm.runInContext(extract('sendParticipantCheckMessage') + extract('checkAllCoursesParticipants'), context);
  await context.checkAllCoursesParticipants();
  assert.equal(lock.data.participantCounts.course.count, 3);
  assert.deepEqual(notifications, []);
  assert.equal(await lock.update('force'), false);
});
