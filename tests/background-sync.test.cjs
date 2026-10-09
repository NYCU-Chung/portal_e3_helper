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
