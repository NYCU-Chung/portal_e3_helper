const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../desktop-notifications.js'), 'utf8');
function load(response = { success: true, permission: 'granted' }, permissions = ['nativeMessaging']) {
  const messages = [];
  const context = vm.createContext({ chrome: { runtime: {
    getURL: p => `safari-web-extension://test/${p}`,
    getManifest: () => ({ permissions }),
    sendNativeMessage: async (app, value) => { messages.push({ app, value }); return response; }
  } } });
  vm.runInContext(source, context);
  return { adapter: context.E3DesktopNotifications, messages };
}
test('Safari native transport supports permission and desktop delivery without browser notifications API', async () => {
  const { adapter, messages } = load();
  assert.equal(adapter.supported, true);
  assert.equal(await adapter.authorize(), 'granted');
  await adapter.create('test', { title: 'Reminder', message: 'Deadline' }, 'https://e3p.nycu.edu.tw/');
  assert.equal(messages[0].value.action, 'authorizeNotifications');
  assert.equal(messages[1].value.action, 'showNotification');
  assert.equal(messages[1].value.url, 'https://e3p.nycu.edu.tw/');
});
test('iPad without nativeMessaging disables desktop delivery despite the Safari API existing', async () => {
  const { adapter, messages } = load(undefined, ['storage', 'alarms']);
  assert.equal(adapter.supported, false);
  assert.equal(adapter.native, false);
  assert.equal(await adapter.authorize(), 'denied');
  assert.equal(messages.length, 0);
});
test('native errors are surfaced rather than marking delivery successful', async () => {
  const { adapter } = load({ success: false, error: 'permission denied' });
  await assert.rejects(adapter.create('test', { title: 'test', message: '' }), /permission denied/);
});
