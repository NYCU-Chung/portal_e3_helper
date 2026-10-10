const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const root = path.join(__dirname, '..');
const content = fs.readFileSync(path.join(root, 'content.js'), 'utf8');
const helper = content.slice(content.indexOf('async function downloadE3File'), content.indexOf('// 綁定下載按鈕事件'));
function load(permissions, response = { success: true }) {
  const links = [], messages = [];
  const context = vm.createContext({ URL,
    chrome: { runtime: { getManifest: () => ({ permissions }), sendMessage: async message => { messages.push(message); return response; } } },
    document: { createElement: () => ({ click() { links.push({ ...this }); }, remove() {} }), body: { appendChild() {} } }
  });
  vm.runInContext(helper, context);
  return { download: context.downloadE3File, links, messages };
}
test('Safari starts a browser-managed file link without using downloads API', async () => {
  const app = load(['storage']);
  await app.download('https://e3p.nycu.edu.tw/pluginfile.php/1/test.pdf', 'test.pdf');
  assert.equal(app.links.length, 1);
  assert.equal(app.messages.length, 0);
  assert.equal(app.links[0].download, 'test.pdf');
  assert.equal(app.links[0].target, '_blank');
  assert.equal(app.links[0].rel, 'noopener noreferrer');
});
test('downloads reject untrusted hosts and non-HTTP schemes before browser actions', async () => {
  const app = load([]);
  for (const url of ['https://evile3p.nycu.edu.tw/file', 'https://e3p.nycu.edu.tw.evil.test/file', 'javascript:alert(1)']) {
    await assert.rejects(app.download(url, 'test.pdf'));
  }
  assert.equal(app.links.length, 0);
});
test('Chrome retains its downloads API and surfaces failures', async () => {
  const app = load(['downloads']);
  await app.download('https://e3.nycu.edu.tw/file', 'test.pdf');
  assert.equal(app.messages.length, 1);
  assert.equal(app.links.length, 0);
  const failed = load(['downloads'], { success: false, error: 'denied' });
  await assert.rejects(failed.download('https://e3.nycu.edu.tw/file', 'test.pdf'), /denied/);
});
test('Safari package omits unsupported permissions and matches shared sources', () => {
  const resources = path.join(root, 'safari/NYCU E3 Helper/NYCU E3 Helper Extension/Resources');
  const manifest = JSON.parse(fs.readFileSync(path.join(resources, 'manifest.json')));
  assert.ok(!manifest.permissions.includes('downloads'));
  assert.ok(!manifest.permissions.includes('notifications'));
  assert.ok(manifest.permissions.includes('nativeMessaging'));
  assert.ok(!('open_in_tab' in manifest.options_ui));
  for (const name of ['content.js', 'background.js', 'desktop-notifications.js', 'notification-api.js', 'notification-engine.js', 'notification-settings.js']) {
    assert.equal(fs.readFileSync(path.join(resources, name), 'utf8'), fs.readFileSync(path.join(root, name), 'utf8'));
  }
});
test('iPad package keeps common sources but does not advertise native desktop notifications', () => {
  const resources = path.join(root, 'safari/ios/NYCU E3 Helper iOS/NYCU E3 Helper iOS Extension/Resources');
  const manifest = JSON.parse(fs.readFileSync(path.join(resources, 'manifest.json')));
  for (const permission of ['downloads', 'notifications', 'nativeMessaging']) {
    assert.ok(!manifest.permissions.includes(permission));
  }
  assert.equal(manifest.version, JSON.parse(fs.readFileSync(path.join(root, 'manifest.json'))).version);
  for (const name of ['content.js', 'background.js', 'desktop-notifications.js', 'notification-api.js', 'notification-engine.js', 'notification-settings.js']) {
    assert.equal(fs.readFileSync(path.join(resources, name), 'utf8'), fs.readFileSync(path.join(root, name), 'utf8'));
  }
});
