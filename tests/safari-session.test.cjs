const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.join(__dirname, '..');
const background = fs.readFileSync(path.join(root, 'background.js'), 'utf8');
const content = fs.readFileSync(path.join(root, 'content.js'), 'utf8');
const pageSource = content.slice(content.indexOf('async function fetchE3Session('), content.indexOf('// 監聽來自 background script 的訊息'));
const workerSource = background.slice(background.indexOf('async function getE3SyncOrigin('), background.indexOf('async function syncAssignments('));
const loginSource = background.slice(background.indexOf('async function checkLoginStatus('), background.indexOf('// 取得 sesskey'));
function load({ safari = true, tabs = [{ id: 1, active: true, url: 'https://e3p.nycu.edu.tw/' }], origin = 'https://e3p.nycu.edu.tw', pageHTML = '<a href="/login/logout.php?sesskey=fixture">Logout</a>', pageError = false, callbackOnly = false, sessions = null } = {}) {
  const pageRequests = [], directRequests = [];
  const page = vm.createContext({ URL, AbortController, setTimeout, clearTimeout,
    chrome: { runtime: { id: 'e3-extension' } }, location: { origin },
    fetch: async (url, options) => {
      pageRequests.push({ url, options });
      if (pageError) throw new Error('offline');
      return { status: 200, statusText: 'OK', url, text: async () => pageHTML };
    }
  });
  vm.runInContext(pageSource, page);
  const worker = vm.createContext({ URL, Response, AbortController, setTimeout, clearTimeout,
    console: { warn() {} },
    chrome: { runtime: { getURL: () => safari ? 'safari-web-extension://fixture/' : 'chrome-extension://fixture/' },
      tabs: { query: async ({ url }) => tabs.filter(tab => (Array.isArray(url) ? url : [url]).some(pattern => tab.url.startsWith(pattern.replace('/*', '/')))), sendMessage: (id, request, callback) => {
        if (sessions) {
          const session = sessions[id];
          callback(session === null ? { success: false, error: 'offline' } : { success: true, status: 200, statusText: 'OK', url: request.url, text: session ? '{"sesskey":"fixture"}' : '<form id="loginform"></form>' });
          return;
        }
        const response = page.fetchE3Session(request, { id: 'e3-extension' })
          .catch(error => ({ success: false, error: error.message }));
        if (callback) response.then(callback);
        return callbackOnly ? undefined : response;
      } }
    },
    fetch: async (url, options) => {
      directRequests.push({ url, options });
      return { ok: true, url: 'https://e3p.nycu.edu.tw/login/index.php', text: async () => '<form id="loginform"></form>' };
    }
  });
  const sesskeyStart = background.indexOf('async function getSesskey(');
  const sesskeySource = background.slice(sesskeyStart, background.indexOf('\n}\n', sesskeyStart) + 2);
  vm.runInContext(workerSource + loginSource + sesskeySource, worker);
  return { worker, page, pageRequests, directRequests };
}
test('Safari login check uses the signed-in page session despite unauthenticated background fetch', async () => {
  const app = load();
  assert.equal(await app.worker.checkLoginStatus(), true);
  assert.equal(app.pageRequests.length, 2);
  assert.equal(app.pageRequests[0].options.credentials, 'same-origin');
  assert.equal(app.directRequests.length, 0);
});
test('Chrome keeps its existing background fetch path', async () => {
  const app = load({ safari: false });
  assert.equal(await app.worker.checkLoginStatus(), false);
  assert.equal(app.directRequests.length, 1);
  assert.equal(app.pageRequests.length, 0);
});
test('genuinely expired page sessions still report signed out', async () => {
  const app = load({ pageHTML: '<form id="loginform"></form>' });
  assert.equal(await app.worker.checkLoginStatus(), false);
});
test('no tab or network error is unknown, rather than a false expired-session result', async () => {
  for (const seed of [{ tabs: [] }, { pageError: true }]) {
    const app = load(seed);
    assert.equal(await app.worker.checkLoginStatus(), null);
  }
});
test('page relay rejects external origins, other extensions, and mutation requests', async () => {
  const app = load();
  for (const url of ['https://api.openai.com/', 'https://e3.nycu.edu.tw/', 'https://e3p.nycu.edu.tw.evil.test/', 'http://e3p.nycu.edu.tw/']) {
    await assert.rejects(app.page.fetchE3Session({ url }, { id: 'e3-extension' }));
  }
  await assert.rejects(app.page.fetchE3Session({ url: 'https://e3p.nycu.edu.tw/' }, { id: 'other-extension' }));
  await assert.rejects(app.page.fetchE3Session({ url: 'https://e3p.nycu.edu.tw/lib/ajax/service.php', method: 'POST', body: JSON.stringify([{ methodname: 'core_user_update_users' }]) }, { id: 'e3-extension' }));
  assert.equal(app.pageRequests.length, 0);
});
test('read-only calendar API requests keep the body and response JSON', async () => {
  const app = load({ pageHTML: JSON.stringify([{ data: { events: [] } }]) });
  const body = JSON.stringify([{ methodname: 'core_calendar_get_action_events_by_timesort', args: {} }]);
  const response = await app.worker.fetchWithTimeout('https://e3p.nycu.edu.tw/lib/ajax/service.php?sesskey=fixture', { method: 'POST', body }, 15000);
  assert.equal(response.ok, true);
  assert.deepEqual(await response.json(), [{ data: { events: [] } }]);
  assert.equal(app.pageRequests[0].options.body, body);
});

test('login-check failure retains the underlying relay error for diagnostics', async () => {
  const app = load({ pageError: true });
  let reason;
  assert.equal(await app.worker.checkLoginStatus(error => { reason = error.message; }), null);
  assert.equal(reason, 'offline');
});

test('Safari waits for callback responses when tabs.sendMessage returns no promise', async () => {
  const app = load({ callbackOnly: true });
  assert.equal(await app.worker.checkLoginStatus(), true);
  assert.equal(app.pageRequests.length, 2);
  assert.equal(app.directRequests.length, 0);
});

test('Safari uses the available E3 origin for login, sesskey and both sync APIs', async () => {
  const origin = 'https://e3.nycu.edu.tw';
  const app = load({ origin, tabs: [{ id: 1, active: true, url: origin + '/my/' }], pageHTML: '{"sesskey":"fixture"}' });
  assert.equal(await app.worker.checkLoginStatus(), true);
  assert.equal(await app.worker.getSesskey(), 'fixture');
  assert.ok(app.pageRequests.every(request => new URL(request.url).origin === origin));
  // Run the real sync functions through their first AJAX request.
  for (const name of ['syncAssignments', 'syncCourses']) {
    const start = background.indexOf(`async function ${name}(`);
    const code = background.slice(start, background.indexOf('\n}\n', start) + 2);
    vm.runInContext(code, app.worker);
    app.worker.console = { log() {}, error() {}, warn() {} };
    app.worker.chrome.storage = { local: { get: async () => ({}) } };
    app.worker.fetchWithTimeout = async url => {
      app.pageRequests.push({ url });
      return { ok: true, url, text: async () => '{"sesskey":"fixture"}', json: async () => { throw new Error('stop after AJAX'); } };
    };
    await assert.rejects(app.worker[name]());
    assert.ok(app.pageRequests.at(-1).url.startsWith(origin + '/lib/ajax/service.php?'));
  }
});

test('Safari chooses a signed-in host over an active signed-out or unreachable host', async () => {
  for (const unavailable of [false, null]) {
    const app = load({ callbackOnly: true, tabs: [
      { id: 1, active: true, url: 'https://e3.nycu.edu.tw/my/' },
      { id: 2, active: false, url: 'https://e3p.nycu.edu.tw/my/' }
    ], sessions: { 1: unavailable, 2: true } });
    assert.equal(await app.worker.getE3SyncOrigin(), 'https://e3p.nycu.edu.tw');
    assert.equal(await app.worker.checkLoginStatus(), true);
  }
});

test('a whole Safari sync keeps the authenticated host after selection', async () => {
  const app = load({ tabs: [
    { id: 1, active: true, url: 'https://e3.nycu.edu.tw/my/' },
    { id: 2, active: false, url: 'https://e3p.nycu.edu.tw/my/' }
  ], sessions: { 1: false, 2: true } });
  const used = [];
  app.worker.console = { log() {}, error() {} };
  app.worker.chrome.storage = { local: { set: async () => {} } };
  app.worker.scanTimetableExams = async () => {};
  app.worker.syncAssignments = async origin => { used.push(origin); return []; };
  app.worker.syncCourses = async origin => { used.push(origin); return []; };
  const start = background.indexOf('async function syncE3Data(');
  vm.runInContext(background.slice(start, background.indexOf('\n}\n', start) + 2), app.worker);
  assert.equal((await app.worker.syncE3Data()).success, true);
  assert.deepEqual(used, ['https://e3p.nycu.edu.tw', 'https://e3p.nycu.edu.tw']);
});
