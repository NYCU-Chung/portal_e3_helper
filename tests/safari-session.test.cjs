const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.join(__dirname, '..');
const background = fs.readFileSync(path.join(root, 'background.js'), 'utf8');
const content = fs.readFileSync(path.join(root, 'content.js'), 'utf8');
const pageSource = content.slice(content.indexOf('async function fetchE3Session('), content.indexOf('// 監聽來自 background script 的訊息'));
const workerSource = background.slice(background.indexOf('async function fetchE3FromPage('), background.indexOf('async function syncAssignments('));
const loginSource = background.slice(background.indexOf('async function checkLoginStatus('), background.indexOf('// 取得 sesskey'));
function load({ safari = true, tabs = [{ id: 1, active: true }], pageHTML = '<a href="/login/logout.php?sesskey=fixture">Logout</a>', pageError = false, callbackOnly = false } = {}) {
  const pageRequests = [], directRequests = [];
  const page = vm.createContext({ URL, AbortController, setTimeout, clearTimeout,
    chrome: { runtime: { id: 'e3-extension' } }, location: { origin: 'https://e3p.nycu.edu.tw' },
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
      tabs: { query: async () => tabs, sendMessage: (_, request, callback) => {
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
  vm.runInContext(workerSource + loginSource, worker);
  return { worker, page, pageRequests, directRequests };
}
test('Safari login check uses the signed-in page session despite unauthenticated background fetch', async () => {
  const app = load();
  assert.equal(await app.worker.checkLoginStatus(), true);
  assert.equal(app.pageRequests.length, 1);
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
  assert.equal(app.pageRequests.length, 1);
  assert.equal(app.directRequests.length, 0);
});
