const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const root = path.join(__dirname, '..');
const background = fs.readFileSync(path.join(root, 'background.js'), 'utf8');
const content = fs.readFileSync(path.join(root, 'content.js'), 'utf8');

function sliceBetween(source, start, end) {
  const a = source.indexOf(start);
  const b = source.indexOf(end, a + start.length);
  assert.notEqual(a, -1, `missing start marker: ${start}`);
  assert.notEqual(b, -1, `missing end marker: ${end}`);
  return source.slice(a, b);
}

test('E3 URL validation requires exact HTTPS hosts', () => {
  const code = sliceBetween(background, 'const E3_ALLOWED_HOSTS', '// NYCU E3 Helper - Background Script');
  const context = { URL, Set };
  vm.createContext(context);
  vm.runInContext(`${code}\nthis.parseAllowedE3Url = parseAllowedE3Url;`, context);
  assert.equal(context.parseAllowedE3Url('https://e3.nycu.edu.tw/x').hostname, 'e3.nycu.edu.tw');
  assert.equal(context.parseAllowedE3Url('https://e3p.nycu.edu.tw/x').hostname, 'e3p.nycu.edu.tw');
  assert.throws(() => context.parseAllowedE3Url('http://e3.nycu.edu.tw/x'));
  assert.throws(() => context.parseAllowedE3Url('https://attacker-e3.nycu.edu.tw/x'));
  assert.throws(() => context.parseAllowedE3Url('https://e3.nycu.edu.tw.evil.invalid/x'));
  assert.throws(() => context.parseAllowedE3Url('https://e3.nycu.edu.tw@evil.invalid/x'));
});

test('rich-content URL policy blocks executable and unsafe data schemes', () => {
  const code = sliceBetween(content, 'function getRichContentUrlScheme', '// HTML 白名單清理');
  const context = {};
  vm.createContext(context);
  vm.runInContext(`${code}\nthis.safe = isSafeRichContentUrl;`, context);
  assert.equal(context.safe('https://e3.nycu.edu.tw/mod/forum/view.php', 'href'), true);
  assert.equal(context.safe('/mod/forum/view.php', 'href'), true);
  assert.equal(context.safe('mailto:user@example.com', 'href'), true);
  assert.equal(context.safe('java\nscript:alert(1)', 'href'), false);
  assert.equal(context.safe(' data:text/html,<script>alert(1)</script>', 'href'), false);
  assert.equal(context.safe('data:image/png;base64,AAAA', 'src'), true);
  assert.equal(context.safe('data:image/svg+xml,<svg/>', 'src'), false);
});

test('hardening is wired into downloads, redirects, links, and manifest permissions', () => {
  assert.ok(background.includes('parseAllowedE3Url(request.url);'));
  assert.ok(background.includes('parseAllowedE3Url(response.url);'));
  assert.ok(background.includes("/[\\\\/]/.test(request.filename)"));
  assert.ok(content.includes("rel.add('noopener')"));
  assert.ok(content.includes("rel.add('noreferrer')"));
  const manifest = JSON.parse(fs.readFileSync(path.join(root, 'manifest.json'), 'utf8'));
  assert.equal(manifest.minimum_chrome_version, '93');
  assert.equal(manifest.version, '2.3.0');
  assert.equal(manifest.host_permissions.includes('http://localhost:11434/*'), false);
});


function fakeRichElement(tagName, attributes = {}) {
  const values = new Map(Object.entries(attributes));
  return {
    tagName: tagName.toUpperCase(),
    removed: false,
    get attributes() { return [...values].map(([name, value]) => ({ name, value })); },
    remove() { this.removed = true; },
    removeAttribute(name) { values.delete(name); },
    getAttribute(name) { return values.has(name) ? values.get(name) : null; },
    setAttribute(name, value) { values.set(name, String(value)); },
    has(name) { return values.has(name); },
    value(name) { return values.get(name); }
  };
}

test('sanitizer executes form-control and DOM-clobbering protections', () => {
  const code = sliceBetween(content, 'function getRichContentUrlScheme', '// HTML 白名單清理');
  const context = { Set };
  vm.createContext(context);
  vm.runInContext(`${code}\nthis.clean = sanitizeRichContentElement;`, context);

  const button = fakeRichElement('button', { form: 'known-form', formaction: 'javascript:alert(1)' });
  assert.equal(context.clean(button), false);
  assert.equal(button.removed, true);

  const div = fakeRichElement('div', { id: 'e3-helper-add-assignment-form', name: 'helper', form: 'host-form' });
  assert.equal(context.clean(div), true);
  assert.equal(div.has('id'), false);
  assert.equal(div.has('name'), false);
  assert.equal(div.has('form'), false);

  const anchor = fakeRichElement('a', { href: 'java\nscript:alert(1)', target: '_blank', rel: 'ugc' });
  assert.equal(context.clean(anchor), true);
  assert.equal(anchor.has('href'), false);
  assert.match(anchor.value('rel'), /\bugc\b/);
  assert.match(anchor.value('rel'), /\bnoopener\b/);
  assert.match(anchor.value('rel'), /\bnoreferrer\b/);
});

test('E3 content fetch rejects redirects and off-host responses before reading body', async () => {
  const urlCode = sliceBetween(background, 'const E3_ALLOWED_HOSTS', '// NYCU E3 Helper - Background Script');
  const fetchCode = sliceBetween(background, 'async function fetchContentFromE3', '// ==================== 載入公告和信件功能');
  let response;
  let options;
  let bodyReads = 0;
  const context = {
    URL, Set,
    console: { log() {}, error() {} },
    fetchWithTimeout: async (_url, opts) => { options = opts; return response; }
  };
  vm.createContext(context);
  vm.runInContext(`${urlCode}\n${fetchCode}\nthis.fetchContentFromE3 = fetchContentFromE3;`, context);

  response = {
    url: 'https://e3.nycu.edu.tw/mod/forum/view.php', ok: true, status: 200,
    text: async () => { bodyReads++; return '<p>ok</p>'; }
  };
  assert.equal(await context.fetchContentFromE3('https://e3.nycu.edu.tw/mod/forum/view.php'), '<p>ok</p>');
  assert.equal(options.redirect, 'error');
  assert.equal(options.credentials, 'include');
  assert.equal(bodyReads, 1);

  response = {
    url: 'https://evil.invalid/redirected', ok: true, status: 200,
    text: async () => { bodyReads++; return 'evil'; }
  };
  await assert.rejects(context.fetchContentFromE3('https://e3.nycu.edu.tw/mod/forum/view.php'), /無法抓取內容/);
  assert.equal(bodyReads, 1);
});

test('download final URL guard cancels off-domain redirects', () => {
  const urlCode = sliceBetween(background, 'const E3_ALLOWED_HOSTS', '// NYCU E3 Helper - Background Script');
  const guardCode = sliceBetween(background, 'const trackedE3Downloads', '// 監聽來自 content script 的訊息');
  const cancelled = [];
  const removed = [];
  const context = {
    URL, Set,
    console: { warn() {} },
    chrome: {
      runtime: { lastError: null },
      downloads: {
        cancel(id, cb) { cancelled.push(id); cb(); },
        removeFile(id, cb) { removed.push(id); cb(); },
        search(_query, cb) { cb([]); },
        onChanged: { addListener() {} }
      }
    }
  };
  vm.createContext(context);
  vm.runInContext(`${urlCode}\n${guardCode}\nthis.guard = cancelDownloadIfFinalUrlDisallowed;`, context);
  assert.equal(context.guard(7, 'https://e3.nycu.edu.tw/pluginfile.php/1/file.pdf'), false);
  assert.deepEqual(cancelled, []);
  assert.equal(context.guard(8, 'https://evil.invalid/file.pdf'), true);
  assert.deepEqual(cancelled, [8]);
  assert.deepEqual(removed, [8]);
});
