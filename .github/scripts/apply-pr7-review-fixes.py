from pathlib import Path
import re

background_path = Path('background.js')
content_path = Path('content.js')
tests_path = Path('tests/security-hardening.test.cjs')
background = background_path.read_text()
content = content_path.read_text()
tests = tests_path.read_text()


def replace_once(text, old, new, label):
    if new in text:
        return text
    if text.count(old) != 1:
        raise RuntimeError(f'{label}: expected one match, got {text.count(old)}')
    return text.replace(old, new, 1)


download_guard = r'''const trackedE3Downloads = new Set();

function cancelDownloadIfFinalUrlDisallowed(downloadId, finalUrl) {
  try {
    parseAllowedE3Url(finalUrl);
    return false;
  } catch {
    trackedE3Downloads.delete(downloadId);
    chrome.downloads.cancel(downloadId, () => {
      void chrome.runtime.lastError;
      if (chrome.downloads.removeFile) {
        chrome.downloads.removeFile(downloadId, () => void chrome.runtime.lastError);
      }
    });
    console.warn(`E3 Helper: 已取消重新導向至非 E3 網域的下載 - ${finalUrl}`);
    return true;
  }
}

function validateTrackedDownload(downloadId) {
  chrome.downloads.search({ id: downloadId }, items => {
    if (chrome.runtime.lastError) return;
    const item = items?.[0];
    if (item?.finalUrl) cancelDownloadIfFinalUrlDisallowed(downloadId, item.finalUrl);
  });
}

chrome.downloads.onChanged.addListener(delta => {
  if (!trackedE3Downloads.has(delta.id)) return;
  if (delta.finalUrl?.current && cancelDownloadIfFinalUrlDisallowed(delta.id, delta.finalUrl.current)) return;
  if (delta.state?.current === 'complete' || delta.error) {
    validateTrackedDownload(delta.id);
    trackedE3Downloads.delete(delta.id);
  }
});

'''
marker = '// 監聽來自 content script 的訊息\n'
if 'const trackedE3Downloads = new Set();' not in background:
    if marker not in background:
        raise RuntimeError('download guard insertion marker missing')
    background = background.replace(marker, download_guard + marker, 1)

old_download_success = '''      } else {
        console.log(`E3 Helper: 下載已開始，ID: ${downloadId}`);
        sendResponse({ success: true, downloadId: downloadId });
      }
'''
new_download_success = '''      } else {
        trackedE3Downloads.add(downloadId);
        validateTrackedDownload(downloadId);
        console.log(`E3 Helper: 下載已開始，ID: ${downloadId}`);
        sendResponse({ success: true, downloadId: downloadId });
      }
'''
background = replace_once(background, old_download_success, new_download_success, 'download tracking')

old_fetch = '''  // 驗證初始 URL；重新導向完成後會再次驗證 response.url。
  parseAllowedE3Url(url);

  try {
    const response = await fetchWithTimeout(url, {
      method: 'GET',
      credentials: 'include'
    }, 15000); // 15 秒超時

    parseAllowedE3Url(response.url);
'''
new_fetch = '''  // 驗證初始 URL；禁止自動重新導向，response.url 再做防禦性檢查。
  parseAllowedE3Url(url);

  try {
    const response = await fetchWithTimeout(url, {
      method: 'GET',
      credentials: 'include',
      redirect: 'error'
    }, 15000); // 15 秒超時

    parseAllowedE3Url(response.url);
'''
background = replace_once(background, old_fetch, new_fetch, 'redirect blocking')

new_sanitizer = r'''const RICH_CONTENT_BLOCKED_TAGS = new Set([
  'script', 'style', 'iframe', 'form', 'object', 'embed', 'applet', 'link', 'meta',
  'svg', 'math', 'template', 'button', 'input', 'select', 'textarea', 'option',
  'optgroup', 'fieldset', 'label', 'datalist', 'output'
]);
const RICH_CONTENT_BLOCKED_ATTRIBUTES = new Set([
  'id', 'name', 'form', 'formaction', 'formenctype', 'formmethod', 'formnovalidate',
  'formtarget', 'action', 'method', 'enctype', 'srcdoc', 'style', 'contenteditable',
  'autofocus'
]);

function sanitizeRichContentElement(el) {
  const tagName = String(el.tagName || '').toLowerCase();
  if (RICH_CONTENT_BLOCKED_TAGS.has(tagName)) {
    el.remove();
    return false;
  }
  Array.from(el.attributes).forEach(attr => {
    const name = attr.name.toLowerCase();
    if (name.startsWith('on') || RICH_CONTENT_BLOCKED_ATTRIBUTES.has(name) ||
        ((name === 'href' || name === 'src') && !isSafeRichContentUrl(attr.value, name))) {
      el.removeAttribute(attr.name);
    }
  });
  if (tagName === 'a' && el.getAttribute('target') === '_blank') {
    const rel = new Set((el.getAttribute('rel') || '').split(/\s+/).filter(Boolean));
    rel.add('noopener');
    rel.add('noreferrer');
    el.setAttribute('rel', [...rel].join(' '));
  }
  return true;
}

// HTML 白名單清理（用於公告/信件等 rich content）
function sanitizeHtml(html) {
  const div = document.createElement('div');
  div.innerHTML = html;
  div.querySelectorAll('*').forEach(el => sanitizeRichContentElement(el));
  return div.innerHTML;
}
'''
if 'function sanitizeRichContentElement' not in content:
    pattern = re.compile(r'// HTML 白名單清理（用於公告/信件等 rich content）\nfunction sanitizeHtml\(html\) \{.*?\n\}\n', re.S)
    content, count = pattern.subn(new_sanitizer, content, count=1)
    if count != 1:
        raise RuntimeError(f'sanitizer replacement: expected one match, got {count}')

behavior_tests = r'''

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
'''
if "sanitizer executes form-control and DOM-clobbering protections" not in tests:
    tests += behavior_tests

background_path.write_text(background)
content_path.write_text(content)
tests_path.write_text(tests)
