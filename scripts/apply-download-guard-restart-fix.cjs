const fs = require('node:fs');

const backgroundPath = 'background.js';
const testPath = 'tests/security-hardening.test.cjs';

let background = fs.readFileSync(backgroundPath, 'utf8');
let tests = fs.readFileSync(testPath, 'utf8');

const oldGuard = `const trackedE3Downloads = new Set();

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
    console.warn(\`E3 Helper: 已取消重新導向至非 E3 網域的下載 - \${finalUrl}\`);
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

chrome.downloads?.onChanged?.addListener(delta => {
  if (!trackedE3Downloads.has(delta.id)) return;
  if (delta.finalUrl?.current && cancelDownloadIfFinalUrlDisallowed(delta.id, delta.finalUrl.current)) return;
  if (delta.state?.current === 'complete' || delta.error) {
    validateTrackedDownload(delta.id);
    trackedE3Downloads.delete(delta.id);
  }
});
`;

const newGuard = `function isExtensionInitiatedE3Download(item) {
  if (!item || item.byExtensionId !== chrome.runtime.id) return false;
  try {
    parseAllowedE3Url(item.url);
    return true;
  } catch {
    return false;
  }
}

function cancelDownloadIfFinalUrlDisallowed(downloadId, finalUrl) {
  try {
    parseAllowedE3Url(finalUrl);
    return false;
  } catch {
    chrome.downloads.cancel(downloadId, () => {
      void chrome.runtime.lastError;
      if (chrome.downloads.removeFile) {
        chrome.downloads.removeFile(downloadId, () => void chrome.runtime.lastError);
      }
    });
    console.warn(\`E3 Helper: 已取消重新導向至非 E3 網域的下載 - \${finalUrl}\`);
    return true;
  }
}

function validateExtensionE3Download(downloadId) {
  chrome.downloads.search({ id: downloadId }, items => {
    if (chrome.runtime.lastError) return;
    const item = items?.[0];
    if (!isExtensionInitiatedE3Download(item)) return;
    if (item.finalUrl) cancelDownloadIfFinalUrlDisallowed(downloadId, item.finalUrl);
  });
}

chrome.downloads?.onChanged?.addListener(delta => {
  if (!delta.finalUrl && !delta.state && !delta.error) return;
  // Reconstruct ownership from the Downloads API instead of relying on worker
  // memory, so redirect validation survives Manifest V3 worker suspension.
  validateExtensionE3Download(delta.id);
});
`;

if (!background.includes(oldGuard)) {
  throw new Error('Expected download guard block was not found');
}
background = background.replace(oldGuard, newGuard);

const oldDownloadStart = `        trackedE3Downloads.add(downloadId);
        validateTrackedDownload(downloadId);
        console.log(\`E3 Helper: 下載已開始，ID: \${downloadId}\`);`;
const newDownloadStart = `        validateExtensionE3Download(downloadId);
        console.log(\`E3 Helper: 下載已開始，ID: \${downloadId}\`);`;
if (!background.includes(oldDownloadStart)) {
  throw new Error('Expected download-start tracking block was not found');
}
background = background.replace(oldDownloadStart, newDownloadStart);

const testStart = "test('download final URL guard cancels off-domain redirects', () => {";
const testIndex = tests.indexOf(testStart);
if (testIndex === -1) throw new Error('Expected download guard test was not found');

tests = tests.slice(0, testIndex) + `test('download final URL guard survives worker restarts by reconstructing ownership', () => {
  const urlCode = sliceBetween(background, 'const E3_ALLOWED_HOSTS', '// NYCU E3 Helper - Background Script');
  const guardCode = sliceBetween(background, 'function isExtensionInitiatedE3Download', '// 監聽來自 content script 的訊息');
  const cancelled = [];
  const removed = [];
  const items = new Map();
  let onChanged;
  const context = {
    URL, Set,
    console: { warn() {} },
    chrome: {
      runtime: { id: 'helper-extension-id', lastError: null },
      downloads: {
        cancel(id, cb) { cancelled.push(id); cb(); },
        removeFile(id, cb) { removed.push(id); cb(); },
        search(query, cb) { cb(items.has(query.id) ? [items.get(query.id)] : []); },
        onChanged: { addListener(listener) { onChanged = listener; } }
      }
    }
  };
  vm.createContext(context);
  vm.runInContext(\`${'${urlCode}'}\\n${'${guardCode}'}\\nthis.validateDownload = validateExtensionE3Download;\`, context);

  items.set(7, {
    id: 7,
    byExtensionId: 'helper-extension-id',
    url: 'https://e3.nycu.edu.tw/pluginfile.php/1/file.pdf',
    finalUrl: 'https://e3.nycu.edu.tw/pluginfile.php/1/file.pdf'
  });
  context.validateDownload(7);
  assert.deepEqual(cancelled, []);

  items.set(8, {
    id: 8,
    byExtensionId: 'helper-extension-id',
    url: 'https://e3.nycu.edu.tw/pluginfile.php/2/file.pdf',
    finalUrl: 'https://evil.invalid/file.pdf'
  });
  context.validateDownload(8);
  assert.deepEqual(cancelled, [8]);
  assert.deepEqual(removed, [8]);

  items.set(9, {
    id: 9,
    byExtensionId: 'helper-extension-id',
    url: 'https://e3.nycu.edu.tw/pluginfile.php/3/file.pdf',
    finalUrl: 'https://evil.invalid/restarted.pdf'
  });
  // No in-memory registration is performed here: the event simulates a fresh
  // service worker waking after the original worker was suspended.
  onChanged({ id: 9, state: { current: 'complete' } });
  assert.deepEqual(cancelled, [8, 9]);
  assert.deepEqual(removed, [8, 9]);

  items.set(10, {
    id: 10,
    byExtensionId: 'another-extension',
    url: 'https://e3.nycu.edu.tw/pluginfile.php/4/file.pdf',
    finalUrl: 'https://evil.invalid/not-ours.pdf'
  });
  onChanged({ id: 10, finalUrl: { current: 'https://evil.invalid/not-ours.pdf' } });
  assert.deepEqual(cancelled, [8, 9]);
});
`;

fs.writeFileSync(backgroundPath, background);
fs.writeFileSync(testPath, tests);
