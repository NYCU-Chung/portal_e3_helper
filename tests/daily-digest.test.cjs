const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../content.js'), 'utf8');
const digestCode = source.slice(source.indexOf('// 顯示公告與信件列表'), source.indexOf('// 翻譯文字（'));
const rendering = source.slice(source.indexOf('// 只接受既有來源編號'), source.indexOf('async function generateDailyDigest'));
const fixture = () => ({ aiSettings: { enabled: true, openaiSummaryApiKey: 'fixture' } });
function load(storage) {
  let elements = {};
  let calls = 0;
  const makeElement = () => ({ dataset: {}, style: {}, innerHTML: '', addEventListener(type, fn) { this[type] = fn; } });
  const list = { set innerHTML(html) {
    elements = {};
    if (!html.includes('id="e3-helper-daily-digest"')) return;
    elements['e3-helper-generate-daily-digest'] = makeElement();
    const container = elements['e3-helper-daily-digest'] = makeElement();
    container.style.display = html.match(/id="e3-helper-daily-digest" style="display: ([^;]+);/)[1];
    container.innerHTML = html.split('aria-live="polite">')[1].split('\n      </section>')[0].replace(/<\/div>\s*$/, '');
  } };
  const item = { id: 1, timestamp: Date.now(), title: '課中提問', courseName: '課程', author: '老師', url: 'https://e3.nycu.edu.tw/source' };
  const context = vm.createContext({
    console, URL, Date, window: { location: { href: 'https://e3.nycu.edu.tw/' } },
    E3HelperI18n: { language: 'zh-TW' },
    ui: (parts, ...values) => parts.reduce((result, part, i) => result + part + (values[i] ?? ''), ''),
    uiText: text => text, escapeHtml: text => String(text).replaceAll('<', '&lt;'), getTimeAgoText: () => '剛剛',
    allAnnouncements: [item], allMessages: [], readAnnouncements: new Set(), readMessages: new Set(),
    document: { querySelector: selector => selector.includes('assignment-list') ? list : null, querySelectorAll: () => [], getElementById: id => elements[id] || null },
    chrome: { storage: { local: { get: async () => structuredClone(storage), set: async update => Object.assign(storage, structuredClone(update)) } } },
    generateDailyDigest: async () => { calls++; return '{"highlights":[{"source":1,"summary":"保留下來的重點"}],"priority":[]}'; },
    showTemporaryMessage() {}
  });
  vm.runInContext(digestCode + rendering, context);
  return { context, display: () => context.displayAnnouncements(), generate: () => elements['e3-helper-generate-daily-digest'].click(), container: () => elements['e3-helper-daily-digest'], calls: () => calls };
}
test('generated digest survives helper reopening and a fresh content-script session without another AI call', async () => {
  const storage = fixture();
  const first = load(storage);
  await first.display();
  await first.generate();
  assert.match(first.container().innerHTML, /保留下來的重點/);
  await first.display();
  assert.equal(first.container().style.display, 'block');
  assert.match(first.container().innerHTML, /保留下來的重點/);
  assert.equal(first.calls(), 1);
  const reopened = load(storage);
  await reopened.display();
  assert.equal(reopened.container().style.display, 'block');
  assert.match(reopened.container().innerHTML, /保留下來的重點/);
  assert.equal(reopened.calls(), 0);
});

test('yesterday or another language does not appear as today’s digest', async () => {
  const storage = fixture();
  const first = load(storage);
  await first.display();
  await first.generate();
  const saved = structuredClone(storage.dailyDigestCache);
  storage.dailyDigestCache.day -= 86400000;
  const yesterday = load(storage);
  await yesterday.display();
  assert.equal(yesterday.container().style.display, 'none');
  storage.dailyDigestCache = { ...saved, language: 'en' };
  const otherLanguage = load(storage);
  await otherLanguage.display();
  assert.equal(otherLanguage.container().style.display, 'none');
});

test('generation completing after helper reopening updates the current container', async () => {
  const storage = fixture();
  const app = load(storage);
  let finish;
  app.context.generateDailyDigest = () => new Promise(resolve => { finish = resolve; });
  await app.display();
  const pending = app.generate();
  await Promise.resolve();
  await app.display();
  finish('{"highlights":[{"source":1,"summary":"稍後完成的重點"}],"priority":[]}');
  await pending;
  assert.equal(app.container().style.display, 'block');
  assert.match(app.container().innerHTML, /稍後完成的重點/);
  const reopened = load(storage);
  await reopened.display();
  assert.match(reopened.container().innerHTML, /稍後完成的重點/);
});

test('failed regeneration preserves the previous successful digest', async () => {
  const storage = fixture();
  const app = load(storage);
  await app.display();
  await app.generate();
  const previous = structuredClone(storage.dailyDigestCache);
  app.context.generateDailyDigest = async () => { throw new Error('fixture failure'); };
  await app.generate();
  assert.equal(app.container().style.display, 'block');
  assert.match(app.container().innerHTML, /保留下來的重點/);
  assert.deepEqual(storage.dailyDigestCache, previous);
});
