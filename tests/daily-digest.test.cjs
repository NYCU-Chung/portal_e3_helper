const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../content.js'), 'utf8');
const digestCode = source.slice(source.indexOf('// 顯示公告與信件列表'), source.indexOf('// 翻譯文字（'));
const rendering = source.slice(source.indexOf('// 只接受既有來源編號'), source.indexOf('async function generateDailyDigest'));
const fixtureTimestamp = new Date(2026, 9, 6, 12).getTime();
class FixtureDate extends Date {
  constructor(...args) { super(...(args.length ? args : [fixtureTimestamp])); }
  static now() { return fixtureTimestamp; }
}
function storedCacheKey(storage) { return Object.keys(storage).find(key => key.startsWith('dailyDigestCache:')) || 'dailyDigestCache'; }
function storedCache(storage) { return storage[storedCacheKey(storage)]; }
const fixture = () => ({ aiSettings: { enabled: true, openaiSummaryApiKey: 'fixture' } });
function load(storage, now = fixtureTimestamp) {
  class ControlledDate extends Date {
    constructor(...args) { super(...(args.length ? args : [now])); }
    static now() { return now; }
  }
  let elements = {};
  let calls = 0, lastConfig;
  const makeElement = () => ({ dataset: {}, style: {}, innerHTML: '', addEventListener(type, fn) { this[type] = fn; } });
  const list = { set innerHTML(html) {
    elements = {};
    if (!html.includes('id="e3-helper-daily-digest"')) return;
    elements['e3-helper-generate-daily-digest'] = makeElement();
    const container = elements['e3-helper-daily-digest'] = makeElement();
    container.style.display = html.match(/id="e3-helper-daily-digest" style="display: ([^;]+);/)[1];
    container.innerHTML = html.split('aria-live="polite">')[1].split('\n      </section>')[0].replace(/<\/div>\s*$/, '');
  } };
  const item = { id: 1, timestamp: now, title: '課中提問', courseName: '課程', author: '老師', url: 'https://e3.nycu.edu.tw/source' };
  const context = vm.createContext({
    console, URL, Date: ControlledDate, window: { location: { href: 'https://e3.nycu.edu.tw/' } },
    E3HelperI18n: { language: 'zh-TW' },
    ui: (parts, ...values) => parts.reduce((result, part, i) => result + part + (values[i] ?? ''), ''),
    uiText: text => text, escapeHtml: text => String(text).replaceAll('<', '&lt;'), getTimeAgoText: () => '剛剛',
    allAnnouncements: [item], allMessages: [], readAnnouncements: new Set(), readMessages: new Set(),
    document: { querySelector: selector => selector.includes('assignment-list') ? list : null, querySelectorAll: () => [], getElementById: id => elements[id] || null },
    chrome: { storage: { local: { get: async () => structuredClone(storage), set: async update => Object.assign(storage, structuredClone(update)) } } },
    generateDailyDigest: async (items, config) => { lastConfig = config; calls++; return '{"highlights":[{"source":1,"summary":"保留下來的重點"}],"priority":[]}'; },
    showTemporaryMessage() {}
  });
  const configCode = source.slice(source.indexOf('function getAISummaryConfig('), source.indexOf('function updateAIProviderFields('));
  vm.runInContext(configCode + digestCode + rendering, context);
  return { context, setTime: value => {now=value;}, display: () => context.displayAnnouncements(), generate: () => elements['e3-helper-generate-daily-digest'].click(), container: () => elements['e3-helper-daily-digest'], button: () => elements['e3-helper-generate-daily-digest'], calls: () => calls, config: () => lastConfig };
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
  const saved = structuredClone(storedCache(storage));
  storedCache(storage).day -= 86400000;
  const yesterday = load(storage);
  await yesterday.display();
  assert.equal(yesterday.container().style.display, 'none');
  storage[storedCacheKey(storage)] = { ...saved, language: 'en' };
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
  const previous = structuredClone(storedCache(storage));
  app.context.generateDailyDigest = async () => { throw new Error('fixture failure'); };
  await app.generate();
  assert.equal(app.container().style.display, 'block');
  assert.match(app.container().innerHTML, /保留下來的重點/);
  assert.deepEqual(storedCache(storage), previous);
});


test('digest generation uses the chosen Gemini provider and preserves cached output', async () => {
  const storage = { aiSettings: { enabled: true, summaryProvider: 'gemini', geminiApiKey: 'gemini-fixture', geminiModel: 'future-model', openaiSummaryApiKey: 'other-provider' } };
  const app = load(storage);
  await app.display();
  await app.generate();
  assert.equal(app.config().provider, 'gemini');
  assert.equal(app.config().apiKey, 'gemini-fixture');
  assert.equal(app.config().model, 'future-model');
  assert.match(storedCache(storage).text, /保留下來的重點/);
});

test('cached digest is hidden when current source records belong to another account', async () => {
  const storage = fixture(), app = load(storage);
  await app.display(); await app.generate();
  const reopened = load(storage);
  reopened.context.allAnnouncements[0].title = 'Different account announcement';
  await reopened.display();
  assert.equal(reopened.container().style.display,'none');
});
test('in-flight generation stays locked after list rerender', async () => {
  const app = load(fixture());
  let finish, requests = 0;
  app.context.generateDailyDigest = () => { requests++; return new Promise(resolve => {finish=resolve;}); };
  await app.display();
  const pending = app.generate(); await Promise.resolve();
  await app.display();
  assert.equal(app.button().disabled,true);
  const duplicate = app.generate(); await new Promise(resolve => setImmediate(resolve));
  assert.equal(requests,1);
  await duplicate;
  finish('{"highlights":[{"source":1,"summary":"finished"}],"priority":[]}');
  await pending;
  assert.equal(app.button().disabled,false);
});
for (const output of ['not JSON', '{"highlights":[{"source":99,"summary":"invalid"}],"priority":[]}']) {
  test(`unusable digest preserves previous successful cache: ${output}`, async () => {
    const storage = fixture(), app = load(storage);
    await app.display(); await app.generate();
    const previous = structuredClone(storedCache(storage));
    app.context.generateDailyDigest = async () => output;
    await app.generate();
    assert.deepEqual(storedCache(storage),previous);
  });
}
test('generation cannot save a result after source records change', async () => {
  const storage = fixture(), app = load(storage);
  let finish;
  app.context.generateDailyDigest = () => new Promise(resolve => {finish=resolve;});
  await app.display(); const pending=app.generate(); await new Promise(resolve => setImmediate(resolve));
  app.context.allAnnouncements[0] = {...app.context.allAnnouncements[0],title:'Other account'};
  finish('{"highlights":[{"source":1,"summary":"stale"}],"priority":[]}');
  await pending;
  assert.equal(storedCache(storage),undefined);
});
test('missing settings release the generation lock for a later configured attempt', async () => {
  const storage = {aiSettings:{enabled:false}}, app=load(storage);
  await app.display(); await app.generate();
  storage.aiSettings=fixture().aiSettings;
  await app.generate();
  assert.equal(app.calls(),1);
});

test('refresh during a delayed cache write cannot replace the current source overview', async () => {
  const storage = fixture(), app = load(storage);
  await app.display(); await app.generate();
  let finishWrite, writeStarted;
  const writing = new Promise(resolve => {writeStarted=resolve;});
  app.context.chrome.storage.local.set = update => new Promise(resolve => {
    finishWrite = () => {Object.assign(storage,structuredClone(update));resolve();};
    writeStarted();
  });
  const pending=app.generate();
  await writing;
  app.context.allAnnouncements[0] = {...app.context.allAnnouncements[0],title:'Refreshed source'};
  const current = {day:new FixtureDate().setHours(0,0,0,0),language:'zh-TW',text:'{"highlights":[{"source":1,"summary":"current overview"}],"priority":[]}',items:[{...app.context.allAnnouncements[0],type:'announcement'}]};
  const key=app.context.getDailyDigestCacheKey(current.items);
  storage[key]=structuredClone(current);
  await app.display();
  finishWrite(); await pending;
  assert.deepEqual(storage[key],current);
  assert.match(app.container().innerHTML,/current overview/);
});

test('legacy cache remains readable with source validation', async () => {
  const storage=fixture(), app=load(storage);
  await app.display(); await app.generate();
  const key=storedCacheKey(storage);
  storage.dailyDigestCache=storage[key]; delete storage[key];
  const reopened=load(storage); await reopened.display();
  assert.match(reopened.container().innerHTML,/保留下來的重點/);
});
test('controlled midnight during cache writing discards yesterday completion', async () => {
  const beforeMidnight = new Date(2026,9,6,23,59,59).getTime();
  const storage=fixture(), app=load(storage,beforeMidnight);
  let finishWrite, started;
  const writing=new Promise(resolve => {started=resolve;});
  app.context.chrome.storage.local.set=update=>new Promise(resolve=>{
    finishWrite=()=>{Object.assign(storage,structuredClone(update));resolve();}; started();
  });
  await app.display(); const pending=app.generate(); await writing;
  app.setTime(beforeMidnight+2000);
  finishWrite(); await pending;
  assert.equal(app.container().style.display,'none');
  assert.equal(app.button().disabled,false);
});
