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
// Share the real background message handler across simulated tabs.
const background = fs.readFileSync(path.join(__dirname, '../background.js'), 'utf8');
const workers = new WeakMap();
function sharedWorker(storage) {
  if (workers.has(storage)) return workers.get(storage);
  const local = {
    get: async () => structuredClone(storage),
    set: async update => Object.assign(storage, structuredClone(update)),
    remove: async keys => keys.forEach(key => { delete storage[key]; })
  };
  let listener;
  const worker = vm.createContext({ console: { ...console, warn() {} },
    chrome: { storage: { local }, runtime: { onMessage: { addListener(fn) { listener = fn; } } } }
  });
  const start = background.indexOf('chrome.runtime.onMessage.addListener(');
  const end = background.indexOf('\n});', start) + 4;
  vm.runInContext(background.slice(background.indexOf('// Serialize daily digest'), start) + background.slice(start, end), worker);
  const runtime = { sendMessage(request, callback) {
    const promise = new Promise(resolve => {
      assert.equal(listener(request, {}, resolve), true);
    });
    promise.then(callback);
    return promise;
  } };
  const chrome = { storage: { local }, runtime };
  workers.set(storage, chrome);
  return chrome;
}
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
    chrome: sharedWorker(storage),
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
  const refreshing = app.display();
  finishWrite(); await Promise.all([pending, refreshing]);
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


test('repeated source snapshots and days retain at most five digest entries', async () => {
  const storage = fixture(), app = load(storage);
  for (let day = 0; day < 10; day++) {
    const now = fixtureTimestamp + day * 86400000;
    app.setTime(now);
    for (let snapshot = 0; snapshot < 8; snapshot++) {
      app.context.allAnnouncements[0] = { ...app.context.allAnnouncements[0], timestamp: now, title: `Day ${day} snapshot ${snapshot}` };
      await app.display();
      await app.generate();
      const keys = Object.keys(storage).filter(key => key.startsWith('dailyDigestCache:'));
      assert.ok(keys.length <= 5, `retained ${keys.length} entries`);
      assert.ok(storage[app.context.getDailyDigestCacheKey()]);
    }
  }
  assert.equal(app.calls(), 80);
  assert.deepEqual(storage.aiSettings, fixture().aiSettings);
});

test('maintenance expires old entries without trimming recent snapshots or unrelated data', async () => {
  const storage = fixture(), app = load(storage);
  await app.display(); await app.generate();
  const currentKey = app.context.getDailyDigestCacheKey();
  const current = structuredClone(storage[currentKey]);
  storage.dailyDigestCache = structuredClone(current);
  storage.notifications = [{ title: 'unrelated' }];
  for (let i = 0; i < 12; i++) {
    storage[`dailyDigestCache:old-${i}`] = { ...current, day: current.day - 8 * 86400000 };
  }
  // A later write for a different source must not evict the current snapshot.
  for (let i = 0; i < 8; i++) {
    storage[`dailyDigestCache:recent-${i}`] = { ...current, savedAt: fixtureTimestamp + i + 1 };
  }
  await app.display();
  const keys = Object.keys(storage).filter(key => key.startsWith('dailyDigestCache:'));
  assert.equal(keys.length, 9);
  assert.ok(keys.every(key => !key.includes('old-')));
  assert.deepEqual(storage[currentKey], current);
  assert.deepEqual(storage.dailyDigestCache, current);
  assert.deepEqual(storage.notifications, [{ title: 'unrelated' }]);
  assert.match(app.container().innerHTML, /保留下來的重點/);
});


test('cleanup leaves a concurrently written new snapshot untouched', async () => {
  const storage = fixture(), app = load(storage);
  await app.display(); await app.generate();
  const previous = structuredClone(storedCache(storage));
  for (let i = 0; i < 8; i++) storage[`dailyDigestCache:recent-${i}`] = { ...previous };
  let finishRead;
  app.context.chrome.storage.local.get = () => new Promise(resolve => {
    const snapshot = structuredClone(storage);
    finishRead = () => resolve(snapshot);
  });
  const pruning = app.context.pruneDailyDigestCaches();
  await new Promise(resolve => setImmediate(resolve));
  app.context.allAnnouncements[0].title = 'Concurrent source';
  const key = app.context.getDailyDigestCacheKey();
  const current = { ...previous, items: [{ ...app.context.allAnnouncements[0], type: 'announcement' }] };
  storage[key] = current;
  finishRead(); await pruning;
  assert.deepEqual(storage[key], current);
});

test('cleanup failure does not hide a successfully generated digest', async () => {
  const storage = fixture(), app = load(storage);
  await app.display();
  for (let i = 0; i < 8; i++) storage[`dailyDigestCache:expired-${i}`] = { day: 0 };
  app.context.console = { ...console, warn() {} };
  app.context.chrome.storage.local.remove = async () => { throw new Error('fixture cleanup failure'); };
  await app.generate();
  assert.match(storage[app.context.getDailyDigestCacheKey()].text, /保留下來的重點/);
  assert.match(app.container().innerHTML, /保留下來的重點/);
  assert.equal(app.button().disabled, false);
});


test('a second tab regenerating the same key survives cleanup of its old value', async () => {
  const storage = fixture(), first = load(storage), second = load(storage);
  await second.display(); await second.generate();
  const key = second.context.getDailyDigestCacheKey();
  const old = structuredClone(storage[key]);
  for (let i = 0; i < 5; i++) storage[`dailyDigestCache:recent-${i}`] = { ...old, savedAt: fixtureTimestamp + i + 1 };
  first.context.allAnnouncements[0].title = 'Different source in first tab';
  let finishRead, readStarted;
  const reading = new Promise(resolve => { readStarted = resolve; });
  const local = first.context.chrome.storage.local;
  const originalGet = local.get;
  let delayed = false;
  local.get = keys => {
    if (keys !== null || delayed) return originalGet(keys);
    delayed = true;
    const snapshot = structuredClone(storage);
    return new Promise(resolve => { finishRead = () => resolve(snapshot); readStarted(); });
  };
  const cleanup = first.context.pruneDailyDigestCaches();
  await reading;
  second.context.generateDailyDigest = async () => '{"highlights":[{"source":1,"summary":"regenerated in second tab"}],"priority":[]}';
  const generation = second.generate();
  await new Promise(resolve => setImmediate(resolve));
  finishRead();
  await Promise.all([cleanup, generation]);
  assert.match(storage[key]?.text || '', /regenerated in second tab/);
  await second.display();
  assert.match(second.container().innerHTML, /regenerated in second tab/);
  assert.ok(Object.keys(storage).filter(key => key.startsWith('dailyDigestCache:')).length <= 5);
});


test('same-key regeneration waits for an already pending removal, then remains saved', async () => {
  const storage = fixture(), first = load(storage), second = load(storage);
  await second.display(); await second.generate();
  const key = second.context.getDailyDigestCacheKey();
  const old = structuredClone(storage[key]);
  for (let i = 0; i < 5; i++) storage[`dailyDigestCache:recent-${i}`] = { ...old, savedAt: fixtureTimestamp + i + 1 };
  const expired = { ...old, day: old.day - 8 * 86400000 };
  storage[key] = expired;
  first.context.allAnnouncements[0].title = 'Other tab source';
  const local = first.context.chrome.storage.local;
  const originalRemove = local.remove;
  let finishRemove, removeStarted;
  const removing = new Promise(resolve => { removeStarted = resolve; });
  local.remove = keys => new Promise(resolve => {
    assert.ok(keys.includes(key));
    finishRemove = async () => { await originalRemove(keys); resolve(); };
    removeStarted();
  });
  const cleanup = first.context.pruneDailyDigestCaches();
  await removing;
  second.context.generateDailyDigest = async () => '{"highlights":[{"source":1,"summary":"saved after removal"}],"priority":[]}';
  const generation = second.generate();
  await new Promise(resolve => setImmediate(resolve));
  assert.deepEqual(storage[key], expired, 'writer waits until removal finishes');
  local.remove = originalRemove;
  await finishRemove();
  await Promise.all([cleanup, generation]);
  assert.match(storage[key].text, /saved after removal/);
  await second.display();
  assert.match(second.container().innerHTML, /saved after removal/);
});

test('a failed queued write preserves the previous result and allows a later write', async () => {
  const storage = fixture(), app = load(storage);
  await app.display(); await app.generate();
  const local = app.context.chrome.storage.local;
  const originalSet = local.set;
  local.set = async () => { throw new Error('fixture write failure'); };
  app.context.generateDailyDigest = async () => '{"highlights":[{"source":1,"summary":"retry succeeded"}],"priority":[]}';
  await app.generate();
  assert.match(app.container().innerHTML, /保留下來的重點/);
  local.set = originalSet;
  await app.generate();
  assert.match(storage[app.context.getDailyDigestCacheKey()].text, /retry succeeded/);
  assert.match(app.container().innerHTML, /retry succeeded/);
});


for (const invalidRequest of [false, true]) {
  test(`queued maintenance preserves all recent source snapshots${invalidRequest ? ' despite an invalid request' : ''}`, async () => {
    const storage = fixture(), first = load(storage), second = load(storage);
    await first.display(); await first.generate();
    const oldKey = first.context.getDailyDigestCacheKey();
    const old = structuredClone(storage[oldKey]);
    second.context.allAnnouncements[0].title = 'Newest source';
    const newKey = second.context.getDailyDigestCacheKey();
    const current = { ...old, savedAt: fixtureTimestamp - 1,
      text: '{"highlights":[{"source":1,"summary":"newest source overview"}],"priority":[]}',
      items: [{ ...second.context.allAnnouncements[0], type: 'announcement' }] };
    storage[newKey] = structuredClone(current);
    for (let i = 0; i < 4; i++) storage[`dailyDigestCache:recent-${i}`] = { ...old, savedAt: fixtureTimestamp + i + 1 };
    const local = first.context.chrome.storage.local;
    const originalGet = local.get;
    let finishRead, readStarted, delayed = false;
    const reading = new Promise(resolve => { readStarted = resolve; });
    local.get = keys => {
      if (keys !== null || delayed) return originalGet(keys);
      delayed = true;
      const snapshot = structuredClone(storage);
      return new Promise(resolve => { finishRead = () => resolve(snapshot); readStarted(); });
    };
    const staleMaintenance = first.context.pruneDailyDigestCaches();
    await reading;
    const newerMaintenance = second.context.pruneDailyDigestCaches();
    let invalid;
    if (invalidRequest) {
      invalid = first.context.chrome.runtime.sendMessage({
        action: 'updateDailyDigestCache', currentKey: oldKey,
        oldestDay: old.day - 6 * 86400000, cacheKey: 'unrelated-storage', cache: old
      }, () => {});
    }
    finishRead();
    await Promise.all([staleMaintenance, newerMaintenance]);
    if (invalidRequest) assert.equal((await invalid).success, false);
    assert.deepEqual(storage[newKey], current);
    assert.deepEqual(storage[oldKey], old);
    await second.display();
    assert.match(second.container().innerHTML, /newest source overview/);
    assert.equal(Object.keys(storage).filter(key => key.startsWith('dailyDigestCache:')).length, 6);
  });
}


for (const blockedStage of ['write', 'read']) {
  test(`successful write protects its own snapshot when unrelated maintenance arrives during ${blockedStage}`, async () => {
    const storage = fixture(), writer = load(storage), other = load(storage);
    await writer.display(); await writer.generate();
    const key = writer.context.getDailyDigestCacheKey();
    const old = structuredClone(storage[key]);
    other.context.allAnnouncements[0].title = 'Other tab source';
    const otherKey = other.context.getDailyDigestCacheKey();
    storage[otherKey] = { ...old, savedAt: fixtureTimestamp + 1,
      items: [{ ...other.context.allAnnouncements[0], type: 'announcement' }] };
    for (let i = 0; i < 4; i++) storage[`dailyDigestCache:recent-${i}`] = { ...old, savedAt: fixtureTimestamp + i + 2 };
    const local = writer.context.chrome.storage.local;
    const originalSet = local.set, originalGet = local.get;
    let release, started, delayed = false;
    const blocked = new Promise(resolve => { started = resolve; });
    if (blockedStage === 'write') {
      local.set = update => new Promise(resolve => {
        release = async () => { await originalSet(update); resolve(); };
        started();
      });
    } else {
      local.get = keys => {
        if (keys !== null || delayed) return originalGet(keys);
        delayed = true;
        const snapshot = structuredClone(storage);
        return new Promise(resolve => { release = () => resolve(snapshot); started(); });
      };
    }
    writer.context.generateDailyDigest = async () => '{"highlights":[{"source":1,"summary":"successful write survives maintenance"}],"priority":[]}';
    const generation = writer.generate();
    await blocked;
    const maintenance = other.context.pruneDailyDigestCaches();
    await release();
    await Promise.all([generation, maintenance]);
    assert.match(storage[key]?.text || '', /successful write survives maintenance/);
    assert.ok(Object.keys(storage).filter(key => key.startsWith('dailyDigestCache:')).length <= 5);
    await writer.display();
    assert.match(writer.container().innerHTML, /successful write survives maintenance/);
  });
}
