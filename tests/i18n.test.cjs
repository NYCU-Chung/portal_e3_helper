const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../i18n.js'), 'utf8');
async function load(browserLanguage, saved, worker = false) {
  let listener;
  const writes = [];
  const context = vm.createContext({
    ...(worker ? {} : { document: {} }),
    chrome: {
      i18n: { getUILanguage: () => browserLanguage },
      storage: {
        local: { get: async () => ({ interfaceLanguage: saved }), set: async value => writes.push(value) },
        onChanged: { addListener: callback => { listener = callback; } }
      }
    }
  });
  vm.runInContext(source, context);
  await context.E3HelperI18n.ready;
  return { i18n: context.E3HelperI18n, listener, writes };
}
test('browser default and saved preference', async () => {
  assert.equal((await load('zh-TW')).i18n.language, 'zh-TW');
  assert.equal((await load('en-US')).i18n.language, 'en');
  assert.equal((await load('zh-TW', 'en')).i18n.text('新增作業'), 'Add assignment');
  assert.equal((await load('en-US', 'zh-TW')).i18n.text('新增作業'), '新增作業');
  assert.equal((await load('en-US', 'invalid')).i18n.language, 'en');
});
test('static templates preserve source data, URLs and interpolation', async () => {
  const { i18n } = await load('en');
  const ui = i18n.template;
  const course = '作業與課程的英文翻譯';
  assert.equal(ui`課程：${course}`, `Course: ${course}`);
  const body = '<p>課程：原文不應被翻譯</p>';
  assert.equal(ui`<div title="更多操作">${body}</div>`, `<div title="More actions">${body}</div>`);
  assert.equal(ui`還有 ${0} 分鐘`, 'Due in 0 m');
  assert.equal(i18n.text('儲存設定'), 'Save settings');
});
test('save persists language, rejects invalid values and restores Chinese', async () => {
  const { i18n, writes } = await load('en');
  await i18n.save('zh-TW');
  assert.equal(writes[0].interfaceLanguage, 'zh-TW');
  assert.equal(i18n.text('儲存設定'), '儲存設定');
  await assert.rejects(i18n.save('fr'));
  assert.equal(writes.length, 1);
});
test('background follows changes while open pages remain consistent until refresh', async () => {
  const worker = await load('zh-TW', 'zh-TW', true);
  worker.listener({ interfaceLanguage: { newValue: 'en' } }, 'local');
  assert.equal(worker.i18n.text('公告'), 'Announcements');
  const page = await load('zh-TW', 'zh-TW');
  page.listener({ interfaceLanguage: { newValue: 'en' } }, 'local');
  assert.equal(page.i18n.text('公告'), '公告');
});
test('manifest loads localization first and has valid store locales', () => {
  const root = require('node:path').join(__dirname, '..');
  const manifest = JSON.parse(fs.readFileSync(require('node:path').join(root, 'manifest.json')));
  const scripts = manifest.content_scripts[0].js;
  assert.ok(scripts.indexOf('i18n.js') < scripts.indexOf('content.js'));
  for (const locale of ['en', 'zh_TW']) {
    const messages = JSON.parse(fs.readFileSync(require('node:path').join(root, '_locales', locale, 'messages.json')));
    assert.ok(messages.extensionDescription.message);
    assert.ok(messages.extensionDescription.message.length <= 112, `${locale} Safari store description exceeds 112 characters`);
    assert.equal(typeof messages.extensionDescription.description, 'string');
    assert.ok(messages.extensionDescription.description.length <= 112);
  }
});

function callbackOnlyStorage(saved, errorMethod) {
  const data = { interfaceLanguage: saved };
  const context = vm.createContext({ document: {}, chrome: {
    i18n: { getUILanguage: () => 'zh-TW' }, runtime: {},
    storage: { onChanged: { addListener() {} }, local: {
      get(keys, callback) {
        queueMicrotask(() => {
          if (errorMethod === 'get') context.chrome.runtime.lastError = { message: 'Read denied' };
          callback({ ...data });
          delete context.chrome.runtime.lastError;
        });
      },
      set(update, callback) {
        queueMicrotask(() => {
          if (errorMethod === 'set') context.chrome.runtime.lastError = { message: 'Write denied' };
          else Object.assign(data, update);
          callback();
          delete context.chrome.runtime.lastError;
        });
      }
    } }
  }, queueMicrotask });
  vm.runInContext(source, context);
  return { i18n: context.E3HelperI18n, data };
}
test('Chrome 93 callback-only storage initializes and persists the language', async () => {
  const { i18n, data } = callbackOnlyStorage('en');
  await i18n.ready;
  assert.equal(i18n.language, 'en');
  assert.equal(i18n.text('新增作業'), 'Add assignment');
  await i18n.save('zh-TW');
  assert.equal(data.interfaceLanguage, 'zh-TW');
  assert.equal(i18n.language, 'zh-TW');
});
test('callback storage errors reject without changing the language', async () => {
  const readFailure = callbackOnlyStorage('en', 'get');
  await assert.rejects(readFailure.i18n.ready, /Read denied/);
  const writeFailure = callbackOnlyStorage('en', 'set');
  await writeFailure.i18n.ready;
  await assert.rejects(writeFailure.i18n.save('zh-TW'), /Write denied/);
  assert.equal(writeFailure.i18n.language, 'en');
  assert.equal(writeFailure.data.interfaceLanguage, 'en');
});
