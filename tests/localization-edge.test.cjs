const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const content = fs.readFileSync(path.join(__dirname, '../content.js'), 'utf8');
const catalog = fs.readFileSync(path.join(__dirname, '../i18n.js'), 'utf8');
function fn(name) {
  const start = content.search(new RegExp(`(?:async )?function ${name}\\(`));
  assert.ok(start >= 0, name);
  return content.slice(start, content.indexOf('\n}', start) + 2);
}
async function app(language = 'en', data = {}, extra = {}) {
  const chrome = { i18n: { getUILanguage: () => language }, runtime: {}, storage: {
    onChanged: { addListener() {} }, local: {
      get: async () => ({ ...data }), set: async value => Object.assign(data, value)
    }
  } };
  const context = vm.createContext({ console: { log() {}, warn() {}, error() {} }, document: {}, chrome, ...extra });
  vm.runInContext(catalog, context);
  await context.E3HelperI18n.ready;
  context.uiText = context.E3HelperI18n.text;
  context.ui = context.E3HelperI18n.template;
  return context;
}
test('unrelated settings saves in an older tab preserve the newer shared language', async () => {
  const data = { interfaceLanguage: 'zh-TW' };
  const fields = Object.fromEntries(['language','enable-ai','ai-provider','gemini-key','gemini-model-id','openai-summary-key','openai-summary-model','theme'].map(id => [`e3-helper-${id}`, { value: '', checked: false }]));
  fields['e3-helper-language'].value = 'zh-TW';
  fields['e3-helper-theme'].value = 'dark';
  let reloads = 0;
  const context = await app('zh-TW', data, { document: { getElementById: id => fields[id] }, applyThemePreference() {}, showTemporaryMessage() {}, window: { location: { reload() { reloads++; } } } });
  vm.runInContext(fn('saveAISettings'), context);
  data.interfaceLanguage = 'en'; // Another tab saves after this tab initialized.
  await context.saveAISettings();
  assert.equal(data.interfaceLanguage, 'en');
  assert.equal(data.themePreference, 'dark');
  assert.equal(context.E3HelperI18n.language, 'zh-TW');
  assert.equal(reloads, 0);
  // A deliberate change in the language control still persists and reloads.
  fields['e3-helper-language'].value = 'en';
  await context.saveAISettings();
  assert.equal(data.interfaceLanguage, 'en');
  assert.equal(context.E3HelperI18n.language, 'en');
  assert.equal(reloads, 1);
});
test('grade points and countdown minutes have different English units', async () => {
  const container = {};
  const context = await app('en', {}, { document: { querySelector: () => container, getElementById: () => null }, escapeHtml: value => value });
  vm.runInContext(fn('displayGradeStats') + fn('showCourseGradeDetails') + fn('formatCountdown') + fn('getTimeAgoCompact'), context);
  context.displayGradeStats({ evaluatedWeight: 1, progress: 100, currentPerformance: 85, optimisticScore: 85 }, { items: [{ evaluated: true, name: '原文分數', score: 85, weight: 100 }] });
  assert.match(container.innerHTML, /85 points/);
  assert.ok(container.innerHTML.includes('原文分數'));
  context.gradeData = { 1: { course: { fullname: '課程原文' }, stats: { progress: 100, currentPerformance: 85, optimisticScore: 85 }, grades: { items: [{ evaluated: true, name: '原文分數', score: 85, weight: 100 }] } } };
  context.showCourseGradeDetails(1);
  assert.match(container.innerHTML, /85 points/);
  assert.match(context.getTimeAgoCompact(Date.now() - 300000), /^5m$/);
  const countdown = context.formatCountdown(Date.now() + 300000);
  assert.match(countdown.text, /\dm /);
  assert.ok(!countdown.text.includes('points'));
  await context.E3HelperI18n.save('zh-TW');
  context.displayGradeStats({ evaluatedWeight: 1, progress: 100, currentPerformance: 85, optimisticScore: 85 }, { items: [{ evaluated: true, name: '原文分數', score: 85, weight: 100 }] });
  assert.match(container.innerHTML, /85 分/);
});
test('urgent alerts localize missing courses while preserving actual course text', async () => {
  const data = {};
  const context = await app('en', data, { updateNotificationBadge: async () => {} });
  vm.runInContext(fn('checkUrgentAssignments'), context);
  const now = Date.now();
  await context.checkUrgentAssignments([{ eventId: '1', name: '作業原文', deadline: now + 60000 }, { eventId: '2', name: '另一作業', course: '課程中文原文', deadline: now + 60000 }], now);
  assert.ok(data.urgentAssignmentNotifications[0].message.includes('(Unknown course)'));
  assert.ok(data.urgentAssignmentNotifications[1].message.includes('課程中文原文'));
});
test('sync stores missing message subjects as raw empty data and renders in the current language', async () => {
  const data = {};
  const row = { classList: { contains: () => false }, querySelector(selector) {
    if (selector === 'a.mail_link') return { href: 'https://e3p.nycu.edu.tw/mail?m=1' };
    if (selector === '.mail_summary') return { textContent: '', querySelector: () => null };
    return null;
  } };
  const context = await app('en', data, { allCourses: [{ id: 1, fullname: '課程原文' }], allMessages: [], isOnE3Site: () => true,
    document: { querySelector: () => ({}) }, showTemporaryMessage() {}, fetch: async () => ({ ok: true, text: async () => 'fixture' }),
    DOMParser: class { parseFromString() { return { querySelectorAll: () => [row] }; } } });
  vm.runInContext(fn('loadMessages'), context);
  await context.loadMessages();
  assert.equal(data.messages[0].title, '');
  vm.runInContext(fn('getItemTitle'), context);
  assert.equal(context.getItemTitle(data.messages[0]), '(No subject)');
  await context.E3HelperI18n.save('zh-TW');
  assert.equal(context.getItemTitle(data.messages[0]), '(無主旨)');
  assert.equal(context.getItemTitle({ type: 'message', title: '原文主旨' }), '原文主旨');
  assert.equal(data.messages[0].title, '');
});
