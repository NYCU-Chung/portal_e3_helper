const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('content.js', 'utf8');
const fn = name => { const start = source.search(new RegExp(`(?:async )?function ${name}\\(`)); return start < 0 ? '' : source.slice(start, source.indexOf('\n}', start) + 2); };
async function load(date) {
  const row = { querySelector(selector) { return ({ 'a.mail_link': { href: 'https://e3p.nycu.edu.tw/local/dcpcmail/view.php?m=1' }, '.mail_summary': { textContent: 'Old message', querySelector: () => null }, '.mail_users': { textContent: 'Teacher' }, '.mail_date': { textContent: date, getAttribute: attr => attr === 'title' ? date : null, querySelector: () => null } })[selector]; }, classList: { contains: () => false } };
  const context = vm.createContext({ console: { log() {}, error() {}, warn() {} }, Date, allCourses: [{ id: 1, fullname: 'course' }], allMessages: [], isOnE3Site: () => true, document: { querySelector: () => ({}) }, fetch: async () => ({ ok: true, text: async () => '' }), DOMParser: class { parseFromString() { return { querySelectorAll: () => [row] }; } }, chrome: { storage: { local: { set: async () => {} } } }, uiText: s => s, ui: (s, ...v) => s.reduce((a,p,i)=>a+p+(v[i]??''),''), showTemporaryMessage() {} });
  vm.runInContext(fn('parseE3Timestamp') + fn('getMessageTimestamp') + fn('loadMessages') + fn('getTimeAgoText'), context);
  await context.loadMessages();
  return context;
}
test('localized mail date with weekday preserves the sent date instead of sync time', async () => {
  const app = await load('2026年9月23日, 星期三, 21:02');
  assert.equal(app.allMessages[0].timestamp, new Date(2026,8,23,21,2).getTime());
});
test('unparseable mail dates stay unknown instead of appearing newly received', async () => {
  const app = await load('');
  assert.equal(app.allMessages[0].timestamp, null);
  assert.equal(app.getTimeAgoText(null), '時間未知');
});

for (const [date, expected] of [
  ['14 October 2026, 12:08 PM', new Date(2026,9,14,12,8).getTime()],
  ['30 September 2026, 8:14 PM', new Date(2026,8,30,20,14).getTime()],
  ['22 September 2026, 10:20 AM', new Date(2026,8,22,10,20).getTime()],
  ['2026年9月23日,21:02', new Date(2026,8,23,21,2).getTime()],
  ['2026-09-23T21:02:00+08:00', Date.parse('2026-09-23T21:02:00+08:00')],
  ['2026年2月30日,21:02', null]
]) {
  test(`mail date ${date} is parsed without fabricating a sync timestamp`, async () => {
    const app = await load(date);
    assert.equal(app.allMessages[0].timestamp, expected);
  });
}

test('daily digest excludes unknown dates, yesterday and future dates', () => {
  const today = new Date(); today.setHours(0,0,0,0);
  const tomorrow = new Date(today); tomorrow.setDate(tomorrow.getDate()+1);
  const context = vm.createContext({ Date, allAnnouncements: [], allMessages: [
    { id: 'unknown', timestamp: null }, { id: 'invalid', timestamp: NaN },
    { id: 'old', timestamp: today.getTime()-1 }, { id: 'future', timestamp: tomorrow.getTime() },
    { id: 'today', timestamp: today.getTime()+1 }
  ] });
  vm.runInContext(fn('getCurrentDailyDigestItems'), context);
  assert.deepEqual(Array.from(context.getCurrentDailyDigestItems(), item => item.id), ['today']);
});
