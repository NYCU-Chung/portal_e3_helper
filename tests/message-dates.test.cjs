const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('content.js', 'utf8');
const fn = name => { const start = source.search(new RegExp(`(?:async )?function ${name}\\(`)); return start < 0 ? '' : source.slice(start, source.indexOf('\n}', start) + 2); };
async function load(date) {
  const row = { querySelector(selector) { return ({ 'a.mail_link': { href: 'https://e3p.nycu.edu.tw/local/dcpcmail/view.php?m=1' }, '.mail_summary': { textContent: 'Old message', querySelector: () => null }, '.mail_users': { textContent: 'Teacher' }, '.mail_date': { textContent: date, getAttribute: attr => attr === 'title' ? date : null, querySelector: () => null } })[selector]; }, classList: { contains: () => false } };
  const context = vm.createContext({ console: { log() {}, error() {}, warn() {} }, Date, allCourses: [{ id: 1, fullname: 'course' }], allMessages: [], isOnE3Site: () => true, document: { querySelector: () => ({}) }, fetch: async () => ({ ok: true, text: async () => '' }), DOMParser: class { parseFromString() { return { querySelectorAll: () => [row] }; } }, chrome: { storage: { local: { set: async () => {} } } }, uiText: s => s, ui: (s, ...v) => s.reduce((a,p,i)=>a+p+(v[i]??''),''), showTemporaryMessage() {} });
  vm.runInContext(fn('parseE3Timestamp') + fn('getMessageTimestamp') + fn('loadMessages') + fn('restoreMessagesFromStorage') + fn('getTimeAgoText'), context);
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

test('upgrading cached mail refreshes legacy timestamps before they enter the digest', async () => {
  const app = await load('30 September 2026, 8:14 PM');
  await app.restoreMessagesFromStorage([{ id: 'legacy', timestamp: Date.now(), title: 'old mail' }]);
  assert.equal(app.allMessages[0].timestamp, new Date(2026,8,30,20,14).getTime());
  assert.equal(app.allMessages[0].timestampVersion, 1);
});
test('legacy mail stays unknown off E3 and already parsed cached mail is preserved', async () => {
  const app = await load('14 October 2026, 12:08 PM');
  app.isOnE3Site = () => false;
  const before = Date.now();
  const old = { id: 'legacy', timestamp: before, title: 'original', isRead: true };
  await app.restoreMessagesFromStorage([old]);
  assert.equal(app.allMessages[0].timestamp, null);
  assert.equal(app.allMessages[0].title, 'original');
  assert.equal(app.allMessages[0].isRead, true);
  assert.equal(old.timestamp, before); // Input storage data is not mutated.
  const parsed = { id: 'parsed', timestamp: new Date(2026,9,14,12,8).getTime(), timestampVersion: 1 };
  app.loadMessages = () => { throw new Error('Parsed mail should not refresh'); };
  app.isOnE3Site = () => true;
  await app.restoreMessagesFromStorage([parsed]);
  assert.equal(app.allMessages[0].timestamp, parsed.timestamp);
});
test('digest cache with legacy mail is invalidated even if its fabricated date is today', () => {
  const today = new Date(); today.setHours(0,0,0,0);
  const context = vm.createContext({ Date, E3HelperI18n: { language: 'zh-TW' }, dailyDigestCache: {
    day: today.getTime(), language: 'zh-TW', text: 'old summary', items: [{ type: 'message', timestamp: Date.now() }]
  } });
  vm.runInContext(fn('getSavedDailyDigestHTML'), context);
  assert.equal(context.getSavedDailyDigestHTML(), '');
});

test('upgrade refresh keeps saved mail with unknown dates when the network fails', async () => {
  const app = await load('14 October 2026, 12:08 PM');
  app.fetch = async () => { throw new Error('offline'); };
  let saved;
  app.chrome.storage.local.set = async data => { saved = data; };
  await app.restoreMessagesFromStorage([{ id: 'msg-1-8', title: 'saved', timestamp: Date.now(), isRead: true }]);
  assert.equal(app.allMessages.length, 1);
  assert.equal(app.allMessages[0].title, 'saved');
  assert.equal(app.allMessages[0].timestamp, null);
  assert.equal(app.allMessages[0].isRead, true);
  assert.equal(saved.messages[0].title, 'saved');
});
test('an E3 login page cannot overwrite saved mail with an empty mailbox', async () => {
  const app = await load('14 October 2026, 12:08 PM');
  app.DOMParser = class { parseFromString() { return { querySelectorAll: () => [], querySelector: () => null, body: { textContent: 'Log in', innerHTML: '<form>Log in</form>' } }; } };
  await app.restoreMessagesFromStorage([{ id: 'msg-1-8', title: 'saved', timestamp: Date.now() }]);
  assert.equal(app.allMessages.length, 1);
  assert.equal(app.allMessages[0].timestamp, null);
  assert.equal(app.allMessages[0].title, 'saved');
});
test('a partial course refresh replaces fetched mail and preserves the failed course', async () => {
  const app = await load('14 October 2026, 12:08 PM');
  app.allCourses = [{ id: 1, fullname: 'works' }, { id: 2, fullname: 'fails' }];
  app.fetch = async url => ({ ok: url.includes('c=1&'), status: 503, text: async () => '' });
  await app.restoreMessagesFromStorage([
    { id: 'msg-1-8', title: 'removed from successful course', timestamp: Date.now() },
    { id: 'msg-2-9', title: 'saved from failed course', timestamp: Date.now(), isRead: false }
  ]);
  assert.equal(app.allMessages.length, 2);
  assert.equal(app.allMessages.find(item => item.id === 'msg-1-1').timestampVersion, 1);
  assert.equal(app.allMessages.find(item => item.id === 'msg-2-9').timestamp, null);
  assert.equal(app.allMessages.find(item => item.id === 'msg-2-9').isRead, false);
});

test('a confirmed empty mailbox removes the saved mail for that course', async () => {
  const app = await load('14 October 2026, 12:08 PM');
  app.DOMParser = class { parseFromString() { return { querySelectorAll: () => [], querySelector: selector => selector === '.mail_list' ? {} : null, body: { textContent: 'No messages', innerHTML: '<div class="mail_list"></div>' } }; } };
  await app.restoreMessagesFromStorage([{ id: 'msg-1-8', title: 'removed mail', timestamp: Date.now() }]);
  assert.equal(app.allMessages.length, 0);
});
