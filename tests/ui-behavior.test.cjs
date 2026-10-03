const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '../content.js'), 'utf8');
function functionSource(name) {
  const start = source.search(new RegExp(`(?:async )?function ${name}\\(`));
  assert.ok(start >= 0, `${name} exists`);
  return source.slice(start, source.indexOf('\n}', start) + 2);
}
const now = 1800000000000;
class FixedDate extends Date {
  constructor(...args) { super(...(args.length ? args : [now])); }
  static now() { return now; }
}
const text = value => value;
const template = (parts, ...values) => parts.reduce((s, part, i) => s + part + (values[i] ?? ''), '');
test('countdown keeps the original one-hour and 24-hour status boundaries', () => {
  const context = vm.createContext({ Date: FixedDate, uiText: text, ui: template });
  vm.runInContext(functionSource('formatCountdown'), context);
  const hour = 3600000;
  for (const [left, status] of [[-1, 'overdue'], [0, 'overdue'], [1, 'urgent'],
    [hour - 1, 'urgent'], [hour, 'warning'], [hour + 1, 'warning'],
    [24 * hour - 1, 'warning'], [24 * hour, 'normal'], [24 * hour + 1, 'normal'],
    [3 * 24 * hour, 'normal'], [7 * 24 * hour, 'normal']]) {
    assert.equal(context.formatCountdown(now + left).status, status, `${left} ms remaining`);
  }
});
for (const deliveryFails of [false, true]) {
  test(`new announcements request a desktop alert before saving the center entry${deliveryFails ? ', even when delivery fails' : ''}`, async () => {
    const events = [], data = {};
    const context = vm.createContext({ Date: FixedDate, console, uiText: text, ui: template,
      updateNotificationBadge: () => events.push('badge'),
      chrome: { runtime: { sendMessage(message) {
        events.push(message);
        return deliveryFails ? Promise.reject(new Error('fixture delivery failure')) : Promise.resolve({ success: true });
      } }, storage: { local: {
        get: async () => { events.push('read'); return data; },
        set: async update => { events.push('save'); Object.assign(data, update); }
      } } }
    });
    vm.runInContext(functionSource('notifyNewAnnouncement'), context);
    await context.notifyNewAnnouncement({ id: 'a1', courseName: '原文課程', title: '公告原文', url: 'https://e3.nycu.edu.tw/source' });
    assert.equal(events[0].action, 'showNotification');
    assert.ok(events[0].title.includes('原文課程'));
    assert.equal(events[0].message, '公告原文');
    assert.deepEqual(events.slice(1), ['read', 'save', 'badge']);
    assert.equal(data.notifications[0].title, '公告原文');
    assert.equal(data.notifications[0].url, 'https://e3.nycu.edu.tw/source');
    assert.equal(data.notifications[0].read, false);
  });
}
