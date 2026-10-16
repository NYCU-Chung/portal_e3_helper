const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.join(__dirname, '..');
const content = fs.readFileSync(path.join(root, 'content.js'), 'utf8');
const sites = ['https://e3.nycu.edu.tw/*', 'https://e3p.nycu.edu.tw/*', 'https://timetable.nycu.edu.tw/*', 'https://portal.nycu.edu.tw/*'];

test('Chrome injects only into the HTTPS E3, timetable and portal hosts by default', () => {
  for (const directory of ['']) {
    const manifest = JSON.parse(fs.readFileSync(path.join(root, directory, 'manifest.json')));
    assert.deepEqual(manifest.content_scripts.flatMap(script => script.matches), sites);
    assert.ok(!manifest.host_permissions.includes('<all_urls>'));
  }
});

test('unrelated pages exit before touching storage, DOM, listeners or helper dependencies', () => {
  for (const url of ['https://example.com/', 'https://www.nycu.edu.tw/',
    'http://timetable.nycu.edu.tw/', 'https://portal.nycu.edu.tw.evil.test/', 'https://timetable.nycu.edu.tw.evil.test/', 'http://e3p.nycu.edu.tw/',
    'https://evile3p.nycu.edu.tw/', 'https://e3p.nycu.edu.tw.evil.test/',
    'https://e3p.nycu.edu.tw@evil.test/', 'about:blank', 'file:///demo/index.html']) {
    vm.runInNewContext(content, { location: new URL(url) }, { timeout: 1000 });
  }
});

test('E3, timetable and portal pages still reach initialization', () => {
  for (const url of ['https://e3.nycu.edu.tw/', 'https://e3p.nycu.edu.tw/', 'https://timetable.nycu.edu.tw/', 'https://portal.nycu.edu.tw/']) {
    const sentinel = new Error('initialization reached');
    const context = { location: new URL(url) };
    context.E3HelperI18n = { get text() { throw sentinel; } };
    assert.throws(() => vm.runInNewContext(content, context), error => error === sentinel);
  }
});
