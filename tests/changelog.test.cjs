const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const root = path.join(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'background.js'), 'utf8');
const start = source.indexOf('function parseChangelogForVersion');
const parse = vm.runInNewContext(`(${source.slice(start, source.indexOf('\n}\n', start) + 2)})`);
const changelog = fs.readFileSync(path.join(root, 'CHANGELOG.md'), 'utf8');

test("What's New shows the release notes of the current version", () => {
  const { version } = JSON.parse(fs.readFileSync(path.join(root, 'manifest.json'), 'utf8'));
  const html = parse(changelog, version);
  assert.match(html, /<h4/);
  assert.match(html, /<li/);
});
test('a version block stops before the next version and handles the last block', () => {
  const notes = '## [2.0.0] - b\n\n### Added\n- **new** thing\n\n---\n\n## [1.0.0] - a\n\n- first release';
  assert.ok(parse(notes, '2.0.0').includes('<strong>new</strong> thing'));
  assert.ok(!parse(notes, '2.0.0').includes('first release'));
  assert.ok(parse(notes, '1.0.0').includes('first release'));
  assert.equal(parse(notes, '9.9.9'), '');
});
