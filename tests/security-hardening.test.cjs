const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const root = path.join(__dirname, '..');
const background = fs.readFileSync(path.join(root, 'background.js'), 'utf8');
const content = fs.readFileSync(path.join(root, 'content.js'), 'utf8');

function sliceBetween(source, start, end) {
  const a = source.indexOf(start);
  const b = source.indexOf(end, a + start.length);
  assert.notEqual(a, -1, `missing start marker: ${start}`);
  assert.notEqual(b, -1, `missing end marker: ${end}`);
  return source.slice(a, b);
}

test('E3 URL validation requires exact HTTPS hosts', () => {
  const code = sliceBetween(background, 'const E3_ALLOWED_HOSTS', '// NYCU E3 Helper - Background Script');
  const context = { URL, Set };
  vm.createContext(context);
  vm.runInContext(`${code}\nthis.parseAllowedE3Url = parseAllowedE3Url;`, context);
  assert.equal(context.parseAllowedE3Url('https://e3.nycu.edu.tw/x').hostname, 'e3.nycu.edu.tw');
  assert.equal(context.parseAllowedE3Url('https://e3p.nycu.edu.tw/x').hostname, 'e3p.nycu.edu.tw');
  assert.throws(() => context.parseAllowedE3Url('http://e3.nycu.edu.tw/x'));
  assert.throws(() => context.parseAllowedE3Url('https://attacker-e3.nycu.edu.tw/x'));
  assert.throws(() => context.parseAllowedE3Url('https://e3.nycu.edu.tw.evil.invalid/x'));
  assert.throws(() => context.parseAllowedE3Url('https://e3.nycu.edu.tw@evil.invalid/x'));
});

test('rich-content URL policy blocks executable and unsafe data schemes', () => {
  const code = sliceBetween(content, 'function getRichContentUrlScheme', '// HTML 白名單清理');
  const context = {};
  vm.createContext(context);
  vm.runInContext(`${code}\nthis.safe = isSafeRichContentUrl;`, context);
  assert.equal(context.safe('https://e3.nycu.edu.tw/mod/forum/view.php', 'href'), true);
  assert.equal(context.safe('/mod/forum/view.php', 'href'), true);
  assert.equal(context.safe('mailto:user@example.com', 'href'), true);
  assert.equal(context.safe('java\nscript:alert(1)', 'href'), false);
  assert.equal(context.safe(' data:text/html,<script>alert(1)</script>', 'href'), false);
  assert.equal(context.safe('data:image/png;base64,AAAA', 'src'), true);
  assert.equal(context.safe('data:image/svg+xml,<svg/>', 'src'), false);
});

test('hardening is wired into downloads, redirects, links, and manifest permissions', () => {
  assert.ok(background.includes('parseAllowedE3Url(request.url);'));
  assert.ok(background.includes('parseAllowedE3Url(response.url);'));
  assert.ok(background.includes("/[\\\\/]/.test(request.filename)"));
  assert.ok(content.includes("rel.add('noopener')"));
  assert.ok(content.includes("rel.add('noreferrer')"));
  const manifest = JSON.parse(fs.readFileSync(path.join(root, 'manifest.json'), 'utf8'));
  assert.equal(manifest.minimum_chrome_version, '93');
  assert.equal(manifest.version, '2.3.0');
  assert.equal(manifest.host_permissions.includes('http://localhost:11434/*'), false);
});
