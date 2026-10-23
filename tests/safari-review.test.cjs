const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { execFileSync } = require('node:child_process');
const root = path.join(__dirname, '..');
const content = fs.readFileSync(path.join(root, 'content.js'), 'utf8');
const background = fs.readFileSync(path.join(root, 'background.js'), 'utf8');

function downloadApp(permissions = []) {
  const nodes = new Map(), messages = [], opened = [], timers = [];
  function element(tag) {
    return { tag, children: [], style: {}, disabled: false, textContent: '',
      appendChild(child) { this.children.push(child); if (child.id) nodes.set(child.id, child); },
      replaceChildren() { this.children = []; },
      insertAdjacentElement(_, child) { nodes.set(child.id, child); },
      click() { opened.push(this.href); }, remove() {}, focus() {} };
  }
  for (const id of ['e3-helper-download-separate', '.e3-helper-download-status', '.e3-helper-progress-container', '.e3-helper-progress-fill', '.e3-helper-progress-text']) nodes.set(id, element('div'));
  const app = { URL, console: { log() {}, error() {} },
    selectedPDFs: new Set([0, 1]),
    allPDFs: [
      { url: 'https://e3p.nycu.edu.tw/pluginfile.php/1/a.pdf', filename: 'a.pdf', course: 'Course' },
      { url: 'https://e3.nycu.edu.tw/pluginfile.php/2/b.pdf', filename: '<b>.pdf', course: 'Course' }
    ],
    SUPPORTED_FILE_TYPES: [{ ext: '.pdf' }], sanitizeFilename: value => value,
    uiText: value => value, ui: (strings, ...values) => String.raw({ raw: strings }, ...values), showTemporaryMessage() {},
    setTimeout(fn, delay) { timers.push(delay); if (delay === 200) fn(); },
    window: { open: url => opened.push(url) },
    document: { createElement: element, getElementById: id => nodes.get(id), querySelector: id => nodes.get(id), body: { appendChild() {} } },
    chrome: { runtime: { getManifest: () => ({ permissions }), sendMessage: async msg => { messages.push(msg); return { success: true }; } } }
  };
  vm.createContext(app);
  vm.runInContext(content.slice(content.indexOf('async function downloadE3File'), content.indexOf('// 綁定下載按鈕事件')), app);
  vm.runInContext(content.slice(content.indexOf('async function downloadSeparately()'), content.indexOf('// 批量下載選取的檔案')), app);
  return { app, nodes, messages, opened, timers };
}

test('Safari batch action exposes each selected file as a real link without opening tabs or claiming completion', async () => {
  const { app, nodes, messages, opened, timers } = downloadApp();
  await app.downloadSeparately();
  assert.equal(opened.length, 0, 'batch must not consume a user gesture by opening tabs');
  assert.equal(messages.length, 0);
  assert.equal(timers.length, 0);
  const list = nodes.get('e3-helper-safari-download-links');
  assert.ok(list, 'selected files need explicit user-clickable links');
  const links = list.children.filter(child => child.tag === 'a');
  assert.equal(links.length, 2);
  assert.equal(links[1].textContent, 'Course — <b>.pdf');
  for (const link of links) {
    assert.equal(link.target, '_blank');
    assert.equal(link.rel, 'noopener noreferrer');
  }
  assert.doesNotMatch(nodes.get('.e3-helper-download-status').textContent, /下載完成/);
  assert.equal(nodes.get('e3-helper-download-separate').disabled, false);
  links[1].click();
  assert.deepEqual(opened, [app.allPDFs[1].url]);
  app.selectedPDFs = new Set([0]);
  await app.downloadSeparately();
  assert.equal(list.children.filter(child => child.tag === 'a').length, 1, 'reopening replaces the previous selection');
});

test('Safari batch links reject unsafe URLs and retain explicit external video links', async () => {
  const { app, nodes } = downloadApp();
  app.allPDFs = [
    { url: 'javascript:alert(1)', filename: 'bad', course: 'Course' },
    { url: 'https://www.youtube.com/embed/fixture', filename: 'Video', course: 'Course', isIframe: true }
  ];
  await app.downloadSeparately();
  const links = nodes.get('e3-helper-safari-download-links').children.filter(child => child.tag === 'a');
  assert.equal(links.length, 1);
  assert.equal(links[0].href, app.allPDFs[1].url);
});

test('Chrome batch action retains the background downloads API', async () => {
  const { app, nodes, messages, opened } = downloadApp(['downloads']);
  await app.downloadSeparately();
  assert.equal(messages.length, 2);
  assert.ok(messages.every(message => message.action === 'download'));
  assert.equal(opened.length, 0);
  assert.equal(nodes.has('e3-helper-safari-download-links'), false);
});

test('background log broadcast tolerates callback-only messaging and consumes delivery errors', async () => {
  let consumed = 0, sent = 0;
  const runtime = { get lastError() { consumed++; return { message: 'No receiver' }; } };
  const app = { console: { log() {}, info() {}, warn() {}, error() {}, debug() {} },
    chrome: { runtime, storage: { local: { get(_, cb) { cb({}); }, set(_, cb) { cb?.(); } } },
      tabs: { query(_, cb) { cb([{ id: 1 }, { id: 2 }]); },
        sendMessage(_, message, callback) { sent++; callback?.(); return undefined; } } }
  };
  vm.createContext(app);
  vm.runInContext(background.slice(background.indexOf('const originalConsole'), background.indexOf('// 攔截 console 方法')), app);
  assert.doesNotThrow(() => app.sendLogToContentScript('log', ['fixture']));
  await new Promise(setImmediate);
  assert.equal(sent, 2);
  assert.ok(consumed >= 2);
});

test('Safari sync copies release notes into a fresh resource directory', () => {
  // Run the real sync function against a temporary project, so an omitted FILES
  // entry cannot be hidden by an older checked-in copy of CHANGELOG.md.
  execFileSync('python3', ['-c', `
import pathlib, plistlib, runpy, shutil, tempfile
source = pathlib.Path.cwd()
sync = runpy.run_path(str(source / 'scripts/sync-safari.py'))
with tempfile.TemporaryDirectory() as directory:
    root = pathlib.Path(directory)
    for name in set(sync['FILES']) | {'manifest.json', 'CHANGELOG.md'}:
        src, dest = source / name, root / name
        if src.is_dir(): shutil.copytree(src, dest)
        else: shutil.copy2(src, dest)
    project = root / 'Safari'
    (project / 'Extension/Resources').mkdir(parents=True)
    pbx = project / 'Safari.xcodeproj/project.pbxproj'
    pbx.parent.mkdir()
    pbx.write_bytes(plistlib.dumps({'objects': {}}))
    sync['sync_project'].__globals__['ROOT'] = root
    sync['sync_project'](project, 'Extension', False)
    assert (project / 'Extension/Resources/CHANGELOG.md').read_bytes() == (source / 'CHANGELOG.md').read_bytes()
`], { cwd: root, encoding: 'utf8' });
});

for (const [project, extension] of [
  ['safari/NYCU E3 Helper', 'NYCU E3 Helper Extension'],
  ['safari/ios/NYCU E3 Helper iOS', 'NYCU E3 Helper iOS Extension']
]) {
  test(`${extension} bundles current release notes in its actual Xcode resource phase`, t => {
    const dir = path.join(root, project);
    assert.equal(fs.readFileSync(path.join(dir, extension, 'Resources/CHANGELOG.md'), 'utf8'), fs.readFileSync(path.join(root, 'CHANGELOG.md'), 'utf8'));
    if (process.platform !== 'darwin') {
      t.diagnostic('Bundle source checked; Xcode resource graph verification requires macOS plutil.');
      return;
    }
    const pbx = path.join(dir, fs.readdirSync(dir).find(name => name.endsWith('.xcodeproj')), 'project.pbxproj');
    const { objects } = JSON.parse(execFileSync('plutil', ['-convert', 'json', '-o', '-', pbx], { encoding: 'utf8' }));
    const target = Object.values(objects).find(obj => obj.isa === 'PBXNativeTarget' && obj.name === extension);
    const phase = target.buildPhases.map(id => objects[id]).find(obj => obj.isa === 'PBXResourcesBuildPhase');
    assert.ok(phase.files.some(id => objects[objects[id].fileRef]?.path === 'Resources/CHANGELOG.md'));
  });
}

test('starting any new download scan removes prepared links from the previous selection', async () => {
  for (const name of ['scanSelectedCourses', 'scanCurrentPage', 'scanAllCourses']) {
    const { app, nodes } = downloadApp();
    await app.downloadSeparately();
    nodes.get('e3-helper-safari-download-links').remove = () => nodes.delete('e3-helper-safari-download-links');
    nodes.set('.e3-helper-pdf-list', {});
    app.allCourses = [];
    app.selectedCourses = new Set();
    app.loadCourseList = async () => {};
    // Stop current-page scanning immediately after its reset, before fetching any results.
    app.getCurrentCourseName = () => { throw new Error('scan paused'); };
    const start = content.indexOf('async function ' + name + '(');
    vm.runInContext(content.slice(start, content.indexOf('\n}\n', start) + 2), app);
    await app[name]().catch(error => { assert.equal(error.message, 'scan paused'); });
    assert.equal(nodes.has('e3-helper-safari-download-links'), false, name);
  }
});

test('Safari exam module is synced and included in both Xcode resource graphs', t => {
  const file = 'exam-detection.js';
  const rootModule = fs.readFileSync(path.join(root, file), 'utf8');
  assert.match(fs.readFileSync(path.join(root, 'scripts/sync-safari.py'), 'utf8'), /'exam-detection\.js'/);
  const rootManifest = JSON.parse(fs.readFileSync(path.join(root, 'manifest.json'), 'utf8'));
  assert.ok(rootManifest.content_scripts.some(script => script.js.includes(file)));
  for (const [project, extension] of [
    ['safari/NYCU E3 Helper', 'NYCU E3 Helper Extension'],
    ['safari/ios/NYCU E3 Helper iOS', 'NYCU E3 Helper iOS Extension']
  ]) {
    const resources = path.join(root, project, extension, 'Resources');
    assert.equal(fs.readFileSync(path.join(resources, file), 'utf8'), rootModule);
    const manifest = JSON.parse(fs.readFileSync(path.join(resources, 'manifest.json'), 'utf8'));
    assert.ok(manifest.content_scripts.some(script => script.js.includes(file)));
    assert.match(fs.readFileSync(path.join(resources, 'background.js'), 'utf8'), /importScripts\([^)]*exam-detection\.js/);
    if (process.platform !== 'darwin') continue;
    const projectPath = path.join(root, project);
    const pbx = path.join(projectPath, fs.readdirSync(projectPath).find(name => name.endsWith('.xcodeproj')), 'project.pbxproj');
    const { objects } = JSON.parse(execFileSync('plutil', ['-convert', 'json', '-o', '-', pbx], { encoding: 'utf8' }));
    const target = Object.values(objects).find(obj => obj.isa === 'PBXNativeTarget' && obj.name === extension);
    assert.ok(target, 'extension target must exist');
    const phase = target.buildPhases.map(id => objects[id]).find(obj => obj.isa === 'PBXResourcesBuildPhase');
    assert.ok(phase.files.some(id => objects[objects[id].fileRef]?.path === 'Resources/exam-detection.js'));
  }
});
