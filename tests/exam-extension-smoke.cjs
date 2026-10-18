const { chromium } = require(process.env.PLAYWRIGHT_PATH || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
(async () => {
  const root = process.env.EXTENSION_PATH || path.resolve(__dirname, '..');
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'e3-exam-regression-'));
  let context;
  try {
    context = await chromium.launchPersistentContext(profile, {
      headless: true, channel: 'chromium',
      args: [`--disable-extensions-except=${root}`, `--load-extension=${root}`, '--host-resolver-rules=MAP * ~NOTFOUND']
    });
    await context.route('https://**/*', route => route.fulfill({ contentType: 'text/html', body: '<!doctype html><html><head></head><body>Offline exam fixture</body></html>' }));
    const worker = context.serviceWorkers()[0] || await context.waitForEvent('serviceworker');
    await worker.evaluate(async () => {
      const deadline = Date.now() + 15 * 86400000;
      const candidates = ['A', 'B'].map(courseId => ({ courseId, course: courseId, examType: 'final', deadline, timeKnown: 'exact', source: 'syllabusFile', url: `https://e3p.nycu.edu.tw/file/${courseId}`, confidence: 'high' }));
      await chrome.storage.local.set({ interfaceLanguage: 'en', lastSeenVersion: chrome.runtime.getManifest().version, lastSyncTime: Date.now(), courses: [], announcements: [], messages: [],
        notificationSettings: { desktop: false, reminders: [] }, assignmentStatuses: { manual: 'submitted' },
        assignments: [{ eventId: 'manual', name: 'Submitted manual task', course: 'Manual course', deadline, isManual: true, manualStatus: 'submitted', url: '#' }],
        examSuggestions: E3ExamDetection.mergeSuggestions([], candidates) });
    });
    const pages = await Promise.all([context.newPage(), context.newPage()]);
    for (const page of pages) {
      await page.goto('https://e3p.nycu.edu.tw/');
      await page.locator('.e3-helper-sidebar-toggle').click();
      await page.locator('[data-exam-suggestion="exam:A:final"]').waitFor();
    }
    const waitForStorage = async predicate => {
      for (let attempt = 0; attempt < 100; attempt++) {
        const data = await worker.evaluate(() => chrome.storage.local.get(['assignments', 'assignmentStatuses', 'examSuggestions']));
        if (predicate(data)) return data;
        await pages[0].waitForTimeout(50);
      }
      throw new Error('Exam storage did not reach the expected state');
    };
    await worker.evaluate(async () => {
      const { assignments } = await chrome.storage.local.get('assignments');
      await chrome.storage.local.set({ assignments: [...assignments, { eventId: 'synced', name: 'Newly synced task', deadline: Date.now() + 86400000 }] });
    });
    await pages[0].locator('[data-exam-suggestion="exam:A:final"] [data-exam-action="accept"]').click();
    await waitForStorage(data => data.assignments.some(item => item.examKey === 'exam:A:final'));
    await pages[1].locator('[data-exam-suggestion="exam:B:final"] [data-exam-action="accept"]').click();
    const accepted = await waitForStorage(data => data.assignments.some(item => item.examKey === 'exam:B:final'));
    assert.equal(accepted.assignments.length, 4);
    assert.ok(accepted.assignments.some(item => item.eventId === 'synced'));
    assert.equal(accepted.assignments.filter(item => item.kind === 'exam').length, 2);
    const original = accepted.assignments.find(item => item.examKey === 'exam:A:final');
    const updatedTime = original.deadline + 3600000;
    await worker.evaluate(async ({ deadline }) => {
      await mergeExamCandidates([{ courseId: 'A', course: 'A', examType: 'final', deadline, timeKnown: 'exact', source: 'syllabusFile', url: 'https://e3p.nycu.edu.tw/file/A', confidence: 'high' }]);
    }, { deadline: updatedTime });
    const updateCard = pages[0].locator('[data-exam-suggestion="exam:A:final"]');
    await updateCard.getByRole('button', { name: 'Update date', exact: true }).click();
    const updated = await waitForStorage(data => data.assignments.some(item => item.eventId === original.eventId && item.deadline === updatedTime));
    assert.equal(updated.assignments.length, 4);
    await pages[0].locator('.e3-helper-edit-assignment[data-event-id="manual"]').click();
    await pages[0].locator('#e3-helper-assignment-kind').selectOption('exam');
    await pages[0].locator('#e3-helper-exam-type').selectOption('midterm');
    await pages[0].locator('#e3-helper-add-assignment-form button[type="submit"]').click();
    const converted = await waitForStorage(data => data.assignments.find(item => item.eventId === 'manual')?.kind === 'exam');
    assert.equal(converted.assignments.find(item => item.eventId === 'manual').manualStatus, 'pending');
    assert.equal(converted.assignmentStatuses.manual, 'pending');
    for (const timeKnown of ['class-slot', 'date-only']) {
      const expected = await worker.evaluate(async timeKnown => {
        const data = await chrome.storage.local.get('assignments');
        const exam = data.assignments.find(item => item.eventId === 'manual');
        exam.deadline = Math.floor(exam.deadline / 60000) * 60000;
        exam.timeKnown = timeKnown;
        exam.endTime = exam.deadline + 3 * 3600000;
        await chrome.storage.local.set({ assignments: data.assignments });
        return exam;
      }, timeKnown);
      await pages[0].reload();
      await pages[0].locator('.e3-helper-sidebar-toggle').click();
      await pages[0].locator('[data-tab="assignments"]').click();
      await pages[0].locator('.e3-helper-edit-assignment[data-event-id="manual"]').click();
      const newName = 'Renamed ' + timeKnown;
      await pages[0].locator('#e3-helper-assignment-name').fill(newName);
      await pages[0].locator('#e3-helper-add-assignment-form button[type="submit"]').click();
      const renamed = (await waitForStorage(data => data.assignments.some(item => item.eventId === 'manual' && item.name === newName))).assignments.find(item => item.eventId === 'manual');
      assert.equal(renamed.timeKnown, expected.timeKnown);
      assert.equal(renamed.endTime, expected.endTime);
      assert.equal(renamed.deadline, expected.deadline);
    }
    console.log('Passed: exam rename preserves date-only and class-slot timing metadata.');
    console.log('Passed: real MV3 two-tab acceptance preserves saved coursework, corrected start times update in place, and converting submitted coursework resets both statuses.');
  } finally {
    await context?.close();
    fs.rmSync(profile, { recursive: true, force: true });
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
