# Exam suggestions

Detected exams stay in suggestions until the user accepts them. Acceptance runs in the background against the latest saved coursework; the exam and the resolved suggestion are written together. Tabs send a selected candidate and its date/time fingerprint, so an outdated card cannot silently accept a revised date. Bulk acceptance uses the same operation.

Confirmed exams ask to update when a preferred source changes the date or known start time, including a previously unknown time becoming known. A source losing its time does not replace a confirmed exact time. Quiz revisions retain the original suggestion identity when their source carries the same quiz number, or when that source has exactly one removed and one added occurrence. Distinct unchanged quizzes remain separate; ambiguous revisions are not guessed.

Announcement and mail extraction caches expire after one day, invalidate when source metadata changes, and refresh on an explicit scan. Failed or empty body reads remain retryable and do not remove earlier suggestions. Converting submitted manual coursework into an exam resets both its inline and saved submission status, since exams have no submission toggle.

Run `node --test tests/exam-detection.test.cjs` for deterministic extraction, revision, cache, status, and multi-tab regressions. With Playwright Chromium available, run `node tests/exam-extension-smoke.cjs` for actual MV3 storage/messages and two E3 fixture tabs. The browser fixture blocks external DNS and uses local pages; it does not authenticate to E3 or call an AI provider.
