# Notification preferences and deadline reminders

Depends on the English interface change. Open More → Settings → Configure notifications and reminders, or the extension options page. Choose desktop delivery, update notifications, immediate/daily delivery, a local daily time, and multiple deadline reminders. Save changes explicitly; Test notification checks browser delivery.

Desktop update delivery moves from the content-script announcement call to the background engine watching announcement/message storage. This is deliberately part of this notification feature; preceding UI/AI/language branches preserve the original call. Keeping both paths would duplicate desktop alerts.

The queue is processed every minute while the browser runs. Submitted, deleted, changed, and expired assignment deadlines are skipped. The first announcement/message snapshot is silent. Subsequent new items are queued, deduplicated, and retried on delivery failure. A disabled or unsupported desktop transport leaves sidebar alerts available. Safari's native notification transport is a separate follow-up.

## Validation

Storage, delivery, permission checks, settings navigation and notification clicks use callback-compatible APIs for Chrome 93–94 as well as Promise implementations.

Run `node --test tests/*.test.cjs` and `node tests/notification-ui-smoke.cjs` with Playwright available. Repeat the browser fixture with `UI_LANGUAGE=zh-TW` and `SAFARI_FIXTURE=1` to check Chinese and unsupported APIs. Add `CHROME_93_FIXTURE=1` to exercise callback-only APIs. Fixtures test dropdowns, timing, save status, responsive layout, deduplication, daily delivery, deadline cancellation, and retries. Actual OS banners and authenticated E3 updates require manual checks.

### Real extension integration

With Playwright's test Chromium installed (`npx playwright install chromium`), run `node tests/notification-extension-smoke.cjs`. Set `EXTENSION_PATH` to an extracted runtime ZIP to test the package itself. The fixture uses an isolated temporary profile and blocks external DNS; it exercises the actual MV3 worker, Chrome notifications/storage/alarms, options save/test, baseline/update delivery, daily queuing, browser restart persistence, delivery-error retry, deadline filtering and alarm dispatch. It clears its notifications and removes the profile afterward. OS banner visibility and physical notification clicks remain manual checks.

Queue capacity is 200 active alerts. A full queue reports an error rather than success; announcement/message updates retain source references and retry after space becomes available, including after worker restart. Assignment/grade sidebar entries remain available when desktop scheduling fails. Unrelated settings saves preserve queued due times, and changes to daily time or delivery mode reschedule only future alerts. Click destinations are saved before each notification is displayed.
