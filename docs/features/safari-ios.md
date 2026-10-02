# Safari on macOS and iPad

Depends on notification preferences and the previous feature branches. Includes macOS and iPad Xcode projects, shared-resource synchronization, Safari-managed individual downloads, read-only requests through a signed-in same-origin E3 tab, macOS native notifications, and pointer gestures for touch resizing/dragging/cancellation.

Run `python3 scripts/sync-safari.py` after editing the root extension. Both manifests and Xcode marketing versions follow the root manifest. This feature branch preserves upstream version 2.2.0; releases and store publishing are separate decisions.

## Validation

- `node --test tests/*.test.cjs`: 37 passed, covering locale preferences, digest caches, notification queues, native adapters, downloads, session relay restrictions/callbacks and source consistency.
- `node tests/touch-ui-smoke.cjs`: language/settings, all six tabs, touch tap/resize/drag/cancel and narrow split views passed in Chromium fixtures.
- Notification UI fixtures passed for English, Chinese, unsupported desktop APIs and simulated native messaging. The native fixture advertises its required permission explicitly.
- Unsigned Release builds succeeded for both macOS and generic iOS destinations using `CODE_SIGNING_ALLOWED=NO` and temporary DerivedData paths.

Real Safari/iPad login, download, translation, API summaries, OS notification banners/clicks, signing and store distribution require manual verification. See `safari/README.md` and `safari/ios/README.md` for setup. These fixture and build checks do not claim device verification.
