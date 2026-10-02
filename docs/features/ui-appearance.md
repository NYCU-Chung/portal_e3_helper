# Sidebar appearance

The sidebar defaults to a light theme. Choose Light, Dark, or System under More → Settings → Appearance. The preference is saved locally and follows changes in other open helper tabs.

The header, six navigation tabs, cards, icons, menus, and help content share theme tokens. The navigation fits sidebar widths from 280 to 800 px. The existing Gemini translation and summary workflows remain available.

## Validation

- `node --check content.js` and `node --check background.js`
- `node --test tests/changelog.test.cjs`
- `node tests/ui-smoke.cjs` with Playwright available (optionally set `PLAYWRIGHT_PATH` and `CHROME_PATH`). Uses fixture data without E3 login or live API calls.
- Manually load the unpacked extension, check all six tabs, change the theme, reopen Settings, and verify the persisted preference. Check What's New renders the current release notes.
