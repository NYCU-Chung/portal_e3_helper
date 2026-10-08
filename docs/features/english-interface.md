# English interface

Depends on the sidebar appearance and OpenAI summary changes. The initial language follows the browser, with Traditional Chinese for Chinese locales and English otherwise. More → Settings → Interface language saves an override and reloads the current page; refresh other open pages manually.

Static labels, templates, settings, background messages, and new AI output follow the selected language. Dynamic course names, assignments, original messages, URLs, and scraping selectors remain unchanged. Existing saved notifications can retain their previous language.

## Validation

- `node --test tests/*.test.cjs` covers browser/saved defaults, interpolation, persistence, invalid preferences, worker updates, manifests, and digest language invalidation.
- `node tests/language-ui-smoke.cjs` checks six tabs for untranslated labels, settings, save/reload, original source text, and 280–800 px navigation.
- Manually switch between English and Traditional Chinese and verify both open pages and newly generated summaries.
