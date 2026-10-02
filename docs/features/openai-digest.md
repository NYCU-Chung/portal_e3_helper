# OpenAI summaries and daily digest

Depends on the sidebar appearance change. Enable AI summaries in More → Settings, enter an OpenAI API key, select a model, and test the connection. This replaces the active Gemini summary workflow; the existing key is not converted into an OpenAI key. Translation uses Google Translate without an AI key.

Single-item summaries send the item's content to OpenAI through the background worker. Generate overview in Announcements summarizes today's latest 40 announcement/message titles, courses, senders, and timestamps. Only recognized source numbers become links. Successful output is cached locally for the day and interface language; failed regeneration preserves the previous successful output. Requests use the Responses endpoint with `store: false`.

## Validation

Run `node --test tests/*.test.cjs` and the fixture UI smoke test. Digest tests cover reopening, a fresh content-script session, date/language invalidation, asynchronous completion, and failed regeneration. For a live check, enter your own key, test the connection, generate a summary and an overview, and reopen the sidebar. Live API and authenticated E3 requests are not exercised by fixtures.
