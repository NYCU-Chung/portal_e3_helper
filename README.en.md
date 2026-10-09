# NYCU E3 Helper

[繁體中文](README.md) · English

Bring NYCU E3 assignments, courses, announcements, messages and materials together in one sidebar.

[Chrome Web Store](https://chromewebstore.google.com/detail/nycu-e3-helper/cmagonljljocpkfojkabhiedjafamoef) · [E3](https://e3p.nycu.edu.tw/) · [Changelog (Chinese)](CHANGELOG.md)

This guide describes the source in this repository. The store release may not include the latest changes.

## Features

- A resizable sidebar with a draggable floating button, with a light theme by default and optional dark or follow-the-system themes.
- Assignment countdowns, submission status and custom assignments.
- Course lists, membership statistics, membership changes and grades.
- Course announcements and system messages, with filters, read status and full previews.
- A daily AI digest of today's announcements and messages, with source links.
- Chinese–English translation with Google Translate and AI summaries with selectable Gemini or OpenAI.
- Material scanning and individual or ZIP downloads.
- Notifications, unread badges and background synchronization.
- Traditional Chinese and English interface languages.

## Installation

Install from the [Chrome Web Store](https://chromewebstore.google.com/detail/nycu-e3-helper/cmagonljljocpkfojkabhiedjafamoef), or load the source in Chrome, Edge or Brave:

1. Clone this repository or download its source.
2. Open `chrome://extensions/`, `edge://extensions/` or `brave://extensions/`.
3. Enable **Developer mode**, select **Load unpacked**, and choose the folder containing `manifest.json`.
4. Open a webpage and click the floating **E3 Helper** button.

After updating the source, reload the extension and refresh open webpages. Browser internal pages and some restricted pages do not allow the sidebar.

## Getting started and language

1. Sign in to [E3](https://e3p.nycu.edu.tw/).
2. Open the sidebar and click **Sync** to load assignments and courses.
3. Use **Assignments**, **Courses**, **Downloads**, **Announcements**, **Alerts** and **Help** to navigate. The tabs wrap onto additional rows when the sidebar is narrow, keeping every label visible.

The initial interface language follows your browser: Chinese uses Traditional Chinese; other languages use English. To change it, open **More (⋯) → Settings → Interface language**, select **English** or **繁體中文**, and click **Save settings**. Saving a new language refreshes the current page. Refresh other open pages manually.

Your language choice is stored locally. Course names, assignment names and original announcement/message content keep their source text. New AI summaries and daily digests use the selected interface language. Previously saved notifications and release notes may retain the language in which they were created.

Appearance is light by default. To change it, open **More (⋯) → Settings → Appearance → Theme**, choose **Dark** or **Follow system**, and click **Save settings**; the change applies immediately. Announcement and message previews stay on a light background in the dark theme so their original colors remain readable.

## Daily use

### Assignments

Click a card to open its assignment page. Use **Mark as submitted** to change the status, or **Submitted** to mark it as pending again. Use **Add assignment**, **Edit** and **Delete** to manage custom assignments.

Deadlines use your local time zone. Countdowns are red with less than one hour remaining, amber with less than 24 hours remaining, normal otherwise, and gray after the deadline. Submitted assignments disappear after their deadline. Deleted E3 assignments can reappear after syncing.

### Announcements and messages

Open **Announcements** to load data, or click **Reload**. Filter by type and read status. A red dot indicates an unread item. **View content** opens a preview with read-status controls and **Open full page**.

Use **Chinese → English**, **English → Chinese**, **AI summary** or **Show original** in the preview. Translation uses Google Translate and requires no OpenAI API key.

### Today's digest

Load announcements and messages, then click **Generate digest** at the top of the Announcements tab. Enable AI summaries, select Gemini or OpenAI, and save the selected provider’s API key and model in Settings first.

The digest uses today's local date and up to 40 recent items, regardless of filters. It summarizes titles, courses, senders and timestamps, rather than reading each item's full body. Open the source links for full instructions and deadlines. Retry if generation fails.

### Courses and downloads

Sync courses, select a course in **Courses**, and open **Statistics** or **Grades**. Loading data requires an active E3 login.

In **Downloads**, use **Scan this page** on an E3 page, or **Select courses** to scan selected courses. Select the files, then choose **Download separately** or **Download ZIP**. Documents, videos and archives are supported. For large files, use smaller batches or download from the source page.

### Notifications and synchronization

**Alerts** collects assignment, grading, announcement, membership and version notifications. Assignments due within 24 hours also contribute to reminders.

Background sync runs every 60 minutes for assignments, courses and membership checks, and every 30 minutes for announcements and messages. The open sidebar also checks for stale assignment/course data. Manual **Sync** updates assignments and courses; use **Reload** in Announcements for announcements and messages.

The helper can display saved data and load E3 data in the background from ordinary websites. Scheduled announcement and message synchronization only uses existing E3 tabs and skips when none are open. Manual loading may briefly create an inactive E3 tab and close it afterward. Scanning the current page requires opening the relevant E3 page.

## AI settings and privacy

Open **More → Settings**, enable AI summaries, select **Gemini** or **OpenAI**, enter that provider’s API key, choose a model and click **Save settings**. Both providers’ settings are retained when you switch. Gemini models load from the API when Gemini is active; you can reload the list or enter a model ID manually. Select a model explicitly before generating a summary. **Test connection** checks the selected provider. Account access, pricing and quotas depend on your provider account.

Course data, read status, settings and API keys are stored locally in your browser. Translation sends content to Google Translate. Individual summaries send the item body to the selected Gemini or OpenAI provider; daily digests send today's titles, courses, senders and timestamps. OpenAI summary requests use the Responses API with `store: false`. The extension uses your existing E3 login and does not ask for your E3 password.

## Troubleshooting

- **Sync or loading fails:** sign in to E3 again, check your connection and retry.
- **AI summary fails:** check that summaries are enabled and your API key is saved; test the connection in Settings.
- **Translation fails:** check your connection and retry. Translation does not require an AI provider key.
- **The interface is outdated after an update:** reload the extension, then refresh open webpages.
- **ZIP download fails:** retry with fewer files or use individual downloads.
- **Need to report a problem:** use **More → View logs** or **More → Report an issue**. Review logs for personal information before sharing.

## Development

`content.js` implements the sidebar and E3 interaction. `background.js` handles downloads, background synchronization and API calls. `i18n.js` contains the shared interface catalog, language preference and static-template translation. `_locales/` supplies localized extension descriptions.

Translations apply to source-authored text before dynamic values are interpolated, preserving course names, URLs and original content. Keep E3 scraping selectors and parsing rules separate from translated UI text.

Run the localization checks with:

```bash
node --test tests/*.test.cjs
node --check i18n.js
node --check content.js
node --check background.js
```

Optional browser smoke test (requires Playwright and a browser):

```bash
node tests/ui-smoke.cjs
node tests/language-ui-smoke.cjs
node tests/translation-html-smoke.cjs
```

Set `CHROME_PATH` to use an existing Chrome executable, or `PLAYWRIGHT_PATH` to use a Playwright installation outside this project. The smoke test uses local fixtures and does not contact E3 or an AI service.
