// Persist desktop notification queues across MV3 worker restarts.
(function () {
  const defaults = { desktop: true, mode: 'instant', time: '20:00', reminders: [24, 1], updates: true };
  function normalize(value = {}) {
    return { updates: value.updates ?? defaults.updates, desktop: E3DesktopNotifications.supported && (value.desktop ?? defaults.desktop), mode: value.mode === 'daily' ? 'daily' : 'instant',
      time: /^([01]\d|2[0-3]):[0-5]\d$/.test(value.time || '') ? value.time : defaults.time,
      reminders: (Array.isArray(value.reminders) ? value.reminders : defaults.reminders).filter(h => Number.isFinite(h) && h > 0 && h <= 720) };
  }
  function dueAt(settings, now) {
    if (settings.mode !== 'daily') return now;
    const date = new Date(now);
    const [hours, minutes] = settings.time.split(':').map(Number);
    date.setHours(hours, minutes, 0, 0);
    if (date.getTime() <= now) date.setDate(date.getDate() + 1);
    return date.getTime();
  }
  let work = Promise.resolve();
  function serial(fn) {
    const result = work.then(fn);
    work = result.catch(() => {});
    return result;
  }
  async function add(id, options, url, now = Date.now(), deadline = null) {
    await E3HelperI18n.ready;
    const data = await E3NotificationAPI.get(['notificationSettings', 'notificationQueue', 'notificationHistory']);
    const settings = normalize(data.notificationSettings);
    if (!settings.desktop) return;
    const queue = data.notificationQueue || [];
    if ((data.notificationHistory || []).includes(id) || queue.some(item => item.id === id)) return;
    if (queue.length >= 200) {
      await E3NotificationAPI.set({ notificationDeliveryError: '通知佇列已滿，請檢查桌面通知設定。' });
      return;
    }
    queue.push({ id, options, url, deadline, created: now, due: dueAt(settings, now), delivered: {}, attempts: 0 });
    await E3NotificationAPI.set({ notificationQueue: queue });
  }
  async function process(now = Date.now()) {
    await E3HelperI18n.ready;
    let data = await E3NotificationAPI.get(['assignments', 'assignmentStatuses', 'notificationSettings']);
    const settings = normalize(data.notificationSettings);
    // Strip retired channel credentials while retaining timing and reminder preferences.
    if (data.notificationSettings && Object.keys(data.notificationSettings).some(key => !(key in settings))) {
      await E3NotificationAPI.set({ notificationSettings: settings });
    }
    for (const assignment of data.assignments || []) {
      if (assignment.manualStatus === 'submitted' || data.assignmentStatuses?.[assignment.eventId] === 'submitted') continue;
      const left = Number(assignment.deadline) - now;
      if (!Number.isFinite(left) || left <= 0) continue;
      for (const hours of settings.reminders) {
        if (left > hours * 3600000) continue;
        await add(`deadline-${assignment.eventId}-${assignment.deadline}-${hours}`, {
          type: 'basic', iconUrl: chrome.runtime.getURL('128.png'),
          title: E3HelperI18n.text('作業即將截止'),
          message: `${assignment.name}\n${assignment.course || ''}\n${new Date(assignment.deadline).toLocaleString(E3HelperI18n.language)}`
        }, assignment.url, now, { eventId: assignment.eventId, timestamp: Number(assignment.deadline), hours });
      }
    }
    data = await E3NotificationAPI.get(['notificationQueue', 'notificationHistory', 'notificationLinks']);
    const history = data.notificationHistory || [];
    const links = data.notificationLinks || {};
    const remaining = [];
    let deliveryError = '';
    const assignments = await E3NotificationAPI.get(['assignments', 'assignmentStatuses']);
    for (const item of data.notificationQueue || []) {
      item.delivered = { desktop: Boolean(item.delivered?.desktop) };
      if (!settings.updates && /^(announcements|messages)-/.test(item.id)) { history.push(item.id); continue; }
      if (item.deadline && !settings.reminders.includes(item.deadline.hours)) { history.push(item.id); continue; }
      if (item.deadline) {
        const assignment = (assignments.assignments || []).find(a => String(a.eventId) === String(item.deadline.eventId));
        if (!assignment || Number(assignment.deadline) !== item.deadline.timestamp || item.deadline.timestamp <= now || assignment.manualStatus === 'submitted' || assignments.assignmentStatuses?.[assignment.eventId] === 'submitted') {
          history.push(item.id); continue;
        }
      }
      if (item.due > now) { remaining.push(item); continue; }
      if (settings.desktop && !item.delivered.desktop) {
        try {
          const id = `e3-alert-${item.id}`;
          await E3DesktopNotifications.create(id, item.options, item.url);
          links[id] = item.url || 'https://e3p.nycu.edu.tw/';
          item.delivered.desktop = true;
        } catch { deliveryError = '桌面通知失敗，請檢查瀏覽器與系統通知設定。'; }
      }
      if (!settings.desktop || item.delivered.desktop) history.push(item.id);
      else {
        item.attempts++;
        item.due = now + Math.min(3600000, 60000 * 2 ** Math.min(item.attempts, 6));
        remaining.push(item);
      }
    }
    await E3NotificationAPI.set({ notificationQueue: remaining, notificationHistory: history.slice(-2000),
      notificationLinks: Object.fromEntries(Object.entries(links).slice(-200)), notificationDeliveryError: deliveryError });
  }
  globalThis.E3Notifications = {
    normalize, dueAt,
    enqueue: (id, options, url) => serial(async () => { await add(id, options, url); await process(); }),
    tick: () => serial(process)
  };
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== 'local' || !['notificationSettings', 'assignments', 'assignmentStatuses', 'announcements', 'messages'].some(key => changes[key])) return;
    serial(async () => {
      await E3HelperI18n.ready;
      if (changes.notificationSettings) {
        const { notificationQueue = [] } = await E3NotificationAPI.get('notificationQueue');
        const settings = normalize(changes.notificationSettings.newValue);
        notificationQueue.forEach(item => { item.due = dueAt(settings, Date.now()); });
        await E3NotificationAPI.set({ notificationQueue });
      }
      const { notificationSettings } = await E3NotificationAPI.get('notificationSettings');
      if (normalize(notificationSettings).updates) {
        for (const key of ['announcements', 'messages']) {
          const change = changes[key];
          // The initial snapshot is a baseline, not a flood of old alerts.
          if (!change || !Array.isArray(change.oldValue)) continue;
          const old = new Set(change.oldValue.map(item => String(item.id)));
          for (const item of change.newValue || []) {
            if (item.id == null || old.has(String(item.id))) continue;
            await add(`${key}-${item.id}`, { type: 'basic', iconUrl: chrome.runtime.getURL('128.png'),
              title: E3HelperI18n.text(key === 'messages' ? '新信件' : '新公告'), message: `${item.title || item.subject || ''}\n${item.courseName || item.course || ''}` }, item.url);
          }
        }
      }
      if (['notificationSettings', 'assignments', 'assignmentStatuses', 'announcements', 'messages'].some(key => changes[key])) await process();
    }).catch(() => console.warn('E3 Helper: 通知處理失敗'));
  });
  chrome.notifications?.onClicked?.addListener(id => {
    if (!id.startsWith('e3-alert-')) return;
    E3NotificationAPI.get('notificationLinks').then(async data => {
      const url = data.notificationLinks?.[id];
      if (url && /^https:\/\/e3p?\.nycu\.edu\.tw\//.test(url)) await E3NotificationAPI.openTab(url);
      await E3NotificationAPI.clear(id);
    }).catch(() => console.warn('E3 Helper: 通知連結處理失敗'));
  });
})();
