// Persist desktop notification queues across MV3 worker restarts.
(function () {
  const queueFullError = '通知佇列已滿，請檢查桌面通知設定。';
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
      await E3NotificationAPI.set({ notificationDeliveryError: queueFullError });
      return false;
    }
    queue.push({ id, options, url, deadline, created: now, due: dueAt(settings, now), delivered: {}, attempts: 0 });
    await E3NotificationAPI.set({ notificationQueue: queue });
  }
  function updateOptions(source, item) {
    return { type: 'basic', iconUrl: chrome.runtime.getURL('128.png'),
      title: E3HelperI18n.text(source === 'messages' ? '新信件' : '新公告'),
      message: `${item.title || item.subject || ''}\n${item.courseName || item.course || ''}` };
  }
  async function process(now = Date.now()) {
    await E3HelperI18n.ready;
    let data = await E3NotificationAPI.get(['assignments', 'assignmentStatuses', 'notificationSettings', 'notificationPendingUpdates', 'announcements', 'messages']);
    const settings = normalize(data.notificationSettings);
    // Strip retired channel credentials while retaining timing and reminder preferences.
    if (data.notificationSettings && Object.keys(data.notificationSettings).some(key => !(key in settings))) {
      await E3NotificationAPI.set({ notificationSettings: settings });
    }
    // Overflow references use the already persisted source snapshots, not a
    // second unbounded queue of notification payloads. Retry on the next tick.
    const pendingUpdates = [];
    for (const pending of data.notificationPendingUpdates || []) {
      if (!settings.updates || !settings.desktop) continue;
      const item = data[pending.source]?.find(item => String(item.id) === pending.id);
      if (!item) continue;
      if (await add(`${pending.source}-${item.id}`, updateOptions(pending.source, item), item.url, now) === false) pendingUpdates.push(pending);
    }
    await E3NotificationAPI.set({ notificationPendingUpdates: pendingUpdates });
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
    data = await E3NotificationAPI.get(['notificationQueue', 'notificationHistory', 'notificationLinks', 'notificationDeliveryError']);
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
          links[id] = item.url || 'https://e3p.nycu.edu.tw/';
          await E3NotificationAPI.set({ notificationLinks: Object.fromEntries(Object.entries(links).slice(-200)) });
          await E3DesktopNotifications.create(id, item.options, item.url);
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
    if (!deliveryError && data.notificationDeliveryError === queueFullError && (remaining.length >= 200 || pendingUpdates.length)) deliveryError = queueFullError;
    await E3NotificationAPI.set({ notificationQueue: remaining, notificationHistory: history.slice(-2000),
      notificationLinks: Object.fromEntries(Object.entries(links).slice(-200)), notificationDeliveryError: deliveryError });
  }
  globalThis.E3Notifications = {
    normalize, dueAt,
    enqueue: (id, options, url) => serial(async () => {
      const accepted = await add(id, options, url);
      await process();
      if (accepted === false) {
        if (await add(id, options, url) === false) throw new Error(queueFullError);
        await process();
      }
    }),
    tick: () => serial(process)
  };
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== 'local' || !['notificationSettings', 'assignments', 'assignmentStatuses', 'announcements', 'messages'].some(key => changes[key])) return;
    serial(async () => {
      await E3HelperI18n.ready;
      if (changes.notificationSettings) {
        const { notificationQueue = [] } = await E3NotificationAPI.get('notificationQueue');
        const settings = normalize(changes.notificationSettings.newValue);
        const previous = normalize(changes.notificationSettings.oldValue);
        const timingChanged = previous.mode !== settings.mode || (settings.mode === 'daily' && previous.time !== settings.time);
        const now = Date.now();
        if (timingChanged) notificationQueue.forEach(item => {
          if (item.due > now) item.due = dueAt(settings, now);
        });
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
            if (await add(`${key}-${item.id}`, updateOptions(key, item), item.url) === false) {
              const { notificationPendingUpdates = [] } = await E3NotificationAPI.get('notificationPendingUpdates');
              if (!notificationPendingUpdates.some(pending => pending.source === key && pending.id === String(item.id))) {
                notificationPendingUpdates.push({ source: key, id: String(item.id) });
                await E3NotificationAPI.set({ notificationPendingUpdates });
              }
            }
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
      if (!url) return;
      if (/^https:\/\/e3p?\.nycu\.edu\.tw\//.test(url)) await E3NotificationAPI.openTab(url);
      await E3NotificationAPI.clear(id);
    }).catch(() => console.warn('E3 Helper: 通知連結處理失敗'));
  });
})();
