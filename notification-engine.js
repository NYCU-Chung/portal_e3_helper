// Persist desktop notification queues across MV3 worker restarts.
(function () {
  const queueFullError = '通知佇列已滿，請檢查桌面通知設定。';
  const events = { assignment: true, grading: true, announcement: true, message: true, participant: true };
  const defaults = { desktop: true, mode: 'instant', time: '20:00', reminders: [24, 1], events };
  // Notification IDs carry their event type; deadline reminders are governed by `reminders` instead of `events`.
  const prefixes = [['deadline', 'deadline-'], ['assignment', 'assignment-'], ['grading', 'grading-'],
    ['announcement', 'announcements-'], ['message', 'messages-'], ['participant', 'participant-']];
  const labels = { assignment: '新作業', grading: '作業評分', announcement: '公告', message: '信件', participant: '課程成員變動', deadline: '截止前提醒', other: '其他' };
  function categoryOf(id) { return prefixes.find(([, prefix]) => String(id).startsWith(prefix))?.[0] || 'other'; }
  function normalize(value = {}) {
    // Before per-event switches, `updates` covered both announcements and messages.
    const saved = value.events && typeof value.events === 'object' ? value.events : {};
    const legacy = value.updates !== false;
    return { desktop: E3DesktopNotifications.supported && (value.desktop ?? defaults.desktop), mode: value.mode === 'daily' ? 'daily' : 'instant',
      time: /^([01]\d|2[0-3]):[0-5]\d$/.test(value.time || '') ? value.time : defaults.time,
      reminders: (Array.isArray(value.reminders) ? value.reminders : defaults.reminders).filter(h => Number.isFinite(h) && h > 0 && h <= 720),
      events: Object.fromEntries(Object.keys(events).map(key => [key, typeof saved[key] === 'boolean' ? saved[key] : (key === 'announcement' || key === 'message') ? legacy : true])) };
  }
  function allowed(settings, id) { return settings.events[categoryOf(id)] ?? true; }
  function currentReminder(settings, left) {
    if (!Number.isFinite(left) || left <= 0) return;
    // On wake-up, use the nearest elapsed cutoff rather than backfilling each
    // historical threshold. A later, closer cutoff can still notify normally.
    return settings.reminders.filter(hours => left <= hours * 3600000).sort((a, b) => a - b)[0];
  }
  // Deadline reminders always go out immediately; holding them for the daily time would let short reminders expire.
  function dueAt(settings, now, deadline = null) {
    if (settings.mode !== 'daily' || deadline) return now;
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
    if (!settings.desktop || !allowed(settings, id)) return;
    const queue = data.notificationQueue || [];
    if ((data.notificationHistory || []).includes(id) || queue.some(item => item.id === id)) return;
    if (queue.length >= 200) {
      await E3NotificationAPI.set({ notificationDeliveryError: queueFullError });
      return false;
    }
    queue.push({ id, options, url, deadline, created: now, due: dueAt(settings, now, deadline), delivered: {}, attempts: 0 });
    await E3NotificationAPI.set({ notificationQueue: queue });
  }
  function deadlineId(assignment, hours) { return `deadline-${assignment.eventId}-${assignment.deadline}-${hours}`; }
  function unsubmitted(assignment, statuses) { return assignment.manualStatus !== 'submitted' && statuses?.[assignment.eventId] !== 'submitted'; }
  // One daily notification summarizing several held events instead of a burst of separate alerts.
  function digest(items) {
    const text = E3HelperI18n.text, counts = {};
    for (const item of items) { const category = categoryOf(item.id); counts[category] = (counts[category] || 0) + 1; }
    const summary = Object.entries(counts).map(([category, count]) => `${text(labels[category])} ${count}`).join(' · ');
    const lines = items.slice(0, 3).map(item => `• ${String(item.options?.message || item.options?.title || '').split('\n')[0]}`);
    if (items.length > 3) lines.push(`…+${items.length - 3}`);
    return { type: 'basic', iconUrl: chrome.runtime.getURL('128.png'), title: E3HelperI18n.template`E3 Helper：${items.length} 則新通知`,
      message: `${summary}\n${lines.join('\n')}`.slice(0, 300) };
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
      if (!allowed(settings, `${pending.source}-`) || !settings.desktop) continue;
      const item = data[pending.source]?.find(item => String(item.id) === pending.id);
      if (!item) continue;
      if (await add(`${pending.source}-${item.id}`, updateOptions(pending.source, item), item.url, now) === false) pendingUpdates.push(pending);
    }
    await E3NotificationAPI.set({ notificationPendingUpdates: pendingUpdates });
    for (const assignment of data.assignments || []) {
      if (!unsubmitted(assignment, data.assignmentStatuses)) continue;
      const left = Number(assignment.deadline) - now;
      const hours = currentReminder(settings, left);
      if (hours === undefined) continue;
      await add(deadlineId(assignment, hours), {
        type: 'basic', iconUrl: chrome.runtime.getURL('128.png'),
        title: E3HelperI18n.text(assignment.kind === 'exam' ? '考試即將開始' : '作業即將截止'),
        message: `${assignment.name}\n${assignment.course || ''}\n${new Date(assignment.deadline).toLocaleString(E3HelperI18n.language)}`
      }, assignment.url, now, { eventId: assignment.eventId, timestamp: Number(assignment.deadline), hours });
    }
    data = await E3NotificationAPI.get(['notificationQueue', 'notificationHistory', 'notificationLinks', 'notificationDeliveryError', 'notificationLog']);
    const history = data.notificationHistory || [];
    const links = data.notificationLinks || {};
    const log = data.notificationLog || [];
    const remaining = [], ready = [];
    let deliveryError = '';
    const assignments = await E3NotificationAPI.get(['assignments', 'assignmentStatuses']);
    for (const item of data.notificationQueue || []) {
      item.delivered = { desktop: Boolean(item.delivered?.desktop) };
      if (!allowed(settings, item.id)) { history.push(item.id); continue; }
      if (item.deadline && !settings.reminders.includes(item.deadline.hours)) { history.push(item.id); continue; }
      if (item.deadline) {
        const assignment = (assignments.assignments || []).find(a => String(a.eventId) === String(item.deadline.eventId));
        if (!assignment || Number(assignment.deadline) !== item.deadline.timestamp || item.deadline.timestamp <= now || !unsubmitted(assignment, assignments.assignmentStatuses) || item.deadline.hours !== currentReminder(settings, item.deadline.timestamp - now)) {
          history.push(item.id); continue;
        }
      }
      if (item.deadline && !item.attempts && !item.delivered.desktop && item.due > now) item.due = now;
      if (item.due > now) { remaining.push(item); continue; }
      ready.push(item);
    }
    const single = item => ({ id: item.id, options: item.options, url: item.url, items: [item] });
    const grouped = settings.desktop && settings.mode === 'daily' ? ready.filter(item => !item.deadline && !item.delivered.desktop) : [];
    const units = grouped.length >= 2
      ? [{ id: `digest-${Math.min(...grouped.map(item => item.due))}`, options: digest(grouped), url: 'https://e3p.nycu.edu.tw/', items: grouped },
        ...ready.filter(item => !grouped.includes(item)).map(single)]
      : ready.map(single);
    for (const unit of units) {
      let deliverySettings = settings, items = unit.items;
      if (settings.desktop && items.some(item => !item.delivered.desktop)) {
        try {
          links[`e3-alert-${unit.id}`] = unit.url || 'https://e3p.nycu.edu.tw/';
          await E3NotificationAPI.set({ notificationLinks: Object.fromEntries(Object.entries(links).slice(-200)) });
          // Settings writes happen in the options page independently of this
          // serial worker queue. Read again after link persistence, immediately
          // before create, so a saved disable applies to this delivery too.
          const latest = await E3NotificationAPI.get('notificationSettings');
          deliverySettings = normalize(latest.notificationSettings);
          const stale = items.filter(item => !allowed(deliverySettings, item.id) ||
            (item.deadline && item.deadline.hours !== currentReminder(deliverySettings, item.deadline.timestamp - now)));
          stale.forEach(item => history.push(item.id));
          items = items.filter(item => !stale.includes(item));
          // A summary that lost items to a newly disabled event is rebuilt, or sent as the one remaining alert.
          const target = items.length === unit.items.length ? unit : items.length > 1 ? { ...unit, options: digest(items), items } : items.length ? single(items[0]) : null;
          if (target && deliverySettings.desktop) {
            const id = `e3-alert-${target.id}`;
            if (!links[id]) {
              links[id] = target.url || 'https://e3p.nycu.edu.tw/';
              await E3NotificationAPI.set({ notificationLinks: Object.fromEntries(Object.entries(links).slice(-200)) });
            }
            await E3DesktopNotifications.create(id, target.options, target.url);
            items.forEach(item => { item.delivered.desktop = true; });
            log.push({ id: target.id, category: items.length > 1 ? 'digest' : categoryOf(target.id), title: target.options?.title || '', message: target.options?.message || '', url: target.url || '', time: now });
          }
        } catch { deliveryError = '桌面通知失敗，請檢查瀏覽器與系統通知設定。'; }
      }
      for (const item of items) {
        if (!deliverySettings.desktop || item.delivered.desktop) history.push(item.id);
        else {
          item.attempts++;
          item.due = now + Math.min(3600000, 60000 * 2 ** Math.min(item.attempts, 6));
          remaining.push(item);
        }
      }
    }
    if (!deliveryError && data.notificationDeliveryError === queueFullError && (remaining.length >= 200 || pendingUpdates.length)) deliveryError = queueFullError;
    await E3NotificationAPI.set({ notificationQueue: remaining, notificationHistory: history.slice(-2000), notificationLog: log.slice(-50),
      notificationLinks: Object.fromEntries(Object.entries(links).slice(-200)), notificationDeliveryError: deliveryError });
  }
  // A newly added reminder only applies to deadlines that cross it later, not to every assignment already inside it.
  async function baselineReminders(previous, next, now = Date.now()) {
    const added = next.reminders.filter(hours => !previous.reminders.includes(hours));
    if (!added.length) return;
    const data = await E3NotificationAPI.get(['assignments', 'assignmentStatuses', 'notificationHistory']);
    const history = data.notificationHistory || [];
    for (const assignment of data.assignments || []) {
      const left = Number(assignment.deadline) - now;
      if (!unsubmitted(assignment, data.assignmentStatuses) || !Number.isFinite(left) || left <= 0) continue;
      for (const hours of added) if (left <= hours * 3600000 && !history.includes(deadlineId(assignment, hours))) history.push(deadlineId(assignment, hours));
    }
    await E3NotificationAPI.set({ notificationHistory: history.slice(-2000) });
  }
  globalThis.E3Notifications = {
    normalize, dueAt, categoryOf,
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
          if (item.due > now) item.due = dueAt(settings, now, item.deadline);
        });
        await E3NotificationAPI.set({ notificationQueue });
        await baselineReminders(previous, settings, now);
      }
      const { notificationSettings } = await E3NotificationAPI.get('notificationSettings');
      const current = normalize(notificationSettings);
      for (const key of ['announcements', 'messages']) {
        const change = changes[key];
        // The initial snapshot is a baseline, not a flood of old alerts.
        if (!allowed(current, `${key}-`) || !change || !Array.isArray(change.oldValue)) continue;
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
