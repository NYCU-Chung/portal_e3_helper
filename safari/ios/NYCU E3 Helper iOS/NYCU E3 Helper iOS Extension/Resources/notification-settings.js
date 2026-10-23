(async function () {
  await E3HelperI18n.ready;
  // Same appearance setting as the sidebar: light unless dark or follow-the-system is saved.
  const { themePreference } = await E3NotificationAPI.get(['themePreference']);
  if (themePreference === 'system' || themePreference === 'dark') document.documentElement.dataset.theme = themePreference;
  const t = E3HelperI18n.text;
  const english = E3HelperI18n.language === 'en';
  const supported = E3DesktopNotifications.supported;
  const eventTypes = [
    ['assignment', '新作業', '老師上架新作業時通知。'],
    ['grading', '作業評分', '作業評分完成時通知。'],
    ['announcement', '公告', '同步發現新公告時通知。'],
    ['message', '信件', '同步發現新信件時通知。'],
    ['participant', '課程成員變動', '課程人數增加或減少時通知。']
  ];
  const categoryLabels = { ...Object.fromEntries(eventTypes.map(([key, name]) => [key, name])), deadline: '截止前提醒', digest: '每日摘要', other: '其他' };
  const deliveryErrors = ['桌面通知失敗，請檢查瀏覽器與系統通知設定。', '通知佇列已滿，請檢查桌面通知設定。'];
  document.documentElement.lang = E3HelperI18n.language;
  document.title = `${t('通知設定')} · E3 Helper`;
  document.getElementById('settings').innerHTML = E3HelperI18n.template`
    <header class="page-header"><span class="eyebrow">E3 HELPER</span><h1>通知設定</h1><p>選擇你需要的提醒，讓課程資訊在適合的時間出現。</p></header>
    <p id="delivery-error" class="notice error-notice" role="alert" hidden></p>
    <form id="form">
      <section class="card"><div class="setting-row"><div><h2 id="desktop-label">桌面通知</h2><p id="desktop-hint">在電腦的通知中心顯示提醒，點擊即可開啟 E3。</p></div><label class="switch"><input id="desktop" type="checkbox" role="switch" aria-labelledby="desktop-label" aria-describedby="desktop-hint"><span aria-hidden="true"></span></label></div>
      <p id="permission" class="permission" aria-live="polite" hidden><span class="dot" aria-hidden="true"></span><span id="permission-text"></span></p>
      <p id="support-notice" class="notice" hidden>此瀏覽器不支援桌面通知，仍可在側欄查看通知。</p></section>
      <fieldset id="preferences">
        <section class="card"><div class="section-heading"><h2>通知內容</h2><p>選擇要通知的事件。</p></div>
          <div id="event-list" class="event-list"></div>
          <div class="reminder-heading"><label id="reminder-label">截止前提醒</label><span class="tag">已繳交作業不提醒</span></div>
          <div id="reminder-list" aria-labelledby="reminder-label"></div>
          <button id="add-reminder" type="button" class="text-button">＋ 新增提醒時間</button>
          <p class="hint">可以設定多個提醒；移除所有時間即可關閉截止提醒。</p>
        </section>
        <section class="card"><div class="section-heading"><h2>通知時間</h2><p>所有通知都依照下方選擇的方式送出。</p></div>
          <div class="mode-options">
            <label class="mode-card"><input type="radio" name="mode" value="instant"><span><strong>即時通知</strong><small>同步發現更新或進入提醒時間時通知。</small></span></label>
            <label class="mode-card"><input type="radio" name="mode" value="daily"><span><strong>每天固定時間</strong><small>將待處理的通知留到指定時間送出。</small></span></label>
          </div>
          <div id="time-row" class="time-row" hidden><label for="time">每日通知時間</label><input id="time" type="time" required><span id="timezone" class="hint"></span><p class="hint">多則通知會合併成一則摘要；截止前提醒仍會準時送出。</p></div>
        </section>
      </fieldset>
      <aside class="summary"><span class="summary-label">目前設定</span><p id="summary" aria-live="polite"></p></aside>
      <footer class="actions"><p id="status" role="status" aria-live="polite"></p><div class="buttons"><button id="test" type="button" class="secondary">測試通知</button><button id="save" type="submit" class="primary" disabled>儲存設定</button></div></footer>
    </form>
    <section class="card history"><div class="section-heading"><h2>最近通知</h2><p>最近送出的 50 則桌面通知。</p></div><ol id="history-list" class="history-list"></ol><p id="history-empty" class="hint" hidden>尚無通知紀錄</p></section>
    <p class="footnote">瀏覽器需保持開啟；電腦休眠可能延後通知。新資料依 E3 同步週期更新。</p>`;
  const data = await E3NotificationAPI.get(['notificationSettings', 'notificationDeliveryError', 'notificationLog']);
  const saved = data.notificationSettings || {};
  const settings = { desktop: true, mode: 'instant', time: '20:00', reminders: [24, 1], ...saved };
  // Older settings had one switch (`updates`) for both announcements and messages.
  const savedEvents = saved.events && typeof saved.events === 'object' ? saved.events : {};
  const eventEnabled = key => typeof savedEvents[key] === 'boolean' ? savedEvents[key] : (key === 'announcement' || key === 'message') ? saved.updates !== false : true;
  const desktop = document.getElementById('desktop');
  const list = document.getElementById('reminder-list');
  const status = document.getElementById('status');
  const save = document.getElementById('save');
  const test = document.getElementById('test');
  const presets = [168, 72, 48, 24, 12, 6, 3, 1, 0.5];
  let dirty = false;
  let revision = 0;
  let saving = false;
  function label(hours) {
    if (hours === 0.5) return english ? '30 minutes before' : '截止前 30 分鐘';
    if (hours % 24 === 0) return english ? `${hours / 24} day${hours === 24 ? '' : 's'} before` : `截止前 ${hours / 24} 天`;
    return english ? `${hours} hour${hours === 1 ? '' : 's'} before` : `截止前 ${hours} 小時`;
  }
  const eventList = document.getElementById('event-list');
  for (const [key, name, hint] of eventTypes) {
    const row = document.createElement('label'); row.className = 'check-row';
    const input = document.createElement('input'); input.type = 'checkbox'; input.dataset.event = key; input.checked = eventEnabled(key);
    const text = document.createElement('span'), strong = document.createElement('strong'), small = document.createElement('small');
    strong.textContent = t(name); small.textContent = t(hint);
    text.append(strong, small); row.append(input, text); eventList.appendChild(row);
  }
  const eventInputs = [...eventList.querySelectorAll('input[data-event]')];
  function values() { return [...new Set([...list.querySelectorAll('select')].map(select => Number(select.value)))].sort((a, b) => b - a); }
  function refresh(changed = false) {
    const enabled = supported && desktop.checked;
    document.getElementById('preferences').disabled = !enabled;
    const daily = document.querySelector('input[name=mode]:checked').value === 'daily';
    document.getElementById('time-row').hidden = !daily;
    document.getElementById('time').required = daily && enabled;
    // A time chosen in one row is unavailable in the others.
    const selects = [...list.querySelectorAll('select')];
    for (const select of selects) {
      const others = selects.filter(other => other !== select).map(other => Number(other.value));
      for (const option of select.options) option.disabled = others.includes(Number(option.value));
    }
    const separator = english ? ', ' : '、';
    const reminders = values().map(label).join(separator);
    const off = eventInputs.filter(input => !input.checked).map(input => t(categoryLabels[input.dataset.event])).join(separator);
    document.getElementById('summary').textContent = !enabled ? t('桌面通知已關閉。') : (english
      ? `${daily ? `Daily at ${document.getElementById('time').value}` : 'Notify immediately'}. ${reminders ? `Reminders: ${reminders}.` : 'Deadline reminders are off.'}${off ? ` Not notifying: ${off}.` : ''}`
      : `${daily ? `每天 ${document.getElementById('time').value} 通知` : '即時通知'}。${reminders ? `作業${reminders.replaceAll('截止前 ', '截止前')}提醒。` : '截止提醒已關閉。'}${off ? `不通知：${off}。` : ''}`);
    document.getElementById('add-reminder').disabled = !enabled || presets.every(value => values().includes(value));
    if (changed) { revision++; dirty = true; save.disabled = saving; status.dataset.state = 'pending'; status.textContent = t('尚未儲存'); }
  }
  function addReminder(hours) {
    const row = document.createElement('div'); row.className = 'reminder-row';
    const select = document.createElement('select');
    select.setAttribute('aria-label', t('截止前提醒'));
    for (const value of [...new Set([...presets, hours])].sort((a, b) => b - a)) {
      const option = document.createElement('option'); option.value = value; option.textContent = label(value); select.appendChild(option);
    }
    select.value = hours;
    const remove = document.createElement('button'); remove.type = 'button'; remove.className = 'remove-button'; remove.textContent = t('移除');
    remove.setAttribute('aria-label', `${t('移除')} ${label(hours)}`);
    remove.addEventListener('click', () => { const next = row.nextElementSibling?.querySelector('select') || row.previousElementSibling?.querySelector('select') || document.getElementById('add-reminder'); row.remove(); refresh(true); next.focus(); });
    select.addEventListener('change', () => remove.setAttribute('aria-label', `${t('移除')} ${label(Number(select.value))}`));
    row.append(select, remove); list.appendChild(row);
    return select;
  }
  async function showPermission() {
    const row = document.getElementById('permission');
    if (!supported) { row.hidden = true; return; }
    let level = 'unknown';
    try { level = await E3DesktopNotifications.getPermissionLevel(); } catch {}
    row.dataset.state = level === 'granted' ? 'granted' : level === 'denied' ? 'denied' : 'unknown';
    document.getElementById('permission-text').textContent = t(level === 'granted' ? '系統通知權限：已允許'
      : level === 'denied' ? '系統通知權限：已封鎖，請到瀏覽器與系統設定開啟' : '系統通知權限：尚未決定');
    row.hidden = false;
  }
  // Delivery problems stay visible until the engine clears them, independent of save status.
  function showDeliveryError(value) {
    const element = document.getElementById('delivery-error');
    element.hidden = !deliveryErrors.includes(value);
    element.textContent = element.hidden ? '' : t(value);
  }
  function showHistory(log) {
    const items = (Array.isArray(log) ? log : []).slice(-50).reverse();
    const target = document.getElementById('history-list');
    target.replaceChildren();
    document.getElementById('history-empty').hidden = items.length > 0;
    for (const entry of items) {
      const item = document.createElement('li');
      const meta = document.createElement('span'); meta.className = 'history-meta';
      const time = new Date(Number(entry.time));
      meta.textContent = `${t(categoryLabels[entry.category] || categoryLabels.other)} · ${Number.isFinite(time.getTime()) ? time.toLocaleString(E3HelperI18n.language) : ''}`;
      const url = String(entry.url || '');
      const title = /^https:\/\/e3p?\.nycu\.edu\.tw\//.test(url) ? document.createElement('a') : document.createElement('strong');
      if (title.tagName === 'A') { title.href = url; title.target = '_blank'; title.rel = 'noopener noreferrer'; }
      title.className = 'history-title'; title.textContent = String(entry.title || '');
      const message = document.createElement('small'); message.textContent = String(entry.message || '');
      item.append(meta, title, message); target.appendChild(item);
    }
  }
  desktop.checked = supported && settings.desktop;
  desktop.disabled = !supported; test.disabled = !supported;
  document.getElementById('support-notice').hidden = supported;
  document.querySelector(`input[name=mode][value="${settings.mode === 'daily' ? 'daily' : 'instant'}"]`).checked = true;
  document.getElementById('time').value = /^([01]\d|2[0-3]):[0-5]\d$/.test(settings.time) ? settings.time : '20:00';
  document.getElementById('timezone').textContent = Intl.DateTimeFormat().resolvedOptions().timeZone;
  for (const hours of Array.isArray(settings.reminders) ? settings.reminders : [24, 1]) if (Number.isFinite(hours) && hours > 0 && hours <= 720 && !values().includes(hours)) addReminder(hours);
  document.getElementById('add-reminder').addEventListener('click', () => { const hours = presets.find(value => !values().includes(value)); if (hours === undefined) return; const select = addReminder(hours); refresh(true); select.focus(); });
  document.getElementById('form').addEventListener('change', () => refresh(true));
  document.getElementById('time').addEventListener('input', () => refresh(true));
  window.addEventListener('beforeunload', event => { if (!dirty) return; event.preventDefault(); event.returnValue = ''; });
  refresh();
  showDeliveryError(data.notificationDeliveryError);
  showHistory(data.notificationLog);
  showPermission();
  chrome.storage.onChanged.addListener((changes, area) => {
    if (area !== 'local') return;
    if (changes.notificationDeliveryError) showDeliveryError(changes.notificationDeliveryError.newValue);
    if (changes.notificationLog) showHistory(changes.notificationLog.newValue);
  });
  document.getElementById('form').addEventListener('submit', async event => {
    event.preventDefault();
    if (saving) return;
    saving = true; save.disabled = true; save.textContent = t('儲存中...');
    const savedRevision = revision;
    const notificationSettings = {
      desktop: desktop.checked, mode: document.querySelector('input[name=mode]:checked').value, time: document.getElementById('time').value, reminders: values(),
      events: Object.fromEntries(eventInputs.map(input => [input.dataset.event, input.checked]))
    };
    try {
      if (notificationSettings.desktop && E3DesktopNotifications.native && await E3DesktopNotifications.authorize() !== 'granted') {
        throw new Error(t('請在瀏覽器與系統設定允許通知。'));
      }
      await E3NotificationAPI.set({ notificationSettings });
      dirty = revision !== savedRevision;
      status.textContent = t(dirty ? '尚未儲存' : '設定已儲存！'); status.dataset.state = dirty ? 'pending' : 'success';
    } catch (error) { status.textContent = error.message || t('儲存失敗，請再試一次。'); status.dataset.state = 'error'; save.disabled = false; }
    finally { saving = false; save.disabled = !dirty; save.textContent = t('儲存設定'); showPermission(); }
  });
  test.addEventListener('click', async () => {
    test.disabled = true; test.textContent = t('測試中...');
    try {
      const permission = await E3DesktopNotifications.authorize();
      if (permission !== 'granted') throw new Error(t('請在瀏覽器與系統設定允許通知。'));
      await E3DesktopNotifications.create('e3-notification-test', { type: 'basic', iconUrl: chrome.runtime.getURL('128.png'), title: 'E3 Helper', message: t('這是一則測試通知。') });
      status.textContent = t('已送出測試通知，請查看系統通知。'); status.dataset.state = 'success';
    } catch (error) { status.textContent = error.message; status.dataset.state = 'error'; }
    finally { test.disabled = false; test.textContent = t('測試通知'); showPermission(); }
  });
})();
