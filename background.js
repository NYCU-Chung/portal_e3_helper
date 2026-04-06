// NYCU E3 Helper - Background Script (Service Worker)
// 處理下載請求和自動同步

// ==================== 日誌系統 ====================
// 保存原始 console 方法
const originalConsole = {
  log: console.log.bind(console),
  info: console.info.bind(console),
  warn: console.warn.bind(console),
  error: console.error.bind(console),
  debug: console.debug.bind(console)
};

// 儲存日誌到 storage 並通知 content script
function sendLogToContentScript(type, args) {
  // 調用原始 console
  originalConsole[type](...args);

  // 轉換參數為可序列化的格式
  const serializedArgs = args.map(arg => {
    if (arg === null) return 'null';
    if (arg === undefined) return 'undefined';
    if (typeof arg === 'object') {
      try {
        return JSON.stringify(arg, null, 2);
      } catch (e) {
        return String(arg);
      }
    }
    return String(arg);
  });

  // 儲存到 chrome.storage（供後續查詢）
  chrome.storage.local.get(['backgroundLogs'], (result) => {
    const logs = result.backgroundLogs || [];
    logs.push({
      type: type,
      args: serializedArgs,
      timestamp: Date.now(),
      time: new Date().toLocaleTimeString('zh-TW', { hour12: false })
    });

    // 只保留最近 200 條日誌
    if (logs.length > 200) {
      logs.splice(0, logs.length - 200);
    }

    chrome.storage.local.set({ backgroundLogs: logs });
  });

  // 廣播日誌到有 E3 Helper sidebar 的 tabs（不廣播到所有 tabs）
  chrome.tabs.query({ url: ['*://*.nycu.edu.tw/*'] }, (tabs) => {
    tabs.forEach(tab => {
      chrome.tabs.sendMessage(tab.id, {
        action: 'backgroundLog',
        type: type,
        args: serializedArgs,
        time: new Date().toLocaleTimeString('zh-TW', { hour12: false })
      }).catch(() => {}); // 忽略錯誤（某些 tab 可能沒有 content script）
    });
  });
}

// 攔截 console 方法
console.log = (...args) => sendLogToContentScript('log', args);
console.info = (...args) => sendLogToContentScript('info', args);
console.warn = (...args) => sendLogToContentScript('warn', args);
console.error = (...args) => sendLogToContentScript('error', args);
console.debug = (...args) => sendLogToContentScript('debug', args);

console.log('E3 Helper Background Script 已載入');

// 監聽來自 content script 的訊息
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'download') {
    console.log(`E3 Helper: 收到下載請求 - ${request.filename}`);

    // 驗證下載 URL 域名
    const downloadAllowedDomains = ['e3.nycu.edu.tw', 'e3p.nycu.edu.tw'];
    try {
      const urlObj = new URL(request.url);
      if (!downloadAllowedDomains.some(d => urlObj.hostname.endsWith(d))) {
        sendResponse({ success: false, error: 'Download URL not from allowed domain' });
        return true;
      }
    } catch {
      sendResponse({ success: false, error: 'Invalid URL' });
      return true;
    }
    // 驗證 filename 無路徑遍歷
    if (request.filename && (request.filename.includes('..') || request.filename.includes('/'))) {
      sendResponse({ success: false, error: 'Invalid filename' });
      return true;
    }

    // 使用 Chrome Downloads API 下載檔案
    chrome.downloads.download({
      url: request.url,
      filename: request.filename,
      saveAs: false // 不顯示儲存對話框，直接下載到預設位置
    }, (downloadId) => {
      if (chrome.runtime.lastError) {
        console.error('E3 Helper: 下載失敗', chrome.runtime.lastError);
        sendResponse({ success: false, error: chrome.runtime.lastError.message });
      } else {
        console.log(`E3 Helper: 下載已開始，ID: ${downloadId}`);
        sendResponse({ success: true, downloadId: downloadId });
      }
    });

    // 返回 true 表示會異步回應
    return true;
  } else if (request.action === 'syncNow') {
    // 手動觸發同步
    console.log('E3 Helper: 收到手動同步請求');
    syncE3Data().then(result => {
      sendResponse(result);
    });
    return true;
  } else if (request.action === 'updateBadge') {
    // 更新擴充功能圖標 badge
    const count = request.count || 0;
    console.log(`E3 Helper: 更新 badge 計數 - ${count}`);

    if (chrome.action) {
      if (count > 0) {
        chrome.action.setBadgeText({ text: count > 99 ? '99+' : count.toString() });
        chrome.action.setBadgeBackgroundColor({ color: '#dc3545' });
      } else {
        chrome.action.setBadgeText({ text: '' });
      }
    } else {
      console.warn('E3 Helper: chrome.action API 不可用');
    }

    sendResponse({ success: true });
    return true;
  } else if (request.action === 'showNotification') {
    // 發送桌面通知
    console.log(`E3 Helper: 發送通知 - ${request.title}`);

    chrome.notifications.create({
      type: 'basic',
      iconUrl: 'chrome-extension://' + chrome.runtime.id + '/128.png',
      title: request.title,
      message: request.message,
      priority: 2,
      requireInteraction: false
    }, (notificationId) => {
      if (chrome.runtime.lastError) {
        console.error('E3 Helper: 發送通知失敗', chrome.runtime.lastError);
        sendResponse({ success: false });
      } else {
        console.log(`E3 Helper: 通知已發送，ID: ${notificationId}`);
        sendResponse({ success: true });
      }
    });

    return true;
  } else if (request.action === 'checkParticipants') {
    // 手動觸發成員檢測
    console.log('E3 Helper: 收到手動成員檢測請求');
    checkParticipantsInTabs();
    sendResponse({ success: true, message: '已觸發成員檢測' });
    return true;
  } else if (request.action === 'callGeminiApi') {
    // 代理 Gemini API 呼叫（避免 API key 暴露在 content script）
    (async () => {
      try {
        const { model, apiKey, content, generationConfig } = request;
        const requestBody = {
          contents: [{ parts: [{ text: content }] }]
        };
        if (generationConfig) {
          requestBody.generationConfig = generationConfig;
        }
        const response = await fetch(
          `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`,
          {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(requestBody)
          }
        );
        if (!response.ok) {
          const errorData = await response.json();
          sendResponse({ success: false, error: errorData.error?.message || `HTTP ${response.status}` });
          return;
        }
        const data = await response.json();
        if (!data.candidates || data.candidates.length === 0) {
          sendResponse({ success: false, error: data.promptFeedback?.blockReason || 'Gemini API 返回空結果' });
          return;
        }
        const candidate = data.candidates[0];
        if (candidate.content?.parts?.[0]?.text) {
          sendResponse({ success: true, data: candidate.content.parts[0].text.trim() });
        } else {
          sendResponse({ success: false, error: candidate.finishReason || 'Gemini API 返回格式錯誤' });
        }
      } catch (error) {
        sendResponse({ success: false, error: error.message });
      }
    })();
    return true;
  } else if (request.action === 'callAI') {
    // 處理 AI API 請求
    handleAIRequest(request)
      .then(response => sendResponse({ success: true, data: response }))
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true;
  } else if (request.action === 'loadAnnouncementsAndMessages') {
    // 從非 E3 網站請求載入公告和信件
    console.log('E3 Helper: 收到載入公告和信件的請求');
    loadAnnouncementsAndMessagesInBackground()
      .then(result => sendResponse(result))
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true;
  } else if (request.action === 'fetchContent') {
    // 從非 E3 網站抓取公告/信件內容
    console.log(`E3 Helper: 收到抓取內容請求 - ${request.url}`);
    fetchContentFromE3(request.url)
      .then(html => sendResponse({ success: true, html: html }))
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true;
  }
});

// ==================== 自動同步功能 ====================

// 安裝時設定定時同步
chrome.runtime.onInstalled.addListener(() => {
  console.log('E3 Helper: 擴充功能已安裝/更新');

  // 設定每小時同步一次（作業、課程）
  chrome.alarms.create('syncE3Data', {
    periodInMinutes: 60
  });

  // 設定每小時檢查課程成員變動
  chrome.alarms.create('checkParticipants', {
    periodInMinutes: 60
  });

  // 設定每 6 小時同步公告和信件
  chrome.alarms.create('syncAnnouncementsAndMessages', {
    periodInMinutes: 360  // 6 小時
  });

  // 立即執行一次同步
  syncE3Data();

  // 初始化 badge 計數
  updateBadgeFromStorage();
});

// Service Worker 啟動時也更新 badge
chrome.runtime.onStartup.addListener(() => {
  console.log('E3 Helper: Service Worker 啟動');
  updateBadgeFromStorage();
});

// 從 storage 更新 badge 計數
async function updateBadgeFromStorage() {
  try {
    if (!chrome.action) {
      console.warn('E3 Helper: chrome.action API 不可用，跳過 badge 更新');
      return;
    }

    const storage = await chrome.storage.local.get(['notifications', 'participantChangeNotifications', 'urgentAssignmentNotifications']);
    const assignmentNotifications = storage.notifications || [];
    const participantNotifications = storage.participantChangeNotifications || [];
    const urgentNotifications = storage.urgentAssignmentNotifications || [];

    const unreadCount = assignmentNotifications.filter(n => !n.read).length +
                        participantNotifications.filter(n => !n.read).length +
                        urgentNotifications.filter(n => !n.read).length;

    if (unreadCount > 0) {
      chrome.action.setBadgeText({ text: unreadCount > 99 ? '99+' : unreadCount.toString() });
      chrome.action.setBadgeBackgroundColor({ color: '#dc3545' });
    } else {
      chrome.action.setBadgeText({ text: '' });
    }

    console.log(`E3 Helper: Badge 已更新 - ${unreadCount} 個未讀通知`);
  } catch (error) {
    console.error('E3 Helper: 更新 badge 時發生錯誤', error);
  }
}

// 監聽定時器
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === 'syncE3Data') {
    console.log('E3 Helper: 定時同步觸發');
    syncE3Data();
  } else if (alarm.name === 'checkParticipants') {
    console.log('E3 Helper: 定時檢查課程成員變動');
    checkParticipantsInTabs();
  } else if (alarm.name === 'syncAnnouncementsAndMessages') {
    console.log('E3 Helper: 定時同步公告和信件');
    syncAnnouncementsAndMessagesSilently();
  }
});

// 同步 E3 資料
async function syncE3Data() {
  console.log('E3 Helper: 開始同步資料...', new Date().toLocaleTimeString());

  const syncResult = {
    success: false,
    timestamp: Date.now(),
    assignments: 0,
    courses: 0,
    error: null,
    loginRequired: false
  };

  try {
    // 檢查登入狀態
    console.log('E3 Helper: 檢查登入狀態...');
    const isLoggedIn = await checkLoginStatus();
    console.log('E3 Helper: 登入狀態:', isLoggedIn);

    if (!isLoggedIn) {
      console.log('E3 Helper: 未登入 E3，無法同步');
      syncResult.loginRequired = true;
      syncResult.error = '請先登入 E3';
      await chrome.storage.local.set({ lastSync: syncResult });
      return syncResult;
    }

    // 同步作業資料
    console.log('E3 Helper: 開始同步作業...');
    const assignments = await syncAssignments();
    console.log(`E3 Helper: 作業同步完成，共 ${assignments.length} 個`);
    syncResult.assignments = assignments.length;

    // 同步課程列表
    console.log('E3 Helper: 開始同步課程...');
    const courses = await syncCourses();
    console.log(`E3 Helper: 課程同步完成，共 ${courses.length} 個`);
    syncResult.courses = courses.length;

    // 同步成績資料（僅在手動同步時執行，避免自動同步太慢）
    // 注意：這裡先不同步成績，因為會花很長時間

    // 儲存同步結果
    syncResult.success = true;
    await chrome.storage.local.set({
      lastSync: syncResult,
      lastSyncTime: Date.now()
    });

    console.log('E3 Helper: 同步完成', syncResult);
    return syncResult;

  } catch (error) {
    console.error('E3 Helper: 同步失敗', error);
    syncResult.error = error.message;

    // 檢查是否是登入問題
    if (error.message.includes('login') || error.message.includes('401')) {
      syncResult.loginRequired = true;
    }

    await chrome.storage.local.set({ lastSync: syncResult });
    return syncResult;
  }
}

// 檢查 E3 登入狀態
async function checkLoginStatus() {
  try {
    const response = await fetchWithTimeout('https://e3p.nycu.edu.tw/', {
      method: 'GET',
      credentials: 'include'
    }, 10000);

    // 被重導向到登入頁面 → 確定未登入
    if (response.url.includes('/login/')) {
      return false;
    }

    const text = await response.text();

    // 包含登入表單 → 確定未登入
    if (text.includes('loginform')) {
      return false;
    }

    // sesskey 是最可靠的登入指標（未登入時不會有 sesskey）
    if (text.includes('sesskey') || text.includes('data-userid') ||
        text.includes('logout') || text.includes('登出')) {
      return true;
    }

    return false;
  } catch (error) {
    // 網路錯誤時返回 null，由呼叫方決定如何處理
    console.warn('E3 Helper: 檢查登入狀態時網路錯誤', error.message);
    return null;
  }
}

// 取得 sesskey
async function getSesskey() {
  try {
    const response = await fetchWithTimeout('https://e3p.nycu.edu.tw/', {
      credentials: 'include'
    }, 10000); // 增加到 10秒
    const html = await response.text();

    // 從 HTML 中提取 sesskey
    const sesskeyMatch = html.match(/"sesskey"\s*:\s*"([a-zA-Z0-9]+)"/) || html.match(/sesskey=([a-zA-Z0-9]+)&/);
    if (sesskeyMatch) {
      return sesskeyMatch[1];
    }

    // 嘗試從 M.cfg.sesskey 提取
    const mConfigMatch = html.match(/M\.cfg\s*=\s*\{[^}]*"sesskey"\s*:\s*"([^"]+)"/);
    if (mConfigMatch) {
      return mConfigMatch[1];
    }

    return null;
  } catch (error) {
    console.error('E3 Helper: 取得 sesskey 失敗', error);
    return null;
  }
}

// 帶超時的 fetch
async function fetchWithTimeout(url, options, timeout = 10000) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeout);

  try {
    const response = await fetch(url, {
      ...options,
      signal: controller.signal
    });
    clearTimeout(timeoutId);
    return response;
  } catch (error) {
    clearTimeout(timeoutId);
    if (error.name === 'AbortError') {
      throw new Error('請求超時');
    }
    throw error;
  }
}

// 同步作業資料
async function syncAssignments() {
  console.log('E3 Helper: 正在同步作業...');

  const sesskey = await getSesskey();
  if (!sesskey) {
    throw new Error('無法取得 sesskey，請重新登入');
  }

  const url = `https://e3p.nycu.edu.tw/lib/ajax/service.php?sesskey=${sesskey}`;

  // 獲取日曆事件（15秒超時）
  const response = await fetchWithTimeout(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    credentials: 'include',
    body: JSON.stringify([{
      index: 0,
      methodname: 'core_calendar_get_action_events_by_timesort',
      args: {
        limitnum: 50,
        timesortfrom: Math.floor(Date.now() / 1000) - 86400 * 30, // 過去 30 天（確保較舊的作業也能被抓取）
        timesortto: Math.floor(Date.now() / 1000) + 86400 * 90  // 未來 90 天
      }
    }])
  }, 15000); // 增加到 15秒

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  }

  const data = await response.json();
  console.log('E3 Helper: API 回應:', data);

  if (data && data[0] && data[0].error) {
    console.error('E3 Helper: API 錯誤:', data[0].error);
    throw new Error(data[0].error);
  }

  if (data && data[0] && data[0].data && data[0].data.events) {
    const events = data[0].data.events;

    // 過濾出作業
    const assignments = events
      .filter(event =>
        event.modulename === 'assign' ||
        event.name.includes('作業')
      )
      .map(event => ({
        eventId: event.id.toString(),
        name: event.name,
        course: event.course ? event.course.fullname : '',
        deadline: event.timesort * 1000,
        url: event.url,
        manualStatus: 'pending'
      }));

    console.log(`E3 Helper: E3 API 返回了 ${assignments.length} 個作業:`);
    console.log('E3 Helper: 作業 ID 列表:', assignments.map(a => ({
      id: a.eventId,
      name: a.name,
      hasURL: !!a.url,
      urlValid: a.url && a.url.includes('mod/assign'),
      url: a.url ? a.url.substring(0, 50) + '...' : '無',
      deadline: new Date(a.deadline).toLocaleString()
    })));

    // 先補齊新作業的 URL 和課程名稱（在合併之前）
    const newAssignmentsNeedingDetails = assignments.filter(a =>
      (!a.course || a.course === '') || (!a.url || a.url === '' || !a.url.includes('mod/assign'))
    );

    if (newAssignmentsNeedingDetails.length > 0) {
      console.log(`E3 Helper: 發現 ${newAssignmentsNeedingDetails.length} 個新作業需要補齊詳細資訊，從 API 獲取...`);

      for (const assignment of newAssignmentsNeedingDetails) {
        try {
          const eventDetailUrl = `https://e3p.nycu.edu.tw/lib/ajax/service.php?sesskey=${sesskey}`;
          const eventResponse = await fetchWithTimeout(eventDetailUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            credentials: 'include',
            body: JSON.stringify([{
              index: 0,
              methodname: 'core_calendar_get_calendar_event_by_id',
              args: { eventid: parseInt(assignment.eventId) }
            }])
          }, 10000);

          if (eventResponse.ok) {
            const eventData = await eventResponse.json();
            if (eventData && eventData[0] && eventData[0].data && eventData[0].data.event) {
              const event = eventData[0].data.event;

              // 補齊課程名稱
              if (event.course && event.course.fullname && (!assignment.course || assignment.course === '')) {
                assignment.course = event.course.fullname;
                console.log(`E3 Helper: 新作業 ${assignment.eventId} (${assignment.name}) 補齊課程: ${event.course.fullname}`);
              }

              // 補齊 URL
              if (event.url && (!assignment.url || assignment.url === '')) {
                assignment.url = event.url;
                console.log(`E3 Helper: 新作業 ${assignment.eventId} (${assignment.name}) 補齊 URL`);
              }
            }
          }
        } catch (error) {
          console.error(`E3 Helper: 獲取新作業 ${assignment.eventId} 詳細資訊失敗:`, error);
        }
      }
    }

    // 載入現有的手動狀態和舊作業列表
    const storage = await chrome.storage.local.get(['assignmentStatuses', 'assignments']);
    const statuses = storage.assignmentStatuses || {};
    const oldAssignments = storage.assignments || [];
    console.log('E3 Helper: 讀取到的手動狀態:', statuses);
    console.log('E3 Helper: 手動狀態數量:', Object.keys(statuses).length);

    // 合併手動狀態和舊資料（包括課程名稱、URL 等）
    let mergedCount = 0;
    const oldAssignmentMap = new Map(oldAssignments.map(a => [a.eventId, a]));

    // 定義無效的課程名稱（這些是頁面標題，不是真正的課程名稱）
    const invalidCourseNames = ['焦點綜覽', '通知', '時間軸', 'Timeline', 'Notifications', '概覽', 'Overview'];
    const isInvalidCourse = (course) => !course || course === '' || invalidCourseNames.includes(course);

    assignments.forEach(assignment => {
      const oldAssignment = oldAssignmentMap.get(assignment.eventId);

      // 合併手動狀態
      if (statuses[assignment.eventId]) {
        assignment.manualStatus = statuses[assignment.eventId];
        mergedCount++;
        console.log(`E3 Helper: 合併狀態 - 作業 ${assignment.eventId}: ${statuses[assignment.eventId]}`);
      }

      // 如果新作業沒有課程名稱（或是無效名稱），但舊資料有有效課程名稱，則保留舊的
      if (isInvalidCourse(assignment.course) && oldAssignment && oldAssignment.course && !isInvalidCourse(oldAssignment.course)) {
        assignment.course = oldAssignment.course;
        console.log(`E3 Helper: 保留課程名稱 - 作業 ${assignment.eventId}: ${oldAssignment.course}`);
      }

      // 如果新作業的課程名稱無效，清空它（讓後續 API 補齊）
      if (isInvalidCourse(assignment.course)) {
        assignment.course = '';
      }

      // 如果新作業沒有 URL，但舊資料有，則保留舊的 URL
      if (!assignment.url && oldAssignment && oldAssignment.url) {
        assignment.url = oldAssignment.url;
      }
    });
    console.log(`E3 Helper: 已合併 ${mergedCount} 個手動狀態到 ${assignments.length} 個作業`);

    // 找出那些已標記為「已繳交」但不在新列表中的舊作業（可能已過期但用戶想保留）
    const newAssignmentIds = new Set(assignments.map(a => a.eventId));

    console.log('E3 Helper: 檢查舊作業是否需要保留...');
    console.log('E3 Helper: 舊作業列表:', oldAssignments.map(a => ({
      id: a.eventId,
      name: a.name,
      manualStatus: a.manualStatus,
      inNewList: newAssignmentIds.has(a.eventId),
      inStatuses: !!statuses[a.eventId]
    })));

    const keptOldAssignments = oldAssignments
      .filter(oldAssignment => {
        // 如果作業在新列表中，不需要特別保留（已經在新列表中了）
        if (newAssignmentIds.has(oldAssignment.eventId)) {
          return false;
        }

        // 作業不在新列表中，檢查是否應該保留
        const now = Date.now();
        const isExpired = oldAssignment.deadline < now;

        if (!isExpired) {
          // 未過期：保留（不管是否已繳交）
          console.log(`E3 Helper: 保留未過期作業 - ${oldAssignment.name} (${oldAssignment.eventId})`);
          return true;
        } else {
          // 已過期：只保留已繳交的
          const isManuallySubmitted = statuses[oldAssignment.eventId] === 'submitted';
          const isAutoSubmitted = oldAssignment.manualStatus === 'submitted';
          const shouldKeep = isManuallySubmitted || isAutoSubmitted;

          if (shouldKeep) {
            console.log(`E3 Helper: 保留已過期但已繳交的作業 - ${oldAssignment.name} (${oldAssignment.eventId})`);
          }

          return shouldKeep;
        }
      })
      .map(oldAssignment => {
        // 確保 manualStatus 是最新的，並清理無效的課程名稱
        const cleanedAssignment = {
          ...oldAssignment,
          manualStatus: statuses[oldAssignment.eventId]
        };

        // 如果課程名稱無效，清空它（讓後續 API 補齊）
        if (isInvalidCourse(cleanedAssignment.course)) {
          cleanedAssignment.course = '';
          console.log(`E3 Helper: 保留的舊作業 ${cleanedAssignment.eventId} 有無效課程名稱，已清空`);
        }

        return cleanedAssignment;
      });

    if (keptOldAssignments.length > 0) {
      console.log(`E3 Helper: 保留 ${keptOldAssignments.length} 個已繳交的舊作業:`,
                  keptOldAssignments.map(a => ({ id: a.eventId, name: a.name, status: a.manualStatus })));
      // 將舊作業加到列表末尾
      assignments.push(...keptOldAssignments);
    }

    // 對於保留的舊作業，也需要補齊 URL 和課程名稱
    // （因為舊作業可能也缺少這些資訊）
    const keptAssignmentsNeedingDetails = keptOldAssignments.filter(a =>
      (!a.course || a.course === '') || (!a.url || a.url === '')
    );

    if (keptAssignmentsNeedingDetails.length > 0) {
      console.log(`E3 Helper: 發現 ${keptAssignmentsNeedingDetails.length} 個保留的舊作業需要補齊詳細資訊，從 API 獲取...`);

      for (const assignment of keptAssignmentsNeedingDetails) {
        try {
          const eventDetailUrl = `https://e3p.nycu.edu.tw/lib/ajax/service.php?sesskey=${sesskey}`;
          const eventResponse = await fetchWithTimeout(eventDetailUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            credentials: 'include',
            body: JSON.stringify([{
              index: 0,
              methodname: 'core_calendar_get_calendar_event_by_id',
              args: { eventid: parseInt(assignment.eventId) }
            }])
          }, 10000);

          if (eventResponse.ok) {
            const eventData = await eventResponse.json();
            if (eventData && eventData[0] && eventData[0].data && eventData[0].data.event) {
              const event = eventData[0].data.event;

              // 補齊課程名稱
              if (event.course && event.course.fullname && (!assignment.course || assignment.course === '')) {
                assignment.course = event.course.fullname;
                console.log(`E3 Helper: 保留的作業 ${assignment.eventId} (${assignment.name}) 補齊課程: ${event.course.fullname}`);
              }

              // 補齊 URL
              if (event.url && (!assignment.url || assignment.url === '')) {
                assignment.url = event.url;
                console.log(`E3 Helper: 保留的作業 ${assignment.eventId} (${assignment.name}) 補齊 URL`);
              }
            }
          }
        } catch (error) {
          console.error(`E3 Helper: 獲取保留作業 ${assignment.eventId} 詳細資訊失敗:`, error);
        }
      }
    }

    // 自動檢測作業繳交狀態
    console.log('E3 Helper: 開始檢測作業繳交狀態...');
    const updatedStatuses = await checkAssignmentSubmissionStatus(assignments, sesskey, statuses);

    // 如果有新的自動檢測狀態，保存到 storage
    if (updatedStatuses) {
      await chrome.storage.local.set({ assignmentStatuses: updatedStatuses });
      console.log('E3 Helper: 已更新自動檢測的繳交狀態到 storage');
    }

    // 檢測新作業並發送通知
    await detectAndNotifyNewAssignments(assignments, oldAssignments);

    // 儲存作業列表
    await chrome.storage.local.set({ assignments: assignments });
    console.log(`E3 Helper: 已同步 ${assignments.length} 個作業（包含 ${keptOldAssignments.length} 個保留的已繳交作業）`);

    return assignments;
  }

  return [];
}

// 檢查作業繳交狀態與評分狀態（使用 HTML 解析）
async function checkAssignmentSubmissionStatus(assignments, sesskey, statuses) {
  let checkedCount = 0;
  let submittedCount = 0;
  let gradedCount = 0;
  let statusUpdated = false;
  const updatedStatuses = { ...statuses }; // 複製一份狀態字典

  // 載入已通知評分的作業列表（避免重複通知）
  const gradingStorage = await chrome.storage.local.get(['gradedNotified']);
  const gradedNotified = new Set(gradingStorage.gradedNotified || []);

  console.log(`E3 Helper: 開始使用 HTML 解析檢測 ${assignments.length} 個作業的繳交與評分狀態...`);

  for (const assignment of assignments) {
    // 跳過手動新增的作業
    if (assignment.isManual || assignment.eventId.startsWith('manual-')) {
      continue;
    }

    // 檢查 URL 有效性
    if (!assignment.url || !assignment.url.includes('mod/assign/view.php')) {
      continue;
    }

    try {
      // 直接訪問作業頁面並解析 HTML
      const htmlResponse = await fetchWithTimeout(assignment.url, {
        method: 'GET',
        credentials: 'include'
      }, 8000);

      if (htmlResponse.ok) {
        const html = await htmlResponse.text();

        // 檢查多個繳交狀態指示器
        const isSubmitted =
          html.includes('submissionstatussubmitted') ||  // CSS class
          html.includes('已繳交') ||  // 中文
          html.includes('已提交供評分') ||  // 中文變體
          html.includes('Submitted for grading') ||  // 英文
          html.includes('修改已繳交的作業') ||  // 按鈕文字
          /class="[^"]*submissionstatus[^"]*submitted[^"]*"/.test(html);  // Regex

        if (isSubmitted && assignment.manualStatus !== 'submitted') {
          // 檢測到已繳交，更新狀態
          assignment.manualStatus = 'submitted';
          assignment.autoDetected = true;
          updatedStatuses[assignment.eventId] = 'submitted';
          statusUpdated = true;
          submittedCount++;
          console.log(`E3 Helper: ✓ 檢測到已繳交 - ${assignment.name}`);
        } else if (!isSubmitted && assignment.autoDetected) {
          // 之前檢測為已繳交，但現在顯示未繳交，保持原狀態（可能是暫時錯誤）
          console.log(`E3 Helper: ${assignment.name} 保持已繳交狀態`);
        }

        // 檢查評分狀態
        const isGraded =
          html.includes('已評分') ||  // 中文
          html.includes('Graded') ||  // 英文
          html.includes('submissiongraded');  // CSS class

        if (isGraded && !gradedNotified.has(assignment.eventId)) {
          gradedNotified.add(assignment.eventId);
          gradedCount++;
          console.log(`E3 Helper: ✓ 偵測到已評分 - ${assignment.name}`);
          await sendGradingNotification(assignment);
        }

        checkedCount++;
      } else {
        console.warn(`E3 Helper: 無法訪問 ${assignment.name} (status: ${htmlResponse.status})`);
      }
    } catch (error) {
      console.error(`E3 Helper: 檢查 ${assignment.name} 失敗:`, error.message);
    }

    // 避免請求過快
    await new Promise(resolve => setTimeout(resolve, 200));
  }

  // 儲存已通知評分的作業列表
  await chrome.storage.local.set({ gradedNotified: [...gradedNotified] });

  console.log(`E3 Helper: 檢測完成 - 已檢查 ${checkedCount} 個作業，其中 ${submittedCount} 個已繳交，${gradedCount} 個新評分`);

  // 如果有更新狀態，返回更新後的字典
  return statusUpdated ? updatedStatuses : null;
}

// 同步課程列表
async function syncCourses() {
  console.log('E3 Helper: 正在同步課程...');

  const sesskey = await getSesskey();
  if (!sesskey) {
    throw new Error('無法取得 sesskey，請重新登入');
  }

  const url = `https://e3p.nycu.edu.tw/lib/ajax/service.php?sesskey=${sesskey}`;

  const response = await fetchWithTimeout(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    credentials: 'include',
    body: JSON.stringify([{
      index: 0,
      methodname: 'core_course_get_enrolled_courses_by_timeline_classification',
      args: {
        offset: 0,
        limit: 0,
        classification: 'inprogress',
        sort: 'fullname'
      }
    }])
  }, 15000); // 增加到 15秒

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  }

  const data = await response.json();
  console.log('E3 Helper: 課程 API 回應:', data);

  if (data && data[0] && data[0].error) {
    console.error('E3 Helper: 課程 API 錯誤:', data[0].error);
    throw new Error(data[0].error);
  }

  if (data && data[0] && data[0].data && data[0].data.courses) {
    const courses = data[0].data.courses;

    // 儲存課程列表
    await chrome.storage.local.set({ courses: courses });
    console.log(`E3 Helper: 已同步 ${courses.length} 個課程`);

    return courses;
  }

  return [];
}

// 檢測新作業並發送通知
async function detectAndNotifyNewAssignments(newAssignments, oldAssignments) {
  try {
    // 獲取已通知的作業列表
    const storage = await chrome.storage.local.get(['notifiedAssignments']);
    const notifiedAssignments = new Set(storage.notifiedAssignments || []);

    // 建立舊作業 ID 集合
    const oldAssignmentIds = new Set(oldAssignments.map(a => a.eventId));

    // 找出真正的新作業（不在舊列表中，且未通知過）
    const newlyAddedAssignments = newAssignments.filter(assignment => {
      return !oldAssignmentIds.has(assignment.eventId) &&
             !notifiedAssignments.has(assignment.eventId);
    });

    if (newlyAddedAssignments.length > 0) {
      console.log(`E3 Helper: 發現 ${newlyAddedAssignments.length} 個新作業`);

      // 為每個新作業發送通知
      for (const assignment of newlyAddedAssignments) {
        await sendAssignmentNotification(assignment);
        notifiedAssignments.add(assignment.eventId);
      }

      // 儲存已通知的作業列表
      await chrome.storage.local.set({
        notifiedAssignments: Array.from(notifiedAssignments)
      });
    }
  } catch (error) {
    console.error('E3 Helper: 檢測新作業時發生錯誤', error);
  }
}

// 發送作業通知
async function sendAssignmentNotification(assignment) {
  try {
    // 計算剩餘時間
    const now = Date.now();
    const deadline = assignment.deadline;
    const timeLeft = deadline - now;
    const daysLeft = Math.floor(timeLeft / (1000 * 60 * 60 * 24));
    const hoursLeft = Math.floor((timeLeft % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));

    let timeText = '';
    if (daysLeft > 0) {
      timeText = `剩餘 ${daysLeft} 天 ${hoursLeft} 小時`;
    } else if (hoursLeft > 0) {
      timeText = `剩餘 ${hoursLeft} 小時`;
    } else if (timeLeft > 0) {
      const minutesLeft = Math.floor(timeLeft / (1000 * 60));
      timeText = `剩餘 ${minutesLeft} 分鐘`;
    } else {
      timeText = '已逾期';
    }

    // 發送桌面通知
    await chrome.notifications.create(`assignment-${assignment.eventId}`, {
      type: 'basic',
      iconUrl: 'chrome-extension://' + chrome.runtime.id + '/128.png',
      title: '📝 新作業上架！',
      message: `${assignment.name}\n📚 課程：${assignment.course}\n⏰ ${timeText}`,
      priority: 2,
      requireInteraction: false
    });

    // 儲存到通知中心
    const storage = await chrome.storage.local.get(['notifications']);
    const notifications = storage.notifications || [];

    const notification = {
      id: `assignment-${assignment.eventId}-${now}`,
      type: 'assignment',
      title: assignment.name,
      message: `📚 課程：${assignment.course}\n⏰ ${timeText}`,
      timestamp: now,
      read: false,
      url: assignment.url
    };

    notifications.unshift(notification); // 新通知放在最前面

    // 只保留最近 50 個通知
    if (notifications.length > 50) {
      notifications.splice(50);
    }

    await chrome.storage.local.set({ notifications });

    // 更新 badge 計數
    if (chrome.action) {
      const unreadCount = notifications.filter(n => !n.read).length;
      if (unreadCount > 0) {
        chrome.action.setBadgeText({ text: unreadCount > 99 ? '99+' : unreadCount.toString() });
        chrome.action.setBadgeBackgroundColor({ color: '#dc3545' });
      }
    }

    console.log(`E3 Helper: 已發送作業通知 - ${assignment.name}`);
  } catch (error) {
    console.error('E3 Helper: 發送通知時發生錯誤', error);
  }
}

// 發送評分通知
async function sendGradingNotification(assignment) {
  try {
    const now = Date.now();

    // 發送桌面通知
    await chrome.notifications.create(`grading-${assignment.eventId}`, {
      type: 'basic',
      iconUrl: 'chrome-extension://' + chrome.runtime.id + '/128.png',
      title: '📊 作業已評分！',
      message: `${assignment.name}\n📚 課程：${assignment.course}`,
      priority: 2,
      requireInteraction: false
    });

    // 儲存到通知中心
    const storage = await chrome.storage.local.get(['notifications']);
    const notifications = storage.notifications || [];

    const notification = {
      id: `grading-${assignment.eventId}-${now}`,
      type: 'grading',
      title: assignment.name,
      message: `📚 課程：${assignment.course}\n📊 作業已評分，點擊查看成績`,
      timestamp: now,
      read: false,
      url: assignment.url
    };

    notifications.unshift(notification);

    if (notifications.length > 50) {
      notifications.splice(50);
    }

    await chrome.storage.local.set({ notifications });

    // 更新 badge 計數
    if (chrome.action) {
      const unreadCount = notifications.filter(n => !n.read).length;
      if (unreadCount > 0) {
        chrome.action.setBadgeText({ text: unreadCount > 99 ? '99+' : unreadCount.toString() });
        chrome.action.setBadgeBackgroundColor({ color: '#dc3545' });
      }
    }

    console.log(`E3 Helper: 已發送評分通知 - ${assignment.name}`);
  } catch (error) {
    console.error('E3 Helper: 發送評分通知時發生錯誤', error);
  }
}

// 監聽通知點擊事件
chrome.notifications.onClicked.addListener((notificationId) => {
  if (notificationId.startsWith('assignment-') || notificationId.startsWith('grading-')) {
    // 提取作業 ID
    const eventId = notificationId.replace('assignment-', '').replace('grading-', '');

    // 獲取作業資料
    chrome.storage.local.get(['assignments'], (result) => {
      const assignments = result.assignments || [];
      const assignment = assignments.find(a => a.eventId === eventId);

      if (assignment && assignment.url) {
        // 開啟作業頁面
        chrome.tabs.create({ url: assignment.url });
      } else {
        // 開啟 E3 首頁
        chrome.tabs.create({ url: 'https://e3p.nycu.edu.tw/' });
      }
    });

    // 清除通知
    chrome.notifications.clear(notificationId);
  }
});

// 監聽來自 content script 的連接請求
chrome.runtime.onConnect.addListener((port) => {
  if (port.name === 'e3-helper') {
    console.log('E3 Helper: Content script 已連接');

    // 發送最後同步時間
    chrome.storage.local.get(['lastSync', 'lastSyncTime'], (result) => {
      port.postMessage({
        type: 'syncStatus',
        data: result
      });
    });
  }
});

// ==================== 課程成員檢測功能 ====================

// 在所有 E3 tabs 中觸發成員檢測
async function checkParticipantsInTabs() {
  console.log('E3 Helper: 開始檢查課程成員變動...');

  try {
    // 查找所有 E3 網站的 tabs
    const tabs = await chrome.tabs.query({
      url: ['https://e3.nycu.edu.tw/*', 'https://e3p.nycu.edu.tw/*']
    });

    if (tabs.length > 0) {
      // 向第一個 E3 tab 發送檢查請求
      const tab = tabs[0];
      console.log(`E3 Helper: 向 tab ${tab.id} 發送成員檢測請求`);

      chrome.tabs.sendMessage(tab.id, {
        action: 'checkParticipants'
      }, (response) => {
        if (chrome.runtime.lastError) {
          console.error('E3 Helper: 無法與 content script 通訊', chrome.runtime.lastError);
        } else {
          console.log('E3 Helper: 成員檢測完成', response);
        }
      });
    } else {
      console.log('E3 Helper: 沒有開啟的 E3 tabs，無法檢測成員變動');
    }
  } catch (error) {
    console.error('E3 Helper: 檢查課程成員時發生錯誤', error);
  }
}

// ==================== AI API 請求處理 ====================

// 帶重試的 fetch 函數（處理 503 等臨時錯誤）
async function fetchWithRetry(url, options, maxRetries = 3) {
  let lastError;
  let lastResponse;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const response = await fetch(url, options);

      // 如果響應成功，直接返回
      if (response.ok) {
        return response;
      }

      // 記錄最後一次響應
      lastResponse = response;

      // 如果是 503 或 429 錯誤且還有重試機會，則重試
      if ((response.status === 503 || response.status === 429) && attempt < maxRetries) {
        const retryDelay = Math.pow(2, attempt) * 1000; // 指數退避：1s, 2s, 4s
        await new Promise(resolve => setTimeout(resolve, retryDelay));
        continue;
      }

      // 對於其他錯誤狀態碼，或最後一次嘗試，返回響應
      return response;
    } catch (error) {
      lastError = error;

      // 網路錯誤也重試
      if (attempt < maxRetries) {
        const retryDelay = Math.pow(2, attempt) * 1000;
        await new Promise(resolve => setTimeout(resolve, retryDelay));
        continue;
      }
    }
  }

  // 所有重試都失敗，返回最後的響應或拋出錯誤
  if (lastResponse) {
    return lastResponse;
  }
  throw lastError || new Error('請求失敗');
}

// 處理 AI API 請求
async function handleAIRequest(request) {
  const { provider, config, prompt } = request;

  switch (provider) {
    case 'ollama':
      return await callOllamaAPI(config, prompt);
    case 'openai':
      return await callOpenAIAPI(config, prompt);
    case 'gemini':
      return await callGeminiAPI(config, prompt);
    case 'custom':
      return await callCustomAPI(config, prompt);
    default:
      throw new Error('未知的 AI 提供商: ' + provider);
  }
}

// 調用 Ollama API
async function callOllamaAPI(config, prompt) {
  const { url, model, temperature } = config;

  try {
    const requestBody = {
      model: model,
      prompt: prompt,
      stream: false
    };

    // 如果提供了 temperature，則添加到請求中
    if (temperature !== undefined) {
      requestBody.temperature = temperature;
    }

    const response = await fetchWithRetry(`${url}/api/generate`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Ollama API 請求失敗: ${response.status} - ${errorText}`);
    }

    const data = await response.json();
    return data.response.trim();
  } catch (error) {
    throw error;
  }
}

// 調用 OpenAI API
async function callOpenAIAPI(config, prompt) {
  const { key, model, temperature } = config;

  try {
    const requestBody = {
      model: model,
      messages: [{
        role: 'user',
        content: prompt
      }]
    };

    // 如果提供了 temperature，則添加到請求中，否則使用默認值 0.3
    requestBody.temperature = temperature !== undefined ? temperature : 0.3;

    const response = await fetchWithRetry('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${key}`
      },
      body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
      throw new Error(`OpenAI API 請求失敗: ${response.status}`);
    }

    const data = await response.json();
    return data.choices[0].message.content.trim();
  } catch (error) {
    throw error;
  }
}

// 調用 Gemini API
async function callGeminiAPI(config, prompt) {
  const { key, model, temperature, thinkingBudget } = config;

  try {
    const generationConfig = {
      temperature: temperature !== undefined ? temperature : 0.3,
      candidateCount: 1
    };

    if (thinkingBudget !== undefined) {
      generationConfig.thinkingConfig = {
        thinkingBudget: thinkingBudget
      };
    }

    const requestBody = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }],
      generationConfig: generationConfig
    };

    const response = await fetchWithRetry(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Gemini API 請求失敗: ${response.status} - ${errorText}`);
    }

    const data = await response.json();

    // 檢查響應結構
    if (!data.candidates || data.candidates.length === 0) {
      throw new Error('Gemini API 沒有返回候選結果');
    }

    const candidate = data.candidates[0];

    if (candidate.content && candidate.content.parts && candidate.content.parts[0] && candidate.content.parts[0].text) {
      return candidate.content.parts[0].text.trim();
    } else if (candidate.text) {
      return candidate.text.trim();
    } else if (candidate.output) {
      return candidate.output.trim();
    } else {
      if (candidate.finishReason === 'MAX_TOKENS') {
        throw new Error('Gemini MAX_TOKENS 錯誤且未返回任何文本，可能是輸入 prompt 太長。');
      }
      throw new Error('無法解析 Gemini 響應結構: ' + JSON.stringify(candidate));
    }
  } catch (error) {
    throw error;
  }
}

// 調用自定義 API
async function callCustomAPI(config, prompt) {
  const { url, key, model, temperature } = config;

  try {
    const headers = {
      'Content-Type': 'application/json'
    };

    if (key) {
      headers['Authorization'] = `Bearer ${key}`;
    }

    const requestBody = {
      model: model,
      messages: [{
        role: 'user',
        content: prompt
      }],
      temperature: temperature !== undefined ? temperature : 0.3
    };

    const response = await fetchWithRetry(url, {
      method: 'POST',
      headers,
      body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
      throw new Error(`自定義 API 請求失敗: ${response.status}`);
    }

    const data = await response.json();
    return data.choices[0].message.content.trim();
  } catch (error) {
    throw error;
  }
}

// ==================== 抓取內容功能 ====================

// 從 E3 抓取公告/信件內容
async function fetchContentFromE3(url) {
  console.log(`E3 Helper: 開始抓取內容 - ${url}`);

  // 驗證 URL 域名
  const fetchAllowedDomains = ['e3.nycu.edu.tw', 'e3p.nycu.edu.tw'];
  try {
    const urlObj = new URL(url);
    if (!fetchAllowedDomains.some(d => urlObj.hostname.endsWith(d))) {
      throw new Error('URL not from allowed E3 domain');
    }
  } catch (e) { throw e; }

  try {
    const response = await fetchWithTimeout(url, {
      method: 'GET',
      credentials: 'include'
    }, 15000); // 15 秒超時

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const html = await response.text();
    console.log(`E3 Helper: 內容抓取成功，長度: ${html.length}`);
    return html;
  } catch (error) {
    console.error('E3 Helper: 抓取內容失敗', error);
    throw new Error(`無法抓取內容: ${error.message}`);
  }
}

// ==================== 載入公告和信件功能 ====================

// 定時靜默同步公告和信件（不打開新標籤頁，只在有 E3 標籤頁時同步）
async function syncAnnouncementsAndMessagesSilently() {
  console.log('E3 Helper: 開始靜默同步公告和信件...', new Date().toLocaleTimeString());

  try {
    // 只查找已開啟的 E3 標籤頁，不主動開啟新的
    const tabs = await chrome.tabs.query({
      url: ['https://e3.nycu.edu.tw/*', 'https://e3p.nycu.edu.tw/*']
    });

    if (tabs.length === 0) {
      console.log('E3 Helper: 沒有開啟的 E3 標籤頁，跳過公告/信件同步');
      return { success: false, reason: 'no_e3_tab' };
    }

    // 使用第一個 E3 標籤頁
    const tab = tabs[0];
    console.log(`E3 Helper: 使用標籤頁 ${tab.id} 同步公告和信件`);

    return new Promise((resolve) => {
      chrome.tabs.sendMessage(tab.id, {
        action: 'loadAnnouncementsAndMessagesInTab'
      }, (response) => {
        if (chrome.runtime.lastError) {
          console.error('E3 Helper: 靜默同步失敗', chrome.runtime.lastError);
          resolve({ success: false, error: chrome.runtime.lastError.message });
        } else if (response && response.success) {
          console.log('E3 Helper: 公告和信件靜默同步完成');
          // 更新 badge
          updateBadgeFromStorage();
          resolve({ success: true });
        } else {
          console.log('E3 Helper: 靜默同步未完成', response);
          resolve({ success: false, error: response?.error });
        }
      });
    });
  } catch (error) {
    console.error('E3 Helper: 靜默同步公告和信件時發生錯誤', error);
    return { success: false, error: error.message };
  }
}

// 在背景載入公告和信件（通過 E3 標籤頁）
async function loadAnnouncementsAndMessagesInBackground() {
  console.log('E3 Helper: 開始在背景載入公告和信件...');

  try {
    // 查找所有 E3 網站的標籤頁
    const tabs = await chrome.tabs.query({
      url: ['https://e3.nycu.edu.tw/*', 'https://e3p.nycu.edu.tw/*']
    });

    if (tabs.length > 0) {
      // 使用第一個 E3 標籤頁來載入資料
      const tab = tabs[0];
      console.log(`E3 Helper: 使用標籤頁 ${tab.id} 載入資料`);

      // 向該標籤頁發送載入請求
      return new Promise((resolve, reject) => {
        chrome.tabs.sendMessage(tab.id, {
          action: 'loadAnnouncementsAndMessagesInTab'
        }, (response) => {
          if (chrome.runtime.lastError) {
            console.error('E3 Helper: 無法與 content script 通訊', chrome.runtime.lastError);
            reject(new Error('無法與 E3 標籤頁通訊'));
          } else if (response && response.success) {
            console.log('E3 Helper: 資料載入完成');
            resolve({ success: true, message: '資料已在背景載入完成' });
          } else {
            console.error('E3 Helper: 載入失敗', response);
            reject(new Error(response?.error || '載入失敗'));
          }
        });
      });
    } else {
      // 沒有打開的 E3 標籤頁，打開一個新的
      console.log('E3 Helper: 沒有開啟的 E3 標籤頁，將打開新標籤頁');

      const newTab = await chrome.tabs.create({
        url: 'https://e3p.nycu.edu.tw/',
        active: false // 在背景開啟
      });

      // 等待標籤頁載入完成
      return new Promise((resolve, reject) => {
        const timeoutId = setTimeout(() => {
          reject(new Error('載入超時，請確認已登入 E3'));
        }, 30000); // 30 秒超時

        // 監聽標籤頁載入完成
        const listener = (tabId, changeInfo, tab) => {
          if (tabId === newTab.id && changeInfo.status === 'complete') {
            chrome.tabs.onUpdated.removeListener(listener);

            // 延遲一下確保 content script 已載入
            setTimeout(() => {
              chrome.tabs.sendMessage(newTab.id, {
                action: 'loadAnnouncementsAndMessagesInTab'
              }, (response) => {
                clearTimeout(timeoutId);

                if (chrome.runtime.lastError) {
                  console.error('E3 Helper: 無法與新標籤頁通訊', chrome.runtime.lastError);
                  reject(new Error('無法與 E3 標籤頁通訊'));
                } else if (response && response.success) {
                  console.log('E3 Helper: 資料載入完成（新標籤頁）');
                  // 關閉新開的標籤頁
                  chrome.tabs.remove(newTab.id);
                  resolve({ success: true, message: '資料已在背景載入完成' });
                } else {
                  console.error('E3 Helper: 載入失敗（新標籤頁）', response);
                  reject(new Error(response?.error || '載入失敗'));
                }
              });
            }, 1000); // 延遲 1 秒
          }
        };

        chrome.tabs.onUpdated.addListener(listener);
      });
    }
  } catch (error) {
    console.error('E3 Helper: 載入公告和信件時發生錯誤', error);
    throw error;
  }
}
