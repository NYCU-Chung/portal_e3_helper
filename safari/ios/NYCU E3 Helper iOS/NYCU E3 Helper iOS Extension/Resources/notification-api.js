// Callback APIs keep notification delivery compatible with Chrome 93–94.
(function () {
  function call(target, method, ...args) {
    return new Promise((resolve, reject) => {
      const result = target[method](...args, value => {
        const error = chrome.runtime?.lastError;
        if (error) reject(new Error(error.message));
        else resolve(value);
      });
      if (result && typeof result.then === 'function') result.then(resolve, reject);
    });
  }
  globalThis.E3NotificationAPI = {
    get: keys => call(chrome.storage.local, 'get', keys),
    set: data => call(chrome.storage.local, 'set', data),
    create: (id, options) => call(chrome.notifications, 'create', id, options),
    permission: () => call(chrome.notifications, 'getPermissionLevel'),
    clear: id => call(chrome.notifications, 'clear', id),
    openTab: url => call(chrome.tabs, 'create', { url })
  };
})();
