// Browser notifications on Chrome; native macOS notifications in the Safari package.
(function () {
  const native = !chrome.notifications?.create && chrome.runtime.getURL('').startsWith('safari-web-extension:') && Boolean(chrome.runtime.sendNativeMessage) && Boolean(chrome.runtime.getManifest?.().permissions?.includes('nativeMessaging'));
  async function request(action, payload = {}) {
    const response = await chrome.runtime.sendNativeMessage('com.yoyo1112.nycu-e3-helper', { action, ...payload });
    if (!response?.success) throw new Error(response?.error || 'Native notification service unavailable');
    return response;
  }
  globalThis.E3DesktopNotifications = {
    supported: Boolean(chrome.notifications?.create) || native,
    native,
    async getPermissionLevel() {
      if (native) return (await request('notificationStatus')).permission;
      return chrome.notifications ? E3NotificationAPI.permission() : 'denied';
    },
    async authorize() {
      return native ? (await request('authorizeNotifications')).permission : this.getPermissionLevel();
    },
    async create(id, options, url) {
      if (!this.supported) throw new Error('Desktop notification service unavailable');
      if (native) return request('showNotification', { id, title: options.title, message: options.message, url: url || '' });
      return E3NotificationAPI.create(id, options);
    }
  };
})();
