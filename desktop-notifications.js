// Browser notification transport. Platform adapters can extend this interface.
(function () {
  globalThis.E3DesktopNotifications = {
    supported: Boolean(chrome.notifications?.create),
    native: false,
    async getPermissionLevel() {
      return chrome.notifications ? E3NotificationAPI.permission() : 'denied';
    },
    async authorize() {
      return this.getPermissionLevel();
    },
    async create(id, options, url) {
      if (!this.supported) throw new Error('Desktop notification service unavailable');
      return E3NotificationAPI.create(id, options);
    }
  };
})();
