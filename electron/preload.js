const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("desktopMonitor", {
  onStats(callback) {
    ipcRenderer.on("system-stats", (_event, payload) => callback(payload));
  },
});
