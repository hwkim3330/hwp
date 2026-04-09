const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("desktopMonitor", {
  onStats(callback) {
    ipcRenderer.on("system-stats", (_event, payload) => callback(payload));
  },
  onCursorOverlayState(callback) {
    ipcRenderer.on("cursor-overlay-state", (_event, payload) => callback(payload));
  },
  onCursorOverlayTick(callback) {
    ipcRenderer.on("cursor-overlay-tick", (_event, payload) => callback(payload));
  },
  openVisionResource(resource) {
    return ipcRenderer.invoke("vision:open-resource", resource);
  },
  toggleCursorOverlay(forceState) {
    return ipcRenderer.invoke("cursor-overlay:toggle", forceState);
  },
  getCursorOverlayState() {
    return ipcRenderer.invoke("cursor-overlay:state");
  },
  runOperatorPreset(presetId) {
    return ipcRenderer.invoke("operator:run-preset", presetId);
  },
  captureScreenshot() {
    return ipcRenderer.invoke("capture:screenshot");
  },
  getPermissionStatus() {
    return ipcRenderer.invoke("permissions:get-status");
  },
  requestPermission(kind) {
    return ipcRenderer.invoke("permissions:request", kind);
  },
  playCue(kind) {
    return ipcRenderer.invoke("sound:play-cue", kind);
  },
});
