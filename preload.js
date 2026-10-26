/**
 * =====================================================================
 * SENTINEL STUDIO™ — SECURE CONTEXT BRIDGE
 * Preload Script exposing window.sentinelDesktop
 * =====================================================================
 */

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('sentinelDesktop', {
  selectDirectory: () => ipcRenderer.invoke('dialog:select-directory'),
  readFile: (filePath) => ipcRenderer.invoke('fs:read-file', filePath),
  writeFile: (filePath, content) => ipcRenderer.invoke('fs:write-file', filePath, content),
  saveBinary: (filePath, arrayBuffer) => ipcRenderer.invoke('fs:save-binary', filePath, arrayBuffer),
  getVersion: () => ipcRenderer.invoke('app:get-version'),
  isDesktop: true
});
