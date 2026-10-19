// Gives the API key window one thing it can ask the app to do: save (and check) the key.
const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("sheetcraft", {
  saveKey: (key) => ipcRenderer.invoke("save-key", key),
});
