const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("pigeonApi", {
  request(command, params = {}) {
    return ipcRenderer.invoke("backend:request", command, params);
  },
  chooseImage() {
    return ipcRenderer.invoke("dialog:open-image");
  },
  choosePngPath() {
    return ipcRenderer.invoke("dialog:save-png");
  },
});
