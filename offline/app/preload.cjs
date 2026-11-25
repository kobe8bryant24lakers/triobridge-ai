const { contextBridge, ipcRenderer } = require('electron');
const host = process.argv
  .find((value) => /^--trio-host=(word|excel|powerpoint)$/.test(value))
  ?.split('=')[1];
if (host)
  contextBridge.exposeInMainWorld(
    'trioWorkbench',
    Object.freeze({
      host,
      model: (id, action, payload) => ipcRenderer.invoke('trio:model', id, action, payload),
      cancel: (id) => ipcRenderer.send('trio:cancel', id),
      office: (action, data) => ipcRenderer.invoke('trio:office', action, data),
    }),
  );
