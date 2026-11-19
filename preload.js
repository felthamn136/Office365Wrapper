'use strict';

const { contextBridge, ipcRenderer } = require('electron');

if (location.protocol === 'file:' && location.pathname.endsWith('/dashboard.html')) {
  contextBridge.exposeInMainWorld('graphSearch', {
    search: (query) => ipcRenderer.invoke('graph:search', query),
    signOut: () => ipcRenderer.invoke('graph:sign-out'),
  });
}
