'use strict';

const { app, BrowserWindow, dialog, ipcMain, Menu, session, shell } = require('electron');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { destinationFromArgs, destinations } = require('./destinations');
const { configureNotifications } = require('./notifications');
const { createGraphSearch } = require('./graph-search');
const { configureWebContents } = require('./window-policy');

async function start() {
  const destination = destinationFromArgs(process.argv.slice(app.isPackaged ? 1 : 2));
  await app.whenReady();

  if (process.platform === 'win32' && !app.isPackaged) {
    app.setAppUserModelId(process.execPath);
  }
  let notificationsEnabled = false;
  configureNotifications(session.defaultSession, () => notificationsEnabled);

  const window = new BrowserWindow({
    width: 1200,
    height: 800,
    title: destination.title,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: true,
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  const graphSearch = createGraphSearch({
    clientId: process.env.OFFICE365_CLIENT_ID,
    openBrowser: (url) => shell.openExternal(url),
  });
  const dashboardUrl = pathToFileURL(path.join(__dirname, 'dashboard.html')).href;
  function requireDashboard(event) {
    if (event.sender !== window.webContents ||
        event.senderFrame !== window.webContents.mainFrame ||
        event.senderFrame.url !== dashboardUrl) {
      throw new Error('Graph access is available only from the Office Dashboard.');
    }
  }
  ipcMain.handle('graph:search', (event, query) => {
    requireDashboard(event);
    return graphSearch.search(query);
  });
  ipcMain.handle('graph:sign-out', (event) => {
    requireDashboard(event);
    return graphSearch.signOut();
  });

  configureWebContents(window.webContents);

  let navigationId = 0;
  const navigate = async (target) => {
    const currentNavigation = ++navigationId;
    window.setTitle(target.title);
    try {
      if (target.file) {
        await window.loadFile(path.join(__dirname, target.file));
      } else {
        await window.loadURL(target.url);
      }
    } catch (error) {
      if (currentNavigation === navigationId && !window.isDestroyed()) {
        console.error(error);
        dialog.showErrorBox('Navigation error', `Could not open ${target.title}: ${error.message}`);
      }
    }
  };

  const template = [
    ...(process.platform === 'darwin' ? [{ role: 'appMenu' }] : []),
    {
      label: 'Navigate',
      submenu: [
        { label: 'Office Dashboard', click: () => navigate(destinations.dashboard) },
        { label: 'Microsoft 365', click: () => navigate(destinations.microsoft365) },
        { label: 'Outlook (work or school)', click: () => navigate(destinations.outlook) },
        { label: 'Outlook (personal)', click: () => navigate(destinations.outlookPersonal) },
        { type: 'separator' },
        { role: 'reload' },
      ],
    },
    {
      label: 'Notifications',
      submenu: [{
        label: 'Enable Microsoft website notifications for this session',
        type: 'checkbox',
        checked: false,
        click: (item) => { notificationsEnabled = item.checked; },
      }],
    },
    { role: 'editMenu' },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
  await navigate(destination);
}

app.on('window-all-closed', () => app.quit());

start().catch((error) => {
  console.error(error);
  dialog.showErrorBox('Office365Wrapper error', error.message);
  app.exit(1);
});
