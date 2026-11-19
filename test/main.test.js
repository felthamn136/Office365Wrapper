'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { test } = require('node:test');
const vm = require('node:vm');
const { destinationFromArgs, destinations } = require('../destinations');
const { configureNotifications } = require('../notifications');
const { configureWebContents } = require('../window-policy');

test('opens Outlook in an isolated window and restricts navigation', async () => {
  const handlers = {};
  const sessionHandlers = {};
  const ipcHandlers = {};
  let window;
  let menu;
  class BrowserWindow {
    constructor(options) {
      this.options = options;
      this.webContents = {
        mainFrame: { url: pathToFileURL(path.join(__dirname, '../dashboard.html')).href },
        getURL: () => this.url || this.webContents.mainFrame.url,
        setWindowOpenHandler: (handler) => { handlers.open = handler; },
        on: (event, handler) => { handlers[event] = handler; },
      };
      window = this;
    }

    async loadURL(url) {
      this.url = url;
      if (this.failLoad) {
        throw new Error('network unavailable');
      }
    }

    async loadFile(file) {
      this.file = file;
    }

    setTitle(title) {
      this.title = title;
    }

    isDestroyed() {
      return false;
    }
  }

  let reportedError;
  const electron = {
    app: { isPackaged: false, whenReady: async () => {}, on: () => {}, quit: () => {} },
    BrowserWindow,
    ipcMain: { handle: (name, handler) => { ipcHandlers[name] = handler; } },
    shell: { openExternal: async () => {} },
    dialog: { showErrorBox: (title, message) => { reportedError = { title, message }; } },
    Menu: {
      buildFromTemplate: (template) => template,
      setApplicationMenu: (template) => { menu = template; },
    },
    session: {
      defaultSession: {
        setPermissionRequestHandler: (handler) => { sessionHandlers.request = handler; },
        setPermissionCheckHandler: (handler) => { sessionHandlers.check = handler; },
      },
    },
  };
  const source = fs.readFileSync(path.join(__dirname, '../main.js'), 'utf8');
  vm.runInNewContext(source, {
    require: (name) => {
      if (name === 'electron') return electron;
      if (name === './notifications') return { configureNotifications };
      if (name === './window-policy') return { configureWebContents };
      if (name === 'node:path') return path;
      if (name === 'node:url') return { pathToFileURL };
      if (name === './graph-search') return { createGraphSearch: () => ({
        search: async (value) => value,
        signOut: async () => {},
      }) };
      return { destinationFromArgs, destinations };
    },
    __dirname: path.join(__dirname, '..'),
    process: { argv: ['electron', '.', '--outlook'], platform: 'linux', env: {} },
    console: { error: () => {} },
  });
  await new Promise(setImmediate);

  assert.equal(window.url, 'https://outlook.office365.com/');
  assert.equal(window.options.webPreferences.nodeIntegration, false);
  assert.equal(window.options.webPreferences.contextIsolation, true);
  assert.equal(window.options.webPreferences.sandbox, true);
  const dashboardEvent = { sender: window.webContents, senderFrame: window.webContents.mainFrame };
  assert.equal(await ipcHandlers['graph:search'](dashboardEvent, 'report'), 'report');
  assert.throws(() => ipcHandlers['graph:search']({
    sender: window.webContents,
    senderFrame: { url: 'https://outlook.office.com/' },
  }, 'report'), /only from the Office Dashboard/);
  assert.equal(handlers.open({ url: 'https://word.office.com/' }).action, 'allow');
  let prevented = false;
  handlers['will-navigate']({ preventDefault: () => { prevented = true; } }, 'file:///tmp/test');
  assert.equal(prevented, true);
  prevented = false;
  handlers['will-navigate']({ preventDefault: () => { prevented = true; } }, 'https://login.microsoftonline.com/');
  assert.equal(prevented, false);
  sessionHandlers.request(null, 'media', (allowed) => assert.equal(allowed, false), {});
  assert.equal(sessionHandlers.check(null, 'media', 'https://outlook.office.com', {}), false);
  assert.equal(reportedError, undefined);

  const notificationToggle = menu.find((item) => item.label === 'Notifications').submenu[0];
  const notificationDetails = { isMainFrame: true, requestingUrl: 'https://outlook.office.com/' };
  assert.equal(sessionHandlers.check(null, 'notifications', 'https://outlook.office.com', notificationDetails), false);
  notificationToggle.click({ checked: true });
  assert.equal(sessionHandlers.check(null, 'notifications', 'https://outlook.office.com', notificationDetails), true);
  notificationToggle.click({ checked: false });
  assert.equal(sessionHandlers.check(null, 'notifications', 'https://outlook.office.com', notificationDetails), false);

  const navigation = menu.find((item) => item.label === 'Navigate').submenu;
  await navigation.find((item) => item.label === 'Office Dashboard').click();
  assert.equal(window.file, path.join(__dirname, '../dashboard.html'));
  assert.equal(window.title, 'Office Dashboard');
  await navigation.find((item) => item.label === 'Outlook (personal)').click();
  assert.equal(window.url, 'https://outlook.live.com/mail/');
  assert.equal(window.title, 'Outlook');
  await navigation.find((item) => item.label === 'Microsoft 365').click();
  assert.equal(window.url, 'https://www.microsoft365.com/');
  await navigation.find((item) => item.label === 'Outlook (work or school)').click();
  assert.equal(window.url, 'https://outlook.office365.com/');
  assert.equal(navigation.some((item) => item.label === 'Microsoft Teams'), false);

  let rejectOldNavigation;
  const loadURL = window.loadURL.bind(window);
  window.loadURL = () => new Promise((_resolve, reject) => { rejectOldNavigation = reject; });
  const oldNavigation = navigation.find((item) => item.label === 'Microsoft 365').click();
  window.loadURL = loadURL;
  await navigation.find((item) => item.label === 'Outlook (work or school)').click();
  rejectOldNavigation(new Error('navigation superseded'));
  await oldNavigation;
  assert.equal(reportedError, undefined);

  window.failLoad = true;
  await navigation.find((item) => item.label === 'Outlook (personal)').click();
  assert.equal(reportedError.title, 'Navigation error');
  assert.match(reportedError.message, /network unavailable/);
});

test('permits Microsoft launcher windows with secure navigation and nested restrictions', () => {
  const createContents = (url) => {
    const handlers = {};
    return {
      url,
      handlers,
      getURL() { return this.url; },
      setWindowOpenHandler(handler) { handlers.open = handler; },
      on(event, handler) { handlers[event] = handler; },
    };
  };
  const contents = createContents('https://www.microsoft365.com/');
  configureWebContents(contents);
  const open = contents.handlers.open({ url: 'https://tenant.sharepoint.com/' });
  assert.equal(open.action, 'allow');
  assert.deepEqual(open.overrideBrowserWindowOptions.webPreferences, {
    nodeIntegration: false, contextIsolation: true, sandbox: true,
  });
  assert.equal(contents.handlers.open({ url: 'about:blank' }).action, 'allow');
  for (const url of ['file:///tmp/test', 'http://example.com/', 'javascript:alert(1)', 'https://user@example.com/']) {
    assert.equal(contents.handlers.open({ url }).action, 'deny');
  }
  for (const url of ['https://microsoft365.com.example.com/', 'file:///dashboard.html', 'https://example.com/']) {
    contents.url = url;
    assert.equal(contents.handlers.open({ url: 'https://example.com/' }).action, 'deny');
  }
  const child = createContents('https://tenant.sharepoint.com/');
  contents.handlers['did-create-window']({ webContents: child });
  assert.equal(child.handlers.open({ url: 'https://word.office.com/' }).action, 'allow');
  let prevented = false;
  child.handlers['will-redirect']({ preventDefault: () => { prevented = true; } }, 'file:///tmp/test');
  assert.equal(prevented, true);
  prevented = false;
  child.handlers['will-navigate']({ preventDefault: () => { prevented = true; } }, 'https://login.microsoftonline.com/');
  assert.equal(prevented, false);
});
