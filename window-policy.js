'use strict';

const microsoftDomains = [
  'microsoft.com', 'microsoft365.com', 'office.com', 'office365.com',
  'cloud.microsoft', 'live.com', 'microsoftonline.com', 'onedrive.com', 'sharepoint.com',
];

function isMicrosoftPage(value) {
  const url = URL.parse(value);
  return Boolean(url && url.protocol === 'https:' && !url.port && !url.username && !url.password &&
    microsoftDomains.some((domain) => url.hostname === domain || url.hostname.endsWith(`.${domain}`)));
}

function isSecurePage(value) {
  const url = URL.parse(value);
  return Boolean(url && url.protocol === 'https:' && !url.username && !url.password);
}

function configureWebContents(contents) {
  contents.setWindowOpenHandler(({ url }) => {
    if (!isMicrosoftPage(contents.getURL()) || (url !== 'about:blank' && !isSecurePage(url))) {
      return { action: 'deny' };
    }
    return {
      action: 'allow',
      overrideBrowserWindowOptions: {
        width: 1200,
        height: 800,
        webPreferences: { nodeIntegration: false, contextIsolation: true, sandbox: true },
      },
    };
  });
  const restrictNavigation = (event, url) => {
    if (!isSecurePage(url)) event.preventDefault();
  };
  contents.on('will-navigate', restrictNavigation);
  contents.on('will-redirect', restrictNavigation);
  contents.on('did-create-window', (child) => configureWebContents(child.webContents));
}

module.exports = { configureWebContents };
