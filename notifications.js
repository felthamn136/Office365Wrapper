'use strict';

const allowedOrigins = new Set([
  'https://www.microsoft365.com',
  'https://m365.cloud.microsoft',
  'https://outlook.office.com',
  'https://outlook.office365.com',
  'https://outlook.cloud.microsoft',
  'https://outlook.live.com',
]);

function isAllowedOrigin(value) {
  if (typeof value !== 'string') return false;
  const url = URL.parse(value);
  return url !== null && allowedOrigins.has(url.origin);
}

function configureNotifications(session, isEnabled) {
  session.setPermissionRequestHandler((_webContents, permission, callback, details) => {
    callback(
      isEnabled() &&
      permission === 'notifications' &&
      details.isMainFrame &&
      isAllowedOrigin(details.requestingUrl),
    );
  });
  session.setPermissionCheckHandler((_webContents, permission, origin, details) => {
    return isEnabled() &&
      permission === 'notifications' &&
      isAllowedOrigin(origin) &&
      (!details.requestingUrl || details.isMainFrame) &&
      (!details.embeddingOrigin || isAllowedOrigin(details.embeddingOrigin));
  });
}

module.exports = { configureNotifications };
