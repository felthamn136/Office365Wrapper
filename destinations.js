'use strict';

const destinations = {
  dashboard: { title: 'Dashboard', file: 'dashboard.html' },
  microsoft365: { title: 'Microsoft 365', url: 'https://www.microsoft365.com/' },
  outlook: { title: 'Outlook', url: 'https://outlook.office365.com/' },
  outlookPersonal: { title: 'Outlook', url: 'https://outlook.live.com/mail/' },
};

function destinationFromArgs(args) {
  if (args.length > 1) {
    throw new Error('Only one destination can be specified');
  }

  switch (args[0]) {
    case undefined:
      return destinations.dashboard;
    case '--microsoft365':
      return destinations.microsoft365;
    case '--outlook':
      return destinations.outlook;
    case '--outlook-personal':
      return destinations.outlookPersonal;
    default:
      throw new Error(`Unknown argument: ${args[0]}`);
  }
}

module.exports = { destinationFromArgs, destinations };
