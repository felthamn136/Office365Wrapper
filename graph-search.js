'use strict';

const { PublicClientApplication, InteractionRequiredAuthError } = require('@azure/msal-node');

const scopes = ['Mail.Read', 'Calendars.Read', 'Files.Read.All'];
const categories = [
  { type: 'message', label: 'Mail' },
  { type: 'event', label: 'Calendar' },
  { type: 'driveItem', label: 'Files' },
];

function createGraphSearch({ clientId, openBrowser, fetchImpl = fetch }) {
  let client;
  let account;
  let tokenPromise;

  function getClient() {
    if (!client) {
      if (typeof clientId !== 'string' || !/^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(clientId)) {
        throw new Error('Set OFFICE365_CLIENT_ID to the application (client) ID of your Microsoft Entra app registration.');
      }
      client = new PublicClientApplication({
        auth: { clientId, authority: 'https://login.microsoftonline.com/organizations' },
      });
    }
    return client;
  }

  async function getToken() {
    if (tokenPromise) return tokenPromise;
    tokenPromise = (async () => {
      const application = getClient();
      if (account) {
        try {
          return await application.acquireTokenSilent({ account, scopes });
        } catch (error) {
          if (!(error instanceof InteractionRequiredAuthError)) throw error;
        }
      }
      const response = await application.acquireTokenInteractive({
        scopes,
        openBrowser: async (url) => {
          const parsed = URL.parse(url);
          if (!parsed || parsed.origin !== 'https://login.microsoftonline.com' || parsed.username || parsed.password) {
            throw new Error('Microsoft sign-in returned an unexpected authorization URL.');
          }
          await openBrowser(url);
        },
      });
      if (!response || !response.account || !response.accessToken) {
        throw new Error('Microsoft sign-in did not return an account and access token.');
      }
      account = response.account;
      return response;
    })();
    try {
      return await tokenPromise;
    } finally {
      tokenPromise = null;
    }
  }

  async function searchCategory(category, query, accessToken) {
    try {
      const response = await fetchImpl('https://graph.microsoft.com/v1.0/search/query', {
        method: 'POST',
        headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          requests: [{ entityTypes: [category.type], query: { queryString: query }, from: 0, size: 25 }],
        }),
        signal: AbortSignal.timeout(30000),
      });
      if (!response.ok) throw new Error(`Microsoft Graph returned HTTP ${response.status}.`);
      const data = await response.json();
      if (!Array.isArray(data.value) || !Array.isArray(data.value[0]?.hitsContainers)) {
        throw new Error('Microsoft Graph returned an unexpected search response.');
      }
      const containers = data.value[0].hitsContainers;
      if (containers.some((container) => !Array.isArray(container.hits))) {
        throw new Error('Microsoft Graph returned an unexpected search results container.');
      }
      const results = containers.flatMap((container) => container.hits.map((hit) => {
        const resource = hit.resource || {};
        const title = category.type === 'driveItem' ? resource.name : resource.subject;
        const url = URL.parse(resource.webUrl || resource.webLink || '');
        return {
          title: typeof title === 'string' && title.trim() ? title : '(Untitled)',
          url: url && url.protocol === 'https:' && !url.username && !url.password ? url.href : null,
        };
      }));
      return { label: category.label, results, more: containers.some((container) => container.moreResultsAvailable === true) };
    } catch (error) {
      return { label: category.label, error: error.message || 'Search failed.' };
    }
  }

  return {
    async search(query) {
      if (typeof query !== 'string' || !query.trim() || query.trim().length > 200) {
        throw new Error('Enter a search term between 1 and 200 characters.');
      }
      const token = await getToken();
      return Promise.all(categories.map((category) => searchCategory(category, query.trim(), token.accessToken)));
    },
    async signOut() {
      if (tokenPromise) throw new Error('Wait for sign-in to finish before signing out.');
      if (client && account) await client.getTokenCache().removeAccount(account);
      account = null;
    },
  };
}

module.exports = { createGraphSearch };
