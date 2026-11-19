'use strict';

const form = document.getElementById('search-form');
const query = document.getElementById('search-query');
const submit = document.getElementById('search-submit');
const signOut = document.getElementById('search-signout');
const status = document.getElementById('search-status');
const results = document.getElementById('search-results');
const graphSearch = window.graphSearch || (window.__TAURI__ && {
  search: (value) => window.__TAURI__.core.invoke('graph_search', { query: value }),
  signOut: () => window.__TAURI__.core.invoke('graph_sign_out'),
});

function setBusy(busy) {
  submit.disabled = busy;
  signOut.disabled = busy;
  query.disabled = busy;
}

function showResults(groups) {
  results.replaceChildren();
  for (const group of groups) {
    const section = document.createElement('section');
    const heading = document.createElement('h3');
    heading.textContent = group.label;
    section.append(heading);
    if (group.error) {
      const error = document.createElement('p');
      error.textContent = group.error;
      error.className = 'search-error';
      section.append(error);
    } else if (group.results.length === 0) {
      const empty = document.createElement('p');
      empty.textContent = 'No results.';
      section.append(empty);
    } else {
      const list = document.createElement('ul');
      for (const item of group.results) {
        const row = document.createElement('li');
        if (item.url) {
          const link = document.createElement('a');
          link.href = item.url;
          link.textContent = item.title;
          row.append(link);
        } else {
          row.textContent = `${item.title} (open the relevant app to view)`;
        }
        list.append(row);
      }
      section.append(list);
      if (group.more) {
        const more = document.createElement('p');
        more.textContent = 'More results are available in the Microsoft app.';
        section.append(more);
      }
    }
    results.append(section);
  }
}

if (!graphSearch) {
  submit.disabled = true;
  signOut.disabled = true;
  status.textContent = 'Graph search is available in the desktop app.';
} else {
  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    if (!query.value.trim()) {
      status.textContent = 'Enter a search term.';
      return;
    }
    setBusy(true);
    results.replaceChildren();
    status.textContent = 'Signing in if needed and searching...';
    try {
      const groups = await graphSearch.search(query.value);
      showResults(groups);
      status.textContent = 'Search complete. Up to 25 results per category are shown.';
    } catch (error) {
      status.textContent = `Search failed: ${error.message || error}`;
    } finally {
      setBusy(false);
    }
  });

  signOut.addEventListener('click', async () => {
    setBusy(true);
    try {
      await graphSearch.signOut();
      results.replaceChildren();
      status.textContent = 'Graph session cleared. Website and browser sign-in are unchanged.';
    } catch (error) {
      status.textContent = `Sign-out failed: ${error.message || error}`;
    } finally {
      setBusy(false);
    }
  });
}
