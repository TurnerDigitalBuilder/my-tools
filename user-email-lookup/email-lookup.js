const EmailLookupApp = (() => {
  const GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/users';
  const MAX_RESULTS = 10;

  const state = {
    token: '',
    isRunning: false,
    results: [],
    analysis: { names: [], rawCount: 0, uniqueCount: 0, skippedDuplicates: 0 }
  };

  const elements = {};

  function cacheElements() {
    elements.namesInput = document.getElementById('namesInput');
    elements.lookupButton = document.getElementById('lookupButton');
    elements.clearButton = document.getElementById('clearButton');
    elements.copyButton = document.getElementById('copyButton');
    elements.downloadButton = document.getElementById('downloadButton');
    elements.lookupStatus = document.getElementById('lookupStatus');
    elements.namesCount = document.getElementById('namesCount');
    elements.resultsBody = document.getElementById('resultsBody');
    elements.resultsSummary = document.getElementById('resultsSummary');
    elements.resultsActions = document.getElementById('resultsActions');
    elements.exactMatchToggle = document.getElementById('exactMatchToggle');
    elements.dedupeToggle = document.getElementById('dedupeToggle');
    elements.throttleInput = document.getElementById('throttleInput');
  }

  function analyzeNames(text, removeDuplicates) {
    const lines = text
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(line => line.length > 0);

    if (!removeDuplicates) {
      return {
        names: lines,
        rawCount: lines.length,
        uniqueCount: lines.length,
        skippedDuplicates: 0
      };
    }

    const seen = new Set();
    const unique = [];
    for (const name of lines) {
      const key = name.toLowerCase();
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);
      unique.push(name);
    }

    return {
      names: unique,
      rawCount: lines.length,
      uniqueCount: unique.length,
      skippedDuplicates: lines.length - unique.length
    };
  }

  function updateNameCountDisplay() {
    const { rawCount, uniqueCount, skippedDuplicates } = state.analysis;

    if (!rawCount) {
      elements.namesCount.textContent = '0 names detected.';
      return;
    }

    if (skippedDuplicates > 0) {
      elements.namesCount.textContent = `${rawCount} names detected, ${uniqueCount} after removing duplicates.`;
      return;
    }

    elements.namesCount.textContent = `${uniqueCount} name${uniqueCount === 1 ? '' : 's'} ready to lookup.`;
  }

  function setStatus(type, message) {
    if (!elements.lookupStatus) {
      return;
    }

    if (!message) {
      elements.lookupStatus.style.display = 'none';
      elements.lookupStatus.textContent = '';
      elements.lookupStatus.className = 'status-message';
      return;
    }

    const statusClass = type ? ` status-${type}` : '';
    elements.lookupStatus.className = `status-message${statusClass}`;
    elements.lookupStatus.textContent = message;
    elements.lookupStatus.style.display = 'flex';
  }

  function updateButtonState() {
    const hasNames = state.analysis.names.length > 0;
    const isDisabled = !state.token || !hasNames || state.isRunning;

    if (elements.lookupButton) {
      elements.lookupButton.disabled = isDisabled;
    }
    if (elements.clearButton) {
      elements.clearButton.disabled = state.isRunning;
    }
    if (elements.copyButton) {
      elements.copyButton.disabled = state.results.length === 0;
    }
    if (elements.downloadButton) {
      elements.downloadButton.disabled = state.results.length === 0;
    }
  }

  function renderResults(summaryInfo = { rawCount: 0, uniqueCount: 0, skippedDuplicates: 0 }) {
    const tbody = elements.resultsBody;
    if (!tbody) {
      return;
    }

    tbody.innerHTML = '';

    if (!state.results.length) {
      const row = document.createElement('tr');
      row.className = 'empty-row';
      const cell = document.createElement('td');
      cell.colSpan = 4;
      cell.textContent = 'No results yet.';
      row.appendChild(cell);
      tbody.appendChild(row);
    } else {
      state.results.forEach(result => {
        const row = document.createElement('tr');

        const inputCell = document.createElement('td');
        inputCell.textContent = result.inputName;
        row.appendChild(inputCell);

        const emailCell = document.createElement('td');
        if (result.email) {
          const emailCode = document.createElement('code');
          emailCode.textContent = result.email;
          emailCell.appendChild(emailCode);
        } else {
          emailCell.innerHTML = '<span class="text-muted">Not available</span>';
        }
        row.appendChild(emailCell);

        const displayCell = document.createElement('td');
        displayCell.textContent = result.displayName || '—';
        row.appendChild(displayCell);

        const statusCell = document.createElement('td');
        const pill = document.createElement('span');
        const statusClass = result.statusKey === 'error'
          ? 'status-pill status-error-pill'
          : `status-pill status-${result.statusKey}`;
        pill.className = statusClass;
        pill.textContent = result.statusLabel;
        statusCell.appendChild(pill);

        if (result.note) {
          const note = document.createElement('div');
          note.className = 'status-note';
          note.textContent = result.note;
          statusCell.appendChild(note);
        }

        row.appendChild(statusCell);
        tbody.appendChild(row);
      });
    }

    if (elements.resultsActions) {
      elements.resultsActions.style.display = state.results.length ? 'flex' : 'none';
    }

    if (elements.resultsSummary) {
      if (!state.results.length) {
        elements.resultsSummary.textContent = 'Paste names and run a lookup to populate this table.';
      } else {
        const matchedCount = state.results.filter(result => ['success', 'partial', 'multiple'].includes(result.statusKey)).length;
        const missingCount = state.results.filter(result => result.statusKey === 'missing').length;
        const errorCount = state.results.filter(result => result.statusKey === 'error').length;

        const summaryParts = [];
        summaryParts.push(`Processed ${summaryInfo.uniqueCount} name${summaryInfo.uniqueCount === 1 ? '' : 's'}.`);
        if (matchedCount) {
          summaryParts.push(`${matchedCount} matched.`);
        }
        if (missingCount) {
          summaryParts.push(`${missingCount} not found.`);
        }
        if (errorCount) {
          summaryParts.push(`${errorCount} error${errorCount === 1 ? '' : 's'}.`);
        }
        if (summaryInfo.skippedDuplicates > 0) {
          summaryParts.push(`Skipped ${summaryInfo.skippedDuplicates} duplicate${summaryInfo.skippedDuplicates === 1 ? '' : 's'}.`);
        }
        elements.resultsSummary.textContent = summaryParts.join(' ');
      }
    }
  }

  function getLookupOptions() {
    const preferExact = elements.exactMatchToggle ? elements.exactMatchToggle.checked : true;
    const dedupe = elements.dedupeToggle ? elements.dedupeToggle.checked : true;
    const parsedDelay = elements.throttleInput ? parseInt(elements.throttleInput.value, 10) : 0;
    const delay = Number.isFinite(parsedDelay) && parsedDelay >= 0 ? parsedDelay : 0;

    return { preferExact, dedupe, delay };
  }

  function buildSearchUrl(name) {
    const sanitized = name.replace(/"/g, '\\"');
    const query = encodeURIComponent(`"displayName:${sanitized}"`);
    return `${GRAPH_ENDPOINT}?$search=${query}&$select=displayName,mail,userPrincipalName,jobTitle,department,id&$top=${MAX_RESULTS}`;
  }

  function buildFilterUrl(name) {
    const sanitized = name.replace(/'/g, "''");
    const filter = encodeURIComponent(`startswith(displayName,'${sanitized}')`);
    return `${GRAPH_ENDPOINT}?$filter=${filter}&$select=displayName,mail,userPrincipalName,jobTitle,department,id&$top=${MAX_RESULTS}`;
  }

  async function executeQuery(url) {
    try {
      const response = await fetch(url, {
        headers: {
          Authorization: `Bearer ${state.token}`,
          'ConsistencyLevel': 'eventual'
        }
      });

      const data = await response.json();
      if (!response.ok) {
        const message = data && data.error && data.error.message
          ? data.error.message
          : `Request failed with status ${response.status}`;
        return { ok: false, status: response.status, value: [], error: message };
      }

      return { ok: true, status: response.status, value: Array.isArray(data.value) ? data.value : [] };
    } catch (error) {
      return { ok: false, status: 0, value: [], error: error.message || 'Network error' };
    }
  }

  async function fetchCandidates(name) {
    const searchOutcome = await executeQuery(buildSearchUrl(name));

    if (searchOutcome.ok && searchOutcome.value.length) {
      return searchOutcome.value;
    }

    if (searchOutcome.ok && !searchOutcome.value.length) {
      const filterOutcome = await executeQuery(buildFilterUrl(name));
      if (filterOutcome.ok) {
        return filterOutcome.value;
      }
      if (filterOutcome.error) {
        throw new Error(filterOutcome.error);
      }
      return [];
    }

    if (!searchOutcome.ok) {
      // If the search query isn't permitted (e.g., missing permissions), attempt a filter fallback before surfacing the error.
      if (searchOutcome.status === 400 || searchOutcome.status === 403) {
        const filterOutcome = await executeQuery(buildFilterUrl(name));
        if (filterOutcome.ok) {
          return filterOutcome.value;
        }
        if (filterOutcome.error) {
          throw new Error(filterOutcome.error);
        }
        return [];
      }

      throw new Error(searchOutcome.error || 'Lookup failed');
    }

    return [];
  }

  function formatResult(name, candidates, options) {
    if (!Array.isArray(candidates) || !candidates.length) {
      return {
        inputName: name,
        email: '',
        displayName: '',
        statusKey: 'missing',
        statusLabel: 'Not found',
        note: 'No Microsoft Graph users matched this name.'
      };
    }

    const normalizedTarget = name.trim().toLowerCase();
    let match = null;

    if (options.preferExact) {
      match = candidates.find(candidate => (candidate.displayName || '').trim().toLowerCase() === normalizedTarget) || null;
    }

    if (!match) {
      match = candidates[0];
    }

    const displayName = match.displayName || '';
    const primaryEmail = match.mail || match.userPrincipalName || '';
    const upn = match.userPrincipalName || '';
    const totalCandidates = candidates.length;
    const isExact = displayName.trim().toLowerCase() === normalizedTarget;

    let statusKey = 'partial';
    let statusLabel = 'Closest match';
    const notes = [];

    if (isExact && totalCandidates === 1) {
      statusKey = 'success';
      statusLabel = 'Exact match';
    } else if (isExact && totalCandidates > 1) {
      statusKey = 'multiple';
      statusLabel = 'Exact match (multiple found)';
      notes.push(`${totalCandidates} results returned.`);
    } else if (!isExact && totalCandidates === 1) {
      statusKey = 'partial';
      statusLabel = 'Single result';
      notes.push('Display name differs from the input.');
    } else {
      statusKey = 'partial';
      statusLabel = 'Closest match';
      notes.push(`${totalCandidates} results returned. Display name differs from the input.`);
    }

    if (!primaryEmail && upn) {
      notes.push('No primary email on file. Using user principal name instead.');
    }

    if (primaryEmail && upn && primaryEmail !== upn) {
      notes.push(`User principal name: ${upn}`);
    }

    return {
      inputName: name,
      email: primaryEmail,
      displayName,
      statusKey,
      statusLabel,
      note: notes.join(' ')
    };
  }

  async function lookupSingleName(name, options) {
    try {
      const candidates = await fetchCandidates(name);
      return formatResult(name, candidates, options);
    } catch (error) {
      return {
        inputName: name,
        email: '',
        displayName: '',
        statusKey: 'error',
        statusLabel: 'Lookup failed',
        note: error.message || 'Unexpected error.'
      };
    }
  }

  async function handleLookup() {
    if (state.isRunning) {
      return;
    }

    const options = getLookupOptions();
    state.analysis = analyzeNames(elements.namesInput.value, options.dedupe);
    updateNameCountDisplay();

    if (!state.token) {
      setStatus('error', 'Paste a Microsoft Graph access token before running a lookup.');
      return;
    }

    if (!state.analysis.names.length) {
      setStatus('warning', 'Add at least one name to lookup.');
      return;
    }

    state.isRunning = true;
    state.results = [];
    updateButtonState();

    const total = state.analysis.names.length;
    setStatus('info', `Looking up ${total} name${total === 1 ? '' : 's'}...`);

    for (let index = 0; index < total; index += 1) {
      const name = state.analysis.names[index];
      setStatus('info', `Looking up ${name} (${index + 1} of ${total})...`);
      const result = await lookupSingleName(name, options);
      state.results.push(result);
      renderResults({
        rawCount: state.analysis.rawCount,
        uniqueCount: state.analysis.uniqueCount,
        skippedDuplicates: state.analysis.skippedDuplicates
      });

      if (options.delay > 0 && index < total - 1) {
        await new Promise(resolve => setTimeout(resolve, options.delay));
      }
    }

    setStatus('success', `Lookup complete for ${total} name${total === 1 ? '' : 's'}.`);
    AppUI.setSectionCollapsed('loadSection', true);
    state.isRunning = false;
    updateButtonState();
  }

  function handleClear() {
    if (state.isRunning) {
      return;
    }

    elements.namesInput.value = '';
    state.analysis = analyzeNames('', getLookupOptions().dedupe);
    state.results = [];
    renderResults({ rawCount: 0, uniqueCount: 0, skippedDuplicates: 0 });
    updateNameCountDisplay();
    setStatus(null, '');
    updateButtonState();
    AppUI.setSectionCollapsed('loadSection', false);
  }

  function buildTabDelimited(results) {
    const header = ['Input Name', 'Primary Email', 'Matched Display Name', 'Status', 'Notes'];
    const lines = [header.join('\t')];

    results.forEach(result => {
      lines.push([
        result.inputName || '',
        result.email || '',
        result.displayName || '',
        result.statusLabel || '',
        result.note || ''
      ].join('\t'));
    });

    return lines.join('\n');
  }

  function quoteCsvField(value) {
    if (value === null || value === undefined) {
      return '""';
    }

    const stringValue = String(value);
    const escaped = stringValue.replace(/"/g, '""');
    return `"${escaped}"`;
  }

  function buildCsv(results) {
    const header = ['Input Name', 'Primary Email', 'Matched Display Name', 'Status', 'Notes'];
    const rows = [header.map(quoteCsvField).join(',')];

    results.forEach(result => {
      const row = [
        quoteCsvField(result.inputName || ''),
        quoteCsvField(result.email || ''),
        quoteCsvField(result.displayName || ''),
        quoteCsvField(result.statusLabel || ''),
        quoteCsvField(result.note || '')
      ].join(',');
      rows.push(row);
    });

    return rows.join('\n');
  }

  async function handleCopy() {
    if (!state.results.length) {
      return;
    }

    if (!navigator.clipboard || typeof navigator.clipboard.writeText !== 'function') {
      setStatus('warning', 'Clipboard copy is not supported in this browser.');
      return;
    }

    try {
      await navigator.clipboard.writeText(buildTabDelimited(state.results));
      setStatus('success', 'Results copied to the clipboard.');
    } catch (error) {
      setStatus('error', `Unable to copy results: ${error.message || error}`);
    }
  }

  function handleDownload() {
    if (!state.results.length) {
      return;
    }

    try {
      const csv = buildCsv(state.results);
      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      link.href = url;
      link.download = `user-email-lookup-${timestamp}.csv`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      setStatus('success', 'CSV download ready.');
    } catch (error) {
      setStatus('error', `Unable to generate CSV: ${error.message || error}`);
    }
  }

  function handleTokenChange(token) {
    state.token = token || '';
    AppUI.updateTokenStatus('tokenStatus', token);
    updateButtonState();
  }

  function handleNamesInput() {
    const options = getLookupOptions();
    state.analysis = analyzeNames(elements.namesInput.value, options.dedupe);
    updateNameCountDisplay();
    updateButtonState();
  }

  function initialize() {
    AppUI.initialize();
    cacheElements();

    GraphTokenManager.initialize({
      onTokenChange: handleTokenChange,
      textareaId: 'graphToken'
    });

    if (elements.namesInput) {
      elements.namesInput.addEventListener('input', handleNamesInput);
    }
    if (elements.lookupButton) {
      elements.lookupButton.addEventListener('click', handleLookup);
    }
    if (elements.clearButton) {
      elements.clearButton.addEventListener('click', handleClear);
    }
    if (elements.copyButton) {
      elements.copyButton.addEventListener('click', handleCopy);
    }
    if (elements.downloadButton) {
      elements.downloadButton.addEventListener('click', handleDownload);
    }
    if (elements.dedupeToggle) {
      elements.dedupeToggle.addEventListener('change', handleNamesInput);
    }

    AppUI.setSectionCollapsed('tokenSection', true);
    AppUI.setSectionCollapsed('settingsSection', true);

    state.analysis = analyzeNames('', getLookupOptions().dedupe);
    updateNameCountDisplay();
    renderResults();
    updateButtonState();
  }

  return { initialize };
})();

document.addEventListener('DOMContentLoaded', () => {
  EmailLookupApp.initialize();
});
