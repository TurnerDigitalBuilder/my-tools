const ColumnComparer = (() => {
  const state = {
    datasets: { A: null, B: null },
    selections: { A: '', B: '' },
    results: null
  };

  const elements = {};

  function cacheElements() {
    elements.datasetInputs = {
      A: document.getElementById('datasetAFile'),
      B: document.getElementById('datasetBFile')
    };
    elements.datasetSummaries = {
      A: document.getElementById('datasetASummary'),
      B: document.getElementById('datasetBSummary')
    };
    elements.datasetStatuses = {
      A: document.getElementById('datasetAStatus'),
      B: document.getElementById('datasetBStatus')
    };
    elements.columnSelects = {
      A: document.getElementById('columnSelectA'),
      B: document.getElementById('columnSelectB')
    };
    elements.optionToggles = {
      ignoreCase: document.getElementById('ignoreCaseToggle'),
      trim: document.getElementById('trimToggle'),
      skipBlank: document.getElementById('skipBlankToggle')
    };
    elements.compareButton = document.getElementById('compareButton');
    elements.comparisonStatus = document.getElementById('comparisonStatus');
    elements.introState = document.getElementById('introState');
    elements.overviewSection = document.getElementById('overviewSection');
    elements.matchesSection = document.getElementById('matchesSection');
    elements.onlyASection = document.getElementById('onlyASection');
    elements.onlyBSection = document.getElementById('onlyBSection');
    elements.duplicatesSection = document.getElementById('duplicatesSection');
    elements.overviewDescription = document.getElementById('overviewDescription');
    elements.selectedColumns = document.getElementById('selectedColumns');
    elements.summaryGrid = document.getElementById('summaryGrid');
    elements.matchesTableBody = document.getElementById('matchesTableBody');
    elements.onlyATableBody = document.getElementById('onlyATableBody');
    elements.onlyBTableBody = document.getElementById('onlyBTableBody');
    elements.duplicatesATableBody = document.getElementById('duplicatesATableBody');
    elements.duplicatesBTableBody = document.getElementById('duplicatesBTableBody');
    elements.tableEmptyStates = {
      matches: document.getElementById('matchesEmpty'),
      onlyA: document.getElementById('onlyAEmpty'),
      onlyB: document.getElementById('onlyBEmpty'),
      duplicatesA: document.getElementById('duplicatesAEmpty'),
      duplicatesB: document.getElementById('duplicatesBEmpty')
    };
  }

  function attachEventListeners() {
    elements.datasetInputs.A?.addEventListener('change', event => {
      const [file] = event.target.files;
      loadDataset('A', file);
    });

    elements.datasetInputs.B?.addEventListener('change', event => {
      const [file] = event.target.files;
      loadDataset('B', file);
    });

    elements.columnSelects.A?.addEventListener('change', event => {
      state.selections.A = event.target.value;
      invalidateResults('Column selection changed. Run the comparison again to refresh results.');
      updateCompareButtonState();
    });

    elements.columnSelects.B?.addEventListener('change', event => {
      state.selections.B = event.target.value;
      invalidateResults('Column selection changed. Run the comparison again to refresh results.');
      updateCompareButtonState();
    });

    elements.optionToggles.ignoreCase?.addEventListener('change', () => {
      invalidateResults('Comparison options changed. Run the comparison again to refresh results.');
    });

    elements.optionToggles.trim?.addEventListener('change', () => {
      invalidateResults('Comparison options changed. Run the comparison again to refresh results.');
    });

    elements.optionToggles.skipBlank?.addEventListener('change', () => {
      invalidateResults('Comparison options changed. Run the comparison again to refresh results.');
    });

    elements.compareButton?.addEventListener('click', () => runComparison());
  }

  function invalidateResults(message) {
    if (state.results) {
      setStatus(elements.comparisonStatus, message, 'info');
    }

    state.results = null;
    elements.overviewSection.hidden = true;
    elements.matchesSection.hidden = true;
    elements.onlyASection.hidden = true;
    elements.onlyBSection.hidden = true;
    elements.duplicatesSection.hidden = true;
    if (elements.introState) {
      elements.introState.style.display = '';
    }
  }

  function setStatus(element, message, type = 'info') {
    if (!element) {
      return;
    }

    element.className = 'status-message';
    element.style.display = 'none';

    if (!message) {
      element.textContent = '';
      return;
    }

    element.textContent = message;
    element.classList.add(`status-${type}`);
    element.style.display = 'flex';
  }

  async function loadDataset(key, file) {
    const summaryEl = elements.datasetSummaries[key];
    const statusEl = elements.datasetStatuses[key];
    const selectEl = elements.columnSelects[key];

    if (!file) {
      state.datasets[key] = null;
      state.selections[key] = '';
      if (summaryEl) {
        summaryEl.textContent = 'No file selected yet.';
      }
      if (selectEl) {
        resetColumnSelect(selectEl, key);
      }
      setStatus(statusEl, 'File selection cleared.', 'info');
      updateCompareButtonState();
      invalidateResults();
      return;
    }

    setStatus(statusEl, `Loading ${file.name}…`, 'info');

    try {
      const dataset = await parseFile(file);
      dataset.fileName = file.name;
      state.datasets[key] = dataset;
      state.selections[key] = '';
      if (summaryEl) {
        summaryEl.textContent = `${file.name} • ${dataset.rows.length} rows, ${dataset.columns.length} columns`;
      }
      if (selectEl) {
        populateColumnSelect(selectEl, dataset.columns, key);
      }
      setStatus(statusEl, `Loaded ${dataset.rows.length} rows and ${dataset.columns.length} columns.`, 'success');
      invalidateResults('Dataset changed. Run the comparison again to refresh results.');
      if (state.datasets.A && state.datasets.B) {
        AppUI.setSectionCollapsed('loadSection', true);
      }
    } catch (error) {
      console.error('Failed to load dataset', error);
      state.datasets[key] = null;
      state.selections[key] = '';
      if (summaryEl) {
        summaryEl.textContent = 'Unable to read file.';
      }
      if (selectEl) {
        resetColumnSelect(selectEl, key);
      }
      setStatus(statusEl, `Unable to load file: ${error.message || 'Unknown error'}`, 'error');
    } finally {
      updateCompareButtonState();
    }
  }

  function resetColumnSelect(selectEl, key) {
    selectEl.innerHTML = '';
    const option = document.createElement('option');
    option.value = '';
    option.textContent = key === 'A' ? 'Load Dataset A to choose a column' : 'Load Dataset B to choose a column';
    selectEl.appendChild(option);
    selectEl.disabled = true;
  }

  function populateColumnSelect(selectEl, columns, key) {
    const previousSelection = state.selections[key];

    selectEl.innerHTML = '';
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = columns.length
      ? `Select a column (${columns.length} available)`
      : 'No columns detected';
    selectEl.appendChild(placeholder);

    columns.forEach(columnName => {
      const option = document.createElement('option');
      option.value = columnName;
      option.textContent = columnName;
      selectEl.appendChild(option);
    });

    selectEl.disabled = columns.length === 0;

    if (columns.length > 0) {
      if (previousSelection && columns.includes(previousSelection)) {
        selectEl.value = previousSelection;
        state.selections[key] = previousSelection;
      } else {
        selectEl.value = columns[0];
        state.selections[key] = columns[0];
      }
      updateCompareButtonState();
    } else {
      state.selections[key] = '';
      updateCompareButtonState();
    }
  }

  function updateCompareButtonState() {
    const ready = Boolean(state.datasets.A && state.datasets.B && state.selections.A && state.selections.B);
    if (elements.compareButton) {
      elements.compareButton.disabled = !ready;
    }
  }

  function getOptions() {
    return {
      ignoreCase: Boolean(elements.optionToggles.ignoreCase?.checked),
      trim: Boolean(elements.optionToggles.trim?.checked),
      skipBlank: Boolean(elements.optionToggles.skipBlank?.checked)
    };
  }

  function normalizeValue(value, options) {
    let text = '';

    if (value !== undefined && value !== null) {
      if (value instanceof Date) {
        text = value.toISOString();
      } else {
        text = String(value);
      }
    }

    if (options.trim) {
      text = text.trim();
    }

    if (text.length === 0 && options.skipBlank) {
      return null;
    }

    const normalized = options.ignoreCase ? text.toLowerCase() : text;

    return {
      normalized,
      display: text
    };
  }

  function collectColumnValues(dataset, columnName, options) {
    const map = new Map();
    let valueCount = 0;

    dataset.rows.forEach((row, index) => {
      const raw = Object.prototype.hasOwnProperty.call(row, columnName) ? row[columnName] : '';
      const normalized = normalizeValue(raw, options);
      if (!normalized) {
        return;
      }

      valueCount += 1;
      if (!map.has(normalized.normalized)) {
        map.set(normalized.normalized, {
          key: normalized.normalized,
          display: normalized.display,
          count: 1
        });
      } else {
        const entry = map.get(normalized.normalized);
        entry.count += 1;
        if (!entry.display && normalized.display) {
          entry.display = normalized.display;
        }
      }
    });

    return { map, valueCount };
  }

  function formatValueForDisplay(value) {
    if (value === null || value === undefined || value === '') {
      return '(blank)';
    }
    return value;
  }

  function combineDisplay(entryA, entryB) {
    const displays = new Set();
    if (entryA?.display) {
      displays.add(entryA.display);
    }
    if (entryB?.display) {
      displays.add(entryB.display);
    }

    if (!displays.size) {
      return '(blank)';
    }

    return Array.from(displays).join(' | ');
  }

  function computeComparison(datasetA, datasetB, columnA, columnB, options) {
    const valuesA = collectColumnValues(datasetA, columnA, options);
    const valuesB = collectColumnValues(datasetB, columnB, options);

    const matches = [];
    const onlyInA = [];
    const onlyInB = [];
    const duplicatesA = [];
    const duplicatesB = [];

    valuesA.map.forEach(entryA => {
      const counterpart = valuesB.map.get(entryA.key);
      if (counterpart) {
        matches.push({
          value: combineDisplay(entryA, counterpart),
          countA: entryA.count,
          countB: counterpart.count,
          sortKey: entryA.key
        });
      } else {
        onlyInA.push({ value: formatValueForDisplay(entryA.display), count: entryA.count, sortKey: entryA.key });
      }

      if (entryA.count > 1) {
        duplicatesA.push({ value: formatValueForDisplay(entryA.display), count: entryA.count, sortKey: entryA.key });
      }
    });

    valuesB.map.forEach(entryB => {
      if (!valuesA.map.has(entryB.key)) {
        onlyInB.push({ value: formatValueForDisplay(entryB.display), count: entryB.count, sortKey: entryB.key });
      }

      if (entryB.count > 1) {
        duplicatesB.push({ value: formatValueForDisplay(entryB.display), count: entryB.count, sortKey: entryB.key });
      }
    });

    const sortAlpha = (a, b) => a.sortKey.localeCompare(b.sortKey, undefined, { sensitivity: 'base', numeric: true });
    const sortByCount = (a, b) => b.count - a.count || sortAlpha(a, b);

    matches.sort(sortAlpha);
    onlyInA.sort(sortAlpha);
    onlyInB.sort(sortAlpha);
    duplicatesA.sort(sortByCount);
    duplicatesB.sort(sortByCount);

    return {
      matches,
      onlyInA,
      onlyInB,
      duplicatesA,
      duplicatesB,
      totals: {
        rowsA: datasetA.rows.length,
        rowsB: datasetB.rows.length,
        valuesComparedA: valuesA.valueCount,
        valuesComparedB: valuesB.valueCount,
        uniqueA: valuesA.map.size,
        uniqueB: valuesB.map.size,
        matches: matches.length,
        onlyA: onlyInA.length,
        onlyB: onlyInB.length,
        duplicatesA: duplicatesA.length,
        duplicatesB: duplicatesB.length
      }
    };
  }

  function escapeHtml(text) {
    const value = text === null || text === undefined ? '' : String(text);
    return value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function renderTableRows(tbody, emptyState, rows, renderRow) {
    if (!tbody || !emptyState) {
      return;
    }

    if (!rows.length) {
      tbody.innerHTML = '';
      emptyState.style.display = 'block';
      return;
    }

    const fragments = rows.map(renderRow).join('');
    tbody.innerHTML = fragments;
    emptyState.style.display = 'none';
  }

  function renderResults(results, datasetA, datasetB, columnA, columnB) {
    if (elements.introState) {
      elements.introState.style.display = 'none';
    }

    elements.overviewSection.hidden = false;
    elements.matchesSection.hidden = false;
    elements.onlyASection.hidden = false;
    elements.onlyBSection.hidden = false;
    elements.duplicatesSection.hidden = false;

    const overviewParts = [];
    overviewParts.push(`Comparing ${formatNumber(results.totals.valuesComparedA)} values from ${datasetA.fileName}`);
    overviewParts.push(`against ${formatNumber(results.totals.valuesComparedB)} values from ${datasetB.fileName}.`);
    elements.overviewDescription.textContent = overviewParts.join(' ');

    if (elements.selectedColumns) {
      elements.selectedColumns.innerHTML = `<span>A: ${escapeHtml(columnA)}</span><span>B: ${escapeHtml(columnB)}</span>`;
    }

    if (elements.summaryGrid) {
      const metrics = [
        {
          title: 'Rows Analyzed · Dataset A',
          value: formatNumber(results.totals.rowsA),
          subtext: datasetA.fileName
        },
        {
          title: 'Rows Analyzed · Dataset B',
          value: formatNumber(results.totals.rowsB),
          subtext: datasetB.fileName
        },
        {
          title: 'Unique Values · Dataset A',
          value: formatNumber(results.totals.uniqueA),
          subtext: `${formatNumber(results.totals.valuesComparedA)} values considered`
        },
        {
          title: 'Unique Values · Dataset B',
          value: formatNumber(results.totals.uniqueB),
          subtext: `${formatNumber(results.totals.valuesComparedB)} values considered`
        },
        {
          title: 'Values in Both Files',
          value: formatNumber(results.totals.matches),
          subtext: 'Shared across the selected columns'
        },
        {
          title: 'Only in Dataset A',
          value: formatNumber(results.totals.onlyA),
          subtext: 'Missing from Dataset B'
        },
        {
          title: 'Only in Dataset B',
          value: formatNumber(results.totals.onlyB),
          subtext: 'Missing from Dataset A'
        }
      ];

      elements.summaryGrid.innerHTML = metrics
        .map(metric => {
          return `
            <article class="summary-card">
              <h4>${escapeHtml(metric.title)}</h4>
              <div class="metric-value">${escapeHtml(metric.value)}</div>
              <div class="metric-subtext">${escapeHtml(metric.subtext)}</div>
            </article>
          `;
        })
        .join('');
    }

    renderTableRows(
      elements.matchesTableBody,
      elements.tableEmptyStates.matches,
      results.matches,
      row => `
        <tr>
          <td>${escapeHtml(formatValueForDisplay(row.value))}</td>
          <td>${escapeHtml(formatNumber(row.countA))}</td>
          <td>${escapeHtml(formatNumber(row.countB))}</td>
        </tr>
      `
    );

    renderTableRows(
      elements.onlyATableBody,
      elements.tableEmptyStates.onlyA,
      results.onlyInA,
      row => `
        <tr>
          <td>${escapeHtml(formatValueForDisplay(row.value))}</td>
          <td>${escapeHtml(formatNumber(row.count))}</td>
        </tr>
      `
    );

    renderTableRows(
      elements.onlyBTableBody,
      elements.tableEmptyStates.onlyB,
      results.onlyInB,
      row => `
        <tr>
          <td>${escapeHtml(formatValueForDisplay(row.value))}</td>
          <td>${escapeHtml(formatNumber(row.count))}</td>
        </tr>
      `
    );

    renderTableRows(
      elements.duplicatesATableBody,
      elements.tableEmptyStates.duplicatesA,
      results.duplicatesA,
      row => `
        <tr>
          <td>${escapeHtml(formatValueForDisplay(row.value))}</td>
          <td>${escapeHtml(formatNumber(row.count))}</td>
        </tr>
      `
    );

    renderTableRows(
      elements.duplicatesBTableBody,
      elements.tableEmptyStates.duplicatesB,
      results.duplicatesB,
      row => `
        <tr>
          <td>${escapeHtml(formatValueForDisplay(row.value))}</td>
          <td>${escapeHtml(formatNumber(row.count))}</td>
        </tr>
      `
    );
  }

  function formatNumber(value) {
    return Number.isFinite(value) ? value.toLocaleString() : '0';
  }

  async function runComparison() {
    if (!state.datasets.A || !state.datasets.B || !state.selections.A || !state.selections.B) {
      setStatus(elements.comparisonStatus, 'Load both datasets and choose a column from each before running a comparison.', 'warning');
      return;
    }

    const options = getOptions();
    setStatus(elements.comparisonStatus, 'Comparing selected columns…', 'info');

    try {
      const results = computeComparison(state.datasets.A, state.datasets.B, state.selections.A, state.selections.B, options);
      state.results = results;
      renderResults(results, state.datasets.A, state.datasets.B, state.selections.A, state.selections.B);
      setStatus(elements.comparisonStatus, 'Comparison complete. Scroll down to review the results.', 'success');
      AppUI.setSectionCollapsed('selectionSection', true);
    } catch (error) {
      console.error('Comparison failed', error);
      setStatus(elements.comparisonStatus, `Comparison failed: ${error.message || 'Unknown error'}`, 'error');
    }
  }

  function parseFile(file) {
    const extension = (file.name.split('.').pop() || '').toLowerCase();

    if (extension === 'csv') {
      return parseCsv(file);
    }

    if (extension === 'xlsx' || extension === 'xls') {
      return parseWorkbook(file);
    }

    return Promise.reject(new Error('Unsupported file type. Please upload a CSV or Excel file.'));
  }

  function parseCsv(file) {
    const headerCounts = {};

    return new Promise((resolve, reject) => {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: 'greedy',
        transformHeader: (header, index) => {
          const base = (header || '').trim() || `Column ${index + 1}`;
          const count = headerCounts[base] || 0;
          headerCounts[base] = count + 1;
          return count === 0 ? base : `${base} (${count + 1})`;
        },
        complete: results => {
          if (results.errors && results.errors.length) {
            reject(new Error(results.errors[0].message));
            return;
          }

          const columns = ensureUniqueColumnNames(results.meta?.fields || []);
          const rows = (results.data || []).filter(row => !isRowEmpty(row));
          resolve({ rows, columns });
        },
        error: error => {
          reject(error);
        }
      });
    });
  }

  function parseWorkbook(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error('Unable to read file.'));
      reader.onload = event => {
        try {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const [firstSheet] = workbook.SheetNames;
          if (!firstSheet) {
            resolve({ rows: [], columns: [] });
            return;
          }

          const worksheet = workbook.Sheets[firstSheet];
          const rowsArray = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
          if (!rowsArray.length) {
            resolve({ rows: [], columns: [] });
            return;
          }

          const [headerRow, ...dataRows] = rowsArray;
          const columns = ensureUniqueColumnNames(headerRow);

          const rows = dataRows
            .map(cells => {
              const row = {};
              columns.forEach((columnName, index) => {
                row[columnName] = cells[index] ?? '';
              });
              return row;
            })
            .filter(row => !isRowEmpty(row));

          resolve({ rows, columns });
        } catch (error) {
          reject(error);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  function isRowEmpty(row) {
    return Object.values(row).every(value => {
      if (value === null || value === undefined) {
        return true;
      }
      if (typeof value === 'string') {
        return value.trim().length === 0;
      }
      return String(value).trim().length === 0;
    });
  }

  function ensureUniqueColumnNames(columns) {
    const seen = new Map();
    return columns.map((rawName, index) => {
      const base = (rawName && String(rawName).trim()) || `Column ${index + 1}`;
      const count = seen.get(base) || 0;
      seen.set(base, count + 1);
      if (count === 0) {
        return base;
      }
      return `${base} (${count + 1})`;
    });
  }

  return {
    initialize() {
      cacheElements();
      attachEventListeners();
      if (elements.columnSelects.A) {
        resetColumnSelect(elements.columnSelects.A, 'A');
      }
      if (elements.columnSelects.B) {
        resetColumnSelect(elements.columnSelects.B, 'B');
      }
    }
  };
})();
