const TemplateApp = (() => {
  function handleTokenChange(token) {
    AppUI.updateTokenStatus('tokenStatus', token, { start: 12, end: 6 });
  }

  function formatSettingsSummary() {
    const toggle = document.getElementById('settingToggle');
    const numberField = document.getElementById('settingNumber');
    const notes = document.getElementById('settingNotes');

    const parts = [];
    if (toggle) {
      parts.push(`Toggle ${toggle.checked ? 'enabled' : 'disabled'}`);
    }
    if (numberField) {
      parts.push(`Number set to ${numberField.value || 'n/a'}`);
    }
    if (notes && notes.value.trim()) {
      parts.push(`Notes captured (${notes.value.trim().length} chars)`);
    }

    return parts.join(' • ') || 'Default settings in use';
  }

  function runTemplateAction() {
    const inputField = document.getElementById('inputText');
    const selectField = document.getElementById('inputSelect');
    const fileField = document.getElementById('inputFile');
    const status = document.getElementById('inputStatus');
    const resultBadge = document.getElementById('resultBadge');

    if (!status || !resultBadge) {
      return;
    }

    const details = [];
    if (inputField && inputField.value.trim()) {
      details.push(`Text: "${inputField.value.trim()}"`);
    }
    if (selectField && selectField.value) {
      details.push(`Dropdown: ${selectField.options[selectField.selectedIndex].text}`);
    }
    if (fileField && fileField.files && fileField.files.length > 0) {
      details.push(`File: ${fileField.files[0].name}`);
    }

    const message = details.length > 0
      ? `Template action triggered with ${details.join(', ')}.`
      : 'Template action triggered with default values.';

    status.textContent = `${message} ${formatSettingsSummary()}.`;
    status.className = 'status-message status-success';
    status.style.display = 'flex';

    resultBadge.textContent = 'Template action completed successfully';
    resultBadge.classList.remove('muted');
    resultBadge.classList.add('success');
  }

  function bindEvents() {
    const actionButton = document.getElementById('inputActionButton');
    if (actionButton) {
      actionButton.addEventListener('click', event => {
        event.preventDefault();
        runTemplateAction();
      });
    }

    const notesField = document.getElementById('settingNotes');
    if (notesField) {
      notesField.addEventListener('input', () => {
        const resultBadge = document.getElementById('resultBadge');
        if (!resultBadge) {
          return;
        }
        if (notesField.value.trim()) {
          resultBadge.classList.remove('muted');
        } else if (!resultBadge.classList.contains('success')) {
          resultBadge.classList.add('muted');
        }
      });
    }
  }

  function collapseDefaults() {
    AppUI.setSectionCollapsed('tokenSection', true);
    AppUI.setSectionCollapsed('inputSection', true);
    AppUI.setSectionCollapsed('settingsSection', false);
  }

  function initialize() {
    AppUI.initialize();
    GraphTokenManager.initialize({ onTokenChange: handleTokenChange });
    collapseDefaults();
    bindEvents();

    const resultBadge = document.getElementById('resultBadge');
    if (resultBadge) {
      resultBadge.classList.add('muted');
    }
  }

  return { initialize };
})();

window.addEventListener('DOMContentLoaded', () => {
  TemplateApp.initialize();
});
