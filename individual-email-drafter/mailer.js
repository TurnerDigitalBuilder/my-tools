const EmailDraftTool = (() => {
  const state = {
    recipients: [],
    duplicateCount: 0,
    subject: '',
    cc: '',
    bcc: '',
    body: '',
    htmlBody: '',
    useHtml: false
  };

  const emailRegex = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;

  function init() {
    AppUI.initialize();
    wireEvents();
    seedSampleContent();
    updateSummary();
    renderRecipientRows();
    updateDraftStatus();
    announceStatus('Load a CSV/Excel file or paste addresses to get started.');
  }

  function wireEvents() {
    document.getElementById('recipientFile').addEventListener('change', handleFileSelect);
    document.getElementById('addPastedButton').addEventListener('click', handlePastedRecipients);
    document.getElementById('clearPastedButton').addEventListener('click', () => {
      document.getElementById('pastedRecipients').value = '';
    });

    document.getElementById('subject').addEventListener('input', handleComposeChange);
    document.getElementById('cc').addEventListener('input', handleComposeChange);
    document.getElementById('bcc').addEventListener('input', handleComposeChange);
    document.getElementById('body').addEventListener('input', handleComposeChange);
    document.getElementById('htmlBody').addEventListener('input', handleComposeChange);
    document.getElementById('useHtml').addEventListener('change', handleComposeChange);

    document.getElementById('updateDraftsButton').addEventListener('click', updateDrafts);
    document.getElementById('resetFormButton').addEventListener('click', resetForm);
    document.getElementById('copyAllButton').addEventListener('click', copyAllDrafts);
    document.getElementById('clearRecipientsButton').addEventListener('click', clearRecipients);
  }

  function seedSampleContent() {
    const sampleSubject = 'Help Us Measure the Impact of ChatGPT - Complete Survey by 12/11';
    const sampleBody = `Hi team,\n\nAs part of our rollout of ChatGPT Enterprise, we’re running a short Impact Survey to better understand how ChatGPT is supporting your work.\n\nWhy this matters: Your feedback helps us see where ChatGPT is delivering value today, where we can improve support and training, and how to shape our roadmap going forward. The survey was created by OpenAI, which means your responses will also help Turner benchmark our adoption and impact against other enterprise customers.\n\nThe survey covers how you use ChatGPT and how it’s affecting your productivity and workflow and should take approximately 5 minutes to complete\n\n👉 Impact Survey Link\n\nPlease complete by EOW: Dec 12th\n\nYour perspective is essential in helping us ensure ChatGPT continues to deliver meaningful impact for you and your teams. Thank you for taking the time to share your input.\n\nBest,\nBen`;

    document.getElementById('subject').value = sampleSubject;
    document.getElementById('body').value = sampleBody;

    handleComposeChange();
  }

  function handleFileSelect(event) {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    document.getElementById('fileStatus').textContent = `${file.name} selected.`;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: '' });
        ingestRows(rows);
      } catch (error) {
        console.error('Failed to parse file', error);
        announceStatus('Unable to read the file. Please confirm it is a CSV or Excel document with an Email column.', 'warning');
      }
    };

    reader.readAsArrayBuffer(file);
  }

  function ingestRows(rows) {
    if (!Array.isArray(rows) || !rows.length) {
      announceStatus('No rows detected. Confirm the sheet has data and an Email column.', 'warning');
      return;
    }

    let added = 0;
    let duplicates = 0;

    const existing = new Set(state.recipients.map((r) => r.email.toLowerCase()));

    rows.forEach((row) => {
      const parsed = parseRow(row);
      if (!parsed) {
        return;
      }

      const emailKey = parsed.email.toLowerCase();
      if (existing.has(emailKey)) {
        duplicates += 1;
        return;
      }

      existing.add(emailKey);
      state.recipients.push(parsed);
      added += 1;
    });

    state.duplicateCount += duplicates;

    announceStatus(`${added} recipient${added === 1 ? '' : 's'} added from file. ${duplicates ? `${duplicates} duplicate${duplicates === 1 ? '' : 's'} skipped.` : ''}`, added ? 'success' : 'warning');
    updateSummary();
    renderRecipientRows();
    updateDraftStatus();
  }

  function parseRow(row) {
    if (!row) {
      return null;
    }

    if (typeof row === 'string') {
      const email = (row.match(emailRegex) || [])[0];
      return email ? { email: email.trim(), name: '' } : null;
    }

    const keys = Object.keys(row);
    const emailKey = keys.find((key) => key.toLowerCase().includes('email') || key.toLowerCase() === 'mail');
    const emailValue = emailKey ? row[emailKey] : null;

    if (typeof emailValue !== 'string' || !emailRegex.test(emailValue)) {
      return null;
    }

    const nameKey = keys.find((key) => ['name', 'full name', 'display name'].includes(key.toLowerCase()));
    const nameValue = nameKey && typeof row[nameKey] === 'string' ? row[nameKey].trim() : '';

    return {
      email: emailValue.trim(),
      name: nameValue
    };
  }

  function handlePastedRecipients() {
    const textarea = document.getElementById('pastedRecipients');
    const pasted = textarea.value.trim();
    if (!pasted) {
      announceStatus('Nothing to add. Paste one or more email addresses first.', 'warning');
      return;
    }

    const candidates = pasted.split(/[,;\n]+/).map((entry) => entry.trim()).filter(Boolean);

    let added = 0;
    let duplicates = 0;
    const existing = new Set(state.recipients.map((r) => r.email.toLowerCase()));

    candidates.forEach((candidate) => {
      const email = (candidate.match(emailRegex) || [])[0];
      if (!email) {
        return;
      }

      const normalized = email.toLowerCase();
      if (existing.has(normalized)) {
        duplicates += 1;
        return;
      }

      existing.add(normalized);
      state.recipients.push({ email, name: '' });
      added += 1;
    });

    state.duplicateCount += duplicates;

    announceStatus(`${added} recipient${added === 1 ? '' : 's'} added from pasted list. ${duplicates ? `${duplicates} duplicate${duplicates === 1 ? '' : 's'} skipped.` : ''}`, added ? 'success' : 'warning');
    textarea.value = '';
    updateSummary();
    renderRecipientRows();
    updateDraftStatus();
  }

  function handleComposeChange() {
    state.subject = document.getElementById('subject').value.trim();
    state.cc = document.getElementById('cc').value.trim();
    state.bcc = document.getElementById('bcc').value.trim();
    state.body = document.getElementById('body').value;
    state.htmlBody = document.getElementById('htmlBody').value;
    state.useHtml = document.getElementById('useHtml').checked;
    updateDraftStatus();
  }

  function updateDraftStatus() {
    const ready = state.recipients.length > 0 && !!state.subject;
    const statusEl = document.getElementById('draftStatus');
    const metaEl = document.getElementById('draftMeta');

    if (!state.recipients.length) {
      statusEl.textContent = 'Awaiting recipients';
      metaEl.textContent = 'Load at least one email address to generate drafts.';
    } else if (!state.subject) {
      statusEl.textContent = 'Missing subject';
      metaEl.textContent = 'Add a subject so drafts can be generated cleanly.';
    } else {
      statusEl.textContent = 'Ready to generate';
      metaEl.textContent = 'Use the Update Drafts button or copy actions to personalize messages.';
    }

    document.getElementById('copyAllButton').disabled = !ready;
  }

  function updateSummary() {
    const count = state.recipients.length;
    document.getElementById('recipientCount').textContent = count;
    document.getElementById('recipientMeta').textContent = count ? 'Recipients are ready for individualized drafts.' : 'No recipients added yet.';
    document.getElementById('duplicateCount').textContent = state.duplicateCount;
  }

  function renderRecipientRows() {
    const container = document.getElementById('recipientRows');
    container.innerHTML = '';

    if (!state.recipients.length) {
      const empty = document.createElement('div');
      empty.className = 'table-row';
      empty.setAttribute('role', 'row');
      empty.innerHTML = '<div class="cell">No recipients yet.</div><div class="cell"></div>';
      container.appendChild(empty);
      return;
    }

    state.recipients.forEach((recipient, index) => {
      const row = document.createElement('div');
      row.className = 'table-row';
      row.setAttribute('role', 'row');

      const nameLabel = recipient.name ? `<span class="tag">${recipient.name}</span>` : '';
      const recipientCell = document.createElement('div');
      recipientCell.className = 'cell';
      recipientCell.innerHTML = `${nameLabel} <span>${recipient.email}</span>`;

      const actions = document.createElement('div');
      actions.className = 'cell table-actions';

      const openButton = document.createElement('a');
      openButton.className = 'btn btn-secondary';
      openButton.href = buildMailtoLink(recipient);
      openButton.textContent = 'Open draft';
      openButton.target = '_blank';
      openButton.rel = 'noopener noreferrer';

      const copyButton = document.createElement('button');
      copyButton.className = 'btn btn-primary';
      copyButton.textContent = 'Copy email';
      copyButton.type = 'button';
      copyButton.addEventListener('click', () => copySingleDraft(recipient));

      const removeButton = document.createElement('button');
      removeButton.className = 'btn btn-secondary';
      removeButton.textContent = 'Remove';
      removeButton.type = 'button';
      removeButton.addEventListener('click', () => removeRecipient(index));

      actions.appendChild(openButton);
      actions.appendChild(copyButton);
      actions.appendChild(removeButton);

      row.appendChild(recipientCell);
      row.appendChild(actions);
      container.appendChild(row);
    });
  }

  function buildMailtoLink(recipient) {
    const params = new URLSearchParams();
    if (state.subject) {
      params.set('subject', state.subject);
    }
    if (state.cc) {
      params.set('cc', state.cc);
    }
    if (state.bcc) {
      params.set('bcc', state.bcc);
    }

    const body = state.useHtml && state.htmlBody ? state.htmlBody : state.body;
    if (body) {
      params.set('body', body);
    }

    return `mailto:${encodeURIComponent(recipient.email)}?${params.toString()}`;
  }

  async function copySingleDraft(recipient) {
    const content = buildEmailText(recipient);
    try {
      await navigator.clipboard.writeText(content);
      announceStatus(`Draft for ${recipient.email} copied to clipboard.`, 'success');
    } catch (error) {
      console.error('Copy failed', error);
      announceStatus('Unable to copy to clipboard. Try manually selecting the text.', 'warning');
    }
  }

  function buildEmailText(recipient) {
    const lines = [
      `To: ${recipient.email}`,
      state.cc ? `CC: ${state.cc}` : null,
      state.bcc ? `BCC: ${state.bcc}` : null,
      state.subject ? `Subject: ${state.subject}` : null,
      '',
      state.useHtml && state.htmlBody ? state.htmlBody : state.body
    ].filter(Boolean);

    return lines.join('\n');
  }

  async function copyAllDrafts() {
    if (!state.recipients.length || !state.subject) {
      announceStatus('Add recipients and a subject before copying drafts.', 'warning');
      return;
    }

    const drafts = state.recipients.map((recipient, index) => `--- Draft ${index + 1}: ${recipient.email} ---\n${buildEmailText(recipient)}`).join('\n\n');

    try {
      await navigator.clipboard.writeText(drafts);
      announceStatus('All drafts copied to clipboard.', 'success');
    } catch (error) {
      console.error('Bulk copy failed', error);
      announceStatus('Unable to copy all drafts. Try copying a single draft instead.', 'warning');
    }
  }

  function removeRecipient(index) {
    state.recipients.splice(index, 1);
    updateSummary();
    renderRecipientRows();
    updateDraftStatus();
  }

  function clearRecipients() {
    state.recipients = [];
    state.duplicateCount = 0;
    renderRecipientRows();
    updateSummary();
    updateDraftStatus();
    announceStatus('Recipients cleared.', 'success');
  }

  function updateDrafts() {
    updateDraftStatus();
    renderRecipientRows();
    announceStatus('Drafts refreshed with the latest subject and body content.', 'success');
  }

  function resetForm() {
    document.getElementById('subject').value = '';
    document.getElementById('cc').value = '';
    document.getElementById('bcc').value = '';
    document.getElementById('body').value = '';
    document.getElementById('htmlBody').value = '';
    document.getElementById('useHtml').checked = false;
    handleComposeChange();
    announceStatus('Compose inputs reset. Recipients remain loaded.', 'success');
  }

  function announceStatus(message, variant = '') {
    const statusEl = document.getElementById('loadStatus');
    statusEl.textContent = message;
    statusEl.className = 'status-message';
    if (variant) {
      statusEl.classList.add(variant);
    }
  }

  return { init };
})();

window.addEventListener('DOMContentLoaded', EmailDraftTool.init);
