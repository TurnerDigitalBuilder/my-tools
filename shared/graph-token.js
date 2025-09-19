const GraphTokenManager = (() => {
  const STORAGE_KEY = 'graphAccessToken';
  const listeners = new Set();
  const storageAreas = { session: null, local: null };
  let token = '';
  let textarea = null;
  let persist = true;
  let storageMode = 'session';

  function resolveStorage(mode) {
    if (!storageAreas.hasOwnProperty(mode)) {
      return null;
    }

    if (storageAreas[mode]) {
      return storageAreas[mode];
    }

    try {
      storageAreas[mode] = mode === 'local' ? window.localStorage : window.sessionStorage;
    } catch (error) {
      console.warn('GraphTokenManager: Storage access unavailable.', error);
      storageAreas[mode] = null;
    }

    return storageAreas[mode];
  }

  function readStoredToken(mode) {
    const storage = resolveStorage(mode);
    if (!storage) {
      return '';
    }

    try {
      return storage.getItem(STORAGE_KEY) || '';
    } catch (error) {
      console.warn('GraphTokenManager: Unable to read stored token.', error);
      return '';
    }
  }

  function writeStoredToken(value, mode) {
    const storage = resolveStorage(mode);
    if (!storage) {
      return;
    }

    try {
      if (value) {
        storage.setItem(STORAGE_KEY, value);
      } else {
        storage.removeItem(STORAGE_KEY);
      }
    } catch (error) {
      console.warn('GraphTokenManager: Unable to persist token.', error);
    }
  }

  function notifyListeners() {
    listeners.forEach(callback => {
      try {
        callback(token);
      } catch (error) {
        console.error('GraphTokenManager listener error:', error);
      }
    });
  }

  function updateToken(newValue, { fromInput = false } = {}) {
    token = newValue.trim();

    if (persist) {
      writeStoredToken(token, storageMode);
    }

    if (!fromInput && textarea) {
      textarea.value = token;
    }

    notifyListeners();
  }

  return {
    initialize(options = {}) {
      const {
        textareaId = 'graphToken',
        onTokenChange,
        persistToken = true,
        restoreFromStorage = true,
        storage = 'session'
      } = options;

      persist = persistToken;
      storageMode = storage === 'local' ? 'local' : 'session';

      if (typeof onTokenChange === 'function') {
        listeners.add(onTokenChange);
      }

      textarea = document.getElementById(textareaId);

      if (restoreFromStorage && persist) {
        const stored = readStoredToken(storageMode);
        if (stored) {
          token = stored;
          if (textarea) {
            textarea.value = token;
          }
        }
      }

      if (textarea) {
        textarea.addEventListener('input', event => {
          updateToken(event.target.value, { fromInput: true });
        });
      }

      notifyListeners();
    },

    subscribe(callback, { immediate = true } = {}) {
      if (typeof callback !== 'function') {
        return () => {};
      }

      listeners.add(callback);
      if (immediate) {
        callback(token);
      }

      return () => listeners.delete(callback);
    },

    getToken() {
      return token;
    },

    setToken(value) {
      updateToken(value, { fromInput: false });
    },

    clearToken() {
      updateToken('', { fromInput: false });
      if (persist) {
        writeStoredToken('', storageMode);
      }
    }
  };
})();
