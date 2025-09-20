const AppUI = (() => {
  const THEME_KEY = 'appThemePreference';
  let themeStorage = null;

  function getStorage() {
    if (themeStorage) {
      return themeStorage;
    }

    try {
      themeStorage = window.localStorage;
    } catch (error) {
      console.warn('AppUI: Unable to access localStorage for theme persistence.', error);
      themeStorage = null;
    }

    return themeStorage;
  }

  function persistTheme(isDark) {
    const storage = getStorage();
    if (!storage) {
      return;
    }

    try {
      storage.setItem(THEME_KEY, isDark ? 'dark' : 'light');
    } catch (error) {
      console.warn('AppUI: Unable to persist theme preference.', error);
    }
  }

  function readStoredTheme() {
    const storage = getStorage();
    if (!storage) {
      return null;
    }

    try {
      return storage.getItem(THEME_KEY);
    } catch (error) {
      console.warn('AppUI: Unable to read theme preference.', error);
      return null;
    }
  }

  function applyTheme(theme) {
    if (theme === 'dark') {
      document.body.classList.add('dark-mode');
    } else {
      document.body.classList.remove('dark-mode');
    }
  }

  function resolveElement(target) {
    if (!target) {
      return null;
    }

    if (typeof target === 'string') {
      return document.getElementById(target);
    }

    if (target instanceof HTMLElement) {
      return target;
    }

    return null;
  }

  function formatTokenPreview(token, { start = 12, end = 6 } = {}) {
    if (typeof token !== 'string' || !token.length) {
      return '';
    }

    if (token.length <= start + end + 1) {
      return token;
    }

    return `${token.slice(0, start)}…${token.slice(-end)}`;
  }

  return {
    initialize(options = {}) {
      const { defaultTheme = 'light', rememberTheme = true } = options;

      if (!rememberTheme) {
        themeStorage = null;
      }

      const storedTheme = rememberTheme ? readStoredTheme() : null;
      const themeToApply = storedTheme || defaultTheme;
      applyTheme(themeToApply);
    },

    toggleDarkMode() {
      const isDark = !document.body.classList.contains('dark-mode');
      applyTheme(isDark ? 'dark' : 'light');
      persistTheme(isDark);
      return isDark;
    },

    setDarkMode(isDark) {
      applyTheme(isDark ? 'dark' : 'light');
      persistTheme(isDark);
    },

    toggleSection(sectionId) {
      const section = document.getElementById(sectionId);
      if (!section) {
        return;
      }

      const shouldCollapse = section.style.display !== 'none';
      this.setSectionCollapsed(sectionId, shouldCollapse);
    },

    setSectionCollapsed(sectionId, collapsed) {
      const section = document.getElementById(sectionId);
      const icon = document.getElementById(`${sectionId}Icon`);
      if (!section || !icon) {
        return;
      }

      if (collapsed) {
        section.style.display = 'none';
        icon.textContent = '▶';
      } else {
        section.style.display = 'block';
        icon.textContent = '▼';
      }
    },

    updateTokenStatus(target, token, options = {}) {
      const element = resolveElement(target);
      if (!element) {
        return;
      }

      const {
        emptyMessage = 'No token stored yet.',
        prefix = 'Token stored',
        start = 12,
        end = 6
      } = options;

      if (!token) {
        element.textContent = emptyMessage;
        return;
      }

      const preview = formatTokenPreview(token, { start, end });
      element.textContent = `${prefix} (${token.length} characters, preview: ${preview}).`;
    }
  };
})();
