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
    }
  };
})();
