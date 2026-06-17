// Storage shim: provides window.storage backed by localStorage
// so the same code works in Claude artifacts AND on Netlify/standalone browsers.
// Data persists in the user's browser localStorage and survives refreshes/closes.

if (typeof window !== 'undefined' && !window.storage) {
  window.storage = {
    get: async (key) => {
      try {
        const value = localStorage.getItem(key);
        return value !== null ? { key, value, shared: false } : null;
      } catch (e) {
        return null;
      }
    },
    set: async (key, value) => {
      try {
        localStorage.setItem(key, value);
        return { key, value, shared: false };
      } catch (e) {
        console.error('Storage set failed', e);
        return null;
      }
    },
    delete: async (key) => {
      try {
        localStorage.removeItem(key);
        return { key, deleted: true, shared: false };
      } catch (e) {
        return null;
      }
    },
    list: async (prefix = '') => {
      try {
        const keys = [];
        for (let i = 0; i < localStorage.length; i++) {
          const k = localStorage.key(i);
          if (k && k.startsWith(prefix)) keys.push(k);
        }
        return { keys, prefix, shared: false };
      } catch (e) {
        return { keys: [], prefix, shared: false };
      }
    },
  };
}
