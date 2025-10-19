(function () {
  const STORAGE_NAMESPACE = 'viberDoomer';
  const DEFAULT_PRESENTATION_ID = 'boston-code-camp-2025';
  const REFRESH_THROTTLE_MS = 2000;

  function getDocumentSettings() {
    if (!Office || !Office.context || !Office.context.document) {
      throw new Error('Office context is not available.');
    }
    return Office.context.document.settings;
  }

  function getStorageKey(name) {
    return STORAGE_NAMESPACE + '.' + name;
  }

  async function saveSettingsAsync() {
    return new Promise((resolve, reject) => {
      const settings = getDocumentSettings();
      settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(result.error || new Error('Unable to save settings.'));
        }
      });
    });
  }

  function getPresentationId() {
    const settings = getDocumentSettings();
    return settings.get(getStorageKey('presentationId')) || DEFAULT_PRESENTATION_ID;
  }

  async function setPresentationId(id) {
    const settings = getDocumentSettings();
    settings.set(getStorageKey('presentationId'), id);
    await saveSettingsAsync();
    return id;
  }

  function getSlideMappings() {
    const settings = getDocumentSettings();
    return settings.get(getStorageKey('slideMappings')) || {};
  }

  async function setSlideMappings(mappings) {
    const settings = getDocumentSettings();
    settings.set(getStorageKey('slideMappings'), mappings);
    await saveSettingsAsync();
    return mappings;
  }

  async function saveSlideMapping(slideId, mapping) {
    const mappings = getSlideMappings();
    mappings[slideId] = mapping;
    await setSlideMappings(mappings);
    return mapping;
  }

  async function deleteSlideMapping(slideId) {
    const mappings = getSlideMappings();
    if (mappings[slideId]) {
      delete mappings[slideId];
      await setSlideMappings(mappings);
    }
  }

  function getRefreshToken() {
    const settings = getDocumentSettings();
    return settings.get(getStorageKey('refreshToken')) || 0;
  }

  async function bumpRefreshToken() {
    const settings = getDocumentSettings();
    settings.set(getStorageKey('refreshToken'), Date.now());
    await saveSettingsAsync();
  }

  async function getCurrentSlideContext() {
    if (!PowerPoint || !PowerPoint.run) {
      throw new Error('PowerPoint API is not available.');
    }
    return PowerPoint.run(async (context) => {
      const selectedSlides = context.presentation.getSelectedSlides();
      selectedSlides.load('items');
      await context.sync();
      if (selectedSlides.items.length === 0) {
        return null;
      }
      const slide = selectedSlides.items[0];
      slide.load('id,title');
      const indexResult = slide.getIndex ? slide.getIndex() : null;
      await context.sync();
      return {
        id: slide.id,
        title: slide.title,
        index: indexResult ? indexResult.value : undefined,
      };
    });
  }

  async function getSlideContextById(slideId) {
    if (!PowerPoint || !PowerPoint.run) {
      throw new Error('PowerPoint API is not available.');
    }
    return PowerPoint.run(async (context) => {
      const slide = context.presentation.slides.getItem(slideId);
      slide.load('id,title');
      const indexResult = slide.getIndex ? slide.getIndex() : null;
      await context.sync();
      return {
        id: slide.id,
        title: slide.title,
        index: indexResult ? indexResult.value : undefined,
      };
    });
  }

  async function postJson(url, payload) {
    const response = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });
    if (!response.ok) {
      const message = await response.text();
      throw new Error(`Request failed (${response.status}): ${message}`);
    }
    return response.json();
  }

  async function getJson(url) {
    const response = await fetch(url, {
      method: 'GET',
      headers: {
        'Content-Type': 'application/json',
      },
    });
    if (!response.ok) {
      const message = await response.text();
      throw new Error(`Request failed (${response.status}): ${message}`);
    }
    return response.json();
  }

  const apiClient = {
    async notifySlideChange(presentationId, slideId, title) {
      const payload = {
        presentation_id: presentationId,
        slide_id: slideId,
        title,
      };
      return postJson('https://api.agenticfeedback.app/slides/current', payload);
    },
    async fetchLiveResults(slideId) {
      return getJson(`https://api.agenticfeedback.app/slides/${encodeURIComponent(slideId)}/results`);
    },
    async fetchPresentationMetadata(presentationId) {
      return getJson(`https://api.agenticfeedback.app/presentations/${encodeURIComponent(presentationId)}`);
    },
  };

  const slideSyncState = {
    lastUpdate: 0,
    pendingTimeout: null,
  };

  function scheduleSlideSync(syncFn) {
    const now = Date.now();
    const elapsed = now - slideSyncState.lastUpdate;
    if (elapsed >= REFRESH_THROTTLE_MS) {
      slideSyncState.lastUpdate = now;
      Promise.resolve()
        .then(syncFn)
        .catch((error) => console.error('Viber vs Doomer:', error));
    } else {
      if (slideSyncState.pendingTimeout) {
        clearTimeout(slideSyncState.pendingTimeout);
      }
      slideSyncState.pendingTimeout = setTimeout(() => {
        slideSyncState.lastUpdate = Date.now();
        Promise.resolve()
          .then(syncFn)
          .catch((error) => console.error('Viber vs Doomer:', error));
      }, REFRESH_THROTTLE_MS - elapsed);
    }
  }

  window.ViberDoomer = {
    constants: {
      DEFAULT_PRESENTATION_ID,
      REFRESH_THROTTLE_MS,
    },
    storage: {
      getPresentationId,
      setPresentationId,
      getSlideMappings,
      setSlideMappings,
      saveSlideMapping,
      deleteSlideMapping,
      getRefreshToken,
      bumpRefreshToken,
    },
    ppt: {
      getCurrentSlideContext,
      getSlideContextById,
    },
    api: apiClient,
    throttling: {
      scheduleSlideSync,
    },
  };
})();
