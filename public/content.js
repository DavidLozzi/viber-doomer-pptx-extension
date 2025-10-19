(function () {
  const REFRESH_INTERVAL = 2000;
  let modeIndicator;
  let chartArea;
  let summaryArea;
  let pollingHandle = null;
  let lastRefreshToken = 0;

  function renderPlaceholder(slideInfo, reason) {
    chartArea.innerHTML = '';
    const placeholder = document.createElement('div');
    placeholder.className = 'placeholder';
    if (!slideInfo) {
      placeholder.textContent = reason || 'Live results will appear here during slideshow.';
    } else {
      placeholder.textContent = reason || `📊 Live Results for ${slideInfo.title || slideInfo.id}`;
    }
    chartArea.appendChild(placeholder);
    if (summaryArea) {
      summaryArea.textContent = reason || '';
    }
  }

  function renderModeIndicator(isPresentation) {
    modeIndicator.textContent = isPresentation
      ? 'Presenting — fetching live results.'
      : 'Edit Mode — placeholder shown.';
  }

  function renderChart(results) {
    chartArea.innerHTML = '';
    const total = results.total_responses || 0;
    if (!results.results || results.results.length === 0) {
      renderPlaceholder(null, 'No responses yet. The chart will update automatically.');
      return {
        total,
        updatedAt: results.updated_at || Date.now(),
      };
    }
    const maxCount = Math.max(...results.results.map((entry) => entry.count), 1);
    for (const entry of results.results) {
      const row = document.createElement('div');
      row.className = 'chart-bar';

      const label = document.createElement('div');
      label.className = 'chart-bar-label';
      label.textContent = entry.option;

      const track = document.createElement('div');
      track.className = 'chart-bar-track';

      const fill = document.createElement('div');
      fill.className = 'chart-bar-fill';
      fill.style.width = `${Math.round((entry.count / maxCount) * 100)}%`;

      const value = document.createElement('div');
      value.className = 'chart-bar-value';
      value.textContent = `${entry.count}`;

      track.appendChild(fill);
      row.appendChild(label);
      row.appendChild(track);
      row.appendChild(value);
      chartArea.appendChild(row);
    }

    return {
      total,
      updatedAt: results.updated_at || Date.now(),
    };
  }

  async function loadSlideMapping() {
    try {
      const slide = await ViberDoomer.ppt.getCurrentSlideContext();
      if (!slide) {
        return null;
      }
      const mappings = ViberDoomer.storage.getSlideMappings();
      const mapping = mappings[slide.id] || { slideId: slide.id, surveyId: slide.id };
      return {
        slide,
        mapping,
      };
    } catch (error) {
      console.error('Viber vs Doomer:', error);
      return null;
    }
  }

  async function fetchAndRender() {
    try {
      const context = await loadSlideMapping();
      if (!context) {
        renderPlaceholder(null, 'Select this slide during slideshow to display live results.');
        return;
      }
      const { slide, mapping } = context;
      const results = await ViberDoomer.api.fetchLiveResults(mapping.surveyId || mapping.slideId);
      const chartMeta = renderChart(results);
      if (chartMeta) {
        summaryArea.textContent = `${results.total_responses} responses · ${slide.title || mapping.slideId} · Updated ${new Date(
          chartMeta.updatedAt
        ).toLocaleTimeString()}`;
      }
    } catch (error) {
      renderPlaceholder(null, error.message);
    }
  }

  function startPolling() {
    if (pollingHandle) {
      clearInterval(pollingHandle);
    }
    pollingHandle = setInterval(async () => {
      try {
        const currentToken = ViberDoomer.storage.getRefreshToken();
        if (currentToken !== lastRefreshToken) {
          lastRefreshToken = currentToken;
          await fetchAndRender();
          return;
        }
        await fetchAndRender();
      } catch (error) {
        console.error('Viber vs Doomer:', error);
      }
    }, REFRESH_INTERVAL);
  }

  async function initialize() {
    modeIndicator = document.getElementById('mode-indicator');
    chartArea = document.getElementById('chart-area');
    summaryArea = document.getElementById('summary');

    const isPresentation =
      Office &&
      Office.context &&
      Office.context.document &&
      Office.context.document.mode === Office.DocumentMode.Read;
    renderModeIndicator(isPresentation);

    if (!isPresentation) {
      const context = await loadSlideMapping();
      renderPlaceholder(context ? context.slide : null);
      summaryArea.textContent = 'Enter presentation mode to load live results automatically.';
      return;
    }

    lastRefreshToken = ViberDoomer.storage.getRefreshToken();
    await fetchAndRender();
    startPolling();
  }

  Office.onReady(() => {
    initialize().catch((error) => {
      if (summaryArea) {
        summaryArea.textContent = error.message;
      } else {
        console.error('Viber vs Doomer:', error);
      }
    });
  });
})();
