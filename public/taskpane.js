(function () {
  let currentSlideContext = null;
  let statusElement;
  let slideContextElement;
  let presentationIdInput;
  let slideIdInput;
  let surveyIdInput;
  let presentationMetaElement;
  let slideTableElement;

  function setStatus(message, isError) {
    if (!statusElement) {
      return;
    }
    statusElement.textContent = message;
    statusElement.classList.toggle('error', Boolean(isError));
  }

  function renderSlideTable(mappings) {
    slideTableElement.innerHTML = '';
    const entries = Object.entries(mappings);
    if (entries.length === 0) {
      const row = document.createElement('tr');
      const cell = document.createElement('td');
      cell.colSpan = 3;
      cell.className = 'placeholder';
      cell.textContent = 'No slides configured yet.';
      row.appendChild(cell);
      slideTableElement.appendChild(row);
      return;
    }

    entries.sort((a, b) => a[0].localeCompare(b[0]));
    for (const [slideKey, mapping] of entries) {
      const row = document.createElement('tr');
      const slideCell = document.createElement('td');
      const slideIdCell = document.createElement('td');
      const surveyIdCell = document.createElement('td');

      slideCell.textContent = slideKey;
      slideIdCell.textContent = mapping.slideId || '—';
      surveyIdCell.textContent = mapping.surveyId || '—';

      row.appendChild(slideCell);
      row.appendChild(slideIdCell);
      row.appendChild(surveyIdCell);
      slideTableElement.appendChild(row);
    }
  }

  async function refreshSlideContext() {
    try {
      currentSlideContext = await ViberDoomer.ppt.getCurrentSlideContext();
    } catch (error) {
      setStatus(error.message, true);
      return;
    }

    if (!currentSlideContext) {
      slideContextElement.textContent = 'No slide selected. Select a slide to configure mapping.';
      slideIdInput.value = '';
      surveyIdInput.value = '';
      return;
    }

    const { id, title, index } = currentSlideContext;
    const slideLabel = typeof index === 'number' ? `Slide ${index + 1}` : 'Slide';
    slideContextElement.textContent = `${slideLabel}: ${title || id}`;
    const mappings = ViberDoomer.storage.getSlideMappings();
    const mapping = mappings[id] || {};
    slideIdInput.value = mapping.slideId || id;
    surveyIdInput.value = mapping.surveyId || mapping.slideId || id;
  }

  async function loadSettings() {
    presentationIdInput.value = ViberDoomer.storage.getPresentationId();
    renderSlideTable(ViberDoomer.storage.getSlideMappings());
    await refreshSlideContext();
  }

  async function handleSavePresentation() {
    const value = presentationIdInput.value.trim();
    if (!value) {
      setStatus('Presentation ID cannot be empty.', true);
      return;
    }
    try {
      await ViberDoomer.storage.setPresentationId(value);
      setStatus(`Presentation ID saved as "${value}".`);
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  async function handleFetchMetadata() {
    const presentationId = presentationIdInput.value.trim();
    if (!presentationId) {
      setStatus('Set a presentation ID first.', true);
      return;
    }
    presentationMetaElement.style.display = 'block';
    presentationMetaElement.textContent = 'Fetching presentation metadata...';
    try {
      const metadata = await ViberDoomer.api.fetchPresentationMetadata(presentationId);
      presentationMetaElement.textContent = metadata.slides
        .map((slide, index) => `${index + 1}. ${slide.title} (${slide.slide_id})`)
        .join('\n');
    } catch (error) {
      presentationMetaElement.textContent = error.message;
    }
  }

  async function handleSaveSlide() {
    if (!currentSlideContext) {
      setStatus('No slide selected.', true);
      return;
    }
    const slideId = slideIdInput.value.trim();
    const surveyId = surveyIdInput.value.trim();
    if (!slideId) {
      setStatus('Slide ID cannot be empty.', true);
      return;
    }
    const mapping = {
      slideId,
      surveyId: surveyId || slideId,
    };
    try {
      await ViberDoomer.storage.saveSlideMapping(currentSlideContext.id, mapping);
      await ViberDoomer.storage.bumpRefreshToken();
      renderSlideTable(ViberDoomer.storage.getSlideMappings());
      setStatus(`Mapping saved for slide ${currentSlideContext.id}.`);
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  async function handleDeleteSlide() {
    if (!currentSlideContext) {
      setStatus('No slide selected.', true);
      return;
    }
    try {
      await ViberDoomer.storage.deleteSlideMapping(currentSlideContext.id);
      await ViberDoomer.storage.bumpRefreshToken();
      renderSlideTable(ViberDoomer.storage.getSlideMappings());
      slideIdInput.value = currentSlideContext.id;
      surveyIdInput.value = currentSlideContext.id;
      setStatus(`Mapping removed for slide ${currentSlideContext.id}.`);
    } catch (error) {
      setStatus(error.message, true);
    }
  }

  async function initialize() {
    statusElement = document.getElementById('status');
    slideContextElement = document.getElementById('slide-context');
    presentationIdInput = document.getElementById('presentation-id');
    slideIdInput = document.getElementById('slide-id');
    surveyIdInput = document.getElementById('survey-id');
    presentationMetaElement = document.getElementById('presentation-meta');
    slideTableElement = document.getElementById('slide-table');

    document.getElementById('save-presentation').addEventListener('click', handleSavePresentation);
    document.getElementById('refresh-presentation').addEventListener('click', handleFetchMetadata);
    document.getElementById('save-slide').addEventListener('click', handleSaveSlide);
    document.getElementById('delete-slide').addEventListener('click', handleDeleteSlide);

    setStatus('Loading settings...');
    await loadSettings();
    if (Office.context && Office.context.document && Office.context.document.addHandlerAsync) {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
        refreshSlideContext();
      });
    }
    setStatus('Ready. Configure slide mappings and go present!');
  }

  Office.onReady(() => {
    initialize().catch((error) => {
      setStatus(error.message, true);
    });
  });
})();
