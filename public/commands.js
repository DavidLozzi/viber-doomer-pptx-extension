(function () {
  async function insertLiveResultsControl() {
    try {
      const slide = await ViberDoomer.ppt.getCurrentSlideContext();
      if (!slide) {
        throw new Error('Select a slide before inserting live results.');
      }
      await PowerPoint.run(async (context) => {
        const selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load('items');
        await context.sync();
        if (selectedSlides.items.length === 0) {
          throw new Error('No slide selected.');
        }
        const current = selectedSlides.items[0];
        const shapes = current.shapes;
        let placeholder;
        if (shapes.addGeometricShape) {
          const enumSet = PowerPoint && PowerPoint.GeometricShapeType;
          const rectangleEnum = enumSet && (enumSet.rectangle || enumSet.Rectangle || enumSet.roundRectangle);
          if (!rectangleEnum) {
            throw new Error('Unable to resolve rectangle geometry for placeholder.');
          }
          placeholder = shapes.addGeometricShape(rectangleEnum);
        } else if (shapes.addTextBox) {
          placeholder = shapes.addTextBox();
        } else {
          throw new Error('Current PowerPoint API does not support inserting shapes from add-ins.');
        }
        placeholder.left = 100;
        placeholder.top = 100;
        placeholder.width = 480;
        placeholder.height = 320;
        if (placeholder.textFrame && placeholder.textFrame.textRange) {
          placeholder.textFrame.textRange.text =
            '📊 Live Results placeholder. In slideshow, embedded add-in will render live data.';
        }
        placeholder.name = 'ViberDoomerLiveResults';
        await context.sync();
      });
      await ViberDoomer.storage.saveSlideMapping(slide.id, {
        slideId: slide.id,
        surveyId: slide.id,
      });
      await ViberDoomer.storage.bumpRefreshToken();
      if (Office.addin && Office.addin.showAsTaskpane) {
        Office.addin.showAsTaskpane();
      } else if (Office.context && Office.context.ui && Office.context.ui.displayDialogAsync) {
        let taskpaneUrl = 'taskpane.html';
        if (typeof window !== 'undefined' && window.location) {
          taskpaneUrl = new URL('taskpane.html', window.location.href).href;
        }
        Office.context.ui.displayDialogAsync(taskpaneUrl, { height: 40, width: 30 });
      }
    } catch (error) {
      showToast(error.message);
    }
  }

  async function syncSlide(event) {
    try {
      let slideContext;
      if (event && event.slideId) {
        slideContext = await ViberDoomer.ppt.getSlideContextById(event.slideId);
      } else {
        slideContext = await ViberDoomer.ppt.getCurrentSlideContext();
      }
      if (!slideContext) {
        return;
      }
      const mappings = ViberDoomer.storage.getSlideMappings();
      const mapping = mappings[slideContext.id] || { slideId: slideContext.id };
      const presentationId = ViberDoomer.storage.getPresentationId();

      ViberDoomer.throttling.scheduleSlideSync(async () => {
        try {
          await ViberDoomer.api.notifySlideChange(
            presentationId,
            mapping.slideId,
            slideContext.title || mapping.slideId
          );
        } catch (error) {
          showToast(error.message);
        }
      });
    } catch (error) {
      showToast(error.message);
    }
  }

  function showToast(message) {
    if (typeof window !== 'undefined' && window.alert) {
      window.alert(message);
    } else {
      console.log('Viber vs Doomer:', message);
    }
  }

  async function registerSlideChangeHandler() {
    try {
      await PowerPoint.run(async (context) => {
        if (context.presentation.onSlideSelectionChanged) {
          context.presentation.onSlideSelectionChanged.add(syncSlide);
        } else if (context.presentation.onSelectionChanged) {
          context.presentation.onSelectionChanged.add(syncSlide);
        }
        await context.sync();
      });
      if (Office.context && Office.context.document && Office.context.document.addHandlerAsync) {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, syncSlide);
      }
    } catch (error) {
      showToast(error.message);
    }
  }

  async function refreshAllData() {
    try {
      await ViberDoomer.storage.bumpRefreshToken();
      showToast('Refresh broadcast to all live results instances.');
    } catch (error) {
      showToast(error.message);
    }
  }

  const ribbon = {
    insertLiveResultsControl,
    refreshAllData,
  };

  Office.onReady(() => {
    registerSlideChangeHandler();
    syncSlide();
  });

  if (typeof module !== 'undefined') {
    module.exports = ribbon;
  } else {
    window.ViberDoomerCommands = ribbon;
  }
})();
