/**
 * @file dragIsolation.ts
 * @description Isolates card boundary dragging and rich text resizing/editing
 * from SharePoint Online's native "Move Web Part" drag-and-drop canvas controller.
 */

let activeDragCleanup: (() => void) | undefined = undefined;

/**
 * Disables SharePoint's native canvas drag-and-drop move controller
 * while content within the full-width dashboard is actively being resized or edited.
 * 
 * @param disableMove True to lock SharePoint move web part handler, False to restore.
 * @param targetElement The DOM element initiating the resize or edit action.
 */
export const suppressSharePointWebPartDrag = (
  disableMove: boolean,
  targetElement?: HTMLElement
): void => {
  if (typeof window === 'undefined' || typeof document === 'undefined') return;

  if (disableMove) {
    // Clean up any existing suppressor first
    if (activeDragCleanup) {
      activeDragCleanup();
      activeDragCleanup = undefined;
    }

    // 1. Capture-phase handler to prevent any HTML5 drag events from bubbling to SharePoint Canvas
    const stopDragEvent = (e: DragEvent): void => {
      e.preventDefault();
      e.stopPropagation();
      e.stopImmediatePropagation();
    };

    window.addEventListener('dragstart', stopDragEvent, { capture: true, passive: false });

    // 2. Identify SharePoint Canvas Control & Section host wrappers
    const disabledElements: Array<{ el: HTMLElement; origDraggable: string | undefined; origUserSelect: string }> = [];

    // Traverse upwards to find SharePoint draggable zones
    let current: HTMLElement | undefined = targetElement ? (targetElement.parentElement || undefined) : undefined;
    while (current && current !== document.body) {
      if (
        current.getAttribute('draggable') === 'true' ||
        current.hasAttribute('data-automation-id') ||
        current.classList.contains('CanvasControl') ||
        current.classList.contains('CanvasSection') ||
        current.classList.contains('CanvasZone')
      ) {
        const rawAttr = current.getAttribute('draggable');
        disabledElements.push({
          el: current,
          origDraggable: rawAttr !== null ? rawAttr : undefined,
          origUserSelect: current.style.userSelect
        });
        current.setAttribute('draggable', 'false');
        current.style.userSelect = 'none';
      }
      current = current.parentElement || undefined;
    }

    // Set body user-select
    const prevBodyUserSelect = document.body.style.userSelect;
    document.body.style.userSelect = 'none';

    // Store cleanup callback
    activeDragCleanup = (): void => {
      window.removeEventListener('dragstart', stopDragEvent, { capture: true });
      disabledElements.forEach(({ el, origDraggable, origUserSelect }) => {
        if (origDraggable !== undefined) {
          el.setAttribute('draggable', origDraggable);
        } else {
          el.removeAttribute('draggable');
        }
        el.style.userSelect = origUserSelect;
      });
      document.body.style.userSelect = prevBodyUserSelect;
    };
  } else {
    // Restore
    if (activeDragCleanup) {
      activeDragCleanup();
      activeDragCleanup = undefined;
    }
  }
};
