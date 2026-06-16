(function (global) {
  const STORAGE_KEY = "operationalAnalytics.sidebarWidth";
  const DESKTOP_MIN_WIDTH = 1100;
  const MIN_WIDTH = 250;
  const MAX_WIDTH = 420;
  const MAX_VIEWPORT_RATIO = 0.42;

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    const handle = document.querySelector(".sidebar-resize-handle");

    if (!handle) {
      return;
    }

    applyStoredWidth();

    handle.addEventListener("pointerdown", function (event) {
      if (!isDesktop()) {
        return;
      }

      event.preventDefault();
      handle.setPointerCapture(event.pointerId);
      document.body.classList.add("is-sidebar-resizing");

      function handleMove(moveEvent) {
        setSidebarWidth(moveEvent.clientX, false);
      }

      function handleUp(upEvent) {
        handle.releasePointerCapture(upEvent.pointerId);
        document.body.classList.remove("is-sidebar-resizing");
        setSidebarWidth(upEvent.clientX, true);
        global.removeEventListener("pointermove", handleMove);
        global.removeEventListener("pointerup", handleUp);
      }

      global.addEventListener("pointermove", handleMove);
      global.addEventListener("pointerup", handleUp);
    });

    global.addEventListener("resize", applyStoredWidth);
  }

  function applyStoredWidth() {
    if (!isDesktop()) {
      return;
    }

    const storedWidth = readStoredWidth();

    if (storedWidth) {
      setSidebarWidth(storedWidth, false);
    }
  }

  function setSidebarWidth(width, shouldPersist) {
    const nextWidth = clampWidth(Number(width));

    document.documentElement.style.setProperty("--sidebar-width", nextWidth + "px");

    if (shouldPersist) {
      try {
        global.localStorage.setItem(STORAGE_KEY, String(nextWidth));
      } catch (_error) {
        // Width persistence is a convenience; resizing should keep working without storage.
      }
    }
  }

  function readStoredWidth() {
    try {
      const value = Number(global.localStorage.getItem(STORAGE_KEY));
      return Number.isFinite(value) ? value : 0;
    } catch (_error) {
      return 0;
    }
  }

  function clampWidth(width) {
    const viewportMax = Math.max(MIN_WIDTH, Math.min(MAX_WIDTH, Math.round(global.innerWidth * MAX_VIEWPORT_RATIO)));
    return Math.min(Math.max(Math.round(width), MIN_WIDTH), viewportMax);
  }

  function isDesktop() {
    return global.innerWidth >= DESKTOP_MIN_WIDTH;
  }
})(window);
