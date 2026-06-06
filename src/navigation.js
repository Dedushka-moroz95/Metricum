(function (global) {
  const App = (global.OperationalAnalytics = global.OperationalAnalytics || {});
  const LINK_SELECTOR = ".sidebar-link[href^='#']";
  const ACTIVE_CLASS = "is-active";
  const visibleSections = new Map();
  let links = [];
  let sections = [];
  let observer = null;
  let scrollTicking = false;

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    links = Array.from(document.querySelectorAll(LINK_SELECTOR));
    sections = getUniqueSections(links);

    if (!links.length || !sections.length) {
      return;
    }

    links.forEach(function (link) {
      link.addEventListener("click", function () {
        const section = getSectionFromLink(link);

        if (section) {
          setActiveSection(section.id);
        }
      });
    });

    if ("IntersectionObserver" in global) {
      observer = new IntersectionObserver(handleIntersection, {
        root: null,
        rootMargin: "-18% 0px -58% 0px",
        threshold: [0, 0.12, 0.28, 0.5],
      });

      sections.forEach(function (section) {
        observer.observe(section);
      });
    } else {
      global.addEventListener("scroll", scheduleFallbackUpdate, { passive: true });
    }

    setActiveSection(getInitialSectionId());
    scheduleFallbackUpdate();
  }

  function getUniqueSections(navLinks) {
    const seen = new Set();

    return navLinks
      .map(getSectionFromLink)
      .filter(function (section) {
        if (!section || seen.has(section.id)) {
          return false;
        }

        seen.add(section.id);
        return true;
      });
  }

  function getSectionFromLink(link) {
    const hash = link.getAttribute("href");

    if (!hash || hash.charAt(0) !== "#") {
      return null;
    }

    return document.getElementById(hash.slice(1));
  }

  function handleIntersection(entries) {
    entries.forEach(function (entry) {
      if (entry.isIntersecting) {
        visibleSections.set(entry.target.id, entry.intersectionRatio);
        return;
      }

      visibleSections.delete(entry.target.id);
    });

    const activeSection = getBestVisibleSection() || getNearestSection();

    if (activeSection) {
      setActiveSection(activeSection.id);
    }
  }

  function getBestVisibleSection() {
    if (!visibleSections.size) {
      return null;
    }

    return Array.from(visibleSections.keys())
      .map(function (id) {
        return document.getElementById(id);
      })
      .filter(Boolean)
      .sort(compareSectionPosition)[0] || null;
  }

  function getNearestSection() {
    return sections.slice().sort(compareSectionPosition)[0] || null;
  }

  function compareSectionPosition(left, right) {
    const offset = getNavigationOffset();
    const leftDistance = Math.abs(left.getBoundingClientRect().top - offset);
    const rightDistance = Math.abs(right.getBoundingClientRect().top - offset);

    return leftDistance - rightDistance;
  }

  function getNavigationOffset() {
    if (global.innerWidth < 761) {
      return 250;
    }

    return global.innerWidth < 1100 ? 260 : 32;
  }

  function setActiveSection(sectionId) {
    let activated = false;

    links.forEach(function (link) {
      const targetId = (link.getAttribute("href") || "").slice(1);
      const shouldActivate = !activated && targetId === sectionId;

      link.classList.toggle(ACTIVE_CLASS, shouldActivate);

      if (shouldActivate) {
        link.setAttribute("aria-current", "page");
      } else {
        link.removeAttribute("aria-current");
      }

      if (shouldActivate) {
        activated = true;
        keepActiveLinkVisible(link);
      }
    });
  }

  function keepActiveLinkVisible(link) {
    const nav = link.closest(".sidebar-nav");

    if (!nav || nav.scrollWidth <= nav.clientWidth) {
      return;
    }

    link.scrollIntoView({
      behavior: "smooth",
      block: "nearest",
      inline: "nearest",
    });
  }

  function getInitialSectionId() {
    const hash = global.location.hash ? global.location.hash.slice(1) : "";
    const hashSection = hash ? document.getElementById(hash) : null;

    if (hashSection) {
      return hashSection.id;
    }

    return sections[0] ? sections[0].id : "";
  }

  function scheduleFallbackUpdate() {
    if (scrollTicking) {
      return;
    }

    scrollTicking = true;
    global.requestAnimationFrame(function () {
      scrollTicking = false;
      const activeSection = getNearestSection();

      if (activeSection) {
        setActiveSection(activeSection.id);
      }
    });
  }

  App.Navigation = {
    init,
  };
})(window);
