(function () {
  const root = document.documentElement;
  const fa = n => String(n).replace(/[0-9]/g, d => "۰۱۲۳۴۵۶۷۸۹"[d]);

  // ---------- theme ----------
  const themeBtn = document.querySelector("[data-theme-toggle]");
  const themeLabel = document.querySelector("[data-theme-label]");
  const themeIcon = document.querySelector(".theme-icon");

  function syncTheme(theme) {
    const dark = theme === "dark";
    root.dataset.theme = dark ? "dark" : "light";
    root.style.colorScheme = dark ? "dark" : "light";
    if (themeLabel) themeLabel.textContent = dark ? "حالت روشن" : "حالت تاریک";
    if (themeIcon) themeIcon.textContent = dark ? "☀" : "◐";
    if (themeBtn) themeBtn.setAttribute("aria-label", dark ? "تغییر به حالت روشن" : "تغییر به حالت تاریک");
  }

  syncTheme(root.dataset.theme === "dark" ? "dark" : "light");

  if (themeBtn) {
    themeBtn.addEventListener("click", () => {
      const next = root.dataset.theme === "dark" ? "light" : "dark";
      syncTheme(next);
      try { localStorage.setItem("guide-theme", next); } catch (e) {}
    });
  }

  // ---------- presenter filter ----------
  const filterBtns = Array.from(document.querySelectorAll("[data-filter]"));
  const cards = Array.from(document.querySelectorAll(".slide-card"));
  const tocItems = Array.from(document.querySelectorAll(".toc li"));
  const mapRows = Array.from(document.querySelectorAll(".overview tbody tr"));
  const countEl = document.querySelector("[data-filter-count]");
  const links = Array.from(document.querySelectorAll(".toc a"));
  let spy;

  function applyFilter(filter) {
    cards.forEach(card => {
      card.hidden = filter !== "all" && card.dataset.speaker !== filter;
    });
    tocItems.forEach(item => {
      item.hidden = filter !== "all" && item.dataset.speaker !== filter;
    });
    mapRows.forEach(row => {
      row.hidden = filter !== "all" && row.dataset.speaker !== filter;
    });
    filterBtns.forEach(btn => {
      const active = btn.dataset.filter === filter;
      btn.classList.toggle("is-active", active);
      btn.setAttribute("aria-pressed", active ? "true" : "false");
    });
    if (countEl) {
      const n = cards.filter(card => !card.hidden).length;
      countEl.textContent = filter === "all" ? "۲۰ اسلاید" : `${fa(n)} اسلاید`;
    }
    setupSpy();
  }

  filterBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      const filter = btn.dataset.filter;
      applyFilter(filter);
      try { localStorage.setItem("guide-filter", filter); } catch (e) {}
    });
  });

  let savedFilter = "all";
  try { savedFilter = localStorage.getItem("guide-filter") || "all"; } catch (e) {}
  if (!["all", "diana", "mozhan"].includes(savedFilter)) savedFilter = "all";
  applyFilter(savedFilter);

  // ---------- scroll spy ----------
  function setupSpy() {
    if (spy) spy.disconnect();
    const visibleCards = cards.filter(card => !card.hidden);
    if (!("IntersectionObserver" in window) || !visibleCards.length) return;
    const byId = new Map(links.map(a => [a.getAttribute("href").slice(1), a]));
    spy = new IntersectionObserver(entries => {
      entries.forEach(entry => {
        if (!entry.isIntersecting) return;
        links.forEach(a => a.classList.remove("active"));
        const link = byId.get(entry.target.id);
        if (link) link.classList.add("active");
      });
    }, { rootMargin: "-15% 0px -65% 0px", threshold: 0.01 });
    visibleCards.forEach(card => spy.observe(card));
  }

  setupSpy();

  document.querySelectorAll("[data-print]").forEach(btn => {
    btn.addEventListener("click", () => window.print());
  });
})();
