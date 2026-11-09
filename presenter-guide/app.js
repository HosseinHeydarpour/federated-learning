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

  // ---------- podcasts ----------
  const podBtns = Array.from(document.querySelectorAll("[data-podcast]"));
  const podBox = document.querySelector("[data-podcast-player]");
  const podAudio = podBox ? podBox.querySelector("audio") : null;
  const podMeta = document.querySelector("[data-podcast-meta]");
  const PODCASTS = {
    diana: { src: "audio/diana.m4a", label: "بخش دیانا — اسلایدهای ۱–۴، ۸–۱۰ و ۱۵–۱۷ · ۱۴:۲۶" },
    mozhan: { src: "audio/mozhan.m4a", label: "بخش موژان — اسلایدهای ۵–۷، ۱۱–۱۴ و ۱۸–۲۰ · ۱۶:۲۰" },
    all: { src: "audio/all.m4a", label: "کل ارائه — هر ۲۰ اسلاید · ۱۳:۴۳" },
  };

  function choosePodcast(who, opts) {
    const item = PODCASTS[who];
    if (!item || !podBtns.length) return;
    opts = opts || {};
    podBtns.forEach(btn => {
      const on = btn.dataset.podcast === who;
      btn.classList.toggle("is-active", on);
      btn.setAttribute("aria-pressed", on ? "true" : "false");
    });
    if (podAudio) {
      if (podAudio.getAttribute("src") !== item.src) podAudio.src = item.src;
      if (opts.play) podAudio.play().catch(() => {});
    }
    if (podMeta) podMeta.textContent = item.label;
    if (podBox) podBox.hidden = false;
    if (opts.filter !== false) applyFilter(who);
    try { localStorage.setItem("guide-podcast", who); } catch (e) {}
  }

  podBtns.forEach(btn => {
    btn.addEventListener("click", () => choosePodcast(btn.dataset.podcast, { play: true }));
  });
  window.guideChoosePodcast = choosePodcast;

  let savedPodcast = null;
  try { savedPodcast = localStorage.getItem("guide-podcast"); } catch (e) {}
  if (savedPodcast && savedPodcast === savedFilter && PODCASTS[savedPodcast]) {
    choosePodcast(savedPodcast, { filter: false });
  }

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

// ---------- enlarge a slide ----------
(function () {
  if (!window.Fancybox) return;
  window.Fancybox.bind('[data-fancybox="slides"]', {
    animated: true,
    Images: { zoom: true },
    Toolbar: { display: { left: ["infobar"], middle: [], right: ["zoomIn", "zoomOut", "toggle1to1", "close"] } },
  });
})();

// ---------- read / unread marks ----------
(function () {
  const KEY = "guide-read";
  const fa = n => String(n).replace(/[0-9]/g, d => "۰۱۲۳۴۵۶۷۸۹"[d]);
  const btns = Array.from(document.querySelectorAll("[data-read]"));
  const textEl = document.querySelector("[data-read-text]");
  const barEl = document.querySelector("[data-read-bar]");
  const read = new Set();

  try {
    const saved = JSON.parse(localStorage.getItem(KEY) || "[]");
    if (Array.isArray(saved)) saved.forEach(n => read.add(Number(n)));
  } catch (e) {}

  function sync() {
    btns.forEach(btn => {
      const n = Number(btn.dataset.read);
      const on = read.has(n);
      btn.classList.toggle("is-read", on);
      btn.setAttribute("aria-pressed", on ? "true" : "false");
      btn.textContent = on ? "خوانده شد ✓" : "خواندم";
      const card = btn.closest(".slide-card");
      if (card) card.classList.toggle("is-read", on);
    });
    const total = btns.length;
    const done = btns.filter(b => read.has(Number(b.dataset.read))).length;
    if (textEl) textEl.textContent = `${fa(done)} از ${fa(total)} اسلاید خوانده شده`;
    if (barEl) barEl.style.width = total ? (done / total) * 100 + "%" : "0%";
    try { localStorage.setItem(KEY, JSON.stringify([...read])); } catch (e) {}
  }

  btns.forEach(btn => {
    btn.addEventListener("click", () => {
      const n = Number(btn.dataset.read);
      read.has(n) ? read.delete(n) : read.add(n);
      sync();
    });
  });

  sync();
})();

// ---------- guided tour ----------
(function () {
  const KEY = "guide-tour-done";
  const fa = n => String(n).replace(/[0-9]/g, d => "۰۱۲۳۴۵۶۷۸۹"[d]);
  const steps = [
    {
      sel: '[data-tour="podcast-order"]',
      title: "به چه ترتیبی گوش کنم؟",
      text: "سه نسخه هست؛ این ترتیب بیشترین فایده را دارد:",
      list: [
        "۱) «کل ارائه» (۱۳:۴۳) را اول گوش کن تا کل مسیر، داستان و نقطهٔ تحویل را یک‌جا ببینی.",
        "۲) بعد بخش خودت را گوش کن: دیانا ۱۴:۲۶ یا موژان ۱۶:۲۰؛ هم‌زمان راهنما و اسلایدها را جلو ببر.",
        "۳) اگر وقت داشتی، آخر بخش هم‌گروهی‌ات را هم بشنو تا بدانی حرف را کجا تحویل می‌گیری.",
      ],
    },
    {
      sel: '[data-tour="podcasts"]',
      title: "اول، پادکست خودت را گوش کن",
      text: "قبل از خواندن راهنما، نسخهٔ صوتی بخش خودت را گوش کن. بگو دیانا هستی، موژان، یا کل ارائه را می‌خواهی؛ پخش‌کننده باز می‌شود و فیلتر صفحه هم روی همان بخش می‌ماند.",
      choice: true,
    },
    {
      sel: '[data-tour="filters"]',
      interactive: true,
      title: "این صفحه چه کاری می‌کند؟",
      text: "راهنمای اسلاید به اسلاید ارائه است. روی «دیانا» یا «موژان» بزن تا فقط اسلایدهای خودت را ببینی؛ تور سر جایش می‌ماند.",
    },
    {
      sel: '[data-tour="files"]',
      title: "سه فایل برای دانلود",
      text: "پاورپوینت ارائه، بستهٔ فونت‌ها و PDF راهنما. هر سه از همین صفحه دانلود می‌شوند.",
    },
    {
      sel: '[data-tour="deck"]',
      title: "۱ · پاورپوینت",
      text: "۲۰ اسلاید آمادهٔ پخش؛ همین فایل را در سالن باز می‌کنید. زمان‌بندی ۳۰ دقیقه در یادداشت اسلایدهاست.",
    },
    {
      sel: '[data-tour="fonts"]',
      title: "۲ · فونت‌ها",
      text: "B Nazanin و B Titr. قبل از ارائه نصبشان کنید؛ بدون این دو، متن اسلایدها در سیستم شما بد باز می‌شود.",
    },
    {
      sel: '[data-tour="pdf"]',
      title: "۳ · PDF راهنما",
      text: "نسخهٔ چاپی ۲۲ صفحه‌ای برای تمرین و خواندن روی کاغذ. اگر لپ‌تاپ ندارید، همین را چاپ کنید.",
    },
  ];

  const startBtn = document.querySelector("[data-tour-start]");
  let layer, spot, card, stepEl, titleEl, textEl, prevBtn, nextBtn, dotsEl, choiceEl, listEl;
  let idx = 0;
  let lastFocus = null;

  function build() {
    layer = document.createElement("div");
    layer.className = "tour-layer no-print";
    layer.hidden = true;
    layer.innerHTML =
      '<div class="tour-spot"></div>' +
      '<div class="tour-card" role="dialog" aria-modal="true" aria-labelledby="tour-title" tabindex="-1">' +
      '<div class="tour-head"><span class="tour-step"></span>' +
      '<button class="tour-close" type="button" aria-label="بستن تور">×</button></div>' +
      '<h3 id="tour-title"></h3><p></p>' +
      '<ul class="tour-list" hidden></ul>' +
      '<div class="tour-choice" hidden>' +
      '<button type="button" data-tour-pick="diana">من دیانا هستم</button>' +
      '<button type="button" data-tour-pick="mozhan">من موژان هستم</button>' +
      '<button type="button" data-tour-pick="all">کل ارائه</button>' +
      '</div>' +
      '<div class="tour-actions"><button class="tour-skip" type="button">رد کردن</button>' +
      '<span class="tour-dots"></span>' +
      '<button class="tour-prev" type="button">قبلی</button>' +
      '<button class="tour-next" type="button">بعدی</button></div>' +
      "</div>";
    document.body.appendChild(layer);

    spot = layer.querySelector(".tour-spot");
    card = layer.querySelector(".tour-card");
    stepEl = layer.querySelector(".tour-step");
    titleEl = layer.querySelector("#tour-title");
    textEl = layer.querySelector(".tour-card p");
    prevBtn = layer.querySelector(".tour-prev");
    nextBtn = layer.querySelector(".tour-next");
    dotsEl = layer.querySelector(".tour-dots");
    choiceEl = layer.querySelector(".tour-choice");
    listEl = layer.querySelector(".tour-list");

    dotsEl.innerHTML = steps.map(() => "<span></span>").join("");
    syncDots();

    choiceEl.querySelectorAll("[data-tour-pick]").forEach(btn => {
      btn.addEventListener("click", () => {
        if (window.guideChoosePodcast) window.guideChoosePodcast(btn.dataset.tourPick, { play: true });
        show(idx + 1);
      });
    });

    nextBtn.addEventListener("click", () => (idx === steps.length - 1 ? close() : show(idx + 1)));
    prevBtn.addEventListener("click", () => show(idx - 1));
    layer.querySelector(".tour-skip").addEventListener("click", close);
    layer.querySelector(".tour-close").addEventListener("click", close);
    layer.addEventListener("click", e => {
      if (e.target !== layer) return;
      const st = steps[idx];
      if (st.interactive) {
        const r = spot.getBoundingClientRect();
        const inside =
          e.clientX >= r.left && e.clientX <= r.right &&
          e.clientY >= r.top && e.clientY <= r.bottom;
        if (inside) {
          // let the page element under the spotlight receive the click
          layer.style.pointerEvents = "none";
          const hit = document.elementFromPoint(e.clientX, e.clientY);
          layer.style.pointerEvents = "";
          const btn = hit && hit.closest && hit.closest("button");
          if (btn) {
            btn.click();
            requestAnimationFrame(place);
            return;
          }
        }
      }
      close();
    });
    window.addEventListener("resize", place);
    window.addEventListener("scroll", () => { if (!layer.hidden) place(); }, { passive: true });
  }

  function syncDots() {
    Array.from(dotsEl.children).forEach((d, i) => d.classList.toggle("on", i === idx));
  }

  function place() {
    if (!layer || layer.hidden) return;
    const el = document.querySelector(steps[idx].sel);
    if (!el) return close();
    const r = el.getBoundingClientRect();
    const pad = 8;
    const top = r.top - pad;
    const left = r.left - pad;
    const w = r.width + pad * 2;
    const h = r.height + pad * 2;
    spot.style.top = top + "px";
    spot.style.left = left + "px";
    spot.style.width = w + "px";
    spot.style.height = h + "px";

    const vw = window.innerWidth;
    const vh = window.innerHeight;
    const cw = card.offsetWidth;
    const ch = card.offsetHeight;
    let t = top + h + 14;
    if (t + ch > vh - 10) t = top - ch - 14;
    if (t < 10) t = Math.max(10, vh - ch - 10);
    let l = left + w / 2 - cw / 2;
    l = Math.min(Math.max(10, l), vw - cw - 10);
    card.style.top = t + "px";
    card.style.left = l + "px";
  }

  function show(i) {
    idx = Math.max(0, Math.min(steps.length - 1, i));
    const el = document.querySelector(steps[idx].sel);
    if (!el) return close();

    stepEl.textContent = "گام " + fa(idx + 1) + " از " + fa(steps.length);
    titleEl.textContent = steps[idx].title;
    textEl.textContent = steps[idx].text;
    prevBtn.hidden = idx === 0;
    nextBtn.textContent = idx === steps.length - 1 ? "پایان" : "بعدی";
    if (choiceEl) choiceEl.hidden = !steps[idx].choice;
    const list = steps[idx].list || [];
    if (listEl) {
      listEl.innerHTML = list.map(item => "<li>" + item + "</li>").join("");
      listEl.hidden = !list.length;
    }
    syncDots();

    const prevBehavior = document.documentElement.style.scrollBehavior;
    document.documentElement.style.scrollBehavior = "auto";
    el.scrollIntoView({ block: "center", inline: "nearest" });
    document.documentElement.style.scrollBehavior = prevBehavior;
    place();
  }

  function open(i) {
    if (!layer) build();
    lastFocus = document.activeElement;
    layer.hidden = false;
    document.body.classList.add("tour-on");
    show(i);
    card.focus({ preventScroll: true });
  }

  function close() {
    if (!layer) return;
    layer.hidden = true;
    document.body.classList.remove("tour-on");
    try { localStorage.setItem(KEY, "1"); } catch (e) {}
    if (lastFocus && lastFocus.focus) lastFocus.focus({ preventScroll: true });
  }

  document.addEventListener("keydown", e => {
    if (!layer || layer.hidden) return;
    if (e.key === "Escape") {
      e.preventDefault();
      close();
    } else if (e.key === "ArrowLeft") {
      e.preventDefault();
      idx === steps.length - 1 ? close() : show(idx + 1);
    } else if (e.key === "ArrowRight") {
      e.preventDefault();
      show(idx - 1);
    }
  });

  if (startBtn) startBtn.addEventListener("click", () => open(0));

  let done = false;
  try { done = localStorage.getItem(KEY) === "1"; } catch (e) {}
  const printable = window.matchMedia && window.matchMedia("print").matches;
  if (!done && !printable) setTimeout(() => open(0), 900);
})();
