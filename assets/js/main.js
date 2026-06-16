/* =========================================================================
   George Johnson — Portfolio interactions
   - Mobile nav toggle
   - Active link highlighting
   - Sticky header shadow on scroll
   - Image fallbacks (project thumbnails + company logos -> generated placeholders)
   - Project category filtering
   - Scroll reveal
   - Footer year
   ========================================================================= */
(function () {
  "use strict";

  /* ---- Mobile nav ---- */
  var toggle = document.querySelector(".nav-toggle");
  var links = document.getElementById("nav-links");
  if (toggle && links) {
    toggle.addEventListener("click", function () {
      var open = links.classList.toggle("open");
      toggle.setAttribute("aria-expanded", open ? "true" : "false");
    });
    links.querySelectorAll("a").forEach(function (a) {
      a.addEventListener("click", function () {
        links.classList.remove("open");
        toggle.setAttribute("aria-expanded", "false");
      });
    });
  }

  /* ---- Sticky header shadow ---- */
  var header = document.querySelector(".site-header");
  if (header) {
    var onScroll = function () {
      header.classList.toggle("scrolled", window.scrollY > 8);
    };
    onScroll();
    window.addEventListener("scroll", onScroll, { passive: true });
  }

  /* ---- Active nav link based on current page ---- */
  var path = location.pathname.split("/").pop() || "index.html";
  document.querySelectorAll(".nav-links a").forEach(function (a) {
    var href = (a.getAttribute("href") || "").split("/").pop();
    if (href === path || (path === "" && href === "index.html")) a.classList.add("active");
  });

  /* ---- Image fallback -> branded placeholder ---- */
  // Deterministic hue from a string so each project/company gets a stable colour.
  function hue(str) {
    var h = 0;
    for (var i = 0; i < str.length; i++) h = (h * 31 + str.charCodeAt(i)) % 360;
    return h;
  }
  function initials(str) {
    var parts = str.replace(/[^A-Za-z0-9 ]/g, " ").trim().split(/\s+/);
    if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }

  function buildThumbPlaceholder(img) {
    var label = img.getAttribute("data-label") || img.alt || "Project";
    var tag = img.getAttribute("data-tag") || "";
    var h = hue(label);
    var ph = document.createElement("div");
    ph.className = "thumb-ph";
    ph.style.background =
      "linear-gradient(140deg, hsl(" + h + ",62%,52%) 0%, hsl(" + ((h + 40) % 360) + ",58%,32%) 100%)";
    ph.innerHTML =
      '<span class="ph-mono">' + initials(label) + "</span>" +
      (tag ? '<span class="ph-tag">' + tag + "</span>" : "");
    img.replaceWith(ph);
  }

  function buildLogoPlaceholder(img) {
    var label = img.getAttribute("data-label") || img.alt || "Company";
    var h = hue(label);
    var mono = document.createElement("div");
    mono.className = "mono";
    mono.style.background =
      "linear-gradient(140deg, hsl(" + h + ",55%,48%), hsl(" + ((h + 35) % 360) + ",50%,30%))";
    mono.textContent = initials(label);
    img.replaceWith(mono);
  }

  function handleImg(img) {
    var kind = img.getAttribute("data-fallback");
    if (!kind) return;
    var fail = function () {
      if (img.dataset.handled) return;
      img.dataset.handled = "1";
      if (kind === "logo") buildLogoPlaceholder(img);
      else buildThumbPlaceholder(img);
    };
    if (img.complete && img.naturalWidth === 0) fail();      // already broken
    else { img.addEventListener("error", fail); }
  }
  document.querySelectorAll("img[data-fallback]").forEach(handleImg);

  /* ---- Project filtering ---- */
  var filterBar = document.querySelector(".filter-bar");
  if (filterBar) {
    var cards = Array.prototype.slice.call(document.querySelectorAll(".project-card"));
    filterBar.addEventListener("click", function (e) {
      var btn = e.target.closest(".filter-btn");
      if (!btn) return;
      filterBar.querySelectorAll(".filter-btn").forEach(function (b) { b.classList.remove("active"); });
      btn.classList.add("active");
      var f = btn.getAttribute("data-filter");
      cards.forEach(function (c) {
        var show = f === "all" || c.getAttribute("data-cat") === f;
        c.classList.toggle("hide", !show);
      });
    });
  }

  /* ---- Scroll reveal ---- */
  var revealEls = document.querySelectorAll(".reveal");
  if ("IntersectionObserver" in window && revealEls.length) {
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) {
        if (en.isIntersecting) { en.target.classList.add("is-visible"); io.unobserve(en.target); }
      });
    }, { threshold: 0.12, rootMargin: "0px 0px -40px 0px" });
    revealEls.forEach(function (el) { io.observe(el); });
  } else {
    revealEls.forEach(function (el) { el.classList.add("is-visible"); });
  }

  /* ---- Footer year ---- */
  var y = document.getElementById("year");
  if (y) y.textContent = new Date().getFullYear();
})();
