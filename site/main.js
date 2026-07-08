/* 도하커피 — interactions
   Vanilla JS, no dependencies. Everything degrades gracefully
   and respects prefers-reduced-motion. */
(function () {
  'use strict';

  var reduce = window.matchMedia &&
    window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  /* ---- current year ---- */
  var yearEl = document.querySelector('[data-year]');
  if (yearEl) yearEl.textContent = new Date().getFullYear();

  /* ---- sticky nav shadow ---- */
  var nav = document.querySelector('[data-nav]');
  function onScroll() {
    if (nav) nav.classList.toggle('is-stuck', window.scrollY > 24);
  }
  onScroll();
  window.addEventListener('scroll', onScroll, { passive: true });

  /* ---- mobile drawer ---- */
  var toggle = document.querySelector('[data-menu-toggle]');
  var drawer = document.querySelector('[data-drawer]');
  if (toggle && drawer) {
    function closeDrawer() {
      drawer.hidden = true;
      toggle.setAttribute('aria-expanded', 'false');
      toggle.setAttribute('aria-label', '메뉴 열기');
    }
    toggle.addEventListener('click', function () {
      var open = toggle.getAttribute('aria-expanded') === 'true';
      if (open) { closeDrawer(); return; }
      drawer.hidden = false;
      toggle.setAttribute('aria-expanded', 'true');
      toggle.setAttribute('aria-label', '메뉴 닫기');
    });
    drawer.querySelectorAll('a').forEach(function (a) {
      a.addEventListener('click', closeDrawer);
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') closeDrawer();
    });
  }

  /* ---- count-up ---- */
  function countUp(el) {
    var target = parseInt(el.getAttribute('data-count'), 10);
    var suffix = el.getAttribute('data-suffix') || '';
    if (reduce || !target) { el.textContent = target + suffix; return; }
    var dur = 1100, start = null;
    function frame(t) {
      if (start === null) start = t;
      var p = Math.min((t - start) / dur, 1);
      var eased = 1 - Math.pow(1 - p, 3);
      el.textContent = Math.round(target * eased) + suffix;
      if (p < 1) requestAnimationFrame(frame);
    }
    requestAnimationFrame(frame);
  }

  /* ---- roast meter fill ---- */
  function litRoast(el) {
    var level = parseInt(el.getAttribute('data-roast'), 10) || 0;
    var bars = el.querySelectorAll('.roast__meter i');
    el.classList.add('is-lit');
    bars.forEach(function (b, i) {
      setTimeout(function () {
        if (i <= level) b.classList.add('on'); // roast level -> lit bars
      }, reduce ? 0 : i * 90);
    });
  }

  /* ---- spectrum marker ---- */
  function fillSpectrum(el) {
    var marker = el.querySelector('[data-spectrum-marker]');
    if (marker) marker.style.left = reduce ? '82%' : '82%';
  }

  /* ---- signature curve ---- */
  var curve = document.querySelector('.curve');
  var curvePath = document.querySelector('.curve__path');
  if (curvePath) {
    try {
      var len = curvePath.getTotalLength();
      curvePath.style.strokeDasharray = len;
      curvePath.style.strokeDashoffset = reduce ? 0 : len;
    } catch (e) { /* older browsers: leave visible */ }
  }
  function drawCurve() {
    if (curve) curve.classList.add('is-drawn');
  }

  /* ---- IntersectionObserver reveal + triggers ---- */
  var revealEls = document.querySelectorAll('[data-reveal]');

  if (!('IntersectionObserver' in window)) {
    // fallback: show everything, run triggers immediately
    revealEls.forEach(function (el) { el.classList.add('is-visible'); });
    document.querySelectorAll('[data-count]').forEach(countUp);
    document.querySelectorAll('.roast').forEach(litRoast);
    document.querySelectorAll('[data-spectrum]').forEach(fillSpectrum);
    drawCurve();
    return;
  }

  var io = new IntersectionObserver(function (entries) {
    entries.forEach(function (entry) {
      if (!entry.isIntersecting) return;
      var el = entry.target;
      el.classList.add('is-visible');

      el.querySelectorAll('[data-count]').forEach(countUp);
      el.querySelectorAll('.roast').forEach(litRoast);
      if (el.matches('[data-spectrum]') || el.querySelector('[data-spectrum]')) {
        var sp = el.matches('[data-spectrum]') ? el : el.querySelector('[data-spectrum]');
        if (sp) fillSpectrum(sp);
      }
      if (el.classList.contains('curve')) drawCurve();

      io.unobserve(el);
    });
  }, { threshold: 0.18, rootMargin: '0px 0px -8% 0px' });

  revealEls.forEach(function (el) { io.observe(el); });

  // Hero reveals: stagger slightly for a page-load sequence
  var heroReveals = document.querySelectorAll('.hero [data-reveal]');
  heroReveals.forEach(function (el, i) {
    el.style.transitionDelay = reduce ? '0ms' : (i * 90) + 'ms';
  });
})();
