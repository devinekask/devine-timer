/* Slide Timer – a PowerPoint content add-in.
 *
 * Edit view:  digits + a config row (duration, presets, style). Saved per add-in instance in the .pptx.
 * Slideshow:  digits only. Click / Space / S = start-pause, double-click / R = reset.
 *             Arrow keys, PageUp/PageDown (presenter clickers) are forwarded to the slideshow.
 * Auto-start: optional; starts the countdown when the slide appears in the slideshow.
 */
(function () {
  "use strict";

  var DEFAULT_SECONDS = 300;
  var TICK_MS = 100;

  var el = {
    body: document.body,
    display: document.getElementById("display"),
    digits: document.getElementById("digits"),
    measure: document.getElementById("measure"),
    config: document.getElementById("config"),
    min: document.getElementById("min"),
    sec: document.getElementById("sec"),
    theme: document.getElementById("theme"),
    auto: document.getElementById("auto"),
    presets: document.querySelectorAll("[data-preset]"),
  };

  var inOffice = false;
  var view = null; // "edit" | "read" once known; the config row stays hidden until then

  var duration = DEFAULT_SECONDS; // seconds
  var remainingMs = duration * 1000; // only meaningful while not running
  var endAt = null; // timestamp (ms) when running, else null
  var ticker = null;
  var autoStart = false;

  // ---------------------------------------------------------------- storage
  // Office settings live inside the presentation, per add-in instance.
  // Outside Office (plain browser testing) fall back to localStorage.
  var store = {
    get: function (key) {
      if (inOffice) return Office.context.document.settings.get(key);
      try { return JSON.parse(localStorage.getItem("slide-timer:" + key)); } catch (e) { return null; }
    },
    set: function (key, value) {
      if (inOffice) {
        Office.context.document.settings.set(key, value);
        return;
      }
      try { localStorage.setItem("slide-timer:" + key, JSON.stringify(value)); } catch (e) { /* ignore */ }
    },
    save: debounce(function () {
      if (inOffice) Office.context.document.settings.saveAsync();
    }, 300),
  };

  // --------------------------------------------------------------- display
  function format(totalSeconds) {
    var h = Math.floor(totalSeconds / 3600);
    var m = Math.floor((totalSeconds % 3600) / 60);
    var s = totalSeconds % 60;
    var ss = s < 10 ? "0" + s : "" + s;
    if (h > 0 || duration >= 3600) return h + ":" + (m < 10 ? "0" + m : m) + ":" + ss;
    return m + ":" + ss;
  }

  function currentRemainingMs() {
    return endAt === null ? remainingMs : Math.max(0, endAt - Date.now());
  }

  function render() {
    var ms = currentRemainingMs();
    // Ceil so a 5:00 timer shows 5:00 at start and only hits 0:00 at the very end.
    el.digits.textContent = format(Math.ceil(ms / 1000));

    var running = endAt !== null;
    var done = ms <= 0;
    el.body.classList.toggle("running", running && !done);
    el.body.classList.toggle("paused", !running && !done && ms < duration * 1000);
    el.body.classList.toggle("done", done);
  }

  // Fit the digits to the box in px. (CSS container units are not re-evaluated reliably by
  // PowerPoint's Mac web view when the add-in is resized or scaled for the slideshow.)
  function fit() {
    var w = el.display.clientWidth;
    var h = el.display.clientHeight;
    if (!w || !h) return;
    // Measure the widest text this timer will show, so the size doesn't jump mid-countdown.
    el.measure.textContent = format(duration).replace(/\d/g, "0");
    var box = el.measure.getBoundingClientRect(); // rendered at 100px
    if (!box.width || !box.height) return;
    var size = Math.floor(100 * Math.min((w * 0.92) / box.width, (h * 0.88) / box.height));
    el.digits.style.fontSize = Math.max(8, size) + "px";
  }

  // ----------------------------------------------------------------- timer
  function start() {
    if (endAt !== null) return;
    if (remainingMs <= 0) remainingMs = duration * 1000; // restart after finishing
    endAt = Date.now() + remainingMs;
    ticker = setInterval(tick, TICK_MS);
    render();
  }

  function pause() {
    if (endAt === null) return;
    remainingMs = currentRemainingMs();
    endAt = null;
    clearInterval(ticker);
    render();
  }

  function reset() {
    endAt = null;
    clearInterval(ticker);
    remainingMs = duration * 1000;
    render();
  }

  function toggle() {
    if (endAt === null) start();
    else pause();
  }

  function tick() {
    if (currentRemainingMs() <= 0) {
      remainingMs = 0;
      endAt = null;
      clearInterval(ticker);
    }
    render();
  }

  // ---------------------------------------------------------------- config
  function setDuration(seconds, persist) {
    seconds = Math.max(1, Math.min(599 * 60 + 59, Math.round(seconds) || 0));
    duration = seconds;
    el.min.value = Math.floor(seconds / 60);
    el.sec.value = seconds % 60;
    el.presets.forEach(function (b) {
      b.setAttribute("aria-pressed", String(Number(b.dataset.preset) === seconds));
    });
    fit();
    reset();
    if (persist) {
      store.set("durationSeconds", seconds);
      store.save();
    }
  }

  function setTheme(theme, persist) {
    var legacy = { "dark-text": "light", "light-box": "light", "light-text": "dark", "dark-box": "dark" };
    var valid = ["light", "dark"];
    if (legacy[theme]) theme = legacy[theme];
    if (valid.indexOf(theme) === -1) theme = "light";
    valid.forEach(function (t) { el.body.classList.toggle("theme-" + t, t === theme); });
    el.theme.value = theme;
    if (persist) {
      store.set("theme", theme);
      store.save();
    }
  }

  function setAutoStart(on, persist) {
    autoStart = !!on;
    el.auto.checked = autoStart;
    if (persist) {
      store.set("autoStart", autoStart);
      store.save();
    }
  }

  function onConfigInput() {
    var m = parseInt(el.min.value, 10) || 0;
    var s = parseInt(el.sec.value, 10) || 0;
    setDuration(m * 60 + s, true);
  }

  // ------------------------------------------------------------------ view
  function setView(v) {
    v = v === "read" ? "read" : "edit";
    if (v === view) return;
    view = v;
    el.body.classList.toggle("view-read", view === "read");
    el.body.classList.toggle("view-edit", view === "edit");
    reset(); // every time the slideshow starts / ends, begin from the full duration
    fit(); // the config row appearing/disappearing changes the space for the digits
    if (view === "read" && autoStart) start();
  }

  function checkView() {
    if (!inOffice) return;
    Office.context.document.getActiveViewAsync(function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) setView(r.value);
    });
  }

  function goToSlide(direction) {
    if (!inOffice) return;
    var index = direction > 0 ? Office.Index.Next : Office.Index.Previous;
    Office.context.document.goToByIdAsync(index, Office.GoToType.Index);
  }

  function onKeyDown(e) {
    if (e.target && e.target.closest && e.target.closest("#config")) return; // typing in the config row

    switch (e.key) {
      case " ":
      case "Enter":
      case "s":
      case "S":
        e.preventDefault();
        toggle();
        return;
      case "r":
      case "R":
        e.preventDefault();
        reset();
        return;
    }

    // While presenting, the timer can steal focus from the slideshow after you click it.
    // Pass navigation keys (incl. presenter clickers, which send PageUp/PageDown) back.
    if (view !== "read") return;
    switch (e.key) {
      case "ArrowRight":
      case "ArrowDown":
      case "PageDown":
      case "n":
      case "N":
        e.preventDefault();
        goToSlide(1);
        break;
      case "ArrowLeft":
      case "ArrowUp":
      case "PageUp":
      case "p":
      case "P":
      case "Backspace":
        e.preventDefault();
        goToSlide(-1);
        break;
    }
  }

  // ------------------------------------------------------------------ init
  // Controls work immediately, even before (or if never) Office finishes loading.
  function bindControls() {
    el.display.addEventListener("click", toggle);
    el.display.addEventListener("dblclick", reset);
    document.addEventListener("keydown", onKeyDown);

    el.min.addEventListener("change", onConfigInput);
    el.sec.addEventListener("change", onConfigInput);
    el.config.addEventListener("submit", function (e) { e.preventDefault(); onConfigInput(); });
    el.theme.addEventListener("change", function () { setTheme(el.theme.value, true); });
    el.auto.addEventListener("change", function () { setAutoStart(el.auto.checked, true); });
    el.presets.forEach(function (b) {
      b.addEventListener("click", function () { setDuration(Number(b.dataset.preset), true); });
    });

    var onResize = debounce(checkView, 250); // entering/leaving the slideshow also resizes us
    if (typeof ResizeObserver !== "undefined") {
      new ResizeObserver(function () { fit(); onResize(); }).observe(el.display);
    }
    window.addEventListener("resize", function () { fit(); onResize(); });
    window.addEventListener("focus", checkView);
  }

  function init(officeHost) {
    inOffice = officeHost;

    setTheme(store.get("theme"), false);
    var saved = Number(store.get("durationSeconds"));
    setDuration(saved > 0 ? saved : DEFAULT_SECONDS, false);
    setAutoStart(store.get("autoStart"), false);

    if (!inOffice) {
      // Browser testing: ?view=read previews slideshow mode, ?s=3 sets a quick duration,
      // ?auto=1 enables auto-start.
      var params = new URLSearchParams(location.search);
      if (params.get("s")) setDuration(Number(params.get("s")), false);
      if (params.get("auto")) setAutoStart(params.get("auto") === "1", false);
      setView(params.get("view") || "edit");
      return;
    }

    checkView();
    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, function (e) {
      setView(e.activeView);
    });
  }

  function debounce(fn, ms) {
    var t;
    return function () {
      clearTimeout(t);
      t = setTimeout(fn, ms);
    };
  }

  bindControls();
  fit();
  render();
  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function (info) { init(!!(info && info.host)); });
  } else {
    init(false);
  }
})();
