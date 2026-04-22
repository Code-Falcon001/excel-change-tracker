(function () {
  const environmentValue = document.getElementById("environmentValue");
  const officeValue = document.getElementById("officeValue");
  const statusValue = document.getElementById("statusValue");
  const logOutput = document.getElementById("logOutput");

  const startBtn = document.getElementById("startBtn");
  const stopBtn = document.getElementById("stopBtn");
  const testBtn = document.getElementById("testBtn");
  const clearLogBtn = document.getElementById("clearLogBtn");

  function now() {
    return new Date().toLocaleString();
  }

  function log(message) {
    const line = `[${now()}] ${message}`;
    if (logOutput.textContent.trim()) {
      logOutput.textContent += `\n${line}`;
    } else {
      logOutput.textContent = line;
    }
    logOutput.scrollTop = logOutput.scrollHeight;
    console.log(line);
  }

  function setText(el, value, cssClass) {
    if (!el) return;
    el.textContent = value;
    el.className = `value ${cssClass || ""}`.trim();
  }

  function initButtons() {
    startBtn.addEventListener("click", () => {
      setText(statusValue, "Tracking started", "ok");
      log("Start Tracking clicked.");
    });

    stopBtn.addEventListener("click", () => {
      setText(statusValue, "Tracking stopped", "warn");
      log("Stop Tracking clicked.");
    });

    testBtn.addEventListener("click", () => {
      log("Test Page clicked. UI is responsive.");
      alert("Change Tracker page is loading correctly.");
    });

    clearLogBtn.addEventListener("click", () => {
      logOutput.textContent = "";
    });
  }

  function initBrowserMode() {
    setText(environmentValue, "Browser / GitHub Pages", "ok");
    setText(officeValue, "Office.js loaded, waiting for host...", "warn");
    setText(statusValue, "Page loaded successfully", "ok");
    log("Page opened in a browser-safe mode.");
  }

  function initOfficeMode(info) {
    const host = info && info.host ? String(info.host) : "Unknown Host";
    const platform = info && info.platform ? String(info.platform) : "Unknown Platform";

    setText(environmentValue, `${host} (${platform})`, "ok");
    setText(officeValue, "Office host ready", "ok");
    setText(statusValue, "Ready inside Office", "ok");
    log(`Office.onReady triggered. Host=${host}, Platform=${platform}`);
  }

  function init() {
    initButtons();
    initBrowserMode();

    if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
      Office.onReady((info) => {
        initOfficeMode(info || {});
      });
    } else {
      setText(officeValue, "Office.js not available in this context", "warn");
      log("Office.js host context not detected. This is normal in a regular browser.");
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})();