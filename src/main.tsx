import React from "react";
import ReactDOM from "react-dom/client";
import { App } from "./App";
import { checkAndHandleStaleCache } from "./services/storageService";
import "./styles/index.css";

// 1. Suppress harmless ResizeObserver loop warnings that Fluent UI v9 can trigger in narrow panes
if (typeof window !== "undefined") {
  window.addEventListener("error", (event) => {
    const msg = String(event.message || "");
    if (
      msg.includes("ResizeObserver loop completed with undelivered notifications") ||
      msg.includes("ResizeObserver loop limit exceeded")
    ) {
      event.stopImmediatePropagation();
      event.preventDefault();
    }
  });

  // 2. Auto-reload once if Vite detects a stale cached chunk after a GitHub Pages deploy
  window.addEventListener("vite:preloadError", () => {
    const reloadKey = "zee-ppt:preload-reloaded";
    if (!sessionStorage.getItem(reloadKey)) {
      sessionStorage.setItem(reloadKey, "1");
      window.location.reload();
    }
  });
}

const isReloading = checkAndHandleStaleCache();

function mountApp() {
  const rootEl = document.getElementById("root");
  if (!rootEl) return;
  ReactDOM.createRoot(rootEl).render(
    <React.StrictMode>
      <App />
    </React.StrictMode>
  );
}

if (!isReloading) {
  if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
    let mounted = false;
    Office.onReady()
      .then(() => {
        if (!mounted) {
          mounted = true;
          mountApp();
        }
      })
      .catch(() => {
        if (!mounted) {
          mounted = true;
          mountApp();
        }
      });

    // Fallback mount in case Office.onReady hangs in browser preview mode
    window.setTimeout(() => {
      if (!mounted) {
        mounted = true;
        mountApp();
      }
    }, 600);
  } else {
    mountApp();
  }
}
