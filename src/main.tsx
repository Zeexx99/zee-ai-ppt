import React from 'react';
import ReactDOM from 'react-dom/client';
import { App } from './App';
import './styles/index.css';

/* global Office */

// 1. Filter harmless ResizeObserver loop warnings
window.addEventListener('error', e => {
  const msg = (e && e.message) ? e.message.toLowerCase() : '';
  if (
    msg.includes('resizeobserver loop') ||
    msg.includes('resizeobserver loop completed with undelivered notifications')
  ) {
    e.stopImmediatePropagation();
    e.preventDefault();
  }
});

let isMounted = false;
let root: ReactDOM.Root | null = null;

function renderApp(isOfficeReady: boolean) {
  const container = document.getElementById('root');
  if (!container) return;

  if (!root) {
    root = ReactDOM.createRoot(container);
  }
  root.render(
    <React.StrictMode>
      <App isOfficeReady={isOfficeReady} />
    </React.StrictMode>
  );
  isMounted = true;
}

// 2. Repeated Office readiness check to fix the "Preview mode" banner bug
let attempts = 0;
const maxAttempts = 8;
const checkInterval = 300;

function checkOfficeReadiness() {
  if (typeof Office !== 'undefined' && typeof Office.onReady === 'function') {
    Office.onReady(info => {
      if (info && info.host === Office.HostType.PowerPoint) {
        renderApp(true);
      } else if (attempts < maxAttempts) {
        attempts++;
        setTimeout(checkOfficeReadiness, checkInterval);
      } else {
        // Fallback: render in browser preview mode
        renderApp(false);
      }
    });
  } else if (attempts < maxAttempts) {
    attempts++;
    setTimeout(checkOfficeReadiness, checkInterval);
  } else {
    renderApp(false);
  }
}

// Initial readiness check
checkOfficeReadiness();
