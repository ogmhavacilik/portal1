import {StrictMode} from 'react';
import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import { ErrorBoundary } from './components/ErrorBoundary.tsx';
import './index.css';

// Global error handler to prevent unhandled rejections and suppress benign WebSocket/WebAssembly warnings
if (typeof window !== 'undefined') {
  const isIgnoredMessage = (msg: any): boolean => {
    if (!msg) return false;
    const str = typeof msg === 'string' ? msg : (msg.message || msg.reason || String(msg));
    const lower = String(str).toLowerCase();
    return (
      lower.includes('websocket') ||
      lower.includes('webassembly') ||
      lower.includes('iş parçacığı') ||
      lower.includes('davet') ||
      lower.includes('script error') ||
      lower.includes('failed to connect to websocket') ||
      lower.includes('açılmadan kapatıldı')
    );
  };

  const origWarn = console.warn;
  console.warn = (...args: any[]) => {
    if (args.some(isIgnoredMessage)) return;
    origWarn.apply(console, args);
  };

  const origError = console.error;
  console.error = (...args: any[]) => {
    if (args.some(isIgnoredMessage)) return;
    origError.apply(console, args);
  };

  window.addEventListener('error', (event) => {
    if (isIgnoredMessage(event.message) || isIgnoredMessage(event.error)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      return true;
    }
  }, true);

  window.addEventListener('unhandledrejection', (event) => {
    if (isIgnoredMessage(event.reason)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      return true;
    }
  }, true);
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </StrictMode>,
);



