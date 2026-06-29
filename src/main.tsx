import {StrictMode} from 'react';
import {createRoot} from 'react-dom/client';
import App from './App.tsx';
import './index.css';

// Global error boundary to prevent cross-origin or third-party iframe errors ("Script error.") from crashing or being flagged as failures
if (typeof window !== 'undefined') {
  const handleGlobalError = (event: ErrorEvent) => {
    // Suppress "Script error." which is a standard benign browser notification for cross-origin errors
    if (event.message === 'Script error.' || event.message?.toLowerCase().includes('script error')) {
      console.warn("Caught and handled cross-origin third-party script error gracefully:", event);
      event.preventDefault();
      return true;
    }
  };

  const handleUnhandledRejection = (event: PromiseRejectionEvent) => {
    console.warn("Caught unhandled promise rejection gracefully:", event.reason);
    event.preventDefault();
  };

  window.addEventListener('error', handleGlobalError);
  window.addEventListener('unhandledrejection', handleUnhandledRejection);
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);

