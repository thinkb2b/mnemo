import React from 'react';
import ReactDOM from 'react-dom/client';
// WICHTIG: Im Browser ohne Build-Tool müssen wir die Dateiendung (.tsx) angeben!
import App from './App.tsx';

const rootElement = document.getElementById('root');
if (rootElement) {
  const root = ReactDOM.createRoot(rootElement);
  root.render(<App />);
}