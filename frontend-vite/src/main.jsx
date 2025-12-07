import React, { useEffect, useState } from 'react';
import ReactDOM from 'react-dom/client';
import { BrowserRouter, Routes, Route, useNavigate, Navigate } from 'react-router-dom';
import { supabase } from './supabaseClient';
import * as serviceWorkerRegistration from './serviceWorkerRegistration';
import Dashboard from './pages/Dashboard.jsx';
import { ErrorBoundary } from './components/ErrorBoundary.jsx';
import Payments from './components/Payments';
import Support from './components/Support.jsx';
import Analytics from './components/Analytics.jsx';
import Pricing from './components/PricingControls.jsx';
import Legal from './components/LegalEditor.jsx';
import mainDashboard from './components/Dashboard.jsx';

import './index.css';


function App() {
  const [session, setSession] = useState(null);
  const navigate = useNavigate();

  useEffect(() => {
    // Optional: ping backend
    const apiBase = import.meta.env.VITE_API_URL;
    if (apiBase) {
      fetch(`${apiBase}/ping`)
        .then(res => res.json())
        .then(data => console.log('Backend says:', data.message))
        .catch(err => console.error('Backend error:', err.message));
    }

    // Supabase session logic
    supabase.auth.getSession()
      .then(({ data }) => setSession(data.session))
      .catch(err => console.error('Supabase error:', err));

    const { data: listener } = supabase.auth.onAuthStateChange((_event, newSession) => {
      setSession(newSession);
      navigate('/'); // always return to dashboard
    });

    return () => {
      listener.subscription.unsubscribe();
    };
  }, [navigate]);

  return (
    <Routes>
      <Route path="/" element={<Dashboard session={session} />} />
      <Route path="/index" element={<Navigate to="/" replace />} />
      <Route path="*" element={<Navigate to="/" replace />} />
      <Route path="/payments" element={<Payments />} />
    </Routes>
  );
}

// Render root
ReactDOM.createRoot(document.getElementById('root')).render(
  <BrowserRouter>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </BrowserRouter>
);

// Enables caching and offline
serviceWorkerRegistration.register();

export default App;
