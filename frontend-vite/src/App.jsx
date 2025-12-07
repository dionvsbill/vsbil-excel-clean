// src/App.jsx
import React, { useEffect, useState } from 'react';
import ReactDOM from 'react-dom/client';
import { BrowserRouter, Routes, Route, Navigate, useNavigate, Link } from 'react-router-dom';
import { supabase } from './supabaseClient';
import * as serviceWorkerRegistration from './serviceWorkerRegistration';
import { ErrorBoundary } from './components/ErrorBoundary.jsx';

import Dashboard from './pages/Dashboard.jsx';
import Support from './components/Support.jsx';
import Analytics from './components/Analytics.jsx';
import Pricing from './components/PricingControls.jsx';
import Legal from './components/LegalEditor.jsx';

import './index.css';

function App() {
  const [session, setSession] = useState(null);
  const [role, setRole] = useState('user');
  const navigate = useNavigate();

  useEffect(() => {
    // Ping backend
    const apiBase = import.meta.env.VITE_API_URL;
    if (apiBase) {
      fetch(`${apiBase}/ping`)
        .then(res => res.json())
        .then(data => console.log('Backend says:', data.message))
        .catch(err => console.error('Backend error:', err.message));
    }

    // Supabase session logic
    supabase.auth.getSession()
      .then(({ data }) => {
        setSession(data.session);
        if (data.session?.user) {
          fetchUserRole(data.session.user.id);
        }
      })
      .catch(err => console.error('Supabase error:', err));

    const { data: listener } = supabase.auth.onAuthStateChange((_event, newSession) => {
      setSession(newSession);
      if (newSession?.user) {
        fetchUserRole(newSession.user.id);
      }
      // If admin, land on dashboard
      navigate('/dashboard');
    });

    return () => {
      listener.subscription.unsubscribe();
    };
  }, [navigate]);

  const fetchUserRole = async (userId) => {
    const { data: profile } = await supabase
      .from('profiles')
      .select('role')
      .eq('id', userId)
      .single();
    if (profile?.role) {
      setRole(profile.role);
    }
  };

  const logout = async () => {
    await supabase.auth.signOut();
    setSession(null);
    setRole('user');
    navigate('/');
  };

  return (
    <>
      <nav className="nav">
        <Link to="/dashboard">Dashboard</Link>
        <Link to="/support">Support</Link>
        <Link to="/analytics">Analytics</Link>
        <Link to="/PricingControls">Pricing</Link>
        <Link to="/legalEditor">Legal</Link>
        {session ? (
          <button onClick={logout}>Logout</button>
        ) : (
          <Link to="/dashboard">Login</Link>
        )}
      </nav>
      <Routes>
        {/* Default landing: if admin, go to dashboard, else redirect */}
        <Route
          path="/"
          element={
            role === 'admin'
              ? <Navigate to="/dashboard" replace />
              : <Navigate to="/support" replace />
          }
        />
        <Route path="/dashboard" element={<Dashboard session={session} role={role} />} />
        <Route path="/support" element={<Support session={session} role={role} />} />
        <Route path="/analytics" element={<Analytics session={session} role={role} />} />
        <Route path="/PricingControls" element={<Pricing session={session} role={role} />} />
        <Route path="/legalEditor" element={<Legal session={session} role={role} />} />
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </>
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
