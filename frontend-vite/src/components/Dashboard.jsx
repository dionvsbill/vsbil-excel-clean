// src/components/Dashboard.jsx
import { useEffect, useState } from 'react';
import RealtimeSSE from './RealtimeSSE';
import Sheets from './Sheets';
import Audit from './Audit';
import Payments from './Payments';
import OwnerPanel from './OwnerPanel';
import SuperadminPanel from './SuperadminPanel';
import { API_BASE } from '../api/client';

export default function Dashboard({ token, user }) {
  const [events, setEvents] = useState([]);
  const [role, setRole] = useState(user?.role || 'user');
  const [email, setEmail] = useState(user?.email || '');

  useEffect(() => {
    setRole(user?.role || 'user');
    setEmail(user?.email || '');
  }, [user]);

  const onEvent = (type, data) => {
    setEvents((prev) => [{ type, data, ts: Date.now() }, ...prev].slice(0, 100));
  };

  const isOwner = import.meta.env.VITE_OWNER_EMAIL ? email === import.meta.env.VITE_OWNER_EMAIL : false;
  const isSuperadmin = role === 'superadmin';

  return (
    <div className="dashboard">
      {(isOwner || isSuperadmin) && (
        <RealtimeSSE token={token} apiBase={API_BASE} onEvent={onEvent} />
      )}

      <header className="header">
        <h1>Admin Dashboard</h1>
        <span>{email} — {role}</span>
      </header>

      <div className="grid two-col">
        <section>
          <Sheets />
          <Audit />
          <Payments />
        </section>

        <section>
          {isOwner && <OwnerPanel />}
          {!isOwner && isSuperadmin && <SuperadminPanel />}

          <div className="card">
            <h2>Realtime events</h2>
            <ul className="event-list">
              {events.map((ev, idx) => (
                <li key={idx}>
                  <strong>{ev.type}</strong> — {new Date(ev.ts).toLocaleTimeString()}
                  <pre>{JSON.stringify(ev.data, null, 2)}</pre>
                </li>
              ))}
            </ul>
          </div>
        </section>
      </div>
    </div>
  );
}
