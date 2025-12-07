// src/components/RealtimeSSE.jsx
import { useEffect, useRef } from 'react';

export default function RealtimeSSE({ token, apiBase, onEvent }) {
  const sseRef = useRef(null);

  useEffect(() => {
    if (!token) return;
    const url = `${apiBase}/realtime/events`;
    const es = new EventSource(url, { withCredentials: false });
    sseRef.current = es;

    const handler = (e) => {
      try {
        const { type } = e;
        const data = JSON.parse(e.data);
        onEvent?.(type, data);
      } catch {}
    };

    es.addEventListener('connected', handler);
    es.addEventListener('ping', handler);
    es.addEventListener('excel:add_sheet', handler);
    es.addEventListener('excel:delete_sheet', handler);
    es.addEventListener('excel:save_all', handler);
    es.addEventListener('excel:upload', handler);
    es.addEventListener('excel:convert', handler);
    es.addEventListener('admin:promote', handler);
    es.addEventListener('admin:denote', handler);
    es.addEventListener('admin:ban', handler);
    es.addEventListener('admin:soft_delete', handler);
    es.addEventListener('admin:permadelete', handler);
    es.addEventListener('support:session_created', handler);
    es.addEventListener('support:ticket_created', handler);
    es.addEventListener('support:ticket_responded', handler);
    es.addEventListener('pricing:update', handler);
    es.addEventListener('legal:terms_update', handler);
    es.addEventListener('legal:privacy_update', handler);
    es.addEventListener('payments:success', handler);
    es.addEventListener('payments:verified', handler);

    return () => {
      try { es.close(); } catch {}
      sseRef.current = null;
    };
  }, [token, apiBase, onEvent]);

  return null;
}
