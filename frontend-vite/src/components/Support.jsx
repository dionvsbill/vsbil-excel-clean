// src/components/Support.jsx
import { useEffect, useState } from 'react';
import { apiGet, apiPost } from '../api/client';

export default function Support() {
  const [tickets, setTickets] = useState([]);
  const [subject, setSubject] = useState('');
  const [body, setBody] = useState('');
  const [response, setResponse] = useState('');
  const [ticketId, setTicketId] = useState('');

  const load = async () => {
    try {
      const res = await apiGet('/support/tickets');
      setTickets(res.tickets || []);
    } catch (e) {
      console.error(e);
    }
  };

  useEffect(() => { load(); }, []);

  const create = async () => {
    try {
      await apiPost('/support/tickets', { subject, body });
      setSubject(''); setBody('');
      await load();
    } catch (e) {
      alert(e.message);
    }
  };

  const respond = async () => {
    try {
      await apiPost('/support/tickets/respond', { ticket_id: ticketId, response });
      setTicketId(''); setResponse('');
      await load();
    } catch (e) {
      alert(e.message);
    }
  };

  return (
    <div className="grid">
      <section>
        <h3>Create Ticket</h3>
        <input placeholder="Subject" value={subject} onChange={(e) => setSubject(e.target.value)} />
        <textarea placeholder="Body" value={body} onChange={(e) => setBody(e.target.value)} />
        <button onClick={create}>Create</button>
      </section>
      <section>
        <h3>Respond</h3>
        <input placeholder="Ticket ID" value={ticketId} onChange={(e) => setTicketId(e.target.value)} />
        <textarea placeholder="Response" value={response} onChange={(e) => setResponse(e.target.value)} />
        <button onClick={respond}>Respond</button>
      </section>
      <section>
        <h3>Tickets</h3>
        <ul>
          {tickets.map(t => (
            <li key={t.id}>
              <strong>{t.subject}</strong> — {t.status} — by {t.user_id}
            </li>
          ))}
        </ul>
      </section>
    </div>
  );
}
