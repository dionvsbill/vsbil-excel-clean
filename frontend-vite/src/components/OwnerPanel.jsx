// src/components/OwnerPanel.jsx
import { useState } from 'react';
import { apiPost } from '../api/client';
import UsersTable from './UsersTable';
import PricingControls from './PricingControls';
import Analytics from './Analytics';
import LegalEditor from './LegalEditor';

export default function OwnerPanel() {
  const [promoteId, setPromoteId] = useState('');
  const [denoteId, setDenoteId] = useState('');
  const [denoteRole, setDenoteRole] = useState('user');
  const [permaId, setPermaId] = useState('');

  const act = async (path, body) => {
    try {
      await apiPost(path, body);
      alert('Action success');
    } catch (e) {
      alert(e.message);
    }
  };

  return (
    <div className="grid">
      <section>
        <h2>Owner controls</h2>
        <div className="row">
          <input placeholder="User ID to promote" value={promoteId} onChange={(e) => setPromoteId(e.target.value)} />
          <button onClick={() => act('/admin/users/promote', { user_id: promoteId })}>Promote to superadmin</button>
        </div>
        <div className="row">
          <input placeholder="User ID to denote" value={denoteId} onChange={(e) => setDenoteId(e.target.value)} />
          <select value={denoteRole} onChange={(e) => setDenoteRole(e.target.value)}>
            <option value="user">user</option>
            <option value="admin">admin</option>
            <option value="superadmin">superadmin</option>
          </select>
          <button onClick={() => act('/admin/users/denote', { user_id: denoteId, role: denoteRole })}>Denote role</button>
        </div>
        <div className="row">
          <input placeholder="User ID to permanently delete" value={permaId} onChange={(e) => setPermaId(e.target.value)} />
          <button onClick={() => act('/admin/users/permadelete', { user_id: permaId })}>Permanent delete</button>
        </div>
      </section>

      <section>
        <h2>Users</h2>
        <UsersTable onAction={() => {}} />
      </section>

      <section>
        <h2>Pricing controls</h2>
        <PricingControls />
      </section>

      <section>
        <h2>Subscription analytics</h2>
        <Analytics />
      </section>

      <section>
        <h2>Legal pages</h2>
        <LegalEditor />
      </section>
    </div>
  );
}
