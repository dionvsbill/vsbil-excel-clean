// src/components/UsersTable.jsx
import { useEffect, useState } from 'react';
import { apiGet, apiPost } from '../api/client';

export default function UsersTable({ onAction }) {
  const [users, setUsers] = useState([]);
  const [q, setQ] = useState('');
  const [loading, setLoading] = useState(false);

  const load = async () => {
    setLoading(true);
    try {
      const res = await apiGet(`/admin/users/list?search=${encodeURIComponent(q)}`);
      setUsers(res.users || []);
    } catch (e) {
      console.error(e);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const act = async (path, body) => {
    try {
      await apiPost(path, body);
      onAction?.(path, body);
      await load();
    } catch (e) {
      alert(e.message);
    }
  };

  return (
    <div className="card">
      <div className="row mb-2">
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Search by email" />
        <button onClick={load} disabled={loading}>Search</button>
      </div>
      <table className="table">
        <thead>
          <tr>
            <th>Email</th><th>Role</th><th>Plan</th><th>Status</th><th>Verified</th><th>Actions</th>
          </tr>
        </thead>
        <tbody>
          {users.map(u => (
            <tr key={u.id}>
              <td>{u.email}</td>
              <td>{u.role}</td>
              <td>{u.plan}</td>
              <td>{u.status || 'active'}</td>
              <td>{u.verified ? 'Yes' : 'No'}</td>
              <td>
                <button onClick={() => act('/admin/users/ban', { user_id: u.id, reason: 'Violation' })}>Ban</button>
                <button onClick={() => act('/admin/users/verify', { user_id: u.id })}>Verify</button>
                <button onClick={() => act('/admin/users/soft-delete', { user_id: u.id })}>Soft delete</button>
                <button onClick={() => act('/admin/support/assist', { user_id: u.id })}>Assist login</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
