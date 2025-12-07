// src/components/Audit.jsx
import { useAudit } from '../api/hooks';

export default function Audit() {
  const { entries, loading, refresh } = useAudit();

  return (
    <div className="card">
      <h2>Audit logs</h2>
      {loading && <p>Loading...</p>}
      <button onClick={refresh}>Refresh</button>
      <table className="table">
        <thead>
          <tr><th>Time</th><th>Actor</th><th>Action</th><th>Sheet</th><th>Details</th></tr>
        </thead>
        <tbody>
          {(entries || []).map((e, idx) => (
            <tr key={idx}>
              <td>{e.ts}</td>
              <td>{e.actor}</td>
              <td>{e.action}</td>
              <td>{e.sheet}</td>
              <td><pre>{JSON.stringify(e.details || e.metadata || {}, null, 2)}</pre></td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
