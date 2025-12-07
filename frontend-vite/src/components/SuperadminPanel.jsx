// src/components/SuperadminPanel.jsx
import UsersTable from './UsersTable';
import Support from './Support';

export default function SuperadminPanel() {
  return (
    <div className="grid">
      <section>
        <h2>Users</h2>
        <UsersTable onAction={() => {}} />
      </section>

      <section>
        <h2>Support</h2>
        <Support />
      </section>
    </div>
  );
}
