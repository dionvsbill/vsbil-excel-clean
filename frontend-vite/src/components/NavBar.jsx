import { useNavigate } from 'react-router-dom';
import { useState } from 'react';
import { supabase } from '../supabaseClient';
import Toast from './Toast';

export default function NavBar({ user, onLogin, onRegister }) {
  const [busy, setBusy] = useState(false);
  const navigate = useNavigate();

  async function handleLogout() {
    setBusy(true);
    const { error } = await supabase.auth.signOut();
    setBusy(false);

    if (error) Toast.error(error.message);
    else {
      Toast.success('Logged out successfully');
      navigate('/'); // back to dashboard
    }
  }

  return (
    <nav className="navbar">
      <div className="nav-left">
        <button className="nav-link" onClick={() => navigate('/')}>Dashboard</button>
        {!user && (
          <>
            <button className="nav-link" onClick={onLogin}>Login</button>
            <button className="nav-link" onClick={onRegister}>Register</button>
          </>
        )}
      </div>
      <div className="nav-right">
        {user ? (
          <>
            <span className="role-badge">{user.user_metadata?.role || 'user'}</span>
            <button onClick={handleLogout} disabled={busy}>
              {busy ? 'Logging out...' : 'Logout'}
            </button>
          </>
        ) : (
          <span className="role-badge">Guest</span>
        )}
      </div>
    </nav>
  );
}
