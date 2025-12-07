import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { supabase } from '../supabaseClient';
import Toast from '../components/Toast';
import Modal from '../components/Modal';

export default function Login({ show, onClose, onShowRegister }) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [busy, setBusy] = useState(false);
  const navigate = useNavigate();

  async function handleLogin() {
    if (!email || !password) {
      Toast.warn('Email and password required');
      return;
    }
    setBusy(true);
    const { data, error } = await supabase.auth.signInWithPassword({
      email: email.trim(),
      password
    });
    setBusy(false);

    if (error) {
      const msg = error.message || 'Login failed';
      if (msg.toLowerCase().includes('email not confirmed')) {
        Toast.error('Email not confirmed. Check your inbox.');
      } else {
        Toast.error(msg);
      }
      return;
    }

    if (data?.session) {
      Toast.success('Logged in successfully');
      onClose();
      location.reload()
      // Preload sheets
      const res = await fetch('/excel/sheets');
      const { sheets } = await res.json();
      setSheets(sheets);

      navigate('/');
    } else {
      Toast.warn('Login succeeded but no session. Check email confirmation.');
    }
  }

  return (
    <Modal
      show={show}
      title="Login"
      onClose={onClose}
      footer={
        <div style={{ display: 'flex', alignItems: 'center', width: '100%' }}>
          <button onClick={handleLogin} disabled={busy}>
            {busy ? 'Signing in...' : 'Login'}
          </button>
          <button onClick={onClose} type="button" style={{ background: '#6b7280', marginLeft: '8px' }}>
            Cancel
          </button>
          <span style={{ marginLeft: 'auto', fontSize: '0.85em', color: '#9ca3af' }}>
            Developer contact: 233591857827
          </span>
        </div>
      }
    >
      <label>Email</label>
      <input
        type="email"
        autoFocus
        value={email}
        onChange={e => setEmail(e.target.value)}
        placeholder="Email"
      />
      <label>Password</label>
      <input
        type="password"
        value={password}
        onChange={e => setPassword(e.target.value)}
        placeholder="Password"
      />

      <p style={{ marginTop: '12px', fontSize: '0.9em', color: '#9ca3af' }}>
        Don’t have an account?{' '}
        <button
          type="button"
          onClick={() => {
            onClose();
            if (onShowRegister) onShowRegister();
            else navigate('/register');
          }}
          style={{ background: 'none', border: 'none', color: '#3b82f6', cursor: 'pointer', padding: 0 }}
        >
          Register
        </button>
      </p>
    </Modal>
  );
}
