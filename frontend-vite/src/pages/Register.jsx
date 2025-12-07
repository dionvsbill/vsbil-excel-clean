import { useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { supabase } from '../supabaseClient';
import Toast from '../components/Toast';
import Modal from '../components/Modal';

export default function Register({ show, onClose, onShowLogin }) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [busy, setBusy] = useState(false);
  const navigate = useNavigate();

  async function handleRegister() {
    if (!email || !password) {
      Toast.warn('Email and password required');
      return;
    }
    if (password.length < 8) {
      Toast.warn('Use at least 8 characters');
      return;
    }
    setBusy(true);
    const { data, error } = await supabase.auth.signUp({
      email: email.trim(),
      password
    });
    setBusy(false);

    if (error) {
      const msg = error.message || 'Registration failed';
      Toast.error(msg);
      return;
    }

    if (data?.user) {
      Toast.success('Check your email to confirm registration');
      onClose();
    } else {
      Toast.warn('Registration submitted. Awaiting email confirmation.');
    }
  }

  return (
    <Modal
      show={show}
      title="Register"
      onClose={onClose}
      footer={
        <div style={{ display: 'flex', alignItems: 'center', width: '100%' }}>
          <button onClick={handleRegister} disabled={busy}>
            {busy ? 'Registering...' : 'Register'}
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
        Already have an account?{' '}
        <button
          type="button"
          onClick={() => {
            onClose();
            if (onShowLogin) onShowLogin();
            else navigate('/login');
          }}
          style={{ background: 'none', border: 'none', color: '#3b82f6', cursor: 'pointer', padding: 0 }}
        >
          Login
        </button>
      </p>
    </Modal>
  );
}
