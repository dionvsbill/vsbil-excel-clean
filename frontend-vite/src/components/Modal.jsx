// src/components/Modal.jsx
import { useEffect, useState } from 'react';

export default function Modal({ show, title, children, onClose, footer }) {
  const [closing, setClosing] = useState(false);
  const [theme, setTheme] = useState('system'); // system | dark | light

  // ✅ Load saved theme preference
  useEffect(() => {
    const saved = localStorage.getItem('app-theme');
    if (saved) setTheme(saved);
  }, []);

  // ✅ Save theme preference
  useEffect(() => {
    localStorage.setItem('app-theme', theme);
  }, [theme]);

  // ✅ Close on Escape key
  useEffect(() => {
    function handleKey(e) {
      if (e.key === 'Escape') triggerClose();
    }
    window.addEventListener('keydown', handleKey);
    return () => window.removeEventListener('keydown', handleKey);
  }, []);

  function triggerClose() {
    setClosing(true);
    setTimeout(() => {
      setClosing(false);
      if (onClose) onClose();
    }, 400); // match animation duration
  }

  if (!show && !closing) return null;

  // ✅ Theme logic
  const isDark =
    theme === 'dark' ||
    (theme === 'system' &&
      window.matchMedia('(prefers-color-scheme: dark)').matches);

  const modalBg = isDark ? '#121826' : '#ffffff';
  const modalText = isDark ? '#e6eefc' : '#111827';

  return (
    <div
      style={{
        position: 'fixed',
        inset: 0,
        background: 'rgba(0,0,0,0.5)',
        backdropFilter: 'blur(6px)', // ✅ frosted blur
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 1000,
        animation: closing ? 'fadeOut 0.4s ease forwards' : 'fadeIn 0.4s ease',
      }}
      onClick={triggerClose}
    >
      <div
        style={{
          background: modalBg,
          color: modalText,
          borderRadius: '12px',
          padding: '20px',
          width: '90%',
          maxWidth: '420px',
          boxShadow:
            '0 10px 25px rgba(0,0,0,0.4), 0 0 20px rgba(59,130,246,0.4)', // ✅ glow
          position: 'relative',
          display: 'flex',
          flexDirection: 'column',
          animation: closing
            ? 'slideOutScale 0.4s cubic-bezier(0.6, -0.28, 0.74, 0.05) forwards'
            : 'slideInScale 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275) forwards, pulseGlow 0.8s ease-out 0.2s',
        }}
        onClick={e => e.stopPropagation()}
      >
        {/* Close button */}
        <button
          onClick={triggerClose}
          aria-label="Close modal"
          style={{
            position: 'absolute',
            top: '10px',
            right: '10px',
            background: 'transparent',
            border: 'none',
            color: modalText,
            fontSize: '1.2em',
            cursor: 'pointer',
          }}
        >
          ✕
        </button>

        {/* Title */}
        {title && (
          <h2 style={{ marginTop: 0, marginBottom: '16px' }}>{title}</h2>
        )}

        {/* Body */}
        <div style={{ flex: 1 }}>{children}</div>

        {/* Footer with theme toggle */}
        <div
          style={{
            marginTop: '20px',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            gap: '10px',
          }}
        >
          <div style={{ fontSize: '0.85em' }}>
            Theme:{' '}
            <select
              value={theme}
              onChange={e => setTheme(e.target.value)}
              style={{
                background: isDark ? '#1f2937' : '#f3f4f6',
                color: modalText,
                border: '1px solid #9ca3af',
                borderRadius: '6px',
                padding: '4px 8px',
              }}
            >
              <option value="system">System</option>
              <option value="dark">Dark</option>
              <option value="light">Light</option>
            </select>
          </div>

          {footer && (
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px' }}>
              {footer}
            </div>
          )}
        </div>
      </div>

      {/* ✅ Keyframes */}
      <style>
        {`
          @keyframes fadeIn {
            from { opacity: 0; }
            to { opacity: 1; }
          }
          @keyframes fadeOut {
            from { opacity: 1; }
            to { opacity: 0; }
          }
          @keyframes slideInScale {
            from { opacity: 0; transform: translateY(-40px) scale(0.95); }
            to { opacity: 1; transform: translateY(0) scale(1); }
          }
          @keyframes slideOutScale {
            from { opacity: 1; transform: translateY(0) scale(1); }
            to { opacity: 0; transform: translateY(-40px) scale(0.95); }
          }
          @keyframes pulseGlow {
            0%   { box-shadow: 0 10px 25px rgba(0,0,0,0.4), 0 0 0 rgba(59,130,246,0.0); }
            50%  { box-shadow: 0 10px 25px rgba(0,0,0,0.4), 0 0 30px rgba(59,130,246,0.6); }
            100% { box-shadow: 0 10px 25px rgba(0,0,0,0.4), 0 0 20px rgba(59,130,246,0.4); }
          }
        `}
      </style>
    </div>
  );
}
