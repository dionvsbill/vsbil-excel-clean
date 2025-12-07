// src/components/Toast.jsx
import { createRoot } from 'react-dom/client';
import { useEffect, useState } from 'react';

let root;
function ensureRoot() {
  if (!root) {
    const container = document.createElement('div');
    container.id = 'toast-root';
    container.style.position = 'fixed';
    container.style.top = '10px';
    container.style.right = '10px';
    container.style.zIndex = 1000;
    document.body.appendChild(container);
    root = createRoot(container);
  }
  return root;
}

function ToastMessage({ msg, type, onDone }) {
  const [visible, setVisible] = useState(true);

  const bg =
    type === 'error' ? '#ef4444' :
    type === 'warn'  ? '#f59e0b' :
                       '#10b981';

  // auto remove after 3s with fade out
  useEffect(() => {
    const t = setTimeout(() => setVisible(false), 2500); // start fade
    const r = setTimeout(onDone, 3000); // remove after fade
    return () => {
      clearTimeout(t);
      clearTimeout(r);
    };
  }, [onDone]);

  return (
    <div
      onClick={() => {
        setVisible(false);
        setTimeout(onDone, 500); // allow fade out before removal
      }}
      style={{
        background: bg,
        color: 'white',
        padding: '10px 12px',
        margin: '5px',
        borderRadius: '8px',
        boxShadow: '0 2px 8px rgba(0,0,0,0.15)',
        fontSize: '14px',
        transition: 'all 0.5s ease',
        opacity: visible ? 1 : 0,
        transform: visible ? 'translateX(0)' : 'translateX(100%)',
        cursor: 'pointer',
        overflow: 'hidden'
      }}
    >
      {msg}
      {/* Progress bar */}
      <div
        style={{
          height: '3px',
          background: 'rgba(255,255,255,0.7)',
          marginTop: '6px',
          borderRadius: '2px',
          width: '100%',
          animation: 'toast-progress 3s linear forwards'
        }}
      />
      <style>
        {`
          @keyframes toast-progress {
            from { width: 100%; }
            to { width: 0%; }
          }
        `}
      </style>
    </div>
  );
}

// keep an array of toasts
let toasts = [];

function render() {
  ensureRoot().render(
    <div>
      {toasts.map((t, i) => (
        <ToastMessage
          key={i}
          msg={t.msg}
          type={t.type}
          onDone={() => {
            toasts = toasts.filter((_, j) => j !== i);
            render();
          }}
        />
      ))}
    </div>
  );
}

export default {
  success(msg) {
    toasts.push({ msg, type: 'success' });
    render();
  },
  error(msg) {
    toasts.push({ msg, type: 'error' });
    render();
  },
  warn(msg) {
    toasts.push({ msg, type: 'warn' });
    render();
  }
};
