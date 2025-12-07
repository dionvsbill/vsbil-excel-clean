import { useEffect, useState } from 'react';
import Toast from './components/Toast';

export default function InstallButton() {
  const [deferredPrompt, setDeferredPrompt] = useState(null);

  useEffect(() => {
    const handler = (e) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };
    window.addEventListener('beforeinstallprompt', handler);
    return () => window.removeEventListener('beforeinstallprompt', handler);
  }, []);

  const handleInstall = async () => {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    const { outcome } = await deferredPrompt.userChoice;
    Toast.success(`User response to install: ${outcome}`);
    setDeferredPrompt(null);
  };

  return (
    deferredPrompt && (
      <button onClick={handleInstall} aria-label="Install this app">
        Install App
      </button>
    )
  );
}
