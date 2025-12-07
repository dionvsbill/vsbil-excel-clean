// src/components/LegalEditor.jsx
import { useEffect, useState } from 'react';
import { apiGet, apiPost } from '../api/client';

export default function LegalEditor() {
  const [terms, setTerms] = useState('');
  const [privacy, setPrivacy] = useState('');
  const [loading, setLoading] = useState(false);

  const load = async () => {
    setLoading(true);
    try {
      const t = await apiGet('/legal/terms');
      const p = await apiGet('/legal/privacy');
      setTerms(t.content || '');
      setPrivacy(p.content || '');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const saveTerms = async () => {
    try {
      await apiPost('/legal/terms', { content: terms });
      alert('Terms updated');
    } catch (e) { alert(e.message); }
  };

  const savePrivacy = async () => {
    try {
      await apiPost('/legal/privacy', { content: privacy });
      alert('Privacy updated');
    } catch (e) { alert(e.message); }
  };

  return (
    <div className="grid">
      <section>
        <h3>Terms & Conditions</h3>
        {loading ? <p>Loading...</p> : (
          <>
            <textarea value={terms} onChange={(e) => setTerms(e.target.value)} rows={10} />
            <button onClick={saveTerms}>Save</button>
          </>
        )}
      </section>
      <section>
        <h3>Privacy Policy</h3>
        {loading ? <p>Loading...</p> : (
          <>
            <textarea value={privacy} onChange={(e) => setPrivacy(e.target.value)} rows={10} />
            <button onClick={savePrivacy}>Save</button>
          </>
        )}
      </section>
    </div>
  );
}

