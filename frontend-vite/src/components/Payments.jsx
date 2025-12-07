// src/components/Payments.jsx
import { useState } from 'react';
import { apiPost } from '../api/client';

export default function Payments() {
  const [email, setEmail] = useState('');
  const [amount, setAmount] = useState('10000'); // kobo or pesewas depending on currency
  const [reference, setReference] = useState('');
  const [initData, setInitData] = useState(null);

  const init = async () => {
    try {
      const data = await apiPost('/payments/init', { email, amount: parseInt(amount, 10) });
      setInitData(data);
      alert('Initialized — redirect user to Paystack authorization URL');
    } catch (e) {
      alert(e.message);
    }
  };

  const verify = async () => {
    try {
      await apiPost('/payments/verify', { reference, email });
      alert('Payment verified and plan updated');
    } catch (e) {
      alert(e.message);
    }
  };

  return (
    <div className="card">
      <h2>Payments (Paystack)</h2>
      <div className="row">
        <input placeholder="Email" value={email} onChange={(e) => setEmail(e.target.value)} />
        <input placeholder="Amount (minor unit)" value={amount} onChange={(e) => setAmount(e.target.value)} />
        <button onClick={init}>Initialize</button>
      </div>
      {initData && <p>Ref: {initData?.data?.reference} — Auth URL: {initData?.data?.authorization_url}</p>}
      <div className="row">
        <input placeholder="Reference" value={reference} onChange={(e) => setReference(e.target.value)} />
        <button onClick={verify}>Verify</button>
      </div>
    </div>
  );
}
