// src/components/PricingControls.jsx
import { useState, useEffect } from 'react';
import { usePricing } from '../api/hooks';

export default function PricingControls() {
  const { pricing, loading, update } = usePricing();
  const [monthly, setMonthly] = useState('');
  const [yearly, setYearly] = useState('');
  const [currency, setCurrency] = useState('GHS');

  useEffect(() => {
    if (pricing) {
      setMonthly(String(pricing.monthly_amount));
      setYearly(String(pricing.yearly_amount));
      setCurrency(pricing.currency || 'GHS');
    }
  }, [pricing]);

  const save = async () => {
    await update({
      monthly_amount: parseInt(monthly, 10),
      yearly_amount: parseInt(yearly, 10),
      currency,
    });
    alert('Pricing updated');
  };

  return (
    <div className="card">
      {loading && <p>Loading...</p>}
      <div className="row">
        <label>Monthly amount</label>
        <input value={monthly} onChange={(e) => setMonthly(e.target.value)} />
      </div>
      <div className="row">
        <label>Yearly amount</label>
        <input value={yearly} onChange={(e) => setYearly(e.target.value)} />
      </div>
      <div className="row">
        <label>Currency</label>
        <input value={currency} onChange={(e) => setCurrency(e.target.value)} />
      </div>
      <button onClick={save} disabled={loading}>Save</button>
    </div>
  );
}
