// src/components/Analytics.jsx
import { useSubscriptionsMetrics } from '../api/hooks';

export default function Analytics() {
  const { metrics, loading, refresh } = useSubscriptionsMetrics();

  if (loading) return <p>Loading metrics...</p>;
  if (!metrics) return <p>No metrics.</p>;

  const { totals, revenue, period } = metrics;

  return (
    <div className="card">
      <h3>Current period</h3>
      <p>{period.monthStart} → {period.monthEnd}</p>
      <div className="grid">
        <div className="stat"><strong>Total users:</strong> {totals.totalUsers}</div>
        <div className="stat"><strong>Paid:</strong> {totals.paidUsers}</div>
        <div className="stat"><strong>Free:</strong> {totals.freeUsers}</div>
        <div className="stat"><strong>Banned:</strong> {totals.bannedUsers}</div>
        <div className="stat"><strong>Active:</strong> {totals.activeUsers}</div>
      </div>
      <div className="stat">
        <strong>Monthly revenue:</strong> {revenue.monthlyRevenue}
      </div>
      <button onClick={refresh}>Refresh</button>
    </div>
  );
}
