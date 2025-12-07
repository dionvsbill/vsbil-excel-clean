// src/api/hooks.js
import { useEffect, useState, useCallback } from 'react';
import { apiGet, apiPost } from './client';

export function useSheets() {
  const [sheets, setSheets] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  const refresh = useCallback(async () => {
    try {
      setLoading(true);
      setError(null);
      const data = await apiGet('/excel/sheets');
      setSheets(data.sheets || []);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    refresh();
  }, [refresh]);

  return { sheets, loading, error, refresh };
}

export function useAudit(limit = 100) {
  const [entries, setEntries] = useState([]);
  const [loading, setLoading] = useState(false);

  const fetchAudit = useCallback(async () => {
    setLoading(true);
    try {
      const res = await apiGet(`/audit/list?limit=${limit}`);
      setEntries(res.entries || []);
    } finally {
      setLoading(false);
    }
  }, [limit]);

  useEffect(() => {
    fetchAudit();
  }, [fetchAudit]);

  return { entries, loading, refresh: fetchAudit };
}

export function useSubscriptionsMetrics() {
  const [metrics, setMetrics] = useState(null);
  const [loading, setLoading] = useState(false);

  const refresh = useCallback(async () => {
    setLoading(true);
    try {
      const res = await apiGet('/admin/subscriptions/metrics');
      setMetrics(res);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    refresh();
  }, [refresh]);

  return { metrics, loading, refresh };
}

export function usePricing() {
  const [pricing, setPricing] = useState(null);
  const [loading, setLoading] = useState(false);

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const res = await apiGet('/admin/subscriptions/pricing');
      setPricing(res.pricing);
    } finally {
      setLoading(false);
    }
  }, []);

  const update = useCallback(async (payload) => {
    await apiPost('/admin/subscriptions/pricing', payload);
    await load();
  }, [load]);

  useEffect(() => {
    load();
  }, [load]);

  return { pricing, loading, update };
}
