// src/api/client.js
export const API_BASE = import.meta.env.VITE_API_BASE || 'http://localhost:3000';

const getToken = () => localStorage.getItem('token') || '';

export async function apiGet(path, opts = {}) {
  const res = await fetch(`${API_BASE}${path}`, {
    ...opts,
    method: 'GET',
    headers: {
      ...(opts.headers || {}),
      Authorization: `Bearer ${getToken()}`,
    },
  });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

export async function apiPost(path, body, opts = {}) {
  const res = await fetch(`${API_BASE}${path}`, {
    ...opts,
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      ...(opts.headers || {}),
      Authorization: `Bearer ${getToken()}`,
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

export async function apiDownload(path, opts = {}) {
  const res = await fetch(`${API_BASE}${path}`, {
    ...opts,
    method: 'GET',
    headers: {
      ...(opts.headers || {}),
      Authorization: `Bearer ${getToken()}`,
    },
  });
  if (!res.ok) throw new Error(await res.text());
  return res.blob();
}

export function getAuthHeaders(extra = {}) {
  return {
    Authorization: `Bearer ${getToken()}`,
    ...extra,
  };
}
