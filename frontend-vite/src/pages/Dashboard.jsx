// src/pages/Dashboard.jsx
import React, { useEffect, useMemo, useRef, useState } from 'react';
import { supabase } from '../supabaseClient';
import Toast from '../components/Toast';
import LoadingOverlay from '../components/LoadingOverlay';
import Login from '../pages/Login';
import Register from '../pages/Register';
import InstallButton from '../InstallButton';
import PayButton from '../components/PayButton';
import { io } from 'socket.io-client';
import {
  FaFileExcel,
  FaBars,
  FaTrash,
  FaEdit,
  FaSearch,
  FaPlus,
  FaSave,
  FaTimes,
  FaFileExport,
  FaDownload,
} from 'react-icons/fa';

import Handsontable from 'handsontable';
import 'handsontable/dist/handsontable.full.css';
import { HyperFormula } from 'hyperformula';
import './dashboard.css';

/* =======================================================
   Utilities
======================================================= */
function dedupeNames(list) {
  const seen = new Set();
  const out = [];
  for (const n of list) {
    if (!seen.has(n)) {
      seen.add(n);
      out.push(n);
    }
  }
  return out;
}

function formatTS(ts) {
  try {
    const d = new Date(ts);
    return d.toLocaleString();
  } catch {
    return '—';
  }
}

// Stable sort helper (keeps insertion order, pins latest at top without shuffle)
function pinLatestStable(allSheets, latest) {
  if (!Array.isArray(allSheets) || allSheets.length === 0) return [];
  if (!latest || !allSheets.includes(latest)) return allSheets.slice();
  const rest = allSheets.filter((s) => s !== latest);
  return [latest, ...rest];
}

/* =======================================================
   Dashboard component
======================================================= */
export default function Dashboard() {
  const [user, setUser] = useState(null);
  const [role, setRole] = useState('guest');
  const [plan, setPlan] = useState('free');
  const [canEdit, setCanEdit] = useState(false);
  const [appName, setAppName] = useState('vsbil Excel');

  const [sheets, setSheets] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState(null);
  const [stagedGrid, setStagedGrid] = useState([]);
  const [loading, setLoading] = useState(false);
  const containerRef = useRef(null);
  const hotRef = useRef(null);
  const hfRef = useRef(null);

  const [showLogin, setShowLogin] = useState(false);
  const [showRegister, setShowRegister] = useState(false);

  const [actionsOpen, setActionsOpen] = useState(false);
  const [showAudit, setShowAudit] = useState(false);
  const [auditEntries, setAuditEntries] = useState([]);

  const [theme, setTheme] = useState('system');
  const [searchQuery, setSearchQuery] = useState('');

  const [socketConnected, setSocketConnected] = useState(false);
  const socketRef = useRef(null);

  const [isDragOver, setIsDragOver] = useState(false);

  const API_BASE = import.meta.env.VITE_API_URL || 'http://localhost:4000';

  const [isMobile, setIsMobile] = useState(window.innerWidth < 992);
  useEffect(() => {
    const onResize = () => setIsMobile(window.innerWidth < 992);
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, []);

  const [hoveredItem, setHoveredItem] = useState(null);
  const [swipeX, setSwipeX] = useState({});
  const touchStartXRef = useRef({});

  // Read-only override used when opening a sheet from Audit
  const [forceReadOnly, setForceReadOnly] = useState(false);

  // Ads disabled across app; UI hint remains only as a message
  const shouldShowAds = false;

  // Persist latest edited sheet locally to pin at top even after refresh
  const [lastEditedSheet, setLastEditedSheet] = useState(
    localStorage.getItem('latestSheet') || null
  );

  // Cache original stable order (first load), to keep list stable
  const [originalOrder, setOriginalOrder] = useState([]);

  // Unified response parsing to avoid <!DOCTYPE... blowing up JSON.parse
  async function parseResponse(res) {
    const contentType = res.headers.get('content-type') || '';
    if (contentType.includes('application/json')) {
      try {
        return await res.json();
      } catch (error) {
        return { error: 'Invalid JSON response from server' };
      }
    }

    const text = await res.text();
    if (/<!DOCTYPE|<html/i.test(text)) {
      return { error: 'Unexpected HTML response from server', raw: text, status: res.status };
    }

    try {
      return JSON.parse(text);
    } catch (error) {
      return { error: 'Unexpected response from server', raw: text };
    }
  }

  /* -------------------------------------------------------
     Session sync
  ------------------------------------------------------- */
  useEffect(() => {
    supabase.auth.getUser().then(({ data }) => {
      if (data.user) updateUser(data.user);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      if (session?.user) updateUser(session.user);
      else {
        clearBrowserState();
        setUser(null);
        setRole('guest');
        setCanEdit(false);
        setPlan('free');
        setAppName('vsbil Excel');
      }
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  async function updateUser(u) {
    setUser(u);
    try {
      const { data, error } = await supabase
        .from('profiles')
        .select('role, can_edit, plan, app_name')
        .eq('id', u.id)
        .single();
      if (error) {
        Toast.error(error.message);
        setRole('user');
        setCanEdit(false);
        setPlan('free');
        return;
      }
      const nextRole = data?.role || 'user';
      const nextPlan = data?.plan || 'free';

      // Premium updates: ensure UI flips immediately from free -> paid
      setRole(nextRole);
      setPlan(nextPlan);
      setCanEdit(Boolean(data?.can_edit) || nextRole === 'admin' || nextRole === 'superadmin' || nextPlan === 'paid');
      if (data?.app_name) setAppName(data.app_name);
    } catch (err) {
      Toast.error(err.message);
      setRole('user');
      setCanEdit(false);
      setPlan('free');
    }
  }

  function clearBrowserState() {
    try {
      localStorage.clear();
      sessionStorage.clear();
      if ('caches' in window) {
        caches.keys().then((names) => names.forEach((n) => caches.delete(n)));
      }
    } catch {}
  }

  async function authHeader() {
    const { data: { session } } = await supabase.auth.getSession();
    return session?.access_token ? { Authorization: `Bearer ${session.access_token}` } : {};
  }

  async function apiGet(path) {
    setLoading(true);
    try {
      const headers = await authHeader();
      const res = await fetch(`${API_BASE}${path}`, { headers });
      const j = await parseResponse(res);
      if (j.error) Toast.error(j.error);
      return j;
    } catch (e) {
      Toast.error(e.message);
      return { error: e.message };
    } finally {
      setLoading(false);
    }
  }

  async function apiPost(path, body) {
    setLoading(true);
    try {
      const headers = { 'Content-Type': 'application/json', ...(await authHeader()) };
      const res = await fetch(`${API_BASE}${path}`, { method: 'POST', headers, body: JSON.stringify(body) });
      const j = await parseResponse(res);
      if (j.error) Toast.error(j.error);
      return j;
    } catch (e) {
      Toast.error(e.message);
      return { error: e.message };
    } finally {
      setLoading(false);
    }
  }

  // Fetch latest edited/saved sheet from audit and pin it
  async function loadLatestSheet() {
    const j = await apiGet('/excel/latest');
    if (!j.error && j.sheet) {
      setLastEditedSheet(j.sheet);
      localStorage.setItem('latestSheet', j.sheet);
    }
  }

  // Sheets list (initial load)
  useEffect(() => {
    (async () => {
      const j = await apiGet('/excel/sheets');
      if (!j.error) {
        const names = dedupeNames(j.sheets || []);
        setSheets(names);
        if (originalOrder.length === 0) {
          setOriginalOrder(names.slice());
        }
        await loadLatestSheet();
      }
    })();
  }, []);

  // Audit entries list when opened
  useEffect(() => {
    (async () => {
      if (!showAudit) return;
      const j = await apiGet('/audit/list');
      if (!j.error) setAuditEntries(Array.isArray(j.entries) ? j.entries : []);
    })();
  }, [showAudit]);

  // Theme
  useEffect(() => {
    const saved = localStorage.getItem('app-theme');
    if (saved) setTheme(saved);
  }, []);
  useEffect(() => {
    localStorage.setItem('app-theme', theme);
    document.documentElement.setAttribute('data-theme', theme);
  }, [theme]);

  /* -------------------------------------------------------
     Real-time
  ------------------------------------------------------- */
  useEffect(() => {
    if (!selectedSheet) return;
    const url = import.meta.env.VITE_SOCKET_URL || API_BASE;
    const socket = io(url, { transports: ['websocket'] });
    socketRef.current = socket;

    socket.on('connect', () => {
      setSocketConnected(true);
      socket.emit('join', { room: `sheet:${selectedSheet}` });
    });

    socket.on('cell-edit', (op) => {
      setStagedGrid((prev) => {
        if (!prev?.length) return prev;
        const next = prev.map((row) => row.slice());
        const { rowIndex, colIndex, value } = op;
        if (next[rowIndex] && colIndex >= 0) next[rowIndex][colIndex] = value;
        return next;
      });
    });

    socket.on('disconnect', () => setSocketConnected(false));

    return () => {
      socket.emit('leave', { room: `sheet:${selectedSheet}` });
      socket.disconnect();
      setSocketConnected(false);
    };
  }, [selectedSheet, API_BASE]);

  /* -------------------------------------------------------
     Stable list + pin latest (no shuffle)
  ------------------------------------------------------- */
  const filteredSheets = useMemo(() => {
    const base = !searchQuery.trim()
      ? sheets
      : sheets.filter((s) => s.toLowerCase().includes(searchQuery.toLowerCase()));

    // Stabilize: order by originalOrder index (keeps initial ordering)
    const stabilized = base.slice().sort((a, b) => {
      const ia = originalOrder.indexOf(a);
      const ib = originalOrder.indexOf(b);
      return ia - ib;
    });

    // Pin latest (top) but preserve stable order for rest
    return pinLatestStable(stabilized, lastEditedSheet);
  }, [sheets, searchQuery, lastEditedSheet, originalOrder]);

  async function previewSheet(name, opts = { readOnly: false }) {
    const j = await apiGet(`/excel/preview?sheet=${encodeURIComponent(name)}`);
    if (!j.error) {
      setSelectedSheet(name);
      setStagedGrid(j.preview || []);
      setForceReadOnly(Boolean(opts.readOnly));
    } else {
      Toast.error(j.error || 'Failed to open sheet');
    }
  }

  async function saveAllChanges() {
    if (!selectedSheet) {
      Toast.error('No sheet selected');
      return;
    }

    const safeGridSource = hotRef?.current ? hotRef.current.getData() : stagedGrid;
    if (!Array.isArray(safeGridSource) || safeGridSource.length === 0) {
      Toast.warn('Nothing to save');
      return;
    }

    try {
      const safeGrid = safeGridSource.map((row) => Array.isArray(row) ? row : [row]);
      const response = await apiPost('/excel/save-all', { sheet: selectedSheet, data: safeGrid });

      if (response.error) {
        Toast.error(response.error);
        return;
      }

      Toast.success(`Saved changes to ${selectedSheet}`);
      setLastEditedSheet(selectedSheet);
      localStorage.setItem('latestSheet', selectedSheet);

      setStagedGrid(safeGrid);

      const previewResponse = await apiGet(`/excel/preview?sheet=${encodeURIComponent(selectedSheet)}`);
      if (previewResponse.error) {
        Toast.error(previewResponse.error);
        return;
      }

      setStagedGrid(previewResponse.preview || safeGrid);

      // Close the editor after successful save (UX choice)
      setForceReadOnly(false);
      setSelectedSheet(null);
    } catch (error) {
      Toast.error('Failed to save changes');
    }
  }

  async function deleteSheet(name) {
    // Allow delete for admin, superadmin, or paid users (premium update)
    const isAdminLike = role === 'admin' || role === 'superadmin';
    const isPremium = plan === 'paid';
    if (!isAdminLike && !isPremium) {
      Toast.warn('Only admin or premium users can delete sheets');
      return;
    }
    const confirmDelete = window.confirm(`Delete sheet "${name}"?`);
    if (!confirmDelete) return;
    const res = await apiPost('/excel/delete-sheet', { name });
    if (!res.error) {
      Toast.success(`Deleted "${name}"`);
      setSheets((prev) => prev.filter((s) => s !== name));
      if (selectedSheet === name) {
        setSelectedSheet(null);
        setStagedGrid([]);
        setForceReadOnly(false);
      }
      // Update original order cache if needed
      setOriginalOrder((prev) => prev.filter((s) => s !== name));
    }
  }

  function uniqueName(base, existingList) {
    if (!existingList.includes(base)) return base;
    let i = 1;
    while (existingList.includes(`${base} ${i}`)) i++;
    return `${base} ${i}`;
  }

  async function addSheet() {
    const nameInput = (prompt('Enter new sheet name:') || '').trim();
    if (!nameInput) return;
    const name = uniqueName(nameInput, sheets);
    let overwrite = false;

    if (sheets.includes(nameInput)) {
      overwrite = window.confirm(`"${nameInput}" exists. Overwrite it? Press Cancel to save as "${name}".`);
    }

    const finalTarget = overwrite ? nameInput : name;
    const j = await apiPost('/excel/add-sheet', { name: finalTarget, overwrite });
    if (!j.error) {
      const finalName = finalTarget;
      Toast.success(`Sheet ${finalName} added`);
      const updated = j.sheets ? dedupeNames(j.sheets) : dedupeNames([...sheets, finalName]);
      setSheets(updated);

      // refresh original order to keep stabilized ordering in sync
      setOriginalOrder((prev) => {
        const next = prev.slice();
        if (!next.includes(finalName)) next.push(finalName);
        return next;
      });

      setSelectedSheet(finalName);
      await previewSheet(finalName);
      setLastEditedSheet(finalName);
      localStorage.setItem('latestSheet', finalName);
    }
  }

  async function handleUpload(file) {
    try {
      const formData = new FormData();
      formData.append('file', file);
      const headers = await authHeader();
      const res = await fetch(`${API_BASE}/excel/upload`, {
        method: 'POST',
        headers,
        body: formData,
      });
      const j = await parseResponse(res);
      if (j.error) {
        Toast.error(j.error);
        return;
      }

      let incomingNames = [];
      if (Array.isArray(j.sheetNames) && j.sheetNames.length > 0) {
        incomingNames = j.sheetNames;
      } else if (j.fileName) {
        incomingNames = [j.fileName.replace(/\.[^/.]+$/, '')];
      }

      const deduped = [...sheets];
      const finalNames = [];
      for (const nm of incomingNames) {
        if (deduped.includes(nm)) {
          const overwrite = window.confirm(`"${nm}" exists. Overwrite? Press Cancel to save as new.`);
          const unique = overwrite ? nm : uniqueName(nm, deduped);
          finalNames.push(unique);
          if (!deduped.includes(unique)) deduped.push(unique);
          if (overwrite) {
            await apiPost('/excel/overwrite', { name: nm });
          }
        } else {
          finalNames.push(nm);
          deduped.push(nm);
        }
      }

      setSheets(dedupeNames(deduped));
      setOriginalOrder((prev) => {
        const next = prev.slice();
        for (const nm of finalNames) {
          if (!next.includes(nm)) next.push(nm);
        }
        return next;
      });

      const first = finalNames[0];
      if (first) {
        setSelectedSheet(first);
        setStagedGrid(j.preview || []);
        setLastEditedSheet(first);
        localStorage.setItem('latestSheet', first);
      }

      Toast.success(`Uploaded ${j.fileName || finalNames.join(', ')}`);
    } catch (err) {
      Toast.error(err.message || 'Upload failed');
    }
  }

  const addRowAbove = () => {
    const hot = window.__hotInstance;
    if (!canEdit) return;
    if (!hot) {
      setStagedGrid((prev) => (prev && prev.length ? prev : [[]]));
      return;
    }
    const sel = hot.getSelectedLast();
    const rowIndex = sel ? sel[0] : 0;
    hot.alter('insert_row', rowIndex);
  };

  const addRowBelow = () => {
    const hot = window.__hotInstance;
    if (!canEdit) return;
    if (!hot) {
      setStagedGrid((prev) => (prev && prev.length ? prev : [[]]));
      return;
    }
    const sel = hot.getSelectedLast();
    const rowIndex = sel ? sel[0] : hot.countRows() - 1;
    hot.alter('insert_row', rowIndex + 1);
  };

  const addColLeft = () => {
    const hot = window.__hotInstance;
    if (!canEdit) return;
    if (!hot) {
      setStagedGrid((prev) => (prev && prev.length ? prev : [[]]));
      return;
    }
    const sel = hot.getSelectedLast();
    const colIndex = sel ? sel[1] : 0;
    hot.alter('insert_col', colIndex);
  };

  const addColRight = () => {
    const hot = window.__hotInstance;
    if (!canEdit) return;
    if (!hot) {
      setStagedGrid((prev) => (prev && prev.length ? prev : [[]]));
      return;
    }
    const sel = hot.getSelectedLast();
    const colIndex = sel ? sel[1] : hot.countCols() - 1;
    hot.alter('insert_col', colIndex + 1);
  };

  /* -------------------------------------------------------
     Mobile swipe handlers
  ------------------------------------------------------- */
  const onTouchStart = (sheet, e) => {
    touchStartXRef.current[sheet] = e.touches[0].clientX;
    setSwipeX((prev) => ({ ...prev, [sheet]: 0 }));
  };
  const onTouchMove = (sheet, e) => {
    const startX = touchStartXRef.current[sheet];
    if (startX == null) return;
    const delta = e.touches[0].clientX - startX;
    setSwipeX((prev) => ({ ...prev, [sheet]: Math.min(0, delta) }));
  };
  const onTouchEnd = (sheet) => {
    const delta = swipeX[sheet] || 0;
    if (delta < -40) {
      setSwipeX((prev) => ({ ...prev, [sheet]: -72 }));
    } else {
      setSwipeX((prev) => ({ ...prev, [sheet]: 0 }));
    }
    touchStartXRef.current[sheet] = null;
  };

  /* -------------------------------------------------------
     Export & Download
  ------------------------------------------------------- */
  async function exportSheet() {
    if (!selectedSheet) {
      Toast.error('No sheet selected');
      return;
    }
    try {
      setLoading(true);
      const headers = await authHeader();
      const res = await fetch(`${API_BASE}/excel/export/csv?sheet=${encodeURIComponent(selectedSheet)}`, {
        method: 'GET',
        headers,
      });
      if (!res.ok) {
        const txt = await res.text();
        Toast.error(txt || 'Export failed');
        return;
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const cd = res.headers.get('content-disposition') || '';
      const m = cd.match(/filename="?([^"]+)"?/i);
      const filename = m ? m[1] : `${selectedSheet}.csv`;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      Toast.success(`Exported ${selectedSheet}`);
    } catch (err) {
      Toast.error(err.message);
    } finally {
      setLoading(false);
    }
  }

  async function downloadSheet() {
    if (!selectedSheet) {
      Toast.error('No sheet selected');
      return;
    }
    try {
      setLoading(true);
      const headers = await authHeader();
      // Backend route serves the workbook (xlsx). It ignores ?sheet; we download the full workbook.
      const res = await fetch(`${API_BASE}/excel/download`, {
        method: 'GET',
        headers,
      });
      if (!res.ok) {
        const txt = await res.text();
        Toast.error(txt || 'Download failed');
        return;
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      // Attempt to derive filename from content-disposition; fallback
      const cd = res.headers.get('content-disposition') || '';
      const m = cd.match(/filename="?([^"]+)"?/i);
      const filename = m ? m[1] : `${selectedSheet}.xlsx`;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      Toast.success(`Downloaded ${filename}`);
    } catch (err) {
      Toast.error(err.message);
    } finally {
      setLoading(false);
    }
  }

  /* -------------------------------------------------------
     Upgrade success -> refresh profile and pin latest
  ------------------------------------------------------- */
  async function onUpgradeSuccess() {
    Toast.success('Upgrade successful');
    if (user) {
      await updateUser(user);
      await loadLatestSheet();
    } else {
      // Try to re-fetch session and profile if user is missing
      const { data: { session } } = await supabase.auth.getSession();
      if (session?.user) {
        await updateUser(session.user);
        await loadLatestSheet();
      }
    }
  }

  /* -------------------------------------------------------
     Render
  ------------------------------------------------------- */
  return (
    <div
      className={`dashboard-container ${isMobile ? 'mobile' : 'desktop'}`}
      style={{
        display: 'grid',
        gridTemplateRows: 'auto 1fr auto',
        minHeight: '100vh',
        overflow: 'hidden',
      }}
    >
      {/* Header */}
      <header
        className="navbar"
        style={{
          gridColumn: '1 / -1',
          display: 'flex',
          alignItems: 'center',
          gap: 12,
          padding: '10px 16px',
          borderBottom: '1px solid #e5e7eb',
          position: 'sticky',
          top: 0,
          background: 'var(--bg, --text)',
          zIndex: 50,
        }}
      >
        <div style={{ position: 'relative', display: 'flex', alignItems: 'center', width: '100%' }}>
          <button
            className="menu-toggle"
            onClick={() => setActionsOpen((v) => !v)}
            aria-label="Open actions"
            style={{
              background: 'none',
              color: 'var(--accent-hover)',
              border: 'none',
              fontSize: 18,
              cursor: 'pointer',
              marginRight: 8,
            }}
          >
            <FaBars />
          </button>

          <div className="brand" style={{ fontWeight: 600 }}>
            {appName}{' '}
          </div>

          <div
            className="searchbar"
            style={{ marginLeft: 16, display: 'flex', alignItems: 'center', gap: 8 }}
          >
            <FaSearch />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Search sheets"
              style={{ border: '1px solid #e5e7eb', borderRadius: 8, padding: '6px 8px', width: '100%' }}
            />
          </div>

          <div
            className="nav-actions"
            style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 8 }}
          >
            <InstallButton />
            {user && plan === 'free' && (
              <PayButton
                user={user}
                amount={500000}
                onSuccess={onUpgradeSuccess}
                onFailure={(e) => Toast.error(e?.message || 'Upgrade failed')}
              />
            )}
          </div>

          {/* Actions dropdown */}
          {actionsOpen && (
            <div
              className="actions-dropdown"
              style={{
                position: 'absolute',
                top: '100%',
                left: 0,
                marginTop: 8,
                background: 'var(--bg, --text)',
                border: '1px solid #e5e7eb',
                borderRadius: 8,
                boxShadow: '0 8px 20px rgba(0,0,0,0.12)',
                padding: 12,
                width: 280,
                zIndex: 99999,
              }}
            >
              <div
                className="dropdown-header"
                style={{
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                  marginBottom: 8,
                }}
              >
                <h3 style={{ margin: 0, fontSize: 16 }}>Actions</h3>
                <span
                  className="role-badge"
                  style={{
                    marginLeft: 8,
                    fontSize: 12,
                    padding: '2px 6px',
                    border: '1px solid #ddd',
                    borderRadius: 999,
                  }}
                >
                  {role}/{plan}
                </span>
                <button
                  className="secondary"
                  onClick={() => setActionsOpen(false)}
                  style={{ background: 'var(--danger, #fff)' }}
                >
                  <FaTimes /> Close
                </button>
              </div>
              <div className="dropdown-body" style={{ display: 'grid', gap: 8 }}>
                {!user && (
                  <button
                    className="primary"
                    onClick={() => {
                      setShowLogin(true);
                      setActionsOpen(false);
                    }}
                  >
                    Login
                  </button>
                )}
                {!user && (
                  <button
                    className="secondary"
                    onClick={() => {
                      setShowRegister(true);
                      setActionsOpen(false);
                    }}
                  >
                    Register
                  </button>
                )}

                {user && (
                  <button
                    className="secondary"
                    onClick={() => {
                      setShowAudit(true);
                      setActionsOpen(false);
                    }}
                  >
                    View Audit Log
                  </button>
                )}

                {user && (
                  <button
                    className="secondary small"
                    onClick={async () => {
                      await supabase.auth.signOut();
                      clearBrowserState();
                      window.location.reload();
                    }}
                    style={{ background: 'var(--danger, #fff)' }}
                  >
                    Logout
                  </button>
                )}

                <label className="drawer-field" style={{ display: 'grid', gap: 6 }}>
                  <span>Theme</span>
                  <select
                    value={theme}
                    onChange={(e) => setTheme(e.target.value)}
                    aria-label="Theme"
                  >
                    <option value="system">System</option>
                    <option value="dark">Dark</option>
                    <option value="light">Light</option>
                  </select>
                </label>
              </div>
            </div>
          )}
        </div>
      </header>

      {/* Content grid */}
      <div
        className="content-grid"
        style={{
          display: 'grid',
          gridTemplateColumns: isMobile ? '1fr' : '320px 1fr',
          gridTemplateRows: '1fr',
          height: 'calc(100vh - 120px)',
          overflow: 'hidden',
        }}
      >
        {/* Left: Sheet list */}
        {(!isMobile || !selectedSheet) && (
          <aside
            className="sheet-list-area"
            style={{
              background: 'var(--bg, --text)',
              borderRight: isMobile ? 'none' : '1px solid #e5e7eb',
              overflowY: 'auto',
              overflowX: 'hidden',
            }}
          >
            <div
              className="sheet-list-top"
              style={{
                position: 'sticky',
                top: 0,
                background: 'var(--bg, --text)',
                padding: 12,
                borderBottom: '1px solid #e5e7eb',
                zIndex: 2,
                display: 'flex',
              }}
            >
              <button className="primary small" onClick={addSheet}>
                <FaPlus /> New sheet
              </button>
              <div
                className="upload-area"
                onDragOver={(e) => {
                  e.preventDefault();
                  setIsDragOver(true);
                }}
                onDragLeave={(e) => {
                  e.preventDefault();
                  setIsDragOver(false);
                }}
                onDrop={async (e) => {
                  e.preventDefault();
                  setIsDragOver(false);
                  const file = e.dataTransfer.files[0];
                  if (!file) return;
                  await handleUpload(file);
                }}
                style={{
                  margin: 7,
                  padding: 7,
                  border: '2px dashed #e5e7eb',
                  borderRadius: 8,
                  textAlign: 'center',
                  color: '#6b7280',
                  background: '#fdfdfd',
                  transition: 'transform 0.2s ease, border-color 0.2s ease',
                  transform: isDragOver ? 'scale(1.05)' : 'scale(1)',
                  borderColor: isDragOver ? '#2563eb' : '#e5e7eb',
                }}
              >
                <p>Drag & drop Excel/CSV/PDF to add a sheet</p>
              </div>
            </div>

            <ul
              className="sheet-list"
              style={{
                listStyle: 'none',
                margin: 0,
                padding: 12,
                display: 'flex',
                flexDirection: 'column',
                gap: 8,
              }}
            >
              {filteredSheets.map((s) => {
                const isSwiped = (swipeX[s] || 0) !== 0;
                const translate = swipeX[s] || 0;
                return (
                  <li
                    key={s}
                    className={`sheet-item lis ${selectedSheet === s ? 'active' : ''}`}
                    title={s}
                    onClick={() => previewSheet(s)}
                    onMouseEnter={() => setHoveredItem(s)}
                    onMouseLeave={() => setHoveredItem(null)}
                    onTouchStart={(e) => onTouchStart(s, e)}
                    onTouchMove={(e) => onTouchMove(s, e)}
                    onTouchEnd={() => onTouchEnd(s)}
                    style={{
                      display: 'grid',
                      gridTemplateColumns: '1fr auto',
                      alignItems: 'center',
                      gap: 5,
                      padding: '10px 12px',
                      borderRadius: 10,
                      cursor: 'pointer',
                      background: 'var(--bg)',
                      position: 'relative',
                      transform: `translateX(${translate}px)`,
                      transition: isSwiped ? 'transform 0.08s ease' : 'transform 0.18s ease',
                    }}
                  >
                    <div
                      className="bubble"
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: 8,
                        minWidth: 0,
                      }}
                    >
                      <FaFileExcel className="sheet-icon" style={{ color: '#2563eb' }} />
                      <span
                        className="sheet-name"
                        style={{
                          fontWeight: 500,
                          color: 'var(--text)',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                        }}
                      >
                        {s}
                      </span>
                    </div>

                    { (role === 'admin' || role === 'superadmin' || plan === 'paid') && !isMobile && hoveredItem === s && (
                      <div className="actions" style={{ display: 'flex', gap: 6 }}>
                        <button
                          className="icon-btn"
                          onClick={(e) => {
                            e.stopPropagation();
                            previewSheet(s);
                          }}
                          aria-label={`Open ${s}`}
                          title="Open"
                          style={{
                            background: '#fff',
                            border: '1px solid #e5e7eb',
                            borderRadius: 6,
                            padding: '4px 6px',
                          }}
                        >
                          <FaEdit />
                        </button>
                        <button
                          className="icon-btn caution"
                          onClick={(e) => {
                            e.stopPropagation();
                            deleteSheet(s);
                          }}
                          aria-label={`Delete ${s}`}
                          title="Delete"
                          style={{
                            background: '#fff',
                            border: '1px solid #fca5a5',
                            color: '#dc2626',
                            borderRadius: 6,
                            padding: '4px 6px',
                          }}
                        >
                          <FaTrash />
                        </button>
                      </div>
                    )}

                    {(role === 'admin' || role === 'superadmin' || plan === 'paid') && isMobile && (
                      <div
                        style={{
                          position: 'absolute',
                          right: 8,
                          top: '50%',
                          transform: 'translateY(-50%)',
                          display: translate <= -40 ? 'flex' : 'none',
                          gap: 6,
                        }}
                      >
                        <button
                          className="icon-btn caution"
                          onClick={(e) => {
                            e.stopPropagation();
                            deleteSheet(s);
                          }}
                          aria-label={`Delete ${s}`}
                          title="Delete"
                          style={{
                            background: '#fff',
                            border: '1px solid #fca5a5',
                            color: '#dc2626',
                            borderRadius: 6,
                            padding: '4px 6px',
                          }}
                        >
                          <FaTrash />
                        </button>
                      </div>
                    )}
                  </li>
                );
              })}
            </ul>

            {/* Ads disabled */}
            {false && (
              <div
                className="ad-slot"
                style={{
                  margin: 12,
                  padding: 12,
                  border: '1px dashed #e5e7eb',
                  borderRadius: 8,
                  textAlign: 'center',
                  color: '#6b7280',
                }}
              >
                <p>Ad space — Upgrade to remove ads</p>
              </div>
            )}
          </aside>
        )}

        {/* Right: Editor or Audit */}
        {(!isMobile || selectedSheet || showAudit) && (
          <main
            className="editor-area"
            style={{
              background: 'var(--bg, --text)',
              position: 'relative',
              height: '100%',
              display: 'flex',
              flexDirection: 'column',
            }}
          >
            {selectedSheet ? (
              <div
                className="card editor-card"
                style={{
                  display: 'flex',
                  flexDirection: 'column',
                  flex: 1,
                  borderRadius: 12,
                  border: '1px solid #e5e7eb',
                  padding: 0,
                  minHeight: 0,
                }}
              >
                <div
                  className="editor-header"
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    borderBottom: '1px solid #e5e7eb',
                    padding: '10px 16px',
                    flexShrink: 0,
                  }}
                >
                  <h2 style={{ margin: 0, fontSize: 18 }}>
                    {selectedSheet}
                    {forceReadOnly && (
                      <span style={{ marginLeft: 8, fontSize: 12, color: '#6b7280' }}>
                        (read-only)
                      </span>
                    )}
                  </h2>
                  <div className="editor-actions" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    {canEdit && !forceReadOnly && (
                      <button
                        className="primary small"
                        onClick={saveAllChanges}
                        disabled={loading}
                        style={{ padding: '4px 8px' }}
                      >
                        <FaSave />
                      </button>
                    )}
                    <button className="secondary small" onClick={exportSheet} style={{ padding: '4px 8px' }}>
                      <FaFileExport />
                    </button>
                    <button className="secondary small" onClick={downloadSheet} style={{ padding: '4px 8px' }}>
                      <FaDownload />
                    </button>
                    <button
                      className="secondary"
                      onClick={() => {
                        setSelectedSheet(null);
                        setForceReadOnly(false);
                      }}
                    >
                      <FaTimes />
                    </button>
                    <span className={`presence-indicator ${socketConnected ? 'online' : 'offline'}`}>
                      {socketConnected ? 'Live' : 'Offline'}
                    </span>
                  </div>
                </div>

                {/* Grid area scrolls fully; header/footer fixed */}
                <div
                  className="editor-scroll-wrapper"
                  style={{
                    flex: 1,
                    padding: 8,
                    overflow: 'hidden',
                    height: 'calc(100vh - 120px)',
                  }}
                >
                  <div className="editor">
                    <SheetEditor
                      data={stagedGrid}
                      canEdit={canEdit && !forceReadOnly}
                      onChange={(next) => setStagedGrid(next)}
                      onCellEdit={(op) => {
                        if (socketRef.current && selectedSheet && !forceReadOnly) {
                          socketRef.current.emit('cell-edit', {
                            room: `sheet:${selectedSheet}`,
                            ...op,
                          });
                        }
                      }}
                    />
                  </div>
                </div>
              </div>
            ) : showAudit ? (
              <div
                className="card editor-card"
                style={{
                  flex: 1,
                  borderRadius: 12,
                  border: '1px solid #e5e7eb',
                  display: 'flex',
                  flexDirection: 'column',
                  minHeight: 0,
                }}
              >
                <div
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    borderBottom: '1px solid #e5e7eb',
                    padding: '10px 16px',
                    flexShrink: 0,
                  }}
                >
                  <h2 style={{ margin: 0, fontSize: 18 }}>Audit Log</h2>
                  <button className="secondary" onClick={() => setShowAudit(false)}>
                    <FaTimes /> Close
                  </button>
                </div>
                <div style={{ flex: 1, overflowY: 'auto', minHeight: 0 }}>
                  <AuditTable
                    entries={auditEntries}
                    onOpen={(sheetName) => {
                      setShowAudit(false);
                      previewSheet(sheetName, { readOnly: true });
                    }}
                  />
                </div>
              </div>
            ) : (
              <div className="card" style={{ flex: 1 }}>
                <h2>Select a sheet to start editing</h2>
                <p className="muted">
                  Your sheets are listed {isMobile ? 'above' : 'on the left'}.
                </p>
              </div>
            )}
          </main>
        )}
      </div>

      {/* Footer */}
      <footer
        className="footer"
        style={{
          borderTop: '1px solid #e5e7eb',
          padding: '10px 16px',
          textAlign: 'center',
        }}
      >
        <p style={{ margin: 0 }}>
          {appName} © {new Date().getFullYear()}
        </p>
      </footer>

      <LoadingOverlay show={loading} />

      {/* Modals */}
      <Login
        show={showLogin}
        onClose={() => setShowLogin(false)}
        onShowRegister={() => {
          setShowLogin(false);
          setShowRegister(true);
        }}
      />
      <Register
        show={showRegister}
        onClose={() => setShowRegister(false)}
        onShowLogin={() => {
          setShowRegister(false);
          setShowLogin(true);
        }}
      />
    </div>
  );
}

/* =======================================================
   SheetEditor with HyperFormula + real-time hooks
======================================================= */
function SheetEditor({ data, onChange, canEdit, onCellEdit }) {
  const containerRef = useRef(null);
  const hotRef = useRef(null);
  const hfRef = useRef(null);

  // UI state for modals
  const [showFormulaPopup, setShowFormulaPopup] = useState(false);
  const [showRowPopup, setShowRowPopup] = useState(false);
  const [showFormulaHelp, setShowFormulaHelp] = useState(false);

  // Formula state
  const [selectedFormula, setSelectedFormula] = useState('');
  const [formulaInfo, setFormulaInfo] = useState('');
  const [inputCols, setInputCols] = useState('');
  const [outputCol, setOutputCol] = useState('');
  const [stopRow, setStopRow] = useState('');
  const [selectedHelp, setSelectedHelp] = useState('');

  // Row adding state
  const [rowsToAdd, setRowsToAdd] = useState(1);

  // Formula descriptions
  const formulaDescriptions = {
    SUM: 'Adds up all selected numbers.',
    AVERAGE: 'Calculates the average of selected numbers.',
    MULTIPLY: 'Multiplies two selected cells.',
    DIVIDE: 'Divides one cell by another.',
    SUBTRACT: 'Subtracts one cell or range from another.',
  };

  // Initialize Handsontable
  useEffect(() => {
    hfRef.current = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });

    const container = containerRef.current;
    const parent = container?.parentElement;

    let initialWidth = parent ? parent.clientWidth - 8 : 800;
    let initialHeight = parent ? parent.clientHeight - 8 : 500;

    if (window.innerWidth < initialWidth) {
      initialWidth = window.innerWidth - 16;
    }

    if (window.innerWidth >= 600) {
      initialHeight = 350;
    } else {
      initialHeight = window.innerHeight - 16;
    }

    hotRef.current = new Handsontable(container, {
      data: Array.isArray(data) ? data : [[]],
      licenseKey: 'non-commercial-and-evaluation',
      rowHeaders: true,
      colHeaders: true,
      formulas: { engine: hfRef.current },
      readOnly: !canEdit,
      autoFill: true,
      contextMenu: canEdit ? true : false,
      dropdownMenu: true,
      manualRowMove: true,
      manualColumnMove: true,
      manualColumnResize: true,
      stretchH: 'none',
      width: '100%',
      height: initialHeight,
      afterChange: (changes, source) => {
        if (!changes || source === 'loadData') return;
        const next = hotRef.current.getData();
        onChange(next);
        changes.forEach(([row, col, _oldVal, newVal]) => {
          onCellEdit?.({ rowIndex: row, colIndex: col, value: newVal });
        });
      },
    });

    window.__hotInstance = hotRef.current;

    return () => {
      try {
        hotRef.current?.destroy();
        window.__hotInstance = null;
      } catch {}
    };
  }, []);

  // Update data when props change
  useEffect(() => {
    if (!hotRef.current) return;
    hotRef.current.updateSettings({
      data: Array.isArray(data) ? data : [[]],
      readOnly: !canEdit,
    });
  }, [data, canEdit]);

  // Apply formula logic
  const applyFormula = () => {
    const hot = hotRef.current;
    if (!hot) return;

    const inputs = inputCols.split(',')
      .map((c) => c.trim().toUpperCase())
      .filter(Boolean);

    const out = (outputCol || '').trim().toUpperCase();
    const stop = parseInt(stopRow, 10) || hot.countRows();

    if (!selectedFormula || !out) {
      Toast.warn('Select a formula and output column');
      return;
    }

    const colIndex = (letter) => {
      const A = 'A'.charCodeAt(0);
      return letter ? letter.charCodeAt(0) - A : 0;
    };

    if (selectedFormula === 'SUM' && inputs.length === 1) {
      const startRow = 1;
      const endRow = stop - 1;
      const targetRow = stop;
      const formula = `=SUM(${inputs[0]}${startRow}:${inputs[0]}${endRow})`;
      hot.setDataAtCell(targetRow - 1, colIndex(out), formula);

    } else if (selectedFormula === 'AVERAGE' && inputs.length === 1) {
      const startRow = 1;
      const endRow = stop - 1;
      const targetRow = stop;
      const formula = `=AVERAGE(${inputs[0]}${startRow}:${inputs[0]}${endRow})`;
      hot.setDataAtCell(targetRow - 1, colIndex(out), formula);

    } else if (selectedFormula === 'MULTIPLY' && inputs.length === 1) {
      const startRow = 1;
      const endRow = stop - 1;
      const targetRow = stop;
      const formula = `=PRODUCT(${inputs[0]}${startRow}:${inputs[0]}${endRow})`;
      hot.setDataAtCell(targetRow - 1, colIndex(out), formula);

    } else if (selectedFormula === 'DIVIDE' && inputs.length === 1) {
      const startRow = 1;
      const endRow = stop - 1;
      const targetRow = stop;
      const formula = `=${inputs[0]}${startRow}/PRODUCT(${inputs[0]}${startRow+1}:${inputs[0]}${endRow})`;
      hot.setDataAtCell(targetRow - 1, colIndex(out), formula);

    } else if (selectedFormula === 'SUBTRACT' && inputs.length === 1) {
      const startRow = 1;
      const endRow = stop - 1;
      const targetRow = stop;
      const formula = `=${inputs[0]}${startRow}-SUM(${inputs[0]}${startRow+1}:${inputs[0]}${endRow})`;
      hot.setDataAtCell(targetRow - 1, colIndex(out), formula);

    } else {
      for (let r = 0; r < stop; r++) {
        let formula = '';
        if (selectedFormula === 'SUM') {
          formula = `=SUM(${inputs.map((i) => `${i}${r + 1}`).join(',')})`;
        } else if (selectedFormula === 'AVERAGE') {
          formula = `=AVERAGE(${inputs.map((i) => `${i}${r + 1}`).join(',')})`;
        } else if (selectedFormula === 'MULTIPLY') {
          if (inputs.length < 2) continue;
          formula = `=${inputs[0]}${r + 1}*${inputs[1]}${r + 1}`;
        } else if (selectedFormula === 'DIVIDE') {
          if (inputs.length < 2) continue;
          formula = `=${inputs[0]}${r + 1}/${inputs[1]}${r + 1}`;
        } else if (selectedFormula === 'SUBTRACT') {
          if (inputs.length < 2) continue;
          formula = `=${inputs[0]}${r + 1}-${inputs[1]}${r + 1}`;
        }
        if (formula) {
          hot.setDataAtCell(r, colIndex(out), formula);
        }
      }
    }

    setShowFormulaPopup(false);
    Toast.success('Formula applied');
  };

  // Add rows logic
  const addRows = () => {
    const hot = hotRef.current;
    if (!hot) return;
    const count = parseInt(rowsToAdd, 10) || 1;
    const lastIndex = Math.max(hot.countRows() - 1, 0);
    hot.alter('insert_row_below', lastIndex, count);
    setShowRowPopup(false);
  };

  // Update formula info when selection changes
  useEffect(() => {
    const info = selectedFormula ? formulaDescriptions[selectedFormula] : '';
    setFormulaInfo(info || '');
  }, [selectedFormula]);

  // Render
  return (
    <div className="sheet-editor-wrapper">
      {/* Toolbar */}
      <div
        className="sticky-toolbar"
        style={{
          position: 'sticky',
          top: 0,
          display: 'flex',
          alignItems: 'center',
          gap: 8,
          padding: 2,
          background: 'var(--bg, --text)',
          zIndex: 5,
          borderBottom: '1px solid #e5e7eb',
        }}
      >
        <button className="secondary small" onClick={() => setShowFormulaPopup(true)}>
          Formula
        </button>
        <button className="secondary small" onClick={() => setShowRowPopup(true)}>
          +Row
        </button>
      </div>

      {/* Formula modal */}
      {showFormulaPopup && (
        <div
          style={{
            position: 'fixed',
            top: '20%',
            left: '50%',
            transform: 'translateX(-50%)',
            background: 'var(--bg, --text)',
            border: '1px solid #e5e7eb',
            borderRadius: 8,
            padding: 16,
            zIndex: 1000,
            boxShadow: '0 8px 20px rgba(0,0,0,0.25)',
            width: 320,
          }}
        >
          <h3 style={{ marginTop: 0 }}>Apply Formula</h3>
          <label style={{ display: 'block', marginBottom: 8 }}>
            Choose formula:
            <select
              value={selectedFormula}
              onChange={(e) => {
                const val = e.target.value;
                setSelectedFormula(val);
                setFormulaInfo(formulaDescriptions[val] || '');
              }}
              style={{ marginLeft: 8 }}
            >
              <option value="">-- Select --</option>
              <option value="SUM">SUM</option>
              <option value="AVERAGE">AVERAGE</option>
              <option value="MULTIPLY">MULTIPLY</option>
              <option value="DIVIDE">DIVIDE</option>
              <option value="SUBTRACT">SUBTRACT</option>
            </select>
            <button
              type="button"
              style={{ marginLeft: 12 }}
              onClick={() => setShowFormulaHelp(true)}
            >
              Help
            </button>
          </label>
          {formulaInfo && (
            <p style={{ fontSize: 12, color: '#6b7280' }}>{formulaInfo}</p>
          )}
          <label>
            Input cols:
            <input
              value={inputCols}
              onChange={(e) => setInputCols(e.target.value)}
              style={{ width: 50, marginLeft: 8 }}
              placeholder="B,C"
            />
          </label>
          <label>
            Output col:
            <input
              value={outputCol}
              onChange={(e) => setOutputCol(e.target.value)}
              style={{ width: 50, marginLeft: 8 }}
              placeholder="D"
            />
          </label>
          <label>
            Stop row:
            <input
              value={stopRow}
              onChange={(e) => setStopRow(e.target.value)}
              style={{ width: 50, marginLeft: 8 }}
              placeholder="20"
            />
          </label>
          <div
            style={{
              marginTop: 12,
              display: 'flex',
              gap: 8,
              justifyContent: 'flex-end',
            }}
          >
            <button
              className="secondary small"
              onClick={() => setShowFormulaPopup(false)}
            >
              Cancel
            </button>
            <button className="primary small" onClick={applyFormula}>
              Apply
            </button>
          </div>
        </div>
      )}

      {/* Formula Help Modal */}
      {showFormulaHelp && (
        <div
          style={{
            position: 'fixed',
            top: '15%',
            left: '50%',
            transform: 'translateX(-50%)',
            background: 'var(--bg, --text)',
            border: '1px solid #e5e7eb',
            borderRadius: 8,
            padding: 16,
            zIndex: 1100,
            boxShadow: '0 8px 20px rgba(0,0,0,0.25)',
            width: 420,
          }}
        >
          <h3 style={{ marginTop: 0 }}>Formula Examples</h3>
          <ul style={{ display: 'flex', listStyle: 'none', padding: 0, fontSize: 10, gap: 8 }}>
            <li><button onClick={() => setSelectedHelp('SUM')}>SUM</button></li>
            <li><button onClick={() => setSelectedHelp('AVERAGE')}>AVERAGE</button></li>
            <li><button onClick={() => setSelectedHelp('MULTIPLY')}>MULTIPLY</button></li>
            <li><button onClick={() => setSelectedHelp('DIVIDE')}>DIVIDE</button></li>
            <li><button onClick={() => setSelectedHelp('SUBTRACT')}>SUBTRACT</button></li>
          </ul>

          {selectedHelp && (
            <div style={{ marginTop: 12, fontSize: 14 }}>
              {selectedHelp === 'SUM' && (
                <>
                  <h4>SUM</h4>
                  <p>Adds numbers together.</p>
                  <p><b>Row-wise:</b> Input cols: <code>B,C</code>, Output col: <code>D</code>, Stop row: <code>14</code></p>
                  <p>Result: <code>D1 = SUM(B1:C1), D2 = SUM(B2:C2) … up to D14 = SUM(B14:C14)</code></p>
                  <p><b>Aggregate:</b> Input col: <code>F</code>, Output col: <code>F</code>, Stop row: <code>14</code></p>
                  <p>Result: <code>F14 = SUM(F1:F13)</code> → sums F1 through F13, total in F14</p>
                </>
              )}

              {selectedHelp === 'AVERAGE' && (
                <>
                  <h4>AVERAGE</h4>
                  <p>Calculates the mean of numbers.</p>
                  <p><b>Row-wise:</b> Input cols: <code>B,C</code>, Output col: <code>D</code>, Stop row: <code>14</code></p>
                  <p>Result: <code>D1 = AVERAGE(B1:C1), D2 = AVERAGE(B2:C2) … up to D14</code></p>
                  <p><b>Aggregate:</b> Input col: <code>E</code>, Output col: <code>E</code>, Stop row: <code>14</code></p>
                  <p>Result: <code>E14 = AVERAGE(E1:E13)</code> → average of E1 through E13 in E14</p>
                </>
              )}

              {selectedHelp === 'MULTIPLY' && (
                <>
                  <h4>MULTIPLY</h4>
                  <p>Multiplies values.</p>
                  <p><b>Row-wise:</b> Input cols: <code>A,B</code>, Output col: <code>C</code>, Stop row: <code>10</code></p>
                  <p>Result: <code>C1 = A1*B1, C2 = A2*B2 … up to C10</code></p>
                  <p><b>Aggregate:</b> Input col: <code>F</code>, Output col: <code>F</code>, Stop row: <code>14</code></p>
                  <p>Result: <code>F14 = PRODUCT(F1:F13)</code> → multiplies all values F1 through F13, result in F14</p>
                </>
              )}

              {selectedHelp === 'DIVIDE' && (
                <>
                  <h4>DIVIDE</h4>
                  <p>Divides values.</p>
                  <p><b>Row-wise:</b> Input cols: <code>A,B</code>, Output col: <code>C</code>, Stop row: <code>10</code></p>
                  <p>Result: <code>C1 = A1/B1, C2 = A2/B2 … up to C10</code></p>
                  <p><b>Aggregate:</b> Input col: <code>F</code>, Output col: <code>F</code>, Stop row: <code>14</code></p>
                  <p>Result: <code>F14 = F1/PRODUCT(F2:F13)</code> → divides F1 by the product of F2 through F13, result in F14</p>
                </>
              )}

              {selectedHelp === 'SUBTRACT' && (
                <>
                  <h4>SUBTRACT</h4>
                  <p>Subtracts values.</p>
                  <p><b>Row-wise:</b> Input cols: <code>A,B</code>, Output col: <code>C</code>, Stop row: <code>10</code></p>
                  <p>Result: <code>C1 = A1-B1, C2 = A2-B2 … up to C10</code></p>
                  <p><b>Aggregate:</b> Input col: <code>F</code>, Output col: <code>F</code>, Stop row: <code>14</code></p>
                  <p>Result: <code>F14 = F1-SUM(F2:F13)</code> → subtracts the sum of F2 through F13 from F1, result in F14</p>
                </>
              )}
            </div>
          )}

          <div style={{ marginTop: 16, textAlign: 'right' }}>
            <button className="secondary small" onClick={() => setShowFormulaHelp(false)}>Close</button>
          </div>
        </div>
      )}

      {/* Add rows modal */}
      {showRowPopup && (
        <div
          style={{
            position: 'fixed',
            top: '30%',
            left: '50%',
            transform: 'translateX(-50%)',
            background: 'var(--bg, --text)',
            border: '1px solid #e5e7eb',
            borderRadius: 8,
            padding: 16,
            zIndex: 1000,
            boxShadow: '0 8px 20px rgba(0,0,0,0.25)',
            width: 220,
          }}
        >
          <h3 style={{ marginTop: 0 }}>Add Rows</h3>
          <label>
            Number of rows:
            <input
              type="number"
              value={rowsToAdd}
              onChange={(e) => setRowsToAdd(parseInt(e.target.value, 10))}
              style={{ width: 50, marginLeft: 8 }}
            />
          </label>
          <div style={{ marginTop: 12, display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
            <button className="secondary small" onClick={() => setShowRowPopup(false)}>Cancel</button>
            <button className="primary small" onClick={addRows}>Add</button>
          </div>
        </div>
      )}

      {/* Handsontable container */}
      <div
        className="hot-container"
        ref={containerRef}
        style={{
          background: 'var(--bg, --text)',
          zIndex: 1,
          width: '100%',
          height: '100%',
        }}
      />
    </div>
  );
}

/* =======================================================
   Read-only audit list that opens a sheet in read-only mode
======================================================= */
function AuditTable({ entries, onOpen }) {
  return (
    <div
      className="audit-table"
      style={{
        display: 'grid',
        gap: 8,
        padding: 12,
      }}
    >
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: '160px 1fr 1fr 120px 80px',
          fontWeight: 600,
          borderBottom: '1px solid #e5e7eb',
          paddingBottom: 8,
        }}
      >
        <span>Time</span>
        <span>Actor</span>
        <span>Action</span>
        <span>Sheet</span>
        <span>Open</span>
      </div>

      {entries && entries.length > 0 ? (
        entries.map((e, idx) => (
          <div
            key={idx}
            style={{
              display: 'grid',
              gridTemplateColumns: '160px 1fr 1fr 120px 80px',
              padding: '8px 0',
              borderBottom: '1px dashed #eee',
              alignItems: 'center',
            }}
          >
            <span style={{ color: '#6b7280' }}>{formatTS(e.ts)}</span>
            <span>{e.actor || e.email || '—'}</span>
            <span>{e.action || '—'}</span>
            <span>{e.sheet || '—'}</span>
            <button
              className="secondary small"
              onClick={() => e.sheet && onOpen?.(e.sheet)}
              disabled={!e.sheet}
            >
              Open
            </button>
          </div>
        ))
      ) : (
        <p className="muted" style={{ color: '#6b7280' }}>No audit entries yet.</p>
      )}
    </div>
  );
}
