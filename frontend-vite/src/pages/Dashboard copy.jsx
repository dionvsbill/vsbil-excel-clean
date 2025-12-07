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
import './preview-table.css';
import './dashboard.css';

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

  const [showLogin, setShowLogin] = useState(false);
  const [showRegister, setShowRegister] = useState(false);

  // Actions dropdown (header-only)
  const [actionsOpen, setActionsOpen] = useState(false);
  const [showAudit, setShowAudit] = useState(false);

  const [theme, setTheme] = useState('system');
  const [searchQuery, setSearchQuery] = useState('');

  const [socketConnected, setSocketConnected] = useState(false);
  const socketRef = useRef(null);

  const [isDragOver, setIsDragOver] = useState(false);

  const API_BASE = import.meta.env.VITE_API_URL || 'http://localhost:4000';

  // Mobile breakpoint
  const [isMobile, setIsMobile] = useState(window.innerWidth < 992);
  useEffect(() => {
    const onResize = () => setIsMobile(window.innerWidth < 992);
    window.addEventListener('resize', onResize);
    return () => window.removeEventListener('resize', onResize);
  }, []);

  // Hover and swipe states
  const [hoveredItem, setHoveredItem] = useState(null);
  const [swipeX, setSwipeX] = useState({});
  const touchStartXRef = useRef({});

  // Audit entries
  const [auditEntries, setAuditEntries] = useState([]);

  // Session sync
  useEffect(() => {
    supabase.auth.getUser().then(({ data }) => {
      if (data.user) updateUser(data.user);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      if (session?.user) updateUser(session.user);
      else {
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
      setRole(data?.role || 'user');
      setCanEdit(Boolean(data?.can_edit) || data?.role === 'admin');
      setPlan(data?.plan || 'free');
      if (data?.app_name) setAppName(data.app_name);
    } catch (err) {
      Toast.error(err.message);
      setRole('user');
      setCanEdit(false);
      setPlan('free');
    }
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
      return await res.json();
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
      return await res.json();
    } catch (e) {
      Toast.error(e.message);
      return { error: e.message };
    } finally {
      setLoading(false);
    }
  }

  // Load sheets
  useEffect(() => {
    (async () => {
      const j = await apiGet('/excel/sheets');
      if (j.error) Toast.error(j.error);
      else setSheets(j.sheets || []);
    })();
  }, []);

  // Load audit entries when showAudit toggles on
  useEffect(() => {
    (async () => {
      if (!showAudit) return;
      const j = await apiGet('/audit/list');
      if (j.error) Toast.error(j.error);
      else setAuditEntries(Array.isArray(j.entries) ? j.entries : []);
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

  // Real-time
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

  const filteredSheets = useMemo(() => {
    if (!searchQuery.trim()) return sheets;
    return sheets.filter((s) => s.toLowerCase().includes(searchQuery.toLowerCase()));
  }, [sheets, searchQuery]);

  async function previewSheet(name) {
    const j = await apiGet(`/excel/preview?sheet=${encodeURIComponent(name)}`);
    if (j.error) {
      Toast.error(j.error);
    } else {
      setSelectedSheet(name);
      setStagedGrid(j.preview || []);
    }
  }

  async function saveAllChanges() {
    if (!selectedSheet || stagedGrid.length === 0) {
      Toast.error('No sheet selected');
      return;
    }
    if (plan === 'free') {
      Toast.warn('Export and Save are premium features. Please upgrade.');
      return;
    }
    const j = await apiPost('/excel/save-all', { sheet: selectedSheet, data: stagedGrid });
    if (j.error) {
      Toast.error(j.error);
    } else {
      Toast.success(`Saved changes to ${selectedSheet}`);
    }
  }

  async function deleteSheet(name) {
    if (role !== 'admin') {
      Toast.warn('Only admin can delete sheets');
      return;
    }
    const confirmDelete = window.confirm(`Delete sheet "${name}"?`);
    if (!confirmDelete) return;
    const res = await apiPost('/excel/delete-sheet', { name });
    if (res.error) {
      Toast.error(res.error);
    } else {
      Toast.success(`Deleted "${name}"`);
      setSheets((prev) => prev.filter((s) => s !== name));
      if (selectedSheet === name) {
        setSelectedSheet(null);
        setStagedGrid([]);
      }
    }
  }

  async function addSheet() {
    const name = prompt('Enter new sheet name:');
    if (!name) return;
    const j = await apiPost('/excel/add-sheet', { name });
    if (j.error) {
      Toast.error(j.error);
    } else {
      Toast.success(`Sheet ${name} added`);
      if (j.sheets) {
        setSheets(j.sheets);
      } else {
        setSheets((prev) => [...prev, name]);
      }
      setSelectedSheet(name);
      previewSheet(name);
    }
  }

  // Editor helpers (Handsontable instance accessed via window.__hotInstance)
  const addRowAbove = () => {
    const hot = window.__hotInstance;
    if (!hot || !canEdit) return;
    const sel = hot.getSelectedLast();
    const rowIndex = sel ? sel[0] : 0;
    hot.alter('insert_row', rowIndex);
  };

  const addRowBelow = () => {
    const hot = window.__hotInstance;
    if (!hot || !canEdit) return;
    const sel = hot.getSelectedLast();
    const rowIndex = sel ? sel[0] : hot.countRows() - 1;
    hot.alter('insert_row', rowIndex + 1);
  };

  const addColLeft = () => {
    const hot = window.__hotInstance;
    if (!hot || !canEdit) return;
    const sel = hot.getSelectedLast();
    const colIndex = sel ? sel[1] : 0;
    hot.alter('insert_col', colIndex);
  };

  const addColRight = () => {
    const hot = window.__hotInstance;
    if (!hot || !canEdit) return;
    const sel = hot.getSelectedLast();
    const colIndex = sel ? sel[1] : hot.countCols() - 1;
    hot.alter('insert_col', colIndex + 1);
  };

  // Swipe for mobile
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

  // Export and Download
  async function exportSheet() {
    if (!selectedSheet) {
      Toast.error('No sheet selected');
      return;
    }
    if (plan === 'free') {
      Toast.warn('Export is a premium feature. Please upgrade.');
      return;
    }
    const j = await apiPost('/excel/export', { sheet: selectedSheet });
    if (j.error) {
      Toast.error(j.error);
    } else {
      Toast.success(`Export started for ${selectedSheet}`);
    }
  }

  async function downloadSheet() {
    if (!selectedSheet) {
      Toast.error('No sheet selected');
      return;
    }
    if (plan === 'free') {
      Toast.warn('Download is a premium feature. Please upgrade.');
      return;
    }
    try {
      setLoading(true);
      const headers = await authHeader();
      const res = await fetch(`${API_BASE}/excel/download?sheet=${encodeURIComponent(selectedSheet)}`, {
        method: 'GET',
        headers,
      });
      if (!res.ok) {
        const txt = await res.text();
        try {
          const j = JSON.parse(txt);
          Toast.error(j.error || 'Download failed');
        } catch {
          Toast.error('Download failed');
        }
        return;
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${selectedSheet}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      Toast.success(`Downloaded ${selectedSheet}`);
    } catch (err) {
      Toast.error(err.message);
    } finally {
      setLoading(false);
    }
  }

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
              style={{ border: '1px solid #e5e7eb', borderRadius: 8, padding: '6px 8px' }}
            />
          </div>

          <div
            className="nav-actions"
            style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 8 }}
          >
            {user && plan === 'free' && <PayButton user={user} amount={500000} currency="GHS" />}
          </div>

          {/* Actions dropdown anchored under the menu button */}
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
                zIndex: 1000,
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
                <button
                  className="secondary"
                  onClick={() => setActionsOpen(false)}
                  style={{ background: 'var(--danger, #fff)' }}
                >
                  <FaTimes /> Close
                </button>
              </div>
              <div className="dropdown-body" style={{ display: 'grid', gap: 8 }}>
                <InstallButton />
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
                {user && (
                  <button
                    className="secondary"
                    onClick={() => {
                      supabase.auth.signOut();
                      setActionsOpen(false);
                    }}
                    style={{ background: 'var(--danger, #fff)' }}
                  >
                    Logout
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
                {/* Audit button */}
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

      {/* Content grid: two columns desktop, single view mobile */}
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
        {/* List area */}
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
              }}
            >
              <button className="primary small" onClick={addSheet}>
                <FaPlus /> New sheet
              </button>
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

                    {/* Desktop hover actions: show edit/delete icons on hover */}
                    {role === 'admin' && !isMobile && hoveredItem === s && (
                      <div className="actions" style={{ display: 'flex', gap: 6 }}>
                        <button
                          className="icon-btn"
                          onClick={(e) => {
                            e.stopPropagation();
                            previewSheet(s);
                          }}
                          aria-label={`Edit ${s}`}
                          title="Edit"
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

                    {/* Mobile swipe-left delete reveal */}
                    {role === 'admin' && isMobile && (
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

                try {
                  const formData = new FormData();
                  formData.append('file', file);

                  const headers = await authHeader();

                  const res = await fetch(`${API_BASE}/excel/upload`, {
                    method: 'POST',
                    headers, // only Authorization
                    body: formData,
                  });

                  const raw = await res.text();
                  let j;
                  try {
                    j = JSON.parse(raw);
                  } catch {
                    console.error('Upload response was not JSON:', raw);
                    Toast.error('Upload failed: bad response');
                    return;
                  }

                  if (j.error) {
                    Toast.error(j.error);
                  } else {
                    Toast.success(`Uploaded ${j.fileName}`);

                    if (Array.isArray(j.sheetNames) && j.sheetNames.length > 0) {
                      setSheets((prev) => [...prev, ...j.sheetNames]);
                      setSelectedSheet(j.sheetNames[0]);
                      setStagedGrid(j.preview || []);
                    } else {
                      const newSheetName = j.fileName.replace(/\.[^/.]+$/, '');
                      setSheets((prev) => [...prev, newSheetName]);
                      setSelectedSheet(newSheetName);
                      setStagedGrid(j.preview || []);
                    }
                  }
                } catch (err) {
                  console.error('Upload failed:', err);
                  Toast.error('Upload failed');
                }
              }}
              style={{
                margin: 12,
                padding: 12,
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
          </aside>
        )}

        {/* Editor area */}
        {(!isMobile || selectedSheet || showAudit) && (
          <main
            className="editor-area"
            style={{
              overflow: 'hidden', // container shouldn't scroll the page
              background: 'var(--bg, --text)',
            }}
          >
            {selectedSheet ? (
              <div
                className="card editor-card"
                style={{
                  display: 'flex',
                  flexDirection: 'column',
                  height: '100%',
                  borderRadius: 12,
                  border: '1px solid #e5e7eb',
                  padding: 0,
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
                  }}
                >
                  <h2 style={{ margin: 0, fontSize: 18 }}>{selectedSheet}</h2>
                  <div className="editor-actions" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    {canEdit && (
                      <button
                        className="primary small"
                        onClick={saveAllChanges}
                        disabled={loading}
                        style={{ padding: '4px 8px' }}
                      >
                        <FaSave />
                      </button>
                    )}
                    {/* Export and Download small buttons */}
                    <button
                      className="secondary small"
                      onClick={exportSheet}
                      style={{ padding: '4px 8px' }}
                    >
                      <FaFileExport />
                    </button>
                    <button
                      className="secondary small"
                      onClick={downloadSheet}
                      style={{ padding: '4px 8px' }}
                    >
                      <FaDownload />
                    </button>
                    <button className="secondary" onClick={() => setSelectedSheet(null)}>
                      <FaTimes />
                    </button>
                    <span className={`presence-indicator ${socketConnected ? 'online' : 'offline'}`}>
                      {socketConnected ? 'Live' : 'Offline'}
                    </span>
                  </div>
                </div>

                {/* Editor scroll wrapper */}
                <div
                  className="editor-scroll-wrapper"
                  style={{
                    flex: 1,
                    overflowX: 'auto',
                    overflowY: 'auto',
                    padding: 8,
                  }}
                >
                  <SheetEditor
                    data={stagedGrid}
                    canEdit={canEdit}
                    onChange={(next) => setStagedGrid(next)}
                    onCellEdit={(op) => {
                      if (socketRef.current && selectedSheet) {
                        socketRef.current.emit('cell-edit', {
                          room: `sheet:${selectedSheet}`,
                          ...op,
                        });
                      }
                    }}
                  />
                </div>

                {canEdit && (
                  <div
                    className="editor-footer-actions"
                    style={{
                      borderTop: '1px solid #e5e7eb',
                      padding: 12,
                      display: 'flex',
                      justifyContent: 'flex-end',
                    }}
                  >
                    <button className="primary" onClick={saveAllChanges} disabled={loading}>
                      <FaSave /> Save changes
                    </button>
                  </div>
                )}
              </div>
            ) : showAudit ? (
              <div className="card editor-card" style={{ height: '100%', overflowY: 'auto' }}>
                <h2 style={{ margin: '0 0 12px 0', fontSize: 18 }}>Audit Log</h2>
                {/* Audit table styled like sheet list, read-only */}
                <AuditTable entries={auditEntries} />
              </div>
            ) : (
              <div className="card" style={{ height: '100%' }}>
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

/* Handsontable-based editor with HyperFormula and real-time hooks */
function SheetEditor({ data, onChange, canEdit, onCellEdit }) {
  const containerRef = useRef(null);
  const hotRef = useRef(null);
  const hfRef = useRef(null);

  const [showFormulaPopup, setShowFormulaPopup] = useState(false);
  const [showRowPopup, setShowRowPopup] = useState(false);
  const [selectedFormula, setSelectedFormula] = useState('');
  const [formulaInfo, setFormulaInfo] = useState('');
  const [inputCols, setInputCols] = useState('');
  const [outputCol, setOutputCol] = useState('');
  const [stopRow, setStopRow] = useState('');
  const [rowsToAdd, setRowsToAdd] = useState(1);

  const formulaDescriptions = {
    SUM: 'Adds up all selected numbers.',
    AVERAGE: 'Calculates the average of selected numbers.',
    MULTIPLY: 'Multiplies two selected cells.',
    DIVIDE: 'Divides one cell by another.',
  };

  // Initialize Handsontable
  useEffect(() => {
    hfRef.current = HyperFormula.buildEmpty({
      licenseKey: 'gpl-v3', // or 'non-commercial-and-evaluation'
    });

    const container = containerRef.current;

    hotRef.current = new Handsontable(container, {
      data,
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
      stretchH: 'all',
      width: container?.parentElement ? container.parentElement.clientWidth - 8 : '100%',
      height: container?.parentElement ? container.parentElement.clientHeight - 8 : 500,
      afterChange: (changes, source) => {
        if (!changes || source === 'loadData') return;
        const next = hotRef.current.getData();
        onChange(next);
        changes.forEach(([row, col, _oldVal, newVal]) => {
          onCellEdit?.({ rowIndex: row, colIndex: col, value: newVal });
        });
      },
    });

    // expose instance for external helpers
    window.__hotInstance = hotRef.current;

    const ro = new ResizeObserver(() => {
      if (!hotRef.current || !container.parentElement) return;
      const { clientWidth, clientHeight } = container.parentElement;
      const newWidth = clientWidth - 8;
      const newHeight = clientHeight - 8;

      const settings = hotRef.current.getSettings();
      if (settings.width !== newWidth || settings.height !== newHeight) {
        hotRef.current.updateSettings({
          width: newWidth,
          height: newHeight,
        });
        hotRef.current.render();
      }
    });

    if (container.parentElement) ro.observe(container.parentElement);

    return () => {
      ro.disconnect();
      hotRef.current?.destroy();
      hfRef.current?.destroy();
      if (window.__hotInstance === hotRef.current) {
        window.__hotInstance = null;
      }
    };
  }, [canEdit]);

  useEffect(() => {
    if (hotRef.current) {
      hotRef.current.loadData(data || []);
    }
  }, [data]);

  const applyFormula = () => {
    try {
      if (!hotRef.current || !selectedFormula || !inputCols || !outputCol) return;

      const isRange = /^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/i.test(inputCols.trim());
      const output = outputCol.trim().toUpperCase();

      // Helper: parse "D7" -> { colIndex, rowIndex }
      const parseCellRef = (ref) => {
        const m = ref.match(/^([A-Z]+)([0-9]+)$/i);
        if (!m) return null;
        const colLetters = m[1].toUpperCase();
        const rowNumber = parseInt(m[2], 10); // spreadsheet 1-based
        return {
          colIndex: Handsontable.helper.spreadsheetColumnIndex(colLetters),
          rowIndex: rowNumber - 1, // convert to 0-based for HOT
        };
      };

      if (isRange) {
        // Grand total mode: inputCols like "D1:D6", outputCol like "D7"
        const out = parseCellRef(output);
        if (!out) return;

        const range = inputCols.trim().toUpperCase(); // e.g., "D1:D6"
        let formula = '';
        switch (selectedFormula) {
          case 'SUM':
            formula = `=SUM(${range})`;
            break;
          case 'AVERAGE':
            formula = `=AVERAGE(${range})`;
            break;
          default:
            formula = `=SUM(${range})`;
        }

        hotRef.current.setDataAtCell(out.rowIndex, out.colIndex, formula);
        return;
      }

      // Per-row builder mode: inputs like "B,C", output like "D", optional stopRow
      const inputs = inputCols.split(',').map(c => c.trim().toUpperCase()).filter(Boolean);
      const outputColLetter = output; // e.g., "D"
      const outColIndex = Handsontable.helper.spreadsheetColumnIndex(outputColLetter);
      const totalRows = hotRef.current.countRows();
      const stop = stopRow ? Math.min(parseInt(stopRow, 10), totalRows) : totalRows;

      for (let r = 0; r < stop; r++) {
        // r is 0-based; spreadsheet row number is r+1
        let formula = '';
        switch (selectedFormula) {
          case 'SUM':
            // SUM of cells in the same row across inputs (e.g., B1+C1)
            formula = `=SUM(${inputs.map(c => `${c}${r + 1}`).join(',')})`;
            break;
          case 'AVERAGE':
            formula = `=AVERAGE(${inputs.map(c => `${c}${r + 1}`).join(',')})`;
            break;
          case 'MULTIPLY':
            if (inputs.length >= 2) formula = `=${inputs[0]}${r + 1}*${inputs[1]}${r + 1}`;
            break;
          case 'DIVIDE':
            if (inputs.length >= 2) formula = `=${inputs[0]}${r + 1}/${inputs[1]}${r + 1}`;
            break;
        }
        if (formula) hotRef.current.setDataAtCell(r, outColIndex, formula);
      }

      // Add a total formula one row after stop (guard if within bounds)
      if (stop > 0) {
        const totalRowIndex = Math.min(stop, totalRows - 1);
        hotRef.current.setDataAtCell(
          totalRowIndex,
          outColIndex,
          `=SUM(${outputColLetter}1:${outputColLetter}${stop})`
        );
      }
    } finally {
      setShowFormulaPopup(false);
    }
  };

  const addRows = () => {
    try {
      if (!hotRef.current) return;
      const lastRow = hotRef.current.countRows() - 1;
      hotRef.current.alter('insert_row', lastRow >= 0 ? lastRow : 0, rowsToAdd);
    } finally {
      setShowRowPopup(false);
    }
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8, position: 'relative' }}>
      {canEdit && (
        <div
          className="toolbar"
          style={{
            display: 'flex',
            gap: 8,
            padding: '6px 12px',
            borderBottom: '1px solid #e5e7eb',
            background: '#f9fafb',
            position: 'sticky',
            top: 0,
            zIndex: 200,
          }}
        >
          <button className="secondary small" onClick={() => setShowFormulaPopup(true)}>
            Formula Builder
          </button>
          <button
            className="secondary small"
            onClick={() => setShowRowPopup(true)}
            style={{ fontSize: 12, padding: '2px 6px' }}
          >
            + Rows
          </button>
        </div>
      )}

      {(showFormulaPopup || showRowPopup) && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            width: '100%',
            height: '100%',
            background: 'rgba(0,0,0,0.3)',
            zIndex: 999,
          }}
          onClick={() => {
            setShowFormulaPopup(false);
            setShowRowPopup(false);
          }}
        />
      )}

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
          <h3 style={{ marginTop: 0 }}>Formula Builder</h3>
          <label>
            Formula:
            <select
              value={selectedFormula}
              onChange={(e) => {
                setSelectedFormula(e.target.value);
                setFormulaInfo(formulaDescriptions[e.target.value] || '');
              }}
              style={{ marginLeft: 8 }}
            >
              <option value="">-- Select --</option>
              <option value="SUM">SUM</option>
              <option value="AVERAGE">AVERAGE</option>
              <option value="MULTIPLY">MULTIPLY</option>
              <option value="DIVIDE">DIVIDE</option>
            </select>
          </label>
          {formulaInfo && <p style={{ fontSize: 12, color: '#6b7280' }}>{formulaInfo}</p>}

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

          <div style={{ marginTop: 12, display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
            <button className="secondary small" onClick={() => setShowFormulaPopup(false)}>Cancel</button>
            <button className="primary small" onClick={applyFormula}>Apply</button>
          </div>
        </div>
      )}

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

      <div
        className="hot-container"
        ref={containerRef}
        style={{
          minWidth: 900,
          minHeight: 500,
          background: 'var(--bg, --text)',
          zIndex: 1,
        }}
      />
    </div>
  );
}

/* Simple read-only audit table */
function AuditTable({ entries }) {
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
          gridTemplateColumns: '160px 1fr 1fr 120px',
          fontWeight: 600,
          borderBottom: '1px solid #e5e7eb',
          paddingBottom: 8,
        }}
      >
        <span>Time</span>
        <span>Actor</span>
        <span>Action</span>
        <span>Sheet</span>
      </div>

      {entries && entries.length > 0 ? (
        entries.map((e, idx) => (
          <div
            key={idx}
            style={{
              display: 'grid',
              gridTemplateColumns: '160px 1fr 1fr 120px',
              padding: '8px 0',
              borderBottom: '1px dashed #eee',
              alignItems: 'center',
            }}
          >
            <span style={{ color: '#6b7280' }}>{formatTS(e.ts)}</span>
            <span>{e.actor || e.email || '—'}</span>
            <span>{e.action || '—'}</span>
            <span>{e.sheet || '—'}</span>
          </div>
        ))
      ) : (
        <p className="muted" style={{ color: '#6b7280' }}>No audit entries yet.</p>
      )}
    </div>
  );
}

function formatTS(ts) {
  try {
    const d = new Date(ts);
    return d.toLocaleString();
  } catch {
    return '—';
  }
}


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
import './preview-table.css';
import './dashboard.css';

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

  // Ads gating: show ads only for free plan
  const shouldShowAds = plan === 'free';

  // Unified response parsing to avoid <!DOCTYPE... blowing up JSON.parse
  async function parseResponse(res) {
    const contentType = res.headers.get('content-type') || '';
    if (contentType.includes('application/json')) {
      return await res.json();
    }
    // Try blob for file responses; fallback to text
    const text = await res.text();
    // If HTML, return as error
    if (/<!DOCTYPE|<html/i.test(text)) {
      return { error: 'Unexpected HTML response from server', raw: text };
    }
    try {
      return JSON.parse(text);
    } catch {
      return { error: 'Unexpected response from server', raw: text };
    }
  }

  // Session sync
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
      setRole(data?.role || 'user');
      setCanEdit(Boolean(data?.can_edit) || data?.role === 'admin');
      setPlan(data?.plan || 'free');
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

  // Sheets list
  useEffect(() => {
    (async () => {
      const j = await apiGet('/excel/sheets');
      if (!j.error) setSheets(dedupeNames(j.sheets || []));
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

  // Real-time
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

  const filteredSheets = useMemo(() => {
    if (!searchQuery.trim()) return sheets;
    return sheets.filter((s) => s.toLowerCase().includes(searchQuery.toLowerCase()));
  }, [sheets, searchQuery]);

  async function previewSheet(name, opts = { readOnly: false }) {
    const j = await apiGet(`/excel/preview?sheet=${encodeURIComponent(name)}`);
    if (!j.error) {
      setSelectedSheet(name);
      setStagedGrid(j.preview || []);
      setForceReadOnly(Boolean(opts.readOnly));
    }
  }

  async function saveAllChanges() {
    if (!selectedSheet) {
      Toast.error('No sheet selected');
      return;
    }
    if (!Array.isArray(stagedGrid) || stagedGrid.length === 0) {
      Toast.warn('Nothing to save');
      return;
    }
    const j = await apiPost('/excel/save-all', { sheet: selectedSheet, data: stagedGrid });
    if (!j.error) {
      Toast.success(`Saved changes to ${selectedSheet}`);
      // Re-preview to confirm saved data persisted
      const j2 = await apiGet(`/excel/preview?sheet=${encodeURIComponent(selectedSheet)}`);
      if (!j2.error) setStagedGrid(j2.preview || stagedGrid);
      // Close editor
      setSelectedSheet(null);
      setForceReadOnly(false);
    }
  }

  async function deleteSheet(name) {
    if (role !== 'admin') {
      Toast.warn('Only admin can delete sheets');
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
    }
  }

  // Create unique names ("sheet", "sheet 1", "sheet 2", ...)
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

    if (name !== nameInput) {
      // The original exists; offer overwrite
      overwrite = window.confirm(`"${nameInput}" exists. Overwrite it? Press Cancel to save as "${name}".`);
    }

    const j = await apiPost('/excel/add-sheet', { name: overwrite ? nameInput : name, overwrite });
    if (!j.error) {
      const finalName = overwrite ? nameInput : name;
      Toast.success(`Sheet ${finalName} added`);
      const updated = j.sheets ? dedupeNames(j.sheets) : dedupeNames([...sheets, finalName]);
      setSheets(updated);
      setSelectedSheet(finalName);
      previewSheet(finalName);
    }
  }

  // Upload with dedupe and overwrite prompt
  async function handleUpload(file) {
    try {
      const formData = new FormData();
      formData.append('file', file);
      const headers = await authHeader();
      const res = await fetch(`${API_BASE}/excel/upload`, {
        method: 'POST',
        headers, // Authorization only; FormData will set its own content-type
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
            // Inform server to overwrite (optional API flag)
            await apiPost('/excel/overwrite', { name: nm });
          }
        } else {
          finalNames.push(nm);
          deduped.push(nm);
        }
      }

      setSheets(dedupeNames(deduped));

      const first = finalNames[0];
      if (first) {
        setSelectedSheet(first);
        setStagedGrid(j.preview || []);
      }

      Toast.success(`Uploaded ${j.fileName || finalNames.join(', ')}`);
    } catch (err) {
      Toast.error(err.message || 'Upload failed');
    }
  }

  const addRowAbove = () => {
    const hot = window.__hotInstance;
    if (!hot || !canEdit) return;
    const sel = hot.getSelectedLast();
    const rowIndex = sel ? sel[0] : 0;
    hot.alter('insert_row', rowIndex);
  };

  const addRowBelow = () => {
    const hot = window.__hotInstance;
    if (!hot || !canEdit) return;
    const sel = hot.getSelectedLast();
    const rowIndex = sel ? sel[0] : hot.countRows() - 1;
    hot.alter('insert_row', rowIndex + 1);
  };

  const addColLeft = () => {
    const hot = window.__hotInstance;
    if (!hot || !canEdit) return;
    const sel = hot.getSelectedLast();
    const colIndex = sel ? sel[1] : 0;
    hot.alter('insert_col', colIndex);
  };

  const addColRight = () => {
    const hot = window.__hotInstance;
    if (!hot || !canEdit) return;
    const sel = hot.getSelectedLast();
    const colIndex = sel ? sel[1] : hot.countCols() - 1;
    hot.alter('insert_col', colIndex + 1);
  };

  // Mobile swipe
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

  async function exportSheet() {
    if (!selectedSheet) {
      Toast.error('No sheet selected');
      return;
    }
    try {
      setLoading(true);
      const headers = await authHeader();
      const res = await fetch(`${API_BASE}/excel/export?sheet=${encodeURIComponent(selectedSheet)}`, {
        method: 'GET',
        headers,
      });
      if (!res.ok) {
        const txt = await res.text();
        Toast.error(txt || 'Export failed');
        return;
      }
      // Handle file download for CSV/PDF/etc.
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      // Try to infer filename from headers; fallback
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
      const res = await fetch(`${API_BASE}/excel/download?sheet=${encodeURIComponent(selectedSheet)}`, {
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
      a.download = `${selectedSheet}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      Toast.success(`Downloaded ${selectedSheet}`);
    } catch (err) {
      Toast.error(err.message);
    } finally {
      setLoading(false);
    }
  }

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

  // Upgrade flow (PayButton should call onSuccess -> refresh profile, hide ads, unlock features)
  async function onUpgradeSuccess() {
    Toast.success('Upgrade successful');
    await updateUser(user);
  }

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
              style={{ border: '1px solid #e5e7eb', borderRadius: 8, padding: '6px 8px' }}
            />
          </div>

          <div
            className="nav-actions"
            style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 8 }}
          >
            {user && plan === 'free' && (
              <PayButton
                user={user}
                amount={500000}
                currency="GHS"
                onSuccess={onUpgradeSuccess}
                onFailure={(e) => Toast.error(e?.message || 'Upgrade failed')}
              />
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
                zIndex: 1000,
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
                <button
                  className="secondary"
                  onClick={() => setActionsOpen(false)}
                  style={{ background: 'var(--danger, #fff)' }}
                >
                  <FaTimes /> Close
                </button>
              </div>
              <div className="dropdown-body" style={{ display: 'grid', gap: 8 }}>
                <InstallButton />
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
              }}
            >
              <button className="primary small" onClick={addSheet}>
                <FaPlus /> New sheet
              </button>
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

                    {role === 'admin' && !isMobile && hoveredItem === s && (
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

                    {role === 'admin' && isMobile && (
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
                margin: 12,
                padding: 12,
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

            {/* Ads: show only if free plan */}
            {shouldShowAds && (
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
              overflow: 'hidden',
              background: 'var(--bg, --text)',
              position: 'relative',
            }}
          >
            {selectedSheet ? (
              <div
                className="card editor-card"
                style={{
                  display: 'grid',
                  gridTemplateRows: 'auto 1fr auto',
                  height: '100%',
                  borderRadius: 12,
                  border: '1px solid #e5e7eb',
                  padding: 0,
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
                  }}
                >
                  <h2 style={{ margin: 0, fontSize: 18 }}>
                    {selectedSheet}
                    {forceReadOnly && <span style={{ marginLeft: 8, fontSize: 12, color: '#6b7280' }}>(read-only)</span>}
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
                    <button
                      className="secondary small"
                      onClick={exportSheet}
                      style={{ padding: '4px 8px' }}
                    >
                      <FaFileExport />
                    </button>
                    <button
                      className="secondary small"
                      onClick={downloadSheet}
                      style={{ padding: '4px 8px' }}
                    >
                      <FaDownload />
                    </button>
                    <button className="secondary" onClick={() => { setSelectedSheet(null); setForceReadOnly(false); }}>
                      <FaTimes />
                    </button>
                    <span className={`presence-indicator ${socketConnected ? 'online' : 'offline'}`}>
                      {socketConnected ? 'Live' : 'Offline'}
                    </span>
                  </div>
                </div>

                {/* Sticky builder actions removed per request (row/col controls not shown on editor) */}

                <div
                  className="editor-scroll-wrapper"
                  style={{
                    overflowX: 'auto',
                    overflowY: 'auto',
                    padding: 8,
                  }}
                >
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

                {canEdit && !forceReadOnly && (
                  <div
                    className="editor-footer-actions"
                    style={{
                      borderTop: '1px solid #e5e7eb',
                      padding: 12,
                      display: 'flex',
                      justifyContent: 'flex-end',
                    }}
                  >
                    <button className="primary" onClick={saveAllChanges} disabled={loading}>
                      <FaSave /> Save changes
                    </button>
                  </div>
                )}
              </div>
            ) : showAudit ? (
              <div
                className="card editor-card"
                style={{
                  height: '100%',
                  borderRadius: 12,
                  border: '1px solid #e5e7eb',
                  display: 'grid',
                  gridTemplateRows: 'auto 1fr',
                }}
              >
                <div
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    borderBottom: '1px solid #e5e7eb',
                    padding: '10px 16px',
                  }}
                >
                  <h2 style={{ margin: 0, fontSize: 18 }}>Audit Log</h2>
                  <button className="secondary" onClick={() => setShowAudit(false)}>
                    <FaTimes /> Close
                  </button>
                </div>
                <div style={{ overflowY: 'auto' }}>
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
              <div className="card" style={{ height: '100%' }}>
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

/* Handsontable-based editor with HyperFormula and real-time hooks */
function SheetEditor({ data, onChange, canEdit, onCellEdit }) {
  const containerRef = useRef(null);
  const hotRef = useRef(null);
  const hfRef = useRef(null);

  const [showFormulaPopup, setShowFormulaPopup] = useState(false);
  const [showRowPopup, setShowRowPopup] = useState(false);
  const [selectedFormula, setSelectedFormula] = useState('');
  const [formulaInfo, setFormulaInfo] = useState('');
  const [inputCols, setInputCols] = useState('');
  const [outputCol, setOutputCol] = useState('');
  const [stopRow, setStopRow] = useState('');
  const [rowsToAdd, setRowsToAdd] = useState(1);

  const formulaDescriptions = {
    SUM: 'Adds up all selected numbers.',
    AVERAGE: 'Calculates the average of selected numbers.',
    MULTIPLY: 'Multiplies two selected cells.',
    DIVIDE: 'Divides one cell by another.',
  };

  useEffect(() => {
    hfRef.current = HyperFormula.buildEmpty({
      licenseKey: 'gpl-v3',
    });

    const container = containerRef.current;

    hotRef.current = new Handsontable(container, {
      data,
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
      stretchH: 'all',
      // Let parent wrapper handle scrolling; we set max size for natural width/height feel
      width: container?.parentElement ? container.parentElement.clientWidth - 8 : '100%',
      height: container?.parentElement ? container.parentElement.clientHeight - 8 : 500,
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

    const ro = new ResizeObserver(() => {
      if (!hotRef.current || !container.parentElement) return;
      const { clientWidth, clientHeight } = container.parentElement;
      const newWidth = clientWidth - 8;
      const newHeight = clientHeight - 8;
      const settings = hotRef.current.getSettings();
      if (settings.width !== newWidth || settings.height !== newHeight) {
        hotRef.current.updateSettings({
          width: newWidth,
          height: newHeight,
        });
        hotRef.current.render();
      }
    });

    if (container.parentElement) ro.observe(container.parentElement);

    return () => {
      ro.disconnect();
      hotRef.current?.destroy();
      hfRef.current?.destroy();
      if (window.__hotInstance === hotRef.current) {
        window.__hotInstance = null;
      }
    };
  }, [canEdit]);

  useEffect(() => {
    if (hotRef.current) {
      hotRef.current.loadData(data || []);
    }
  }, [data]);

  // Formula builder
  const applyFormula = () => {
    try {
      if (!hotRef.current || !selectedFormula || !inputCols || !outputCol) return;

      const isRange = /^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/i.test(inputCols.trim());
      const output = outputCol.trim().toUpperCase();

      // Helper: parse "D7" -> { colIndex, rowIndex }
      const parseCellRef = (ref) => {
        const m = ref.match(/^([A-Z]+)([0-9]+)$/i);
        if (!m) return null;
        const colLetters = m[1].toUpperCase();
        const rowNumber = parseInt(m[2], 10); // spreadsheet 1-based
        return {
          colIndex: Handsontable.helper.spreadsheetColumnIndex(colLetters),
          rowIndex: rowNumber - 1, // convert to 0-based for HOT
        };
      };

      if (isRange) {
        // Grand total mode: inputCols like "D1:D6", outputCol like "D7"
        const out = parseCellRef(output);
        if (!out) return;

        const range = inputCols.trim().toUpperCase(); // e.g., "D1:D6"
        let formula = '';
        switch (selectedFormula) {
          case 'SUM':
            formula = `=SUM(${range})`;
            break;
          case 'AVERAGE':
            formula = `=AVERAGE(${range})`;
            break;
          default:
            formula = `=SUM(${range})`;
        }

        hotRef.current.setDataAtCell(out.rowIndex, out.colIndex, formula);
        return;
      }

      // Per-row builder mode: inputs like "B,C", output like "D", optional stopRow
      const inputs = inputCols.split(',').map(c => c.trim().toUpperCase()).filter(Boolean);
      const outputColLetter = output; // e.g., "D"
      const outColIndex = Handsontable.helper.spreadsheetColumnIndex(outputColLetter);
      const totalRows = hotRef.current.countRows();
      const stop = stopRow ? Math.min(parseInt(stopRow, 10), totalRows) : totalRows;

      for (let r = 0; r < stop; r++) {
        // r is 0-based; spreadsheet row number is r+1
        let formula = '';
        switch (selectedFormula) {
          case 'SUM':
            // SUM of cells in the same row across inputs (e.g., B1+C1)
            formula = `=SUM(${inputs.map(c => `${c}${r + 1}`).join(',')})`;
            break;
          case 'AVERAGE':
            formula = `=AVERAGE(${inputs.map(c => `${c}${r + 1}`).join(',')})`;
            break;
          case 'MULTIPLY':
            if (inputs.length >= 2) formula = `=${inputs[0]}${r + 1}*${inputs[1]}${r + 1}`;
            break;
          case 'DIVIDE':
            if (inputs.length >= 2) formula = `=${inputs[0]}${r + 1}/${inputs[1]}${r + 1}`;
            break;
        }
        if (formula) hotRef.current.setDataAtCell(r, outColIndex, formula);
      }

      // Add a total formula one row after stop (guard if within bounds)
      if (stop > 0) {
        const totalRowIndex = Math.min(stop, totalRows - 1);
        hotRef.current.setDataAtCell(
          totalRowIndex,
          outColIndex,
          `=SUM(${outputColLetter}1:${outputColLetter}${stop})`
        );
      }
    } finally {
      setShowFormulaPopup(false);
    }
  };

  const addRows = () => {
    try {
      if (!hotRef.current) return;
      const lastRow = hotRef.current.countRows() - 1;
      hotRef.current.alter('insert_row', lastRow >= 0 ? lastRow : 0, rowsToAdd);
    } finally {
      setShowRowPopup(false);
    }
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8, position: 'relative' }}>
      {/* Keep builder controls sticky and not moving with table */}
      {canEdit && (
        <div
          className="toolbar"
          style={{
            display: 'flex',
            gap: 8,
            padding: '6px 12px',
            borderBottom: '1px solid #e5e7eb',
            background: '#f9fafb',
            position: 'sticky',
            top: 0,
            zIndex: 200,
          }}
        >
          <button className="secondary small" onClick={() => setShowFormulaPopup(true)}>
            Formula Builder
          </button>
          <button
            className="secondary small"
            onClick={() => setShowRowPopup(true)}
            style={{ fontSize: 12, padding: '2px 6px' }}
          >
            + Rows
          </button>
        </div>
      )}

      {(showFormulaPopup || showRowPopup) && (
        <div
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            width: '100%',
            height: '100%',
            background: 'rgba(0,0,0,0.3)',
            zIndex: 999,
          }}
          onClick={() => {
            setShowFormulaPopup(false);
            setShowRowPopup(false);
          }}
        />
      )}

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
          <h3 style={{ marginTop: 0 }}>Formula Builder</h3>
          <label>
            Formula:
            <select
              value={selectedFormula}
              onChange={(e) => {
                setSelectedFormula(e.target.value);
                setFormulaInfo(formulaDescriptions[e.target.value] || '');
              }}
              style={{ marginLeft: 8 }}
            >
              <option value="">-- Select --</option>
              <option value="SUM">SUM</option>
              <option value="AVERAGE">AVERAGE</option>
              <option value="MULTIPLY">MULTIPLY</option>
              <option value="DIVIDE">DIVIDE</option>
            </select>
          </label>
          {formulaInfo && <p style={{ fontSize: 12, color: '#6b7280' }}>{formulaInfo}</p>}

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

          <div style={{ marginTop: 12, display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
            <button className="secondary small" onClick={() => setShowFormulaPopup(false)}>Cancel</button>
            <button className="primary small" onClick={applyFormula}>Apply</button>
          </div>
        </div>
      )}

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

      <div
        className="hot-container"
        ref={containerRef}
        style={{
          minWidth: 900,
          minHeight: 500,
          background: 'var(--bg, --text)',
          zIndex: 1,
        }}
      />
    </div>
  );
}

/* Read-only audit list that opens a sheet in read-only mode */
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

function formatTS(ts) {
  try {
    const d = new Date(ts);
    return d.toLocaleString();
  } catch {
    return '—';
  }
}


/* Handsontable-based editor with HyperFormula and real-time hooks */
// function SheetEditor({ data, onChange, canEdit, onCellEdit }) {
//   const containerRef = useRef(null);
//   const hotRef = useRef(null);
//   const hfRef = useRef(null);

//   const [showFormulaPopup, setShowFormulaPopup] = useState(false);
//   const [showRowPopup, setShowRowPopup] = useState(false);
//   const [selectedFormula, setSelectedFormula] = useState('');
//   const [formulaInfo, setFormulaInfo] = useState('');
//   const [inputCols, setInputCols] = useState('');
//   const [outputCol, setOutputCol] = useState('');
//   const [stopRow, setStopRow] = useState('');
//   const [rowsToAdd, setRowsToAdd] = useState(1);

//   const formulaDescriptions = {
//     SUM: 'Adds up all selected numbers.',
//     AVERAGE: 'Calculates the average of selected numbers.',
//     MULTIPLY: 'Multiplies two selected cells.',
//     DIVIDE: 'Divides one cell by another.',
//   };

//   useEffect(() => {
//     hfRef.current = HyperFormula.buildEmpty({
//       licenseKey: 'gpl-v3',
//     });

//     const container = containerRef.current;

//     hotRef.current = new Handsontable(container, {
//       data: Array.isArray(data) ? data : [[]],
//       licenseKey: 'non-commercial-and-evaluation',
//       rowHeaders: true,
//       colHeaders: true,
//       formulas: { engine: hfRef.current },
//       readOnly: !canEdit,
//       autoFill: true,
//       contextMenu: canEdit ? true : false,
//       dropdownMenu: true,
//       manualRowMove: true,
//       manualColumnMove: true,
//       manualColumnResize: true,
//       stretchH: 'all',
//       // Fit into scroll wrapper
//       width: container?.parentElement ? container.parentElement.clientWidth - 8 : '100%',
//       height: container?.parentElement ? container.parentElement.clientHeight - 8 : 500,
//       afterChange: (changes, source) => {
//         if (!changes || source === 'loadData') return;
//         const next = hotRef.current.getData();
//         onChange(next);
//         changes.forEach(([row, col, _oldVal, newVal]) => {
//           onCellEdit?.({ rowIndex: row, colIndex: col, value: newVal });
//         });
//       },
//     });

//     window.__hotInstance = hotRef.current;

//     return () => {
//       try {
//         hotRef.current?.destroy();
//         window.__hotInstance = null;
//       } catch {}
//     };
//     // eslint-disable-next-line react-hooks/exhaustive-deps
//   }, []);

//   useEffect(() => {
//     if (!hotRef.current) return;
//     hotRef.current.updateSettings({
//       data: Array.isArray(data) ? data : [[]],
//       readOnly: !canEdit,
//     });
//   }, [data, canEdit]);

//   const applyFormula = () => {
//     const hot = hotRef.current;
//     if (!hot) return;
//     const inputs = inputCols.split(',').map((c) => c.trim().toUpperCase()).filter(Boolean);
//     const out = (outputCol || '').trim().toUpperCase();
//     const stop = parseInt(stopRow, 10) || hot.countRows();

//     if (!selectedFormula || !out) {
//       Toast.warn('Select a formula and output column');
//       return;
//     }

//     const colIndex = (letter) => {
//       const A = 'A'.charCodeAt(0);
//       return letter ? letter.charCodeAt(0) - A : 0;
//     };

//     for (let r = 0; r < stop; r++) {
//       let formula = '';
//       if (selectedFormula === 'SUM') {
//         formula = `=SUM(${inputs.map((i) => `${i}${r + 1}`).join(',')})`;
//       } else if (selectedFormula === 'AVERAGE') {
//         formula = `=AVERAGE(${inputs.map((i) => `${i}${r + 1}`).join(',')})`;
//       } else if (selectedFormula === 'MULTIPLY') {
//         if (inputs.length < 2) continue;
//         formula = `=${inputs[0]}${r + 1}*${inputs[1]}${r + 1}`;
//       } else if (selectedFormula === 'DIVIDE') {
//         if (inputs.length < 2) continue;
//         formula = `=${inputs[0]}${r + 1}/${inputs[1]}${r + 1}`;
//       }
//       hot.setDataAtCell(r, colIndex(out), formula);
//     }

//     setShowFormulaPopup(false);
//     Toast.success('Formula applied');
//   };

//   const addRows = () => {
//     const hot = hotRef.current;
//     if (!hot) return;
//     const count = parseInt(rowsToAdd, 10) || 1;
//     hot.alter('insert_row', hot.countRows(), count);
//     setShowRowPopup(false);
//   };

//   useEffect(() => {
//     const info = selectedFormula ? formulaDescriptions[selectedFormula] : '';
//     setFormulaInfo(info || '');
//   }, [selectedFormula]);

//   return (
//     <div
//       className="sheet-editor-wrapper"
//       style={{
//         position: 'relative',
//         minWidth: 900,
//         minHeight: 500,
//       }}
//     >
//       {/* Sticky toolbar with formula & row controls */}
//       <div
//         className="sticky-toolbar"
//         style={{
//           position: 'sticky',
//           top: 0,
//           display: 'flex',
//           alignItems: 'center',
//           gap: 8,
//           padding: 8,
//           background: 'var(--bg, --text)',
//           zIndex: 5,
//           borderBottom: '1px solid #e5e7eb',
//         }}
//       >
//         <button className="secondary small" onClick={() => setShowFormulaPopup(true)}>
//           Formula
//         </button>
//         <button className="secondary small" onClick={() => setShowRowPopup(true)}>
//           +Row
//         </button>
//       </div>

//       {/* Formula modal */}
//       {showFormulaPopup && (
//         <div
//           style={{
//             position: 'fixed',
//             top: '20%',
//             left: '50%',
//             transform: 'translateX(-50%)',
//             background: 'var(--bg, --text)',
//             border: '1px solid #e5e7eb',
//             borderRadius: 8,
//             padding: 16,
//             zIndex: 1000,
//             boxShadow: '0 8px 20px rgba(0,0,0,0.25)',
//             width: 320,
//           }}
//         >
//           <h3 style={{ marginTop: 0 }}>Apply Formula</h3>
//           <label style={{ display: 'block', marginBottom: 8 }}>
//             Choose formula:
//             <select
//               value={selectedFormula}
//               onChange={(e) => setSelectedFormula(e.target.value)}
//               style={{ marginLeft: 8 }}
//             >
//               <option value="">-- Select --</option>
//               <option value="SUM">SUM</option>
//               <option value="AVERAGE">AVERAGE</option>
//               <option value="MULTIPLY">MULTIPLY</option>
//               <option value="DIVIDE">DIVIDE</option>
//             </select>
//           </label>
//           {formulaInfo && <p style={{ fontSize: 12, color: '#6b7280' }}>{formulaInfo}</p>}

//           <label>
//             Input cols:
//             <input
//               value={inputCols}
//               onChange={(e) => setInputCols(e.target.value)}
//               style={{ width: 50, marginLeft: 8 }}
//               placeholder="B,C"
//             />
//           </label>
//           <label>
//             Output col:
//             <input
//               value={outputCol}
//               onChange={(e) => setOutputCol(e.target.value)}
//               style={{ width: 50, marginLeft: 8 }}
//               placeholder="D"
//             />
//           </label>
//           <label>
//             Stop row:
//             <input
//               value={stopRow}
//               onChange={(e) => setStopRow(e.target.value)}
//               style={{ width: 50, marginLeft: 8 }}
//               placeholder="20"
//             />
//           </label>

//           <div style={{ marginTop: 12, display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
//             <button className="secondary small" onClick={() => setShowFormulaPopup(false)}>Cancel</button>
//             <button className="primary small" onClick={applyFormula}>Apply</button>
//           </div>
//         </div>
//       )}

//       {/* Add rows modal */}
//       {showRowPopup && (
//         <div
//           style={{
//             position: 'fixed',
//             top: '30%',
//             left: '50%',
//             transform: 'translateX(-50%)',
//             background: 'var(--bg, --text)',
//             border: '1px solid #e5e7eb',
//             borderRadius: 8,
//             padding: 16,
//             zIndex: 1000,
//             boxShadow: '0 8px 20px rgba(0,0,0,0.25)',
//             width: 220,
//           }}
//         >
//           <h3 style={{ marginTop: 0 }}>Add Rows</h3>
//           <label>
//             Number of rows:
//             <input
//               type="number"
//               value={rowsToAdd}
//               onChange={(e) => setRowsToAdd(parseInt(e.target.value, 10))}
//               style={{ width: 50, marginLeft: 8 }}
//             />
//           </label>
//           <div style={{ marginTop: 12, display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
//             <button className="secondary small" onClick={() => setShowRowPopup(false)}>Cancel</button>
//             <button className="primary small" onClick={addRows}>Add</button>
//           </div>
//         </div>
//       )}

//       <div
//         className="hot-container"
//         ref={containerRef}
//         style={{
//           minWidth: 900,
//           minHeight: 500,
//           background: 'var(--bg, --text)',
//           zIndex: 1,
//         }}
//       />
//     </div>
//   );
// }