// src/components/Sheets.jsx
import { useEffect, useState } from 'react';
import { apiGet, apiPost, apiDownload, getAuthHeaders, API_BASE } from '../api/client';

export default function Sheets() {
  const [sheets, setSheets] = useState([]);
  const [selected, setSelected] = useState('');
  const [preview, setPreview] = useState([]);
  const [newSheet, setNewSheet] = useState('');

  const loadSheets = async () => {
    try {
      const res = await apiGet('/excel/sheets', {
        headers: getAuthHeaders({ 'x-ads-watched': '2' }),
      });
      setSheets(res.sheets || []);
    } catch (e) {
      alert(e.message);
    }
  };

  const loadPreview = async (sheet) => {
    try {
      const res = await apiGet(`/excel/preview?sheet=${encodeURIComponent(sheet)}`);
      setPreview(res.preview || []);
    } catch (e) {
      alert(e.message);
    }
  };

  useEffect(() => { loadSheets(); }, []);

  useEffect(() => {
    if (selected) loadPreview(selected);
  }, [selected]);

  const addSheet = async () => {
    try {
      await apiPost('/excel/add-sheet', { name: newSheet, overwrite: false }, {
        headers: getAuthHeaders({ 'x-ads-watched': '2' }),
      });
      setNewSheet('');
      await loadSheets();
    } catch (e) {
      alert(e.message);
    }
  };

  const deleteSheet = async () => {
    try {
      await apiPost('/excel/delete-sheet', { name: selected }, {
        headers: getAuthHeaders({ 'x-ads-watched': '2' }),
      });
      setSelected('');
      setPreview([]);
      await loadSheets();
    } catch (e) {
      alert(e.message);
    }
  };

  const exportCSV = async () => {
    try {
      const blob = await apiDownload(`/excel/export/csv?sheet=${encodeURIComponent(selected)}`, {
        headers: getAuthHeaders({ 'x-ads-watched': '2' }),
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = `${selected}.csv`; a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      alert(e.message);
    }
  };

  const exportPDF = async () => {
    try {
      const res = await fetch(`${API_BASE}/excel/export/pdf?sheet=${encodeURIComponent(selected)}`, {
        method: 'GET',
        headers: getAuthHeaders({ 'x-ads-watched': '2' }),
      });
      if (!res.ok) throw new Error(await res.text());
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = `${selected}.pdf`; a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      alert(e.message);
    }
  };

  return (
    <div className="grid">
      <section>
        <h2>Sheets</h2>
        <div className="row">
          <select value={selected} onChange={(e) => setSelected(e.target.value)}>
            <option value="">Select sheet</option>
            {sheets.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
          <button onClick={loadSheets}>Reload</button>
          <input placeholder="New sheet name" value={newSheet} onChange={(e) => setNewSheet(e.target.value)} />
          <button onClick={addSheet}>Add</button>
          <button onClick={deleteSheet} disabled={!selected}>Delete</button>
        </div>
        <div className="row">
          <button onClick={exportCSV} disabled={!selected}>Export CSV</button>
          <button onClick={exportPDF} disabled={!selected}>Export PDF</button>
        </div>
      </section>
      <section>
        <h3>Preview</h3>
        <table className="table">
          <tbody>
          {preview.map((row, idx) => (
            <tr key={idx}>
              {row.map((cell, i) => <td key={i}>{cell === null ? '' : String(cell)}</td>)}
            </tr>
          ))}
          </tbody>
        </table>
      </section>
    </div>
  );
}
