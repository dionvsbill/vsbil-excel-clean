import { Router } from 'express';
import { supabase, CONFIG } from '../app.js';
import ExcelJS from 'exceljs';
import { nanoid } from 'nanoid';

// Resolve identity and role from Supabase JWT
async function getIdentity(req) {
  const token = (req.headers.authorization || '').replace('Bearer ', '');
  if (!token) return { user: null, role: 'anonymous', can_edit: false };

  const { data, error } = await supabase.auth.getUser(token);
  if (error || !data?.user) return { user: null, role: 'anonymous', can_edit: false };

  const md = data.user.user_metadata || {};
  const role = md.role === 'admin' ? 'admin' : 'user';
  const can_edit = role === 'admin' || Boolean(md.can_edit);

  return { user: data.user, role, can_edit };
}

// Append audit log entry to Supabase Storage (NDJSON, one file per day)
async function writeAuditLog(entry) {
  try {
    const day = new Date().toISOString().slice(0, 10);
    const objectKey = `${CONFIG.LOGS_PREFIX}/${day}.ndjson`;

    let current = '';
    const existing = await supabase.storage.from(CONFIG.LOGS_BUCKET).download(objectKey);
    if (existing.data) {
      const buf = Buffer.from(await existing.data.arrayBuffer());
      current = buf.toString('utf-8');
    }

    const updated = current + JSON.stringify(entry) + '\n';

    const up = await supabase.storage.from(CONFIG.LOGS_BUCKET).upload(objectKey, updated, {
      upsert: true,
      contentType: 'application/x-ndjson'
    });

    if (up.error) {
      console.warn('Audit log upload error:', up.error.message);
      return false;
    }
    return true;
  } catch (e) {
    console.warn('Audit log write failed:', e.message);
    return false;
  }
}

// Helpers
function clientIp(req) {
  return req.headers['x-forwarded-for']?.split(',')[0]?.trim() || req.ip || 'unknown';
}

const router = Router();
const FILE_KEY = CONFIG.EXCEL_FILE_KEY;

// Public view (signed URL, short TTL)
router.get('/public', async (req, res) => {
  const { data, error } = await supabase.storage
    .from(CONFIG.EXCEL_BUCKET)
    .createSignedUrl(FILE_KEY, 60);

  if (error) return res.status(500).json({ error: error.message });

  await writeAuditLog({
    id: nanoid(),
    ts: new Date().toISOString(),
    action: 'view_public',
    file_key: FILE_KEY,
    bucket: CONFIG.EXCEL_BUCKET,
    ip: clientIp(req),
    ua: req.headers['user-agent'] || null,
    user_id: null,
    email: null
  });

  res.json({ url: data.signedUrl, expires_in: 60 });
});

// Authenticated download
router.get('/download', async (req, res) => {
  const { user } = await getIdentity(req);
  if (!user) return res.status(401).json({ error: 'Unauthorized' });

  const { data, error } = await supabase.storage
    .from(CONFIG.EXCEL_BUCKET)
    .createSignedUrl(FILE_KEY, 60);

  if (error) return res.status(500).json({ error: error.message });

  await writeAuditLog({
    id: nanoid(),
    ts: new Date().toISOString(),
    action: 'download',
    file_key: FILE_KEY,
    bucket: CONFIG.EXCEL_BUCKET,
    ip: clientIp(req),
    ua: req.headers['user-agent'] || null,
    user_id: user.id,
    email: user.email || null
  });

  res.json({ url: data.signedUrl, expires_in: 60 });
});

// Admin/editor update: apply cell changes with ExcelJS
router.post('/update', async (req, res) => {
  const { user, can_edit } = await getIdentity(req);
  if (!user) return res.status(401).json({ error: 'Unauthorized' });
  if (!can_edit) return res.status(403).json({ error: 'Forbidden' });

  const { changes = [] } = req.body;
  if (!Array.isArray(changes) || changes.length === 0) {
    return res.status(400).json({ error: 'No changes provided' });
  }

  // Download current workbook
  const dl = await supabase.storage.from(CONFIG.EXCEL_BUCKET).download(FILE_KEY);
  if (dl.error) return res.status(500).json({ error: `Download failed: ${dl.error.message}` });

  const buf = Buffer.from(await dl.data.arrayBuffer());
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);

  const diffSummary = [];
  try {
    for (const c of changes) {
      const { sheet: sheetName, cell: cellAddr, value: newValue, formula, style } = c;
      if (!sheetName || !cellAddr) throw new Error('Invalid change payload');

      const sheet = workbook.getWorksheet(sheetName);
      if (!sheet) throw new Error(`Sheet not found: ${sheetName}`);

      const cell = sheet.getCell(cellAddr);
      const before = cell.value;

      if (formula) {
        cell.value = { formula, result: newValue ?? null };
      } else {
        cell.value = newValue;
      }

      if (style && typeof style === 'object') {
        Object.assign(cell.style, style);
      }

      diffSummary.push({ sheet: sheetName, cell: cellAddr, before, after: cell.value });
    }

    const outBuffer = await workbook.xlsx.writeBuffer();

    const up = await supabase.storage.from(CONFIG.EXCEL_BUCKET).upload(FILE_KEY, outBuffer, {
      upsert: true,
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    if (up.error) return res.status(500).json({ error: `Upload failed: ${up.error.message}` });

    await writeAuditLog({
      id: nanoid(),
      ts: new Date().toISOString(),
      action: 'write',
      file_key: FILE_KEY,
      bucket: CONFIG.EXCEL_BUCKET,
      ip: clientIp(req),
      ua: req.headers['user-agent'] || null,
      user_id: user.id,
      email: user.email || null,
      diff_summary: diffSummary
    });

    res.json({ status: 'ok', changes_applied: diffSummary.length });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

export default router;
