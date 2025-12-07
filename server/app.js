// server/app.js
// Premium, production-grade backend: Excel workflows, role-based admin/owner controls,
// payments (Paystack), support sessions, analytics, pricing controls, legal pages,
// real-time event streaming (SSE), robust auditing, and plan-aware gating.
// ES Modules compatible with CommonJS libraries via createRequire.

import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import rateLimit from 'express-rate-limit';
import { createClient } from '@supabase/supabase-js';
import ExcelJS from 'exceljs';
import multer from 'multer';
import { parse as csvParse } from 'csv-parse/sync';
import crypto from 'crypto';
import fetch from 'node-fetch';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);
const pdfParse = require('pdf-parse'); // CommonJS import fix

// Policies are assumed existing. If not, inline checks below cover the core gating.
// import { attachUserContext, requirePremium, requireAdsSoft } from './policies.js';

// Minimal safe stubs to avoid missing imports; integrate your existing policies if present.
const attachUserContext = async (req, _res, next) => {
  req.context = { ip: req.ip, ua: req.headers['user-agent'] || '' };
  next();
};
const requirePremium = async (req, res, next) => {
  if (
    req.userPlan !== 'paid' &&
    req.userRole !== 'superadmin' &&
    req.userEmail !== CONFIG.OWNER_EMAIL
  ) {
    return res.status(403).json({ error: 'Requires premium plan' });
  }
  next();
};
const requireAdsSoft = async (_req, _res, next) => next();

const app = express();

/* -------------------------------------------------------
   Security + middleware
------------------------------------------------------- */
app.use(helmet({ crossOriginResourcePolicy: false }));
app.use(
  cors({
    origin: (_origin, cb) => cb(null, true),
    credentials: true,
  })
);
app.use(express.json({ limit: '15mb' })); // bump default limit to 15mb

const limiter = rateLimit({
  windowMs: 60 * 1000,
  max: 600, // higher burst, still protective
});
app.use(limiter);

const REQUIRED_ENV = [
  'SUPABASE_URL',
  'SUPABASE_ANON_KEY',
  'SUPABASE_SERVICE_ROLE_KEY',
  'PAYSTACK_SECRET_KEY',
  'OWNER_EMAIL',
];
for (const k of REQUIRED_ENV) {
  if (!process.env[k]) {
    throw new Error(`Missing env: ${k} must be set in .env`);
  }
}

/* -------------------------------------------------------
   Supabase clients
------------------------------------------------------- */
// Public client (anon key) for user-scoped queries
const baseSupabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_ANON_KEY, {
  auth: { persistSession: false, autoRefreshToken: false },
});

// Service role client for privileged operations (bypasses RLS)
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY, {
  auth: { persistSession: false, autoRefreshToken: false },
});

// Token-scoped client for authenticated users
export const getSupabaseForToken = (token) =>
  createClient(process.env.SUPABASE_URL, process.env.SUPABASE_ANON_KEY, {
    auth: { persistSession: false, autoRefreshToken: false },
    global: { headers: { Authorization: `Bearer ${token}` } },
  });

/* -------------------------------------------------------
   Config
------------------------------------------------------- */
export const CONFIG = {
  EXCEL_BUCKET: process.env.EXCEL_BUCKET || 'excel',
  EXCEL_FILE_KEY: process.env.EXCEL_FILE_KEY || 'master.xlsx',
  USER_FILES_PREFIX: process.env.USER_FILES_PREFIX || 'users',
  LOGS_BUCKET: process.env.LOGS_BUCKET || 'logs',
  LOGS_PREFIX: process.env.LOGS_PREFIX || 'excel_access',
  ADS_REQUIRED: 2,
  FREE_LIMITS: {
    maxSheetsInMultiPDF: 1,
    maxRowsSaveAll: 5000,
  },
  DAILY_EXPORT_LIMIT_FREE: 3,
  OWNER_EMAIL: process.env.OWNER_EMAIL,
  SUPPORT_SESSION_TTL_MIN: parseInt(process.env.SUPPORT_SESSION_TTL_MIN || '60', 10),
  REALTIME_PING_INTERVAL_MS: 25000,
  MAX_SSE_CLIENTS: 500,
};

/* -------------------------------------------------------
   Helpers
------------------------------------------------------- */
const getBufferFromStorage = async (supabaseClient, bucket, key) => {
  try {
    const { data, error } = await supabaseClient.storage.from(bucket).download(key);
    if (error) {
      return { error };
    }
    const buffer = Buffer.from(await data.arrayBuffer());
    return { buffer };
  } catch (e) {
    return { error: e };
  }
};

const putBufferToStorage = async (supabaseClient, bucket, key, buffer, contentType) => {
  try {
    const { error } = await supabaseClient.storage.from(bucket).upload(key, buffer, {
      upsert: true,
      contentType,
    });
    return { error };
  } catch (e) {
    return { error: e };
  }
};

const updateBufferToStorage = async (supabaseClient, bucket, key, buffer, contentType) => {
  try {
    const { error } = await supabaseClient.storage.from(bucket).update(key, buffer, {
      contentType,
    });
  return { error };
  } catch (e) {
    return { error: e };
  }
};

const removeFromStorage = async (supabaseClient, bucket, key) => {
  try {
    const { error } = await supabaseClient.storage.from(bucket).remove([key]);
    return { error };
  } catch (e) {
    return { error: e };
  }
};

const loadWorkbook = async (buffer) => {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    return workbook;
  } catch (e) {
    throw e;
  }
};

const sanitizeString = (s) => {
  if (s === null || s === undefined) return '';
  return String(s);
};

/* -------------------------------------------------------
   Audit helper (NDJSON logs in storage)
------------------------------------------------------- */
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
      contentType: 'application/x-ndjson',
    });

    if (up.error) {
      return false;
    }
    return true;
  } catch (e) {
    return false;
  }
}

/* -------------------------------------------------------
   Ads helper (soft gate verification)
------------------------------------------------------- */
async function checkAdsWatched(userId, required = CONFIG.ADS_REQUIRED) {
  try {
    const { data: events, error } = await supabase
      .from('ads_events')
      .select('id')
      .eq('user_id', userId)
      .eq('event', 'watched');
    if (error) {
      return false;
    }
    const cnt = Array.isArray(events) ? events.length : 0;
    return cnt >= required;
  } catch (e) {
    return false;
  }
}

/* -------------------------------------------------------
   Auth middleware
------------------------------------------------------- */
const requireAuth = async (req, res, next) => {
  try {
    const authHeader = req.headers.authorization || '';
    const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;
    if (!token) return res.status(401).json({ error: 'Unauthorized: Login to get access' });

    const authClient = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_ANON_KEY, {
      auth: { persistSession: false, autoRefreshToken: false },
    });
    const { data, error } = await authClient.auth.getUser(token);
    if (error || !data?.user) return res.status(401).json({ error: 'Unauthorized: invalid User' });

    req.user = data.user;
    req.supabase = getSupabaseForToken(token);
    next();
  } catch (e) {
    res.status(401).json({ error: 'Unauthorized' });
  }
};

const attachUserPlanAndFile = async (req, res, next) => {
  try {
    if (!req.user || !req.supabase) {
      req.userPlan = 'anon';
      req.userRole = 'guest';
      req.userEmail = '';
      req.fileKey = CONFIG.EXCEL_FILE_KEY;
      return next();
    }
    const { data, error } = await req.supabase
      .from('profiles')
      .select('plan,user_file_key,role,email,status')
      .eq('id', req.user.id)
      .single();
    if (error) {
      req.userPlan = 'free';
      req.userRole = 'user';
      req.userEmail = req.user.email || '';
      req.fileKey = CONFIG.EXCEL_FILE_KEY;
      return next();
    }
    req.userPlan = data?.plan || 'free';
    req.userRole = data?.role || 'user';
    req.userEmail = data?.email || req.user.email || '';
    req.userStatus = data?.status || 'active';
    req.fileKey =
      data?.user_file_key ||
      `${CONFIG.USER_FILES_PREFIX}/${req.user.id}/uploaded.xlsx`;
    next();
  } catch (e) {
    req.userPlan = 'free';
    req.userRole = 'user';
    req.userEmail = req.user?.email || '';
    req.fileKey = CONFIG.EXCEL_FILE_KEY;
    next();
  }
};

const requireSuperadmin = async (req, res, next) => {
  try {
    const { data, error } = await req.supabase
      .from('profiles')
      .select('role')
      .eq('id', req.user.id)
      .single();
    if (error) return res.status(403).json({ error: 'Forbidden: role check failed' });
    if (data?.role !== 'superadmin') return res.status(403).json({ error: 'Forbidden: superadmin only' });
    next();
  } catch (e) {
    res.status(403).json({ error: 'Forbidden' });
  }
};

const requireOwner = async (req, res, next) => {
  try {
    const email = req.userEmail || req.user?.email;
    if (!CONFIG.OWNER_EMAIL || !email) return res.status(403).json({ error: 'Forbidden: owner email not set' });
    if (email !== CONFIG.OWNER_EMAIL) return res.status(403).json({ error: 'Forbidden: owner only' });
    next();
  } catch (e) {
    res.status(403).json({ error: 'Forbidden' });
  }
};

const requirePremiumPlan = async (req, res, next) => {
  if (
    req.userPlan !== 'paid' &&
    req.userRole !== 'superadmin' &&
    req.userEmail !== CONFIG.OWNER_EMAIL
  ) {
    return res.status(403).json({ error: 'Requires premium plan' });
  }
  next();
};

const requireAdsWatchedHeader = (req, res, next) => {
  const adsWatched = parseInt(req.headers['x-ads-watched'] || '0', 10);
  if (Number.isNaN(adsWatched) || adsWatched < CONFIG.ADS_REQUIRED) {
    return res.status(403).json({ error: `Please watch ${CONFIG.ADS_REQUIRED} ads before downloading/exporting` });
  }
  next();
};

/* -------------------------------------------------------
   Health + utility routes
------------------------------------------------------- */
app.get('/health', (req, res) => res.json({ ok: true, ts: new Date().toISOString() }));
app.get('/ping', (req, res) => res.json({ message: 'pong' }));
app.get('/', (req, res) => res.json({ message: 'Backend running' }));

/* -------------------------------------------------------
   Real-time (SSE) channel: events for dashboard
------------------------------------------------------- */
const sseClients = new Map(); // id -> res
function broadcastSSE(event, payload) {
  const msg = `event: ${event}\ndata: ${JSON.stringify(payload)}\n\n`;
  for (const [, res] of sseClients) {
    try {
      res.write(msg);
    } catch (e) {
      // client may be closed
    }
  }
}

app.get('/realtime/events', requireAuth, attachUserPlanAndFile, (req, res) => {
  if (req.userEmail !== CONFIG.OWNER_EMAIL && req.userRole !== 'superadmin') {
    return res.status(403).json({ error: 'Forbidden' });
  }
  if (sseClients.size >= CONFIG.MAX_SSE_CLIENTS) {
    return res.status(503).json({ error: 'Max clients reached' });
  }
  res.set({
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    Connection: 'keep-alive',
  });
  res.flushHeaders?.();

  const clientId = `${req.user.id}:${Date.now()}`;
  sseClients.set(clientId, res);
  res.write(`event: connected\ndata: ${JSON.stringify({ id: clientId, ts: Date.now() })}\n\n`);

  const ping = setInterval(() => {
    try {
      res.write(`event: ping\ndata: ${JSON.stringify({ ts: Date.now() })}\n\n`);
    } catch (e) {}
  }, CONFIG.REALTIME_PING_INTERVAL_MS);

  req.on('close', () => {
    clearInterval(ping);
    sseClients.delete(clientId);
  });
});

/* -------------------------------------------------------
   Excel public URL (premium-only)
------------------------------------------------------- */
app.get('/excel/public', requireAuth, attachUserPlanAndFile, async (req, res) => {
  try {
    if (req.userPlan === 'free' && req.userRole !== 'superadmin' && req.userEmail !== CONFIG.OWNER_EMAIL) {
      return res.status(403).json({ error: 'Public file access is available only on premium plans' });
    }

    const key = req.userRole === 'superadmin' ? CONFIG.EXCEL_FILE_KEY : req.fileKey;

    if (!req._publicUrlCache) {
      const { data, error } = baseSupabase.storage.from(CONFIG.EXCEL_BUCKET).getPublicUrl(key);
      if (error) return res.status(500).json({ error: error.message });
      if (!data?.publicUrl) return res.status(404).json({ error: 'File not found in bucket' });
      req._publicUrlCache = `${data.publicUrl}?t=${Date.now()}`;
    }

    res.json({ url: req._publicUrlCache });
  } catch (e) {
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});

/* -------------------------------------------------------
   Workbook download (premium-only)
------------------------------------------------------- */
app.get('/excel/download', requireAuth, attachUserContext, attachUserPlanAndFile, async (req, res) => {
  try {
    // Explicitly block only free users (unless superadmin or owner)
    if (req.userPlan === 'free' && req.userRole !== 'superadmin' && req.userEmail !== CONFIG.OWNER_EMAIL) {
      return res.status(403).json({ error: 'Downloads are available only on premium plans' });
    }

    const key = req.userRole === 'superadmin' ? CONFIG.EXCEL_FILE_KEY : req.fileKey;

    // Faster: cache buffer in memory for this request cycle
    if (!req._bufferCache) {
      const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
      if (error || !buffer) return res.status(404).json({ error: 'Workbook not found' });
      req._bufferCache = buffer;
    }

    const buffer = req._bufferCache;
    res.setHeader('Content-Disposition', `attachment; filename="${key.split('/').pop()}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Length', buffer.length);
    res.status(200).send(buffer);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* -------------------------------------------------------
Add sheet (free users: max 3/day, pin latest)
------------------------------------------------------- */
app.post('/excel/add-sheet', requireAuth, attachUserContext, attachUserPlanAndFile, async (req, res) => {
  const { name, overwrite } = req.body;
  if (!name) return res.status(400).json({ error: 'Sheet name is required' });

  try {
    // Check free plan limits
    if (req.userPlan === 'free') {
      const { count } = await req.supabase
        .from('excel_audit')
        .select('*', { count: 'exact', head: true })
        .eq('user_id', req.user.id)
        .eq('action', 'add_sheet')
        .gte('created_at', new Date().toISOString().split('T')[0]);

      if (count >= 3) {
        return res.status(403).json({ error: 'Free plan limit: max 3 new sheets per day' });
      }
    }

    const key = req.userRole === 'superadmin' ? CONFIG.EXCEL_FILE_KEY : req.fileKey;
    if (!key) return res.status(404).json({ error: 'Workbook key not found' });

    let workbook;
    const { buffer, error } = await getBufferFromStorage(req.supabase, CONFIG.EXCEL_BUCKET, key);
    if (error || !buffer) {
      workbook = new ExcelJS.Workbook();
      workbook.addWorksheet('Sheet1');
    } else {
      workbook = await loadWorkbook(buffer);
    }

    const existing = workbook.getWorksheet(name);
    if (existing) {
      if (!overwrite) return res.status(400).json({ error: 'Sheet already exists' });
      workbook.removeWorksheet(existing.id);
    }

    const ws = workbook.addWorksheet(name);
    ws.getCell('A1').value = 'New sheet created';

    const outBuffer = await workbook.xlsx.writeBuffer();
    const { error: uploadError } = await updateBufferToStorage(
      req.supabase,
      CONFIG.EXCEL_BUCKET,
      key,
      outBuffer,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );

    if (uploadError) {
      return res.status(500).json({ error: uploadError.message });
    }

    const reloaded = await loadWorkbook(outBuffer);
    let sheetNames = reloaded.worksheets.map(ws => ws.name);

    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      email: req.userEmail,
      action: 'add_sheet',
      sheet_name: name,
      metadata: { role: req.userRole, plan: req.userPlan }
    });

    broadcastSSE('excel:add_sheet', { by: req.userEmail, sheet: name, key });

    // Reorder sheets: new sheet at top, rest shuffled
    const rest = sheetNames.filter(s => s !== name);
    for (let i = rest.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [rest[i], rest[j]] = [rest[j], rest[i]];
    }
    sheetNames = [name, ...rest];

    res.json({ success: true, sheet: name, sheets: sheetNames });
  } catch (e) {
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});


/* -------------------------------------------------------
   Delete sheet (free users limited by daily edits/saves)
------------------------------------------------------- */
app.post(
  '/excel/delete-sheet',
  requireAuth,
  attachUserContext,
  attachUserPlanAndFile,
  async (req, res) => {
    const { name } = req.body;
    if (!name) return res.status(400).json({ error: 'Sheet name required' });

    try {
      // Free plan limit: max 3 edits/saves/deletes per day
      if (req.userPlan === 'free') {
        const { count } = await req.supabase
          .from('excel_audit')
          .select('*', { count: 'exact', head: true })
          .eq('user_id', req.user.id)
          .in('action', ['save_all', 'delete_sheet', 'edit_sheet', 'add_sheet'])
          .gte('created_at', new Date().toISOString().split('T')[0]);
        if (count >= 3) {
          return res.status(403).json({ error: 'Free plan limit: max 3 edits/saves per day' });
        }
      }

      const key = req.userRole === 'superadmin' ? CONFIG.EXCEL_FILE_KEY : req.fileKey;
      if (!key) return res.status(404).json({ error: 'Workbook key not found' });

      const { buffer, error } = await getBufferFromStorage(req.supabase, CONFIG.EXCEL_BUCKET, key);
      if (error || !buffer) return res.status(404).json({ error: 'Workbook not found' });

      const workbook = await loadWorkbook(buffer);
      const ws = workbook.getWorksheet(name);
      if (!ws) return res.status(400).json({ error: 'Sheet not found' });

      workbook.removeWorksheet(ws.id);

      const outBuffer = await workbook.xlsx.writeBuffer();
      const { error: uploadError } = await updateBufferToStorage(
        req.supabase,
        CONFIG.EXCEL_BUCKET,
        key,
        outBuffer,
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      );
      if (uploadError) return res.status(500).json({ error: uploadError.message });

      const reloaded = await loadWorkbook(outBuffer);
      const remainingSheets = reloaded.worksheets.map(ws => ws.name);

      await req.supabase.from('excel_audit').insert({
        user_id: req.user.id,
        email: req.userEmail,
        action: 'delete_sheet',
        sheet_name: name,
        metadata: { role: req.userRole, plan: req.userPlan },
        details: { remainingSheets }
      });

      broadcastSSE('excel:delete_sheet', { by: req.userEmail, sheet: name, key });

      // Reorder: keep latest edit pinned via frontend using /excel/latest
      res.json({ success: true, deleted: name, sheets: remainingSheets });
    } catch (e) {
      res.status(500).json({ error: `Unexpected server error: ${e.message}` });
    }
  }
);

/* -------------------------------------------------------
   Get cell value (free users limited, no ads gate)
------------------------------------------------------- */
app.get(
  '/excel/get',
  requireAuth,
  attachUserContext,
  attachUserPlanAndFile,
  async (req, res) => {
    const { sheet, cell, scope } = req.query;
    if (!sheet || !cell) return res.status(400).json({ error: 'Sheet and cell are required' });

    try {
      // Free plan limit: max 3 get/save/edit per day
      if (req.userPlan === 'free') {
        const { count } = await req.supabase
          .from('excel_audit')
          .select('*', { count: 'exact', head: true })
          .eq('user_id', req.user.id)
          .in('action', ['get_cell', 'save_all', 'delete_sheet', 'add_sheet'])
          .gte('created_at', new Date().toISOString().split('T')[0]);
        if (count >= 3) {
          return res.status(403).json({ error: 'Free plan limit: max 3 queries/saves per day' });
        }
      }

      const key = req.userRole === 'superadmin' && scope === 'master' ? CONFIG.EXCEL_FILE_KEY : req.fileKey;
      if (!key) return res.status(404).json({ error: 'Workbook key not found' });

      // Load buffer using authenticated client
      const { buffer, error } = await getBufferFromStorage(req.supabase, CONFIG.EXCEL_BUCKET, key);
      if (error || !buffer) return res.status(404).json({ error: 'Workbook not found' });

      const workbook = await loadWorkbook(buffer);
      const ws = workbook.getWorksheet(sheet);
      if (!ws) return res.status(400).json({ error: `Sheet "${sheet}" not found` });

      const value = ws.getCell(cell).value ?? null;

      await req.supabase.from('excel_audit').insert({
        user_id: req.user.id,
        email: req.userEmail,
        action: 'get_cell',
        sheet_name: sheet,
        metadata: { cell, value },
        details: { scope, source: key }
      });

      res.json({ sheet, cell, value, source: key });
    } catch (e) {
      res.status(500).json({ error: `Unexpected server error: ${e.message}` });
    }
  }
);

/* -------------------------------------------------------
   List sheet names (premium logic: pin latest, preserve order)
------------------------------------------------------- */
app.get('/excel/sheets', requireAuth, attachUserContext, attachUserPlanAndFile, async (req, res) => {
  try {
    const key = req.fileKey;
    if (!key) return res.status(404).json({ error: 'Workbook not found' });

    // Load user workbook
    const { buffer, error } = await getBufferFromStorage(req.supabase, CONFIG.EXCEL_BUCKET, key);
    if (error || !buffer) return res.status(404).json({ error: 'Workbook not found' });

    const workbook = await loadWorkbook(buffer);
    let sheets = workbook.worksheets.map(ws => ws.name);

    // If superadmin, merge master workbook
    if (req.userRole === 'superadmin' && CONFIG.EXCEL_FILE_KEY) {
      const { buffer: masterBuffer, error: masterErr } = await getBufferFromStorage(
        req.supabase,
        CONFIG.EXCEL_BUCKET,
        CONFIG.EXCEL_FILE_KEY
      );
      if (!masterErr && masterBuffer) {
        const masterWorkbook = await loadWorkbook(masterBuffer);
        const masterSheets = masterWorkbook.worksheets.map(ws => ws.name);
        sheets = [...new Set([...sheets, ...masterSheets])];
      }
    }

    // Determine latest edited/saved sheet from audit
    let latestSheet = null;
    const { data: auditRows, error: auditErr } = await req.supabase
      .from('excel_audit')
      .select('sheet_name, created_at')
      .eq('user_id', req.user.id)
      .in('action', ['save_all', 'add_sheet', 'delete_sheet'])
      .order('created_at', { ascending: false })
      .limit(1);

    if (!auditErr && Array.isArray(auditRows) && auditRows.length > 0) {
      latestSheet = auditRows[0].sheet_name || null;
    }

    // Pin latest at top, preserve original order for the rest
    if (sheets.length > 0 && latestSheet) {
      const latestIndex = sheets.indexOf(latestSheet);
      if (latestIndex !== -1) {
        sheets.splice(latestIndex, 1); // remove from current position
        sheets.unshift(latestSheet);   // insert at top
      }
    }

    res.json({ sheets });
  } catch (e) {
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});

/* -------------------------------------------------------
   Get latest edited/saved sheet + index in list
------------------------------------------------------- */
app.get('/excel/latest', requireAuth, attachUserContext, attachUserPlanAndFile, async (req, res) => {
  try {
    // Find latest sheet from audit
    const { data: auditRows, error: auditErr } = await req.supabase
      .from('excel_audit')
      .select('sheet_name, created_at')
      .eq('user_id', req.user.id)
      .in('action', ['save_all', 'add_sheet', 'delete_sheet'])
      .order('created_at', { ascending: false })
      .limit(1);

    if (auditErr || !Array.isArray(auditRows) || auditRows.length === 0) {
      return res.json({ sheet: null, index: null });
    }

    const latestSheet = auditRows[0].sheet_name || null;

    // Load workbook sheets in original order
    const key = req.fileKey;
    if (!key) return res.json({ sheet: latestSheet, index: null });

    const { buffer, error } = await getBufferFromStorage(req.supabase, CONFIG.EXCEL_BUCKET, key);
    if (error || !buffer) return res.json({ sheet: latestSheet, index: null });

    const workbook = await loadWorkbook(buffer);
    let sheets = workbook.worksheets.map(ws => ws.name);

    // If superadmin, merge master workbook
    if (req.userRole === 'superadmin' && CONFIG.EXCEL_FILE_KEY) {
      const { buffer: masterBuffer, error: masterErr } = await getBufferFromStorage(
        req.supabase,
        CONFIG.EXCEL_BUCKET,
        CONFIG.EXCEL_FILE_KEY
      );
      if (!masterErr && masterBuffer) {
        const masterWorkbook = await loadWorkbook(masterBuffer);
        const masterSheets = masterWorkbook.worksheets.map(ws => ws.name);
        sheets = [...new Set([...sheets, ...masterSheets])];
      }
    }

    // Pin latest at top
    let index = null;
    if (sheets.length > 0 && latestSheet) {
      const latestIndex = sheets.indexOf(latestSheet);
      if (latestIndex !== -1) {
        sheets.splice(latestIndex, 1);
        sheets.unshift(latestSheet);
        index = 0; // always pinned at top
      }
    }

    res.json({ sheet: latestSheet, index });
  } catch (e) {
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});

/* -------------------------------------------------------
   Save entire grid (free users limited, no ads gate)
   - Persists all provided rows/cols
   - Clears trailing cells if grid shrinks
------------------------------------------------------- */

/* -------------------------------------------------------
Save entire grid (free users limited, no ads gate)
- Persists all provided rows/cols
- Clears trailing cells if grid shrinks
------------------------------------------------------- */
app.post('/excel/save-all', requireAuth, attachUserContext, attachUserPlanAndFile, async (req, res) => {
  const { sheet, data } = req.body;
  if (!sheet || !data) {
    return res.status(400).json({ error: 'Sheet and data are required' });
  }

  try {
    // Check free plan limits
    if (req.userPlan === 'free') {
      const { count } = await req.supabase
        .from('excel_audit')
        .select('*', { count: 'exact', head: true })
        .eq('user_id', req.user.id)
        .eq('action', 'save_all')
        .gte('created_at', new Date().toISOString().split('T')[0]);

      if (count >= 3) {
        return res.status(403).json({ error: 'Free plan limit: max 3 saves per day' });
      }

      const rowsCount = Array.isArray(data) ? data.length : 0;
      if (rowsCount > CONFIG.FREE_LIMITS.maxRowsSaveAll) {
        return res.status(403).json({
          error: `Free plan limit: max ${CONFIG.FREE_LIMITS.maxRowsSaveAll} rows per save`,
        });
      }
    }

    const key = req.fileKey;
    if (!key) {
      return res.status(404).json({ error: 'Workbook key not found' });
    }

    const { buffer, error } = await getBufferFromStorage(req.supabase, CONFIG.EXCEL_BUCKET, key);
    let workbook;
    let ws;

    if (error || !buffer) {
      workbook = new ExcelJS.Workbook();
      ws = workbook.addWorksheet(sheet);
    } else {
      workbook = await loadWorkbook(buffer);
      ws = workbook.getWorksheet(sheet) || workbook.addWorksheet(sheet);
    }

    // Normalize incoming grid and determine bounds
    const grid = Array.isArray(data) ? data.map((r) => (Array.isArray(r) ? r : [r])) : [[]];
    const maxRowIn = grid.length;
    const maxColIn = grid.reduce((m, r) => Math.max(m, r.length), 0);

    // Apply edits to provided bounds
    for (let r = 0; r < maxRowIn; r++) {
      const excelRow = ws.getRow(r + 1);
      const row = grid[r] || [];
      for (let c = 0; c < maxColIn; c++) {
        const cell = excelRow.getCell(c + 1);
        const newVal = row[c] ?? null;
        const oldVal = cell.value ?? null;
        if (oldVal !== newVal) {
          cell.value = newVal;
        }
      }
      excelRow.commit();
    }

    // Clear any trailing cells if the grid shrank compared to existing sheet dimensions
    const existingMaxRow = ws.actualRowCount || ws.rowCount;
    const existingMaxCol = ws.actualColumnCount || (ws.columns ? ws.columns.length : 0);

    // Clear rows beyond new maxRowIn
    for (let r = maxRowIn + 1; r <= existingMaxRow; r++) {
      const excelRow = ws.getRow(r);
      for (let c = 1; c <= existingMaxCol; c++) {
        const cell = excelRow.getCell(c);
        if (cell.value != null) cell.value = null;
      }
      excelRow.commit();
    }

    // Clear cols beyond new maxColIn for existing rows
    for (let r = 1; r <= Math.max(existingMaxRow, maxRowIn); r++) {
      const excelRow = ws.getRow(r);
      for (let c = maxColIn + 1; c <= existingMaxCol; c++) {
        const cell = excelRow.getCell(c);
        if (cell.value != null) cell.value = null;
      }
      excelRow.commit();
    }

    // Write buffer and upload
    const outBuffer = await workbook.xlsx.writeBuffer();
    const { error: uploadError } = await updateBufferToStorage(
      req.supabase,
      CONFIG.EXCEL_BUCKET,
      key,
      outBuffer,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );

    if (uploadError) {
      return res.status(500).json({ error: uploadError.message });
    }

    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      email: req.userEmail,
      action: 'save_all',
      sheet_name: sheet,
      metadata: {
        rows: maxRowIn,
        cols: maxColIn,
        role: req.userRole,
        plan: req.userPlan,
      },
      details: { updated: true, key },
    });

    broadcastSSE('excel:save_all', {
      by: req.userEmail,
      sheet,
      rows: maxRowIn,
      cols: maxColIn,
      key,
    });

    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});

/* -------------------------------------------------------
   Overwrite sheet (utility)
------------------------------------------------------- */
app.post('/excel/overwrite', requireAuth, attachUserPlanAndFile, async (req, res) => {
  const { name } = req.body;
  if (!name) return res.status(400).json({ error: 'Sheet name required' });

  try {
    const key = req.fileKey;
    const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
    if (error || !buffer) return res.status(404).json({ error: 'Workbook not found' });

    const workbook = await loadWorkbook(buffer);
    const ws = workbook.getWorksheet(name);
    if (ws) {
      workbook.removeWorksheet(ws.id);
    }
    const newWs = workbook.addWorksheet(name);
    newWs.getCell('A1').value = 'Overwritten sheet';

    const outBuffer = await workbook.xlsx.writeBuffer();
    const { error: uploadError } = await updateBufferToStorage(
      baseSupabase,
      CONFIG.EXCEL_BUCKET,
      key,
      outBuffer,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    if (uploadError) return res.status(500).json({ error: uploadError.message });

    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      email: req.userEmail,
      action: 'overwrite_sheet',
      sheet_name: name,
      details: { overwritten: true }
    });

    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});

/* -------------------------------------------------------
   Preview full sheet (no ads gate, faster load, robust dims)
------------------------------------------------------- */
app.get(
  '/excel/preview',
  requireAuth,
  attachUserContext,
  attachUserPlanAndFile,
  async (req, res) => {
    const { sheet } = req.query;
    if (!sheet) return res.status(400).json({ error: 'Sheet is required' });

    try {
      // Determine storage key
      const key = req.user ? req.fileKey : CONFIG.EXCEL_FILE_KEY;
      if (!key) return res.status(404).json({ error: 'Workbook key not found' });

      if (!req.supabase || !req.supabase.storage) {
        return res.status(500).json({ error: 'Supabase client not initialized' });
      }

      // Load buffer using authenticated client
      const { buffer, error } = await getBufferFromStorage(req.supabase, CONFIG.EXCEL_BUCKET, key);
      if (error || !buffer) return res.status(404).json({ error: 'Workbook not found' });

      const workbook = await loadWorkbook(buffer);
      const ws = workbook.getWorksheet(sheet);
      if (!ws) return res.status(400).json({ error: `Sheet "${sheet}" not found` });

      const preview = [];
      const maxRow = ws.actualRowCount || ws.rowCount || 0;
      const maxCol = ws.actualColumnCount || (ws.columns ? ws.columns.length : 0);

      // Always return at least a 1x1 grid
      const safeRowCount = Math.max(maxRow, 1);
      const safeColCount = Math.max(maxCol, 1);

      for (let r = 1; r <= safeRowCount; r++) {
        const row = [];
        for (let c = 1; c <= safeColCount; c++) {
          row.push(ws.getRow(r).getCell(c).value ?? null);
        }
        preview.push(row);
      }

      res.json({ sheet, preview, rows: maxRow, cols: maxCol });
    } catch (e) {
      res.status(500).json({ error: `Unexpected server error: ${e.message}` });
    }
  }
);

/* -------------------------------------------------------
   Metadata
------------------------------------------------------- */
app.get('/excel/meta', attachUserPlanAndFile, async (req, res) => {
  try {
    const key = req.user ? req.fileKey : CONFIG.EXCEL_FILE_KEY;
    const dir = key.includes('/') ? key.split('/').slice(0, -1).join('/') : '';
    const { data, error } = await baseSupabase.storage.from(CONFIG.EXCEL_BUCKET).list(dir, { limit: 100 });
    if (error) return res.status(404).json({ error: error.message });
    const file = data.find((f) => f.name === key.split('/').pop());
    if (!file) return res.status(404).json({ error: 'File not found' });
    res.json({ name: file.name, size: file.size, last_modified: file.updated_at });
  } catch (e) {
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});

/* -------------------------------------------------------
   Export CSV (premium-only)
------------------------------------------------------- */
app.get('/excel/export/csv', requireAuth, attachUserContext, attachUserPlanAndFile, async (req, res) => {
  const { sheet } = req.query;
  if (!sheet) return res.status(400).json({ error: 'Sheet is required' });

  try {
    // Block free plan entirely
    if (req.userPlan === 'free') {
      return res.status(403).json({ error: 'CSV export is available only on premium plans' });
    }

    const key = req.user ? req.fileKey : CONFIG.EXCEL_FILE_KEY;
    const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
    if (error || !buffer) return res.status(404).json({ error: 'Workbook not found' });

    const workbook = await loadWorkbook(buffer);
    const ws = workbook.getWorksheet(sheet);
    if (!ws) return res.status(400).json({ error: 'Sheet not found' });

    const maxRow = ws.actualRowCount || ws.rowCount;
    const maxCol = ws.actualColumnCount || (ws.columns ? ws.columns.length : 0);

    const escapeCSV = (v) => {
      if (v === null || v === undefined) return '';
      const s = String(v);
      if (s.includes('"') || s.includes(',') || s.includes('\n') || s.includes('\r')) {
        return `"${s.replace(/"/g, '""')}"`;
      }
      return s;
    };

    let csv = '';
    for (let r = 1; r <= maxRow; r++) {
      const rowVals = [];
      for (let c = 1; c <= maxCol; c++) {
        const cellVal = ws.getRow(r).getCell(c).value;
        rowVals.push(escapeCSV(cellVal));
      }
      csv += rowVals.join(',') + '\n';
    }

    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      email: req.userEmail,
      action: 'export_csv',
      sheet_name: sheet,
      details: { sheet }
    });

    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${sheet}.csv"`);
    res.send(csv);
  } catch (e) {
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});

/* -------------------------------------------------------
   Export PDF single (premium-only)
------------------------------------------------------- */
app.get('/excel/export/pdf', requireAuth, attachUserContext, attachUserPlanAndFile, async (req, res) => {
  const { sheet } = req.query;
  if (!sheet) return res.status(400).json({ error: 'Sheet is required' });

  try {
    // Block free plan entirely
    if (req.userPlan === 'free') {
      return res.status(403).json({ error: 'PDF export is available only on premium plans' });
    }

    const key = req.user ? req.fileKey : CONFIG.EXCEL_FILE_KEY;
    const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
    if (error || !buffer) return res.status(404).json({ error: 'Workbook not found' });

    const workbook = await loadWorkbook(buffer);
    const ws = workbook.getWorksheet(sheet);
    if (!ws) return res.status(400).json({ error: 'Sheet not found' });

    const PDFDocument = (await import('pdfkit')).default;
    const doc = new PDFDocument({ size: 'A4', margin: 36 });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${sheet}.pdf"`);
    doc.pipe(res);

    const cellPaddingX = 6;
    const cellPaddingY = 4;
    const rowHeight = 18;
    const colWidth = 100;
    const maxRow = ws.actualRowCount || ws.rowCount;
    const maxCol = ws.actualColumnCount || (ws.columns ? ws.columns.length : 0);

    const startX = doc.page.margins.left;
    let cursorY = doc.page.margins.top;

    doc.fontSize(16).fillColor('#111111').text(`Sheet: ${sheet}`, { align: 'left' });
    doc.moveDown(0.5);

    doc.save();
    doc.rect(startX, cursorY, colWidth * maxCol, rowHeight).fill('#f3f4f6');
    doc.restore();

    doc.fontSize(10).fillColor('#111827');
    for (let c = 0; c < maxCol; c++) {
      const x = startX + c * colWidth;
      doc.text(`Col ${c + 1}`, x + cellPaddingX, cursorY + cellPaddingY, {
        width: colWidth - cellPaddingX * 2,
        height: rowHeight - cellPaddingY * 2,
        ellipsis: true,
      });
      doc.rect(x, cursorY, colWidth, rowHeight).stroke('#d1d5db');
    }
    cursorY += rowHeight;

    doc.fontSize(10).fillColor('#111111');
    for (let r = 1; r <= maxRow; r++) {
      if (cursorY + rowHeight > doc.page.height - doc.page.margins.bottom) {
        doc.addPage();
        cursorY = doc.page.margins.top;
      }
      for (let c = 1; c <= maxCol; c++) {
        const x = startX + (c - 1) * colWidth;
        const val = ws.getRow(r).getCell(c).value;
        const text = val === null || val === undefined ? '' : String(val);
        doc.text(text, x + cellPaddingX, cursorY + cellPaddingY, {
          width: colWidth - cellPaddingX * 2,
          height: rowHeight - cellPaddingY * 2,
          ellipsis: true,
        });
        doc.rect(x, cursorY, colWidth, rowHeight).stroke('#e5e7eb');
      }
      cursorY += rowHeight;
    }

    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      email: req.userEmail,
      action: 'export_pdf_single',
      sheet_name: sheet,
      details: { sheet }
    });

    doc.end();
  } catch (e) {
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});

/* -------------------------------------------------------
   Export PDF multi (premium-only)
------------------------------------------------------- */
app.get('/excel/export/pdf-multi', requireAuth, attachUserContext, attachUserPlanAndFile, async (req, res) => {
  const { sheets } = req.query;
  try {
    // Block free plan entirely
    if (req.userPlan === 'free') {
      return res.status(403).json({ error: 'Multi-sheet PDF export is available only on premium plans' });
    }

    const key = req.user ? req.fileKey : CONFIG.EXCEL_FILE_KEY;
    const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
    if (error || !buffer) return res.status(404).json({ error: 'Workbook not found' });

    const workbook = await loadWorkbook(buffer);
    let sheetList = sheets
      ? sheets.split(',').map(s => s.trim()).filter(Boolean)
      : workbook.worksheets.map(ws => ws.name);

    const PDFDocument = (await import('pdfkit')).default;
    const doc = new PDFDocument({ size: 'A4', margin: 36 });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="workbook.pdf"');
    doc.pipe(res);

    const drawTable = (ws, title) => {
      doc.fontSize(16).fillColor('#111111').text(`Sheet: ${title}`, { align: 'left' });
      doc.moveDown(0.5);
      const cellPaddingX = 6;
      const cellPaddingY = 4;
      const rowHeight = 18;
      const colWidth = 100;
      const maxRow = ws.actualRowCount || ws.rowCount;
      const maxCol = ws.actualColumnCount || (ws.columns ? ws.columns.length : 0);
      const startX = doc.page.margins.left;
      let cursorY = doc.page.margins.top;

      doc.save();
      doc.rect(startX, cursorY, colWidth * maxCol, rowHeight).fill('#f3f4f6');
      doc.restore();

      doc.fontSize(10).fillColor('#111827');
      for (let c = 0; c < maxCol; c++) {
        const x = startX + c * colWidth;
        doc.text(`Col ${c + 1}`, x + cellPaddingX, cursorY + cellPaddingY, {
          width: colWidth - cellPaddingX * 2,
          height: rowHeight - cellPaddingY * 2,
          ellipsis: true,
        });
        doc.rect(x, cursorY, colWidth, rowHeight).stroke('#d1d5db');
      }
      cursorY += rowHeight;

      doc.fontSize(10).fillColor('#111111');
      for (let r = 1; r <= maxRow; r++) {
        if (cursorY + rowHeight > doc.page.height - doc.page.margins.bottom) {
          doc.addPage();
          cursorY = doc.page.margins.top;
        }
        for (let c = 1; c <= maxCol; c++) {
          const x = startX + (c - 1) * colWidth;
          const val = ws.getRow(r).getCell(c).value;
          const text = val === null || val === undefined ? '' : String(val);
          doc.text(text, x + cellPaddingX, cursorY + cellPaddingY, {
            width: colWidth - cellPaddingX * 2,
            height: rowHeight - cellPaddingY * 2,
            ellipsis: true,
          });
          doc.rect(x, cursorY, colWidth, rowHeight).stroke('#e5e7eb');
        }
        cursorY += rowHeight;
      }
    };

    sheetList.forEach((name, idx) => {
      const ws = workbook.getWorksheet(name);
      if (!ws) return;
      if (idx > 0) doc.addPage();
      drawTable(ws, name);
    });

    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      email: req.userEmail,
      action: 'export_pdf_multi',
      details: { sheets: sheetList }
    });

    doc.end();
  } catch (e) {
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});

/* -------------------------------------------------------
   Audit retrieval for Dashboard (premium-only)
------------------------------------------------------- */
app.get('/audit/list', requireAuth, attachUserPlanAndFile, async (req, res) => {
  try {
    // Free plan restriction
    if (req.userPlan === 'free') {
      return res.status(403).json({ error: 'Audit logs are not available on the free plan' });
    }

    const role = req.userRole;
    const limit = Math.min(parseInt(req.query.limit, 10) || 100, 500); // cap limit for safety
    const offset = parseInt(req.query.offset, 10) || 0;

    let query = req.supabase
      .from('excel_audit')
      .select('created_at, user_id, action, sheet_name, metadata, details, email')
      .order('created_at', { ascending: false })
      .range(offset, offset + limit - 1);

    // Non-superadmin users only see their own logs
    if (role !== 'superadmin' && req.userEmail !== CONFIG.OWNER_EMAIL) {
      query = query.eq('user_id', req.user.id);
    }

    const { data, error } = await query;
    if (error) return res.status(500).json({ error: error.message });

    // Normalize fields for frontend AuditTable
    const entries = (data || []).map((row) => ({
      ts: row.created_at,
      actor: row.email || row.user_id || '—',
      action: row.action,
      sheet: row.sheet_name,
      details: row.details || {},
      metadata: row.metadata || {},
      email: row.email || null
    }));

    res.json({ entries });
  } catch (e) {
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});

/* =======================================================
   PAYMENTS: Paystack init, verify, webhook (raw body)
======================================================= */
// Initialize a Paystack transaction
app.post('/payments/init', requireAuth, async (req, res) => {
  const { email, amount, mode } = req.body;
  if (!email) return res.status(400).json({ error: 'Email is required' });

  try {
    let body;
    if (mode === 'monthly') {
      if (!process.env.PAYSTACK_MONTHLY_PLAN) {
        return res.status(500).json({ error: 'Monthly plan code not configured' });
      }
      body = {
        email,
        plan: process.env.PAYSTACK_MONTHLY_PLAN,
        callback_url: process.env.PAYSTACK_CALLBACK_URL || 'https://vsbil-excel.netlify.app/'
      };
    } else {
      if (!amount) return res.status(400).json({ error: 'Amount is required for one-time payments' });
      body = {
        email,
        amount: parseInt(amount, 10),
        currency: 'GHS',
        callback_url: process.env.PAYSTACK_CALLBACK_URL || 'https://vsbil-excel.netlify.app/'
      };
    }

    const response = await fetch('https://api.paystack.co/transaction/initialize', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${process.env.PAYSTACK_SECRET_KEY}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
    });

    const data = await response.json();
    if (!response.ok || data.status !== true) {
      return res.status(response.status).json({ error: data.message || 'Paystack init failed' });
    }

    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Verify a Paystack transaction manually
app.post('/payments/verify', requireAuth, async (req, res) => {
  const { reference, email, mode } = req.body;
  if (!reference || !email) return res.status(400).json({ error: 'Reference and email are required' });

  try {
    const response = await fetch(`https://api.paystack.co/transaction/verify/${reference}`, {
      method: 'GET',
      headers: { Authorization: `Bearer ${process.env.PAYSTACK_SECRET_KEY}` },
    });

    const data = await response.json();
    if (!response.ok || data.status !== true) {
      return res.status(response.status).json({ error: data.message || 'Verification failed' });
    }

    const tx = data.data;
    const amount = tx.amount;

    // Decide expiry window
    let expiresAt;
    if (mode === 'monthly') {
      expiresAt = new Date();
      expiresAt.setMonth(expiresAt.getMonth() + 1);
    } else {
      expiresAt = new Date();
      expiresAt.setFullYear(expiresAt.getFullYear() + 1);
    }

    // Update profile
    const { data: updated, error: updateErr } = await supabase
      .from('profiles')
      .update({ plan: 'paid', premium_expires_at: expiresAt.toISOString() })
      .eq('email', email.toLowerCase())
      .select();

    if (updateErr) {
      return res.status(500).json({ error: 'Failed to update profile' });
    }

    // Log payment
    const { error: payErr } = await supabase.from('payments').insert({
      email,
      amount,
      reference,
      status: tx.status,
      mode,
      created_at: new Date().toISOString(),
    });
    if (payErr) {
      return res.status(500).json({ error: 'Failed to log payment' });
    }

    // Redirect back to main page after verification
    res.redirect(process.env.PAYSTACK_SUCCESS_REDIRECT || '/');
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Paystack webhook
app.post('/payments/webhook', express.raw({ type: '*/*' }), async (req, res) => {
  try {
    const secret = process.env.PAYSTACK_SECRET_KEY;
    const signature = req.headers['x-paystack-signature'];
    const hash = crypto.createHmac('sha512', secret).update(req.body).digest('hex');
    if (hash !== signature) return res.status(401).json({ error: 'Invalid signature' });

    const event = JSON.parse(req.body.toString());
    if (event.event === 'charge.success') {
      const { email } = event.data.customer;
      const amount = event.data.amount;
      const reference = event.data.reference;
      const planCode = event.data.plan?.plan_code;

      let expiresAt;
      if (planCode === process.env.PAYSTACK_MONTHLY_PLAN) {
        expiresAt = new Date();
        expiresAt.setMonth(expiresAt.getMonth() + 1);
      } else {
        expiresAt = new Date();
        expiresAt.setFullYear(expiresAt.getFullYear() + 1);
      }

      try {
        const { data: updated, error: updateErr } = await supabase
          .from('profiles')
          .update({ plan: 'paid', premium_expires_at: expiresAt.toISOString() })
          .eq('email', email.toLowerCase())
          .select();
        if (updateErr) {}
        if (!updated || updated.length === 0) {}
        const { error: payErr } = await supabase.from('payments').insert({
          email,
          amount,
          reference,
          status: 'success',
          mode: planCode === process.env.PAYSTACK_MONTHLY_PLAN ? 'monthly' : 'one-time',
          created_at: new Date().toISOString(),
        });
        if (payErr) {}

        broadcastSSE('payments:success', { email, amount, reference, plan: 'paid', expiresAt });
      } catch (err) {}
    }
    res.sendStatus(200);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* =======================================================
   ADMIN / SUPERADMIN / OWNER MANAGEMENT
======================================================= */

/* Promote to superadmin — owner only */
app.post('/admin/users/promote', requireAuth, attachUserPlanAndFile, requireOwner, async (req, res) => {
  const { user_id } = req.body;
  if (!user_id) return res.status(400).json({ error: 'user_id required' });
  try {
    const { data, error } = await baseSupabase
      .from('profiles')
      .update({ role: 'superadmin' })
      .eq('id', user_id)
      .select('id, role, email')
      .single();
    if (error) return res.status(500).json({ error: error.message });
    await writeAuditLog({ ts: Date.now(), action: 'promote_superadmin', by: req.userEmail, target: user_id });
    broadcastSSE('admin:promote', { target: data.email, role: 'superadmin' });
    res.json({ success: true, user: data });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* Denote role — owner only */
app.post('/admin/users/denote', requireAuth, attachUserPlanAndFile, requireOwner, async (req, res) => {
  const { user_id, role = 'user' } = req.body;
  if (!user_id) return res.status(400).json({ error: 'user_id required' });
  try {
    const { data, error } = await baseSupabase
      .from('profiles')
      .update({ role })
      .eq('id', user_id)
      .select('id, role, email')
      .single();
    if (error) return res.status(500).json({ error: error.message });
    await writeAuditLog({ ts: Date.now(), action: 'denote_role', by: req.userEmail, target: user_id, role });
    broadcastSSE('admin:denote', { target: data.email, role });
    res.json({ success: true, user: data });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* List users — superadmin or owner */
app.get('/admin/users/list', requireAuth, attachUserPlanAndFile, async (req, res) => {
  try {
    if (req.userRole !== 'superadmin' && req.userEmail !== CONFIG.OWNER_EMAIL) {
      return res.status(403).json({ error: 'Forbidden' });
    }
    const limit = parseInt(req.query.limit, 10) || 100;
    const offset = parseInt(req.query.offset, 10) || 0;
    const search = (req.query.search || '').trim();

    let query = baseSupabase
      .from('profiles')
      .select('id,email,role,plan,status,verified,created_at,user_file_key')
      .order('created_at', { ascending: false })
      .range(offset, offset + limit - 1);

    if (search) query = query.ilike('email', `%${search}%`);

    const { data, error } = await query;
    if (error) return res.status(500).json({ error: error.message });
    res.json({ users: data || [], count: (data || []).length });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* Ban user — superadmin or owner */
app.post('/admin/users/ban', requireAuth, attachUserPlanAndFile, async (req, res) => {
  if (req.userRole !== 'superadmin' && req.userEmail !== CONFIG.OWNER_EMAIL) {
    return res.status(403).json({ error: 'Forbidden' });
  }
  const { user_id, reason } = req.body;
  if (!user_id) return res.status(400).json({ error: 'user_id required' });
  try {
    const { data, error } = await baseSupabase
      .from('profiles')
      .update({ status: 'banned', ban_reason: reason || null })
      .eq('id', user_id)
      .select('id,status,email')
      .single();
    if (error) return res.status(500).json({ error: error.message });
    await writeAuditLog({ ts: Date.now(), action: 'ban_user', by: req.userEmail, target: user_id, reason });
    broadcastSSE('admin:ban', { target: data.email, reason: reason || null });
    res.json({ success: true, user: data });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* Verify user — superadmin or owner */
app.post('/admin/users/verify', requireAuth, attachUserPlanAndFile, async (req, res) => {
  if (req.userRole !== 'superadmin' && req.userEmail !== CONFIG.OWNER_EMAIL) {
    return res.status(403).json({ error: 'Forbidden' });
  }
  const { user_id } = req.body;
  if (!user_id) return res.status(400).json({ error: 'user_id required' });
  try {
    const { data, error } = await baseSupabase
      .from('profiles')
      .update({ verified: true })
      .eq('id', user_id)
      .select('id,verified,email')
      .single();
    if (error) return res.status(500).json({ error: error.message });
    await writeAuditLog({ ts: Date.now(), action: 'verify_user', by: req.userEmail, target: user_id });
    broadcastSSE('admin:verify', { target: data.email });
    res.json({ success: true, user: data });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* Soft delete user — superadmin or owner */
app.post('/admin/users/soft-delete', requireAuth, attachUserPlanAndFile, async (req, res) => {
  if (req.userRole !== 'superadmin' && req.userEmail !== CONFIG.OWNER_EMAIL) {
    return res.status(403).json({ error: 'Forbidden' });
  }
  const { user_id } = req.body;
  if (!user_id) return res.status(400).json({ error: 'user_id required' });
  try {
    const { data, error } = await baseSupabase
      .from('profiles')
      .update({ status: 'deleted' })
      .eq('id', user_id)
      .select('id,status,email')
      .single();
    if (error) return res.status(500).json({ error: error.message });
    await writeAuditLog({ ts: Date.now(), action: 'soft_delete_user', by: req.userEmail, target: user_id });
    broadcastSSE('admin:soft_delete', { target: data.email });
    res.json({ success: true, user: data });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* Permanent delete — OWNER ONLY */
app.post('/admin/users/permadelete', requireAuth, attachUserPlanAndFile, requireOwner, async (req, res) => {
  const { user_id } = req.body;
  if (!user_id) return res.status(400).json({ error: 'user_id required' });
  try {
    const { data: profile, error: profileErr } = await baseSupabase
      .from('profiles')
      .select('user_file_key,email')
      .eq('id', user_id)
      .single();
    if (profileErr) return res.status(500).json({ error: profileErr.message });

    const prefix = `${CONFIG.USER_FILES_PREFIX}/${user_id}`;
    const { data: list, error: listErr } = await baseSupabase.storage
      .from(CONFIG.EXCEL_BUCKET)
      .list(prefix, { limit: 1000 });
    if (!listErr && Array.isArray(list) && list.length) {
      const keys = list.map((f) => `${prefix}/${f.name}`);
      const { error: rmErr } = await baseSupabase.storage.from(CONFIG.EXCEL_BUCKET).remove(keys);
      if (rmErr) {}
    }

    const { error: delErr } = await baseSupabase.from('profiles').delete().eq('id', user_id);
    if (delErr) return res.status(500).json({ error: delErr.message });

    await writeAuditLog({ ts: Date.now(), action: 'permanent_delete_user', by: req.userEmail, target: user_id, email: profile?.email });
    broadcastSSE('admin:permadelete', { target: profile?.email });
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* Assist login (support session) — superadmin or owner */
app.post('/admin/support/assist', requireAuth, attachUserPlanAndFile, async (req, res) => {
  if (req.userRole !== 'superadmin' && req.userEmail !== CONFIG.OWNER_EMAIL) {
    return res.status(403).json({ error: 'Forbidden' });
  }
  const { user_id } = req.body;
  if (!user_id) return res.status(400).json({ error: 'user_id required' });
  try {
    const key = `support_${user_id}_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
    const expires_at = new Date(Date.now() + CONFIG.SUPPORT_SESSION_TTL_MIN * 60 * 1000).toISOString();
    const { error } = await baseSupabase.from('support_sessions').insert({
      user_id,
      session_key: key,
      created_by: req.user.id,
      expires_at,
      read_only: true,
    });
    if (error) return res.status(500).json({ error: error.message });
    await writeAuditLog({ ts: Date.now(), action: 'support_assist', by: req.userEmail, target: user_id, session_key: key });
    broadcastSSE('support:session_created', { target: user_id, session_key: key, expires_at });
    res.json({ success: true, session_key: key, expires_at });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* Validate support session key (support page) */
app.get('/support/session/validate', async (req, res) => {
  const { key } = req.query;
  if (!key) return res.status(400).json({ error: 'key required' });
  try {
    const { data, error } = await baseSupabase
      .from('support_sessions')
      .select('user_id,expires_at,read_only')
      .eq('session_key', key)
      .single();
    if (error || !data) return res.status(404).json({ error: 'Invalid session key' });
    if (new Date(data.expires_at).getTime() < Date.now()) {
      return res.status(410).json({ error: 'Session expired' });
    }
    res.json({ valid: true, user_id: data.user_id, read_only: data.read_only });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* Support tickets: create, list, respond */
app.post('/support/tickets', requireAuth, async (req, res) => {
  const { subject, body } = req.body;
  if (!subject || !body) return res.status(400).json({ error: 'subject and body required' });
  try {
    const { data, error } = await baseSupabase.from('support_tickets').insert({
      user_id: req.user.id,
      subject,
      body,
      status: 'open',
      created_at: new Date().toISOString(),
    }).select('*').single();
    if (error) return res.status(500).json({ error: error.message });
    broadcastSSE('support:ticket_created', { ticket_id: data.id, user_id: req.user.id, subject });
    res.json({ ticket: data });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/support/tickets', requireAuth, async (req, res) => {
  try {
    const roleQuery = await baseSupabase
      .from('profiles')
      .select('role,email')
      .eq('id', req.user.id)
      .single();
    const isAdminLike = roleQuery.data?.role === 'superadmin' || (roleQuery.data?.email === CONFIG.OWNER_EMAIL);
    let query = baseSupabase.from('support_tickets').select('*').order('created_at', { ascending: false });
    if (!isAdminLike) query = query.eq('user_id', req.user.id);
    const { data, error } = await query;
    if (error) return res.status(500).json({ error: error.message });
    res.json({ tickets: data || [] });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/support/tickets/respond', requireAuth, attachUserPlanAndFile, async (req, res) => {
  if (req.userRole !== 'superadmin' && req.userEmail !== CONFIG.OWNER_EMAIL) {
    return res.status(403).json({ error: 'Forbidden' });
  }
  const { ticket_id, response } = req.body;
  if (!ticket_id || !response) return res.status(400).json({ error: 'ticket_id and response required' });
  try {
    const { error: insErr } = await baseSupabase.from('support_responses').insert({
      ticket_id,
      user_id: req.user.id,
      response,
      created_at: new Date().toISOString(),
    });
    if (insErr) return res.status(500).json({ error: insErr.message });
    const { error: updErr } = await baseSupabase
      .from('support_tickets')
      .update({ status: 'responded' })
      .eq('id', ticket_id);
    if (updErr) {}
    broadcastSSE('support:ticket_responded', { ticket_id, by: req.userEmail });
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* Subscription analytics — owner only */
app.get('/admin/subscriptions/metrics', requireAuth, attachUserPlanAndFile, requireOwner, async (req, res) => {
  try {
    const { data: profiles, error: pErr } = await baseSupabase
      .from('profiles')
      .select('id,plan,role,status,created_at');
    if (pErr) return res.status(500).json({ error: pErr.message });

    const totalUsers = profiles.length;
    const paidUsers = profiles.filter((p) => p.plan === 'paid').length;
    const freeUsers = profiles.filter((p) => p.plan === 'free').length;
    const bannedUsers = profiles.filter((p) => p.status === 'banned').length;
    const activeUsers = profiles.filter((p) => p.status !== 'banned' && p.status !== 'deleted').length;

    const today = new Date();
    const monthStart = new Date(today.getFullYear(), today.getMonth(), 1).toISOString();
    const monthEnd = new Date(today.getFullYear(), today.getMonth() + 1, 0, 23, 59, 59).toISOString();
    const { data: payments, error: payErr } = await baseSupabase
      .from('payments')
      .select('amount,created_at,status')
      .gte('created_at', monthStart)
      .lte('created_at', monthEnd)
      .eq('status', 'success');
    if (payErr) return res.status(500).json({ error: payErr.message });

    const monthlyRevenue = (payments || []).reduce((sum, p) => sum + (p.amount || 0), 0);

    res.json({
      totals: { totalUsers, paidUsers, freeUsers, bannedUsers, activeUsers },
      revenue: { monthlyRevenue },
      period: { monthStart, monthEnd },
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* Pricing controls — owner only */
app.get('/admin/subscriptions/pricing', requireAuth, attachUserPlanAndFile, requireOwner, async (req, res) => {
  try {
    const { data, error } = await baseSupabase.from('pricing').select('*').limit(1);
    if (error) return res.status(500).json({ error: error.message });
    const cfg = Array.isArray(data) && data.length ? data[0] : { monthly_amount: 500000, yearly_amount: 5000000, currency: 'GHS' };
    res.json({ pricing: cfg });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/admin/subscriptions/pricing', requireAuth, attachUserPlanAndFile, requireOwner, async (req, res) => {
  const { monthly_amount, yearly_amount, currency } = req.body;
  if (!monthly_amount || !yearly_amount || !currency) {
    return res.status(400).json({ error: 'monthly_amount, yearly_amount, and currency are required' });
  }
  try {
    const { data: existing, error: selErr } = await baseSupabase.from('pricing').select('id').limit(1);
    if (selErr) return res.status(500).json({ error: selErr.message });
    if (Array.isArray(existing) && existing.length) {
      const id = existing[0].id;
      const { error: updErr } = await baseSupabase
        .from('pricing')
        .update({ monthly_amount, yearly_amount, currency, updated_at: new Date().toISOString() })
        .eq('id', id);
      if (updErr) return res.status(500).json({ error: updErr.message });
    } else {
      const { error: insErr } = await baseSupabase
        .from('pricing')
        .insert({ monthly_amount, yearly_amount, currency, created_at: new Date().toISOString() });
      if (insErr) return res.status(500).json({ error: insErr.message });
    }
    await writeAuditLog({ ts: Date.now(), action: 'pricing_update', by: req.userEmail, monthly_amount, yearly_amount, currency });
    broadcastSSE('pricing:update', { monthly_amount, yearly_amount, currency });
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* Legal pages: Terms & Privacy */
app.get('/legal/terms', async (req, res) => {
  try {
    const { data, error } = await baseSupabase
      .from('legal_pages')
      .select('content,updated_at')
      .eq('slug', 'terms')
      .single();
    if (error && error.code !== 'PGRST116') {}
    res.json({ content: data?.content || '', updated_at: data?.updated_at || null });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/legal/privacy', async (req, res) => {
  try {
    const { data, error } = await baseSupabase
      .from('legal_pages')
      .select('content,updated_at')
      .eq('slug', 'privacy')
      .single();
    if (error && error.code !== 'PGRST116') {}
    res.json({ content: data?.content || '', updated_at: data?.updated_at || null });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/legal/terms', requireAuth, attachUserPlanAndFile, requireOwner, async (req, res) => {
  const { content } = req.body;
  if (!content) return res.status(400).json({ error: 'content required' });
  try {
    const { data: existing, error: selErr } = await baseSupabase
      .from('legal_pages')
      .select('id')
      .eq('slug', 'terms')
      .limit(1);
    if (selErr) return res.status(500).json({ error: selErr.message });
    if (Array.isArray(existing) && existing.length) {
      const id = existing[0].id;
      const { error: updErr } = await baseSupabase
        .from('legal_pages')
        .update({ content, updated_at: new Date().toISOString() })
        .eq('id', id);
      if (updErr) return res.status(500).json({ error: updErr.message });
    } else {
      const { error: insErr } = await baseSupabase
        .from('legal_pages')
        .insert({ slug: 'terms', content, created_at: new Date().toISOString() });
      if (insErr) return res.status(500).json({ error: insErr.message });
    }
    broadcastSSE('legal:terms_update', { updated_by: req.userEmail });
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post('/legal/privacy', requireAuth, attachUserPlanAndFile, requireOwner, async (req, res) => {
  const { content } = req.body;
  if (!content) return res.status(400).json({ error: 'content required' });
  try {
    const { data: existing, error: selErr } = await baseSupabase
      .from('legal_pages')
      .select('id')
      .eq('slug', 'privacy')
      .limit(1);
    if (selErr) return res.status(500).json({ error: selErr.message });
    if (Array.isArray(existing) && existing.length) {
      const id = existing[0].id;
      const { error: updErr } = await baseSupabase
        .from('legal_pages')
        .update({ content, updated_at: new Date().toISOString() })
        .eq('id', id);
      if (updErr) return res.status(500).json({ error: updErr.message });
    } else {
      const { error: insErr } = await baseSupabase
        .from('legal_pages')
        .insert({ slug: 'privacy', content, created_at: new Date().toISOString() });
      if (insErr) return res.status(500).json({ error: insErr.message });
    }
    broadcastSSE('legal:privacy_update', { updated_by: req.userEmail });
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* =======================================================
   FILE UPLOAD / CONVERT (plan-aware, no ads gate)
======================================================= */
const upload = multer({ storage: multer.memoryStorage() });

/* -------------------------------------------------------
   Upload Excel/CSV/PDF (premium rules enforced)
------------------------------------------------------- */
app.post(
  '/excel/upload',
  requireAuth,
  attachUserContext,
  upload.single('file'),
  async (req, res) => {
    try {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
      let buffer = req.file.buffer;
      const filename = req.file.originalname;

      const isFree = req.userPlan === 'free';
      const alreadyHasFile =
        !!req.fileKey && req.fileKey.startsWith(`${CONFIG.USER_FILES_PREFIX}/${req.user.id}/`);

      // Plan enforcement
      const ext = filename.toLowerCase().split('.').pop();
      if (['csv', 'pdf'].includes(ext)) {
        if (isFree) {
          return res.status(403).json({ error: 'CSV and PDF uploads require a premium plan' });
        }
      }
      if (ext === 'xlsx') {
        if (isFree && alreadyHasFile) {
          return res.status(403).json({ error: 'Free plan allows only one Excel file upload' });
        }
      }

      let preview = [];
      let sheetNames = [];

      try {
        if (ext === 'xlsx') {
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.load(buffer);
          sheetNames = workbook.worksheets.map((ws) => ws.name);

          if (sheetNames.length === 0) {
            const defaultSheet = workbook.addWorksheet('Sheet1');
            defaultSheet.getCell('A1').value = 'New sheet created';
            sheetNames.push('Sheet1');
            buffer = await workbook.xlsx.writeBuffer();
          }

          const ws = workbook.worksheets[0];
          if (ws) {
            const maxRow = ws.actualRowCount || ws.rowCount;
            const maxCol = ws.actualColumnCount || (ws.columns ? ws.columns.length : 0);
            for (let r = 1; r <= maxRow; r++) {
              const row = [];
              for (let c = 1; c <= maxCol; c++) {
                row.push(ws.getRow(r).getCell(c).value);
              }
              preview.push(row);
            }
          }
        } else if (ext === 'csv') {
          const rows = csvParse(buffer.toString(), { columns: false });
          preview = rows.slice(0, 20);
          const workbook = new ExcelJS.Workbook();
          const ws = workbook.addWorksheet('Sheet1');
          rows.forEach((row, rIndex) => {
            row.forEach((val, cIndex) => {
              ws.getRow(rIndex + 1).getCell(cIndex + 1).value = val;
            });
          });
          sheetNames = ['Sheet1'];
          buffer = await workbook.xlsx.writeBuffer();
        } else if (ext === 'pdf') {
          // Use pdf-parse safely
          const pdfData = await require('pdf-parse')(buffer);
          preview = pdfData.text.split('\n').slice(0, 20);
          sheetNames = [filename.replace(/\.[^/.]+$/, '')];
          // Store raw PDF buffer directly (no conversion to Excel)
        } else {
          return res.status(400).json({ error: 'Unsupported file type' });
        }
      } catch (parseErr) {
        return res.status(400).json({ error: 'File parse failed: ' + parseErr.message });
      }

      // Ensure unique filename in storage
      let key = `${CONFIG.USER_FILES_PREFIX}/${req.user.id}/${filename}`;
      const { data: existingList, error: listErr } = await req.supabase.storage
        .from(CONFIG.EXCEL_BUCKET)
        .list(`${CONFIG.USER_FILES_PREFIX}/${req.user.id}`, { limit: 100 });
      if (!listErr && Array.isArray(existingList)) {
        const names = new Set(existingList.map((f) => f.name));
        if (names.has(filename)) {
          const base = filename.replace(/\.[^/.]+$/, '');
          const extn = filename.split('.').pop();
          let i = 1;
          let candidate = `${base} ${i}.${extn}`;
          while (names.has(candidate)) {
            i++;
            candidate = `${base} ${i}.${extn}`;
          }
          key = `${CONFIG.USER_FILES_PREFIX}/${req.user.id}/${candidate}`;
        }
      }

      const { error: uploadError } = await putBufferToStorage(
        req.supabase,
        CONFIG.EXCEL_BUCKET,
        key,
        buffer,
        req.file.mimetype
      );
      if (uploadError) {
        return res.status(500).json({ error: 'Storage upload failed: ' + uploadError.message });
      }

      const { error: updateErr } = await req.supabase
        .from('profiles')
        .update({ user_file_key: key })
        .eq('id', req.user.id);
      if (updateErr) {
        return res.status(500).json({ error: 'Profile update failed: ' + updateErr.message });
      }

      await req.supabase.from('excel_audit').insert({
        user_id: req.user.id,
        email: req.userEmail,
        action: 'upload_file',
        sheet_name: sheetNames[0] || null,
        details: { fileName: filename, type: ext }
      });

      broadcastSSE('excel:upload', {
        by: req.userEmail,
        fileKey: key,
        fileName: key.split('/').pop()
      });

      res.json({
        success: true,
        fileKey: key,
        fileName: key.split('/').pop(),
        sheetNames,
        preview
      });
    } catch (e) {
      return res.status(500).json({ error: `Unexpected server error: ${e.message}` });
    }
  }
);

/* -------------------------------------------------------
   Convert uploaded file to Excel and store (plan-aware)
------------------------------------------------------- */
app.post('/excel/convert', requireAuth, attachUserContext, async (req, res) => {
  try {
    const { fileBase64, fileName, fileUrl } = req.body;
    let buffer;
    if (fileBase64) {
      buffer = Buffer.from(fileBase64, 'base64');
    } else if (fileUrl) {
      const resp = await fetch(fileUrl);
      if (!resp.ok) return res.status(400).json({ error: 'Failed to fetch fileUrl' });
      const arr = await resp.arrayBuffer();
      buffer = Buffer.from(arr);
    } else {
      return res.status(400).json({ error: 'fileBase64 or fileUrl required' });
    }

    const isFree = req.userPlan === 'free';
    const alreadyHasFile = !!req.fileKey && req.fileKey.startsWith(`${CONFIG.USER_FILES_PREFIX}/${req.user.id}/`);
    const safeFileName = fileName || 'uploaded.xlsx';

    // Enforce plan rules
    if (safeFileName.toLowerCase().endsWith('.csv') || safeFileName.toLowerCase().endsWith('.pdf')) {
      if (isFree) return res.status(403).json({ error: 'CSV and PDF conversion requires a premium plan' });
    }
    if (safeFileName.toLowerCase().endsWith('.xlsx')) {
      if (isFree && alreadyHasFile) {
        return res.status(403).json({ error: 'Free plan allows only one Excel file conversion' });
      }
    }

    let workbook;
    try {
      workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      if (workbook.worksheets.length === 0) {
        const ws = workbook.addWorksheet('Sheet1');
        ws.getCell('A1').value = 'New sheet created';
        buffer = await workbook.xlsx.writeBuffer();
      }
    } catch (e) {
      return res.status(400).json({ error: 'Invalid Excel file' });
    }

    // Ensure unique filename
    let key = `${CONFIG.USER_FILES_PREFIX}/${req.user.id}/${safeFileName}`;
    const { data: existingList, error: listErr } = await req.supabase.storage
      .from(CONFIG.EXCEL_BUCKET)
      .list(`${CONFIG.USER_FILES_PREFIX}/${req.user.id}`, { limit: 100 });
    if (!listErr && Array.isArray(existingList)) {
      const names = new Set(existingList.map((f) => f.name));
      if (names.has(safeFileName)) {
        const base = safeFileName.replace(/\.[^/.]+$/, '');
        const ext = safeFileName.split('.').pop();
        let i = 1;
        let candidate = `${base} ${i}.${ext}`;
        while (names.has(candidate)) {
          i++;
          candidate = `${base} ${i}.${ext}`;
        }
        key = `${CONFIG.USER_FILES_PREFIX}/${req.user.id}/${candidate}`;
      }
    }

    const { error: uploadError } = await putBufferToStorage(
      req.supabase,
      CONFIG.EXCEL_BUCKET,
      key,
      buffer,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    if (uploadError) return res.status(500).json({ error: uploadError.message });

    const { error: updateErr } = await req.supabase
      .from('profiles')
      .update({ user_file_key: key })
      .eq('id', req.user.id);
    if (updateErr) return res.status(500).json({ error: updateErr.message });

    // Audit log
    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      email: req.userEmail,
      action: 'convert_file',
      sheet_name: workbook.worksheets[0]?.name || null,
      details: { fileName: safeFileName }
    });

    broadcastSSE('excel:convert', { by: req.userEmail, fileKey: key, fileName: key.split('/').pop() });
    res.json({
      success: true,
      fileKey: key,
      appUrl: `/app/${req.user.id}`,
      fileName: key.split('/').pop(),
      sheetNames: workbook.worksheets.map(ws => ws.name),
    });
  } catch (e) {
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});

export { app, baseSupabase as supabase };
