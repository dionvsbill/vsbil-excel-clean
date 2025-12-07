// server/app.js
import 'dotenv/config';
import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import rateLimit from 'express-rate-limit';
import { createClient } from '@supabase/supabase-js';
import ExcelJS from 'exceljs';
import { nanoid } from 'nanoid';
import multer from 'multer';
import * as pdfParse from 'pdf-parse';
import { parse as csvParse } from 'csv-parse/sync';
import { attachUserContext, requirePremium, requireAdsSoft } from './policies.js';

const app = express();

/* -------------------------------------------------------
   Security + middleware
------------------------------------------------------- */
app.use(helmet({ crossOriginResourcePolicy: false }));
app.use(
  cors({
    origin: (origin, cb) => cb(null, true),
    credentials: true,
  })
);
app.use(express.json({ limit: '10mb' }));

const limiter = rateLimit({
  windowMs: 60 * 1000,
  max: 300,
});
app.use(limiter);

const REQUIRED_ENV = ['SUPABASE_URL', 'SUPABASE_ANON_KEY'];
for (const k of REQUIRED_ENV) {
  if (!process.env[k]) {
    throw new Error(`Missing env: ${k} must be set in .env`);
  }
}
const requirePaidPlan = async (req, res, next) => {
  if (req.userPlan !== 'paid' && req.userRole !== 'superadmin') {
    return res.status(403).json({ error: 'Requires premium plan' });
  }
  next();
};

/* -------------------------------------------------------
   Supabase clients
------------------------------------------------------- */
const baseSupabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_ANON_KEY, {
  auth: { persistSession: false, autoRefreshToken: false },
});

const getSupabaseForToken = (token) =>
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
};

/* -------------------------------------------------------
   Helpers
------------------------------------------------------- */
const getBufferFromStorage = async (supabase, bucket, key) => {
  const { data, error } = await supabase.storage.from(bucket).download(key);
  if (error) return { error };
  const buffer = Buffer.from(await data.arrayBuffer());
  return { buffer };
};

const putBufferToStorage = async (supabase, bucket, key, buffer, contentType) => {
  const { error } = await supabase.storage.from(bucket).upload(key, buffer, {
    upsert: true,
    contentType,
  });
  return { error };
};

const updateBufferToStorage = async (supabase, bucket, key, buffer, contentType) => {
  const { error } = await supabase.storage.from(bucket).update(key, buffer, {
    contentType,
  });
  return { error };
};

const loadWorkbook = async (buffer) => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
};

/* -------------------------------------------------------
   Audit helper (NDJSON logs in storage)
------------------------------------------------------- */
async function writeAuditLog(entry) {
  try {
    const day = new Date().toISOString().slice(0, 10);
    const objectKey = `${CONFIG.LOGS_PREFIX}/${day}.ndjson`;

    let current = '';
    const existing = await baseSupabase.storage.from(CONFIG.LOGS_BUCKET).download(objectKey);
    if (existing.data) {
      const buf = Buffer.from(await existing.data.arrayBuffer());
      current = buf.toString('utf-8');
    }

    const updated = current + JSON.stringify(entry) + '\n';

    const up = await baseSupabase.storage.from(CONFIG.LOGS_BUCKET).upload(objectKey, updated, {
      upsert: true,
      contentType: 'application/x-ndjson',
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

/* -------------------------------------------------------
   Auth middleware
------------------------------------------------------- */
const requireAuth = async (req, res, next) => {
  try {
    const authHeader = req.headers.authorization || '';
    const token = authHeader.startsWith('Bearer ') ? authHeader.slice(7) : null;
    if (!token) return res.status(401).json({ error: 'Unauthorized: missing bearer token' });

    const authClient = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_ANON_KEY, {
      auth: { persistSession: false, autoRefreshToken: false },
    });
    const { data, error } = await authClient.auth.getUser(token);
    if (error || !data?.user) return res.status(401).json({ error: 'Unauthorized: invalid token' });

    req.user = data.user;
    req.supabase = getSupabaseForToken(token);
    next();
  } catch (e) {
    console.error('requireAuth error:', e);
    res.status(401).json({ error: 'Unauthorized' });
  }
};

const requireAdmin = async (req, res, next) => {
  try {
    const { data, error } = await req.supabase
      .from('profiles')
      .select('role')
      .eq('id', req.user.id)
      .single();

    if (error) {
      console.error('Role check error:', error.message);
      return res.status(403).json({ error: 'Forbidden: role check failed' });
    }
    if (data?.role !== 'admin') {
      return res.status(403).json({ error: 'Forbidden: admin only' });
    }
    next();
  } catch (e) {
    console.error('requireAdmin error:', e);
    res.status(403).json({ error: 'Forbidden' });
  }
};

/* -------------------------------------------------------
   Plan middleware
------------------------------------------------------- */
const attachUserPlanAndFile = async (req, res, next) => {
  try {
    if (!req.user || !req.supabase) {
      req.userPlan = 'anon';
      req.fileKey = CONFIG.EXCEL_FILE_KEY;
      return next();
    }
    const { data, error } = await req.supabase
      .from('profiles')
      .select('plan,user_file_key,role')
      .eq('id', req.user.id)
      .single();
    if (error) {
      console.error('Plan lookup error:', error.message);
      req.userPlan = 'free';
      req.fileKey = CONFIG.EXCEL_FILE_KEY;
      return next();
    }
    req.userPlan = data?.plan || 'free';
    req.userRole = data?.role || 'user';
    req.fileKey =
      data?.user_file_key ||
      `${CONFIG.USER_FILES_PREFIX}/${req.user.id}/uploaded.xlsx`;
    next();
  } catch (e) {
    console.error('attachUserPlanAndFile error:', e);
    req.userPlan = 'free';
    req.fileKey = CONFIG.EXCEL_FILE_KEY;
    next();
  }
};



const requireAdsWatched = (req, res, next) => {
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
   Excel storage interactions (plan-aware)
   - Free plan: denied access
   - Paid plan: can get their own file URL
   - Superadmin: can get master file URL
------------------------------------------------------- */
app.get('/excel/public', requireAuth, async (req, res) => {
  try {
    // Lookup profile + plan
    const { data: profile, error: profileErr } = await req.supabase
      .from('profiles')
      .select('role, plan')
      .eq('id', req.user.id)
      .single();
    if (profileErr || !profile) {
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const role = profile.role;
    const plan = profile.plan || 'free';

    // Block free plan entirely
    if (plan === 'free') {
      return res.status(403).json({ error: 'Public file access is available only on premium plans' });
    }

    // Superadmin gets master file, others get their own
    const key = role === 'superadmin' ? CONFIG.EXCEL_FILE_KEY : req.fileKey;

    const { data, error } = baseSupabase.storage
      .from(CONFIG.EXCEL_BUCKET)
      .getPublicUrl(key);

    if (error) {
      console.error("Public URL generation failed:", error.message);
      return res.status(500).json({ error: error.message });
    }

    if (!data?.publicUrl) {
      console.error("File not found in bucket:", key);
      return res.status(404).json({ error: 'File not found in bucket' });
    }

    const urlWithCacheBust = `${data.publicUrl}?t=${Date.now()}`;
    console.log("Public Excel file URL:", urlWithCacheBust);

    res.json({ url: urlWithCacheBust });
  } catch (e) {
    console.error('/excel/public error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});



/* -------------------------------------------------------
   Download/export — totally premium
   - Free plan: denied access (no ads option)
   - Paid plan: direct download
   - Superadmin (future): unrestricted
------------------------------------------------------- */
app.get('/excel/download', requireAuth, attachUserContext, requirePremium, attachUserPlanAndFile, async (req, res) => {
   try {
    // Lookup plan
    const { data: profile, error: profileErr } = await req.supabase
      .from('profiles')
      .select('role, plan')
      .eq('id', req.user.id)
      .single();

    if (profileErr || !profile) {
      console.error("Profile lookup failed:", profileErr?.message || "No profile found");
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const role = profile.role;
    const plan = profile.plan || 'free';
    console.log('Download request by', req.user.id, 'role:', role, 'plan:', plan);

    // Block free plan entirely
    if (plan === 'free') {
      return res.status(403).json({ error: 'Downloads are available only on premium plans' });
    }

    // Default: user workbook; superadmin can override to master
    const isSuperadmin = role === 'superadmin';
    const key = isSuperadmin ? CONFIG.EXCEL_FILE_KEY : req.fileKey;

    const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
    if (error || !buffer) {
      console.error("Download failed:", error?.message || "No buffer returned");
      return res.status(404).json({ error: "Workbook not found" });
    }

    res.setHeader('Content-Disposition', `attachment; filename="${key.split('/').pop()}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Length', buffer.length);

    console.log(`Workbook downloaded successfully by ${plan} user (${role})`);
    res.status(200).send(buffer);
  } catch (e) {
    console.error('/excel/download error:', e);
    res.status(500).json({ error: e.message });
  }
});


/* -------------------------------------------------------
   List sheet names (plan-aware)
   - Free plan: must watch 2 ads before listing sheets
   - Paid plan: list own workbook sheets directly
   - Superadmin (future): can also list master file sheets
------------------------------------------------------- */
app.get('/excel/sheets', requireAuth, attachUserContext, requireAdsSoft, attachUserPlanAndFile, async (req, res) => { try {
    // Lookup profile + plan
    const { data: profile, error: profileErr } = await req.supabase
      .from('profiles')
      .select('role, plan')
      .eq('id', req.user.id)
      .single();
    if (profileErr || !profile) {
      console.error("Profile lookup failed:", profileErr?.message || "No profile found");
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const role = profile.role;
    const plan = profile.plan || 'free';
    console.log('Sheets request by', req.user.id, 'role:', role, 'plan:', plan);

    // Free plan enforcement: require ads watched
    if (plan === 'free') {
      const adsWatched = await checkAdsWatched(req.user.id, 2); // implement helper
      if (!adsWatched) {
        return res.status(403).json({ error: 'Free plan requires watching 2 ads before listing sheets' });
      }
    }

    // Default: user workbook
    const key = req.fileKey;
    let sheets = [];

    const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
    if (!error && buffer) {
      const workbook = await loadWorkbook(buffer);
      sheets = workbook.worksheets.map(ws => ws.name);
    } else {
      console.warn(`No workbook found for key "${key}"`);
    }

    // Only superadmin can also include master file sheets
    if (role === 'superadmin') {
      const { buffer: masterBuffer, error: masterError } = await getBufferFromStorage(
        baseSupabase,
        CONFIG.EXCEL_BUCKET,
        CONFIG.EXCEL_FILE_KEY
      );
      if (!masterError && masterBuffer) {
        const masterWorkbook = await loadWorkbook(masterBuffer);
        const masterSheets = masterWorkbook.worksheets.map(ws => ws.name);
        sheets = [...new Set([...sheets, ...masterSheets])]; // merge without duplicates
      }
    }

    res.json({ sheets });
  } catch (e) {
    console.error('/excel/sheets error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});



/* -------------------------------------------------------
   Add sheet (plan-aware)
   - Free plan: must watch 2 ads before adding
   - Paid plan: add allowed without ads
   - Superadmin (future): add to master file
------------------------------------------------------- */
app.post('/excel/delete-sheet', requireAuth, attachUserContext, requireAdsSoft, attachUserPlanAndFile, async (req, res) => {
  const { name } = req.body;
  if (!name) {
    return res.status(400).json({ error: 'Sheet name is required' });
  }

  try {
    // Lookup profile + plan
    const { data: profile, error: profileErr } = await req.supabase
      .from('profiles')
      .select('role, plan')
      .eq('id', req.user.id)
      .single();
    if (profileErr || !profile) {
      console.error("Profile lookup failed:", profileErr?.message || "No profile found");
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const role = profile.role;
    const plan = profile.plan || 'free';
    console.log('Add sheet request by', req.user.id, 'role:', role, 'plan:', plan);

    // Free plan enforcement: require ads watched
    if (plan === 'free') {
      const adsWatched = await checkAdsWatched(req.user.id, 2); // implement helper
      if (!adsWatched) {
        return res.status(403).json({ error: 'Free plan requires watching 2 ads before adding sheets' });
      }
    }

    // Default: user workbook; superadmin can override to master
    const isSuperadmin = role === 'superadmin';
    const key = isSuperadmin ? CONFIG.EXCEL_FILE_KEY : req.fileKey;

    // Try to download existing workbook
    const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
    let workbook;
    if (error || !buffer) {
      console.warn(`No existing workbook for key "${key}", creating new one`);
      workbook = new ExcelJS.Workbook();
    } else {
      workbook = await loadWorkbook(buffer);
    }

    // Prevent duplicates
    if (workbook.getWorksheet(name)) {
      return res.status(400).json({ error: 'Sheet already exists' });
    }

    // Add new sheet
    const ws = workbook.addWorksheet(name);
    ws.getCell('A1').value = 'New sheet created';

    // Save workbook back to storage
    const outBuffer = await workbook.xlsx.writeBuffer();
    const { error: uploadError } = await updateBufferToStorage(
      baseSupabase,
      CONFIG.EXCEL_BUCKET,
      key,
      outBuffer,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    if (uploadError) {
      console.error("Storage upload failed:", uploadError.message);
      return res.status(500).json({ error: uploadError.message });
    }

    // Collect updated sheet names for frontend sync
    const sheetNames = workbook.worksheets.map(ws => ws.name);

    // Optional: log add-sheet action
    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      action: 'add_sheet',
      sheet_name: name,
      metadata: { role, plan }
    });

    console.log(`Sheet "${name}" added successfully to ${isSuperadmin ? 'master' : 'user'} workbook`);

    res.json({
      success: true,
      sheet: name,
      sheets: sheetNames,
    });
  } catch (e) {
    console.error('/excel/add-sheet error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});



/* -------------------------------------------------------
   Delete sheet (plan-aware)
   - Free plan: must watch 2 ads before deleting
   - Paid plan: delete allowed without ads
   - Superadmin (future): delete from master file
------------------------------------------------------- */
app.post('/excel/delete-sheet', requireAuth, attachUserContext, requireAdsSoft, attachUserPlanAndFile, async (req, res) => {
  const { name } = req.body;
  if (!name) {
    return res.status(400).json({ error: 'Sheet name required' });
  }

  try {
    // Lookup profile + plan
    const { data: profile, error: profileErr } = await req.supabase
      .from('profiles')
      .select('role, plan')
      .eq('id', req.user.id)
      .single();
    if (profileErr || !profile) {
      console.error("Profile lookup failed:", profileErr?.message || "No profile found");
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const role = profile.role;
    const plan = profile.plan || 'free';
    console.log('Delete sheet request by', req.user.id, 'role:', role, 'plan:', plan);

    // Free plan enforcement: require ads watched
    if (plan === 'free') {
      const adsWatched = await checkAdsWatched(req.user.id, 2); // implement helper
      if (!adsWatched) {
        return res.status(403).json({ error: 'Free plan requires watching 2 ads before deleting sheets' });
      }
    }

    // Default: user workbook; superadmin can override to master
    const isSuperadmin = role === 'superadmin';
    const key = isSuperadmin ? CONFIG.EXCEL_FILE_KEY : req.fileKey;

    // Download workbook
    const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
    if (error || !buffer) {
      console.error("Download failed:", error?.message || "No buffer returned");
      return res.status(404).json({ error: "Workbook not found" });
    }

    const workbook = await loadWorkbook(buffer);
    const ws = workbook.getWorksheet(name);
    if (!ws) {
      return res.status(400).json({ error: 'Sheet not found' });
    }

    workbook.removeWorksheet(ws.id);

    // Save workbook back to storage
    const outBuffer = await workbook.xlsx.writeBuffer();
    const { error: uploadError } = await updateBufferToStorage(
      baseSupabase,
      CONFIG.EXCEL_BUCKET,
      key,
      outBuffer,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    if (uploadError) {
      console.error("Storage upload failed:", uploadError.message);
      return res.status(500).json({ error: uploadError.message });
    }

    // Optional: log delete action
    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      action: 'delete_sheet',
      sheet_name: name,
      metadata: { role, plan }
    });

    console.log(`Sheet "${name}" deleted successfully from ${isSuperadmin ? 'master' : 'user'} workbook`);

    res.json({ success: true, deleted: name });
  } catch (e) {
    console.error('/excel/delete-sheet error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});


/* -------------------------------------------------------
   Get single cell value (plan-aware)
   - Free plan: must watch 2 ads before querying
   - Paid plan: query allowed without ads
   - Superadmin (future): can query master workbook
------------------------------------------------------- */
app.get('/excel/get', requireAuth, attachUserContext, requireAdsSoft, attachUserPlanAndFile, async (req, res) => {
  const { sheet, cell, scope } = req.query; 
  if (!sheet || !cell) {
    return res.status(400).json({ error: 'Sheet and cell are required' });
  }

  try {
    // Lookup profile + plan
    const { data: profile, error: profileErr } = await req.supabase
      .from('profiles')
      .select('role, plan')
      .eq('id', req.user.id)
      .single();
    if (profileErr || !profile) {
      console.error("Profile lookup failed:", profileErr?.message || "No profile found");
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const role = profile.role;
    const plan = profile.plan || 'free';
    console.log('Get cell request by', req.user.id, 'role:', role, 'plan:', plan);

    // Free plan enforcement: require ads watched
    if (plan === 'free') {
      const adsWatched = await checkAdsWatched(req.user.id, 2); // implement helper
      if (!adsWatched) {
        return res.status(403).json({ error: 'Free plan requires watching 2 ads before querying cells' });
      }
    }

    // Default: user workbook; superadmin can override with ?scope=master
    const isSuperadmin = role === 'superadmin';
    const key = isSuperadmin && scope === 'master' 
      ? CONFIG.EXCEL_FILE_KEY 
      : req.fileKey;

    const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
    if (error || !buffer) {
      console.error("Download failed:", error?.message || "No buffer returned");
      return res.status(404).json({ error: "Workbook not found" });
    }

    const workbook = await loadWorkbook(buffer);
    const ws = workbook.getWorksheet(sheet);
    if (!ws) {
      return res.status(400).json({ error: `Sheet "${sheet}" not found` });
    }

    const value = ws.getCell(cell).value;

    // Optional: log cell query in audit
    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      action: 'get_cell',
      sheet_name: sheet,
      metadata: { cell, value }
    });

    res.json({ sheet, cell, value, source: key });
  } catch (e) {
    console.error('/excel/get error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});



/* -------------------------------------------------------
   Save only changed cells (batch commit, plan-aware)
   - Free plan: must watch 2 ads before saving + row limit safeguard
   - Paid plan: unlimited saves, no ads required
------------------------------------------------------- */
app.post('/excel/save-all', requireAuth, attachUserContext, requireAdsSoft, attachUserPlanAndFile, async (req, res) => {
   const { sheet, data } = req.body;
  if (!sheet || !data) {
    return res.status(400).json({ error: 'Sheet and data are required' });
  }

  try {
    // Lookup user plan
    const { data: profile, error: profileErr } = await req.supabase
      .from('profiles')
      .select('plan')
      .eq('id', req.user.id)
      .single();

    if (profileErr || !profile) {
      console.error("Profile lookup failed:", profileErr?.message || "No profile found");
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const plan = profile.plan || 'free';
    console.log('Save-all request by', req.user.id, 'plan:', plan);

    // Free plan enforcement
    if (plan === 'free') {
      // Require ads watched
      const adsWatched = await checkAdsWatched(req.user.id, 2); // implement helper
      if (!adsWatched) {
        return res.status(403).json({ error: 'Free plan requires watching 2 ads before saving' });
      }

      // Optional safeguard: limit rows per save
      const rowsCount = Array.isArray(data) ? data.length : 0;
      if (rowsCount > CONFIG.FREE_LIMITS.maxRowsSaveAll) {
        console.warn(`Free plan limit exceeded: ${rowsCount} rows`);
        return res.status(403).json({
          error: `Free plan limit: max ${CONFIG.FREE_LIMITS.maxRowsSaveAll} rows per save`
        });
      }
    }

    const key = req.fileKey;
    let workbook;
    let ws;

    // Try to download existing workbook
    const { buffer, error } = await getBufferFromStorage(req.supabase, CONFIG.EXCEL_BUCKET, key);
    if (error || !buffer) {
      console.warn("No existing file found, creating new workbook:", error?.message);
      workbook = new ExcelJS.Workbook();
      ws = workbook.addWorksheet(sheet);
    } else {
      workbook = await loadWorkbook(buffer);
      ws = workbook.getWorksheet(sheet);
      if (!ws) {
        console.warn(`Sheet "${sheet}" not found, creating new one`);
        ws = workbook.addWorksheet(sheet);
      }
    }

    // Update cells with staged data
    data.forEach((row, rIndex) => {
      row.forEach((val, cIndex) => {
        const cell = ws.getRow(rIndex + 1).getCell(cIndex + 1);
        const oldVal = cell.value ?? null;
        if (oldVal !== val) {
          cell.value = val;
        }
      });
    });

    // Save workbook back to storage
    const outBuffer = await workbook.xlsx.writeBuffer();
    const { error: uploadError } = await updateBufferToStorage(
      req.supabase,
      CONFIG.EXCEL_BUCKET,
      key,
      outBuffer,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    if (uploadError) {
      console.error("Storage upload failed:", uploadError.message);
      return res.status(500).json({ error: uploadError.message });
    }

    // Optional verification
    try {
      const { buffer: verifyBuffer, error: verifyErr } = await getBufferFromStorage(
        req.supabase,
        CONFIG.EXCEL_BUCKET,
        key
      );
      if (!verifyErr && verifyBuffer) {
        const verifyWb = await loadWorkbook(verifyBuffer);
        const vws = verifyWb.getWorksheet(sheet);
        if (vws) {
          console.log(`Verification: sheet "${sheet}" now has ${vws.rowCount} rows`);
        }
      }
    } catch (verifyEx) {
      console.error("Verification step failed:", verifyEx.message);
    }

    // Optional: log save-all action
    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      action: 'save_all',
      sheet_name: sheet,
      metadata: { rows: Array.isArray(data) ? data.length : 0 }
    });

    res.json({ success: true });
  } catch (e) {
    console.error('/excel/save-all error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});



/* -------------------------------------------------------
   Admin-only cell updates with audit (master file)
   - Free plan: must watch 2 ads before updates
   - Paid plan: updates allowed without ads
------------------------------------------------------- */

app.post('/excel/update', requireAuth, attachUserContext, requireAdsSoft, async (req, res) => {
   const { changes } = req.body;
  if (!Array.isArray(changes) || changes.length === 0) {
    return res.status(400).json({ error: 'No changes provided' });
  }

  try {
    // Lookup profile + plan
    const { data: profile, error: profileError } = await req.supabase
      .from('profiles')
      .select('id, email, plan')
      .eq('id', req.user.id)
      .single();
    if (profileError || !profile) {
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const plan = profile.plan || 'free';
    console.log('Update request by', req.user.id, 'plan:', plan);

    // Free plan enforcement: require ads watched
    if (plan === 'free') {
      const adsWatched = await checkAdsWatched(req.user.id, 2); // implement helper
      if (!adsWatched) {
        return res.status(403).json({ error: 'Free plan requires watching 2 ads before updating' });
      }
    }

    // Each user has their own workbook file
    const userFileKey = `${req.user.id}.xlsx`;

    const { buffer, error: downloadError } = await getBufferFromStorage(
      req.supabase,
      CONFIG.EXCEL_BUCKET,
      userFileKey
    );
    let workbook;
    if (downloadError || !buffer) {
      console.warn("No existing workbook, creating new one for user:", req.user.id);
      workbook = new ExcelJS.Workbook();
    } else {
      workbook = await loadWorkbook(buffer);
    }

    // Apply changes + audit
    for (const c of changes) {
      const ws = workbook.getWorksheet(c.sheet) || workbook.addWorksheet(c.sheet);

      const oldValue = ws.getCell(c.cell).value;
      const nextVal =
        typeof c.value === 'object'
          ? JSON.stringify(c.value)
          : typeof c.value === 'number'
          ? c.value
          : String(c.value ?? '');

      ws.getCell(c.cell).value = nextVal;

      await req.supabase.from('excel_audit').insert({
        user_id: profile.id,
        action: 'update_cell',
        sheet_name: c.sheet,
        metadata: {
          cell: c.cell,
          old_value: oldValue,
          new_value: nextVal,
          email: profile.email,
        },
      });
    }

    // Save workbook back to storage under the user’s key
    const outBuffer = await workbook.xlsx.writeBuffer();
    const { error: uploadError } = await req.supabase.storage
      .from(CONFIG.EXCEL_BUCKET)
      .upload(userFileKey, outBuffer, { upsert: true });
    if (uploadError) {
      return res.status(500).json({ error: `Storage upload failed: ${uploadError.message}` });
    }

    res.json({ changes_applied: changes.length });
  } catch (e) {
    console.error('/excel/update error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});


/* -------------------------------------------------------
   Preview full sheet (2D array, plan-aware)
------------------------------------------------------- */
app.get('/excel/preview', attachUserPlanAndFile, async (req, res) => {
   const { sheet } = req.query;
  if (!sheet) {
    return res.status(400).json({ error: 'Sheet is required' });
  }

  try {
    const key = req.user ? req.fileKey : CONFIG.EXCEL_FILE_KEY;
    const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
    if (error) {
      console.error("Download failed:", error.message);
      return res.status(404).json({ error: error.message });
    }

    const workbook = await loadWorkbook(buffer);
    const ws = workbook.getWorksheet(sheet);
    if (!ws) {
      return res.status(400).json({ error: 'Sheet not found' });
    }

    const preview = [];
    const maxRow = ws.actualRowCount || ws.rowCount;
    const maxCol = ws.actualColumnCount || (ws.columns ? ws.columns.length : 0);

    for (let r = 1; r <= maxRow; r++) {
      const row = [];
      for (let c = 1; c <= maxCol; c++) {
        row.push(ws.getRow(r).getCell(c).value);
      }
      preview.push(row);
    }

    console.log(`Previewed sheet "${sheet}" with ${maxRow} rows and ${maxCol} cols`);
    res.json({ sheet, preview, rows: maxRow, cols: maxCol });
  } catch (e) {
    console.error('/excel/preview error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});

/* -------------------------------------------------------
   Metadata (plan-aware: show meta for user file if exists)
------------------------------------------------------- */
app.get('/excel/meta', attachUserPlanAndFile, async (req, res) => {
   try {
    const key = req.user ? req.fileKey : CONFIG.EXCEL_FILE_KEY;
    const { data, error } = await baseSupabase.storage
      .from(CONFIG.EXCEL_BUCKET)
      .list('', { limit: 100 });
    if (error) {
      console.error("Storage list failed:", error.message);
      return res.status(404).json({ error: error.message });
    }

    const file = data.find((f) => f.name === key) || data.find((f) => f.name === CONFIG.EXCEL_FILE_KEY);
    if (!file) {
      console.error("File not found:", key);
      return res.status(404).json({ error: 'File not found' });
    }

    console.log(`Metadata retrieved for file "${file.name}"`);
    res.json({
      name: file.name,
      size: file.size,
      last_modified: file.updated_at,
    });
  } catch (e) {
    console.error('/excel/meta error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});

/* -------------------------------------------------------
   Export as CSV (full sheet, escaped) — ad-gated, plan-aware
   - Free plan: must watch 2 ads + max 3 exports/day
   - Paid plan: unlimited exports, no ads required
------------------------------------------------------- */
app.get('/excel/export/csv', requireAuth, attachUserContext, requireAdsSoft, attachUserPlanAndFile, async (req, res) => {
   const { sheet } = req.query;
  if (!sheet) {
    return res.status(400).json({ error: 'Sheet is required' });
  }

  try {
    // Lookup user plan
    const { data: profile, error: profileErr } = await req.supabase
      .from('profiles')
      .select('plan')
      .eq('id', req.user.id)
      .single();

    if (profileErr || !profile) {
      console.error("Profile lookup failed:", profileErr?.message || "No profile found");
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const plan = profile.plan || 'free';
    console.log('CSV export request by', req.user.id, 'plan:', plan);

    // Free plan enforcement
    if (plan === 'free') {
      // Require ads watched
      const adsWatched = await checkAdsWatched(req.user.id, 2); // implement helper
      if (!adsWatched) {
        return res.status(403).json({ error: 'Free plan requires watching 2 ads before export' });
      }

      // Enforce daily export limit (3/day)
      const today = new Date().toISOString().split('T')[0]; // YYYY-MM-DD
      const { count, error: countErr } = await req.supabase
        .from('excel_audit')
        .select('*', { count: 'exact', head: true })
        .eq('user_id', req.user.id)
        .eq('action', 'export_csv')
        .gte('created_at', `${today}T00:00:00`)
        .lte('created_at', `${today}T23:59:59`);

      if (countErr) {
        console.error("Export count check failed:", countErr.message);
        return res.status(500).json({ error: 'Export count check failed' });
      }

      if (count >= 3) {
        return res.status(403).json({ error: 'Free plan allows only 3 CSV exports per day' });
      }
    }

    const key = req.user ? req.fileKey : CONFIG.EXCEL_FILE_KEY;
    const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
    if (error) {
      console.error("Download failed:", error.message);
      return res.status(404).json({ error: error.message });
    }

    const workbook = await loadWorkbook(buffer);
    const ws = workbook.getWorksheet(sheet);
    if (!ws) {
      return res.status(400).json({ error: 'Sheet not found' });
    }

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

    console.log(`Exported sheet "${sheet}" as CSV with ${maxRow} rows`);

    // Log export in audit table
    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      action: 'export_csv',
      details: { sheet },
    });

    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${sheet}.csv"`);
    res.send(csv);
  } catch (e) {
    console.error('/excel/export/csv error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});


/* -------------------------------------------------------
   Export as PDF (single sheet) — ad-gated, plan-aware
   - Free plan: must watch 2 ads before unlimited exports
   - Paid plan: unlimited exports, no ads required
------------------------------------------------------- */
app.get('/excel/export/pdf', requireAuth, attachUserContext, requireAdsSoft, attachUserPlanAndFile, async (req, res) => {
  const { sheet } = req.query;
  if (!sheet) {
    return res.status(400).json({ error: 'Sheet is required' });
  }

  try {
    // Lookup user plan
    const { data: profile, error: profileErr } = await req.supabase
      .from('profiles')
      .select('plan')
      .eq('id', req.user.id)
      .single();

    if (profileErr || !profile) {
      console.error("Profile lookup failed:", profileErr?.message || "No profile found");
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const plan = profile.plan || 'free';
    console.log('PDF export request by', req.user.id, 'plan:', plan);

    // Free plan enforcement: require ads watched once before unlimited exports
    if (plan === 'free') {
      const adsWatched = await checkAdsWatched(req.user.id, 2); // implement helper to verify ads
      if (!adsWatched) {
        return res.status(403).json({ error: 'Free plan requires watching 2 ads before exporting' });
      }
    }

    const key = req.user ? req.fileKey : CONFIG.EXCEL_FILE_KEY;
    const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
    if (error) {
      console.error("Download failed:", error.message);
      return res.status(404).json({ error: error.message });
    }

    const workbook = await loadWorkbook(buffer);
    const ws = workbook.getWorksheet(sheet);
    if (!ws) {
      return res.status(400).json({ error: 'Sheet not found' });
    }

    const PDFDocument = (await import('pdfkit')).default;
    const doc = new PDFDocument({ size: 'A4', margin: 36 });

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${sheet}.pdf"`);
    doc.pipe(res);

    // Title
    doc.fontSize(16).fillColor('#111111').text(`Sheet: ${sheet}`, { align: 'left' });
    doc.moveDown(0.5);

    // Table styles
    const cellPaddingX = 6;
    const cellPaddingY = 4;
    const rowHeight = 18;
    const colWidth = 100;
    const maxRow = ws.actualRowCount || ws.rowCount;
    const maxCol = ws.actualColumnCount || (ws.columns ? ws.columns.length : 0);

    const startX = doc.page.margins.left;
    let cursorY = doc.y;

    // Header background
    doc.save();
    doc.rect(startX, cursorY, colWidth * maxCol, rowHeight).fill('#f3f4f6');
    doc.restore();

    // Header text + borders
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

    // Body rows
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

    console.log(`Exported sheet "${sheet}" as PDF with ${maxRow} rows`);

    // Optional: log export in audit table
    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      action: 'export_pdf_single',
      details: { sheet },
    });

    doc.end();
  } catch (e) {
    console.error('/excel/export/pdf error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});


/* -------------------------------------------------------
   Export multiple sheets to one PDF — ad-gated, plan-aware
   - Free plan: must watch 2 ads + sheet limit + max 3 exports/day
   - Paid plan: full export, no ads required
------------------------------------------------------- */
app.get('/excel/export/pdf-multi', requireAuth, attachUserContext, requireAdsSoft, attachUserPlanAndFile, async (req, res) => {
  const { sheets } = req.query; // comma-separated list
  try {
    // Lookup user plan
    const { data: profile, error: profileErr } = await req.supabase
      .from('profiles')
      .select('plan')
      .eq('id', req.user.id)
      .single();

    if (profileErr || !profile) {
      console.error("Profile lookup failed:", profileErr?.message || "No profile found");
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const plan = profile.plan || 'free';
    console.log('PDF export request by', req.user.id, 'plan:', plan);

    // Free plan enforcement
    if (plan === 'free') {
      // Require ads watched
      const adsWatched = await checkAdsWatched(req.user.id, 2); // implement this helper
      if (!adsWatched) {
        return res.status(403).json({ error: 'Free plan requires watching 2 ads before export' });
      }

      // Enforce daily export limit (3/day)
      const today = new Date().toISOString().split('T')[0]; // YYYY-MM-DD
      const { count, error: countErr } = await req.supabase
        .from('excel_audit')
        .select('*', { count: 'exact', head: true })
        .eq('user_id', req.user.id)
        .eq('action', 'export_pdf_multi')
        .gte('created_at', `${today}T00:00:00`)
        .lte('created_at', `${today}T23:59:59`);

      if (countErr) {
        console.error("Export count check failed:", countErr.message);
        return res.status(500).json({ error: 'Export count check failed' });
      }

      if (count >= 3) {
        return res.status(403).json({ error: 'Free plan allows only 3 PDF exports per day' });
      }
    }

    const key = req.user ? req.fileKey : CONFIG.EXCEL_FILE_KEY;
    const { buffer, error } = await getBufferFromStorage(baseSupabase, CONFIG.EXCEL_BUCKET, key);
    if (error) {
      console.error("Download failed:", error.message);
      return res.status(404).json({ error: error.message });
    }

    const workbook = await loadWorkbook(buffer);

    let sheetList = sheets
      ? sheets.split(',').map(s => s.trim()).filter(Boolean)
      : workbook.worksheets.map(ws => ws.name);

    // Free plan sheet limit
    if (plan === 'free') {
      sheetList = sheetList.slice(0, CONFIG.FREE_LIMITS.maxSheetsInMultiPDF);
      console.warn(`Free plan: limiting to ${sheetList.length} sheets`);
    }

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

      // Header
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

      // Rows
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
      if (!ws) {
        console.warn(`Sheet "${name}" not found, skipping`);
        return;
      }
      if (idx > 0) doc.addPage();
      drawTable(ws, name);
    });

    console.log(`Exported ${sheetList.length} sheets to multi-PDF`);

    // Log export in audit table
    await req.supabase.from('excel_audit').insert({
      user_id: req.user.id,
      action: 'export_pdf_multi',
      details: { sheets: sheetList },
    });

    doc.end();
  } catch (e) {
    console.error('/excel/export/pdf-multi error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});



/* -------------------------------------------------------
   Audit retrieval (RLS-aware, plan-aware)
   - Free plan: denied access (admins included)
   - Paid plan: only own audit logs
   - Superadmin (future role): all logs
------------------------------------------------------- */
app.get('/excel/audit', requireAuth, requirePremium, attachUserContext, async (req, res) => {
  try {
    // Lookup user role + plan
    const { data: profile, error: profileErr } = await req.supabase
      .from('profiles')
      .select('role, plan')
      .eq('id', req.user.id)
      .single();

    if (profileErr || !profile) {
      console.error("Profile lookup failed:", profileErr?.message || "No profile found");
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const role = profile.role;
    const isFree = profile.plan === 'free';
    console.log('Audit request by', req.user.id, 'role:', role, 'plan:', profile.plan);

    // Block free plan entirely
    if (isFree) {
      return res.status(403).json({ error: 'Audit logs are not available on the free plan' });
    }

    // Support optional pagination
    const limit = parseInt(req.query.limit, 10) || 100;
    const offset = parseInt(req.query.offset, 10) || 0;

    let query = req.supabase
      .from('excel_audit')
      .select('*')
      .order('created_at', { ascending: false })
      .range(offset, offset + limit - 1);

    // Only superadmin can see all logs
    if (role !== 'superadmin') {
      query = query.eq('user_id', req.user.id);
    }

    const { data, error } = await query;
    if (error) {
      console.error("Audit retrieval failed:", error.message);
      return res.status(500).json({ error: error.message });
    }

    res.json({ logs: data, count: data?.length || 0 });
  } catch (e) {
    console.error('/excel/audit error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});





/* -------------------------------------------------------
   File upload (Excel/CSV/PDF) with preview + profile update
   Plan-aware:
   - Free users: only 1 XLSX file allowed
   - CSV/PDF: premium only
------------------------------------------------------- */
const upload = multer({ storage: multer.memoryStorage() });

app.post('/excel/upload', requireAuth, attachUserContext, requireAdsSoft, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      console.error("No file received");
      return res.status(400).json({ error: 'No file uploaded' });
    }

    let buffer = req.file.buffer; // allow reassignment if we rewrite
    const filename = req.file.originalname;
    console.log("Received upload:", filename);

    // Lookup user plan
    const { data: profile, error: profileErr } = await req.supabase
      .from('profiles')
      .select('plan, user_file_key')
      .eq('id', req.user.id)
      .single();

    if (profileErr || !profile) {
      console.error("Profile lookup failed:", profileErr?.message || "No profile found");
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const isFree = profile.plan === 'free';
    const alreadyHasFile = !!profile.user_file_key;

    // Enforce plan rules
    if (filename.toLowerCase().endsWith('.csv') || filename.toLowerCase().endsWith('.pdf')) {
      if (isFree) {
        return res.status(403).json({ error: 'CSV and PDF uploads require a premium plan' });
      }
    }
    if (filename.toLowerCase().endsWith('.xlsx')) {
      if (isFree && alreadyHasFile) {
        return res.status(403).json({ error: 'Free plan allows only one Excel file upload' });
      }
    }

    let preview = [];
    let sheetNames = [];

    try {
      if (filename.toLowerCase().endsWith('.xlsx')) {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);

        sheetNames = workbook.worksheets.map(ws => ws.name);

        // auto-create default sheet if none exist
        if (sheetNames.length === 0) {
          const defaultSheet = workbook.addWorksheet('Sheet1');
          defaultSheet.getCell('A1').value = 'New sheet created';
          sheetNames.push('Sheet1');
          buffer = await workbook.xlsx.writeBuffer();
        }

        // preview first sheet
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
      } else if (filename.toLowerCase().endsWith('.csv')) {
        // parse CSV rows
        const rows = csvParse(buffer.toString(), { columns: false });
        preview = rows.slice(0, 20); // preview first 20 rows

        // convert into workbook with default sheet
        const workbook = new ExcelJS.Workbook();
        const ws = workbook.addWorksheet('Sheet1');
        rows.forEach((row, rIndex) => {
          row.forEach((val, cIndex) => {
            ws.getRow(rIndex + 1).getCell(cIndex + 1).value = val;
          });
        });
        sheetNames = ['Sheet1'];
        buffer = await workbook.xlsx.writeBuffer();
      } else if (filename.toLowerCase().endsWith('.pdf')) {
        const data = await pdfParse(buffer);
        preview = data.text.split('\n').slice(0, 20);
        sheetNames = [filename.replace(/\.[^/.]+$/, '')];
        // PDFs are stored as-is, not converted to workbook
      } else {
        return res.status(400).json({ error: 'Unsupported file type' });
      }
    } catch (parseErr) {
      console.error("File parse failed:", parseErr.message);
      return res.status(400).json({ error: "File parse failed: " + parseErr.message });
    }

    const key = `${CONFIG.USER_FILES_PREFIX}/${req.user.id}/${filename}`;
    const { error: uploadError } = await putBufferToStorage(
      req.supabase,
      CONFIG.EXCEL_BUCKET,
      key,
      buffer,
      req.file.mimetype
    );
    if (uploadError) {
      console.error("Storage upload failed:", uploadError.message);
      return res.status(500).json({ error: "Storage upload failed: " + uploadError.message });
    }

    const { error: updateErr } = await req.supabase
      .from('profiles')
      .update({ user_file_key: key })
      .eq('id', req.user.id);
    if (updateErr) {
      console.error("Profile update failed:", updateErr.message);
      return res.status(500).json({ error: "Profile update failed: " + updateErr.message });
    }

    console.log("Upload success:", key);
    return res.json({
      success: true,
      fileKey: key,
      fileName: filename,
      sheetNames,
      preview,
    });
  } catch (e) {
    console.error('/excel/upload error:', e);
    return res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});



/* -------------------------------------------------------
   Convert-to-App: upload user Excel, bind to profile
   Plan-aware:
   - Free users: only 1 XLSX file allowed
   - CSV/PDF: premium only
------------------------------------------------------- */
app.post('/excel/convert', requireAuth, attachUserContext, requireAdsSoft, async (req, res) => {
  try {
    const { fileBase64, fileName, fileUrl } = req.body;
    let buffer;

    // Input validation
    if (fileBase64) {
      buffer = Buffer.from(fileBase64, 'base64');
    } else if (fileUrl) {
      const resp = await fetch(fileUrl);
      if (!resp.ok) {
        console.error("Failed to fetch fileUrl:", fileUrl);
        return res.status(400).json({ error: 'Failed to fetch fileUrl' });
      }
      const arr = await resp.arrayBuffer();
      buffer = Buffer.from(arr);
    } else {
      return res.status(400).json({ error: 'fileBase64 or fileUrl required' });
    }

    // Lookup user plan + existing file
    const { data: profile, error: profileErr } = await req.supabase
      .from('profiles')
      .select('plan, user_file_key')
      .eq('id', req.user.id)
      .single();

    if (profileErr || !profile) {
      console.error("Profile lookup failed:", profileErr?.message || "No profile found");
      return res.status(500).json({ error: 'Profile lookup failed' });
    }

    const isFree = profile.plan === 'free';
    const alreadyHasFile = !!profile.user_file_key;
    const safeFileName = fileName || 'uploaded.xlsx';

    // Enforce plan rules
    if (safeFileName.toLowerCase().endsWith('.csv') || safeFileName.toLowerCase().endsWith('.pdf')) {
      if (isFree) {
        return res.status(403).json({ error: 'CSV and PDF conversion requires a premium plan' });
      }
    }
    if (safeFileName.toLowerCase().endsWith('.xlsx')) {
      if (isFree && alreadyHasFile) {
        return res.status(403).json({ error: 'Free plan allows only one Excel file conversion' });
      }
    }

    // Validate Excel file
    let workbook;
    try {
      workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      // Auto-create default sheet if none exist
      if (workbook.worksheets.length === 0) {
        const ws = workbook.addWorksheet('Sheet1');
        ws.getCell('A1').value = 'New sheet created';
        buffer = await workbook.xlsx.writeBuffer();
      }
    } catch (e) {
      console.error("Invalid Excel file:", e.message);
      return res.status(400).json({ error: 'Invalid Excel file' });
    }

    // Upload to storage (per-user key)
    const key = `${CONFIG.USER_FILES_PREFIX}/${req.user.id}/${safeFileName}`;
    const { error: uploadError } = await putBufferToStorage(
      req.supabase,
      CONFIG.EXCEL_BUCKET,
      key,
      buffer,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    if (uploadError) {
      console.error("Storage upload failed:", uploadError.message);
      return res.status(500).json({ error: uploadError.message });
    }

    // Update profile with file key
    const { error: updateErr } = await req.supabase
      .from('profiles')
      .update({ user_file_key: key })
      .eq('id', req.user.id);
    if (updateErr) {
      console.error("Profile update failed:", updateErr.message);
      return res.status(500).json({ error: updateErr.message });
    }

    console.log("Convert-to-App success:", key);

    // Respond success
    res.json({
      success: true,
      fileKey: key,
      appUrl: `/app/${req.user.id}`,
      fileName: safeFileName,
      sheetNames: workbook.worksheets.map(ws => ws.name),
    });
  } catch (e) {
    console.error('/excel/convert error:', e);
    res.status(500).json({ error: `Unexpected server error: ${e.message}` });
  }
});



export { app, baseSupabase as supabase };

