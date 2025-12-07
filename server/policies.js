// server/policies.js
import { CONFIG } from './app.js';

/* -------------------------------------------------------
   Unified policy middleware
   - Centralized enforcement for plan, ads, and roles
------------------------------------------------------- */

// Attach plan + role + fileKey
export const attachUserContext = async (req, res, next) => {
  try {
    if (!req.user || !req.supabase) {
      req.userPlan = 'anon';
      req.userRole = 'anon';
      req.fileKey = CONFIG.EXCEL_FILE_KEY;
      return next();
    }

    const { data, error } = await req.supabase
      .from('profiles')
      .select('plan, role, user_file_key')
      .eq('id', req.user.id)
      .single();

    if (error || !data) {
      req.userPlan = 'free';
      req.userRole = 'user';
      req.fileKey = CONFIG.EXCEL_FILE_KEY;
      return next();
    }

    req.userPlan = data.plan || 'free';
    req.userRole = data.role || 'user';
    req.fileKey = data.user_file_key || `${CONFIG.USER_FILES_PREFIX}/${req.user.id}/uploaded.xlsx`;
    next();
  } catch (e) {
    console.error('attachUserContext error:', e);
    req.userPlan = 'free';
    req.userRole = 'user';
    req.fileKey = CONFIG.EXCEL_FILE_KEY;
    next();
  }
};

// Require premium plan
export const requirePremium = (req, res, next) => {
  if (req.userPlan !== 'paid' && req.userRole !== 'superadmin') {
    return res.status(403).json({ error: 'Premium plan required' });
  }
  next();
};

// Ad-gate (soft enforcement: if ads not found, continue)
export const requireAdsSoft = (req, res, next) => {
  const adsWatched = parseInt(req.headers['x-ads-watched'] || '0', 10);
  if (Number.isNaN(adsWatched) || adsWatched < CONFIG.ADS_REQUIRED) {
    console.warn(`Ads not satisfied: required ${CONFIG.ADS_REQUIRED}, got ${adsWatched}. Continuing anyway.`);
    // Instead of blocking, just continue
  }
  next();
};

// Role checks
export const requireSuperadmin = (req, res, next) => {
  if (req.userRole !== 'superadmin') {
    return res.status(403).json({ error: 'Superadmin only' });
  }
  next();
};
