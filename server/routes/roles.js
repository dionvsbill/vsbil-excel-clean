// routes/roles.js
import { Router } from 'express';
import { supabase } from '../app.js';

const router = Router();

// Resolve identity and role from JWT + profiles table
async function getIdentity(req) {
  const token = (req.headers.authorization || '').replace('Bearer ', '');
  if (!token) return { user: null, role: 'anonymous', can_edit: false };

  // Validate token and get user
  const { data, error } = await supabase.auth.getUser(token);
  if (error || !data?.user) return { user: null, role: 'anonymous', can_edit: false };

  const user = data.user;

  // Look up role/can_edit from profiles table
  const { data: profile, error: profileError } = await supabase
    .from('profiles')
    .select('role, can_edit, email')
    .eq('id', user.id)
    .single();

  if (profileError || !profile) {
    return { user, role: 'user', can_edit: false, email: user.email };
  }

  return {
    user,
    role: profile.role || 'user',
    can_edit: profile.can_edit || profile.role === 'admin',
    email: profile.email || user.email
  };
}

// 🔒 Middleware: require authenticated user
export async function requireAuth(req, res, next) {
  const ident = await getIdentity(req);
  if (!ident.user) return res.status(401).json({ error: 'Unauthorized' });
  req.identity = ident;
  next();
}

// 🔒 Middleware: require admin role
export async function requireAdmin(req, res, next) {
  const ident = await getIdentity(req);
  if (!ident.user) return res.status(401).json({ error: 'Unauthorized' });
  if (ident.role !== 'admin') return res.status(403).json({ error: 'Forbidden' });
  req.identity = ident;
  next();
}

// 🔒 Middleware: require edit permission (admin or editor)
export async function requireEditor(req, res, next) {
  const ident = await getIdentity(req);
  if (!ident.user) return res.status(401).json({ error: 'Unauthorized' });
  if (!ident.can_edit) return res.status(403).json({ error: 'Forbidden' });
  req.identity = ident;
  next();
}

// Self-check: who am I?
router.get('/me', requireAuth, (req, res) => {
  const ident = req.identity;
  res.json({
    id: ident.user.id,
    email: ident.email,
    role: ident.role,
    can_edit: ident.can_edit
  });
});

// Admin-only: list all users with roles and edit flags
router.get('/list', requireAdmin, async (req, res) => {
  try {
    const { data, error } = await supabase
      .from('profiles')
      .select('id, email, role, can_edit')
      .order('email', { ascending: true });

    if (error) return res.status(500).json({ error: error.message });
    res.json({ users: data });
  } catch (e) {
    console.error('/roles/list error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Admin-only: grant edit access
router.patch('/grant/:id', requireAdmin, async (req, res) => {
  const userId = req.params.id;
  try {
    const { error } = await supabase
      .from('profiles')
      .update({ can_edit: true })
      .eq('id', userId);

    if (error) return res.status(500).json({ error: error.message });
    res.json({ success: true, message: `Edit access granted to ${userId}` });
  } catch (e) {
    console.error('/roles/grant error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Admin-only: revoke edit access
router.patch('/revoke/:id', requireAdmin, async (req, res) => {
  const userId = req.params.id;
  try {
    const { error } = await supabase
      .from('profiles')
      .update({ can_edit: false })
      .eq('id', userId);

    if (error) return res.status(500).json({ error: error.message });
    res.json({ success: true, message: `Edit access revoked from ${userId}` });
  } catch (e) {
    console.error('/roles/revoke error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Guidance endpoint
router.get('/guide', (req, res) => {
  res.json({
    note: 'Roles and edit permissions are stored in the profiles table.',
    steps: [
      'Mark first account as admin in profiles table.',
      'Default new users will be role=user.',
      'Admin can grant edit by setting can_edit=true for that user in profiles.',
      'Alternatively, allow editors to self-claim via a moderated frontend that updates profiles.'
    ]
  });
});

export default router;
