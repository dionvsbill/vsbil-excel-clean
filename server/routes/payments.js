// exc/server/routes/payments.js
import express from 'express';
import fetch from 'node-fetch';
import crypto from 'crypto';
import { supabase } from '../app.js'; // must be initialized with SERVICE_ROLE_KEY

const router = express.Router();

/**
 * Initialize a Paystack transaction
 * Expects { email, amount, mode } in body
 * amount must be in minor units (e.g. GHS 1.00 = 100)
 */
router.post('/init', async (req, res) => {
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
        callback_url: process.env.PAYSTACK_CALLBACK_URL || 'https://yourdomain.com/'
      };
    } else {
      if (!amount || isNaN(amount)) {
        return res.status(400).json({ error: 'Valid amount is required for one-time payments' });
      }
      body = {
        email,
        amount: parseInt(amount, 10),
        currency: 'GHS',
        callback_url: process.env.PAYSTACK_CALLBACK_URL || 'https://yourdomain.com/'
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

    console.log(`Initialized ${mode} payment for ${email}`, data.data?.reference);
    res.json(data);
  } catch (err) {
    console.error('Paystack init error:', err);
    res.status(500).json({ error: err.message });
  }
});

/**
 * Verify a Paystack transaction manually
 * Expects { reference, email, mode } in body
 */
router.post('/verify', express.json(), async (req, res) => {
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
      console.error('Supabase profile update error:', updateErr.message);
      return res.status(500).json({ error: 'Failed to update profile' });
    }
    if (!updated || updated.length === 0) {
      console.error('No profile row matched for email:', email);
    } else {
      console.log('Profile updated via verify:', updated);
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
      console.error('Supabase payment insert error:', payErr.message);
      return res.status(500).json({ error: 'Failed to log payment' });
    }

    // Redirect back to main page after verification
    res.redirect(process.env.PAYSTACK_SUCCESS_REDIRECT || '/');
  } catch (err) {
    console.error('Paystack verify error:', err);
    res.status(500).json({ error: err.message });
  }
});

/**
 * Paystack webhook
 */
router.post('/webhook', express.raw({ type: '*/*' }), async (req, res) => {
  const secret = process.env.PAYSTACK_SECRET_KEY;
  const signature = req.headers['x-paystack-signature'];

  const hash = crypto.createHmac('sha512', secret)
    .update(req.body)
    .digest('hex');
  if (hash !== signature) {
    return res.status(401).json({ error: 'Invalid signature' });
  }

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
      if (updateErr) console.error('Supabase plan update error:', updateErr.message);
      if (!updated || updated.length === 0) console.error('No profile row matched for email:', email);
      else console.log('Profile updated via webhook:', updated);

      const { error: payErr } = await supabase.from('payments').insert({
        email,
        amount,
        reference,
        status: 'success',
        mode: planCode === process.env.PAYSTACK_MONTHLY_PLAN ? 'monthly' : 'one-time',
        created_at: new Date().toISOString(),
      });
      if (payErr) console.error('Payments insert error:', payErr.message);

      console.log(`Payment success for ${email}, ref ${reference}, plan paid, expires ${expiresAt}`);
    } catch (err) {
      console.error('Supabase update error:', err.message);
    }
  }

  res.sendStatus(200);
});

export default router;
