// routes/authRoutes.js
const express = require('express');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const { User, AppConfig } = require('../models');
const { ensureUserSheet } = require('../services/tdeeSheetService');
const { seedDefaultExercises } = require('../services/exerciseService');

const router = express.Router();
const JWT_SECRET = process.env.JWT_SECRET || 'resfit-super-secret-jwt-key-change-in-prod';
const COOKIE_NAME = 'resfit_token';

function signUserToken(user) {
  return jwt.sign(
    { sub: user._id.toString(), username: user.username },
    JWT_SECRET,
    { expiresIn: '30d' }
  );
}

function setAuthCookie(res, token) {
  res.cookie(COOKIE_NAME, token, {
    httpOnly: true,
    sameSite: 'lax',
    secure: process.env.NODE_ENV === 'production',
    maxAge: 30 * 24 * 60 * 60 * 1000
  });
}

function requireAuth(req, res, next) {
  const bearer = req.headers.authorization?.startsWith('Bearer ')
    ? req.headers.authorization.slice(7).trim()
    : null;
  const token = req.cookies?.[COOKIE_NAME] || bearer;

  if (!token) {
    return res.status(401).json({ error: 'Authentication required. Please log in.' });
  }

  try {
    const payload = jwt.verify(token, JWT_SECRET);
    req.user = { id: payload.sub, username: payload.username };
    next();
  } catch (_) {
    return res.status(401).json({ error: 'Session expired or invalid. Please log in again.' });
  }
}

router.post('/api/auth/register', async (req, res) => {
  try {
    const username = String(req.body?.username || '').trim().toLowerCase();
    const password = String(req.body?.password || '');

    if (!/^[a-z0-9_.-]{3,30}$/.test(username)) {
      return res.status(400).json({
        error: 'Username must be 3–30 characters using only letters, numbers, _, ., or -.'
      });
    }
    if (password.length < 6) {
      return res.status(400).json({ error: 'Password must be at least 6 characters.' });
    }

    const existing = await User.findOne({ username });
    if (existing) {
      return res.status(409).json({ error: 'Username is already taken.' });
    }

    const passwordHash = await bcrypt.hash(password, 12);
    const todayIso = new Date().toLocaleDateString('sv-SE');

    const user = await User.create({
      username,
      passwordHash,
      sheetConfig: {
        startDate: todayIso,
        weightUnit: 'kg',
        energyUnit: 'Cal.',
        startingWeight: 75,
        goalWeight: 70,
        weeklyGoalRate: 0.5,
        gender: 'Male',
        measureUnit: 'Inch',
        height: 68,
        tdeeAvgWeeks: 12
      }
    });

    await Promise.all([
      AppConfig.create({ userId: user._id, measurementDayOfWeek: 'Monday' }),
      seedDefaultExercises(user._id),
      ensureUserSheet(user._id)
    ]);

    const token = signUserToken(user);
    setAuthCookie(res, token);
    res.status(201).json({
      user: { id: user._id.toString(), username: user.username },
      token
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.post('/api/auth/login', async (req, res) => {
  try {
    const username = String(req.body?.username || '').trim().toLowerCase();
    const password = String(req.body?.password || '');

    const user = await User.findOne({ username });
    if (!user || !(await bcrypt.compare(password, user.passwordHash))) {
      return res.status(401).json({ error: 'Invalid username or password.' });
    }

    await ensureUserSheet(user._id);

    const token = signUserToken(user);
    setAuthCookie(res, token);
    res.json({
      user: { id: user._id.toString(), username: user.username },
      token
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.post('/api/auth/logout', (req, res) => {
  res.clearCookie(COOKIE_NAME);
  res.json({ ok: true });
});

router.get('/api/auth/me', requireAuth, async (req, res) => {
  try {
    const user = await User.findById(req.user.id).select('_id username').lean();
    if (!user) {
      res.clearCookie(COOKIE_NAME);
      return res.status(401).json({ error: 'Account no longer exists.' });
    }
    res.json({ user: { id: user._id.toString(), username: user.username } });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

module.exports = { authRouter: router, requireAuth };