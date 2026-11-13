// server.js
const express = require('express');
const cors = require('cors');
const cookieParser = require('cookie-parser');
const mongoose = require('mongoose');
const multer = require('multer');
const path = require('path');

const models = require('./models');
const { AppConfig, NutritionLog } = models;
const { authRouter, requireAuth } = require('./routes/authRoutes');
const { logToTdeeSheet, getTdeeSheetDashboard, updateSheetConfig } = require('./services/tdeeSheetService');
const {
  logExerciseSession,
  updateExerciseTargetSettings,
  getRecommendationsForWeekday,
  importExercisesFromText,
  getAllExerciseData,
  deleteExerciseLog,
  clearEntireExercise
} = require('./services/exerciseService');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// Needed so `secure` cookies work behind a reverse proxy (Render, Railway, Nginx, etc.)
app.set('trust proxy', 1);

app.use(cors({ origin: true, credentials: true }));
app.use(express.json({ limit: '5mb' }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

// Public endpoints: /api/auth/register, /login, /logout, /me
app.use(authRouter);

// Everything else under /api requires a logged-in user (sets req.user)
app.use('/api', requireAuth);
app.use(require('./routes/dataRoutes')); // clear / export / import / summary endpoints

const MONGO_URI = process.env.MONGO_URI || 'mongodb://127.0.0.1:27017/fitness_tdee_tracker';

async function connectDatabase(uri = MONGO_URI) {
  if (mongoose.connection.readyState === 0) {
    await mongoose.connect(uri);
  }
}

// Removes indexes left over from the old single-user schema (e.g. a unique index on
// exerciseName or date alone) and builds the current per-user indexes.
// syncIndexes() drops any index that is not declared in the schema.
async function syncAllIndexes() {
  for (const model of Object.values(models)) {
    try {
      const dropped = await model.syncIndexes();
      if (dropped && dropped.length) {
        console.log(`Dropped legacy indexes on ${model.collection.name}: ${dropped.join(', ')}`);
      }
    } catch (err) {
      console.error(`Index sync failed for ${model.modelName}:`, err.message);
    }
  }
}

async function getOrCreateAppConfig(userId) {
  let cfg = await AppConfig.findOne({ userId });
  if (!cfg) {
    cfg = await AppConfig.create({ userId });
  }
  return cfg;
}

// --- NUTRITION & TDEE SHEET ROUTES ---

app.get('/api/nutrition/dashboard', async (req, res) => {
  try {
    const userId = req.user.id;
    const [sheetData, config] = await Promise.all([
      getTdeeSheetDashboard(userId),
      getOrCreateAppConfig(userId)
    ]);
    res.json({ ...sheetData, config });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.put('/api/nutrition/sheet-config', async (req, res) => {
  try {
    const userId = req.user.id;
    const result = await updateSheetConfig(userId, req.body || {});
    const [dashboard, config] = await Promise.all([
      getTdeeSheetDashboard(userId),
      getOrCreateAppConfig(userId)
    ]);
    res.json({
      message: 'Excel sheet parameters updated and recalculated.',
      updated: result.updated,
      dashboard: { ...dashboard, config }
    });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

app.post('/api/nutrition/log', async (req, res) => {
  try {
    const userId = req.user.id;
    const { date, weight, calories, waist, neck, hip, clearEmpty, clearMeasurements } = req.body;
    const coords = await logToTdeeSheet(userId, {
      date,
      weight,
      calories,
      waist,
      neck,
      hip,
      clearEmpty: Boolean(clearEmpty),
      clearMeasurements: Boolean(clearMeasurements)
    });

    const updateFields = {
      userId,
      date,
      weekNumber: coords.weekNumber,
      dayName: coords.dayName,
      excelCoordinates: {
        weightRow: coords.weightRow,
        calorieRow: coords.calorieRow,
        dayCol: coords.dayCol
      }
    };
    const setOrClear = (field, val, allowClear) => {
      if (val !== undefined && val !== null && val !== '') {
        updateFields[field] = Number(val);
      } else if (allowClear && (val === '' || val === null)) {
        updateFields[field] = null;
      }
    };
    setOrClear('weight', weight, clearEmpty);
    setOrClear('calories', calories, clearEmpty);
    setOrClear('waist', waist, clearMeasurements);
    setOrClear('neck', neck, clearMeasurements);
    setOrClear('hip', hip, clearMeasurements);

    const mongoDoc = await NutritionLog.findOneAndUpdate(
      { userId, date },
      { $set: updateFields },
      { upsert: true, new: true }
    );

    const updatedDashboard = await getTdeeSheetDashboard(userId);
    res.json({
      message: 'Logged to TDEE Excel sheet and MongoDB.',
      coords,
      mongoDoc,
      dashboard: updatedDashboard
    });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

app.put('/api/nutrition/config', async (req, res) => {
  try {
    const cfg = await getOrCreateAppConfig(req.user.id);
    const { measurementDayOfWeek } = req.body;
    if (measurementDayOfWeek) cfg.measurementDayOfWeek = measurementDayOfWeek;
    await cfg.save();
    res.json(cfg);
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

// --- EXERCISE & PROGRESSIVE OVERLOAD ROUTES ---

app.get('/api/exercises', async (req, res) => {
  try {
    const userId = req.user.id;
    const data = await getAllExerciseData(userId);
    const day = req.query.day || 'Monday';
    const refDate = req.query.date || null;
    const schedule = await getRecommendationsForWeekday(userId, day, refDate);
    res.json({ ...data, schedule });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/exercises/log', async (req, res) => {
  try {
    const entry = await logExerciseSession(req.user.id, req.body);
    res.json(entry);
  } catch (err) {
    res.status(err.status || 400).json({
      error: err.message,
      needsConfirmation: err.needsConfirmation,
      suggestions: err.suggestions
    });
  }
});

app.put('/api/exercises/settings', async (req, res) => {
  try {
    const updated = await updateExerciseTargetSettings(req.user.id, req.body);
    res.json(updated);
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

app.post('/api/exercises/import-txt', upload.single('file'), async (req, res) => {
  try {
    const textContent = req.file ? req.file.buffer.toString('utf8') : req.body.rawText;
    const fallbackDate = req.body.date || new Date().toLocaleDateString('sv-SE');
    const imported = await importExercisesFromText(req.user.id, textContent, fallbackDate);
    res.json({ importedCount: imported.length, imported });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

// --- DELETION ROUTES ---

app.delete('/api/exercises/log/:id', async (req, res) => {
  try {
    await deleteExerciseLog(req.user.id, req.params.id);
    res.json({ message: 'Log entry deleted successfully.' });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

app.delete('/api/exercises/name/:name', async (req, res) => {
  try {
    await clearEntireExercise(req.user.id, req.params.name);
    res.json({ message: `All data for ${req.params.name} cleared.` });
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});

// Unknown API routes return JSON instead of an HTML page
app.use('/api', (req, res) => {
  res.status(404).json({ error: `No such endpoint: ${req.method} ${req.originalUrl}` });
});

if (require.main === module) {
  const PORT = process.env.PORT || 3000;
  const HOST = process.env.HOST || '0.0.0.0';
  connectDatabase()
    .then(syncAllIndexes)
    .then(() => {
      app.listen(PORT, HOST, () => {
        console.log(`MongoDB connected & PWA Server running`);
      });
    })
    .catch((err) => {
      console.error('Failed to connect to MongoDB:', err.message);
      process.exit(1);
    });
}

module.exports = { app, connectDatabase };