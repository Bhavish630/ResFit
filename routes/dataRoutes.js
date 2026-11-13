// routes/dataRoutes.js
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const { User, AppConfig, ExerciseLog, ExerciseSetting, NutritionLog } = require('../models');
const { clearAllExerciseData, syncExerciseCurrentWeight } = require('../services/exerciseService');
const {
  clearTdeeSheetEntries,
  bulkLogToTdeeSheet,
  ensureUserSheet,
  updateSheetConfig,
  getTdeeSheetDashboard,
  getCoordinatesForDate,
  normalizeWeightUnit
} = require('../services/tdeeSheetService');
const { readTdeeWorksheet, findTdeeWorksheet } = require('../services/excelImportService');

const router = express.Router();
const xlsxUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

const WEEKDAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
const SHEET_CONFIG_KEYS = [
  'startDate', 'weightUnit', 'energyUnit', 'startingWeight', 'goalWeight',
  'weeklyGoalRate', 'gender', 'measureUnit', 'height', 'tdeeAvgWeeks'
];

function isYMD(v) {
  return typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v);
}

function shiftDate(ymd, days) {
  const [y, m, d] = ymd.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d + days)).toISOString().split('T')[0];
}

function cleanSheetConfig(raw) {
  const out = {};
  if (!raw || typeof raw !== 'object') return out;
  const mustBePositive = ['startingWeight', 'goalWeight', 'height', 'tdeeAvgWeeks'];
  for (const key of SHEET_CONFIG_KEYS) {
    const v = raw[key];
    if (v === undefined || v === null || v === '') continue;
    if (mustBePositive.includes(key) && !(Number(v) > 0)) continue;
    out[key] = key === 'weightUnit' ? normalizeWeightUnit(v) : v;
  }
  return out;
}

function inferStartDate(nutritionLogs) {
  const valid = (nutritionLogs || []).filter((n) => n && isYMD(n.date));
  if (!valid.length) return null;

  for (const n of valid) {
    const week = Number(n.weekNumber);
    const dayCol = Number(n.excelCoordinates?.dayCol);
    if (Number.isInteger(week) && week >= 1 && Number.isInteger(dayCol) && dayCol >= 4 && dayCol <= 10) {
      return shiftDate(n.date, -((week - 1) * 7 + (dayCol - 4)));
    }
  }
  return valid.map((n) => n.date).sort()[0];
}

router.get('/api/data/summary', async (req, res) => {
  try {
    const userId = req.user.id;
    const [exerciseLogs, exerciseTargets, nutritionLogs] = await Promise.all([
      ExerciseLog.countDocuments({ userId }),
      ExerciseSetting.countDocuments({ userId }),
      NutritionLog.countDocuments({ userId })
    ]);
    res.json({ exerciseLogs, exerciseTargets, nutritionLogs });
  } catch (e) {
    console.error('Error in /api/data/summary:', e.message);
    res.status(500).json({ error: 'An internal server error occurred' });
  }
});

router.get('/api/data/export.json', async (req, res) => {
  try {
    const userId = req.user.id;
    const [userDoc, appConfigDoc, exerciseLogs, exerciseSettings, nutritionLogs] = await Promise.all([
      User.findById(userId).select('username sheetConfig').lean(),
      AppConfig.findOne({ userId }).lean(),
      ExerciseLog.find({ userId }).lean(),
      ExerciseSetting.find({ userId }).lean(),
      NutritionLog.find({ userId }).lean()
    ]);

    let sheetConfig = userDoc?.sheetConfig || null;
    try {
      const dash = await getTdeeSheetDashboard(userId);
      if (dash?.sheetConfig) sheetConfig = { ...(sheetConfig || {}), ...dash.sheetConfig };
    } catch (_) {}

    const safeUser = userDoc?.username || 'user';
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="resfit-${safeUser}-backup-${new Date().toISOString().split('T')[0]}.json"`
    );
    res.json({
      backupVersion: 2,
      exportedAt: new Date().toISOString(),
      username: safeUser,
      sheetConfig,
      appConfig: { measurementDayOfWeek: appConfigDoc?.measurementDayOfWeek || 'Monday' },
      exerciseSettings,
      exerciseLogs,
      nutritionLogs
    });
  } catch (e) {
    console.error('Error in /api/data/export.json:', e.message);
    res.status(500).json({ error: 'An internal server error occurred' });
  }
});

router.post('/api/data/import.json', async (req, res) => {
  try {
    const userId = req.user.id;
    const { sheetConfig, appConfig, exerciseSettings, exerciseLogs, nutritionLogs, mode = 'merge' } = req.body || {};
    if (!Array.isArray(exerciseSettings) && !Array.isArray(exerciseLogs) && !Array.isArray(nutritionLogs)) {
      return res.status(400).json({ error: 'Invalid backup JSON: expected exerciseSettings, exerciseLogs, or nutritionLogs arrays.' });
    }

    const warnings = [];
    await ensureUserSheet(userId);

    if (mode === 'replace') {
      await clearTdeeSheetEntries(userId);
      await Promise.all([
        ExerciseLog.deleteMany({ userId }),
        ExerciseSetting.deleteMany({ userId }),
        NutritionLog.deleteMany({ userId })
      ]);
    }

    const configToApply = cleanSheetConfig(sheetConfig);
    let startDateInferred = false;
    if (!configToApply.startDate && Array.isArray(nutritionLogs) && nutritionLogs.length > 0) {
      const inferred = inferStartDate(nutritionLogs);
      if (inferred) {
        configToApply.startDate = inferred;
        startDateInferred = true;
      }
    }
    if (startDateInferred && configToApply.startingWeight === undefined) {
      const firstWeighIn = nutritionLogs
        .filter((n) => n && isYMD(n.date) && n.weight !== null && n.weight !== undefined && Number(n.weight) > 0)
        .sort((a, b) => a.date.localeCompare(b.date))[0];
      if (firstWeighIn) configToApply.startingWeight = Number(firstWeighIn.weight);
    }
    if (Object.keys(configToApply).length > 0) {
      await updateSheetConfig(userId, configToApply);
    }
    if (startDateInferred) {
      warnings.push(
        `Your backup had no tracker settings, so the Start Date (${configToApply.startDate}) was worked out from your logs. ` +
        'Please check Start Date, Goal Weight, Weekly Rate and Height under Tracker Preferences.'
      );
    }

    const mDay = appConfig?.measurementDayOfWeek;
    if (WEEKDAYS.includes(mDay)) {
      await AppConfig.findOneAndUpdate(
        { userId },
        { $set: { userId, measurementDayOfWeek: mDay } },
        { upsert: true, new: true }
      );
    }

    let settingsCount = 0;
    if (Array.isArray(exerciseSettings)) {
      for (const s of exerciseSettings) {
        if (!s || !s.exerciseName) continue;
        const setFields = {
          userId,
          exerciseName: s.exerciseName,
          targetSets: Math.max(1, Number(s.targetSets) || 3),
          targetReps: Math.max(1, Number(s.targetReps) || 10),
          incrementKg: Math.max(0, Number(s.incrementKg ?? 2.5))
        };
        if (s.currentWeightKg !== undefined && s.currentWeightKg !== null && s.currentWeightKg !== '') {
          setFields.currentWeightKg = Math.max(0, Number(s.currentWeightKg) || 0);
        }
        await ExerciseSetting.findOneAndUpdate(
          { userId, exerciseName: s.exerciseName },
          { $set: setFields },
          { upsert: true, new: true }
        );
        settingsCount++;
      }
    }

    let logsCount = 0;
    const touchedExercises = new Set();
    if (Array.isArray(exerciseLogs)) {
      for (const l of exerciseLogs) {
        if (!l || !l.date || !l.exerciseName) continue;
        const cleanExerciseLog = {
          userId,
          date: String(l.date),
          dayOfWeek: String(l.dayOfWeek || ''),
          nextWeekDate: String(l.nextWeekDate || ''),
          exerciseName: String(l.exerciseName),
          weightKg: Number(l.weightKg) || 0,
          targetSets: Number(l.targetSets) || 1,
          targetReps: Number(l.targetReps) || 1,
          actualReps: Array.isArray(l.actualReps) ? l.actualReps.map(Number) : [],
          isSuccess: Boolean(l.isSuccess),
          nextWeekWeightKg: Number(l.nextWeekWeightKg) || 0,
          evaluationReason: String(l.evaluationReason || ''),
          failedSets: Array.isArray(l.failedSets) ? l.failedSets.map(String) : [],
          notes: String(l.notes || '')
        };

        await ExerciseLog.findOneAndUpdate(
          { userId, date: cleanExerciseLog.date, exerciseName: cleanExerciseLog.exerciseName },
          { $set: cleanExerciseLog },
          { upsert: true, new: true }
        );
        touchedExercises.add(cleanExerciseLog.exerciseName);
        logsCount++;
      }
    }

    for (const exName of touchedExercises) {
      await syncExerciseCurrentWeight(userId, exName);
    }

    let nutritionCount = 0;
    let sheetApplied = 0;
    let sheetSkipped = 0;
    if (Array.isArray(nutritionLogs)) {
      const userDoc = await User.findById(userId).select('sheetConfig').lean();
      const startDate = userDoc?.sheetConfig?.startDate || null;

      for (const n of nutritionLogs) {
        if (!n || !n.date) continue;
        const cleanNut = {
          userId,
          date: String(n.date),
          weekNumber: Number(n.weekNumber) || 1,
          dayName: String(n.dayName || ''),
          weight: n.weight !== undefined && n.weight !== null ? Number(n.weight) : null,
          calories: n.calories !== undefined && n.calories !== null ? Number(n.calories) : null,
          waist: n.waist !== undefined && n.waist !== null ? Number(n.waist) : null,
          neck: n.neck !== undefined && n.neck !== null ? Number(n.neck) : null,
          hip: n.hip !== undefined && n.hip !== null ? Number(n.hip) : null,
          sourceApp: String(n.sourceApp || 'Manual')
        };

        if (startDate) {
          try {
            const c = getCoordinatesForDate(startDate, cleanNut.date);
            cleanNut.weekNumber = c.weekNumber;
            cleanNut.dayName = c.dayName;
            cleanNut.excelCoordinates = { weightRow: c.weightRow, calorieRow: c.calorieRow, dayCol: c.dayCol };
          } catch (_) {}
        }

        await NutritionLog.findOneAndUpdate(
          { userId, date: cleanNut.date },
          { $set: cleanNut },
          { upsert: true, new: true }
        );
        nutritionCount++;
      }

      if (nutritionLogs.length > 0) {
        sheetApplied = await bulkLogToTdeeSheet(userId, nutritionLogs);
        sheetSkipped = Math.max(0, nutritionLogs.length - sheetApplied);
        if (sheetSkipped > 0) {
          warnings.push(
            `${sheetSkipped} nutrition entr${sheetSkipped === 1 ? 'y was' : 'ies were'} outside the 56-week sheet window ` +
            `starting ${startDate || 'your start date'} and could not be written to the spreadsheet. ` +
            'Adjust Start Date in Tracker Preferences and import again (Merge mode).'
          );
        }
      }
    }

    res.json({
      ok: true,
      mode,
      imported: {
        exerciseSettings: settingsCount,
        exerciseLogs: logsCount,
        nutritionLogs: nutritionCount,
        sheetEntriesApplied: sheetApplied,
        sheetEntriesSkipped: sheetSkipped
      },
      warnings
    });
  } catch (e) {
    res.status(400).json({ error: e.message });
  }
});

router.post('/api/data/import/sheet', xlsxUpload.single('file'), async (req, res) => {
  try {
    const userId = req.user.id;
    if (!req.file) return res.status(400).json({ error: 'Choose an .xlsx file first.' });
    const mode = req.body?.mode === 'replace' ? 'replace' : 'merge';

    const looksZip = req.file.buffer.length > 4 && req.file.buffer[0] === 0x50 && req.file.buffer[1] === 0x4b;
    if (!/\.xlsx$/i.test(req.file.originalname || '') || !looksZip) {
      return res.status(415).json({ error: 'Only .xlsx Excel workbooks are supported (not .xls, .csv or renamed files).' });
    }

    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.load(req.file.buffer);
    } catch (_) {
      return res.status(400).json({ error: 'That file could not be opened as an .xlsx workbook.' });
    }
    const sheet = findTdeeWorksheet(workbook);

    await ensureUserSheet(userId);
    const cfgDoc = await AppConfig.findOne({ userId }).lean();
    const measurementDay = cfgDoc?.measurementDayOfWeek || 'Monday';

    const { config, entries, stats, warnings: parseWarnings } = readTdeeWorksheet(sheet, { measurementDay });
    if (entries.length === 0 && Object.keys(config).length <= 1) {
      return res.status(400).json({ error: 'No settings or logged entries were found in that workbook.' });
    }

    const warnings = [...parseWarnings];
    const userBefore = await User.findById(userId).select('sheetConfig').lean();
    const previousStart = userBefore?.sheetConfig?.startDate || null;
    const startChanged = Boolean(previousStart && previousStart !== config.startDate);

    if (mode === 'replace') {
      await Promise.all([clearTdeeSheetEntries(userId), NutritionLog.deleteMany({ userId })]);
    } else if (startChanged) {
      await clearTdeeSheetEntries(userId);
      warnings.push(`Start Date changed from ${previousStart} to ${config.startDate}; existing entries were re-placed on the new calendar.`);
    }

    await updateSheetConfig(userId, config);

    for (const e of entries) {
      const set = { userId, date: e.date };
      for (const f of ['weight', 'calories', 'waist', 'neck', 'hip']) {
        if (e[f] !== undefined) set[f] = e[f];
      }
      try {
        const c = getCoordinatesForDate(config.startDate, e.date);
        set.weekNumber = c.weekNumber;
        set.dayName = c.dayName;
        set.excelCoordinates = { weightRow: c.weightRow, calorieRow: c.calorieRow, dayCol: c.dayCol };
      } catch (_) { continue; }
      set.sourceApp = 'Excel Import';
      await NutritionLog.findOneAndUpdate({ userId, date: e.date }, { $set: set }, { upsert: true, new: true });
    }

    const allLogs = await NutritionLog.find({ userId }).lean();
    for (const n of allLogs) {
      try {
        const c = getCoordinatesForDate(config.startDate, n.date);
        await NutritionLog.updateOne({ _id: n._id }, {
          $set: {
            weekNumber: c.weekNumber,
            dayName: c.dayName,
            excelCoordinates: { weightRow: c.weightRow, calorieRow: c.calorieRow, dayCol: c.dayCol }
          }
        });
      } catch (_) {}
    }
    const applied = await bulkLogToTdeeSheet(userId, allLogs);
    if (applied < allLogs.length) {
      warnings.push(`${allLogs.length - applied} stored entr${allLogs.length - applied === 1 ? 'y falls' : 'ies fall'} outside the 56-week sheet window and ${allLogs.length - applied === 1 ? 'was' : 'were'} not written to the spreadsheet.`);
    }

    res.json({
      ok: true,
      mode,
      sheetName: sheet.name,
      settings: config,
      imported: { days: stats.days, measurementWeeks: stats.measurementWeeks, sheetEntriesApplied: applied },
      warnings
    });
  } catch (e) {
    res.status(e.status || 400).json({ error: e.message });
  }
});

router.get('/api/data/export/sheet', async (req, res) => {
  try {
    const userSheetPath = await ensureUserSheet(req.user.id);
    res.download(userSheetPath, `TDEE-${req.user.username}.xlsx`);
  } catch (e) {
    res.status(404).json({ error: e.message });
  }
});

router.post('/api/data/clear', async (req, res) => {
  try {
    const userId = req.user.id;
    const { scope, confirm, resetTargets } = req.body || {};
    if (confirm !== 'DELETE') return res.status(400).json({ error: 'Type DELETE to confirm.' });
    if (!['exercises', 'nutrition', 'all'].includes(scope)) return res.status(400).json({ error: 'Invalid scope.' });

    const result = { ok: true };
    if (scope === 'nutrition' || scope === 'all') {
      result.nutrition = await clearTdeeSheetEntries(userId);
      await NutritionLog.deleteMany({ userId });
    }
    if (scope === 'exercises' || scope === 'all') {
      result.exercises = await clearAllExerciseData(userId, { resetTargets: !!resetTargets });
    }
    res.json(result);
  } catch (e) {
    console.error('Error in /api/data/clear:', e.message);
    res.status(500).json({ error: 'An internal server error occurred' });
  }
});

module.exports = router;