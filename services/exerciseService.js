// services/exerciseService.js
const { ExerciseSetting, ExerciseLog } = require('../models');

const DAYS_OF_WEEK = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

const DEFAULT_EXERCISES = [
  { exerciseName: 'Back Squat', targetSets: 4, targetReps: 5, incrementKg: 2.5, currentWeightKg: null },
  { exerciseName: 'Romanian Deadlift', targetSets: 3, targetReps: 8, incrementKg: 2.5, currentWeightKg: null },
  { exerciseName: 'Bulgarian Split Squat', targetSets: 3, targetReps: 8, incrementKg: 2.5, currentWeightKg: null },
  { exerciseName: 'Leg Curl', targetSets: 3, targetReps: 12, incrementKg: 2.5, currentWeightKg: null },
  { exerciseName: 'Standing Calf Raise', targetSets: 3, targetReps: 15, incrementKg: 2.5, currentWeightKg: null },
  { exerciseName: 'Hanging Knee Raise', targetSets: 3, targetReps: 12, incrementKg: 2.5, currentWeightKg: null }
];

function toTitleCase(str) {
  return String(str || '')
    .trim()
    .toLowerCase()
    .split(/\s+/)
    .filter(Boolean)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(' ');
}

function editDistance(a, b) {
  const dp = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array(b.length).fill(0)]);
  for (let j = 1; j <= b.length; j++) dp[0][j] = j;
  for (let i = 1; i <= a.length; i++)
    for (let j = 1; j <= b.length; j++)
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
  return dp[a.length][b.length];
}

async function suggestSimilarExercises(userId, name) {
  const lower = name.toLowerCase();
  const all = await ExerciseSetting.find({ userId }).lean();
  return all
    .map((s) => ({ n: s.exerciseName, d: editDistance(lower, s.exerciseName.toLowerCase()) }))
    .filter((x) => x.d <= 3 || x.n.toLowerCase().includes(lower) || lower.includes(x.n.toLowerCase()))
    .sort((a, b) => a.d - b.d)
    .slice(0, 3)
    .map((x) => x.n);
}

async function seedDefaultExercises(userId) {
  const count = await ExerciseSetting.countDocuments({ userId });
  if (count === 0) {
    await ExerciseSetting.insertMany(
      DEFAULT_EXERCISES.map((ex) => ({ ...ex, userId }))
    );
  }
}

function todayLocal() {
  return new Date().toISOString().split('T')[0];
}

function isValidYMD(y, m, d) {
  if (!Number.isInteger(y) || !Number.isInteger(m) || !Number.isInteger(d)) return false;
  if (m < 1 || m > 12 || d < 1 || d > 31) return false;
  const dt = new Date(Date.UTC(y, m - 1, d));
  return dt.getUTCFullYear() === y && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d;
}

function parseFlexibleDate(str) {
  const s = String(str || '').trim().replace(/^\[|\]$/g, '');
  let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (m) {
    const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
    if (!isValidYMD(y, mo, d)) return null;
    return `${String(y).padStart(4, '0')}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  }
  m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/);
  if (m) {
    const yearStr = m[3].length === 2 ? `20${m[3]}` : m[3];
    const y = Number(yearStr), mo = Number(m[2]), d = Number(m[1]);
    if (!isValidYMD(y, mo, d)) return null;
    return `${String(y).padStart(4, '0')}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  }
  return null;
}

function parseRepsInput(repsInput) {
  if (Array.isArray(repsInput)) {
    return repsInput.map((n) => parseInt(n, 10)).filter((n) => !Number.isNaN(n));
  }
  const tokens = String(repsInput || '')
    .trim()
    .split(/[\s,/|-]+/)
    .filter(Boolean);
  if (tokens.some((t) => !/^\d+$/.test(t))) {
    throw new Error(`Invalid reps: "${repsInput}". Use numbers like "12 12 12".`);
  }
  return tokens.map(Number);
}

function getDayOfWeekName(dateStr) {
  const [y, m, d] = dateStr.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  return DAYS_OF_WEEK[dt.getUTCDay()];
}

function addDays(dateStr, days) {
  const [y, m, d] = dateStr.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d + days));
  return dt.toISOString().split('T')[0];
}

function evaluateExerciseProgression({
  weightKg,
  actualRepsArray,
  targetSets,
  targetReps,
  incrementKg = 2.5
}) {
  let numericWeight = Number(weightKg);
  if (Number.isNaN(numericWeight)) {
    numericWeight = 0;
  }

  const reqSets = Number(targetSets);
  const reqReps = Number(targetReps);
  const stepKg = Number(incrementKg);

  const setEvaluations = [];
  const failedSets = [];

  for (let i = 0; i < Math.max(reqSets, actualRepsArray.length); i++) {
    const achieved = actualRepsArray[i] !== undefined ? actualRepsArray[i] : 0;
    const passed = achieved >= reqReps;
    setEvaluations.push({
      setNumber: i + 1,
      achievedReps: achieved,
      targetReps: reqReps,
      passed
    });
    if (i < reqSets && !passed) {
      failedSets.push(`Set ${i + 1} (${achieved}/${reqReps})`);
    }
  }

  const completedAllRequiredSets = actualRepsArray.length >= reqSets;
  const allRequiredSetsMetTarget = setEvaluations.slice(0, reqSets).every((s) => s.passed);
  const isSuccess = completedAllRequiredSets && allRequiredSetsMetTarget;

  const nextWeekWeightKg = isSuccess
    ? Number((numericWeight + stepKg).toFixed(2))
    : numericWeight;

  return {
    isSuccess,
    statusLabel: isSuccess ? `SUCCESS (+${stepKg} kg Next Session)` : 'FAILED (Keep Weight Same)',
    currentWeightKg: numericWeight,
    nextWeekWeightKg,
    targetSets: reqSets,
    targetReps: reqReps,
    actualRepsArray,
    setEvaluations,
    failedSets,
    reason: isSuccess
      ? `All ${reqSets} sets reached target of ${reqReps} reps (${actualRepsArray.join(' ')}). Saved weight increases by ${stepKg} kg.`
      : `Failed to reach ${reqSets}×${reqReps} on: ${failedSets.join(', ')}. Saved weight remains at ${numericWeight} kg.`
  };
}

// Syncs an exercise's saved currentWeightKg from its most recent log across any day of the week
async function syncExerciseCurrentWeight(userId, exerciseName) {
  const cleanName = toTitleCase(String(exerciseName || '').trim());
  if (!cleanName) return null;

  const [setting, latestLog] = await Promise.all([
    ExerciseSetting.findOne({ userId, exerciseName: cleanName }),
    ExerciseLog.findOne({ userId, exerciseName: cleanName }).sort({ date: -1, createdAt: -1 }).lean()
  ]);

  if (!setting) return null;

  if (latestLog) {
    const stepKg = Number(setting.incrementKg ?? 2.5);
    const nextWeight = latestLog.isSuccess
      ? Number((latestLog.weightKg + stepKg).toFixed(2))
      : latestLog.weightKg;
    setting.currentWeightKg = nextWeight;
  } else {
    setting.currentWeightKg = null;
  }

  await setting.save();
  return setting.currentWeightKg;
}

async function logExerciseSession(userId, { date, exerciseName, weightKg, actualRepsInput, notes = '', confirmNew = false }) {
  const cleanName = toTitleCase(String(exerciseName || '').trim());
  const actualRepsArray = parseRepsInput(actualRepsInput);

  if (!cleanName || actualRepsArray.length === 0) {
    throw new Error(`Invalid input for "${cleanName}". Provide space-separated reps like "12 12 12".`);
  }

  if (weightKg === undefined || weightKg === null || String(weightKg).trim() === '' || Number.isNaN(Number(weightKg))) {
    throw new Error('Enter the weight used (0 for bodyweight exercises).');
  }

  let setting = await ExerciseSetting.findOne({ userId, exerciseName: cleanName });
  if (!setting) {
    if (!confirmNew) {
      const err = new Error(`"${cleanName}" is not an existing exercise.`);
      err.status = 409;
      err.needsConfirmation = true;
      err.suggestions = await suggestSimilarExercises(userId, cleanName);
      throw err;
    }
    setting = await ExerciseSetting.create({
      userId,
      exerciseName: cleanName,
      targetSets: Math.max(1, actualRepsArray.length),
      targetReps: Math.max(1, ...actualRepsArray),
      incrementKg: 2.5,
      currentWeightKg: Number(weightKg)
    });
  }

  const evaluation = evaluateExerciseProgression({
    weightKg,
    actualRepsArray,
    targetSets: setting.targetSets,
    targetReps: setting.targetReps,
    incrementKg: setting.incrementKg
  });

  const rawResult = await ExerciseLog.findOneAndUpdate(
    { userId, date, exerciseName: cleanName },
    {
      $set: {
        userId,
        date,
        dayOfWeek: getDayOfWeekName(date),
        nextWeekDate: addDays(date, 7),
        exerciseName: cleanName,
        weightKg: evaluation.currentWeightKg,
        targetSets: evaluation.targetSets,
        targetReps: evaluation.targetReps,
        actualReps: evaluation.actualRepsArray,
        isSuccess: evaluation.isSuccess,
        nextWeekWeightKg: evaluation.nextWeekWeightKg,
        evaluationReason: evaluation.reason,
        failedSets: evaluation.failedSets,
        notes
      }
    },
    { upsert: true, new: true, rawResult: true }
  );

  // Persist the updated working weight on the exercise itself based on its latest session
  const savedWeightKg = await syncExerciseCurrentWeight(userId, cleanName);

  const entry = rawResult.value || rawResult;
  const replaced = Boolean(rawResult.lastErrorObject?.updatedExisting);

  return {
    ...(entry.toObject ? entry.toObject() : entry),
    savedWeightKg,
    replaced: replaced ? 1 : 0
  };
}

async function updateExerciseTargetSettings(userId, { exerciseName, targetSets, targetReps, incrementKg, currentWeightKg }) {
  const cleanName = toTitleCase(String(exerciseName || '').trim());
  const setsNum = Number(targetSets);
  const repsNum = Number(targetReps);
  if (!cleanName) throw new Error('Exercise name is required.');
  if (!Number.isFinite(setsNum) || setsNum < 1) throw new Error('Target sets must be at least 1.');
  if (!Number.isFinite(repsNum) || repsNum < 1) throw new Error('Target reps must be at least 1.');

  const $set = {
    userId,
    exerciseName: cleanName,
    targetSets: Math.round(setsNum),
    targetReps: Math.round(repsNum)
  };
  const $setOnInsert = {};

  if (incrementKg !== undefined && incrementKg !== null && String(incrementKg).trim() !== '') {
    const incNum = Number(incrementKg);
    if (!Number.isFinite(incNum) || incNum < 0) throw new Error('Increment must be 0 or greater.');
    $set.incrementKg = incNum;
  } else {
    $setOnInsert.incrementKg = 2.5;
  }

  const hasExplicitWeight = currentWeightKg !== undefined && currentWeightKg !== null && String(currentWeightKg).trim() !== '';
  if (hasExplicitWeight) {
    const weightNum = Number(currentWeightKg);
    if (!Number.isFinite(weightNum) || weightNum < 0) throw new Error('Working weight must be 0 or greater.');
    $set.currentWeightKg = Number(weightNum.toFixed(2));
  } else if (currentWeightKg === null || currentWeightKg === '') {
    // Allow clearing or recalculating from latest log if left blank
    const latestLog = await ExerciseLog.findOne({ userId, exerciseName: cleanName }).sort({ date: -1, createdAt: -1 }).lean();
    if (latestLog) {
      const step = $set.incrementKg ?? 2.5;
      $set.currentWeightKg = latestLog.isSuccess
        ? Number((latestLog.weightKg + step).toFixed(2))
        : latestLog.weightKg;
    }
  }

  const updateOp = Object.keys($setOnInsert).length ? { $set, $setOnInsert } : { $set };
  const updated = await ExerciseSetting.findOneAndUpdate(
    { userId, exerciseName: cleanName },
    updateOp,
    { upsert: true, new: true, runValidators: true }
  );
  return updated.toObject();
}

async function getRecommendationsForWeekday(userId, dayOfWeekOrDate, referenceDate = null) {
  const rawDay = String(dayOfWeekOrDate || '');
  const candidate = rawDay.includes('-') ? getDayOfWeekName(rawDay) : rawDay;
  const targetDayName = DAYS_OF_WEEK.find((d) => d.toLowerCase() === candidate.toLowerCase());
  if (!targetDayName) throw new Error(`Invalid day: "${rawDay}"`);

  // 1. Identify which exercises are performed on this weekday (keeps the weekly routine schedule feature as-is)
  const dayLogs = await ExerciseLog.find({ userId, dayOfWeek: targetDayName }).sort({ date: 1, createdAt: 1 }).lean();
  const settingsMap = new Map((await ExerciseSetting.find({ userId }).lean()).map((s) => [s.exerciseName, s]));

  const dayLogsByExercise = new Map();
  for (const log of dayLogs) {
    if (!dayLogsByExercise.has(log.exerciseName)) dayLogsByExercise.set(log.exerciseName, []);
    dayLogsByExercise.get(log.exerciseName).push(log);
  }

  const exerciseNames = [...dayLogsByExercise.keys()];
  // 2. Fetch all logs for those exercises across ANY day of the week so repeating an exercise updates its weight everywhere
  const allExerciseLogs = exerciseNames.length
    ? await ExerciseLog.find({ userId, exerciseName: { $in: exerciseNames } }).sort({ date: 1, createdAt: 1 }).lean()
    : [];

  const globalLogsByExercise = new Map();
  for (const log of allExerciseLogs) {
    if (!globalLogsByExercise.has(log.exerciseName)) globalLogsByExercise.set(log.exerciseName, []);
    globalLogsByExercise.get(log.exerciseName).push(log);
  }

  const recommendations = [];
  for (const [exerciseName, weekdayLogs] of dayLogsByExercise.entries()) {
    const priorWeekdayLogs = referenceDate ? weekdayLogs.filter((l) => l.date < referenceDate) : weekdayLogs;
    const lastWeekdayLog = priorWeekdayLogs.length > 0 ? priorWeekdayLogs[priorWeekdayLogs.length - 1] : weekdayLogs[weekdayLogs.length - 1];

    const globalLogs = globalLogsByExercise.get(exerciseName) || weekdayLogs;
    const priorGlobalLogs = referenceDate ? globalLogs.filter((l) => l.date < referenceDate) : globalLogs;
    const lastGlobalLog = priorGlobalLogs.length > 0 ? priorGlobalLogs[priorGlobalLogs.length - 1] : globalLogs[globalLogs.length - 1];
    const isLatestOverall = lastGlobalLog && globalLogs[globalLogs.length - 1]?._id.toString() === lastGlobalLog._id.toString();

    const todayLog = referenceDate ? globalLogs.find((l) => l.date === referenceDate) : null;

    const currentSettings = settingsMap.get(exerciseName) || {
      targetSets: lastGlobalLog.targetSets,
      targetReps: lastGlobalLog.targetReps,
      incrementKg: 2.5,
      currentWeightKg: null
    };

    const stepKg = Number(currentSettings.incrementKg ?? 2.5);
    const lastSessionSuccess = Boolean(lastGlobalLog.isSuccess);
    const computedNextWeightKg = lastSessionSuccess
      ? Number((lastGlobalLog.weightKg + stepKg).toFixed(2))
      : lastGlobalLog.weightKg;

    // Use the exercise's saved weight if we're viewing the current state, or the historical progression for past dates
    const recommendedWeightKg = (isLatestOverall && currentSettings.currentWeightKg !== null && currentSettings.currentWeightKg !== undefined)
      ? Number(currentSettings.currentWeightKg)
      : computedNextWeightKg;

    recommendations.push({
      exerciseName,
      dayOfWeek: targetDayName,
      lastPerformedDate: lastGlobalLog.date,
      recommendedForDate: lastWeekdayLog.nextWeekDate,
      lastWeightKg: lastGlobalLog.weightKg,
      lastActualReps: lastGlobalLog.actualReps.join(' '),
      lastSessionSuccess,
      recommendedWeightKg,
      targetSets: currentSettings.targetSets,
      targetReps: currentSettings.targetReps,
      incrementKg: stepKg,
      loggedToday: Boolean(todayLog),
      rationale: lastGlobalLog.evaluationReason
    });
  }

  return {
    dayOfWeek: targetDayName,
    recommendations
  };
}

async function importExercisesFromText(userId, rawText, fallbackDate) {
  const lines = String(rawText || '')
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l && !l.startsWith('#') && !l.startsWith('//'));

  const importedEntries = [];
  let currentDate = parseFlexibleDate(fallbackDate) || fallbackDate || todayLocal();

  for (const line of lines) {
    const cleanedHeader = line.replace(/^date\s*:\s*/i, '').trim();
    const parsedHeaderDate = parseFlexibleDate(cleanedHeader);
    if (parsedHeaderDate && !line.includes('|')) {
      currentDate = parsedHeaderDate;
      continue;
    }

    const parts = line.split('|').map((p) => p.trim());
    if (parts.length < 3) continue;

    let entryDate = currentDate;
    let idx = 0;

    const firstPartDate = parseFlexibleDate(parts[0]);
    if (firstPartDate) {
      entryDate = firstPartDate;
      idx = 1;
    }

    let rawExercisePart = parts[idx] || '';
    const leadingDateMatch = rawExercisePart.match(/^(\d{1,4}[-/]\d{1,2}[-/]\d{1,4})\s+(.+)$/);
    if (leadingDateMatch) {
      const inlineDate = parseFlexibleDate(leadingDateMatch[1]);
      if (inlineDate) {
        entryDate = inlineDate;
        rawExercisePart = leadingDateMatch[2];
      }
    }

    const exerciseName = toTitleCase(rawExercisePart.trim());
    const weightStr = (parts[idx + 1] || '').replace(/[^\d.]/g, '');
    const weightKg = weightStr === '' ? 0 : parseFloat(weightStr);
    let repsPart = parts[idx + 2];

    if (parts[idx + 3]) {
      const targetPart = parts[idx + 2].replace(/[×✕]/g, 'x');
      const sxr = targetPart.match(/^(\d+)\s*x\s*(\d+)(?:\s*[-–]\s*(\d+))?$/i);
      const setsOnly = targetPart.match(/^(\d+)\s*sets?$/i);

      if (sxr) {
        await updateExerciseTargetSettings(userId, {
          exerciseName,
          targetSets: parseInt(sxr[1], 10),
          targetReps: parseInt(sxr[3] || sxr[2], 10)
        });
        repsPart = parts[idx + 3];
      } else if (setsOnly) {
        const reps = parseRepsInput(parts[idx + 3]);
        if (reps.length === 0) {
          throw new Error(`No reps found for "${exerciseName}" in line: ${line}`);
        }
        await updateExerciseTargetSettings(userId, {
          exerciseName,
          targetSets: parseInt(setsOnly[1], 10),
          targetReps: Math.max(...reps)
        });
        repsPart = parts[idx + 3];
      }
    }

    const logged = await logExerciseSession(userId, {
      date: entryDate,
      exerciseName,
      weightKg,
      actualRepsInput: repsPart,
      confirmNew: true
    });
    importedEntries.push(logged);
  }

  return importedEntries;
}

async function getAllExerciseData(userId) {
  await seedDefaultExercises(userId);
  const settingsDocs = await ExerciseSetting.find({ userId }).lean();
  const logs = await ExerciseLog.find({ userId }).sort({ date: 1, createdAt: 1 }).lean();

  const latestLogByExercise = new Map();
  for (const l of logs) {
    latestLogByExercise.set(l.exerciseName, l);
  }

  const exerciseSettings = {};
  const backfillPromises = [];

  for (const s of settingsDocs) {
    let savedWeight = s.currentWeightKg ?? null;
    const latest = latestLogByExercise.get(s.exerciseName);

    // Automatically backfill saved weight for existing logs if not yet stored
    if (savedWeight === null && latest) {
      const step = Number(s.incrementKg ?? 2.5);
      savedWeight = latest.isSuccess
        ? Number((latest.weightKg + step).toFixed(2))
        : latest.weightKg;
      backfillPromises.push(
        ExerciseSetting.updateOne({ _id: s._id }, { $set: { currentWeightKg: savedWeight } })
      );
    }

    exerciseSettings[s.exerciseName] = {
      targetSets: s.targetSets,
      targetReps: s.targetReps,
      incrementKg: s.incrementKg,
      currentWeightKg: savedWeight
    };
  }

  if (backfillPromises.length > 0) {
    await Promise.all(backfillPromises);
  }

  return { exerciseSettings, logs };
}

async function deleteExerciseLog(userId, logId) {
  const deleted = await ExerciseLog.findOneAndDelete({ _id: logId, userId });
  if (deleted) {
    await syncExerciseCurrentWeight(userId, deleted.exerciseName);
  }
}

async function clearEntireExercise(userId, exerciseName) {
  const cleanName = toTitleCase(String(exerciseName || '').trim());
  await ExerciseLog.deleteMany({ userId, exerciseName: cleanName });
  await ExerciseSetting.findOneAndDelete({ userId, exerciseName: cleanName });
}

async function clearAllExerciseData(userId, { resetTargets = false } = {}) {
  const logs = await ExerciseLog.deleteMany({ userId });
  let targetsDeleted = 0;
  if (resetTargets) {
    targetsDeleted = (await ExerciseSetting.deleteMany({ userId })).deletedCount;
    await seedDefaultExercises(userId);
  } else {
    await ExerciseSetting.updateMany({ userId }, { $set: { currentWeightKg: null } });
  }
  return { logsDeleted: logs.deletedCount, targetsDeleted };
}

module.exports = {
  seedDefaultExercises,
  syncExerciseCurrentWeight,
  clearAllExerciseData,
  parseRepsInput,
  evaluateExerciseProgression,
  logExerciseSession,
  updateExerciseTargetSettings,
  getRecommendationsForWeekday,
  importExercisesFromText,
  getAllExerciseData,
  deleteExerciseLog,
  clearEntireExercise
};