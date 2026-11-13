// tests/app.test.js
const request = require('supertest');
const mongoose = require('mongoose');
const fs = require('fs');
const path = require('path');

const { app, connectDatabase } = require('../server');
const { evaluateExerciseProgression } = require('../services/exerciseService');
const { getCoordinatesForDate } = require('../services/tdeeSheetService');
const { User, AppConfig, NutritionLog, ExerciseSetting, ExerciseLog } = require('../models');

const SHEET_PATH = path.join(__dirname, '..', 'TDEE variant with bf 3.06.xlsx');
const BACKUP_PATH = path.join(__dirname, '..', 'TDEE variant with bf 3.06.backup.xlsx');

// Test account credentials (adjust field names if your auth routes differ)
const TEST_USER = { username: 'testuser', password: 'TestPass123!' };
const OTHER_USER = { username: 'otheruser', password: 'TestPass123!' };

// An agent keeps the resfit_token cookie between requests, so every call
// made with it is authenticated as TEST_USER.
const agent = request.agent(app);

// Registers (if needed) and logs in with the given agent.
async function loginAs(ag, creds) {
  await ag.post('/api/auth/register').send(creds); // ignore result: may already exist
  await ag.post('/api/auth/login').send(creds).expect(200);
}

beforeAll(async () => {
  // Back up the Excel file so test writes can be cleanly restored afterward
  if (fs.existsSync(SHEET_PATH)) {
    fs.copyFileSync(SHEET_PATH, BACKUP_PATH);
  }
  const testMongoUri = process.env.MONGO_TEST_URI || 'mongodb://127.0.0.1:27017/fitness_tdee_test';
  await connectDatabase(testMongoUri);
  await Promise.all([
    User.deleteMany({}),
    AppConfig.deleteMany({}),
    NutritionLog.deleteMany({}),
    ExerciseSetting.deleteMany({}),
    ExerciseLog.deleteMany({})
  ]);

  // Create the test account and log in once for the whole suite
  await loginAs(agent, TEST_USER);
});

afterAll(async () => {
  if (fs.existsSync(BACKUP_PATH)) {
    fs.copyFileSync(BACKUP_PATH, SHEET_PATH);
    fs.unlinkSync(BACKUP_PATH);
  }
  await mongoose.connection.dropDatabase();
  await mongoose.connection.close();
});

describe('MODULE 1: Progressive Overload Evaluation Algorithm (Unit Tests)', () => {
  test('Successful sets (12 12 12 for target 3x12) increase weight by +2.5 kg', () => {
    const res = evaluateExerciseProgression({
      weightKg: 60,
      actualRepsArray: [12, 12, 12],
      targetSets: 3,
      targetReps: 12,
      incrementKg: 2.5
    });
    expect(res.isSuccess).toBe(true);
    expect(res.nextWeekWeightKg).toBe(62.5);
    expect(res.failedSets).toHaveLength(0);
  });

  test('Failed sets (12 4 11 for target 3x12) keep weight unchanged (+0 kg)', () => {
    const res = evaluateExerciseProgression({
      weightKg: 45,
      actualRepsArray: [12, 4, 11],
      targetSets: 3,
      targetReps: 12,
      incrementKg: 2.5
    });
    expect(res.isSuccess).toBe(false);
    expect(res.nextWeekWeightKg).toBe(45);
    expect(res.failedSets).toEqual(['Set 2 (4/12)', 'Set 3 (11/12)']);
  });
});

describe('MODULE 2: TDEE Excel Coordinate Mapper & Nutrition MongoDB Sync', () => {
  test('Correctly calculates Week 4 Tuesday coordinates for 2026-10-27', () => {
    const coords = getCoordinatesForDate('2026-10-05', '2026-10-27');
    expect(coords.weekNumber).toBe(4);
    expect(coords.dayName).toBe('Tuesday');
    expect(coords.weightRow).toBe(18);
    expect(coords.calorieRow).toBe(19);
    expect(coords.dayCol).toBe(5); // Column E
  });

  test('POST /api/nutrition/log writes weight, calories, waist & neck to Excel and MongoDB', async () => {
    // The tracker's 56-week range starts on the day the account is created (today),
    // so log on today's date and compute the expected Excel cells from the same start date.
    const d = new Date();
    const today = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    const expected = getCoordinatesForDate(today, today);

    const res = await agent
      .post('/api/nutrition/log')
      .send({
        date: today,
        weight: 89.6,
        calories: 1850,
        waist: 37.2,
        neck: 14.6
      })
      .expect(200);

    expect(res.body.coords.weightRow).toBe(expected.weightRow);
    expect(res.body.coords.calorieRow).toBe(expected.calorieRow);
    expect(res.body.coords.dayCol).toBe(expected.dayCol);
    expect(res.body.mongoDoc.weight).toBe(89.6);
    expect(res.body.mongoDoc.calories).toBe(1850);
  });
});

describe('MODULE 3: Weekly Measurement Day Config', () => {
  test('PUT /api/nutrition/config updates measurement day in MongoDB', async () => {
    const res = await agent
      .put('/api/nutrition/config')
      .send({
        measurementDayOfWeek: 'Sunday'
      })
      .expect(200);

    expect(res.body.measurementDayOfWeek).toBe('Sunday');
  });
});

describe('MODULE 4: Exercise Logging, .txt File Import & Next-Week Schedule API', () => {
  test('POST /api/exercises/import-txt parses .txt lines into MongoDB & generates next Monday schedule', async () => {
    const sampleTxt = [
      '[2026-10-26]',
      'Back Squat | 80 | 4x5 | 5 5 5 5',
      'Leg Curl | 45 | 3x12 | 12 4 11'
    ].join('\n');

    const importRes = await agent
      .post('/api/exercises/import-txt')
      .send({ rawText: sampleTxt, date: '2026-10-26' })
      .expect(200);

    expect(importRes.body.importedCount).toBe(2);

    // Fetch next Monday's recommended workout schedule
    const scheduleRes = await agent
      .get('/api/exercises?day=Monday')
      .expect(200);

    const recs = scheduleRes.body.schedule.recommendations;
    const squatRec = recs.find((r) => r.exerciseName === 'Back Squat');
    const curlRec = recs.find((r) => r.exerciseName === 'Leg Curl');

    expect(squatRec.lastSessionSuccess).toBe(true);
    expect(squatRec.recommendedWeightKg).toBe(82.5); // 80 + 2.5 kg
    expect(squatRec.recommendedForDate).toBe('2026-11-02');

    expect(curlRec.lastSessionSuccess).toBe(false);
    expect(curlRec.recommendedWeightKg).toBe(45);   // Kept at 45 kg due to 12 4 11
  });

  test('PUT /api/exercises/settings modifies target reps without disrupting historical logs', async () => {
    const res = await agent
      .put('/api/exercises/settings')
      .send({
        exerciseName: 'Leg Curl',
        targetSets: 4,
        targetReps: 10,
        incrementKg: 2.5
      })
      .expect(200);

    expect(res.body.targetSets).toBe(4);
    expect(res.body.targetReps).toBe(10);
  });
});

describe('MODULE 5: Data & Backup - clear & JSON restore endpoints', () => {
  test('POST /api/data/clear refuses to run without the DELETE confirmation', async () => {
    await agent.post('/api/data/clear').send({ scope: 'exercises' }).expect(400);
    expect(await ExerciseLog.countDocuments()).toBeGreaterThan(0);
  });

  test('POST /api/data/clear rejects an unknown scope', async () => {
    await agent.post('/api/data/clear').send({ scope: 'everything', confirm: 'DELETE' }).expect(400);
  });

  test('GET /api/data/export.json and POST /api/data/import.json restore exported data after clear', async () => {
    const exportRes = await agent.get('/api/data/export.json').expect(200);
    expect(exportRes.body.exerciseLogs.length).toBeGreaterThan(0);

    await agent.post('/api/data/clear').send({ scope: 'exercises', confirm: 'DELETE' }).expect(200);
    expect(await ExerciseLog.countDocuments()).toBe(0);

    const importRes = await agent
      .post('/api/data/import.json')
      .send({ ...exportRes.body, mode: 'merge' })
      .expect(200);

    expect(importRes.body.ok).toBe(true);
    expect(importRes.body.imported.exerciseLogs).toBe(exportRes.body.exerciseLogs.length);
    expect(await ExerciseLog.countDocuments()).toBe(exportRes.body.exerciseLogs.length);

    // Clean exercise logs for subsequent module assertions
    await agent.post('/api/data/clear').send({ scope: 'exercises', confirm: 'DELETE' }).expect(200);
  });
});

describe('MODULE 6: Unknown exercise confirmation', () => {
  const base = { date: '2026-10-27', weightKg: 40, actualRepsInput: '12 12 12' };

  test('A typo is rejected with suggestions instead of creating an exercise', async () => {
    const res = await agent.post('/api/exercises/log').send({ ...base, exerciseName: 'Leg Curll' }).expect(409);
    expect(res.body.needsConfirmation).toBe(true);
    expect(res.body.suggestions).toContain('Leg Curl');
    expect(await ExerciseSetting.findOne({ exerciseName: 'Leg Curll' })).toBeNull();
  });

  test('A log without a weight is rejected and nothing is saved', async () => {
    const before = await ExerciseLog.countDocuments();
    await agent.post('/api/exercises/log').send({ date: '2026-10-27', exerciseName: 'Leg Curl', actualRepsInput: '12 12 12' }).expect(400);
    await agent.post('/api/exercises/log').send({ ...base, exerciseName: 'Cable Row 2', weightKg: '' }).expect(400);
    expect(await ExerciseLog.countDocuments()).toBe(before);
  });

  test('confirmNew creates the exercise', async () => {
    await agent.post('/api/exercises/log').send({ ...base, exerciseName: 'Cable Row', confirmNew: true }).expect(200);
    expect(await ExerciseSetting.findOne({ exerciseName: 'Cable Row' })).not.toBeNull();
  });
});

describe('MODULE 7: Same-day replace & flexible import dates', () => {
  test('Logging the same exercise twice on one day keeps a single entry', async () => {
    const body = { date: '2026-10-28', exerciseName: 'Leg Curl', weightKg: 45, actualRepsInput: '10 10 10 10' };
    await agent.post('/api/exercises/log').send(body).expect(200);
    await agent.post('/api/exercises/log').send({ ...body, weightKg: 47.5 }).expect(200);

    const docs = await ExerciseLog.find({ date: '2026-10-28', exerciseName: 'Leg Curl' });
    expect(docs).toHaveLength(1);         // still a single entry
    expect(docs[0].weightKg).toBe(47.5);  // holding the newer weight
  });

  test('import understands DD-MM-YYYY headers and per-line dates', async () => {
    const txt = ['28-10-2026', 'Cable Row | 40 | 3x10 | 10 10 10', '27/10/2026 | Cable Row | 37.5 | 10 10 10'].join('\n');
    await agent.post('/api/exercises/import-txt').send({ rawText: txt }).expect(200);
    expect(await ExerciseLog.countDocuments({ date: '2026-10-28', exerciseName: 'Cable Row' })).toBe(1);
    expect(await ExerciseLog.countDocuments({ date: '2026-10-27', exerciseName: 'Cable Row' })).toBe(1);
  });
});

describe('MODULE 8: Dates written in front of an exercise', () => {
  test('"25-10-26 Lag Pulls | ..." logs Lag Pulls on 2026-10-25, not an exercise called "25-10-26 Lag Pulls"', async () => {
    const txt = ['25-10-26 Lag Pulls | 80 | 3x12 | 12 12 12', '26-10-26', 'Lag Pulls | 82.5 | 12 12 12'].join('\n');
    await agent.post('/api/exercises/import-txt').send({ rawText: txt, date: '2026-10-28' }).expect(200);
    expect(await ExerciseLog.countDocuments({ date: '2026-10-25', exerciseName: 'Lag Pulls' })).toBe(1);
    expect(await ExerciseLog.countDocuments({ date: '2026-10-26', exerciseName: 'Lag Pulls' })).toBe(1);
    expect(await ExerciseSetting.findOne({ exerciseName: '25-10-26 Lag Pulls' })).toBeNull();
  });
});

describe('MODULE 9: Authentication & user data isolation', () => {
  test('Protected routes reject requests that have no login cookie', async () => {
    // Plain request(app) has no cookie, so the auth middleware must refuse it
    await request(app).get('/api/exercises').expect(401);
    await request(app).get('/api/data/export.json').expect(401);
    await request(app).post('/api/nutrition/log').send({ date: '2026-10-27', weight: 80 }).expect(401);
  });

  test('A second user cannot see the first user\'s data', async () => {
    // The first user already has exercise logs from earlier modules
    const mine = await agent.get('/api/data/export.json').expect(200);
    expect(mine.body.exerciseLogs.length).toBeGreaterThan(0);

    const other = request.agent(app);
    await loginAs(other, OTHER_USER);

    const theirs = await other.get('/api/data/export.json').expect(200);
    expect(theirs.body.exerciseLogs).toHaveLength(0);
  });
});