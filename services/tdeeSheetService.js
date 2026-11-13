// services/tdeeSheetService.js
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { User, NutritionLog } = require('../models');

const TEMPLATE_PATH = path.join(__dirname, '..', 'TDEE variant with bf 3.06.xlsx');
const USER_SHEETS_DIR = path.join(__dirname, '..', 'user_sheets');
const SHEET_NAME = 'TDEE';

const CAL_PER_LB = 3500;
const LB_PER_KG = 2.20462;
const KJ_PER_CAL = 4.184;
const DAYS_UTC = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function normalizeWeightUnit(unit) {
  const clean = String(unit || '').trim().toLowerCase();
  if (clean === 'lb' || clean === 'lbs') return 'lbs';
  return 'kg';
}

function getUserSheetPath(userId) {
  if (!userId) throw new Error('userId is required to access a user TDEE spreadsheet.');
  return path.join(USER_SHEETS_DIR, `tdee_${String(userId)}.xlsx`);
}

function parseUTCDate(input) {
  if (!input) return NaN;
  if (input instanceof Date) {
    const t = input.getTime();
    if (Number.isNaN(t)) return NaN;
    return Date.UTC(input.getUTCFullYear(), input.getUTCMonth(), input.getUTCDate());
  }
  const str = String(input).trim();
  const match = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!match) return NaN;
  const y = Number(match[1]), m = Number(match[2]), d = Number(match[3]);
  if (m < 1 || m > 12 || d < 1 || d > 31) return NaN;
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return NaN;
  return dt.getTime();
}

function formatUTCDate(ms) {
  const d = new Date(ms);
  return d.toISOString().split('T')[0];
}

function extractCellValue(cellVal) {
  if (cellVal === null || cellVal === undefined) return null;
  if (typeof cellVal === 'object') {
    if (cellVal instanceof Date) return cellVal;
    if ('result' in cellVal) {
      const res = cellVal.result;
      if (res === '' || res === null || res === undefined) return null;
      if (typeof res === 'object' && 'error' in res) return null;
      return res;
    }
  }
  if (cellVal === '') return null;
  return cellVal;
}

function setFormulaResult(sheet, row, col, resultVal) {
  const cell = typeof row === 'string' ? sheet.getCell(row) : sheet.getCell(row, col);
  const val = cell.value;
  const normalizedResult = resultVal === null || resultVal === undefined ? '' : resultVal;
  if (val && typeof val === 'object' && ('formula' in val || 'sharedFormula' in val)) {
    cell.value = { ...val, result: normalizedResult };
  } else if (typeof val === 'string' && val.startsWith('=')) {
    cell.value = { formula: val.slice(1), result: normalizedResult };
  } else {
    cell.value = resultVal;
  }
}

function mround(val, multiple) {
  if (val === null || val === undefined || isNaN(val) || !multiple) return 0;
  return Math.round(val / multiple) * multiple;
}

function computeBodyFat({ gender, measureUnit, height, waist, neck, hip }) {
  if (!waist || !neck || !height) return null;
  let w = Number(waist), n = Number(neck), h = Number(height), hp = hip ? Number(hip) : 0;
  if (w <= 0 || n <= 0 || h <= 0) return null;
  const isCm = String(measureUnit).toLowerCase() === 'cm';
  if (isCm) { w /= 2.54; n /= 2.54; h /= 2.54; hp /= 2.54; }
  const g = String(gender).toLowerCase();
  if (g === 'male') {
    if (w - n <= 0) return null;
    return Math.round(86.01 * Math.log10(w - n) - 70.041 * Math.log10(h) + 36.76) / 100;
  }
  if (g === 'female') {
    if (!hp || w + hp - n <= 0) return null;
    return Math.round(163.205 * Math.log10(w + hp - n) - 97.684 * Math.log10(h) - 78.387) / 100;
  }
  return null;
}

function applyConfigToWorksheet(sheet, updates) {
  const mapping = {
    startDate: { row: 3, col: 6, transform: v => { const [y, m, d] = String(v).split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); } },
    weightUnit: { row: 4, col: 6, transform: v => normalizeWeightUnit(v) === 'lbs' ? 'Lb' : 'Kg' },
    energyUnit: { row: 5, col: 6, transform: v => String(v) },
    startingWeight: { row: 6, col: 6, transform: v => (v === null || v === '') ? null : Number(v) },
    goalWeight: { row: 7, col: 6, transform: v => (v === null || v === '') ? null : Number(v) },
    weeklyGoalRate: { row: 8, col: 6, transform: v => (v === null || v === '') ? null : Math.abs(Number(v)) },
    gender: { row: 3, col: 19, transform: v => String(v) },
    measureUnit: { row: 4, col: 19, transform: v => String(v) },
    height: { row: 5, col: 19, transform: v => (v === null || v === '') ? null : Number(v) },
    tdeeAvgWeeks: { row: 8, col: 19, transform: v => Math.max(1, Number(v)) }
  };

  const changed = [];
  for (const [key, value] of Object.entries(updates || {})) {
    if (value === undefined) continue;
    const m = mapping[key];
    if (!m) continue;
    const val = m.transform ? m.transform(value) : value;
    sheet.getCell(m.row, m.col).value = val;
    changed.push(key);
  }
  return changed;
}

function wipeWorksheetEntries(sheet, { includeMeasurements = true } = {}) {
  for (let w = 0; w < 56; w++) {
    const wRow = 12 + w * 2, cRow = 13 + w * 2;
    for (let col = 4; col <= 10; col++) {
      sheet.getCell(wRow, col).value = null;
      sheet.getCell(cRow, col).value = null;
    }
    if (includeMeasurements) {
      for (const col of [15, 16, 17]) sheet.getCell(wRow, col).value = null;
    }
  }
}

function recalculateAndRefreshSheet(sheet) {
  const startDateRaw = extractCellValue(sheet.getCell('F3').value);
  const startMs = parseUTCDate(startDateRaw) || parseUTCDate('2026-10-05');
  const weightUnitRaw = String(extractCellValue(sheet.getCell('F4').value) || 'kg');
  const weightUnit = normalizeWeightUnit(weightUnitRaw);
  const isLb = weightUnit === 'lbs';
  const unitLabel = isLb ? 'Lb' : 'Kg';

  const energyUnitRaw = String(extractCellValue(sheet.getCell('F5').value) || 'Cal.');
  const isCal = energyUnitRaw.toLowerCase() === 'cal.';

  const startingWeight = Number(extractCellValue(sheet.getCell('F6').value) || 0);
  const goalWeight = Number(extractCellValue(sheet.getCell('F7').value) || 0);
  const weeklyGoalRate = Math.abs(Number(extractCellValue(sheet.getCell('F8').value) || 0));

  const gender = String(extractCellValue(sheet.getCell('S3').value) || 'Male');
  const measureUnit = String(extractCellValue(sheet.getCell('S4').value) || 'Inch');
  const height = Number(extractCellValue(sheet.getCell('S5').value) || 0);
  const tdeeAvgWeeks = Math.max(1, Number(extractCellValue(sheet.getCell('S8').value) || 12));

  const av4 = isLb ? 1 : 2;
  const av5 = isCal ? 3 : 5;
  const av6 = av4 + av5;

  const aw12 = CAL_PER_LB;
  const ax12 = aw12 * LB_PER_KG;
  const ay12 = aw12 * KJ_PER_CAL;
  const az12 = ax12 * KJ_PER_CAL;

  const aw13 = 13;
  const ax13 = aw13 * LB_PER_KG;
  const ay13 = aw13 * KJ_PER_CAL;
  const az13 = ax13 * KJ_PER_CAL;

  const av7 = av6 === 4 ? aw12 : av6 === 5 ? ax12 : av6 === 6 ? ay12 : az12;
  const aw7 = av6 === 4 ? aw13 : av6 === 5 ? ax13 : av6 === 6 ? ay13 : az13;
  const ai5 = mround(startingWeight * aw7, 5);

  setFormulaResult(sheet, 'AV4', null, av4);
  setFormulaResult(sheet, 'AW4', null, av4);
  setFormulaResult(sheet, 'AV5', null, av5);
  setFormulaResult(sheet, 'AW5', null, av5);
  setFormulaResult(sheet, 'AV6', null, av6);
  setFormulaResult(sheet, 'AW6', null, av6);
  setFormulaResult(sheet, 'AV7', null, av7);
  setFormulaResult(sheet, 'AW7', null, aw7);
  setFormulaResult(sheet, 'AI5', null, ai5);
  setFormulaResult(sheet, 'AM6', null, startingWeight);
  setFormulaResult(sheet, 'AI11', null, startingWeight);
  setFormulaResult(sheet, 'AJ11', null, goalWeight);

  const aj9 = startingWeight > goalWeight ? -weeklyGoalRate : startingWeight < goalWeight ? weeklyGoalRate : 0;
  setFormulaResult(sheet, 'AJ9', null, aj9);

  let prevASWeight = startingWeight;
  let prevASCals = ai5;
  let prevAIWeight = startingWeight;
  let prevAV = ai5;
  let prevAK = startingWeight;

  const mValues = [];
  const kWeights = [];

  for (let w = 0; w < 56; w++) {
    const wRow = 12 + w * 2;
    const cRow = 13 + w * 2;
    const weekDate = new Date(startMs + w * 7 * 86400000);

    setFormulaResult(sheet, wRow, 2, weekDate);
    setFormulaResult(sheet, wRow, 3, 'Weight');
    setFormulaResult(sheet, cRow, 3, energyUnitRaw);

    const rawW = [];
    const rawC = [];
    for (let d = 0; d < 7; d++) {
      const wv = extractCellValue(sheet.getCell(wRow, 4 + d).value);
      const cv = extractCellValue(sheet.getCell(cRow, 4 + d).value);
      rawW.push(wv !== null && wv !== undefined && wv !== '' ? Number(wv) : null);
      rawC.push(cv !== null && cv !== undefined && cv !== '' ? Number(cv) : null);
    }

    const rawWCount = rawW.filter(v => v !== null && !isNaN(v)).length;
    const rawCCount = rawC.filter(v => v !== null && !isNaN(v)).length;
    const hasAnyW = rawWCount > 0;
    const hasAnyC = rawCCount > 0;

    const carriedW = [];
    const carriedC = [];
    for (let d = 0; d < 7; d++) {
      if (!hasAnyW) {
        carriedW.push(null);
      } else if (d === 0) {
        carriedW.push(rawW[0] !== null ? rawW[0] : (prevASWeight !== null ? prevASWeight : startingWeight));
      } else {
        carriedW.push(rawW[d] !== null ? rawW[d] : carriedW[d - 1]);
      }
      setFormulaResult(sheet, wRow, 38 + d, carriedW[d]);

      if (!hasAnyC) {
        carriedC.push(null);
      } else if (d === 0) {
        carriedC.push(rawC[0] !== null ? rawC[0] : (prevASCals !== null ? prevASCals : ai5));
      } else {
        carriedC.push(rawC[d] !== null ? rawC[d] : carriedC[d - 1]);
      }
      setFormulaResult(sheet, cRow, 38 + d, carriedC[d]);
    }

    const ahW = carriedW.filter(v => v !== null && !isNaN(v)).length;
    const ahC = carriedC.filter(v => v !== null && !isNaN(v)).length;
    setFormulaResult(sheet, wRow, 34, ahW);
    setFormulaResult(sheet, cRow, 34, ahC);

    const asW = ahW > 0 ? carriedW.reduce((s, v) => s + v, 0) / ahW : null;
    const asC = ahC > 0 ? carriedC.reduce((s, v) => s + v, 0) / ahC : null;
    setFormulaResult(sheet, wRow, 45, asW);
    setFormulaResult(sheet, cRow, 45, asC);

    const kW = ahW > 0 ? asW : null;
    const kC = ahC > 0 ? asC : null;
    setFormulaResult(sheet, wRow, 11, kW);
    setFormulaResult(sheet, cRow, 11, kC);
    kWeights.push(kW);

    let atW = null;
    if (asW !== null) {
      const refW = w === 0 ? startingWeight : prevASWeight;
      if (refW !== null && refW !== undefined) {
        atW = asW - refW;
      }
    }
    setFormulaResult(sheet, wRow, 46, atW);
    setFormulaResult(sheet, wRow, 12, kW !== null ? atW : null);

    const aiW = kW !== null ? kW : prevAIWeight;
    setFormulaResult(sheet, wRow, 35, aiW);
    setFormulaResult(sheet, cRow, 35, aiW);
    setFormulaResult(sheet, wRow, 36, goalWeight);
    setFormulaResult(sheet, cRow, 36, goalWeight);

    let akW = goalWeight;
    if (startingWeight > goalWeight) {
      akW = prevAK > goalWeight ? prevAK + aj9 : goalWeight;
    } else if (startingWeight < goalWeight) {
      akW = prevAK < goalWeight ? prevAK + aj9 : goalWeight;
    }
    setFormulaResult(sheet, wRow, 37, akW);
    prevAK = akW;

    let avVal;
    if (w === 0) {
      if (ahW === 0 || ahC === 0 || asC === null || atW === null) {
        avVal = ai5;
      } else {
        avVal = asC + (((-atW) * av7) / ahC);
      }
    } else {
      if (ahC < 7 || ahW < 7 || kC === null || atW === null) {
        avVal = prevAV;
      } else {
        const rawWeekEnergy = kC + (((-atW) * av7) / ahC);
        const weekNum = w + 1;
        const windowSize = Math.min(tdeeAvgWeeks, weekNum);
        const prevCount = windowSize - 1;
        const sliceStart = Math.max(0, w - prevCount);
        const prevValidM = mValues
          .slice(sliceStart, w)
          .filter(v => typeof v === 'number' && !isNaN(v));
        const prevMSum = prevValidM.reduce((acc, v) => acc + v, 0);
        avVal = (rawWeekEnergy + prevMSum) / (prevValidM.length + 1);
      }
    }
    setFormulaResult(sheet, wRow, 48, avVal);

    const auVal = avVal !== null && atW !== null ? mround(avVal, 5) : null;
    const mVal = avVal !== null && atW !== null && kW !== null ? auVal : null;
    setFormulaResult(sheet, wRow, 47, auVal);
    setFormulaResult(sheet, wRow, 13, mVal);
    mValues.push(mVal);

    const waist = extractCellValue(sheet.getCell(wRow, 15).value);
    const neck = extractCellValue(sheet.getCell(wRow, 16).value);
    const hip = extractCellValue(sheet.getCell(wRow, 17).value);
    const bfDec = computeBodyFat({ gender, measureUnit, height, waist, neck, hip });
    setFormulaResult(sheet, wRow, 18, bfDec);
    setFormulaResult(sheet, wRow, 19, bfDec !== null && kW !== null ? kW * bfDec : null);
    setFormulaResult(sheet, wRow, 20, bfDec !== null && kW !== null ? kW * (1 - bfDec) : null);

    if (asW !== null) prevASWeight = asW;
    if (asC !== null) prevASCals = asC;
    prevAIWeight = aiW;
      prevAV = avVal;
  }

  for (let blockEndWeek = 3; blockEndWeek < 56; blockEndWeek += 4) {
    const endRow = 12 + blockEndWeek * 2;
    const startRow = endRow - 6;
    const mEnd = mValues[blockEndWeek];
    const kEnd = kWeights[blockEndWeek];
    const kPrevBlock = blockEndWeek === 3 ? startingWeight : kWeights[blockEndWeek - 4];

    if (mEnd !== null && kEnd !== null && kPrevBlock !== null) {
      const diff = Math.abs(kEnd - kPrevBlock);
      const status = kEnd === kPrevBlock ? 'No Weight Change' : kEnd > kPrevBlock ? 'You Gained:' : 'You Lost:';
      setFormulaResult(sheet, startRow + 2, 22, 'From');
      setFormulaResult(sheet, startRow + 4, 22, status);
      setFormulaResult(sheet, startRow + 4, 28, kEnd === kPrevBlock ? null : diff);
      setFormulaResult(sheet, startRow + 5, 22, kEnd === kPrevBlock ? null : 'At a Rate Of');
      setFormulaResult(sheet, startRow + 5, 27, kEnd === kPrevBlock ? null : diff / 4);
      setFormulaResult(sheet, endRow, 22, kEnd === kPrevBlock ? null : 'You have');
      setFormulaResult(sheet, endRow, 26, kEnd === kPrevBlock ? null : Math.abs(kEnd - goalWeight));
    }
  }

  const k122 = kWeights[55];
  setFormulaResult(sheet, 'K124', null, k122);
  const ai124 = mround(prevAIWeight, 0.5);
  const av124 = prevAV;
  setFormulaResult(sheet, 'AI124', null, ai124);
  setFormulaResult(sheet, 'AV124', null, av124);

  const validKWeights = kWeights.filter(v => v !== null);
  const exactCurrentWeight = validKWeights.length > 0 ? validKWeights[validKWeights.length - 1] : startingWeight;
  const l4 = ai124 || startingWeight;
  setFormulaResult(sheet, 'L4', null, l4);
  setFormulaResult(sheet, 'M4', null, unitLabel);
  setFormulaResult(sheet, 'G6', null, unitLabel);
  setFormulaResult(sheet, 'G7', null, unitLabel);
  setFormulaResult(sheet, 'G8', null, unitLabel);

  const h5 = l4 === startingWeight ? 'Your weight hasn’t changed.' : l4 > startingWeight ? "You've Gained" : "You've Lost";
  const l5 = l4 === startingWeight ? null : Math.abs(l4 - startingWeight);
  setFormulaResult(sheet, 'H5', null, h5);
  setFormulaResult(sheet, 'L5', null, l5);
  setFormulaResult(sheet, 'M5', null, l5 !== null ? unitLabel : null);

  const l6 = av124 ? mround(av124, 25) : mround(startingWeight * aw7, 25);
  setFormulaResult(sheet, 'L6', null, l6);
  setFormulaResult(sheet, 'M6', null, isCal ? 'Cal/Day' : 'Kj/Day');

  const f9 = mround((weeklyGoalRate * av7) / 7, 5);
  setFormulaResult(sheet, 'F9', null, f9);
  setFormulaResult(sheet, 'G9', null, isCal ? 'Cal.' : 'Kj');

  let l8;
  let b8Text = '';
  let b9Text = '';
  const isMaintenance = startingWeight === goalWeight || weeklyGoalRate === 0 || Math.abs(goalWeight - exactCurrentWeight) < 0.1;

  if (isMaintenance) {
    l8 = l6;
    b8Text = 'Maintenance Goal Rate:';
    b9Text = 'Target Daily Margin:';
  } else if (goalWeight > startingWeight) {
    l8 = l6 + f9;
    b8Text = 'Goal Weight Gain per Week:';
    b9Text = 'Target Daily Surplus:';
  } else {
    l8 = l6 - f9;
    b8Text = 'Goal Weight Loss per Week:';
    b9Text = 'Target Daily Deficit:';
  }

  setFormulaResult(sheet, 'B8', null, b8Text);
  setFormulaResult(sheet, 'B9', null, b9Text);
  setFormulaResult(sheet, 'L8', null, l8);
  setFormulaResult(sheet, 'M8', null, isCal ? 'Cal/Day' : 'Kj/Day');

  const af4 = weeklyGoalRate > 0 ? Math.ceil(((goalWeight - l4) / weeklyGoalRate) * 2) / 2 : 0;
  const af5 = weeklyGoalRate > 0 ? Math.ceil(((l4 - goalWeight) / weeklyGoalRate) * 2) / 2 : 0;
  setFormulaResult(sheet, 'AF4', null, af4);
  setFormulaResult(sheet, 'AF5', null, af5);
  const h9 = isMaintenance ? 0 : Math.max(0, startingWeight < goalWeight ? af4 : af5);
  setFormulaResult(sheet, 'H9', null, h9);

  // Timezone-safe UTC date marker for today
  const utcNow = new Date();
  const todayUtcDate = new Date(Date.UTC(utcNow.getUTCFullYear(), utcNow.getUTCMonth(), utcNow.getUTCDate()));
  setFormulaResult(sheet, 'L3', null, todayUtcDate);
  if (h9 > 0) {
    setFormulaResult(sheet, 'L7', null, new Date(todayUtcDate.getTime() + h9 * 7 * 86400000));
  } else {
    setFormulaResult(sheet, 'L7', null, null);
  }

  return {
    exactCurrentWeight: Number(exactCurrentWeight.toFixed(2)),
    sheetCurrentWeight: Number(l4.toFixed(2)),
    currentTdee: l6,
    targetDailyDeficit: isMaintenance ? 0 : f9,
    targetDailyCalories: l8
  };
}

function getCoordinatesForDate(startDateValue, targetDateStr) {
  const startMs = parseUTCDate(startDateValue);
  const targetMs = parseUTCDate(targetDateStr);

  if (Number.isNaN(startMs)) {
    throw new Error(`Invalid start date: "${startDateValue}". Must be a valid YYYY-MM-DD date.`);
  }
  if (Number.isNaN(targetMs)) {
    throw new Error(`Invalid target date: "${targetDateStr}". Must be a valid YYYY-MM-DD date.`);
  }

  const daysDiff = Math.floor((targetMs - startMs) / 86400000);

  if (daysDiff < 0 || daysDiff >= 56 * 7) {
    throw new Error(`Date ${targetDateStr} is outside the 56-week range starting ${formatUTCDate(startMs)}.`);
  }

  const weekIndex = Math.floor(daysDiff / 7);
  const dayIndex = daysDiff % 7;
  const startDow = new Date(startMs).getUTCDay();

  return {
    weekNumber: weekIndex + 1,
    weekIndex,
    dayIndex,
    dayName: DAYS_UTC[(startDow + dayIndex) % 7],
    weightRow: 12 + weekIndex * 2,
    calorieRow: 13 + weekIndex * 2,
    dayCol: 4 + dayIndex,
    waistCol: 15,
    neckCol: 16,
    hipCol: 17
  };
}

async function _ensureUserSheetUnlocked(userId) {
  await fs.promises.mkdir(USER_SHEETS_DIR, { recursive: true });
  const userSheetPath = getUserSheetPath(userId);

  if (fs.existsSync(userSheetPath)) {
    return userSheetPath;
  }

  if (!fs.existsSync(TEMPLATE_PATH)) {
    throw new Error(`Master TDEE spreadsheet template not found at ${TEMPLATE_PATH}`);
  }

  await fs.promises.copyFile(TEMPLATE_PATH, userSheetPath);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(userSheetPath);
  const sheet = workbook.getWorksheet(SHEET_NAME);
  if (!sheet) throw new Error(`Worksheet "${SHEET_NAME}" not found.`);

  wipeWorksheetEntries(sheet, { includeMeasurements: true });

  const [userDoc, logs] = await Promise.all([
    User.findById(userId).lean(),
    NutritionLog.find({ userId }).sort({ date: 1 }).lean()
  ]);

  if (userDoc?.sheetConfig && userDoc.sheetConfig.startDate) {
    applyConfigToWorksheet(sheet, userDoc.sheetConfig);
  }

  const startDateVal = extractCellValue(sheet.getCell('F3').value);
  for (const item of logs) {
    if (!item || !item.date) continue;
    try {
      const coords = getCoordinatesForDate(startDateVal, item.date);
      if (item.weight !== undefined && item.weight !== null) sheet.getCell(coords.weightRow, coords.dayCol).value = Number(item.weight);
      if (item.calories !== undefined && item.calories !== null) sheet.getCell(coords.calorieRow, coords.dayCol).value = Number(item.calories);
      if (item.waist !== undefined && item.waist !== null) sheet.getCell(coords.weightRow, coords.waistCol).value = Number(item.waist);
      if (item.neck !== undefined && item.neck !== null) sheet.getCell(coords.weightRow, coords.neckCol).value = Number(item.neck);
      if (item.hip !== undefined && item.hip !== null) sheet.getCell(coords.weightRow, coords.hipCol).value = Number(item.hip);
    } catch (_) {}
  }

  recalculateAndRefreshSheet(sheet);
  await workbook.xlsx.writeFile(userSheetPath);
  return userSheetPath;
}

async function _logToTdeeSheet(userId, { date, weight, calories, waist, neck, hip, clearEmpty = false, clearMeasurements = false }) {
  const userSheetPath = await _ensureUserSheetUnlocked(userId);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(userSheetPath);
  const sheet = workbook.getWorksheet(SHEET_NAME);
  if (!sheet) throw new Error(`Worksheet "${SHEET_NAME}" not found.`);

  const startDateVal = extractCellValue(sheet.getCell('F3').value);
  const coords = getCoordinatesForDate(startDateVal, date);

  const applyCell = (row, col, val, allowClear) => {
    if (val !== undefined && val !== null && val !== '') {
      sheet.getCell(row, col).value = Number(val);
    } else if (allowClear && (val === '' || val === null)) {
      sheet.getCell(row, col).value = null;
    }
  };

  applyCell(coords.weightRow, coords.dayCol, weight, clearEmpty);
  applyCell(coords.calorieRow, coords.dayCol, calories, clearEmpty);
  applyCell(coords.weightRow, coords.waistCol, waist, clearMeasurements);
  applyCell(coords.weightRow, coords.neckCol, neck, clearMeasurements);
  applyCell(coords.weightRow, coords.hipCol, hip, clearMeasurements);

  recalculateAndRefreshSheet(sheet);
  await workbook.xlsx.writeFile(userSheetPath);
  return coords;
}

async function _bulkLogToTdeeSheet(userId, entries = []) {
  const userSheetPath = await _ensureUserSheetUnlocked(userId);
  if (!entries.length) return 0;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(userSheetPath);
  const sheet = workbook.getWorksheet(SHEET_NAME);
  if (!sheet) throw new Error(`Worksheet "${SHEET_NAME}" not found.`);

  const startDateVal = extractCellValue(sheet.getCell('F3').value);
  let appliedCount = 0;

  for (const item of entries) {
    if (!item || !item.date) continue;
    try {
      const coords = getCoordinatesForDate(startDateVal, item.date);
      if (item.weight !== undefined && item.weight !== null) sheet.getCell(coords.weightRow, coords.dayCol).value = Number(item.weight);
      if (item.calories !== undefined && item.calories !== null) sheet.getCell(coords.calorieRow, coords.dayCol).value = Number(item.calories);
      if (item.waist !== undefined && item.waist !== null) sheet.getCell(coords.weightRow, coords.waistCol).value = Number(item.waist);
      if (item.neck !== undefined && item.neck !== null) sheet.getCell(coords.weightRow, coords.neckCol).value = Number(item.neck);
      if (item.hip !== undefined && item.hip !== null) sheet.getCell(coords.weightRow, coords.hipCol).value = Number(item.hip);
      appliedCount++;
    } catch (_) {}
  }

  recalculateAndRefreshSheet(sheet);
  await workbook.xlsx.writeFile(userSheetPath);
  return appliedCount;
}

async function _getTdeeSheetDashboard(userId) {
  const userSheetPath = await _ensureUserSheetUnlocked(userId);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(userSheetPath);
  const sheet = workbook.getWorksheet(SHEET_NAME);
  if (!sheet) throw new Error(`Worksheet "${SHEET_NAME}" not found.`);

  const refreshed = recalculateAndRefreshSheet(sheet);

  const startMs = parseUTCDate(extractCellValue(sheet.getCell('F3').value));
  const weightUnitRaw = String(extractCellValue(sheet.getCell('F4').value) || 'kg');
  const weightUnit = normalizeWeightUnit(weightUnitRaw);
  const energyUnit = String(extractCellValue(sheet.getCell('F5').value) || 'Cal.');
  const startingWeight = Number(extractCellValue(sheet.getCell('F6').value) || 0);
  const goalWeight = Number(extractCellValue(sheet.getCell('F7').value) || 0);
  const weeklyGoalRate = Math.abs(Number(extractCellValue(sheet.getCell('F8').value) || 0));

  const gender = String(extractCellValue(sheet.getCell('S3').value) || 'Male');
  const measureUnit = String(extractCellValue(sheet.getCell('S4').value) || 'Inch');
  const height = Number(extractCellValue(sheet.getCell('S5').value) || 0);
  const tdeeAvgWeeks = Number(extractCellValue(sheet.getCell('S8').value) || 12);

  const weeks = [];
  for (let w = 0; w < 56; w++) {
    const wRow = 12 + w * 2;
    const cRow = 13 + w * 2;
    const weekStartMs = startMs + w * 7 * 86400000;

    const weights = [];
    const calories = [];
    for (let d = 0; d < 7; d++) {
      const wVal = extractCellValue(sheet.getCell(wRow, 4 + d).value);
      const cVal = extractCellValue(sheet.getCell(cRow, 4 + d).value);
      weights.push(wVal !== null && wVal !== undefined ? Number(wVal) : null);
      calories.push(cVal !== null && cVal !== undefined ? Number(cVal) : null);
    }

    const waist = extractCellValue(sheet.getCell(wRow, 15).value);
    const neck = extractCellValue(sheet.getCell(wRow, 16).value);
    const hip = extractCellValue(sheet.getCell(wRow, 17).value);

    const kWeight = extractCellValue(sheet.getCell(wRow, 11).value);
    const lDelta = extractCellValue(sheet.getCell(wRow, 12).value);
    const kCal = extractCellValue(sheet.getCell(cRow, 11).value);
    const mTdee = extractCellValue(sheet.getCell(wRow, 13).value);
    const rBf = extractCellValue(sheet.getCell(wRow, 18).value);

    const avgWeight = kWeight !== null && kWeight !== undefined ? Number(Number(kWeight).toFixed(2)) : null;
    const deltaWeight = lDelta !== null && lDelta !== undefined ? Number(Number(lDelta).toFixed(2)) : null;
    const avgCalories = kCal !== null && kCal !== undefined ? Math.round(Number(kCal)) : null;
    const weekTdee = mTdee !== null && mTdee !== undefined ? Math.round(Number(mTdee)) : null;
    const bfPercent = rBf !== null && rBf !== undefined ? Math.round(Number(rBf) * 100) : null;

    if (weights.some(v => v !== null) || calories.some(v => v !== null) || waist || neck || hip || w < 4) {
      weeks.push({
        weekNumber: w + 1,
        weekStartDate: formatUTCDate(weekStartMs),
        weights,
        calories,
        avgWeight,
        deltaWeight,
        avgCalories,
        weekTdee,
        waist: waist !== null ? Number(waist) : null,
        neck: neck !== null ? Number(neck) : null,
        hip: hip !== null ? Number(hip) : null,
        bfPercent
      });
    }
  }

  const currentWeight = refreshed.exactCurrentWeight;
  const weightChange = Number((currentWeight - startingWeight).toFixed(2));
  const weightLost = Number((startingWeight - currentWeight).toFixed(2));

  let goalMode = 'loss';
  if (startingWeight === goalWeight || weeklyGoalRate === 0 || Math.abs(goalWeight - currentWeight) < 0.1) {
    goalMode = 'maintenance';
  } else if (goalWeight > startingWeight) {
    goalMode = 'gain';
  }

  const sheetConfig = {
    startDate: formatUTCDate(startMs),
    weightUnit,
    energyUnit,
    startingWeight,
    goalWeight,
    weeklyGoalRate,
    gender,
    measureUnit,
    height,
    tdeeAvgWeeks
  };

  return {
    summary: {
      ...sheetConfig,
      currentWeight,
      sheetCurrentWeight: refreshed.sheetCurrentWeight,
      weightLost,
      weightChange,
      goalMode,
      currentTdee: refreshed.currentTdee,
      targetDailyDeficit: refreshed.targetDailyDeficit,
      targetDailyCalories: refreshed.targetDailyCalories
    },
    sheetConfig,
    weeks
  };
}

async function _updateSheetConfig(userId, updates) {
  const userSheetPath = await _ensureUserSheetUnlocked(userId);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(userSheetPath);
  const sheet = workbook.getWorksheet(SHEET_NAME);
  if (!sheet) throw new Error(`Worksheet "${SHEET_NAME}" not found.`);

  if (updates.weightUnit) {
    updates.weightUnit = normalizeWeightUnit(updates.weightUnit);
  }

  const changed = applyConfigToWorksheet(sheet, updates);

  recalculateAndRefreshSheet(sheet);
  await workbook.xlsx.writeFile(userSheetPath);

  const setOps = {};
  for (const key of changed) {
    setOps[`sheetConfig.${key}`] = updates[key];
  }
  if (Object.keys(setOps).length > 0) {
    await User.findByIdAndUpdate(userId, { $set: setOps });
  }

  return { updated: changed };
}

async function _clearTdeeSheetEntries(userId, { includeMeasurements = true } = {}) {
  const userSheetPath = await _ensureUserSheetUnlocked(userId);

  const backupDir = path.join(USER_SHEETS_DIR, 'backups');
  await fs.promises.mkdir(backupDir, { recursive: true });
  const backupFile = `TDEE-${String(userId)}-backup-${new Date().toISOString().replace(/[:.]/g, '-')}.xlsx`;
  await fs.promises.copyFile(userSheetPath, path.join(backupDir, backupFile));

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(userSheetPath);
  const sheet = workbook.getWorksheet(SHEET_NAME);
  if (!sheet) throw new Error(`Worksheet "${SHEET_NAME}" not found.`);

  wipeWorksheetEntries(sheet, { includeMeasurements });

  recalculateAndRefreshSheet(sheet);
  await workbook.xlsx.writeFile(userSheetPath);
  return { backupFile };
}

const userSheetQueues = new Map();
function enqueueUserSheetTask(userId, task) {
  const key = String(userId);
  const prev = userSheetQueues.get(key) || Promise.resolve();
  const run = prev.then(task);
  userSheetQueues.set(key, run.catch(() => {}));
  return run;
}

const ensureUserSheet = (userId) => enqueueUserSheetTask(userId, () => _ensureUserSheetUnlocked(userId));
const logToTdeeSheet = (userId, entry) => enqueueUserSheetTask(userId, () => _logToTdeeSheet(userId, entry));
const bulkLogToTdeeSheet = (userId, entries) => enqueueUserSheetTask(userId, () => _bulkLogToTdeeSheet(userId, entries));
const getTdeeSheetDashboard = (userId) => enqueueUserSheetTask(userId, () => _getTdeeSheetDashboard(userId));
const clearTdeeSheetEntries = (userId, opts) => enqueueUserSheetTask(userId, () => _clearTdeeSheetEntries(userId, opts));
const updateSheetConfig = (userId, updates) => enqueueUserSheetTask(userId, () => _updateSheetConfig(userId, updates));

module.exports = {
  TEMPLATE_PATH,
  getUserSheetPath,
  normalizeWeightUnit,
  ensureUserSheet,
  clearTdeeSheetEntries,
  logToTdeeSheet,
  bulkLogToTdeeSheet,
  getTdeeSheetDashboard,
  getCoordinatesForDate,
  updateSheetConfig
};
