// services/excelImportService.js
// Reads settings + logged entries out of a TDEE workbook worksheet (layout of "TDEE variant with bf 3.06.xlsx").
// Pure parsing - it never touches the database or the user's own sheet.

const DAYS_UTC = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const WEEKS = 56;

function unwrap(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if ('result' in v) return unwrap(v.result);          // formula cell -> cached result
    if ('richText' in v) return v.richText.map((t) => t.text).join('');
    return null;                                          // error cells, hyperlinks, etc.
  }
  if (v === '') return null;
  return v;
}

function toNumber(v) {
  if (v === null || v === undefined || v instanceof Date) return null;
  const n = Number(String(v).replace(',', '.'));
  return Number.isFinite(n) ? n : null;
}

function toText(v) {
  if (v === null || v === undefined || v instanceof Date) return null;
  const s = String(v).trim();
  if (!s || /choose/i.test(s)) return null;               // dropdown placeholder " - Choose - "
  return s;
}

function toYMD(v) {
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v.toISOString().split('T')[0];
  if (typeof v === 'string') {
    const m = v.trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  }
  return null;
}

function addDaysYMD(ymd, days) {
  const [y, m, d] = ymd.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, d + days)).toISOString().split('T')[0];
}

function dayNameOf(ymd) {
  const [y, m, d] = ymd.split('-').map(Number);
  return DAYS_UTC[new Date(Date.UTC(y, m - 1, d)).getUTCDay()];
}


// ---- Layout validation -------------------------------------------------------------------
// Labels that must sit in fixed cells for the file to count as a TDEE tracker workbook.
const REQUIRED_LABELS = [
  { row: 3,  col: 2,  text: 'start date',      shown: 'B3 "Input Start Date to Begin:"' },
  { row: 4,  col: 2,  text: 'lb or kg',        shown: 'B4 "Weights in Lb or Kg?"' },
  { row: 6,  col: 2,  text: 'starting weight', shown: 'B6 "Starting Weight:"' },
  { row: 7,  col: 2,  text: 'goal weight',     shown: 'B7 "Goal Weight:"' },
  { row: 11, col: 2,  text: 'week',            shown: 'B11 "Week"' },
  { row: 11, col: 15, text: 'waist',           shown: 'O11 "waist"' },
  { row: 11, col: 16, text: 'neck',            shown: 'P11 "neck"' }
];

function checkTdeeLayout(sheet) {
  const missing = [];
  for (const r of REQUIRED_LABELS) {
    const v = unwrap(sheet.getCell(r.row, r.col).value);
    const text = v === null || v instanceof Date ? '' : String(v).toLowerCase();
    if (!text.includes(r.text)) missing.push(r.shown);
  }
  return { ok: missing.length === 0, missing };
}

// Prefers a worksheet named "TDEE"; otherwise takes the first sheet that passes the layout check.
function findTdeeWorksheet(workbook) {
  const named = workbook.worksheets.filter((w) => String(w.name).trim().toLowerCase() === 'tdee');
  const candidates = [...named, ...workbook.worksheets.filter((w) => !named.includes(w))];
  let firstFailure = null;
  for (const sheet of candidates) {
    const check = checkTdeeLayout(sheet);
    if (check.ok) return sheet;
    if (!firstFailure) firstFailure = { sheet: sheet.name, ...check };
  }
  const err = new Error(
    'This file is not a TDEE tracker workbook. ' +
    (firstFailure
      ? `Sheet "${firstFailure.sheet}" is missing: ${firstFailure.missing.slice(0, 3).join(', ')}${firstFailure.missing.length > 3 ? ', ...' : ''}.`
      : 'It has no worksheets.') +
    ' Import the "TDEE variant with bf" workbook, or use a JSON backup instead.'
  );
  err.status = 422;
  throw err;
}

// ---- Value validation ----------------------------------------------------------------------
const LIMITS = {
  weight: [20, 700],        // kg or lb
  calories: [100, 40000],   // Cal or kJ
  measure: [5, 300],        // inch or cm (waist / neck / hip)
  height: [20, 300],
  weeklyRate: [0, 10],
  avgWeeks: [1, 56]
};
const inRange = (n, [lo, hi]) => n !== null && n >= lo && n <= hi;

function readTdeeWorksheet(sheet, { measurementDay = 'Monday' } = {}) {
  const layout = checkTdeeLayout(sheet);
  if (!layout.ok) {
    const err = new Error(`Worksheet "${sheet.name}" does not match the TDEE tracker layout (missing ${layout.missing[0]}).`);
    err.status = 422;
    throw err;
  }

  const cell = (row, col) => unwrap(sheet.getCell(row, col).value);
  const warnings = [];

  const startDate = toYMD(cell(3, 6));                    // F3
  if (!startDate) {
    throw new Error('Could not find a Start Date in cell F3 of the TDEE worksheet.');
  }
  const year = Number(startDate.slice(0, 4));
  if (year < 2000 || year > 2100) {
    throw new Error(`Start Date ${startDate} is not a realistic date.`);
  }

  // ---- Settings (a value that is blank or out of range is skipped, never imported) ----
  const config = { startDate };
  const unitRaw = toText(cell(4, 6));                     // F4
  if (unitRaw) config.weightUnit = unitRaw.toLowerCase() === 'lb' ? 'Lb' : 'kg';
  const energy = toText(cell(5, 6));                      // F5
  if (energy) config.energyUnit = energy;

  const setting = (key, row, col, limits, label) => {
    const n = toNumber(cell(row, col));
    if (n === null) return;
    if (inRange(n, limits)) config[key] = n;
    else warnings.push(`${label} (${n}) looked unrealistic and was skipped.`);
  };
  setting('startingWeight', 6, 6, LIMITS.weight, 'Starting Weight');        // F6
  setting('goalWeight', 7, 6, LIMITS.weight, 'Goal Weight');                // F7
  setting('weeklyGoalRate', 8, 6, LIMITS.weeklyRate, 'Weekly Rate');        // F8
  if (config.weeklyGoalRate === 0) delete config.weeklyGoalRate;
  const gender = toText(cell(3, 19));                     // S3
  if (gender && /^(male|female)$/i.test(gender)) config.gender = gender.charAt(0).toUpperCase() + gender.slice(1).toLowerCase();
  const measureUnit = toText(cell(4, 19));                // S4
  if (measureUnit && /^(inch|cm)$/i.test(measureUnit)) config.measureUnit = measureUnit.toLowerCase() === 'cm' ? 'cm' : 'Inch';
  setting('height', 5, 19, LIMITS.height, 'Height');                        // S5
  setting('tdeeAvgWeeks', 8, 19, LIMITS.avgWeeks, 'TDEE averaging weeks');  // S8

  // ---- Daily entries ----
  const byDate = new Map();
  const entryFor = (date) => {
    if (!byDate.has(date)) byDate.set(date, { date });
    return byDate.get(date);
  };

  let filledCells = 0;
  let rejectedCells = 0;
  const accept = (raw, limits, target, field) => {
    if (raw === null || raw === 0) return;                // empty cell
    filledCells++;
    if (inRange(raw, limits)) target[field] = raw;
    else rejectedCells++;
  };

  let measurementWeeks = 0;
  for (let w = 0; w < WEEKS; w++) {
    const weightRow = 12 + w * 2;
    const calorieRow = 13 + w * 2;

    for (let d = 0; d < 7; d++) {
      const date = addDaysYMD(startDate, w * 7 + d);
      const weight = toNumber(cell(weightRow, 4 + d));
      const calories = toNumber(cell(calorieRow, 4 + d));
      const hold = {};
      accept(weight, LIMITS.weight, hold, 'weight');
      accept(calories, LIMITS.calories, hold, 'calories');
      if (Object.keys(hold).length) Object.assign(entryFor(date), hold);
    }

    // Waist / neck / hip are stored once per week, on the weight row (columns O, P, Q)
    const hold = {};
    accept(toNumber(cell(weightRow, 15)), LIMITS.measure, hold, 'waist');
    accept(toNumber(cell(weightRow, 16)), LIMITS.measure, hold, 'neck');
    accept(toNumber(cell(weightRow, 17)), LIMITS.measure, hold, 'hip');
    if (Object.keys(hold).length) {
      // Attach them to the day matching the user's measurement day (fallback: first day of the week)
      let target = addDaysYMD(startDate, w * 7);
      for (let d = 0; d < 7; d++) {
        const date = addDaysYMD(startDate, w * 7 + d);
        if (dayNameOf(date) === measurementDay) { target = date; break; }
      }
      Object.assign(entryFor(target), hold);
      measurementWeeks++;
    }
  }

  if (rejectedCells > 0) {
    if (filledCells > 0 && rejectedCells / filledCells > 0.5) {
      const err = new Error(`Most values in this sheet (${rejectedCells} of ${filledCells}) are outside realistic weight, calorie or measurement ranges, so nothing was imported.`);
      err.status = 422;
      throw err;
    }
    warnings.push(`${rejectedCells} unrealistic value${rejectedCells === 1 ? ' was' : 's were'} skipped.`);
  }

  const entries = [...byDate.values()].sort((a, b) => a.date.localeCompare(b.date));
  return { config, entries, warnings, stats: { days: entries.length, measurementWeeks, rejectedCells } };
}

module.exports = { readTdeeWorksheet, findTdeeWorksheet, checkTdeeLayout };