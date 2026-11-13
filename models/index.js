// models/index.js
const mongoose = require('mongoose');

// 0. User Account Schema
const UserSchema = new mongoose.Schema({
  username: { type: String, required: true, unique: true, trim: true, lowercase: true },
  passwordHash: { type: String, required: true },
  sheetConfig: {
    startDate: { type: String, default: null },
    weightUnit: { type: String, default: 'kg' },
    energyUnit: { type: String, default: 'Cal.' },
    startingWeight: { type: Number, default: null },
    goalWeight: { type: Number, default: null },
    weeklyGoalRate: { type: Number, default: null },
    gender: { type: String, default: 'Male' },
    measureUnit: { type: String, default: 'Inch' },
    height: { type: Number, default: null },
    tdeeAvgWeeks: { type: Number, default: 12 }
  }
}, { timestamps: true });

// 1. Application Configuration Schema (One Document Per User)
const AppConfigSchema = new mongoose.Schema({
  userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true, unique: true },
  measurementDayOfWeek: {
    type: String,
    enum: ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
    default: 'Monday'
  }
}, { timestamps: true });

// 2. Daily Nutrition & TDEE Sheet Sync Audit Schema (Per User)
const NutritionLogSchema = new mongoose.Schema({
  userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true, index: true },
  date: { type: String, required: true }, // YYYY-MM-DD
  weekNumber: { type: Number, required: true },
  dayName: { type: String, required: true },
  weight: { type: Number, default: null },
  calories: { type: Number, default: null },
  waist: { type: Number, default: null },
  neck: { type: Number, default: null },
  hip: { type: Number, default: null },
  sourceApp: { type: String, default: 'Manual' },
  excelCoordinates: {
    weightRow: Number,
    calorieRow: Number,
    dayCol: Number
  }
}, { timestamps: true });

NutritionLogSchema.index({ userId: 1, date: 1 }, { unique: true });

// 3. Exercise Target Settings & Saved Weight Schema (Per User)
const ExerciseSettingSchema = new mongoose.Schema({
  userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true, index: true },
  exerciseName: { type: String, required: true, trim: true },
  targetSets: { type: Number, required: true, default: 3, min: 1 },
  targetReps: { type: Number, required: true, default: 12, min: 1 },
  incrementKg: { type: Number, required: true, default: 2.5, min: 0 },
  currentWeightKg: { type: Number, default: null, min: 0 }
}, { timestamps: true });

ExerciseSettingSchema.index({ userId: 1, exerciseName: 1 }, { unique: true });

// 4. Exercise Workout Session & Progressive Overload Log Schema (Per User)
const ExerciseLogSchema = new mongoose.Schema({
  userId: { type: mongoose.Schema.Types.ObjectId, ref: 'User', required: true, index: true },
  date: { type: String, required: true },         // YYYY-MM-DD
  dayOfWeek: { type: String, required: true },    // e.g., 'Monday'
  nextWeekDate: { type: String, required: true }, // YYYY-MM-DD (+7 days)
  exerciseName: { type: String, required: true, trim: true },
  weightKg: { type: Number, required: true },
  targetSets: { type: Number, required: true, min: 1 },
  targetReps: { type: Number, required: true, min: 1 },
  actualReps: { type: [Number], required: true },
  isSuccess: { type: Boolean, required: true },
  nextWeekWeightKg: { type: Number, required: true },
  evaluationReason: { type: String, required: true },
  failedSets: { type: [String], default: [] },
  notes: { type: String, default: '' }
}, { timestamps: true });

ExerciseLogSchema.index({ userId: 1, date: 1, exerciseName: 1 });

module.exports = {
  User: mongoose.model('User', UserSchema),
  AppConfig: mongoose.model('AppConfig', AppConfigSchema),
  NutritionLog: mongoose.model('NutritionLog', NutritionLogSchema),
  ExerciseSetting: mongoose.model('ExerciseSetting', ExerciseSettingSchema),
  ExerciseLog: mongoose.model('ExerciseLog', ExerciseLogSchema)
};