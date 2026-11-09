const APP_TITLE = 'DAILY CORE';
const SPREADSHEET_PROPERTY = 'DAILY_CORE_SPREADSHEET_ID';
const SHEET_NAME = 'Tasks';
const HEADERS = [
  'ID', '日付', 'タスク名', '期限', '完了', 'メモ', '作成日時', '更新日時',
  '推定所要時間（分）', '重要度', 'ロングタスク'
];
const START_TIME_PROPERTY_PREFIX = 'DAILY_CORE_START_TIME_';

function doGet(e) {
  if (e && e.parameter && e.parameter.mode === 'bridge') {
    return HtmlService.createHtmlOutputFromFile('bridge')
      .setTitle(APP_TITLE + ' Bridge')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  ensureSheet_();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_TITLE)
    .setFaviconUrl('https://morigu812-svg.github.io/DAILY-CORE/favicon.svg?v=20261110')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ensureSheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_PROPERTY);
  if (!spreadsheetId) throw new Error('スプレッドシートIDが未設定です。既存のスプレッドシートIDを確認してください。');
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('既存スプレッドシートに「' + SHEET_NAME + '」シートが見つかりません。');
  let lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  } else {
    const existing = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
    HEADERS.forEach(header => {
      if (existing.indexOf(header) === -1) {
        sheet.getRange(1, ++lastColumn).setValue(header);
        existing.push(header);
      }
    });
  }
  return sheet;
}

function getColumnMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i + 1; });
  HEADERS.forEach(h => { if (!map[h]) throw new Error('必要な列がありません: ' + h); });
  return map;
}

function normalizeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const text = String(value == null ? '' : value).trim();
  if (!text) return '';
  const match = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (!match) throw new Error('日付の形式が正しくありません。');
  const y = Number(match[1]), m = Number(match[2]), d = Number(match[3]);
  const date = new Date(y, m - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== m - 1 || date.getDate() !== d) {
    throw new Error('正しい日付を入力してください。');
  }
  return [String(y).padStart(4, '0'), String(m).padStart(2, '0'), String(d).padStart(2, '0')].join('-');
}

function validateDate_(value, label) {
  const result = normalizeDate_(value);
  if (!result) throw new Error((label || '日付') + 'を入力してください。');
  return result;
}

function cleanText_(value, maxLength) {
  const text = String(value == null ? '' : value).trim();
  if (maxLength && text.length > maxLength) throw new Error('入力できる文字数は最大' + maxLength + '文字です。');
  return text;
}

function normalizeDuration_(value) {
  if (value === '' || value === null || value === undefined) return 0;
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0 || !Number.isInteger(n)) throw new Error('推定所要時間は0以上の整数（分）で入力してください。');
  return n;
}

function normalizeImportance_(value) {
  const n = Number(value);
  return [1, 2, 3].indexOf(n) >= 0 ? n : 2;
}

function rowToTask_(row, columns) {
  const get = h => row[columns[h] - 1];
  const startDate = normalizeDate_(get('日付'));
  const rawDue = get('期限');
  const dueDate = rawDue ? normalizeDate_(rawDue) : startDate;
  return {
    id: String(get('ID') || ''),
    date: startDate,
    title: String(get('タスク名') || ''),
    dueDate: dueDate,
    completed: get('完了') === true || String(get('完了')).toLowerCase() === 'true',
    memo: String(get('メモ') || ''),
    createdAt: get('作成日時') instanceof Date ? get('作成日時').toISOString() : String(get('作成日時') || ''),
    updatedAt: get('更新日時') instanceof Date ? get('更新日時').toISOString() : String(get('更新日時') || ''),
    duration: normalizeDuration_(get('推定所要時間（分）')),
    importance: normalizeImportance_(get('重要度')),
    longTask: get('ロングタスク') === true || String(get('ロングタスク')).toLowerCase() === 'true'
  };
}

function getAllTasks_(sheet, diagnostics) {
  sheet = sheet || ensureSheet_();
  const columns = getColumnMap_(sheet), lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const tasks = [];
  rows.forEach((row, index) => {
    if (!row[columns.ID - 1]) return;
    try {
      tasks.push(rowToTask_(row, columns));
    } catch (error) {
      if (diagnostics) diagnostics.invalidRows++;
      console.warn('DAILY CORE: invalid task at row ' + (index + 2));
    }
  });
  return tasks;
}

function taskVisibleOnDate_(task, date) {
  const start = task.date || task.dueDate;
  const due = task.dueDate || task.date;
  if (!start || !due || due < start) return date === start;
  if (task.longTask) return date >= start && date <= due;
  return date === due;
}

function getTasks(date) {
  const selected = validateDate_(date, '対象日');
  return getAllTasks_().filter(task => taskVisibleOnDate_(task, selected));
}

function buildMonthTasks_(all, year, month, countsOnly) {
  const prefix = String(year).padStart(4, '0') + '-' + String(month).padStart(2, '0');
  const result = {}, lastDay = new Date(year, month, 0).getDate();
  for (let day = 1; day <= lastDay; day++) {
    const date = prefix + '-' + String(day).padStart(2, '0');
    const visible = all.filter(task => taskVisibleOnDate_(task, date));
    if (visible.length) result[date] = countsOnly ? visible.length : visible;
  }
  return result;
}

function getMonthTasks(year, month) {
  const y = Number(year), m = Number(month);
  if (!Number.isInteger(y) || !Number.isInteger(m) || m < 1 || m > 12) throw new Error('年月が正しくありません。');
  return buildMonthTasks_(getAllTasks_(), y, m);
}

function getDailyStartTime_(date) {
  const value = PropertiesService.getScriptProperties().getProperty(START_TIME_PROPERTY_PREFIX + date);
  return value && /^([01]\d|2[0-3]):[0-5]\d$/.test(value) ? value : '09:00';
}

function saveDailyStartTime(date, time) {
  const selected = validateDate_(date, '対象日');
  const normalized = cleanText_(time || '09:00', 5);
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(normalized)) throw new Error('開始時刻を正しく入力してください。');
  PropertiesService.getScriptProperties().setProperty(START_TIME_PROPERTY_PREFIX + selected, normalized);
  return getDashboardData(selected);
}

function getDashboardData(date, compactCalendar) {
  const selected = date ? validateDate_(date, '対象日') : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return buildDashboardData_(selected, ensureSheet_(), compactCalendar === true);
}

function buildDashboardData_(selected, sheet, compactCalendar) {
  const diagnostics = { invalidRows: 0 };
  const all = getAllTasks_(sheet, diagnostics);
  const year = Number(selected.slice(0, 4)), month = Number(selected.slice(5, 7));
  const monthTasks = buildMonthTasks_(all, year, month, compactCalendar);
  const tasks = all.filter(task => taskVisibleOnDate_(task, selected));
  return { selectedDate: selected, tasks: tasks, monthTasks: monthTasks, startTime: getDailyStartTime_(selected), safeMutations: true, invalidRows: diagnostics.invalidRows };
}

function saveTask(input, selectedDate) {
  if (!input || typeof input !== 'object') throw new Error('タスクの入力内容がありません。');
  const selected = selectedDate ? validateDate_(selectedDate, '対象日') : '';
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ensureSheet_(), columns = getColumnMap_(sheet), now = new Date();
    const operation = input.operation;
    const providedId = cleanText_(input.id, 200);
    if (operation && !providedId) throw new Error('タスクIDがありません。');
    const id = providedId || Utilities.getUuid();
    let targetRow = 0, createdAt = now;
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, columns.ID, lastRow - 1, 1).getDisplayValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0] === id) {
          targetRow = i + 2;
          createdAt = sheet.getRange(targetRow, columns['作成日時']).getValue() || now;
          break;
        }
      }
    }
    if (operation === 'create' && targetRow) return { id };
    if ((operation === 'edit' || operation === 'toggle') && !targetRow) {
      throw new Error('タスクが見つかりません。すでに削除されている可能性があります。');
    }
    if (operation === 'toggle') {
      sheet.getRange(targetRow, columns['完了']).setValue(input.completed === true);
      sheet.getRange(targetRow, columns['更新日時']).setValue(now);
      SpreadsheetApp.flush();
      return { id };
    }
    if (operation === 'edit') {
      const current = sheet.getRange(targetRow, columns['更新日時']).getValue();
      const version = current instanceof Date ? current.toISOString() : String(current || '');
      if (version !== input.expectedUpdatedAt) throw new Error('別の画面で更新されました。最新のタスクを開き直してください。');
    }
    const date = validateDate_(input.date, '開始日');
    const title = cleanText_(input.title, 200);
    if (!title) throw new Error('タスク名を入力してください。');
    const dueDate = input.dueDate ? validateDate_(input.dueDate, '期限日') : date;
    if (dueDate < date) throw new Error('期限日は開始日以降の日付にしてください。');
    const completed = input.completed === true;
    const memo = cleanText_(input.memo, 10000);
    const duration = normalizeDuration_(input.duration);
    const importance = normalizeImportance_(input.importance);
    const longTask = input.longTask === true;
    if (!targetRow) targetRow = sheet.getLastRow() + 1;
    const values = {
      'ID': id, '日付': date, 'タスク名': sheetText_(title), '期限': dueDate, '完了': completed,
      'メモ': sheetText_(memo), '作成日時': createdAt, '更新日時': now,
      '推定所要時間（分）': duration, '重要度': importance, 'ロングタスク': longTask
    };
    const firstColumn = Math.min.apply(null, HEADERS.map(h => columns[h]));
    const lastColumn = Math.max.apply(null, HEADERS.map(h => columns[h]));
    const rowValues = sheet.getRange(targetRow, firstColumn, 1, lastColumn - firstColumn + 1).getValues()[0];
    Object.keys(values).forEach(h => { rowValues[columns[h] - firstColumn] = values[h]; });
    sheet.getRange(targetRow, firstColumn, 1, rowValues.length).setValues([rowValues]);
    SpreadsheetApp.flush();
    return operation ? { id } : buildDashboardData_(selected || date, sheet);
  } finally {
    lock.releaseLock();
  }
}

function sheetText_(value) {
  return /^[=']/.test(value) ? "'" + value : value;
}

function deleteTask(input, selectedDate) {
  const ackOnly = input && typeof input === 'object';
  const taskId = cleanText_(ackOnly ? input.id : input, 200);
  const date = selectedDate ? validateDate_(selectedDate, '対象日') : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (!taskId) throw new Error('削除するタスクIDがありません。');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ensureSheet_(), columns = getColumnMap_(sheet), lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, columns.ID, lastRow - 1, 1).getDisplayValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (ids[i][0] === taskId) { sheet.deleteRow(i + 2); break; }
      }
    }
    SpreadsheetApp.flush();
    return ackOnly ? { id: taskId } : buildDashboardData_(date, sheet);
  } finally {
    lock.releaseLock();
  }
}

function setup() {
  const id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_PROPERTY);
  if (!id) throw new Error('既存スプレッドシートIDが未設定です。setup()による新規作成は行いません。');
  const sheet = ensureSheet_();
  return { spreadsheetUrl: sheet.getParent().getUrl(), sheetName: sheet.getName() };
}

