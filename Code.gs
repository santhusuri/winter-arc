/**
 * WINTER ARC — Google Sheets Backend
 *
 * Deploy this file as a Web App:
 *   Execute as: Me
 *   Who has access: Anyone
 *
 * The HTML frontend can POST JSON to the deployed Web App URL.
 *
 * Expected spreadsheet tabs:
 *   Habits
 *   Tasks
 *   Goals
 *   Finance
 *   Transactions
 *   Transformation
 *
 * The script creates missing sheets automatically.
 */

const SPREADSHEET_ID = ''; // Optional: leave blank when this script is bound to your Sheet.

const SHEETS = {
  habits: 'Habits',
  tasks: 'Tasks',
  goals: 'Goals',
  finance: 'Finance',
  transactions: 'Transactions',
  transformation: 'Transformation'
};

function getSpreadsheet_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, headers) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }

  return sh;
}

function setupSheets() {
  sheet_(SHEETS.habits, [
    'id', 'date', 'habit', 'group', 'completed', 'timestamp'
  ]);

  sheet_(SHEETS.tasks, [
    'id', 'date', 'task', 'completed', 'timestamp'
  ]);

  sheet_(SHEETS.goals, [
    'id', 'goal', 'area', 'targetDate', 'progress', 'status', 'timestamp'
  ]);

  sheet_(SHEETS.finance, [
    'id', 'month', 'income', 'expenses', 'savings',
    'assets', 'liabilities', 'netWorth', 'fiTarget', 'fiProgress',
    'timestamp'
  ]);

  sheet_(SHEETS.transactions, [
    'id', 'date', 'type', 'amount', 'category', 'description', 'timestamp'
  ]);

  sheet_(SHEETS.transformation, [
    'id', 'date', 'photoUrl', 'weight', 'waist',
    'energy', 'note', 'timestamp'
  ]);

  return json_({ok: true, message: 'Winter Arc sheets are ready.'});
}

function doGet(e) {
  setupSheets();
  return json_({
    ok: true,
    app: 'Winter Arc',
    status: 'connected',
    message: 'Google Sheets backend is running.'
  });
}

function doPost(e) {
  try {
    setupSheets();

    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action || '';

    switch (action) {
      case 'saveHabit':
        return saveHabit_(body);

      case 'saveTask':
        return saveTask_(body);

      case 'saveGoal':
        return saveGoal_(body);

      case 'saveFinance':
        return saveFinance_(body);

      case 'saveTransaction':
        return saveTransaction_(body);

      case 'saveTransformation':
        return saveTransformation_(body);

      case 'getData':
        return getData_(body);

      default:
        return json_({
          ok: false,
          error: 'Unknown action: ' + action
        });
    }

  } catch (err) {
    return json_({
      ok: false,
      error: String(err)
    });
  }
}

function append_(sheetName, values) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(sheetName);

  if (!sh) throw new Error('Sheet not found: ' + sheetName);

  sh.appendRow(values);
}

function saveHabit_(d) {
  append_(SHEETS.habits, [
    d.id || Utilities.getUuid(),
    d.date || '',
    d.habit || '',
    d.group || '',
    d.completed === true ? true : false,
    new Date()
  ]);

  return json_({ok: true});
}

function saveTask_(d) {
  append_(SHEETS.tasks, [
    d.id || Utilities.getUuid(),
    d.date || '',
    d.task || '',
    d.completed === true ? true : false,
    new Date()
  ]);

  return json_({ok: true});
}

function saveGoal_(d) {
  append_(SHEETS.goals, [
    d.id || Utilities.getUuid(),
    d.goal || '',
    d.area || '',
    d.targetDate || '',
    Number(d.progress || 0),
    d.status || '',
    new Date()
  ]);

  return json_({ok: true});
}

function saveFinance_(d) {
  const income = Number(d.income || 0);
  const expenses = Number(d.expenses || 0);
  const assets = Number(d.assets || 0);
  const liabilities = Number(d.liabilities || 0);
  const savings = income - expenses;
  const netWorth = assets - liabilities;

  const fiTarget = Number(d.fiTarget || 0);
  const fiProgress = fiTarget > 0 ? (netWorth / fiTarget) * 100 : 0;

  append_(SHEETS.finance, [
    d.id || Utilities.getUuid(),
    d.month || '',
    income,
    expenses,
    savings,
    assets,
    liabilities,
    netWorth,
    fiTarget,
    fiProgress,
    new Date()
  ]);

  return json_({
    ok: true,
    calculated: {
      savings: savings,
      netWorth: netWorth,
      fiProgress: fiProgress
    }
  });
}

function saveTransaction_(d) {
  append_(SHEETS.transactions, [
    d.id || Utilities.getUuid(),
    d.date || '',
    d.type || '',
    Number(d.amount || 0),
    d.category || '',
    d.description || '',
    new Date()
  ]);

  return json_({ok: true});
}

function saveTransformation_(d) {
  /*
   * Recommended:
   * Upload the image to Google Drive from the frontend/backend and send
   * the resulting Drive URL as photoUrl.
   *
   * Do NOT store large base64 images directly in Sheets.
   */
  append_(SHEETS.transformation, [
    d.id || Utilities.getUuid(),
    d.date || '',
    d.photoUrl || '',
    d.weight || '',
    d.waist || '',
    d.energy || '',
    d.note || '',
    new Date()
  ]);

  return json_({ok: true});
}

function getData_(d) {
  const requested = d.sheet || '';

  if (requested && SHEETS[requested]) {
    return json_({
      ok: true,
      sheet: requested,
      data: readSheet_(SHEETS[requested])
    });
  }

  const result = {};

  Object.keys(SHEETS).forEach(function(key) {
    result[key] = readSheet_(SHEETS[key]);
  });

  return json_({
    ok: true,
    data: result
  });
}

function readSheet_(name) {
  const sh = getSpreadsheet_().getSheetByName(name);

  if (!sh || sh.getLastRow() < 2) return [];

  const values = sh.getDataRange().getValues();
  const headers = values.shift();

  return values.map(function(row) {
    const item = {};

    headers.forEach(function(header, i) {
      let value = row[i];

      if (value instanceof Date) {
        value = value.toISOString();
      }

      item[header] = value;
    });

    return item;
  });
}
