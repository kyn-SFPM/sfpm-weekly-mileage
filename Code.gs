/**
 * SFPM Weekly Mileage -- Apps Script backend
 *
 * SETUP:
 * 1. Go to script.google.com > New Project, delete the placeholder code,
 *    and paste this whole file in.
 * 2. Update the four sheet IDs below if you ever move/recreate the sheets.
 * 3. Deploy > New deployment > type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the Web app URL it gives you -- that is what goes into the
 *    APPS_SCRIPT_URL constant at the top of the HTML front end.
 * 5. Re-deploy (Manage deployments > Edit > New version) any time you
 *    change this script.
 */

const ROSTER_SHEET_ID = '1KvQ4rcQHtoeDbVnNEXQBV7G2eqzFMdq7lFewjcWmsjY';
const INTERVALS_SHEET_ID = '1sKpBWxmQcUujuZJrvS9T10NEb-0XrhQTaVaWone75Kw';
const BASELINE_SHEET_ID = '1JpffVRqJpO0StT-7QhBpoYLPvKYmEN0BOC9uA9m3Bjw';
const LOG_SHEET_ID = '1MGgd9U_ciUi5HpuVti7uOqOCGioNefsvH1eNuEdl6YI';

const DUE_SOON_BUFFER = 500; // flag as "due soon" within this many miles of the interval

function doGet(e) {
  const roster = getRoster_();
  const intervals = getIntervals_();
  return jsonOut_({ roster: roster, intervals: intervals });
}

function doPost(e) {
  try {
    const params = e.parameter;
    const employee = (params.employee || '').trim();
    const unit = (params.unit || '').trim();
    const mileage = Number(params.mileage);
    const weekEnding = (params.weekEnding || '').trim();
    const notes = (params.notes || '').trim();

    if (!employee || !unit || !mileage) {
      return jsonOut_({ ok: false, error: 'Missing employee, unit, or mileage.' });
    }

    // Log the submission
    const logSheet = SpreadsheetApp.openById(LOG_SHEET_ID).getSheets()[0];
    logSheet.appendRow([new Date(), employee, unit, mileage, weekEnding, notes]);

    // Compute service status against baseline + intervals
    const intervals = getIntervals_();
    const baseline = getBaselineForUnit_(unit);
    const results = [];

    intervals.forEach(function (item) {
      const lastAt = baseline[item.name];
      if (lastAt === null || lastAt === undefined || lastAt === '') {
        results.push({ item: item.name, status: 'unknown', message: item.name + ': no baseline on file yet' });
        return;
      }
      const since = mileage - Number(lastAt);
      const remaining = item.miles - since;
      if (remaining <= 0) {
        results.push({ item: item.name, status: 'overdue', remaining: remaining, message: item.name + ' -- overdue by ' + Math.abs(remaining).toLocaleString() + ' mi' });
      } else if (remaining <= DUE_SOON_BUFFER) {
        results.push({ item: item.name, status: 'due_soon', remaining: remaining, message: item.name + ' -- due in ' + remaining.toLocaleString() + ' mi' });
      } else {
        results.push({ item: item.name, status: 'ok', remaining: remaining, message: item.name + ' -- ok (' + remaining.toLocaleString() + ' mi left)' });
      }
    });

    return jsonOut_({ ok: true, unit: unit, mileage: mileage, results: results });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function getRoster_() {
  const sheet = SpreadsheetApp.openById(ROSTER_SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  return rows
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return { unit: r[0], assigned: r[1], entity: r[2], year: r[3], make: r[4], model: r[5] };
    });
}

function getIntervals_() {
  const sheet = SpreadsheetApp.openById(INTERVALS_SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);
  return rows
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { name: r[0], miles: Number(r[1]) }; });
}

function getBaselineForUnit_(unit) {
  const sheet = SpreadsheetApp.openById(BASELINE_SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const row = data.slice(1).find(function (r) { return r[0] === unit; });
  const out = {};
  if (!row) return out;
  headers.forEach(function (h, i) {
    if (i === 0) return;
    out[h] = row[i];
  });
  return out;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
