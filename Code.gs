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

const ROSTER_SHEET_ID = '1DNDODv5taegzsA7fzowtiYcHrLWzrSXemklV3I_zlac';
const INTERVALS_SHEET_ID = '1sKpBWxmQcUujuZJrvS9T10NEb-0XrhQTaVaWone75Kw';
const BASELINE_SHEET_ID = '1_MVvcxxn2WIPRMwFcFl6_BguYUM9b25ioX_E8K8zW2U';
const LOG_SHEET_ID = '1MGgd9U_ciUi5HpuVti7uOqOCGioNefsvH1eNuEdl6YI';
const RECEIPTS_FOLDER_ID = '1D3ABMoGUb3f6IknmDJapfZ5XvDxxKw-l';
const RECEIPTS_LOG_SHEET_ID = '1Vf3hi4ao7EkcfZ201-Dvw4cHqBB4QcW3Yq2zevBNG0I';

// Earliest/latest known odometer readings per unit, pulled from the historical
// odometer export. Used to estimate mileage on a given date when backfilling
// a known service date without knowing the exact mileage at the time.
// Unit 15 excluded -- its history shows mileage dropping over time, which
// means the vehicle assigned to that slot was swapped, so linear estimation
// would be wrong.
const MILEAGE_ANCHORS = {
  'Unit 1': [['2019-05-17', 119681], ['2025-12-12', 216671]],
  'Unit 2': [['2019-05-17', 62406], ['2025-12-12', 72577]],
  'Unit 3': [['2019-05-17', 63217], ['2025-12-12', 65567]],
  'Unit 4': [['2019-05-17', 113647], ['2025-12-12', 272962]],
  'Unit 5': [['2019-05-17', 52062], ['2025-12-12', 114055]],
  'Unit 6': [['2019-05-17', 233262], ['2025-12-12', 298834]],
  'Unit 8': [['2019-05-17', 86264], ['2025-12-12', 230586]],
  'Unit 9': [['2019-05-17', 91000], ['2025-12-12', 147591]],
  'Unit 10': [['2019-05-17', 211285], ['2025-12-12', 296158]],
  'Unit 11': [['2019-05-17', 120519], ['2025-12-12', 182048]],
  'Unit 12': [['2019-05-17', 87306], ['2025-12-12', 179263]],
  'Unit 13': [['2019-05-17', 240556], ['2025-12-12', 273208]],
  'Unit 14': [['2019-05-17', 106813], ['2025-12-12', 147431]],
  'Unit 16': [['2019-05-17', 132270], ['2025-12-12', 191806]],
  'Unit 19': [['2019-08-09', 124728], ['2025-12-12', 165760]],
  'Unit 20': [['2019-08-09', 131869], ['2025-12-12', 257954]],
  'Unit 21': [['2020-07-31', 66957], ['2025-12-12', 129873]],
  'Unit 22': [['2020-07-31', 85327], ['2025-12-12', 149478]],
  'Unit 23': [['2022-03-21', 1269], ['2025-12-12', 86702]],
  'Unit 24': [['2022-03-21', 2784], ['2025-12-12', 124265]],
  'Unit 25': [['2022-12-05', 14], ['2025-12-12', 59773]],
  'Unit 26': [['2023-03-10', 1989], ['2025-12-12', 47978]],
  'Unit 27': [['2023-03-17', 960], ['2025-12-12', 61835]],
  'Unit 28': [['2024-05-13', 1048], ['2025-12-12', 43003]],
  'Unit 29': [['2024-06-28', 17041], ['2025-12-12', 56165]],
  'Unit 30': [['2025-05-16', 907], ['2025-12-12', 5625]],
  'Unit 31': [['2025-05-16', 942], ['2025-12-12', 7644]]
};

function estimateMileageAtDate_(unit, dateStr) {
  const anchors = MILEAGE_ANCHORS[unit];
  if (!anchors || !dateStr) return null;
  const target = new Date(dateStr + 'T00:00:00').getTime();
  const d0 = new Date(anchors[0][0] + 'T00:00:00').getTime();
  const m0 = anchors[0][1];
  const d1 = new Date(anchors[1][0] + 'T00:00:00').getTime();
  const m1 = anchors[1][1];
  if (d1 === d0) return m0;
  const frac = (target - d0) / (d1 - d0);
  const estimate = m0 + frac * (m1 - m0);
  return Math.max(0, Math.round(estimate));
}

const DUE_SOON_BUFFER = 500; // flag as "due soon" within this many miles of the interval

function doGet(e) {
  const mode = (e && e.parameter && e.parameter.mode) || 'form';
  if (mode === 'admin') {
    return jsonOut_(buildFleetStatus_());
  }
  const roster = getRoster_();
  const intervals = getIntervals_();
  return jsonOut_({ roster: roster, intervals: intervals });
}

function buildFleetStatus_() {
  const roster = getRoster_();
  const intervals = getIntervals_();
  const latestMileage = getLatestMileageByUnit_();
  const baselineSheet = SpreadsheetApp.openById(BASELINE_SHEET_ID).getSheets()[0];
  const baselineData = baselineSheet.getDataRange().getValues();
  const baselineHeaders = baselineData[0];

  const fleet = roster.map(function (r) {
    const unit = r.unit;
    const latest = latestMileage[unit];
    const row = baselineData.slice(1).find(function (br) { return br[0] === unit; });

    const items = intervals.map(function (item) {
      if (!latest) {
        return { item: item.name, status: 'no_data' };
      }
      const colIndex = baselineHeaders.indexOf(item.name);
      const lastAt = row ? row[colIndex] : '';
      if (lastAt === '' || lastAt === undefined || lastAt === null) {
        return { item: item.name, status: 'no_baseline' };
      }
      const since = latest.mileage - Number(lastAt);
      const remaining = item.miles - since;
      if (remaining <= 0) return { item: item.name, status: 'overdue', remaining: remaining };
      if (remaining <= DUE_SOON_BUFFER) return { item: item.name, status: 'due_soon', remaining: remaining };
      return { item: item.name, status: 'ok', remaining: remaining };
    });

    const overdueCount = items.filter(function (i) { return i.status === 'overdue'; }).length;
    const dueSoonCount = items.filter(function (i) { return i.status === 'due_soon'; }).length;

    const docItems = [
      { item: 'Tags/Registration', status: dateStatus_(r.tagsExpire), dateValue: r.tagsExpire },
      { item: 'Insurance', status: dateStatus_(r.insuranceExpires), dateValue: r.insuranceExpires },
      { item: 'Inspection', status: dateStatus_(r.inspectionDate), dateValue: r.inspectionDate }
    ];
    const docOverdue = docItems.filter(function (i) { return i.status === 'overdue'; }).length;
    const docDueSoon = docItems.filter(function (i) { return i.status === 'due_soon'; }).length;

    let worst = 'ok';
    if (!latest) worst = 'no_data';
    if (overdueCount > 0 || docOverdue > 0) worst = 'overdue';
    else if (dueSoonCount > 0 || docDueSoon > 0) worst = 'due_soon';

    const daysSinceReport = latest ? Math.floor((new Date() - new Date(latest.timestamp)) / 86400000) : null;
    const reportStale = daysSinceReport === null || daysSinceReport > 7;

    return {
      unit: unit,
      assigned: r.assigned,
      entity: r.entity,
      active: r.active,
      yearMakeModel: [r.year, r.make, r.model].filter(Boolean).join(' '),
      vin: r.vin,
      plate: r.plate,
      latestMileage: latest ? latest.mileage : null,
      latestDate: latest ? latest.timestamp : null,
      daysSinceReport: daysSinceReport,
      reportStale: reportStale,
      worst: worst,
      overdueCount: overdueCount + docOverdue,
      dueSoonCount: dueSoonCount + docDueSoon,
      items: items,
      docItems: docItems
    };
  });

  const order = { overdue: 0, due_soon: 1, no_data: 2, no_baseline: 2, ok: 3 };
  fleet.sort(function (a, b) { return order[a.worst] - order[b.worst]; });

  return { fleet: fleet, generatedAt: new Date() };
}

function getLatestMileageByUnit_() {
  const sheet = SpreadsheetApp.openById(LOG_SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // Timestamp, Employee, Unit, Mileage, WeekEnding, Notes
  const latest = {};
  rows.forEach(function (r) {
    const unit = r[2];
    const mileage = Number(r[3]);
    const timestamp = r[0];
    if (!unit || !mileage) return;
    if (!latest[unit] || timestamp > latest[unit].timestamp) {
      latest[unit] = { mileage: mileage, timestamp: timestamp };
    }
  });
  return latest;
}

function doPost(e) {
  const formType = (e.parameter.formType || 'mileage');
  if (formType === 'receipt') return handleReceipt_(e);
  if (formType === 'setActive') return handleSetActive_(e);
  if (formType === 'backfillService') return handleBackfillService_(e);
  return handleMileage_(e);
}

function handleSetActive_(e) {
  try {
    const unit = (e.parameter.unit || '').trim();
    const active = e.parameter.active === 'true';
    const sheet = SpreadsheetApp.openById(ROSTER_SHEET_ID).getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(function (r) { return r[0] === unit; });
    if (rowIndex === -1) return jsonOut_({ ok: false, error: 'Unit not found.' });
    sheet.getRange(rowIndex + 1, 12).setValue(active ? 'Yes' : 'No'); // column L = Active
    return jsonOut_({ ok: true, unit: unit, active: active });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function handleBackfillService_(e) {
  try {
    const unit = (e.parameter.unit || '').trim();
    const itemName = (e.parameter.serviceItem || '').trim();
    const dateStr = (e.parameter.date || '').trim(); // date the service was actually done, for your own records
    let mileage = Number(e.parameter.mileage);
    let estimated = false;

    if (!mileage && dateStr) {
      const est = estimateMileageAtDate_(unit, dateStr);
      if (est !== null) {
        mileage = est;
        estimated = true;
      }
    }

    if (!unit || !itemName || !mileage) {
      return jsonOut_({ ok: false, error: 'Missing unit, service item, or mileage -- and no date given to estimate from.' });
    }
    setBaselineCell_(unit, itemName, mileage);

    // Also log it to the receipts log as a record, so there's a dated history entry.
    const sheet = SpreadsheetApp.openById(RECEIPTS_LOG_SHEET_ID).getSheets()[0];
    sheet.appendRow([dateStr ? new Date(dateStr) : new Date(), 'Backfilled (admin)', unit, itemName, '', '', 'Historical baseline entered manually' + (dateStr ? ' -- service date: ' + dateStr : '') + (estimated ? ' -- mileage estimated from date' : ''), '']);

    return jsonOut_({ ok: true, unit: unit, item: itemName, mileage: mileage, estimated: estimated });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function handleReceipt_(e) {
  try {
    const params = e.parameter;
    const employee = (params.employee || '').trim();
    const unit = (params.unit || '').trim();
    const serviceType = (params.serviceType || '').trim();
    const vendor = (params.vendor || '').trim();
    const cost = (params.cost || '').trim();
    const notes = (params.notes || '').trim();
    const photoBase64 = params.photo;
    const photoMime = params.photoMime || 'image/jpeg';

    if (!employee || !unit) {
      return jsonOut_({ ok: false, error: 'Missing employee or unit.' });
    }

    let fileUrl = '';
    if (photoBase64) {
      const bytes = Utilities.base64Decode(photoBase64);
      const blob = Utilities.newBlob(bytes, photoMime, unit + '_' + new Date().getTime() + '.jpg');
      const folder = DriveApp.getFolderById(RECEIPTS_FOLDER_ID);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }

    const sheet = SpreadsheetApp.openById(RECEIPTS_LOG_SHEET_ID).getSheets()[0];
    sheet.appendRow([new Date(), employee, unit, serviceType, vendor, cost, notes, fileUrl]);

    return jsonOut_({ ok: true, unit: unit });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function handleMileage_(e) {
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
    const newlySet = [];

    intervals.forEach(function (item) {
      const lastAt = baseline[item.name];
      if (lastAt === null || lastAt === undefined || lastAt === '') {
        // No baseline yet -- this reading becomes the starting point.
        setBaselineCell_(unit, item.name, mileage);
        newlySet.push(item.name);
        results.push({ item: item.name, status: 'ok', remaining: item.miles, message: item.name + ' -- baseline set today, next check in ' + item.miles.toLocaleString() + ' mi' });
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

    return jsonOut_({ ok: true, unit: unit, mileage: mileage, results: results, baselinesJustSet: newlySet });
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
      return {
        unit: r[0], assigned: r[1], entity: r[2], year: r[3], make: r[4], model: r[5],
        vin: r[6], plate: r[7], tagsExpire: fmtDate_(r[8]), insuranceExpires: fmtDate_(r[9]), inspectionDate: fmtDate_(r[10]),
        active: (r[11] === undefined || r[11] === '' || String(r[11]).toLowerCase() === 'yes')
      };
    });
}

function fmtDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function daysUntil_(dateStr) {
  if (!dateStr) return null;
  const target = new Date(dateStr + 'T00:00:00');
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return Math.round((target - today) / 86400000);
}

function dateStatus_(dateStr) {
  const d = daysUntil_(dateStr);
  if (d === null) return 'no_data';
  if (d < 0) return 'overdue';
  if (d <= 30) return 'due_soon';
  return 'ok';
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

function setBaselineCell_(unit, itemName, mileage) {
  const sheet = SpreadsheetApp.openById(BASELINE_SHEET_ID).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = headers.indexOf(itemName);
  if (colIndex === -1) return;

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === unit) { rowIndex = i; break; }
  }
  if (rowIndex === -1) {
    // Unit isn't on the baseline sheet yet -- add it.
    const newRow = new Array(headers.length).fill('');
    newRow[0] = unit;
    newRow[colIndex] = mileage;
    sheet.appendRow(newRow);
    return;
  }
  sheet.getRange(rowIndex + 1, colIndex + 1).setValue(mileage);
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
