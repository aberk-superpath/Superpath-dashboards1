// ============================================================
// GAVIN DASHBOARD — CallRail Data Puller
// Install as a standalone Google Apps Script
// (NOT bound to the sheet — standalone script with SHEET_ID)
//
// Setup:
//   1. Go to script.google.com → New project
//   2. Paste this entire script
//   3. Run setupTrigger() once to create the daily trigger
//   4. Authorize when prompted
//
// Schedule: Daily at 7:00 AM (runs after the Google Ads scripts)
// ============================================================

var CR_CONFIG = {
  API_KEY:    'YOUR_CALLRAIL_API_KEY',
  ACCOUNT_ID: 'YOUR_CALLRAIL_ACCOUNT_ID',
  SHEET_ID:   'YOUR_GOOGLE_SHEET_ID',
  TIMEZONE:   'America/New_York',
  LOOKBACK_DAYS: 30,
  PER_PAGE:   250
};

var CR_HEADERS = {
  callrail:   ['date','company_id','company_name','ads_account_id',
               'total_calls','answered_calls','missed_calls','first_time_callers',
               'form_submissions','avg_call_duration_sec'],
  accountMap: ['callrail_company_name','ads_account_id','notes']
};

// ─── account_map tab — populate in the Sheet with these rows ──
// (callrail_company_name must match exactly what CallRail shows)
//
// Superpath MCC accounts (exact CallRail company names):
//   A1 Check Cashing                              | 939-906-9856
//   Birmingham Mosquito Control                   | 846-879-6505
//   Connell's Pest Control                        | 645-214-4301
//   DMH Electric                                  | 936-735-8878
//   Elder Law Lexington                           | 637-717-3911  (confirm = Elder Law of Alabama)
//   Fortress Roofing                              | 908-629-9597
//   Heavyweight Waste                             | 177-566-3931
//   Houk Air Conditioning - San Antonio           | 668-998-1948
//   Houk Air Conditioning - Houston               | 668-998-1948
//   Houk Air Conditioning - HLM                   | 668-998-1948
//   Houk Air Conditioning - DFW                   | 668-998-1948
//   Houk Air Conditioning - Austin                | 668-998-1948
//   JBathe Electric                               | 962-984-8921
//   Let It Glow Electric                          | 831-534-5637
//   Mac Electric and Lighting                     | 348-959-6662
//   Paff Electric and Solar                       | 655-187-1568
//   Prantl Heating & Cooling                      | 419-569-2711
//   Rhema Electric                                | 707-844-0497
//   Rooter MD Plumbing                            | 573-217-9782
//   Rooter Solutions - San Jose                   | 939-434-2239  (shared with Jared)
//   Rooter Solutions - Santa Barbara              | 400-550-7148  (shared with Jared)
//   Rooter Solutions - San Diego                  | 370-453-9854  (shared with Jared)
//   Semper Fi Heating and Cooling                 | 966-140-2141
//   Smash Franchise Partners (Smash My Trash)     | 853-232-1784  (shared with Jared)
//   U.S. Smoke & Fire                             | 528-340-7322
//   White Glove Transportation Services           | 264-806-2658
//   Woodworth Heat Treating                       | 324-478-3262  (confirm = Woodworth Electric)
//
// Bell Media MCC accounts (exact CallRail company names):
//   Arobotech                                     | 701-646-5480
//   Nonconnah Corporate Center                    | 116-427-8686
//   Odyssey Electronics Inc.                      | 233-500-2045
//   Guin Service                                  | 395-434-5123
//
// NOT in CallRail (omitted): Hidden Oak Labradors, Southland Goldens, Birmingham Airport Authority
// Note: One Lincoln Park (869-354-5249) is Gaurav's account — do NOT include here.

// ─── MAIN ENTRY ───────────────────────────────────────────────
function pullCallRailData() {
  var ss       = SpreadsheetApp.openById(CR_CONFIG.SHEET_ID);
  var crSheet  = getOrCreate(ss, 'callrail_data', CR_HEADERS.callrail);
  var mapSheet = getOrCreate(ss, 'account_map',   CR_HEADERS.accountMap);

  var today     = fmtDate(0);
  var startDate = fmtDate(CR_CONFIG.LOOKBACK_DAYS);

  // Load company → Google Ads account ID mapping from account_map tab
  var mapping = buildMapping(mapSheet);

  // Fetch all CallRail companies, then filter to only those in account_map
  var allCompanies = fetchCompanies();
  Logger.log('CallRail total companies: ' + allCompanies.length);

  var companies = allCompanies.filter(function(c) {
    return !!mapping[c.name.toLowerCase().trim()];
  });
  Logger.log('Mapped to Gavin accounts: ' + companies.length);

  if (!companies.length) {
    Logger.log('No mapped companies found — fill in the account_map tab first.');
    return;
  }

  // Remove today's existing rows (safe re-run)
  clearTodayRows(crSheet, today);

  var newRows = [];

  companies.forEach(function(company) {
    var cId   = company.id;
    var cName = company.name;
    var adsId = mapping[cName.toLowerCase().trim()] || '';

    var callSummary = fetchCallSummary(cId, startDate, today);
    var formCount   = fetchFormCount(cId, startDate, today);

    newRows.push([
      today,
      cId,
      cName,
      adsId,
      callSummary.total,
      callSummary.answered,
      callSummary.missed,
      callSummary.firstTime,
      formCount,
      callSummary.avgDuration
    ]);

    Logger.log(cName + ': ' + callSummary.total + ' calls, ' + formCount + ' forms');
  });

  if (newRows.length > 0) {
    crSheet.getRange(crSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
           .setValues(newRows);
  }

  Logger.log('Done. Wrote ' + newRows.length + ' company rows for ' + today);
}

// ─── CALLRAIL API CALLS ───────────────────────────────────────
function fetchCompanies() {
  var allCompanies = [];
  var page = 1;
  var perPage = 100;

  while (true) {
    var url = 'https://api.callrail.com/v3/a/' + CR_CONFIG.ACCOUNT_ID +
              '/companies.json?fields=name,id&per_page=' + perPage + '&page=' + page;
    try {
      var res   = UrlFetchApp.fetch(url, authHeaders());
      var data  = JSON.parse(res.getContentText());
      var batch = data.companies || [];
      allCompanies = allCompanies.concat(batch);
      if (batch.length < perPage) break;
      page++;
    } catch(e) {
      Logger.log('fetchCompanies ERR (page ' + page + '): ' + e.message);
      break;
    }
  }

  return allCompanies;
}

function fetchCallSummary(companyId, startDate, endDate) {
  var empty = { total: 0, answered: 0, missed: 0, firstTime: 0, avgDuration: 0 };
  var url =
    'https://api.callrail.com/v3/a/' + CR_CONFIG.ACCOUNT_ID + '/calls.json' +
    '?company_id='  + companyId +
    '&start_date='  + startDate +
    '&end_date='    + endDate +
    '&fields=answered,duration' +
    '&per_page='    + CR_CONFIG.PER_PAGE;

  try {
    var res   = UrlFetchApp.fetch(url, authHeaders());
    var data  = JSON.parse(res.getContentText());
    var calls = data.calls || [];

    var total       = calls.length;
    var answered    = calls.filter(function(c) { return c.answered; }).length;
    var firstTime   = calls.filter(function(c) { return c.first_time_caller; }).length;
    var totalDur    = calls.reduce(function(s, c) { return s + (parseInt(c.duration) || 0); }, 0);
    var avgDuration = total > 0 ? Math.round(totalDur / total) : 0;

    return {
      total:       total,
      answered:    answered,
      missed:      total - answered,
      firstTime:   firstTime,
      avgDuration: avgDuration
    };
  } catch(e) {
    Logger.log('fetchCallSummary ERR [' + companyId + ']: ' + e.message);
    return empty;
  }
}

function fetchFormCount(companyId, startDate, endDate) {
  var url =
    'https://api.callrail.com/v3/a/' + CR_CONFIG.ACCOUNT_ID + '/form_submissions.json' +
    '?company_id=' + companyId +
    '&start_date=' + startDate +
    '&end_date='   + endDate +
    '&per_page='   + CR_CONFIG.PER_PAGE;

  try {
    var res  = UrlFetchApp.fetch(url, authHeaders());
    var data = JSON.parse(res.getContentText());
    return (data.form_submissions || []).length;
  } catch(e) {
    Logger.log('fetchFormCount ERR [' + companyId + ']: ' + e.message);
    return 0;
  }
}

// ─── MAPPING ──────────────────────────────────────────────────
function buildMapping(mapSheet) {
  var map = {};
  if (mapSheet.getLastRow() <= 1) return map;
  var rows = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 2).getValues();
  rows.forEach(function(row) {
    if (row[0] && row[1]) {
      map[String(row[0]).toLowerCase().trim()] = String(row[1]).trim();
    }
  });
  return map;
}

// ─── SHEET UTILITIES ──────────────────────────────────────────
function getOrCreate(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold')
         .setBackground('#1a2a5e').setFontColor('#ffffff');
  }
  return sheet;
}

function clearTodayRows(sheet, today) {
  if (sheet.getLastRow() <= 1) return;
  var data = sheet.getDataRange().getValues();
  var keep = data.slice(1).filter(function(row) {
    if (!row[0]) return false; // drop phantom empty rows
    var cellStr = (row[0] instanceof Date)
      ? Utilities.formatDate(row[0], CR_CONFIG.TIMEZONE, 'yyyy-MM-dd')
      : String(row[0]);
    return cellStr !== today;
  });
  sheet.getRange(2, 1, data.length - 1, data[0].length).clearContent();
  if (keep.length > 0) {
    sheet.getRange(2, 1, keep.length, keep[0].length).setValues(keep);
  }
}

// ─── UTILITIES ────────────────────────────────────────────────
function authHeaders() {
  return { headers: { 'Authorization': 'Token token=' + CR_CONFIG.API_KEY } };
}

function fmtDate(daysAgo) {
  var d = new Date();
  d.setDate(d.getDate() - daysAgo);
  return Utilities.formatDate(d, CR_CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

// ─── TRIGGER SETUP (run once) ─────────────────────────────────
// Run this function manually one time to install the daily 7 AM trigger.
function setupTrigger() {
  // Remove any existing triggers for pullCallRailData
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'pullCallRailData') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Create new daily trigger at 7 AM ET
  ScriptApp.newTrigger('pullCallRailData')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  Logger.log('Daily trigger set for 7 AM.');
}
