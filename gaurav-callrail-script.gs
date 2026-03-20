// ============================================================
// GAURAV DASHBOARD — CallRail Data Puller
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
// Superpath MCC accounts:
//   1 Point Electric                    | 428-467-1850
//   Auburn Leach K9 Solutions           | 410-807-3491
//   Bob Larson Plumbing                 | 904-107-5919
//   Broadco                             | 988-418-1651
//   Byers Electric Service Team         | 656-136-9007
//   C.W.Fischer Electric                | 413-360-8484
//   Country Meadows                     | 595-548-0486
//   Discount Office                     | 164-060-4631
//   EMC Electric, Inc                   | 645-934-8768
//   GHome                               | 792-974-6795
//   Hallmark Tree Service               | 759-738-9590
//   Hassle Free Lawns                   | 385-306-5831
//   Industrial Resin Recycling          | 278-232-9441
//   Kraus Tile & Bath                   | 802-179-9249
//   Mansea Metal                        | 820-249-0394
//   Schonsheck                          | 306-437-9996
//   The Pampered House Maid Services    | 230-396-6254
//   Underwood Plumbing                  | 843-364-4227
//   Maggie's Wigs 4 Kids of Michigan    | 740-352-7070
//
// Bell Media MCC accounts:
//   Alabama Power                       | 592-109-8560
//   Bromberg & Co Inc.                  | 204-572-9484
//   Hawker Powersource                  | 919-746-5666
//   Infinity Med-I-Spa                  | 161-268-8539
//   Koch Dentistry                      | 551-410-7788
//   Shelby Dental                       | 992-412-9662
//
// Note: CallRail company names above are starting points — verify against
// exact names in CallRail and adjust as needed.
//
// Unresolved (not found in either MCC — may be Meta-only or pending setup):
//   Arca Aesthetics, Ellery Milan Beauty, MN Express, Pink Pony Pub,
//   Southern Carriers Inc., Studio 4955 Aesthetics, Trinity Medical Solutions

// ─── MAIN ENTRY ───────────────────────────────────────────────
function pullCallRailData() {
  var ss       = SpreadsheetApp.openById(CR_CONFIG.SHEET_ID);
  var crSheet  = getOrCreate(ss, 'callrail_data', CR_HEADERS.callrail);
  var mapSheet = getOrCreate(ss, 'account_map',   CR_HEADERS.accountMap);

  var today     = fmtDate(0);
  var startDate = fmtDate(CR_CONFIG.LOOKBACK_DAYS);

  var mapping = buildMapping(mapSheet);

  var allCompanies = fetchCompanies();
  Logger.log('CallRail total companies: ' + allCompanies.length);

  var companies = allCompanies.filter(function(c) {
    return !!mapping[c.name.toLowerCase().trim()];
  });
  Logger.log('Mapped to Gaurav accounts: ' + companies.length);

  if (!companies.length) {
    Logger.log('No mapped companies found — fill in the account_map tab first.');
    return;
  }

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
    '&fields=answered,duration,source' +
    '&per_page='    + CR_CONFIG.PER_PAGE;

  try {
    var res   = UrlFetchApp.fetch(url, authHeaders());
    var data  = JSON.parse(res.getContentText());
    var calls = (data.calls || []).filter(function(c) {
      return c.source && c.source.toLowerCase().indexOf('google') !== -1;
    });

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
    '&fields=source' +
    '&per_page='   + CR_CONFIG.PER_PAGE;

  try {
    var res  = UrlFetchApp.fetch(url, authHeaders());
    var data = JSON.parse(res.getContentText());
    return (data.form_submissions || []).filter(function(f) {
      return f.source && f.source.toLowerCase().indexOf('google') !== -1;
    }).length;
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
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'pullCallRailData') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('pullCallRailData')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  Logger.log('Daily trigger set for 7 AM.');
}
