// ============================================================
// HVAC DASHBOARD — Google Ads MCC Script (merged)
// Replaces two separate Campaign + Keyword scripts with one.
// Install in: Superpath MCC (826-157-9112)
// Schedule: Daily at 6:00 AM — DELETE the old 6:00 and 6:30 scripts
// Sheet ID: YOUR_GOOGLE_SHEET_ID
// ============================================================

var CONFIG = {
  SHEET_ID:         'YOUR_GOOGLE_SHEET_ID',
  LABEL:            'HVAC - Dashboard',
  LOOKBACK_DAYS:    30,
  MAX_KEYWORDS:     500,
  TIMEZONE:         'America/New_York'
};

// Headers match the existing tabs exactly — do not reorder
var HEADERS = {
  campaign: ['Timestamp','Account','Customer ID','Campaign','Status','Campaign Type',
             'Cost','Clicks','Impressions','CTR','Avg. CPC','Conversions',
             'Conv. Rate','Cost / Conv.','Search Impr. Share',
             'IS Lost (Budget)','IS Lost (Rank)'],
  keyword:  ['Timestamp','Account','Customer ID','Keyword','Match Type','Ad Group',
             'Campaign','Campaign Status','Cost','Clicks','Impressions','CTR',
             'Avg. CPC','Conversions','Conv. Rate','Cost / Conv.','Quality Score']
};

// ─── MAIN ─────────────────────────────────────────────────────
function main() {
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var today     = fmtDate(0);
  var startDate = fmtDate(CONFIG.LOOKBACK_DAYS);

  var tabs = {
    campaign: getOrCreate(ss, 'Campaigns hvac', HEADERS.campaign),
    keyword:  getOrCreate(ss, 'Keywords hvac',  HEADERS.keyword)
  };

  // Collect labeled accounts — ONE iteration for both data types
  var accounts = [];
  var iter = MccApp.accounts().withCondition("LabelNames CONTAINS '" + CONFIG.LABEL + "'").get();
  while (iter.hasNext()) {
    var acct = iter.next();
    accounts.push({ id: acct.getCustomerId(), name: acct.getName(), obj: acct });
  }

  Logger.log('Found ' + accounts.length + ' accounts labeled "' + CONFIG.LABEL + '"');
  if (!accounts.length) return;

  // Single-MCC script: full clear is safe and simpler than filtering by account ID
  clearSheetData(tabs.campaign);
  clearSheetData(tabs.keyword);

  // ONE account iteration, TWO queries per account
  accounts.forEach(function(acct) {
    MccApp.select(acct.obj);
    Logger.log('Processing: ' + acct.name + ' (' + acct.id + ')');

    try { pullCampaigns(tabs.campaign, acct.id, acct.name, today, startDate); }
    catch(e) { Logger.log('Campaign ERR [' + acct.name + ']: ' + e.message); }

    try { pullKeywords(tabs.keyword,  acct.id, acct.name, today, startDate); }
    catch(e) { Logger.log('Keyword ERR  [' + acct.name + ']: ' + e.message); }
  });

  Logger.log('Done. ' + accounts.length + ' accounts processed.');
}

// ─── DATA PULL: CAMPAIGNS ─────────────────────────────────────
function pullCampaigns(sheet, accountId, accountName, today, startDate) {
  var query =
    'SELECT ' +
    '  campaign.name, campaign.status, campaign.advertising_channel_type, ' +
    '  campaign_budget.amount_micros, ' +
    '  metrics.impressions, metrics.clicks, metrics.ctr, ' +
    '  metrics.average_cpc, metrics.cost_micros, metrics.conversions, ' +
    '  metrics.conversions_from_interactions_rate, metrics.cost_per_conversion, ' +
    '  metrics.search_impression_share, ' +
    '  metrics.search_budget_lost_impression_share, ' +
    '  metrics.search_rank_lost_impression_share ' +
    'FROM campaign ' +
    'WHERE segments.date BETWEEN \'' + startDate + '\' AND \'' + today + '\' ' +
    '  AND campaign.status != \'REMOVED\' ' +
    '  AND campaign.advertising_channel_type != \'LOCAL_SERVICES\'';

  var report = AdsApp.report(query, {});
  var rows = report.rows();
  var data = [];

  while (rows.hasNext()) {
    var r = rows.next();
    data.push([
      today,
      accountName,
      accountId,
      r['campaign.name'],
      r['campaign.status'],
      friendlyType(r['campaign.advertising_channel_type']),
      micros(r['metrics.cost_micros']),
      r['metrics.clicks'],
      r['metrics.impressions'],
      pct(r['metrics.ctr']),
      micros(r['metrics.average_cpc']),
      r['metrics.conversions'],
      pct(r['metrics.conversions_from_interactions_rate']),
      micros(r['metrics.cost_per_conversion']),
      pct(r['metrics.search_impression_share']),
      pct(r['metrics.search_budget_lost_impression_share']),
      pct(r['metrics.search_rank_lost_impression_share'])
    ]);
  }

  appendRows(sheet, data);
  Logger.log('  Campaigns: ' + data.length + ' rows');
}

// ─── DATA PULL: KEYWORDS ──────────────────────────────────────
function pullKeywords(sheet, accountId, accountName, today, startDate) {
  var query =
    'SELECT ' +
    '  campaign.name, campaign.status, ' +
    '  ad_group.name, ' +
    '  ad_group_criterion.keyword.text, ad_group_criterion.keyword.match_type, ' +
    '  ad_group_criterion.status, ad_group_criterion.quality_info.quality_score, ' +
    '  metrics.cost_micros, metrics.clicks, metrics.impressions, metrics.ctr, ' +
    '  metrics.average_cpc, metrics.conversions, ' +
    '  metrics.conversions_from_interactions_rate, metrics.cost_per_conversion ' +
    'FROM keyword_view ' +
    'WHERE segments.date BETWEEN \'' + startDate + '\' AND \'' + today + '\' ' +
    '  AND campaign.status != \'REMOVED\' ' +
    '  AND ad_group_criterion.status != \'REMOVED\' ' +
    '  AND campaign.advertising_channel_type != \'LOCAL_SERVICES\' ' +
    'ORDER BY metrics.cost_micros DESC ' +
    'LIMIT ' + CONFIG.MAX_KEYWORDS;

  var report = AdsApp.report(query, {});
  var rows = report.rows();
  var data = [];

  while (rows.hasNext()) {
    var r = rows.next();
    data.push([
      today,
      accountName,
      accountId,
      r['ad_group_criterion.keyword.text'],
      r['ad_group_criterion.keyword.match_type'],
      r['ad_group.name'],
      r['campaign.name'],
      r['campaign.status'],
      micros(r['metrics.cost_micros']),
      r['metrics.clicks'],
      r['metrics.impressions'],
      pct(r['metrics.ctr']),
      micros(r['metrics.average_cpc']),
      r['metrics.conversions'],
      pct(r['metrics.conversions_from_interactions_rate']),
      micros(r['metrics.cost_per_conversion']),
      r['ad_group_criterion.quality_info.quality_score'] || ''
    ]);
  }

  appendRows(sheet, data);
  Logger.log('  Keywords: ' + data.length + ' rows');
}

// ─── SHEET UTILITIES ──────────────────────────────────────────
function getOrCreate(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  // Always refresh header row so columns stay in sync after script changes
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
       .setFontWeight('bold').setBackground('#1a2a5e').setFontColor('#ffffff');
  return sheet;
}

function appendRows(sheet, data) {
  if (!data.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
}

function clearSheetData(sheet) {
  if (sheet.getLastRow() <= 1) return;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
}

// ─── FORMAT UTILITIES ─────────────────────────────────────────
function fmtDate(daysAgo) {
  var d = new Date();
  d.setDate(d.getDate() - daysAgo);
  return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function micros(val) {
  var n = parseFloat(val) || 0;
  return (n / 1000000).toFixed(2);
}

function pct(val) {
  var n = parseFloat(val) || 0;
  return (n * 100).toFixed(2);
}

function friendlyType(type) {
  var map = {
    'SEARCH':          'Search',
    'DISPLAY':         'Display',
    'SHOPPING':        'Shopping',
    'VIDEO':           'Video',
    'PERFORMANCE_MAX': 'PMax',
    'SMART':           'Smart',
    'MULTI_CHANNEL':   'Multi-Channel'
  };
  return map[type] || type;
}
