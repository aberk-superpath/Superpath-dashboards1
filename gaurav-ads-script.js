// ============================================================
// GAURAV DASHBOARD — Google Ads MCC Script
// Install in BOTH MCCs:
//   Superpath MCC:   826-157-9112
//   Bell Media MCC:  562-306-0457
//
// Schedule: Daily at 6:00 AM (each MCC)
// Sheet ID: *** CREATE SHEET FIRST — paste ID here ***
// ============================================================

var CONFIG = {
  SHEET_ID:            'YOUR_GOOGLE_SHEET_ID',
  LABEL:               'Gaurav - Dashboard',
  LOOKBACK_DAYS:       30,
  HISTORY_DAYS:        90,   // campaign_data retention window
  MAX_KEYWORDS:        500,
  MAX_SEARCH_TERMS:    300,
  MAX_ADS:             500,
  TIMEZONE:            'America/New_York',
  // LSA-only accounts — skip entirely (no standard campaign data to pull)
  // Add any confirmed LSA-only Gaurav accounts here
  EXCLUDED_ACCOUNT_IDS: []
};

// ─── COLUMN HEADERS ──────────────────────────────────────────
var HEADERS = {
  campaign:    ['date','account_id','account_name','campaign_id','campaign_name','campaign_type',
                'campaign_status','daily_budget','impressions','clicks','ctr','avg_cpc','cost',
                'conversions','conv_rate','cost_per_conv','impression_share','is_lost_budget','is_lost_rank'],
  keyword:     ['date','account_id','account_name','campaign_name','ad_group_name','keyword_text',
                'match_type','status','quality_score','impressions','clicks','ctr','avg_cpc',
                'cost','conversions','cost_per_conv'],
  searchTerms: ['date','account_id','account_name','campaign_name','ad_group_name','search_term',
                'impressions','clicks','ctr','avg_cpc','cost','conversions','cost_per_conv'],
  ads:         ['date','account_id','account_name','campaign_name','ad_group_name','ad_id',
                'ad_type','ad_status','final_url','impressions','clicks','ctr','cost',
                'conversions','ad_strength']
};

// ─── MAIN ─────────────────────────────────────────────────────
function main() {
  Logger.log('Script started');
  var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  Logger.log('Sheet opened');
  var today     = fmtDate(0);
  var startDate = fmtDate(CONFIG.LOOKBACK_DAYS);

  // Get or create all tabs
  var tabs = {
    campaign:    getOrCreate(ss, 'campaign_data',  HEADERS.campaign),
    keyword:     getOrCreate(ss, 'keyword_data',   HEADERS.keyword),
    searchTerms: getOrCreate(ss, 'search_terms',   HEADERS.searchTerms),
    ads:         getOrCreate(ss, 'ad_data',        HEADERS.ads)
  };
  Logger.log('Tabs ready');

  // Collect labeled accounts for this MCC
  var accounts = [];
  Logger.log('Fetching labeled accounts...');
  var iter = MccApp.accounts().withCondition("LabelNames CONTAINS '" + CONFIG.LABEL + "'").get();
  while (iter.hasNext()) {
    var acct = iter.next();
    accounts.push({ id: acct.getCustomerId(), name: acct.getName(), obj: acct });
  }

  // Filter out LSA-only accounts
  accounts = accounts.filter(function(a) {
    return CONFIG.EXCLUDED_ACCOUNT_IDS.indexOf(a.id) === -1;
  });

  Logger.log('Found ' + accounts.length + ' accounts labeled "' + CONFIG.LABEL + '"');
  if (!accounts.length) return;

  var accountIds = accounts.map(function(a) { return a.id; });

  // Overwrite tabs: clear this MCC's accounts, then re-append fresh 30d data
  clearAccountRows(tabs.keyword,     accountIds);
  clearAccountRows(tabs.searchTerms, accountIds);
  clearAccountRows(tabs.ads,         accountIds);

  // Append tab: remove today's rows for this MCC (safe re-run), prune old history
  clearTodayAccountRows(tabs.campaign, accountIds, today);
  pruneOldRows(tabs.campaign, CONFIG.HISTORY_DAYS);

  // Process each account
  accounts.forEach(function(acct) {
    MccApp.select(acct.obj);
    Logger.log('Processing: ' + acct.name + ' (' + acct.id + ')');

    try { pullCampaigns(tabs.campaign,    acct.id, acct.name, today, startDate); }
    catch(e) { Logger.log('Campaign ERR [' + acct.name + ']: ' + e.message); }

    try { pullKeywords(tabs.keyword,      acct.id, acct.name, today, startDate); }
    catch(e) { Logger.log('Keyword ERR [' + acct.name + ']: ' + e.message); }

    try { pullSearchTerms(tabs.searchTerms, acct.id, acct.name, today, startDate); }
    catch(e) { Logger.log('SearchTerms ERR [' + acct.name + ']: ' + e.message); }

    try { pullAds(tabs.ads,               acct.id, acct.name, today, startDate); }
    catch(e) { Logger.log('Ads ERR [' + acct.name + ']: ' + e.message); }
  });

  Logger.log('Done. ' + accounts.length + ' accounts processed.');
}

// ─── DATA PULL: CAMPAIGNS ─────────────────────────────────────
function pullCampaigns(sheet, accountId, accountName, today, startDate) {
  var query =
    'SELECT ' +
    '  campaign.id, campaign.name, campaign.advertising_channel_type, ' +
    '  campaign.status, campaign_budget.amount_micros, ' +
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
      accountId,
      accountName,
      r['campaign.id'],
      r['campaign.name'],
      friendlyType(r['campaign.advertising_channel_type']),
      r['campaign.status'],
      micros(r['campaign_budget.amount_micros']),
      r['metrics.impressions'],
      r['metrics.clicks'],
      pct(r['metrics.ctr']),
      micros(r['metrics.average_cpc']),
      micros(r['metrics.cost_micros']),
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
    '  campaign.name, ad_group.name, ' +
    '  ad_group_criterion.keyword.text, ad_group_criterion.keyword.match_type, ' +
    '  ad_group_criterion.status, ad_group_criterion.quality_info.quality_score, ' +
    '  metrics.impressions, metrics.clicks, metrics.ctr, ' +
    '  metrics.average_cpc, metrics.cost_micros, metrics.conversions, ' +
    '  metrics.cost_per_conversion ' +
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
      accountId,
      accountName,
      r['campaign.name'],
      r['ad_group.name'],
      r['ad_group_criterion.keyword.text'],
      r['ad_group_criterion.keyword.match_type'],
      r['ad_group_criterion.status'],
      r['ad_group_criterion.quality_info.quality_score'] || '',
      r['metrics.impressions'],
      r['metrics.clicks'],
      pct(r['metrics.ctr']),
      micros(r['metrics.average_cpc']),
      micros(r['metrics.cost_micros']),
      r['metrics.conversions'],
      micros(r['metrics.cost_per_conversion'])
    ]);
  }

  appendRows(sheet, data);
  Logger.log('  Keywords: ' + data.length + ' rows');
}

// ─── DATA PULL: SEARCH TERMS ──────────────────────────────────
function pullSearchTerms(sheet, accountId, accountName, today, startDate) {
  var query =
    'SELECT ' +
    '  campaign.name, ad_group.name, search_term_view.search_term, ' +
    '  metrics.impressions, metrics.clicks, metrics.ctr, ' +
    '  metrics.average_cpc, metrics.cost_micros, metrics.conversions, ' +
    '  metrics.cost_per_conversion ' +
    'FROM search_term_view ' +
    'WHERE segments.date BETWEEN \'' + startDate + '\' AND \'' + today + '\' ' +
    '  AND campaign.status != \'REMOVED\' ' +
    '  AND campaign.advertising_channel_type != \'LOCAL_SERVICES\' ' +
    '  AND metrics.cost_micros > 0 ' +
    'ORDER BY metrics.cost_micros DESC ' +
    'LIMIT ' + CONFIG.MAX_SEARCH_TERMS;

  var report = AdsApp.report(query, {});
  var rows = report.rows();
  var data = [];

  while (rows.hasNext()) {
    var r = rows.next();
    data.push([
      today,
      accountId,
      accountName,
      r['campaign.name'],
      r['ad_group.name'],
      r['search_term_view.search_term'],
      r['metrics.impressions'],
      r['metrics.clicks'],
      pct(r['metrics.ctr']),
      micros(r['metrics.average_cpc']),
      micros(r['metrics.cost_micros']),
      r['metrics.conversions'],
      micros(r['metrics.cost_per_conversion'])
    ]);
  }

  appendRows(sheet, data);
  Logger.log('  Search terms: ' + data.length + ' rows');
}

// ─── DATA PULL: ADS ───────────────────────────────────────────
function pullAds(sheet, accountId, accountName, today, startDate) {
  var query =
    'SELECT ' +
    '  campaign.name, ad_group.name, ' +
    '  ad_group_ad.ad.id, ad_group_ad.ad.type, ' +
    '  ad_group_ad.ad.final_urls, ad_group_ad.status, ad_group_ad.ad_strength, ' +
    '  metrics.impressions, metrics.clicks, metrics.ctr, ' +
    '  metrics.cost_micros, metrics.conversions ' +
    'FROM ad_group_ad ' +
    'WHERE segments.date BETWEEN \'' + startDate + '\' AND \'' + today + '\' ' +
    '  AND campaign.status != \'REMOVED\' ' +
    '  AND ad_group_ad.status != \'REMOVED\' ' +
    '  AND campaign.advertising_channel_type != \'LOCAL_SERVICES\' ' +
    'ORDER BY metrics.cost_micros DESC ' +
    'LIMIT ' + CONFIG.MAX_ADS;

  var report = AdsApp.report(query, {});
  var rows = report.rows();
  var data = [];

  while (rows.hasNext()) {
    var r = rows.next();
    var urls = r['ad_group_ad.ad.final_urls'] || '';
    if (Array.isArray(urls)) urls = urls[0] || '';
    data.push([
      today,
      accountId,
      accountName,
      r['campaign.name'],
      r['ad_group.name'],
      r['ad_group_ad.ad.id'],
      r['ad_group_ad.ad.type'],
      r['ad_group_ad.status'],
      urls,
      r['metrics.impressions'],
      r['metrics.clicks'],
      pct(r['metrics.ctr']),
      micros(r['metrics.cost_micros']),
      r['metrics.conversions'],
      r['ad_group_ad.ad_strength'] || ''
    ]);
  }

  appendRows(sheet, data);
  Logger.log('  Ads: ' + data.length + ' rows');
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

function appendRows(sheet, data) {
  if (!data.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
}

function clearAccountRows(sheet, accountIds) {
  if (sheet.getLastRow() <= 1) return;
  var data = sheet.getDataRange().getValues();
  var keep = data.slice(1).filter(function(row) {
    return accountIds.indexOf(String(row[1])) === -1;
  });
  sheet.getRange(2, 1, data.length - 1, data[0].length).clearContent();
  if (keep.length > 0) {
    sheet.getRange(2, 1, keep.length, keep[0].length).setValues(keep);
  }
}

function clearTodayAccountRows(sheet, accountIds, today) {
  if (sheet.getLastRow() <= 1) return;
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === today && accountIds.indexOf(String(data[i][1])) !== -1) {
      sheet.deleteRow(i + 1);
    }
  }
}

function pruneOldRows(sheet, keepDays) {
  if (sheet.getLastRow() <= 1) return;
  var cutoff = fmtDate(keepDays);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) < cutoff) sheet.deleteRow(i + 1);
  }
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
