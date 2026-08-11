/*****************************************************************
 * DV SOCIAL — GOOGLE BUSINESS INTELLIGENCE ENGINE
 * VERSION 3.1 — HARDENED DAILY DISCOVERY + REFRESH
 *
 * PURPOSE:
 * Internal lead discovery system for DV Social.
 *
 * STACK:
 * Google Sheets
 * Google Apps Script
 * Google Places API (New) — Text Search + Place Details
 *
 * CORE:
 * Category x Locality search matrix
 * 100-500 new leads/day target, budget- and time-bounded
 * Place-ID deduplication
 * Live category views (QUERY formulas, never overwrite edits)
 * Coverage intelligence
 * Search rotation with priority scoring
 * Daily automation: new-lead discovery AND existing-lead refresh
 *
 * WHAT CHANGED IN 3.1
 * - Setup is idempotent: CONTROL and SEARCH QUEUE are never wiped
 *   by re-running "Setup V3 System".
 * - Fixed a crash when running discovery against an empty queue.
 * - Added a hard per-run API-call budget and execution-time budget
 *   shared across discovery and refresh, so a run can't blow past
 *   Apps Script's execution limit or spend unbounded API quota
 *   chasing a target across low-yield niche searches.
 * - "Run Next 10 Searches" now checks the API key and takes the
 *   same lock as the daily run.
 * - Category tabs (Food & Hospitality, Fashion & Retail, ...) are
 *   now live QUERY() formulas against MASTER DATABASE instead of
 *   static clear+rewrite copies. Editing Sales Status / Notes /
 *   Assigned To directly in a category tab used to get silently
 *   wiped on the next refresh — now there's nothing to edit there,
 *   so do that editing in MASTER DATABASE (use its column filter
 *   to view one category at a time).
 * - Added a daily "refresh existing leads" pass (Place Details)
 *   so rating / reviews / phone / website / open-closed status
 *   stay current instead of only growing new leads.
 * - RAW DATA now auto-trims instead of growing forever.
 * - Column lookups go through colOf_()/idxOf_() driven by the
 *   header list instead of magic numbers.
 *
 *****************************************************************/


/*****************************************************************
 * CONFIG
 *****************************************************************/

const DV3 = {

  VERSION: '3.1',

  ENDPOINT: 'https://places.googleapis.com/v1/places:searchText',
  DETAILS_ENDPOINT: 'https://places.googleapis.com/v1/places/',

  SHEETS: {
    CONTROL: 'CONTROL',
    MASTER: 'MASTER DATABASE',
    QUEUE: 'SEARCH QUEUE',
    COVERAGE: 'COVERAGE',
    LOG: 'RUN LOG',
    RAW: 'RAW DATA',
    FNB: 'Food & Hospitality',
    FASHION: 'Fashion & Retail',
    BEAUTY: 'Beauty & Wellness',
    REAL_ESTATE: 'Real Estate',
    HEALTHCARE: 'Healthcare',
    EDUCATION: 'Education',
    AUTOMOTIVE: 'Automotive',
    OTHER: 'Other'
  },

  API_KEY_PROPERTY: 'DV_PLACES_API_KEY',
  DAILY_TRIGGER_HANDLER: 'runDailyDiscovery',

  DEFAULT_DAILY_TARGET: 250,
  MAX_DAILY_TARGET: 500,
  DEFAULT_MAX_API_CALLS_PER_RUN: 400,
  DEFAULT_LEADS_TO_REFRESH_PER_DAY: 100,
  DEFAULT_MAX_RAW_ROWS: 20000,

  MAX_PAGES_PER_QUERY: 3,
  PAGE_SIZE: 20,
  API_RETRIES: 3,
  WRITE_BATCH_SIZE: 50,

  // Apps Script kills executions around 6 minutes. Stay well clear of
  // that so a run always finishes cleanly, logs, and releases its lock.
  MAX_RUN_MS: 4.5 * 60 * 1000,
  MIN_REFRESH_RESERVE_MS: 60 * 1000,

  /**
   * IMPORTANT: Google billing depends on requested fields.
   * We deliberately request only the fields sales actually uses.
   */
  FIELD_MASK: [
    'places.id',
    'places.displayName',
    'places.formattedAddress',
    'places.primaryType',
    'places.primaryTypeDisplayName',
    'places.types',
    'places.location',
    'places.googleMapsUri',
    'places.businessStatus',
    'places.nationalPhoneNumber',
    'places.internationalPhoneNumber',
    'places.websiteUri',
    'places.rating',
    'places.userRatingCount',
    'places.priceLevel',
    'places.regularOpeningHours',
    'nextPageToken'
  ].join(','),

  DETAILS_FIELD_MASK: [
    'id',
    'displayName',
    'formattedAddress',
    'primaryType',
    'primaryTypeDisplayName',
    'types',
    'location',
    'googleMapsUri',
    'businessStatus',
    'nationalPhoneNumber',
    'internationalPhoneNumber',
    'websiteUri',
    'rating',
    'userRatingCount',
    'priceLevel',
    'regularOpeningHours'
  ].join(','),

  MASTER_HEADERS: [
    'DV Lead ID', 'Business Name', 'Main Group', 'DV Category', 'Google Category',
    'Google Types', 'Search Locality', 'Address', 'Phone', 'International Phone',
    'Website', 'Google Rating', 'Google Reviews', 'Price Level', 'Google Maps URL',
    'Google Place ID', 'Latitude', 'Longitude', 'Business Status', 'Opening Hours',
    'Search Query', 'First Discovered', 'Last Updated', 'Lead Score', 'Lead Priority',
    'Contactability Score', 'Zomato Found', 'Zomato URL', 'Zomato Status',
    'Swiggy Found', 'Swiggy URL', 'Swiggy Status', 'F&B Verification', 'LinkedIn',
    'Instagram', 'Facebook', 'Email', 'Sales Status', 'Assigned To', 'Notes',
    'Last Refreshed'
  ],

  QUEUE_HEADERS: [
    'Queue ID', 'Main Group', 'Category', 'Keyword', 'Locality', 'Full Query',
    'Priority', 'Status', 'Last Run', 'API Results', 'New Leads', 'Duplicates',
    'Run Count'
  ],

  RAW_HEADERS: ['Timestamp', 'Query', 'Group', 'Category', 'Locality', 'Place ID', 'Raw JSON'],

  LOG_HEADERS: ['Timestamp', 'Run Type', 'Searches', 'API Results', 'New Leads', 'Duplicates', 'Execution Seconds', 'Status', 'Message']

};


/*****************************************************************
 * CATEGORY TAXONOMY
 *****************************************************************/

const DV_CATEGORIES = [

  { group: 'Food & Hospitality', category: 'Restaurants', queries: ['restaurants', 'popular restaurants', 'premium restaurants'] },
  { group: 'Food & Hospitality', category: 'Cafes', queries: ['cafes', 'coffee shops', 'premium cafes'] },
  { group: 'Food & Hospitality', category: 'Cloud Kitchens', queries: ['cloud kitchens', 'delivery kitchens', 'ghost kitchens'] },
  { group: 'Food & Hospitality', category: 'Bakeries', queries: ['bakeries', 'cake shops', 'premium bakeries'] },
  { group: 'Food & Hospitality', category: 'Fine Dining', queries: ['fine dining restaurants', 'luxury restaurants'] },
  { group: 'Food & Hospitality', category: 'QSR', queries: ['quick service restaurants', 'fast food restaurants'] },
  { group: 'Food & Hospitality', category: 'Desserts & Sweets', queries: ['dessert shops', 'sweet shops', 'ice cream shops'] },
  { group: 'Food & Hospitality', category: 'Bars & Pubs', queries: ['bars', 'pubs', 'lounges'] },
  { group: 'Food & Hospitality', category: 'Hotels', queries: ['hotels', 'luxury hotels', 'boutique hotels'] },
  { group: 'Food & Hospitality', category: 'Resorts', queries: ['resorts', 'luxury resorts'] },
  { group: 'Food & Hospitality', category: 'Catering', queries: ['catering companies', 'event caterers'] },

  { group: 'Fashion & Retail', category: 'Clothing Brands', queries: ['clothing stores', 'fashion brands', 'clothing brands'] },
  { group: 'Fashion & Retail', category: 'Boutiques', queries: ['fashion boutiques', 'designer boutiques'] },
  { group: 'Fashion & Retail', category: 'Designer Stores', queries: ['designer clothing stores', 'designer fashion stores'] },
  { group: 'Fashion & Retail', category: 'Jewellery', queries: ['jewellery stores', 'jewellery brands', 'luxury jewellery'] },
  { group: 'Fashion & Retail', category: 'Footwear', queries: ['shoe stores', 'footwear stores'] },
  { group: 'Fashion & Retail', category: 'Lifestyle', queries: ['lifestyle stores', 'home decor stores'] },

  { group: 'Beauty & Wellness', category: 'Salons', queries: ['salons', 'premium salons', 'beauty salons'] },
  { group: 'Beauty & Wellness', category: 'Spas', queries: ['spas', 'wellness spas'] },
  { group: 'Beauty & Wellness', category: 'Gyms', queries: ['gyms', 'fitness centres', 'premium gyms'] },
  { group: 'Beauty & Wellness', category: 'Skin Clinics', queries: ['skin clinics', 'dermatology clinics', 'aesthetic clinics'] },
  { group: 'Beauty & Wellness', category: 'Wellness Centres', queries: ['wellness centres', 'wellness clinics'] },

  { group: 'Real Estate', category: 'Developers', queries: ['real estate developers', 'property developers'] },
  { group: 'Real Estate', category: 'Builders', queries: ['builders', 'construction companies'] },
  { group: 'Real Estate', category: 'Interior Designers', queries: ['interior designers', 'interior design studios'] },
  { group: 'Real Estate', category: 'Architects', queries: ['architects', 'architecture firms'] },

  { group: 'Healthcare', category: 'Hospitals', queries: ['hospitals', 'private hospitals'] },
  { group: 'Healthcare', category: 'Dental Clinics', queries: ['dental clinics', 'dentists'] },
  { group: 'Healthcare', category: 'Eye Hospitals', queries: ['eye hospitals', 'eye clinics'] },
  { group: 'Healthcare', category: 'Fertility Clinics', queries: ['fertility clinics', 'IVF centres'] },

  { group: 'Education', category: 'Schools', queries: ['private schools', 'international schools'] },
  { group: 'Education', category: 'Colleges', queries: ['private colleges', 'degree colleges'] },
  { group: 'Education', category: 'Coaching', queries: ['coaching institutes', 'training institutes'] },

  { group: 'Automotive', category: 'Car Dealerships', queries: ['car dealerships', 'car showrooms'] },
  { group: 'Automotive', category: 'Bike Dealerships', queries: ['bike dealerships', 'motorcycle showrooms'] },
  { group: 'Automotive', category: 'Car Detailing', queries: ['car detailing', 'car detailing studios'] }

];


/*****************************************************************
 * HYDERABAD LOCALITY ENGINE
 *****************************************************************/

const DV_LOCALITIES = [
  'Jubilee Hills', 'Banjara Hills', 'Madhapur', 'HITEC City', 'Gachibowli',
  'Kondapur', 'Financial District', 'Kokapet', 'Nanakramguda', 'Manikonda',
  'Film Nagar', 'Begumpet', 'Somajiguda', 'Punjagutta', 'Ameerpet',
  'Secunderabad', 'Kukatpally', 'KPHB', 'Miyapur', 'Nizampet',
  'Kompally', 'Sainikpuri', 'AS Rao Nagar', 'Uppal', 'Nagole',
  'LB Nagar', 'Dilsukhnagar', 'Himayat Nagar', 'Abids', 'Basheerbagh',
  'Mehdipatnam', 'Tolichowki', 'Attapur', 'Shamshabad'
];

const DV_PREMIUM_LOCALITIES = [
  'Jubilee Hills', 'Banjara Hills', 'Madhapur', 'HITEC City', 'Gachibowli',
  'Financial District', 'Kokapet', 'Nanakramguda', 'Film Nagar'
];


/*****************************************************************
 * MENU
 *****************************************************************/

function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('DV INTELLIGENCE')
    .addItem('Setup / Update V3 System', 'setupV3')
    .addSeparator()
    .addItem('Set / Update API Key', 'setPlacesAPIKey')
    .addItem('Test API', 'testPlacesAPI')
    .addSeparator()
    .addItem('Update Search Queue (Add New Categories)', 'generateSearchQueue')
    .addItem('Run Daily Discovery', 'runDailyDiscovery')
    .addItem('Run Next 10 Searches', 'runNext10Searches')
    .addItem('Refresh Existing Leads Now', 'refreshExistingLeadsNow')
    .addSeparator()
    .addItem('Refresh Category Views', 'refreshCategoryViews')
    .addItem('Refresh Coverage', 'refreshCoverage')
    .addSeparator()
    .addItem('Enable Daily Automation', 'enableDailyAutomation')
    .addItem('Disable Daily Automation', 'disableDailyAutomation')
    .addToUi();

}


/*****************************************************************
 * SETUP (idempotent — safe to re-run any time)
 *****************************************************************/

function setupV3() {

  const ss = SpreadsheetApp.getActive();

  const queueSheetExisted = !!ss.getSheetByName(DV3.SHEETS.QUEUE) &&
    ss.getSheetByName(DV3.SHEETS.QUEUE).getLastRow() > 0;

  createControlSheet_(ss);
  createMasterSheet_(ss);
  createQueueSheet_(ss);
  createCoverageSheet_(ss);
  createLogSheet_(ss);
  createRawSheet_(ss);
  createCategorySheets_(ss);

  if (!queueSheetExisted) {
    generateSearchQueue();
  }

  refreshCategoryViews();
  refreshCoverage();
  updateControlStats_();

  SpreadsheetApp.getUi().alert(
    'DV GOOGLE INTELLIGENCE V3 READY\n\n' +
    'Next steps:\n\n' +
    '1. DV INTELLIGENCE -> Set / Update API Key\n' +
    '2. Test API\n' +
    '3. Open CONTROL and set your Daily Target\n' +
    '4. Run Daily Discovery (or Enable Daily Automation)\n\n' +
    'Safe to re-run this any time — it will never erase your\n' +
    'CONTROL settings, search queue progress, or leads.'
  );

}


/*****************************************************************
 * CONTROL (label/value settings sheet — idempotent)
 *****************************************************************/

function createControlSheet_(ss) {

  const sh = getOrCreateSheet_(ss, DV3.SHEETS.CONTROL);

  if (sh.getLastRow() === 0) {

    const rows = [
      ['DV SOCIAL — GOOGLE BUSINESS INTELLIGENCE', ''],
      ['VERSION', DV3.VERSION],
      ['', ''],
      ['SETTING', 'VALUE'],
      ['Daily New Lead Target', DV3.DEFAULT_DAILY_TARGET],
      ['Maximum Daily Target', DV3.MAX_DAILY_TARGET],
      ['Default City', 'Hyderabad'],
      ['Default State', 'Telangana'],
      ['Country', 'India'],
      ['Daily Automation Hour (0-23)', 7],
      ['Max API Calls Per Run', DV3.DEFAULT_MAX_API_CALLS_PER_RUN],
      ['Leads To Refresh Per Day', DV3.DEFAULT_LEADS_TO_REFRESH_PER_DAY],
      ['Max Raw Data Rows', DV3.DEFAULT_MAX_RAW_ROWS],
      ['Daily Automation Enabled', 'FALSE'],
      ['', ''],
      ['SYSTEM STATUS', ''],
      ['Last Discovery Run', ''],
      ['Last New Leads', 0],
      ['Last Leads Refreshed', 0],
      ['Total Master Leads', 0],
      ['Remaining Queue', 0]
    ];

    sh.getRange(1, 1, rows.length, 2).setValues(rows);

    sh.getRange('A1:B1').merge().setFontSize(18).setFontWeight('bold');
    sh.getRange('A4:B4').setFontWeight('bold');
    sh.getRange('A16:B16').setFontWeight('bold');

    sh.setColumnWidth(1, 280);
    sh.setColumnWidth(2, 300);

    return;

  }

  // Sheet already exists — add any settings introduced since, without
  // touching anything the user has already configured.
  ensureControlRow_(sh, 'Max API Calls Per Run', DV3.DEFAULT_MAX_API_CALLS_PER_RUN);
  ensureControlRow_(sh, 'Leads To Refresh Per Day', DV3.DEFAULT_LEADS_TO_REFRESH_PER_DAY);
  ensureControlRow_(sh, 'Max Raw Data Rows', DV3.DEFAULT_MAX_RAW_ROWS);
  ensureControlRow_(sh, 'Daily Automation Enabled', 'FALSE');
  ensureControlRow_(sh, 'Last Leads Refreshed', 0);

}


function ensureControlRow_(sh, label, defaultValue) {

  const lastRow = sh.getLastRow();
  const labels = sh.getRange(1, 1, lastRow, 1).getValues().map(function(r) { return String(r[0]).trim(); });

  if (labels.indexOf(label) === -1) {
    sh.getRange(lastRow + 1, 1, 1, 2).setValues([[label, defaultValue]]);
  }

}


function getControlValue_(label) {

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DV3.SHEETS.CONTROL);

  if (!sh) return null;

  const lastRow = sh.getLastRow();
  if (lastRow === 0) return null;

  const values = sh.getRange(1, 1, lastRow, 2).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === label) return values[i][1];
  }

  return null;

}


function setControlValue_(label, value) {

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DV3.SHEETS.CONTROL);

  if (!sh) return;

  const lastRow = sh.getLastRow();
  const labels = sh.getRange(1, 1, lastRow, 1).getValues().map(function(r) { return String(r[0]).trim(); });
  const idx = labels.indexOf(label);

  if (idx !== -1) {
    sh.getRange(idx + 1, 2).setValue(value);
  }

}


function getDailyTarget_() {
  const raw = Number(getControlValue_('Daily New Lead Target')) || DV3.DEFAULT_DAILY_TARGET;
  return Math.min(raw, DV3.MAX_DAILY_TARGET);
}

function getMaxApiCallsPerRun_() {
  return Math.max(1, Number(getControlValue_('Max API Calls Per Run')) || DV3.DEFAULT_MAX_API_CALLS_PER_RUN);
}

function getLeadsToRefreshPerDay_() {
  return Math.max(0, Number(getControlValue_('Leads To Refresh Per Day')) || DV3.DEFAULT_LEADS_TO_REFRESH_PER_DAY);
}

function getMaxRawRows_() {
  return Math.max(100, Number(getControlValue_('Max Raw Data Rows')) || DV3.DEFAULT_MAX_RAW_ROWS);
}

function getAutomationHour_() {
  const hour = Number(getControlValue_('Daily Automation Hour (0-23)'));
  if (isNaN(hour) || hour < 0 || hour > 23) return 7;
  return hour;
}


/*****************************************************************
 * MASTER DATABASE
 *****************************************************************/

function createMasterSheet_(ss) {

  const sh = getOrCreateSheet_(ss, DV3.SHEETS.MASTER);

  if (sh.getLastRow() > 1) {
    ensureHeaderColumn_(sh, 'Last Refreshed');
    return;
  }

  sh.clear();

  sh.getRange(1, 1, 1, DV3.MASTER_HEADERS.length)
    .setValues([DV3.MASTER_HEADERS])
    .setFontWeight('bold');

  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, DV3.MASTER_HEADERS.length).createFilter();
  sh.autoResizeColumns(1, DV3.MASTER_HEADERS.length);

}


function ensureHeaderColumn_(sh, headerName) {

  const lastCol = sh.getLastColumn();
  if (lastCol === 0) return;

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  if (headers.indexOf(headerName) === -1) {
    sh.getRange(1, lastCol + 1).setValue(headerName).setFontWeight('bold');
  }

}


/*****************************************************************
 * SEARCH QUEUE (idempotent — generateSearchQueue() populates it)
 *****************************************************************/

function createQueueSheet_(ss) {

  const sh = getOrCreateSheet_(ss, DV3.SHEETS.QUEUE);

  if (sh.getLastRow() > 0) {
    return; // never wipe existing queue progress
  }

  sh.getRange(1, 1, 1, DV3.QUEUE_HEADERS.length)
    .setValues([DV3.QUEUE_HEADERS])
    .setFontWeight('bold');

  sh.setFrozenRows(1);

}


/*****************************************************************
 * GENERATE / UPDATE SEARCH QUEUE (preserves existing progress)
 *****************************************************************/

function generateSearchQueue() {

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {

    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(DV3.SHEETS.QUEUE);

    if (!sh) {
      createQueueSheet_(ss);
      sh = ss.getSheetByName(DV3.SHEETS.QUEUE);
    }

    const existing = {};

    if (sh.getLastRow() > 1) {
      const old = sh.getRange(2, 1, sh.getLastRow() - 1, DV3.QUEUE_HEADERS.length).getValues();
      old.forEach(function(row) {
        existing[row[5]] = {
          status: row[7], lastRun: row[8], apiResults: row[9],
          newLeads: row[10], duplicates: row[11], runCount: row[12]
        };
      });
    }

    const rows = [];
    let queueNumber = 1;

    DV_CATEGORIES.forEach(function(category) {
      category.queries.forEach(function(keyword) {
        DV_LOCALITIES.forEach(function(locality) {

          const fullQuery = keyword + ' in ' + locality + ', Hyderabad, Telangana';
          const history = existing[fullQuery] || {};

          rows.push([
            'Q-' + String(queueNumber).padStart(5, '0'),
            category.group, category.category, keyword, locality, fullQuery,
            calculateQueryPriority_(category.group, category.category, locality),
            history.status || 'PENDING',
            history.lastRun || '',
            history.apiResults || 0,
            history.newLeads || 0,
            history.duplicates || 0,
            history.runCount || 0
          ]);

          queueNumber++;

        });
      });
    });

    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    }

    if (rows.length) {
      sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      sh.getRange(2, 1, rows.length, rows[0].length).sort([{ column: 7, ascending: false }]);
    }

    updateControlStats_();

  } finally {
    lock.releaseLock();
  }

}


/*****************************************************************
 * QUERY PRIORITY
 *****************************************************************/

function calculateQueryPriority_(group, category, locality) {

  let score = 50;

  if (group === 'Food & Hospitality') score += 25;
  if (group === 'Fashion & Retail') score += 20;
  if (group === 'Beauty & Wellness') score += 15;
  if (group === 'Real Estate') score += 15;

  if (DV_PREMIUM_LOCALITIES.indexOf(locality) !== -1) score += 20;
  if (category === 'Cloud Kitchens') score += 10;

  return score;

}


/*****************************************************************
 * RUN DAILY DISCOVERY (new leads + existing-lead refresh, in one
 * time- and budget-bounded run)
 *****************************************************************/

function runDailyDiscovery() {

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  const started = new Date();
  const deadline = started.getTime() + DV3.MAX_RUN_MS;

  try {

    const apiKey = getPlacesAPIKey_();
    if (!apiKey) throw new Error('Google Places API key not configured.');

    const target = getDailyTarget_();
    const budget = { remaining: getMaxApiCallsPerRun_() };

    // Reserve time at the tail of the run for the refresh pass so it
    // never gets starved by a long discovery phase.
    const discoveryDeadline = deadline - DV3.MIN_REFRESH_RESERVE_MS;

    const discoveryResult = runDiscoveryBatch_(target, Infinity, budget, discoveryDeadline);
    const refreshResult = refreshExistingLeads_(budget, deadline);

    refreshCategoryViews();
    refreshCoverage();
    updateControlStats_();

    setControlValue_('Last Discovery Run', new Date());
    setControlValue_('Last New Leads', discoveryResult.newLeads);
    setControlValue_('Last Leads Refreshed', refreshResult.updated);

    const seconds = ((new Date() - started) / 1000).toFixed(2);

    logSystemRun_(
      'DAILY DISCOVERY', discoveryResult.searchesRun, discoveryResult.apiResults,
      discoveryResult.newLeads, discoveryResult.duplicates, seconds, 'SUCCESS',
      'Target: ' + target + ' | Refreshed: ' + refreshResult.updated +
      ' | Newly Closed: ' + refreshResult.closed
    );

    SpreadsheetApp.getUi().alert(
      'DV DAILY DISCOVERY COMPLETE\n\n' +
      'Target: ' + target +
      '\nNew Leads: ' + discoveryResult.newLeads +
      '\nSearches Run: ' + discoveryResult.searchesRun +
      '\nAPI Results: ' + discoveryResult.apiResults +
      '\nDuplicates: ' + discoveryResult.duplicates +
      '\n\nLeads Refreshed: ' + refreshResult.updated +
      '\nNewly Closed: ' + refreshResult.closed +
      '\n\nTime: ' + seconds + ' sec'
    );

  } catch (error) {

    logSystemRun_('DAILY DISCOVERY', 0, 0, 0, 0, 0, 'ERROR', error.message);

    // A trigger-fired run has no UI to alert — only show a dialog if
    // this was launched interactively from the menu.
    try { SpreadsheetApp.getUi().alert('V3 ERROR\n\n' + error.message); } catch (e) {}

  } finally {

    lock.releaseLock();

  }

}


/*****************************************************************
 * RUN NEXT 10 (manual, bounded batch)
 *****************************************************************/

function runNext10Searches() {

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {

    const apiKey = getPlacesAPIKey_();

    if (!apiKey) {
      SpreadsheetApp.getUi().alert('Set your Google Places API key first (DV INTELLIGENCE -> Set / Update API Key).');
      return;
    }

    const budget = { remaining: getMaxApiCallsPerRun_() };
    const deadline = new Date().getTime() + DV3.MAX_RUN_MS;

    const result = runDiscoveryBatch_(Infinity, 10, budget, deadline);

    refreshCategoryViews();
    refreshCoverage();
    updateControlStats_();

    SpreadsheetApp.getUi().alert(
      'SEARCHES COMPLETE\n\n' +
      'Searches Run: ' + result.searchesRun +
      '\nNew Leads: ' + result.newLeads +
      '\nDuplicates: ' + result.duplicates
    );

  } finally {

    lock.releaseLock();

  }

}


/*****************************************************************
 * SHARED DISCOVERY CORE
 * Picks the highest-priority PENDING/RETRY queue rows and executes
 * them, bounded by new-lead target, search count, API budget, and
 * a hard deadline — whichever comes first.
 *****************************************************************/

function runDiscoveryBatch_(target, maxSearches, budget, deadline) {

  const ss = SpreadsheetApp.getActive();
  const queue = ss.getSheetByName(DV3.SHEETS.QUEUE);
  const master = ss.getSheetByName(DV3.SHEETS.MASTER);

  const result = { searchesRun: 0, apiResults: 0, newLeads: 0, duplicates: 0 };

  if (!queue || !master) {
    throw new Error('Run Setup / Update V3 System first.');
  }

  if (queue.getLastRow() <= 1) {
    return result; // queue has no rows yet
  }

  const existingIds = getExistingPlaceIds_(master);

  const queueRows = queue.getRange(2, 1, queue.getLastRow() - 1, DV3.QUEUE_HEADERS.length).getValues();

  const candidates = [];

  queueRows.forEach(function(row, index) {
    const status = String(row[7] || '');
    if (status === 'PENDING' || status === 'RETRY') {
      candidates.push({ sheetRow: index + 2, data: row });
    }
  });

  candidates.sort(function(a, b) { return Number(b.data[6]) - Number(a.data[6]); });

  for (let i = 0; i < candidates.length; i++) {

    if (result.newLeads >= target) break;
    if (result.searchesRun >= maxSearches) break;
    if (budget.remaining <= 0) break;
    if (new Date().getTime() >= deadline) break;

    const item = candidates[i];
    const row = item.data;

    const searchResult = executeQueueSearch_(row, existingIds, target - result.newLeads, budget, deadline);

    result.searchesRun++;
    result.newLeads += searchResult.newLeads;
    result.apiResults += searchResult.apiResults;
    result.duplicates += searchResult.duplicates;

    queue.getRange(item.sheetRow, 8, 1, 6).setValues([[
      'DONE', new Date(), searchResult.apiResults, searchResult.newLeads,
      searchResult.duplicates, Number(row[12] || 0) + 1
    ]]);

  }

  // Queue was already fully exhausted before this run started — reset
  // it for tomorrow. (This run still reports 0 new leads; the reset
  // takes effect on the next run.)
  if (candidates.length === 0) {
    resetSearchQueue_();
  }

  return result;

}


/*****************************************************************
 * EXECUTE ONE QUEUE SEARCH (budget- and deadline-aware)
 *****************************************************************/

function executeQueueSearch_(queueRow, existingIds, remainingTarget, budget, deadline) {

  const group = queueRow[1];
  const category = queueRow[2];
  const locality = queueRow[4];
  const query = queueRow[5];

  const apiKey = getPlacesAPIKey_();

  let token = null;
  let page = 0;
  let apiResults = 0;
  let duplicates = 0;
  let newLeads = 0;

  const rows = [];
  const rawRows = [];

  while (page < DV3.MAX_PAGES_PER_QUERY && newLeads < remainingTarget) {

    if (budget.remaining <= 0) break;
    if (new Date().getTime() >= deadline) break;

    budget.remaining--;

    const response = callPlacesAPI_(query, token, apiKey);
    page++;

    const places = response.places || [];
    apiResults += places.length;

    places.forEach(function(place) {

      if (newLeads >= remainingTarget) return;

      const placeId = place.id || '';

      rawRows.push([new Date(), query, group, category, locality, placeId, JSON.stringify(place)]);

      if (placeId && existingIds.has(placeId)) {
        duplicates++;
        return;
      }

      const lead = placeToLeadRow_(place, group, category, locality, query);
      rows.push(lead);

      if (placeId) existingIds.add(placeId);

      newLeads++;

    });

    token = response.nextPageToken || null;
    if (!token) break;

    Utilities.sleep(500);

  }

  appendMasterRows_(rows);
  appendRawRows_(rawRows);

  return { apiResults: apiResults, newLeads: newLeads, duplicates: duplicates };

}


/*****************************************************************
 * PLACE -> MASTER ROW
 *****************************************************************/

function placeToLeadRow_(place, group, category, locality, query) {

  const name = place.displayName ? place.displayName.text : '';
  const googleCategory = place.primaryTypeDisplayName ? place.primaryTypeDisplayName.text : '';
  const googleTypes = place.types ? place.types.join(', ') : '';
  const rating = Number(place.rating || 0);
  const reviews = Number(place.userRatingCount || 0);
  const phone = place.nationalPhoneNumber || '';
  const internationalPhone = place.internationalPhoneNumber || '';
  const website = place.websiteUri || '';
  const businessStatus = place.businessStatus || '';

  const score = calculateLeadScoreV3_(group, category, rating, reviews, website, phone, locality, businessStatus);
  const priority = calculateLeadPriority_(score);
  const contactability = calculateContactability_(website, phone);
  const isFNB = group === 'Food & Hospitality';
  const now = new Date();

  return [
    createDVLeadId_(place.id), name, group, category, googleCategory, googleTypes,
    locality, place.formattedAddress || '', phone, internationalPhone, website,
    rating || '', reviews || '', place.priceLevel || '', place.googleMapsUri || '',
    place.id || '', place.location ? place.location.latitude : '',
    place.location ? place.location.longitude : '', businessStatus,
    extractOpeningHours_(place), query, now, now, score, priority, contactability,
    isFNB ? 'NOT CHECKED' : 'N/A', '', isFNB ? 'PENDING' : 'N/A',
    isFNB ? 'NOT CHECKED' : 'N/A', '', isFNB ? 'PENDING' : 'N/A',
    isFNB ? 'PENDING' : 'N/A', '', '', '', '', 'NEW', '', '', now
  ];

}


/*****************************************************************
 * LEAD SCORE
 *****************************************************************/

function calculateLeadScoreV3_(group, category, rating, reviews, website, phone, locality, businessStatus) {

  let score = 0;

  if (reviews >= 5000) score += 25;
  else if (reviews >= 2000) score += 22;
  else if (reviews >= 1000) score += 20;
  else if (reviews >= 500) score += 17;
  else if (reviews >= 100) score += 12;
  else if (reviews >= 25) score += 7;
  else score += 3;

  if (rating >= 4.5) score += 15;
  else if (rating >= 4.2) score += 13;
  else if (rating >= 4) score += 10;
  else if (rating >= 3.5) score += 7;
  else if (rating > 0) score += 3;

  if (website) score += 15;
  if (phone) score += 10;

  if (group === 'Food & Hospitality') score += 15;
  else if (group === 'Fashion & Retail') score += 15;
  else if (group === 'Beauty & Wellness') score += 12;
  else if (group === 'Real Estate') score += 12;
  else score += 8;

  if (DV_PREMIUM_LOCALITIES.indexOf(locality) !== -1) score += 10;
  if (category === 'Cloud Kitchens') score += 5;

  score = Math.min(score, 100);

  // A closed business is not a sellable lead, however strong its
  // historical numbers were.
  if (businessStatus === 'CLOSED_PERMANENTLY') return 0;
  if (businessStatus === 'CLOSED_TEMPORARILY') return Math.round(score * 0.5);

  return score;

}


/*****************************************************************
 * PRIORITY / CONTACTABILITY
 *****************************************************************/

function calculateLeadPriority_(score) {
  if (score >= 80) return 'HOT';
  if (score >= 65) return 'HIGH';
  if (score >= 45) return 'MEDIUM';
  return 'LOW';
}


function calculateContactability_(website, phone) {
  let score = 0;
  if (website) score += 50;
  if (phone) score += 50;
  return score;
}


/*****************************************************************
 * GOOGLE PLACES API (Text Search + Place Details, with retry)
 *****************************************************************/

function callPlacesAPI_(query, pageToken, apiKey) {

  const payload = { textQuery: query, pageSize: DV3.PAGE_SIZE, languageCode: 'en', regionCode: 'IN' };
  if (pageToken) payload.pageToken = pageToken;

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-Goog-Api-Key': apiKey, 'X-Goog-FieldMask': DV3.FIELD_MASK },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  return fetchWithRetry_(DV3.ENDPOINT, options);

}


function callPlaceDetailsAPI_(placeId, apiKey) {

  const url = DV3.DETAILS_ENDPOINT + encodeURIComponent(placeId);

  const options = {
    method: 'get',
    headers: { 'X-Goog-Api-Key': apiKey, 'X-Goog-FieldMask': DV3.DETAILS_FIELD_MASK },
    muteHttpExceptions: true
  };

  return fetchWithRetry_(url, options);

}


function fetchWithRetry_(url, options) {

  let lastError = null;

  for (let attempt = 1; attempt <= DV3.API_RETRIES; attempt++) {

    try {

      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const text = response.getContentText();

      if (code >= 200 && code < 300) return JSON.parse(text);

      lastError = new Error('Places API ' + code + ': ' + text);

      if (code === 429 || code >= 500) {
        Utilities.sleep(attempt * 1500);
        continue;
      }

      throw lastError;

    } catch (error) {

      lastError = error;
      if (attempt < DV3.API_RETRIES) Utilities.sleep(attempt * 1500);

    }

  }

  throw lastError;

}


/*****************************************************************
 * EXISTING LEAD REFRESH (Place Details lookup)
 *****************************************************************/

function refreshExistingLeadsNow() {

  const apiKey = getPlacesAPIKey_();

  if (!apiKey) {
    SpreadsheetApp.getUi().alert('Set your Google Places API key first (DV INTELLIGENCE -> Set / Update API Key).');
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {

    const budget = { remaining: getMaxApiCallsPerRun_() };
    const deadline = new Date().getTime() + DV3.MAX_RUN_MS;

    const result = refreshExistingLeads_(budget, deadline);

    SpreadsheetApp.getUi().alert(
      'LEAD REFRESH COMPLETE\n\n' +
      'Checked: ' + result.checked +
      '\nUpdated: ' + result.updated +
      '\nNewly Closed: ' + result.closed +
      '\nNot Found (removed from Google): ' + result.notFound
    );

  } finally {

    lock.releaseLock();

  }

}


function refreshExistingLeads_(budget, deadline) {

  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(DV3.SHEETS.MASTER);
  const apiKey = getPlacesAPIKey_();

  const stats = { checked: 0, updated: 0, closed: 0, notFound: 0 };

  if (!apiKey || !master) return stats;

  const lastRow = master.getLastRow();
  if (lastRow <= 1) return stats;

  const numRows = lastRow - 1;
  const numCols = DV3.MASTER_HEADERS.length;
  const data = master.getRange(2, 1, numRows, numCols).getValues();

  const idIdx = idxOf_('Google Place ID');
  const statusIdx = idxOf_('Sales Status');
  const refreshedIdx = idxOf_('Last Refreshed');
  const groupIdx = idxOf_('Main Group');
  const categoryIdx = idxOf_('DV Category');
  const localityIdx = idxOf_('Search Locality');

  const candidates = data
    .map(function(row, i) {
      return {
        rowIndex: i + 2,
        placeId: row[idIdx],
        status: row[statusIdx],
        group: row[groupIdx],
        category: row[categoryIdx],
        locality: row[localityIdx],
        lastRefreshed: row[refreshedIdx] ? new Date(row[refreshedIdx]).getTime() : 0
      };
    })
    .filter(function(c) { return c.placeId; })
    .sort(function(a, b) { return a.lastRefreshed - b.lastRefreshed; });

  const perDayLimit = getLeadsToRefreshPerDay_();
  const toRefresh = candidates.slice(0, perDayLimit);

  for (let i = 0; i < toRefresh.length; i++) {

    const candidate = toRefresh[i];

    if (budget.remaining <= 0) break;
    if (new Date().getTime() >= deadline) break;

    budget.remaining--;
    stats.checked++;

    try {

      const place = callPlaceDetailsAPI_(candidate.placeId, apiKey);
      updateMasterRow_(master, candidate, place);
      stats.updated++;

      if (place.businessStatus === 'CLOSED_PERMANENTLY' || place.businessStatus === 'CLOSED_TEMPORARILY') {
        stats.closed++;
      }

    } catch (error) {

      master.getRange(candidate.rowIndex, colOf_('Business Status')).setValue('NOT FOUND');
      master.getRange(candidate.rowIndex, colOf_('Last Refreshed')).setValue(new Date());
      stats.notFound++;

    }

  }

  return stats;

}


function updateMasterRow_(sheet, candidate, place) {

  const rating = Number(place.rating || 0);
  const reviews = Number(place.userRatingCount || 0);
  const phone = place.nationalPhoneNumber || '';
  const internationalPhone = place.internationalPhoneNumber || '';
  const website = place.websiteUri || '';
  const businessStatus = place.businessStatus || '';
  const openingHours = extractOpeningHours_(place);

  const score = calculateLeadScoreV3_(candidate.group, candidate.category, rating, reviews, website, phone, candidate.locality, businessStatus);
  const priority = calculateLeadPriority_(score);
  const contactability = calculateContactability_(website, phone);

  const row = candidate.rowIndex;

  sheet.getRange(row, colOf_('Phone')).setValue(phone);
  sheet.getRange(row, colOf_('International Phone')).setValue(internationalPhone);
  sheet.getRange(row, colOf_('Website')).setValue(website);
  sheet.getRange(row, colOf_('Google Rating')).setValue(rating || '');
  sheet.getRange(row, colOf_('Google Reviews')).setValue(reviews || '');
  sheet.getRange(row, colOf_('Price Level')).setValue(place.priceLevel || '');
  sheet.getRange(row, colOf_('Business Status')).setValue(businessStatus);
  sheet.getRange(row, colOf_('Opening Hours')).setValue(openingHours);
  sheet.getRange(row, colOf_('Last Updated')).setValue(new Date());
  sheet.getRange(row, colOf_('Lead Score')).setValue(score);
  sheet.getRange(row, colOf_('Lead Priority')).setValue(priority);
  sheet.getRange(row, colOf_('Contactability Score')).setValue(contactability);
  sheet.getRange(row, colOf_('Last Refreshed')).setValue(new Date());

  // Only auto-flip Sales Status to CLOSED if it's still in an automated
  // state — never override a manual pipeline stage like CONTACTED / WON.
  if (businessStatus === 'CLOSED_PERMANENTLY' || businessStatus === 'CLOSED_TEMPORARILY') {
    if (candidate.status === 'NEW' || candidate.status === 'CLOSED') {
      sheet.getRange(row, colOf_('Sales Status')).setValue('CLOSED');
    }
  }

}


/*****************************************************************
 * APPEND MASTER / RAW
 *****************************************************************/

function appendMasterRows_(rows) {

  if (!rows.length) return;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DV3.SHEETS.MASTER);

  for (let i = 0; i < rows.length; i += DV3.WRITE_BATCH_SIZE) {
    const batch = rows.slice(i, i + DV3.WRITE_BATCH_SIZE);
    sh.getRange(sh.getLastRow() + 1, 1, batch.length, batch[0].length).setValues(batch);
  }

}


function createRawSheet_(ss) {

  const sh = getOrCreateSheet_(ss, DV3.SHEETS.RAW);

  if (sh.getLastRow() > 1) return;

  sh.clear();

  sh.getRange(1, 1, 1, DV3.RAW_HEADERS.length)
    .setValues([DV3.RAW_HEADERS])
    .setFontWeight('bold');

  sh.setFrozenRows(1);

}


function appendRawRows_(rows) {

  if (!rows.length) return;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DV3.SHEETS.RAW);

  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

}


function trimRawData_() {

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DV3.SHEETS.RAW);

  if (!sh) return;

  const maxRows = getMaxRawRows_();
  const dataRows = sh.getLastRow() - 1;

  if (dataRows > maxRows) {
    sh.deleteRows(2, dataRows - maxRows); // oldest rows are at the top
  }

}


/*****************************************************************
 * DEDUPLICATION
 *****************************************************************/

function getExistingPlaceIds_(master) {

  const set = new Set();

  if (!master || master.getLastRow() <= 1) return set;

  const values = master.getRange(2, colOf_('Google Place ID'), master.getLastRow() - 1, 1).getValues();

  values.forEach(function(row) {
    const id = String(row[0] || '').trim();
    if (id) set.add(id);
  });

  return set;

}


/*****************************************************************
 * CATEGORY SHEETS — live QUERY() views, never overwritten data
 *****************************************************************/

function createCategorySheets_(ss) {

  const names = [
    DV3.SHEETS.FNB, DV3.SHEETS.FASHION, DV3.SHEETS.BEAUTY, DV3.SHEETS.REAL_ESTATE,
    DV3.SHEETS.HEALTHCARE, DV3.SHEETS.EDUCATION, DV3.SHEETS.AUTOMOTIVE, DV3.SHEETS.OTHER
  ];

  names.forEach(function(name) { getOrCreateSheet_(ss, name); });

}


function refreshCategoryViews() {

  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(DV3.SHEETS.MASTER);

  if (!master) return;

  const groupCol = columnToLetter_(colOf_('Main Group'));
  const range = "'" + DV3.SHEETS.MASTER + "'!A:" + columnToLetter_(DV3.MASTER_HEADERS.length);

  const mapping = {
    'Food & Hospitality': DV3.SHEETS.FNB,
    'Fashion & Retail': DV3.SHEETS.FASHION,
    'Beauty & Wellness': DV3.SHEETS.BEAUTY,
    'Real Estate': DV3.SHEETS.REAL_ESTATE,
    'Healthcare': DV3.SHEETS.HEALTHCARE,
    'Education': DV3.SHEETS.EDUCATION,
    'Automotive': DV3.SHEETS.AUTOMOTIVE
  };

  const groups = Object.keys(mapping);

  groups.forEach(function(group) {

    const sh = ss.getSheetByName(mapping[group]);
    if (!sh) return;

    const formula = '=IFERROR(QUERY(' + range + ',"select * where ' + groupCol +
      " = '" + group.replace(/'/g, "\\'") + '\'", 1), "No leads yet.")';

    writeCategoryFormula_(sh, formula);

  });

  const otherSheet = ss.getSheetByName(DV3.SHEETS.OTHER);

  if (otherSheet) {

    const exclusions = groups.map(function(g) { return groupCol + " != '" + g.replace(/'/g, "\\'") + "'"; }).join(' and ');
    const otherFormula = '=IFERROR(QUERY(' + range + ',"select * where ' + groupCol +
      ' is not null and ' + exclusions + '", 1), "No leads yet.")';

    writeCategoryFormula_(otherSheet, otherFormula);

  }

}


function writeCategoryFormula_(sh, formula) {

  sh.getRange(1, 1).setFormula(formula);
  sh.getRange(1, 1).setNote(
    'This is a live view of MASTER DATABASE — do not edit here.\n' +
    'Update Sales Status / Notes / Assigned To in MASTER DATABASE\n' +
    '(use its column filter to view one category at a time).'
  );
  sh.setFrozenRows(1);

}


/*****************************************************************
 * COVERAGE
 *****************************************************************/

function createCoverageSheet_(ss) {

  const sh = getOrCreateSheet_(ss, DV3.SHEETS.COVERAGE);

  sh.clear();

  sh.getRange(1, 1, 1, 10)
    .setValues([[
      'Main Group', 'Category', 'Total Leads', 'With Phone', 'Phone %',
      'With Website', 'Website %', 'Hot Leads', 'High Leads', 'Last Updated'
    ]])
    .setFontWeight('bold');

  sh.setFrozenRows(1);

}


function refreshCoverage() {

  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(DV3.SHEETS.MASTER);
  const coverage = ss.getSheetByName(DV3.SHEETS.COVERAGE);

  if (!master || !coverage) return;
  if (master.getLastRow() <= 1) return;

  const data = master.getRange(2, 1, master.getLastRow() - 1, DV3.MASTER_HEADERS.length).getValues();

  const groupIdx = idxOf_('Main Group');
  const categoryIdx = idxOf_('DV Category');
  const phoneIdx = idxOf_('Phone');
  const websiteIdx = idxOf_('Website');
  const priorityIdx = idxOf_('Lead Priority');

  const stats = {};

  data.forEach(function(row) {

    const group = row[groupIdx];
    const category = row[categoryIdx];
    const key = group + '|' + category;

    if (!stats[key]) {
      stats[key] = { group: group, category: category, total: 0, phone: 0, website: 0, hot: 0, high: 0 };
    }

    const s = stats[key];
    s.total++;
    if (row[phoneIdx]) s.phone++;
    if (row[websiteIdx]) s.website++;
    if (row[priorityIdx] === 'HOT') s.hot++;
    if (row[priorityIdx] === 'HIGH') s.high++;

  });

  const rows = Object.keys(stats).map(function(key) {
    const s = stats[key];
    return [
      s.group, s.category, s.total, s.phone, s.total ? s.phone / s.total : 0,
      s.website, s.total ? s.website / s.total : 0, s.hot, s.high, new Date()
    ];
  });

  if (coverage.getLastRow() > 1) {
    coverage.getRange(2, 1, coverage.getLastRow() - 1, coverage.getLastColumn()).clearContent();
  }

  if (rows.length) {
    coverage.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    coverage.getRange(2, 5, rows.length, 1).setNumberFormat('0.0%');
    coverage.getRange(2, 7, rows.length, 1).setNumberFormat('0.0%');
  }

}


/*****************************************************************
 * RUN LOG
 *****************************************************************/

function createLogSheet_(ss) {

  const sh = getOrCreateSheet_(ss, DV3.SHEETS.LOG);

  if (sh.getLastRow() > 1) return;

  sh.clear();

  sh.getRange(1, 1, 1, DV3.LOG_HEADERS.length)
    .setValues([DV3.LOG_HEADERS])
    .setFontWeight('bold');

  sh.setFrozenRows(1);

}


function logSystemRun_(type, searches, api, leads, duplicates, seconds, status, message) {

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DV3.SHEETS.LOG);

  if (!sh) return;

  sh.appendRow([new Date(), type, searches, api, leads, duplicates, seconds, status, message]);

}


/*****************************************************************
 * RESET SEARCH QUEUE
 *****************************************************************/

function resetSearchQueue_() {

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DV3.SHEETS.QUEUE);

  if (!sh || sh.getLastRow() <= 1) return;

  sh.getRange(2, 8, sh.getLastRow() - 1, 1).setValue('PENDING');

}


/*****************************************************************
 * CONTROL STATS
 *****************************************************************/

function updateControlStats_() {

  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName(DV3.SHEETS.MASTER);
  const queue = ss.getSheetByName(DV3.SHEETS.QUEUE);

  const total = master ? Math.max(master.getLastRow() - 1, 0) : 0;

  let remaining = 0;

  if (queue && queue.getLastRow() > 1) {
    const statuses = queue.getRange(2, 8, queue.getLastRow() - 1, 1).getValues();
    statuses.forEach(function(row) {
      if (row[0] === 'PENDING' || row[0] === 'RETRY') remaining++;
    });
  }

  setControlValue_('Total Master Leads', total);
  setControlValue_('Remaining Queue', remaining);

  trimRawData_();

}


/*****************************************************************
 * API KEY
 *****************************************************************/

function setPlacesAPIKey() {

  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    'Google Places API Key',
    'Paste your API key. It will be stored in Script Properties, not inside the sheet.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const key = response.getResponseText().trim();

  if (!key) {
    ui.alert('API key cannot be blank.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty(DV3.API_KEY_PROPERTY, key);
  ui.alert('API key saved securely.');

}


function getPlacesAPIKey_() {
  return PropertiesService.getScriptProperties().getProperty(DV3.API_KEY_PROPERTY) || '';
}


/*****************************************************************
 * TEST API
 *****************************************************************/

function testPlacesAPI() {

  try {

    const key = getPlacesAPIKey_();
    if (!key) throw new Error('Set API key first.');

    const response = callPlacesAPI_('cafes in Jubilee Hills, Hyderabad, Telangana', null, key);
    const count = response.places ? response.places.length : 0;

    SpreadsheetApp.getUi().alert('API SUCCESS\n\n' + count + ' businesses returned.');

  } catch (error) {

    SpreadsheetApp.getUi().alert('API TEST FAILED\n\n' + error.message);

  }

}


/*****************************************************************
 * DAILY AUTOMATION
 *****************************************************************/

function enableDailyAutomation() {

  disableDailyAutomation();

  const hour = getAutomationHour_();

  ScriptApp.newTrigger(DV3.DAILY_TRIGGER_HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();

  setControlValue_('Daily Automation Enabled', 'TRUE');

  SpreadsheetApp.getUi().alert('Daily automation enabled around ' + hour + ':00.');

}


function disableDailyAutomation() {

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === DV3.DAILY_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  setControlValue_('Daily Automation Enabled', 'FALSE');

}


/*****************************************************************
 * OPENING HOURS
 *****************************************************************/

function extractOpeningHours_(place) {

  try {
    if (place.regularOpeningHours && place.regularOpeningHours.weekdayDescriptions) {
      return place.regularOpeningHours.weekdayDescriptions.join(' | ');
    }
  } catch (e) {}

  return '';

}


/*****************************************************************
 * LEAD ID
 *****************************************************************/

function createDVLeadId_(placeId) {

  if (placeId) {
    return 'DV-' + placeId.replace(/[^a-zA-Z0-9]/g, '').substring(0, 14).toUpperCase();
  }

  return 'DV-' + Utilities.getUuid().substring(0, 10).toUpperCase();

}


/*****************************************************************
 * COLUMN LOOKUP HELPERS
 * Resolve Master Database columns by header name instead of magic
 * numbers, so adding/reordering headers can't silently break scoring,
 * dedup, coverage, or refresh.
 *****************************************************************/

function colOf_(headerName) {
  const idx = DV3.MASTER_HEADERS.indexOf(headerName);
  if (idx === -1) throw new Error('Unknown MASTER DATABASE column: ' + headerName);
  return idx + 1;
}


function idxOf_(headerName) {
  return colOf_(headerName) - 1;
}


function columnToLetter_(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}


/*****************************************************************
 * SHEET HELPER
 *****************************************************************/

function getOrCreateSheet_(ss, name) {

  let sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
  }

  return sh;

}
