/***************************************************************
 * DV SOCIAL — SALES INTELLIGENCE ENGINE
 * VERSION 2.0 — DAILY AUTO-REFINEMENT
 *
 * Stack:
 * Google Sheets
 * Google Apps Script
 * Google Places API (New) — Text Search + Place Details
 *
 * -------------------------------------------------------------
 * WHAT'S NEW IN V2
 * - SEARCH QUEUE sheet: queue up many searches (not just one) to
 *   run automatically every day, instead of one manual query.
 * - Daily Auto-Refinement trigger: once a day, runs the active
 *   queue AND re-checks your oldest existing leads against Place
 *   Details so rating / reviews / phone / website / open-closed
 *   status stay current instead of going stale.
 * - Leads whose business has closed are automatically flagged
 *   CLOSED (unless you've already moved them into your own
 *   pipeline stage), and lead scores fall to 0 for permanently
 *   closed businesses so dead leads stop ranking as "hot".
 * - API key can be moved out of the sheet into Script Properties
 *   (DV LEADS → Save API Key Securely) so it isn't sitting in
 *   plain text for anyone with view access to the sheet.
 * - "Setup / Update System" is now safe to re-run — it no longer
 *   wipes existing leads, raw data, run log, or settings.
 * - Retries with backoff on transient API errors, and leads are
 *   written to the sheet page-by-page instead of only at the very
 *   end, so a failure mid-search doesn't lose everything found
 *   so far.
 * - RAW DATA sheet is auto-trimmed so it doesn't grow forever.
 * - Realistic pagination cap (3 pages / 60 results per query) —
 *   this is a hard ceiling on Google's side, not a bug.
 *
 * -------------------------------------------------------------
 * SHEETS:
 * 1. SEARCH        — single manual query
 * 2. SEARCH QUEUE   — many queries, run daily
 * 3. LEADS
 * 4. RAW DATA
 * 5. RUN LOG
 * 6. SETTINGS
 *
 ***************************************************************/


const DV = {

  SHEETS: {
    SEARCH: 'SEARCH',
    QUEUE: 'SEARCH QUEUE',
    LEADS: 'LEADS',
    RAW: 'RAW DATA',
    LOG: 'RUN LOG',
    SETTINGS: 'SETTINGS'
  },

  SEARCH_ENDPOINT: 'https://places.googleapis.com/v1/places:searchText',
  DETAILS_ENDPOINT: 'https://places.googleapis.com/v1/places/',

  // Google Places Text Search hard-caps results at 3 pages of 20 (60 total)
  // per query. Requesting more pages just returns an empty nextPageToken.
  HARD_MAX_PAGES: 3,

  API_KEY_PROPERTY: 'DV_PLACES_API_KEY',
  DAILY_TRIGGER_HANDLER: 'runDailyRefinement',

  SEARCH_FIELD_MASK: [
    'places.id',
    'places.displayName',
    'places.formattedAddress',
    'places.primaryTypeDisplayName',
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
    'primaryTypeDisplayName',
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

  LEADS_HEADERS: [
    'Lead ID', 'Business Name', 'Category', 'Address', 'City / Search Location',
    'Phone', 'International Phone', 'Website', 'Rating', 'Review Count',
    'Price Level', 'Google Maps URL', 'Google Place ID', 'Latitude', 'Longitude',
    'Business Status', 'Opening Hours', 'Search Query', 'Date Discovered',
    'Lead Score', 'Status', 'Notes', 'Last Refreshed'
  ],

  LOG_HEADERS: [
    'Timestamp', 'Search Query', 'API Results', 'Passed Filters', 'New Leads',
    'Duplicates', 'Rejected Rating', 'Rejected Reviews', 'Execution Time (Sec)',
    'Status', 'Message', 'Run Type'
  ]

};


/***************************************************************
 * MENU
 ***************************************************************/

function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('DV LEADS')
    .addItem('Setup / Update System', 'setupDVLeadSystem')
    .addSeparator()
    .addItem('Run Lead Search (Manual)', 'runLeadSearch')
    .addItem('Run Search Queue Now', 'runSearchQueueNow')
    .addItem('Refresh Existing Leads Now', 'refreshExistingLeadsNow')
    .addSeparator()
    .addItem('Enable Daily Auto-Refinement', 'enableDailyAutoRefinement')
    .addItem('Disable Daily Auto-Refinement', 'disableDailyAutoRefinement')
    .addSeparator()
    .addItem('Clear Search Results', 'clearLeadResults')
    .addSeparator()
    .addItem('Save API Key Securely', 'saveApiKeySecurely')
    .addItem('Test API Connection', 'testAPIConnection')
    .addToUi();

}


/***************************************************************
 * INITIAL SETUP (safe to re-run — never wipes existing data)
 ***************************************************************/

function setupDVLeadSystem() {

  const ss = SpreadsheetApp.getActive();

  createSearchSheet_(ss);
  createQueueSheet_(ss);
  createLeadsSheet_(ss);
  createRawSheet_(ss);
  createLogSheet_(ss);
  createSettingsSheet_(ss);

  ensureHeaderColumn_(ss.getSheetByName(DV.SHEETS.LEADS), 'Last Refreshed');
  ensureHeaderColumn_(ss.getSheetByName(DV.SHEETS.LOG), 'Run Type');

  SpreadsheetApp.getUi().alert(
    'DV Sales Intelligence V2 is ready.\n\n' +
    'Next:\n' +
    '1. Open SETTINGS, paste your Google Places API key, then run\n' +
    '   DV LEADS → Save API Key Securely\n' +
    '2. Open SEARCH QUEUE and activate the searches you want to run daily\n' +
    '   (or use SEARCH for a one-off manual search)\n' +
    '3. DV LEADS → Enable Daily Auto-Refinement to keep leads fresh\n' +
    '   automatically, or run things on demand from the menu.'
  );

}


/***************************************************************
 * SEARCH SHEET (single manual query)
 ***************************************************************/

function createSearchSheet_(ss) {

  let sh = getOrCreateSheet_(ss, DV.SHEETS.SEARCH);

  if (sh.getLastRow() > 0) {
    return; // already configured — don't touch the user's values
  }

  const values = [
    ['DV SALES INTELLIGENCE', ''],
    ['GOOGLE BUSINESS DISCOVERY — MANUAL SEARCH', ''],
    ['', ''],
    ['SEARCH PARAMETER', 'VALUE'],
    ['Business / Keyword', 'Restaurants'],
    ['Location', 'Hyderabad, Telangana'],
    ['Additional Keywords', ''],
    ['Minimum Rating', 4],
    ['Minimum Reviews', 100],
    ['Maximum Results', 100],
    ['', ''],
    ['Tip', 'For recurring / daily searches, use the SEARCH QUEUE sheet instead.']
  ];

  sh.getRange(1, 1, values.length, 2).setValues(values);

  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 320);

  sh.getRange('A1:B1').merge().setFontSize(18).setFontWeight('bold');
  sh.getRange('A2:B2').merge().setFontSize(10);
  sh.getRange('A4:B4').setFontWeight('bold');

  sh.setFrozenRows(4);

}


/***************************************************************
 * SEARCH QUEUE (many queries, run automatically every day)
 ***************************************************************/

function createQueueSheet_(ss) {

  let sh = getOrCreateSheet_(ss, DV.SHEETS.QUEUE);

  if (sh.getLastRow() > 0) {
    return; // already configured — don't touch the user's queries
  }

  const headers = [
    'Active', 'Business / Keyword', 'Location', 'Additional Keywords',
    'Minimum Rating', 'Minimum Reviews', 'Maximum Results',
    'Last Run', 'New Leads (Last Run)', 'Status'
  ];

  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold');

  sh.setFrozenRows(1);

  // Seeded inactive — flip "Active" to TRUE for the searches you want to run daily.
  const seed = [
    [false, 'Restaurants', 'Hyderabad, Telangana', '', 4, 100, 100, '', '', ''],
    [false, 'Cafes', 'Jubilee Hills Hyderabad', '', 4, 50, 100, '', '', ''],
    [false, 'Hotels', 'Hyderabad, Telangana', '', 4, 100, 100, '', '', ''],
    [false, 'Interior Designers', 'Hyderabad, Telangana', '', 4, 25, 100, '', '', ''],
    [false, 'Jewellery Stores', 'Hyderabad, Telangana', '', 4, 50, 100, '', '', ''],
    [false, 'Hospitals', 'Hyderabad, Telangana', '', 4, 100, 100, '', '', '']
  ];

  sh.getRange(2, 1, seed.length, headers.length).setValues(seed);
  sh.getRange(2, 1, 200, 1).insertCheckboxes();

  sh.autoResizeColumns(1, headers.length);

}


/***************************************************************
 * LEADS SHEET
 ***************************************************************/

function createLeadsSheet_(ss) {

  let sh = getOrCreateSheet_(ss, DV.SHEETS.LEADS);

  if (sh.getLastRow() > 0) {
    return; // never wipe existing leads
  }

  sh.getRange(1, 1, 1, DV.LEADS_HEADERS.length)
    .setValues([DV.LEADS_HEADERS])
    .setFontWeight('bold');

  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, DV.LEADS_HEADERS.length).createFilter();
  sh.autoResizeColumns(1, DV.LEADS_HEADERS.length);

}


/***************************************************************
 * RAW DATA
 ***************************************************************/

function createRawSheet_(ss) {

  let sh = getOrCreateSheet_(ss, DV.SHEETS.RAW);

  if (sh.getLastRow() > 0) {
    return;
  }

  const headers = ['Timestamp', 'Search Query', 'Place ID', 'Raw JSON'];

  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold');

  sh.setFrozenRows(1);

}


/***************************************************************
 * RUN LOG
 ***************************************************************/

function createLogSheet_(ss) {

  let sh = getOrCreateSheet_(ss, DV.SHEETS.LOG);

  if (sh.getLastRow() > 0) {
    return;
  }

  sh.getRange(1, 1, 1, DV.LOG_HEADERS.length)
    .setValues([DV.LOG_HEADERS])
    .setFontWeight('bold');

  sh.setFrozenRows(1);

}


/***************************************************************
 * SETTINGS
 ***************************************************************/

function createSettingsSheet_(ss) {

  let sh = getOrCreateSheet_(ss, DV.SHEETS.SETTINGS);

  if (sh.getLastRow() === 0) {

    const values = [
      ['DV SALES INTELLIGENCE — SETTINGS', ''],
      ['', ''],
      ['SETTING', 'VALUE'],
      ['Google Places API Key', 'PASTE_API_KEY_HERE'],
      ['Default Country', 'IN'],
      ['Language', 'en'],
      ['Maximum Pages Per Search', 3],
      ['Leads To Refresh Per Day', 50],
      ['Max Daily API Calls', 150],
      ['Max Raw Data Rows', 5000],
      ['Daily Auto-Run Hour (0-23)', 6],
      ['Auto-Refresh Enabled', 'FALSE']
    ];

    sh.getRange(1, 1, values.length, 2).setValues(values);

    sh.getRange('A1:B1').merge().setFontSize(16).setFontWeight('bold');
    sh.getRange('A3:B3').setFontWeight('bold');

    sh.setColumnWidth(1, 260);
    sh.setColumnWidth(2, 400);

    return;

  }

  // Sheet already exists from V1 — add any new V2 settings without
  // touching the user's existing values (API key, etc.).
  ensureSettingRow_(sh, 'Leads To Refresh Per Day', 50);
  ensureSettingRow_(sh, 'Max Daily API Calls', 150);
  ensureSettingRow_(sh, 'Max Raw Data Rows', 5000);
  ensureSettingRow_(sh, 'Daily Auto-Run Hour (0-23)', 6);
  ensureSettingRow_(sh, 'Auto-Refresh Enabled', 'FALSE');

}


function ensureSettingRow_(sh, label, defaultValue) {

  const lastRow = sh.getLastRow();

  const labels = sh.getRange(1, 1, lastRow, 1)
    .getValues()
    .map(function(r) { return String(r[0]).trim(); });

  if (labels.indexOf(label) === -1) {
    sh.getRange(lastRow + 1, 1, 1, 2).setValues([[label, defaultValue]]);
  }

}


function ensureHeaderColumn_(sh, headerName) {

  if (!sh) return;

  const lastCol = sh.getLastColumn();

  if (lastCol === 0) return;

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];

  if (headers.indexOf(headerName) === -1) {
    sh.getRange(1, lastCol + 1).setValue(headerName).setFontWeight('bold');
  }

}


/***************************************************************
 * MANUAL SEARCH (DV LEADS → Run Lead Search)
 ***************************************************************/

function runLeadSearch() {

  const ss = SpreadsheetApp.getActive();
  const searchSheet = ss.getSheetByName(DV.SHEETS.SEARCH);
  const leadsSheet = ss.getSheetByName(DV.SHEETS.LEADS);
  const rawSheet = ss.getSheetByName(DV.SHEETS.RAW);

  if (!searchSheet || !leadsSheet) {
    SpreadsheetApp.getUi().alert('Please run "Setup / Update System" first.');
    return;
  }

  const business = String(searchSheet.getRange('B5').getValue()).trim();
  const location = String(searchSheet.getRange('B6').getValue()).trim();
  const additional = String(searchSheet.getRange('B7').getValue()).trim();
  const minRating = Number(searchSheet.getRange('B8').getValue()) || 0;
  const minReviews = Number(searchSheet.getRange('B9').getValue()) || 0;
  const maxResults = Number(searchSheet.getRange('B10').getValue()) || 100;

  if (!business || !location) {
    SpreadsheetApp.getUi().alert('Business/Keyword and Location are required.');
    return;
  }

  let query = business;
  if (additional) query += ' ' + additional;
  query += ' in ' + location;

  const apiKey = getAPIKey_();

  if (!apiKey) {
    SpreadsheetApp.getUi().alert('Add your Google Places API key in SETTINGS.');
    return;
  }

  const existingPlaceIds = getExistingPlaceIds_(leadsSheet);
  let result;

  try {
    result = performSearch_(query, location, apiKey, maxResults, minRating, minReviews,
      leadsSheet, rawSheet, existingPlaceIds, null);
  } catch (error) {
    logRun_(query, emptyStats_(), '0', 'FAILURE', String(error.message || error), 'MANUAL');
    SpreadsheetApp.getUi().alert('Search failed:\n\n' + error.message);
    return;
  }

  logRun_(query, result.stats, result.seconds, 'SUCCESS', 'Search completed', 'MANUAL');

  SpreadsheetApp.getUi().alert(
    'DV LEAD SEARCH COMPLETE\n\n' +
    'Query:\n' + query +
    '\n\nAPI Results: ' + result.stats.apiResults +
    '\nPassed Filters: ' + result.stats.passed +
    '\nNew Leads: ' + result.stats.newLeads +
    '\nDuplicates: ' + result.stats.duplicates +
    '\nRejected by Rating: ' + result.stats.rejectedRating +
    '\nRejected by Reviews: ' + result.stats.rejectedReviews +
    '\n\nExecution Time: ' + result.seconds + ' sec'
  );

}


/***************************************************************
 * SEARCH QUEUE — MANUAL TRIGGER (DV LEADS → Run Search Queue Now)
 ***************************************************************/

function runSearchQueueNow() {

  const apiKey = getAPIKey_();

  if (!apiKey) {
    SpreadsheetApp.getUi().alert('Add your Google Places API key in SETTINGS.');
    return;
  }

  const budget = { remaining: getMaxDailyApiCalls_() };
  const stats = runSearchQueue_(budget);

  SpreadsheetApp.getUi().alert(
    'SEARCH QUEUE COMPLETE\n\n' +
    'Active queries run: ' + stats.queriesRun +
    '\nNew leads found: ' + stats.newLeads
  );

}


/***************************************************************
 * LEAD REFRESH — MANUAL TRIGGER (DV LEADS → Refresh Existing Leads Now)
 ***************************************************************/

function refreshExistingLeadsNow() {

  const apiKey = getAPIKey_();

  if (!apiKey) {
    SpreadsheetApp.getUi().alert('Add your Google Places API key in SETTINGS.');
    return;
  }

  const budget = { remaining: getMaxDailyApiCalls_() };
  const stats = refreshExistingLeads_(budget);

  SpreadsheetApp.getUi().alert(
    'LEAD REFRESH COMPLETE\n\n' +
    'Checked: ' + stats.checked +
    '\nUpdated: ' + stats.updated +
    '\nNewly Closed: ' + stats.closed +
    '\nNot Found (removed from Google): ' + stats.notFound
  );

}


/***************************************************************
 * DAILY AUTO-REFINEMENT (installed trigger entry point)
 ***************************************************************/

function runDailyRefinement() {

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    return; // another run is already in progress
  }

  const startTime = new Date();

  try {

    const apiKey = getAPIKey_();

    if (!apiKey) {
      logRun_('DAILY AUTO-REFINEMENT', emptyStats_(), '0', 'FAILURE', 'No API key configured.', 'AUTO');
      return;
    }

    const budget = { remaining: getMaxDailyApiCalls_() };

    const queueStats = runSearchQueue_(budget);
    const refreshStats = refreshExistingLeads_(budget);

    trimRawData_();

    const seconds = ((new Date() - startTime) / 1000).toFixed(2);

    logDailyRun_(queueStats, refreshStats, seconds);

  } catch (error) {

    logRun_('DAILY AUTO-REFINEMENT', emptyStats_(), '0', 'FAILURE', String(error.message || error), 'AUTO');

  } finally {

    lock.releaseLock();

  }

}


function enableDailyAutoRefinement() {

  removeExistingTriggers_(DV.DAILY_TRIGGER_HANDLER);

  const hour = getDailyRunHour_();

  ScriptApp.newTrigger(DV.DAILY_TRIGGER_HANDLER)
    .timeBased()
    .atHour(hour)
    .everyDays(1)
    .create();

  setSettingValue_('Auto-Refresh Enabled', 'TRUE');

  SpreadsheetApp.getUi().alert(
    'Daily Auto-Refinement enabled.\n\n' +
    'Once a day, around ' + hour + ':00 (script timezone), this will run your ' +
    'active SEARCH QUEUE and refresh your oldest existing leads automatically.'
  );

}


function disableDailyAutoRefinement() {

  removeExistingTriggers_(DV.DAILY_TRIGGER_HANDLER);
  setSettingValue_('Auto-Refresh Enabled', 'FALSE');

  SpreadsheetApp.getUi().alert('Daily Auto-Refinement disabled.');

}


function removeExistingTriggers_(handlerName) {

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

}


/***************************************************************
 * CORE SEARCH (shared by manual search and the queue)
 ***************************************************************/

function performSearch_(query, location, apiKey, maxResults, minRating, minReviews,
    leadsSheet, rawSheet, existingPlaceIds, budget) {

  const startTime = new Date();

  const stats = emptyStats_();

  let pageToken = null;
  let page = 0;

  const maxPages = getMaxPagesPerSearch_();

  while (stats.newLeads < maxResults && page < maxPages) {

    if (budget && budget.remaining <= 0) break;

    page++;
    if (budget) budget.remaining--;

    const response = callGooglePlaces_(query, apiKey, pageToken);

    if (!response) break;

    const places = response.places || [];
    stats.apiResults += places.length;

    const leadRows = [];
    const rawRows = [];

    places.forEach(function(place) {

      rawRows.push([new Date(), query, place.id || '', JSON.stringify(place)]);

      const rating = Number(place.rating || 0);

      if (rating < minRating) {
        stats.rejectedRating++;
        return;
      }

      const reviews = Number(place.userRatingCount || 0);

      if (reviews < minReviews) {
        stats.rejectedReviews++;
        return;
      }

      stats.passed++;

      const placeId = place.id || '';

      if (placeId && existingPlaceIds.has(placeId)) {
        stats.duplicates++;
        return;
      }

      if (stats.newLeads >= maxResults) {
        return; // cap reached — leave remaining places for a future run
      }

      const name = place.displayName ? place.displayName.text : '';
      const category = place.primaryTypeDisplayName ? place.primaryTypeDisplayName.text : '';
      const address = place.formattedAddress || '';
      const phone = place.nationalPhoneNumber || '';
      const internationalPhone = place.internationalPhoneNumber || '';
      const website = place.websiteUri || '';
      const mapsUrl = place.googleMapsUri || '';
      const latitude = place.location ? place.location.latitude : '';
      const longitude = place.location ? place.location.longitude : '';
      const businessStatus = place.businessStatus || '';
      const priceLevel = place.priceLevel || '';
      const openingHours = extractOpeningHours_(place);
      const leadScore = calculateLeadScore_(rating, reviews, website, phone, businessStatus);
      const leadId = createLeadId_(placeId);
      const now = new Date();

      leadRows.push([
        leadId, name, category, address, location, phone, internationalPhone, website,
        rating, reviews, priceLevel, mapsUrl, placeId, latitude, longitude, businessStatus,
        openingHours, query, now, leadScore, 'NEW', '', now
      ]);

      if (placeId) existingPlaceIds.add(placeId);

      stats.newLeads++;

    });

    // Flush per page so a later failure doesn't lose leads already found.
    if (leadRows.length > 0) {
      const startRow = leadsSheet.getLastRow() + 1;
      leadsSheet.getRange(startRow, 1, leadRows.length, leadRows[0].length).setValues(leadRows);
    }

    if (rawRows.length > 0) {
      const rawStart = rawSheet.getLastRow() + 1;
      rawSheet.getRange(rawStart, 1, rawRows.length, rawRows[0].length).setValues(rawRows);
    }

    if (stats.newLeads >= maxResults) break;

    pageToken = response.nextPageToken || null;

    if (!pageToken) break;

    // Google's nextPageToken needs a moment to become valid.
    Utilities.sleep(1500);

  }

  const seconds = ((new Date() - startTime) / 1000).toFixed(2);

  return { stats: stats, seconds: seconds };

}


/***************************************************************
 * SEARCH QUEUE PROCESSING
 ***************************************************************/

function runSearchQueue_(budget) {

  const ss = SpreadsheetApp.getActive();
  const queueSheet = ss.getSheetByName(DV.SHEETS.QUEUE);
  const leadsSheet = ss.getSheetByName(DV.SHEETS.LEADS);
  const rawSheet = ss.getSheetByName(DV.SHEETS.RAW);
  const apiKey = getAPIKey_();

  const stats = { queriesRun: 0, newLeads: 0 };

  if (!queueSheet || !leadsSheet || !apiKey) return stats;

  const lastRow = queueSheet.getLastRow();
  if (lastRow <= 1) return stats;

  const numRows = lastRow - 1;
  const rows = queueSheet.getRange(2, 1, numRows, 10).getValues();
  const existingPlaceIds = getExistingPlaceIds_(leadsSheet);

  rows.forEach(function(row, i) {

    const sheetRow = i + 2;
    const active = row[0] === true;

    if (!active) return;

    if (budget && budget.remaining <= 0) {
      queueSheet.getRange(sheetRow, 10).setValue('SKIPPED (Daily API budget reached)');
      return;
    }

    const business = String(row[1]).trim();
    const location = String(row[2]).trim();
    const additional = String(row[3]).trim();
    const minRating = Number(row[4]) || 0;
    const minReviews = Number(row[5]) || 0;
    const maxResults = Number(row[6]) || 100;

    if (!business || !location) {
      queueSheet.getRange(sheetRow, 10).setValue('SKIPPED (missing business or location)');
      return;
    }

    let query = business;
    if (additional) query += ' ' + additional;
    query += ' in ' + location;

    try {

      const result = performSearch_(query, location, apiKey, maxResults, minRating, minReviews,
        leadsSheet, rawSheet, existingPlaceIds, budget);

      queueSheet.getRange(sheetRow, 8).setValue(new Date());
      queueSheet.getRange(sheetRow, 9).setValue(result.stats.newLeads);
      queueSheet.getRange(sheetRow, 10).setValue('OK');

      stats.queriesRun++;
      stats.newLeads += result.stats.newLeads;

      logRun_(query, result.stats, result.seconds, 'SUCCESS', 'Queue search completed', 'QUEUE');

    } catch (error) {

      queueSheet.getRange(sheetRow, 8).setValue(new Date());
      queueSheet.getRange(sheetRow, 10).setValue('ERROR: ' + error.message);

      logRun_(query, emptyStats_(), '0', 'FAILURE', String(error.message || error), 'QUEUE');

    }

  });

  return stats;

}


/***************************************************************
 * EXISTING LEAD REFRESH (Place Details lookup)
 ***************************************************************/

function refreshExistingLeads_(budget) {

  const ss = SpreadsheetApp.getActive();
  const leadsSheet = ss.getSheetByName(DV.SHEETS.LEADS);
  const apiKey = getAPIKey_();

  const stats = { checked: 0, updated: 0, closed: 0, notFound: 0 };

  if (!apiKey || !leadsSheet) return stats;

  const lastRow = leadsSheet.getLastRow();
  if (lastRow <= 1) return stats;

  const numRows = lastRow - 1;
  const numCols = DV.LEADS_HEADERS.length;
  const data = leadsSheet.getRange(2, 1, numRows, numCols).getValues();

  const candidates = data
    .map(function(row, i) {
      return {
        rowIndex: i + 2,
        placeId: row[12],
        status: row[20],
        lastRefreshed: row[22] ? new Date(row[22]).getTime() : 0
      };
    })
    .filter(function(c) { return c.placeId; })
    .sort(function(a, b) { return a.lastRefreshed - b.lastRefreshed; });

  const perDayLimit = getLeadsToRefreshPerDay_();
  const toRefresh = candidates.slice(0, perDayLimit);

  toRefresh.forEach(function(candidate) {

    if (budget && budget.remaining <= 0) return;
    if (budget) budget.remaining--;

    stats.checked++;

    try {

      const place = callPlaceDetails_(candidate.placeId, apiKey);
      updateLeadRow_(leadsSheet, candidate.rowIndex, place, candidate.status);
      stats.updated++;

      if (place.businessStatus === 'CLOSED_PERMANENTLY' || place.businessStatus === 'CLOSED_TEMPORARILY') {
        stats.closed++;
      }

    } catch (error) {

      leadsSheet.getRange(candidate.rowIndex, 16).setValue('NOT FOUND');
      leadsSheet.getRange(candidate.rowIndex, 23).setValue(new Date());
      stats.notFound++;

    }

  });

  return stats;

}


function updateLeadRow_(sheet, rowIndex, place, currentStatus) {

  const rating = Number(place.rating || 0);
  const reviews = Number(place.userRatingCount || 0);
  const phone = place.nationalPhoneNumber || '';
  const internationalPhone = place.internationalPhoneNumber || '';
  const website = place.websiteUri || '';
  const priceLevel = place.priceLevel || '';
  const businessStatus = place.businessStatus || '';
  const openingHours = extractOpeningHours_(place);
  const leadScore = calculateLeadScore_(rating, reviews, website, phone, businessStatus);

  sheet.getRange(rowIndex, 6).setValue(phone);
  sheet.getRange(rowIndex, 7).setValue(internationalPhone);
  sheet.getRange(rowIndex, 8).setValue(website);
  sheet.getRange(rowIndex, 9).setValue(rating);
  sheet.getRange(rowIndex, 10).setValue(reviews);
  sheet.getRange(rowIndex, 11).setValue(priceLevel);
  sheet.getRange(rowIndex, 16).setValue(businessStatus);
  sheet.getRange(rowIndex, 17).setValue(openingHours);
  sheet.getRange(rowIndex, 20).setValue(leadScore);
  sheet.getRange(rowIndex, 23).setValue(new Date());

  // Only auto-flip Status to CLOSED if it's still in an automated state —
  // never override a manual pipeline stage like CONTACTED / WON / LOST.
  if (businessStatus === 'CLOSED_PERMANENTLY' || businessStatus === 'CLOSED_TEMPORARILY') {
    if (currentStatus === 'NEW' || currentStatus === 'CLOSED') {
      sheet.getRange(rowIndex, 21).setValue('CLOSED');
    }
  }

}


/***************************************************************
 * RAW DATA CLEANUP
 ***************************************************************/

function trimRawData_() {

  const ss = SpreadsheetApp.getActive();
  const raw = ss.getSheetByName(DV.SHEETS.RAW);

  if (!raw) return;

  const maxRows = getMaxRawDataRows_();
  const lastRow = raw.getLastRow();
  const dataRows = lastRow - 1;

  if (dataRows > maxRows) {
    const excess = dataRows - maxRows;
    raw.deleteRows(2, excess); // oldest rows are at the top
  }

}


/***************************************************************
 * GOOGLE PLACES API (with retry/backoff on transient errors)
 ***************************************************************/

function fetchWithRetry_(url, options) {

  const maxAttempts = 3;
  let attempt = 0;

  while (true) {

    attempt++;

    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    if (code >= 200 && code < 300) return response;

    const retryable = (code === 429 || code >= 500);

    if (retryable && attempt < maxAttempts) {
      Utilities.sleep(1000 * Math.pow(2, attempt - 1));
      continue;
    }

    throw new Error('Google Places API Error ' + code + '\n\n' + response.getContentText());

  }

}


function callGooglePlaces_(query, apiKey, pageToken) {

  const payload = {
    textQuery: query,
    languageCode: 'en',
    regionCode: 'IN',
    pageSize: 20
  };

  if (pageToken) payload.pageToken = pageToken;

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'X-Goog-Api-Key': apiKey,
      'X-Goog-FieldMask': DV.SEARCH_FIELD_MASK
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = fetchWithRetry_(DV.SEARCH_ENDPOINT, options);

  return JSON.parse(response.getContentText());

}


function callPlaceDetails_(placeId, apiKey) {

  const url = DV.DETAILS_ENDPOINT + encodeURIComponent(placeId);

  const options = {
    method: 'get',
    headers: {
      'X-Goog-Api-Key': apiKey,
      'X-Goog-FieldMask': DV.DETAILS_FIELD_MASK
    },
    muteHttpExceptions: true
  };

  const response = fetchWithRetry_(url, options);

  return JSON.parse(response.getContentText());

}


/***************************************************************
 * LEAD SCORE
 ***************************************************************/

function calculateLeadScore_(rating, reviews, website, phone, businessStatus) {

  let score = 0;

  // RATING — 30 POINTS
  if (rating >= 4.5) score += 30;
  else if (rating >= 4.2) score += 25;
  else if (rating >= 4) score += 20;
  else if (rating >= 3.5) score += 10;

  // REVIEWS — 30 POINTS
  if (reviews >= 5000) score += 30;
  else if (reviews >= 1000) score += 25;
  else if (reviews >= 500) score += 20;
  else if (reviews >= 100) score += 15;
  else if (reviews >= 25) score += 5;

  // WEBSITE — 20 POINTS
  if (website) score += 20;

  // PHONE — 20 POINTS
  if (phone) score += 20;

  // A permanently closed business is not a sellable lead.
  if (businessStatus === 'CLOSED_PERMANENTLY') return 0;

  // A temporarily closed business is still worth half credit.
  if (businessStatus === 'CLOSED_TEMPORARILY') return Math.round(score * 0.5);

  return score;

}


/***************************************************************
 * OPENING HOURS
 ***************************************************************/

function extractOpeningHours_(place) {

  try {
    if (place.regularOpeningHours && place.regularOpeningHours.weekdayDescriptions) {
      return place.regularOpeningHours.weekdayDescriptions.join(' | ');
    }
  } catch (e) {
    return '';
  }

  return '';

}


/***************************************************************
 * EXISTING PLACE IDS
 ***************************************************************/

function getExistingPlaceIds_(sheet) {

  const ids = new Set();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return ids;

  const values = sheet.getRange(2, 13, lastRow - 1, 1).getValues();

  values.forEach(function(row) {
    const id = String(row[0]).trim();
    if (id) ids.add(id);
  });

  return ids;

}


/***************************************************************
 * LEAD ID
 ***************************************************************/

function createLeadId_(placeId) {

  if (placeId) {
    return 'DV-' + placeId.substring(0, 12).toUpperCase();
  }

  return 'DV-' + Utilities.getUuid().substring(0, 8).toUpperCase();

}


/***************************************************************
 * SETTINGS ACCESS
 ***************************************************************/

function getSettingsMap_() {

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DV.SHEETS.SETTINGS);
  const map = {};

  if (!sh) return map;

  const lastRow = sh.getLastRow();
  if (lastRow === 0) return map;

  const values = sh.getRange(1, 1, lastRow, 2).getValues();

  values.forEach(function(row) {
    const key = String(row[0]).trim();
    if (key) map[key] = row[1];
  });

  return map;

}


function setSettingValue_(label, value) {

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(DV.SHEETS.SETTINGS);

  if (!sh) return;

  const lastRow = sh.getLastRow();

  const labels = sh.getRange(1, 1, lastRow, 1)
    .getValues()
    .map(function(r) { return String(r[0]).trim(); });

  const idx = labels.indexOf(label);

  if (idx !== -1) {
    sh.getRange(idx + 1, 2).setValue(value);
  }

}


function getAPIKey_() {

  const stored = PropertiesService.getScriptProperties().getProperty(DV.API_KEY_PROPERTY);
  if (stored) return stored;

  const map = getSettingsMap_();
  const key = String(map['Google Places API Key'] || '').trim();

  if (!key || key === 'PASTE_API_KEY_HERE' || key.indexOf('STORED SECURELY') !== -1) {
    return null;
  }

  return key;

}


function saveApiKeySecurely() {

  const map = getSettingsMap_();
  const key = String(map['Google Places API Key'] || '').trim();

  if (!key || key === 'PASTE_API_KEY_HERE' || key.indexOf('STORED SECURELY') !== -1) {
    SpreadsheetApp.getUi().alert('Paste a real API key into SETTINGS → Google Places API Key first.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty(DV.API_KEY_PROPERTY, key);
  setSettingValue_('Google Places API Key', '•••• STORED SECURELY (Script Properties) ••••');

  SpreadsheetApp.getUi().alert(
    'API key saved securely.\n\n' +
    'It is now stored in Script Properties instead of the sheet, so it will no ' +
    'longer be visible to anyone with view access to this spreadsheet.'
  );

}


function getMaxPagesPerSearch_() {
  const map = getSettingsMap_();
  const raw = Number(map['Maximum Pages Per Search']) || DV.HARD_MAX_PAGES;
  return Math.max(1, Math.min(raw, DV.HARD_MAX_PAGES));
}


function getLeadsToRefreshPerDay_() {
  const map = getSettingsMap_();
  return Math.max(0, Number(map['Leads To Refresh Per Day']) || 50);
}


function getMaxDailyApiCalls_() {
  const map = getSettingsMap_();
  return Math.max(1, Number(map['Max Daily API Calls']) || 150);
}


function getMaxRawDataRows_() {
  const map = getSettingsMap_();
  return Math.max(100, Number(map['Max Raw Data Rows']) || 5000);
}


function getDailyRunHour_() {
  const map = getSettingsMap_();
  const hour = Number(map['Daily Auto-Run Hour (0-23)']);
  if (isNaN(hour) || hour < 0 || hour > 23) return 6;
  return hour;
}


/***************************************************************
 * LOGGING
 ***************************************************************/

function emptyStats_() {
  return { apiResults: 0, passed: 0, newLeads: 0, duplicates: 0, rejectedRating: 0, rejectedReviews: 0 };
}


function logRun_(query, stats, seconds, status, message, runType) {

  const ss = SpreadsheetApp.getActive();
  const log = ss.getSheetByName(DV.SHEETS.LOG);

  if (!log) return;

  log.appendRow([
    new Date(), query, stats.apiResults, stats.passed, stats.newLeads, stats.duplicates,
    stats.rejectedRating, stats.rejectedReviews, seconds, status, message, runType || 'MANUAL'
  ]);

}


function logDailyRun_(queueStats, refreshStats, seconds) {

  const ss = SpreadsheetApp.getActive();
  const log = ss.getSheetByName(DV.SHEETS.LOG);

  if (!log) return;

  const message =
    'New leads: ' + queueStats.newLeads +
    ' | Refreshed: ' + refreshStats.updated +
    ' | Newly Closed: ' + refreshStats.closed +
    ' | Not Found: ' + refreshStats.notFound;

  log.appendRow([
    new Date(),
    'DAILY AUTO-REFINEMENT (' + queueStats.queriesRun + ' queries, ' + refreshStats.checked + ' leads checked)',
    0, 0, queueStats.newLeads, 0, 0, 0,
    seconds, 'SUCCESS', message, 'AUTO'
  ]);

}


/***************************************************************
 * CLEAR LEADS
 ***************************************************************/

function clearLeadResults() {

  const ss = SpreadsheetApp.getActive();
  const leads = ss.getSheetByName(DV.SHEETS.LEADS);
  const raw = ss.getSheetByName(DV.SHEETS.RAW);

  if (leads && leads.getLastRow() > 1) {
    leads.getRange(2, 1, leads.getLastRow() - 1, leads.getLastColumn()).clearContent();
  }

  if (raw && raw.getLastRow() > 1) {
    raw.getRange(2, 1, raw.getLastRow() - 1, raw.getLastColumn()).clearContent();
  }

  SpreadsheetApp.getUi().alert('Lead results cleared.');

}


/***************************************************************
 * API TEST
 ***************************************************************/

function testAPIConnection() {

  const apiKey = getAPIKey_();

  if (!apiKey) {
    SpreadsheetApp.getUi().alert(
      'API key missing.\n\nAdd it in SETTINGS → Google Places API Key.'
    );
    return;
  }

  try {

    const result = callGooglePlaces_('restaurants in Hyderabad Telangana', apiKey, null);
    const count = result.places ? result.places.length : 0;

    SpreadsheetApp.getUi().alert(
      'API CONNECTION SUCCESSFUL\n\nGoogle returned ' + count + ' businesses.'
    );

  } catch (error) {

    SpreadsheetApp.getUi().alert('API CONNECTION FAILED\n\n' + error.message);

  }

}


/***************************************************************
 * HELPER
 ***************************************************************/

function getOrCreateSheet_(ss, name) {

  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  return sheet;

}
