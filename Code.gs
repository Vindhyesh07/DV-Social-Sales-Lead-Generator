/***************************************************************
 * DV SOCIAL — SALES INTELLIGENCE ENGINE
 * VERSION 1.0
 *
 * Stack:
 * Google Sheets
 * Google Apps Script
 * Google Places API (New)
 *
 * -------------------------------------------------------------
 * SHEETS CREATED:
 * 1. SEARCH
 * 2. LEADS
 * 3. RAW DATA
 * 4. RUN LOG
 * 5. SETTINGS
 *
 ***************************************************************/


const DV = {

  SHEETS: {
    SEARCH: 'SEARCH',
    LEADS: 'LEADS',
    RAW: 'RAW DATA',
    LOG: 'RUN LOG',
    SETTINGS: 'SETTINGS'
  },

  SEARCH_ENDPOINT:
    'https://places.googleapis.com/v1/places:searchText',

  MAX_PAGES_PER_QUERY: 5,

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
  ].join(',')

};


/***************************************************************
 * MENU
 ***************************************************************/

function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('DV LEADS')
    .addItem('Setup V1 System', 'setupDVLeadSystem')
    .addSeparator()
    .addItem('Run Lead Search', 'runLeadSearch')
    .addItem('Clear Search Results', 'clearLeadResults')
    .addSeparator()
    .addItem('Test API Connection', 'testAPIConnection')
    .addToUi();

}


/***************************************************************
 * INITIAL SETUP
 ***************************************************************/

function setupDVLeadSystem() {

  const ss = SpreadsheetApp.getActive();

  createSearchSheet_(ss);
  createLeadsSheet_(ss);
  createRawSheet_(ss);
  createLogSheet_(ss);
  createSettingsSheet_(ss);

  SpreadsheetApp.getUi().alert(
    'DV Sales Intelligence V1 created.\n\n' +
    'Next:\n' +
    '1. Open SETTINGS\n' +
    '2. Paste your Google Places API key\n' +
    '3. Open SEARCH\n' +
    '4. Enter your search parameters\n' +
    '5. Use DV LEADS → Run Lead Search'
  );

}


/***************************************************************
 * SEARCH SHEET
 ***************************************************************/

function createSearchSheet_(ss) {

  let sh = getOrCreateSheet_(ss, DV.SHEETS.SEARCH);

  sh.clear();

  const values = [

    ['DV SALES INTELLIGENCE', ''],
    ['GOOGLE BUSINESS DISCOVERY — V1', ''],

    ['', ''],

    ['SEARCH PARAMETER', 'VALUE'],

    ['Business / Keyword', 'Restaurants'],
    ['Location', 'Hyderabad, Telangana'],
    ['Additional Keywords', ''],
    ['Minimum Rating', 4],
    ['Minimum Reviews', 100],
    ['Maximum Results', 100],

    ['', ''],

    ['Example Searches', ''],
    ['Restaurants', 'Hyderabad'],
    ['Cafes', 'Jubilee Hills Hyderabad'],
    ['Hotels', 'Hyderabad'],
    ['Interior Designers', 'Hyderabad'],
    ['Jewellery Stores', 'Hyderabad'],
    ['Hospitals', 'Hyderabad']

  ];

  sh.getRange(1, 1, values.length, 2).setValues(values);

  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 320);

  sh.getRange('A1:B1')
    .merge()
    .setFontSize(18)
    .setFontWeight('bold');

  sh.getRange('A2:B2')
    .merge()
    .setFontSize(10);

  sh.getRange('A4:B4')
    .setFontWeight('bold');

  sh.setFrozenRows(4);

}


/***************************************************************
 * LEADS SHEET
 ***************************************************************/

function createLeadsSheet_(ss) {

  let sh = getOrCreateSheet_(ss, DV.SHEETS.LEADS);

  sh.clear();

  const headers = [

    'Lead ID',
    'Business Name',
    'Category',

    'Address',
    'City / Search Location',

    'Phone',
    'International Phone',

    'Website',

    'Rating',
    'Review Count',

    'Price Level',

    'Google Maps URL',
    'Google Place ID',

    'Latitude',
    'Longitude',

    'Business Status',

    'Opening Hours',

    'Search Query',

    'Date Discovered',

    'Lead Score',

    'Status',

    'Notes'

  ];

  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold');

  sh.setFrozenRows(1);

  sh.getRange(1, 1, 1, headers.length)
    .createFilter();

  sh.autoResizeColumns(1, headers.length);

}


/***************************************************************
 * RAW DATA
 ***************************************************************/

function createRawSheet_(ss) {

  let sh = getOrCreateSheet_(ss, DV.SHEETS.RAW);

  sh.clear();

  const headers = [
    'Timestamp',
    'Search Query',
    'Place ID',
    'Raw JSON'
  ];

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

  sh.clear();

  const headers = [

    'Timestamp',
    'Search Query',

    'API Results',
    'Passed Filters',

    'New Leads',
    'Duplicates',

    'Rejected Rating',
    'Rejected Reviews',

    'Execution Time (Sec)',

    'Status',
    'Message'

  ];

  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold');

  sh.setFrozenRows(1);

}


/***************************************************************
 * SETTINGS
 ***************************************************************/

function createSettingsSheet_(ss) {

  let sh = getOrCreateSheet_(ss, DV.SHEETS.SETTINGS);

  sh.clear();

  const values = [

    ['DV SALES INTELLIGENCE — SETTINGS', ''],

    ['', ''],

    ['SETTING', 'VALUE'],

    ['Google Places API Key', 'PASTE_API_KEY_HERE'],

    ['Default Country', 'IN'],

    ['Language', 'en'],

    ['Maximum Pages Per Search', 5]

  ];

  sh.getRange(1, 1, values.length, 2)
    .setValues(values);

  sh.getRange('A1:B1')
    .merge()
    .setFontSize(16)
    .setFontWeight('bold');

  sh.getRange('A3:B3')
    .setFontWeight('bold');

  sh.setColumnWidth(1, 250);
  sh.setColumnWidth(2, 400);

}


/***************************************************************
 * MAIN SEARCH
 ***************************************************************/

function runLeadSearch() {

  const startTime = new Date();

  const ss = SpreadsheetApp.getActive();

  const searchSheet =
    ss.getSheetByName(DV.SHEETS.SEARCH);

  const leadsSheet =
    ss.getSheetByName(DV.SHEETS.LEADS);

  const rawSheet =
    ss.getSheetByName(DV.SHEETS.RAW);

  if (!searchSheet || !leadsSheet) {

    SpreadsheetApp.getUi().alert(
      'Please run "Setup V1 System" first.'
    );

    return;

  }


  /*************************************************************
   * READ SEARCH PARAMETERS
   *************************************************************/

  const business =
    String(searchSheet.getRange('B5').getValue()).trim();

  const location =
    String(searchSheet.getRange('B6').getValue()).trim();

  const additional =
    String(searchSheet.getRange('B7').getValue()).trim();

  const minRating =
    Number(searchSheet.getRange('B8').getValue()) || 0;

  const minReviews =
    Number(searchSheet.getRange('B9').getValue()) || 0;

  const maxResults =
    Number(searchSheet.getRange('B10').getValue()) || 100;


  if (!business || !location) {

    SpreadsheetApp.getUi().alert(
      'Business/Keyword and Location are required.'
    );

    return;

  }


  /*************************************************************
   * BUILD QUERY
   *************************************************************/

  let query = business;

  if (additional) {
    query += ' ' + additional;
  }

  query += ' in ' + location;


  /*************************************************************
   * GET API KEY
   *************************************************************/

  const apiKey = getAPIKey_();


  if (!apiKey) {

    SpreadsheetApp.getUi().alert(
      'Add your Google Places API key in SETTINGS.'
    );

    return;

  }


  /*************************************************************
   * EXISTING PLACE IDS
   *************************************************************/

  const existingPlaceIds =
    getExistingPlaceIds_(leadsSheet);


  /*************************************************************
   * STATS
   *************************************************************/

  let stats = {

    apiResults: 0,

    passed: 0,

    newLeads: 0,

    duplicates: 0,

    rejectedRating: 0,

    rejectedReviews: 0

  };


  let rows = [];

  let rawRows = [];

  let pageToken = null;

  let page = 0;


  /*************************************************************
   * PAGINATION LOOP
   *************************************************************/

  while (
    rows.length < maxResults &&
    page < DV.MAX_PAGES_PER_QUERY
  ) {

    page++;


    const response =
      callGooglePlaces_(
        query,
        apiKey,
        pageToken
      );


    if (!response) {
      break;
    }


    const places =
      response.places || [];


    stats.apiResults += places.length;


    /***********************************************************
     * PROCESS PLACES
     ***********************************************************/

    places.forEach(function(place) {


      /*********************************************************
       * SAVE RAW RESPONSE
       *********************************************************/

      rawRows.push([

        new Date(),

        query,

        place.id || '',

        JSON.stringify(place)

      ]);


      /*********************************************************
       * FILTER RATING
       *********************************************************/

      const rating =
        Number(place.rating || 0);


      if (rating < minRating) {

        stats.rejectedRating++;

        return;

      }


      /*********************************************************
       * FILTER REVIEWS
       *********************************************************/

      const reviews =
        Number(place.userRatingCount || 0);


      if (reviews < minReviews) {

        stats.rejectedReviews++;

        return;

      }


      stats.passed++;


      /*********************************************************
       * DEDUPLICATION
       *********************************************************/

      const placeId =
        place.id || '';


      if (
        placeId &&
        existingPlaceIds.has(placeId)
      ) {

        stats.duplicates++;

        return;

      }


      /*********************************************************
       * BUSINESS DATA
       *********************************************************/

      const name =
        place.displayName
          ? place.displayName.text
          : '';


      const category =
        place.primaryTypeDisplayName
          ? place.primaryTypeDisplayName.text
          : '';


      const address =
        place.formattedAddress || '';


      const phone =
        place.nationalPhoneNumber || '';


      const internationalPhone =
        place.internationalPhoneNumber || '';


      const website =
        place.websiteUri || '';


      const mapsUrl =
        place.googleMapsUri || '';


      const latitude =
        place.location
          ? place.location.latitude
          : '';


      const longitude =
        place.location
          ? place.location.longitude
          : '';


      const businessStatus =
        place.businessStatus || '';


      const priceLevel =
        place.priceLevel || '';


      const openingHours =
        extractOpeningHours_(place);


      /*********************************************************
       * LEAD SCORE
       *********************************************************/

      const leadScore =
        calculateLeadScore_(
          rating,
          reviews,
          website,
          phone
        );


      /*********************************************************
       * UNIQUE LEAD ID
       *********************************************************/

      const leadId =
        createLeadId_(placeId);


      /*********************************************************
       * CREATE ROW
       *********************************************************/

      rows.push([

        leadId,

        name,

        category,

        address,

        location,

        phone,

        internationalPhone,

        website,

        rating,

        reviews,

        priceLevel,

        mapsUrl,

        placeId,

        latitude,

        longitude,

        businessStatus,

        openingHours,

        query,

        new Date(),

        leadScore,

        'NEW',

        ''

      ]);


      if (placeId) {
        existingPlaceIds.add(placeId);
      }


      stats.newLeads++;

    });


    /***********************************************************
     * STOP IF LIMIT REACHED
     ***********************************************************/

    if (rows.length >= maxResults) {
      break;
    }


    /***********************************************************
     * NEXT PAGE
     ***********************************************************/

    pageToken =
      response.nextPageToken || null;


    if (!pageToken) {
      break;
    }


    Utilities.sleep(1000);

  }


  /*************************************************************
   * LIMIT RESULTS
   *************************************************************/

  if (rows.length > maxResults) {

    rows =
      rows.slice(
        0,
        maxResults
      );

  }


  /*************************************************************
   * WRITE LEADS
   *************************************************************/

  if (rows.length > 0) {

    const startRow =
      leadsSheet.getLastRow() + 1;


    leadsSheet
      .getRange(
        startRow,
        1,
        rows.length,
        rows[0].length
      )
      .setValues(rows);

  }


  /*************************************************************
   * WRITE RAW DATA
   *************************************************************/

  if (rawRows.length > 0) {

    const rawStart =
      rawSheet.getLastRow() + 1;


    rawSheet
      .getRange(
        rawStart,
        1,
        rawRows.length,
        rawRows[0].length
      )
      .setValues(rawRows);

  }


  /*************************************************************
   * LOG
   *************************************************************/

  const executionSeconds =
    (
      (new Date() - startTime)
      / 1000
    ).toFixed(2);


  logRun_(
    query,
    stats,
    executionSeconds,
    'SUCCESS',
    'Search completed'
  );


  /*************************************************************
   * RESULT MESSAGE
   *************************************************************/

  SpreadsheetApp.getUi().alert(

    'DV LEAD SEARCH COMPLETE\n\n' +

    'Query:\n' +
    query +

    '\n\nAPI Results: ' +
    stats.apiResults +

    '\nPassed Filters: ' +
    stats.passed +

    '\nNew Leads: ' +
    stats.newLeads +

    '\nDuplicates: ' +
    stats.duplicates +

    '\nRejected by Rating: ' +
    stats.rejectedRating +

    '\nRejected by Reviews: ' +
    stats.rejectedReviews +

    '\n\nExecution Time: ' +
    executionSeconds +
    ' sec'

  );

}


/***************************************************************
 * GOOGLE PLACES API
 ***************************************************************/

function callGooglePlaces_(
  query,
  apiKey,
  pageToken
) {

  const payload = {

    textQuery: query,

    languageCode: 'en',

    regionCode: 'IN',

    pageSize: 20

  };


  if (pageToken) {

    payload.pageToken =
      pageToken;

  }


  const options = {

    method: 'post',

    contentType: 'application/json',

    headers: {

      'X-Goog-Api-Key':
        apiKey,

      'X-Goog-FieldMask':
        DV.SEARCH_FIELD_MASK

    },

    payload:
      JSON.stringify(payload),

    muteHttpExceptions: true

  };


  const response =
    UrlFetchApp.fetch(
      DV.SEARCH_ENDPOINT,
      options
    );


  const code =
    response.getResponseCode();


  const text =
    response.getContentText();


  if (
    code < 200 ||
    code >= 300
  ) {

    throw new Error(
      'Google Places API Error ' +
      code +
      '\n\n' +
      text
    );

  }


  return JSON.parse(text);

}


/***************************************************************
 * LEAD SCORE
 ***************************************************************/

function calculateLeadScore_(
  rating,
  reviews,
  website,
  phone
) {

  let score = 0;


  /*************************************************************
   * RATING — 30 POINTS
   *************************************************************/

  if (rating >= 4.5) {

    score += 30;

  } else if (rating >= 4.2) {

    score += 25;

  } else if (rating >= 4) {

    score += 20;

  } else if (rating >= 3.5) {

    score += 10;

  }


  /*************************************************************
   * REVIEWS — 30 POINTS
   *************************************************************/

  if (reviews >= 5000) {

    score += 30;

  } else if (reviews >= 1000) {

    score += 25;

  } else if (reviews >= 500) {

    score += 20;

  } else if (reviews >= 100) {

    score += 15;

  } else if (reviews >= 25) {

    score += 5;

  }


  /*************************************************************
   * WEBSITE — 20 POINTS
   *************************************************************/

  if (website) {

    score += 20;

  }


  /*************************************************************
   * PHONE — 20 POINTS
   *************************************************************/

  if (phone) {

    score += 20;

  }


  return score;

}


/***************************************************************
 * OPENING HOURS
 ***************************************************************/

function extractOpeningHours_(place) {

  try {

    if (
      place.regularOpeningHours &&
      place.regularOpeningHours.weekdayDescriptions
    ) {

      return place
        .regularOpeningHours
        .weekdayDescriptions
        .join(' | ');

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

  const ids =
    new Set();


  const lastRow =
    sheet.getLastRow();


  if (lastRow <= 1) {
    return ids;
  }


  const values =
    sheet
      .getRange(
        2,
        13,
        lastRow - 1,
        1
      )
      .getValues();


  values.forEach(function(row) {

    const id =
      String(row[0]).trim();


    if (id) {

      ids.add(id);

    }

  });


  return ids;

}


/***************************************************************
 * LEAD ID
 ***************************************************************/

function createLeadId_(placeId) {

  if (placeId) {

    return (
      'DV-' +
      placeId
        .substring(0, 12)
        .toUpperCase()
    );

  }


  return (
    'DV-' +
    Utilities
      .getUuid()
      .substring(0, 8)
      .toUpperCase()
  );

}


/***************************************************************
 * API KEY
 ***************************************************************/

function getAPIKey_() {

  const ss =
    SpreadsheetApp.getActive();


  const settings =
    ss.getSheetByName(
      DV.SHEETS.SETTINGS
    );


  if (!settings) {
    return null;
  }


  const key =
    String(
      settings
        .getRange('B4')
        .getValue()
    ).trim();


  if (
    !key ||
    key === 'PASTE_API_KEY_HERE'
  ) {

    return null;

  }


  return key;

}


/***************************************************************
 * LOG RUN
 ***************************************************************/

function logRun_(
  query,
  stats,
  seconds,
  status,
  message
) {

  const ss =
    SpreadsheetApp.getActive();


  const log =
    ss.getSheetByName(
      DV.SHEETS.LOG
    );


  if (!log) {
    return;
  }


  log.appendRow([

    new Date(),

    query,

    stats.apiResults,

    stats.passed,

    stats.newLeads,

    stats.duplicates,

    stats.rejectedRating,

    stats.rejectedReviews,

    seconds,

    status,

    message

  ]);

}


/***************************************************************
 * CLEAR LEADS
 ***************************************************************/

function clearLeadResults() {

  const ss =
    SpreadsheetApp.getActive();


  const leads =
    ss.getSheetByName(
      DV.SHEETS.LEADS
    );


  const raw =
    ss.getSheetByName(
      DV.SHEETS.RAW
    );


  if (
    leads &&
    leads.getLastRow() > 1
  ) {

    leads
      .getRange(
        2,
        1,
        leads.getLastRow() - 1,
        leads.getLastColumn()
      )
      .clearContent();

  }


  if (
    raw &&
    raw.getLastRow() > 1
  ) {

    raw
      .getRange(
        2,
        1,
        raw.getLastRow() - 1,
        raw.getLastColumn()
      )
      .clearContent();

  }


  SpreadsheetApp
    .getUi()
    .alert(
      'Lead results cleared.'
    );

}


/***************************************************************
 * API TEST
 ***************************************************************/

function testAPIConnection() {

  const apiKey =
    getAPIKey_();


  if (!apiKey) {

    SpreadsheetApp.getUi().alert(
      'API key missing.\n\n' +
      'Add it in SETTINGS → Google Places API Key.'
    );

    return;

  }


  try {

    const result =
      callGooglePlaces_(
        'restaurants in Hyderabad Telangana',
        apiKey,
        null
      );


    const count =
      result.places
        ? result.places.length
        : 0;


    SpreadsheetApp.getUi().alert(

      'API CONNECTION SUCCESSFUL\n\n' +

      'Google returned ' +
      count +
      ' businesses.'

    );


  } catch (error) {

    SpreadsheetApp.getUi().alert(

      'API CONNECTION FAILED\n\n' +

      error.message

    );

  }

}


/***************************************************************
 * HELPER
 ***************************************************************/

function getOrCreateSheet_(
  ss,
  name
) {

  let sheet =
    ss.getSheetByName(name);


  if (!sheet) {

    sheet =
      ss.insertSheet(name);

  }


  return sheet;

}
