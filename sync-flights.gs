/**
 * THE BIG 4-0 — Flights & Accommodations sync
 *
 * Pulls whatever people have entered at thebig40.vercel.app/#flights
 * into this spreadsheet, and keeps it up to date on its own.
 *
 * Setup: Extensions > Apps Script, paste this in, Save, then run
 * startAutoSync once. After that it refreshes every hour by itself.
 *
 * NOTE: this tab is overwritten on every refresh, so don't keep notes
 * of your own in it — they'd be wiped. Put those on another tab.
 */

// Where the website stores its data. Public, read-only.
const FIRESTORE_URL =
  'https://firestore.googleapis.com/v1/projects/the-big-40/databases/(default)' +
  '/documents/data/travel?key=AIzaSyAUNBnZALzSIo3bjgk6BQTge0HFEP9hRLc';

// Leave blank to auto-find the tab that has a NAME header.
// If it picks the wrong one, type the tab's exact name here, e.g. 'Flights'.
const TAB_NAME = '';

const HEADERS = [
  'NAME', 'EMAIL', 'DATE OF ARRIVAL', 'TIME OF ARRIVAL', 'AIRPORT OF ARRIVAL',
  'ADDRESS OF STAY', 'DO I NEED HELP FINDING A PLACE?', 'DATE OF DEPARTURE',
  'TIME OF DEPARTURE', 'AIRPORT OF DEPARTURE', 'DIETARY RESTRICTIONS'
];


/** Adds a "Big 4-0" menu to the spreadsheet. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Big 4-0')
    .addItem('Refresh flights now', 'syncFlights')
    .addItem('Turn on hourly auto-refresh', 'startAutoSync')
    .addSeparator()
    .addItem('Restore the lost rows', 'restoreLostRows')
    .addToUi();
}


// ── one-off restore ──────────────────────────────────────────────────
// On 12 Sept a submission from a stale browser tab overwrote the whole
// travel table, destroying eight rows. These are those rows, recovered
// from this spreadsheet's last good sync. Cecile Bugier is not listed
// because her later submission survived and is more recent.
//
// Run restoreLostRows ONCE. It merges these back in without touching
// anything that is already there. Safe to run twice — names that already
// exist are skipped.

const LOST_ROWS = [
  { name: 'Boris & Vicky de Souza', email: 'boris@bdesouza.com',
    arrDate: '2026-11-19', arrTime: '16:54', airport: 'EWR',
    stay: 'Kimpton Eventi', help: 'No',
    depDate: '2026-11-23', depTime: '18:29', depAirport: 'EWR',
    diet: 'Shellfish (Boris)' },

  { name: 'Hortense Moulonguet', email: 'hortense.moulonguet@gmail.com',
    arrDate: '2026-11-20', arrTime: '10:55', airport: 'JFK',
    stay: 'Washington Heights', help: 'No',
    depDate: '2026-11-23', depTime: '18:30', depAirport: 'JFK',
    diet: 'Nope' },

  { name: 'ZoZo \u{1F9D6}\u{1F3FC}\u{200D}\u{2640}\u{FE0F}', email: 'washbrook113@hotmail.co.uk',
    arrDate: '2026-11-20', arrTime: '11:30', airport: 'JFK',
    stay: 'The retreat minus some days with Sud', help: 'No',
    depDate: '2026-12-06', depTime: '21:36', depAirport: 'JFK',
    diet: 'None' },

  { name: 'M\u00e9gu\u00e8re Nelly', email: 'nellymeguere@gmail.com',
    arrDate: '2026-11-20', arrTime: '13:00', airport: 'JFK',
    stay: '', help: 'No',
    depDate: '2026-11-23', depTime: '17:00', depAirport: 'EWR',
    diet: '' },

  { name: 'Emilie merc', email: 'emilie.mercadal@rofim.doctor',
    arrDate: '2026-11-21', arrTime: '10:45', airport: 'JFK',
    stay: 'Your place', help: '',
    depDate: '2026-11-30', depTime: '23:00', depAirport: 'JFK',
    diet: 'Gluten' },

  { name: 'Brooke & Lauren Brown', email: 'b.schlairet@gmail.com',
    arrDate: '2026-11-21', arrTime: '12:30', airport: 'EWR',
    stay: 'The Evelyn, 7 E 27th St', help: 'No',
    depDate: '2026-11-22', depTime: '14:00', depAirport: 'JFK',
    diet: 'None' }
];


/** Puts the lost rows back on the website. Run once. */
function restoreLostRows() {
  var res = UrlFetchApp.fetch(FIRESTORE_URL, { muteHttpExceptions: true });
  var code = res.getResponseCode();
  if (code !== 200 && code !== 404) {
    throw new Error('Could not read the site data (HTTP ' + code + ').');
  }

  var existing = [];
  if (code === 200) {
    var doc = JSON.parse(res.getContentText());
    existing = (doc && doc.fields && doc.fields.list &&
                doc.fields.list.arrayValue && doc.fields.list.arrayValue.values) || [];
  }

  var haveNames = existing.map(function (v) {
    var f = (v.mapValue && v.mapValue.fields) || {};
    return String((f.name && f.name.stringValue) || '').toLowerCase();
  });

  var added = [];
  var now = Date.now();

  LOST_ROWS.forEach(function (row, i) {
    if (haveNames.indexOf(row.name.toLowerCase()) !== -1) return;
    var fields = {};
    Object.keys(row).forEach(function (k) { fields[k] = { stringValue: row[k] }; });
    fields.ts = { integerValue: String(now + i) };
    existing.push({ mapValue: { fields: fields } });
    added.push(row.name);
  });

  if (!added.length) {
    Logger.log('Nothing to restore — all rows already present.');
    return;
  }

  var write = UrlFetchApp.fetch(FIRESTORE_URL, {
    method: 'patch',
    contentType: 'application/json',
    payload: JSON.stringify({ fields: { list: { arrayValue: { values: existing } } } }),
    muteHttpExceptions: true
  });

  if (write.getResponseCode() !== 200) {
    throw new Error('Could not save (HTTP ' + write.getResponseCode() + '): ' + write.getContentText());
  }

  Logger.log('Restored ' + added.length + ' rows: ' + added.join(', '));
  syncFlights();
}


/** Run this ONCE to switch on the hourly refresh. */
function startAutoSync() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncFlights') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncFlights').timeBased().everyHours(1).create();
  syncFlights();
}


/** Pulls the website data into the flights tab. */
function syncFlights() {
  var at = findAnchor_();
  var sheet = at.sheet;
  var rows = fetchRows_();

  sheet.getRange(at.row, at.col, 1, HEADERS.length)
       .setValues([HEADERS])
       .setFontWeight('bold');

  var lastRow = sheet.getLastRow();
  if (lastRow > at.row) {
    sheet.getRange(at.row + 1, at.col, lastRow - at.row, HEADERS.length).clearContent();
  }

  if (!rows.length) return;

  var values = rows.map(function (r) {
    return [
      r.name, r.email, fmtDate_(r.arrDate), fmtTime_(r.arrTime), r.airport,
      r.stay, r.help, fmtDate_(r.depDate), fmtTime_(r.depTime), r.depAirport, r.diet
    ];
  });

  sheet.getRange(at.row + 1, at.col, values.length, HEADERS.length).setValues(values);
}


// ── helpers ──────────────────────────────────────────────────────────

/**
 * Finds the flights tab AND where its header row starts, by hunting for a
 * cell reading NAME in the top-left corner of each tab. Copes with blank
 * rows or a blank column sitting above or left of the header.
 */
function findAnchor_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var candidates = [];

  if (TAB_NAME) {
    var named = ss.getSheetByName(TAB_NAME);
    if (!named) {
      throw new Error('No tab called "' + TAB_NAME + '". Check the spelling, ' +
                      'or clear TAB_NAME at the top of the script to auto-detect it.');
    }
    candidates = [named];
  } else {
    candidates = ss.getSheets();
  }

  for (var i = 0; i < candidates.length; i++) {
    var sheet = candidates[i];
    var rows = Math.min(sheet.getMaxRows(), 20);
    var cols = Math.min(sheet.getMaxColumns(), 10);
    var grid = sheet.getRange(1, 1, rows, cols).getValues();

    for (var r = 0; r < rows; r++) {
      for (var c = 0; c < cols; c++) {
        if (String(grid[r][c]).trim().toUpperCase() === 'NAME') {
          // Skip the RSVP Data tab, whose header row starts with "#" then "Name".
          var leftOfIt = c > 0 ? String(grid[r][c - 1]).trim() : '';
          if (leftOfIt === '#') continue;
          return { sheet: sheet, row: r + 1, col: c + 1 };
        }
      }
    }
  }

  throw new Error(
    'Could not find the flights tab. It needs a cell reading NAME somewhere in ' +
    'its first 20 rows. Otherwise type the tab name into TAB_NAME at the top of this script.'
  );
}


function fetchRows_() {
  var res = UrlFetchApp.fetch(FIRESTORE_URL, { muteHttpExceptions: true });
  var code = res.getResponseCode();

  // 404 just means nobody has filled the form in yet.
  if (code === 404) return [];
  if (code !== 200) throw new Error('Could not reach the site data (HTTP ' + code + ').');

  var doc = JSON.parse(res.getContentText());
  var list = doc && doc.fields && doc.fields.list &&
             doc.fields.list.arrayValue && doc.fields.list.arrayValue.values;
  if (!list) return [];

  return list
    .map(function (v) {
      var f = (v.mapValue && v.mapValue.fields) || {};
      function s(k) { return (f[k] && f[k].stringValue) || ''; }
      return {
        name: s('name'), email: s('email'),
        arrDate: s('arrDate'), arrTime: s('arrTime'), airport: s('airport'),
        stay: s('stay'), help: s('help'),
        depDate: s('depDate'), depTime: s('depTime'), depAirport: s('depAirport'),
        diet: s('diet')
      };
    })
    .filter(function (r) { return r.name; })
    .sort(function (a, b) {
      var byDate = (a.arrDate || '9999').localeCompare(b.arrDate || '9999');
      return byDate || (a.arrTime || '').localeCompare(b.arrTime || '');
    });
}


function fmtDate_(s) {
  if (!s) return '';
  var p = s.split('-');
  if (p.length !== 3) return s;
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  if (isNaN(d.getTime())) return s;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEE d MMM yyyy');
}


function fmtTime_(s) {
  if (!s) return '';
  var p = s.split(':');
  if (p.length < 2) return s;
  var h = Number(p[0]);
  var suffix = h >= 12 ? 'pm' : 'am';
  h = h % 12 || 12;
  return h + ':' + p[1] + suffix;
}
