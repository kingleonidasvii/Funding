// ============================================================
// lovelee — Opportunity Tracker — Google Apps Script Web App
// ============================================================
//
// DEPLOYMENT INSTRUCTIONS
// -----------------------
// 1. Open the Google Sheet: https://docs.google.com/spreadsheets/d/1Si1_GF1DVPq7ZNfH5NPOxQaH6gJDtiDDbqTUYT7kQg4/edit
// 2. Extensions → Apps Script
// 3. Delete any existing code. Paste the entire contents of this file.
// 4. Click Save (Ctrl+S / Cmd+S). Name the project "lovelee-opportunity-tracker".
// 5. Click Deploy → New deployment.
// 6. Click the gear icon next to "Select type" → choose "Web app".
// 7. Set Description: "lovelee Opportunity Tracker API"
// 8. Set Execute as: "Me (your Google account)"
// 9. Set Who has access: "Anyone"  ← required for GitHub Pages to read data
// 10. Click Deploy. Authorize when prompted (review permissions → allow).
// 11. Copy the Web app URL shown (looks like: https://script.google.com/macros/s/AKfy.../exec)
// 12. Open index.html in the GitHub repo. Set SCRIPT_URL at the top to that URL.
// 13. Commit and push index.html. The app is now live.
//
// RE-DEPLOYING AFTER CODE CHANGES
// --------------------------------
// Deploy → Manage deployments → Edit (pencil icon) → Version: "New version" → Deploy
// The URL stays the same — no need to update index.html.
//
// TESTING LOCALLY
// ---------------
// Run → Run function → testGet  (reads first 3 rows and logs them)
// Run → Run function → testPost (adds a test row; delete it from the sheet afterwards)
//
// ============================================================

const SHEET_ID = '1Si1_GF1DVPq7ZNfH5NPOxQaH6gJDtiDDbqTUYT7kQg4';
const OPPORTUNITIES_TAB = 'Opportunities';

// ---- Helpers ------------------------------------------------

function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(OPPORTUNITIES_TAB);
}

function headersFromRow(row) {
  return row.map(h => String(h).trim());
}

// Currency fields where a plain number should get a € prefix
const CURRENCY_FIELDS = new Set(['Amount']);

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    const v = row[i];
    if (v instanceof Date) {
      obj[h] = Utilities.formatDate(v, 'Europe/Berlin', 'yyyy-MM-dd');
    } else if (typeof v === 'number' && CURRENCY_FIELDS.has(h)) {
      // Re-attach € if the sheet stored the value as a number
      obj[h] = '€' + v.toLocaleString('de-DE');
    } else {
      obj[h] = v !== undefined && v !== null ? String(v) : '';
    }
  });
  return obj;
}

function corsResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

function nextId(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'OPP001';
  const idCol = headers.indexOf('ID') + 1;
  if (idCol === 0) return 'OPP001';
  const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues().flat()
    .map(v => String(v).replace(/\D/g, ''))
    .filter(Boolean)
    .map(Number);
  const max = ids.length ? Math.max(...ids) : 0;
  return 'OPP' + String(max + 1).padStart(3, '0');
}

// ---- GET — return all opportunities, or run a source scan ----

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'list';

    if (action === 'scan') {
      return corsResponse(scanSources());
    }

    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return corsResponse({ opportunities: [] });

    const headers = headersFromRow(data[0]);
    const opportunities = data.slice(1).map(row => rowToObject(headers, row));
    return corsResponse({ opportunities });
  } catch (err) {
    return corsResponse({ error: err.message });
  }
}

// ---- SCAN — fetch public sources and add new entries ---------
// Checks each public source URL for keywords signalling a new
// funding round, open call, or festival submission window.
// Login-required sources (e.g. Music Group Berlin) must be
// handled manually by uploading the PDF in Claude chat.

const SCAN_SOURCES = [
  { name: 'Musicboard Berlin',          url: 'https://www.musicboard-berlin.de/foerderung/' },
  { name: 'Initiative Musik',           url: 'https://www.initiative-musik.de/foerderung/' },
  { name: 'kreativkultur.berlin',       url: 'https://kreativkultur.berlin' },
  { name: 'Berlin Music Commission',    url: 'https://www.berlin-music-commission.de' },
  { name: 'Musikfonds',                 url: 'https://www.musikfonds.de/foerderung/' },
  { name: 'GEMA Kulturförderung',       url: 'https://www.gema.de/kulturfoerderung-online' },
  { name: 'Backstage PRO',              url: 'https://www.backstage.de/ratgeber/foerderung/' },
  { name: 'German Music Export',        url: 'https://www.german-music-export.de/en/funding/' },
  { name: 'Eurosonic Noorderslag',      url: 'https://www.eurosonic-noorderslag.nl/showcases/applications/' },
  { name: 'Waves Vienna',               url: 'https://wavesvienna.com/application/' },
  { name: 'Reeperbahn Festival',        url: 'https://www.reeperbahn-festival.de/en/for-acts/' },
  { name: 'MIL Lisboa',                 url: 'https://mil-lisboa.pt' },
  { name: 'SHAPE+',                     url: 'https://shapeplatform.eu/open-call/' },
  { name: 'Creative Europe',            url: 'https://culture.ec.europa.eu/calls' },
  { name: 'Fonds Podiumkunsten',        url: 'https://www.fondspodiumkunsten.nl/subsidies/' },
];

// Keywords that suggest an active open call or funding round
const KEYWORDS = [
  'open call', 'open for', 'apply now', 'applications open',
  'jetzt bewerben', 'bewerbung', 'förderantrag', 'einreichung',
  'deadline', 'frist', 'submission', 'einreichen',
  'förderung', 'stipendium', 'grant', 'funding round'
];

function scanSources() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = headersFromRow(data[0]);
  const existing = data.slice(1).map(r => rowToObject(headers, r));
  const existingTitles = new Set(existing.map(o => o.Title.toLowerCase().trim()));

  const today = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd');
  let added = 0;

  SCAN_SOURCES.forEach(src => {
    try {
      const resp = UrlFetchApp.fetch(src.url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (compatible; lovelee-scanner/1.0)' }
      });
      if (resp.getResponseCode() !== 200) return;

      const html = resp.getContentText().toLowerCase();
      const hasKeyword = KEYWORDS.some(kw => html.includes(kw));
      if (!hasKeyword) return;

      // Extract a deadline date if visible (yyyy-mm-dd or dd.mm.yyyy pattern)
      const dateMatch = html.match(/(\d{4}-\d{2}-\d{2})|(\d{2}\.\d{2}\.\d{4})/);
      let deadline = 'rolling';
      if (dateMatch) {
        const raw = dateMatch[0];
        if (raw.includes('-')) deadline = raw;
        else {
          const [d, m, y] = raw.split('.');
          deadline = `${y}-${m}-${d}`;
        }
      }

      // Build a candidate title
      const titleCandidate = `${src.name} — Open Call (scanned ${today})`;
      if (existingTitles.has(titleCandidate.toLowerCase())) return;

      // Append to sheet
      const row = headers.map(h => {
        switch (h) {
          case 'ID':           return nextId(sheet, headers);
          case 'Title':        return titleCandidate;
          case 'Organization': return src.name;
          case 'Type':         return 'funding';
          case 'Status':       return 'watching';
          case 'Deadline':     return deadline;
          case 'Amount':       return 'variable';
          case 'Region':       return 'Germany';
          case 'Source':       return 'Auto-scan';
          case 'URL':          return src.url;
          case 'Notes':        return 'Auto-detected via keyword scan. Review and update details.';
          case 'Date Added':   return today;
          case 'Date Updated': return today;
          case 'Applied':      return 'FALSE';
          default:             return '';
        }
      });
      sheet.appendRow(row);
      existingTitles.add(titleCandidate.toLowerCase());
      added++;
    } catch (err) {
      Logger.log('Scan error for ' + src.name + ': ' + err.message);
    }
  });

  return { added, scanned: SCAN_SOURCES.length, timestamp: today };
}

// ---- Set up monthly trigger (run this ONCE manually) --------
// In the Apps Script editor: Run → Run function → setupMonthlyTrigger
// This creates a trigger that scans on the 1st of every month at 9am Berlin.
function setupMonthlyTrigger() {
  // Remove any existing monthly triggers first
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runMonthlyScan')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('runMonthlyScan')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();
  Logger.log('Monthly trigger created: 1st of each month at 9am.');
}

function runMonthlyScan() {
  const result = scanSources();
  Logger.log('Monthly scan complete: ' + JSON.stringify(result));
}

// ---- POST — add a new opportunity row -----------------------

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = getSheet();
    const headers = headersFromRow(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);

    const today = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd');
    const id = nextId(sheet, headers);

    const row = headers.map(h => {
      switch (h) {
        case 'ID':           return id;
        case 'Date Added':   return today;
        case 'Date Updated': return today;
        case 'Applied':      return payload['Applied'] || 'FALSE';
        case 'Result':       return payload['Result'] || '';
        default:             return payload[h] || '';
      }
    });

    sheet.appendRow(row);
    return corsResponse({ success: true, id });
  } catch (err) {
    return corsResponse({ error: err.message });
  }
}

// ---- Test functions (run manually from Apps Script editor) --

function testGet() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = headersFromRow(data[0]);
  const rows = data.slice(1, 4).map(r => rowToObject(headers, r));
  Logger.log(JSON.stringify(rows, null, 2));
}

function testPost() {
  const fakePost = {
    postData: {
      contents: JSON.stringify({
        Title: 'TEST ENTRY — delete me',
        Organization: 'Test Org',
        Type: 'funding',
        Status: 'watching',
        Deadline: '2026-12-31',
        Amount: '€0',
        Region: 'Germany',
        Eligibility: 'test',
        URL: 'https://example.com',
        Notes: 'Auto-generated by testPost(). Delete this row.',
        Source: 'test'
      })
    }
  };
  const result = doPost(fakePost);
  Logger.log(result.getContent());
}
