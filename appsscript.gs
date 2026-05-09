// ============================================================
// lovelee — Opportunity Tracker — Google Apps Script Web App
// ============================================================
//
// DEPLOYMENT INSTRUCTIONS
// -----------------------
// 1. Open the Google Sheet: https://docs.google.com/spreadsheets/d/1Si1_GF1DVPq7ZNfH5NPOxQaH6gJDtiDDbqTUYT7kQg4/edit
// 2. Extensions → Apps Script
// 3. Delete any existing code. Paste the entire contents of this file.
// 4. Click Save (Cmd+S). Name the project "lovelee-opportunity-tracker".
// 5. Click Deploy → New deployment → gear icon → Web app.
// 6. Execute as: Me | Who has access: Anyone → Deploy → Authorize → Allow.
// 7. Copy the Web app URL (https://script.google.com/macros/s/AKfy.../exec)
// 8. Paste it into SCRIPT_URL in index.html on GitHub. Commit.
//
// AFTER CODE CHANGES: Deploy → Manage deployments → Edit → New version → Deploy
// URL stays the same — no need to update index.html again.
//
// ONE-TIME SETUP (run these manually in the editor after first deploy):
//   Run → setupPriorityColumn   — adds Priority column to the Sheet
//   Run → setupMonthlyTrigger   — sets up automatic scan on 1st of month
//
// SHEET COLUMNS (Opportunities tab):
//   ID | Title | Organization | Type | Status | Deadline | Amount |
//   Region | Eligibility | URL | Notes | Source | Date Added |
//   Date Updated | Applied | Result | Priority
// ============================================================

const SHEET_ID        = '1Si1_GF1DVPq7ZNfH5NPOxQaH6gJDtiDDbqTUYT7kQg4';
const OPPORTUNITIES_TAB = 'Opportunities';

// ── Helpers ──────────────────────────────────────────────────

function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(OPPORTUNITIES_TAB);
}

function headersFromRow(row) {
  return row.map(h => String(h).trim());
}

const CURRENCY_FIELDS = new Set(['Amount']);

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    const v = row[i];
    if (v instanceof Date) {
      obj[h] = Utilities.formatDate(v, 'Europe/Berlin', 'yyyy-MM-dd');
    } else if (typeof v === 'number' && CURRENCY_FIELDS.has(h)) {
      obj[h] = '€' + v.toLocaleString('de-DE');
    } else {
      obj[h] = (v !== undefined && v !== null) ? String(v) : '';
    }
  });
  return obj;
}

function corsResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function nextId(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'OPP001';
  const idCol = headers.indexOf('ID') + 1;
  if (idCol === 0) return 'OPP001';
  const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues().flat()
    .map(v => String(v).replace(/\D/g, '')).filter(Boolean).map(Number);
  const max = ids.length ? Math.max(...ids) : 0;
  return 'OPP' + String(max + 1).padStart(3, '0');
}

// ── GET — list all opportunities, run scan, or mark applied ──

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'list';

    if (action === 'scan')         return corsResponse(scanSources());
    if (action === 'markApplied')  return corsResponse(markApplied(e.parameter.id));

    const sheet = getSheet();
    const data  = sheet.getDataRange().getValues();
    if (data.length < 2) return corsResponse({ opportunities: [] });

    const headers = headersFromRow(data[0]);
    const opportunities = data.slice(1).map(row => rowToObject(headers, row));
    return corsResponse({ opportunities });
  } catch (err) {
    return corsResponse({ error: err.message });
  }
}

// ── POST — add a new opportunity ─────────────────────────────

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet   = getSheet();
    const headers = headersFromRow(
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    );
    const today = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd');
    const id    = nextId(sheet, headers);

    const row = headers.map(h => {
      switch (h) {
        case 'ID':           return id;
        case 'Date Added':   return today;
        case 'Date Updated': return today;
        case 'Applied':      return payload['Applied']  || 'FALSE';
        case 'Result':       return payload['Result']   || '';
        case 'Priority':     return payload['Priority'] || 'Normal';
        default:             return payload[h]          || '';
      }
    });

    sheet.appendRow(row);
    return corsResponse({ success: true, id });
  } catch (err) {
    return corsResponse({ error: err.message });
  }
}

// ── Mark Applied (called from app's "Mark Applied" button) ───

function markApplied(id) {
  if (!id) return { error: 'No ID provided' };
  const sheet   = getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = headersFromRow(data[0]);
  const idCol   = headers.indexOf('ID');
  const applCol = headers.indexOf('Applied');
  const statCol = headers.indexOf('Status');
  const updCol  = headers.indexOf('Date Updated');
  if (idCol < 0 || applCol < 0) return { error: 'Column not found' };

  const today = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd');
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idCol]).trim() === String(id).trim()) {
      sheet.getRange(r + 1, applCol + 1).setValue('TRUE');
      if (statCol >= 0) sheet.getRange(r + 1, statCol + 1).setValue('submitted');
      if (updCol  >= 0) sheet.getRange(r + 1, updCol  + 1).setValue(today);
      return { success: true, id, status: 'submitted' };
    }
  }
  return { error: 'ID not found: ' + id };
}

// ── One-time setup: add Priority column if missing ───────────

function setupPriorityColumn() {
  const sheet   = getSheet();
  const headers = headersFromRow(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
  if (headers.includes('Priority')) {
    Logger.log('Priority column already exists.');
    return;
  }
  const col = sheet.getLastColumn() + 1;
  sheet.getRange(1, col).setValue('Priority');

  // Pre-fill known high-priority entries
  const HIGH = new Set([
    'International Export Funding', 'Eurosonic Noorderslag 2027',
    'Reeperbahn Festival 2026', 'WOMEX 2026 Delegate',
    'International Live and Showcase Funding'
  ]);
  const data  = sheet.getDataRange().getValues();
  const titleCol = headers.indexOf('Title');
  for (let r = 1; r < data.length; r++) {
    const title = String(data[r][titleCol] || '');
    const prio  = HIGH.has(title) ? 'High' : 'Normal';
    sheet.getRange(r + 1, col).setValue(prio);
  }
  Logger.log('Priority column added and pre-filled.');
}

// ── Scan sources ─────────────────────────────────────────────
// Fetches public source pages, detects active open calls by keyword,
// and appends new "watching" rows. Login-required sources (e.g.
// Music Group Berlin) should be handled by uploading their PDF
// in Claude chat — Claude will extract and add entries via POST.

const SCAN_SOURCES = [
  // ── Germany: Funding ─────────────────────────────────────
  { name:'Musicboard Berlin',           url:'https://www.musicboard-berlin.de/foerderung/',                    priority:'High',   type:'funding'  },
  { name:'Initiative Musik',            url:'https://www.initiative-musik.de/foerderung/',                    priority:'High',   type:'funding'  },
  { name:'kreativkultur.berlin',        url:'https://kreativkultur.berlin',                                   priority:'Normal', type:'funding'  },
  { name:'Berlin Music Commission',     url:'https://www.berlin-music-commission.de',                         priority:'Normal', type:'funding'  },
  { name:'Musikfonds',                  url:'https://www.musikfonds.de/foerderung/',                          priority:'Normal', type:'funding'  },
  { name:'GEMA Kulturförderung',        url:'https://www.gema.de/kulturfoerderung-online',                    priority:'Normal', type:'funding'  },
  { name:'German Music Export',         url:'https://www.german-music-export.de/en/funding/',                 priority:'High',   type:'funding'  },
  { name:'Backstage PRO',               url:'https://www.backstage.de/ratgeber/foerderung/',                  priority:'Normal', type:'funding'  },
  { name:'Fonds Darstellende Künste',   url:'https://www.fonds-daku.de/foerderung/',                          priority:'Normal', type:'funding'  },
  { name:'NRW KULTURsekretariat',       url:'https://www.nrw-kultursekretariat.de/musik/',                    priority:'Normal', type:'funding'  },

  // ── Germany: Festivals & Major Stages ────────────────────
  { name:'Reeperbahn Festival',         url:'https://www.reeperbahn-festival.de/en/for-acts/',                priority:'High',   type:'festival' },
  { name:'c/o pop Convention',          url:'https://copop.de/convention/',                                   priority:'High',   type:'festival' },
  { name:'SWR3 New Pop Festival',       url:'https://www.swr.de/swr3/musik/new-pop-festival/',                priority:'High',   type:'festival' },
  { name:'Lollapalooza Berlin',         url:'https://www.lollapaloozade.com/',                                priority:'High',   type:'festival' },
  { name:'Melt Festival',               url:'https://www.meltfestival.de',                                    priority:'High',   type:'festival' },
  { name:'MS Dockville',                url:'https://msdockville.de',                                         priority:'Normal', type:'festival' },
  { name:'Hurricane / Southside',       url:'https://www.hurricane.de',                                       priority:'High',   type:'festival' },
  { name:'Immergut Festival',           url:'https://www.immergutrocken.de',                                  priority:'Normal', type:'festival' },

  // ── Germany: Radio & Media ───────────────────────────────
  { name:'Fritz (rbb)',                 url:'https://www.fritz.de/programm/aktionen/',                        priority:'High',   type:'radio'    },
  { name:'ByteFM',                      url:'https://www.byte.fm/features/',                                  priority:'High',   type:'radio'    },
  { name:'radioeins (rbb)',             url:'https://www.radioeins.de/musik/',                                 priority:'Normal', type:'radio'    },
  { name:'Deutschlandfunk Kultur',      url:'https://www.deutschlandfunkkultur.de/musik/',                    priority:'Normal', type:'radio'    },
  { name:'MDR Kultur',                  url:'https://www.mdr.de/mdr-kultur/radio/index.html',                 priority:'Normal', type:'radio'    },

  // ── EU: Funding ──────────────────────────────────────────
  { name:'Creative Europe',             url:'https://culture.ec.europa.eu/calls',                            priority:'High',   type:'funding'  },
  { name:'SHAPE+',                      url:'https://shapeplatform.eu/open-call/',                            priority:'High',   type:'open-call'},
  { name:'Liveurope',                   url:'https://liveurope.eu',                                           priority:'High',   type:'funding'  },
  { name:'Fonds Podiumkunsten',         url:'https://www.fondspodiumkunsten.nl/subsidies/',                   priority:'Normal', type:'funding'  },
  { name:'Österreichischer Musikfonds', url:'https://www.musikfonds.at/foerderungen/',                        priority:'Normal', type:'funding'  },
  { name:'Pro Helvetia',                url:'https://prohelvetia.ch/en/funding/',                             priority:'Normal', type:'funding'  },
  { name:'Music Finland Export',        url:'https://musicfinland.com/en/funding/',                           priority:'Normal', type:'funding'  },
  { name:'Music Norway',                url:'https://musicnorway.no/funding/',                                priority:'Normal', type:'funding'  },
  { name:'Swedish Arts Council',        url:'https://www.kulturradet.se/en/apply-for-grants/',                priority:'Normal', type:'funding'  },
  { name:'PRS Foundation (UK)',         url:'https://prsfoundation.com/funding-support/funding/',             priority:'High',   type:'funding'  },
  { name:'Help Musicians UK',           url:'https://www.helpmusicians.org.uk/get-support/funding',           priority:'Normal', type:'funding'  },
  { name:'IFPI Nordic',                 url:'https://www.ifpi.org',                                           priority:'Normal', type:'funding'  },

  // ── Netherlands / Belgium ────────────────────────────────
  { name:'Eurosonic Noorderslag',       url:'https://www.eurosonic-noorderslag.nl/showcases/applications/',  priority:'High',   type:'festival' },
  { name:'Waves Vienna',                url:'https://wavesvienna.com/application/',                           priority:'High',   type:'festival' },
  { name:'Glimps Festival',             url:'https://glimps.be',                                              priority:'Normal', type:'festival' },

  // ── France ───────────────────────────────────────────────
  { name:'MAMA Paris',                  url:'https://www.mama-event.com/en/',                                 priority:'High',   type:'festival' },
  { name:'Trans Musicales Rennes',      url:'https://www.lestrans.com/les-trans-musicales/candidater/',       priority:'High',   type:'festival' },
  { name:'La Route du Rock',            url:'https://www.laroutedurock.com',                                  priority:'High',   type:'festival' },
  { name:'Printemps de Bourges',        url:'https://www.printemps-bourges.com',                              priority:'High',   type:'festival' },
  { name:'CNM France (funding)',        url:'https://cnm.fr/aides/',                                          priority:'High',   type:'funding'  },
  { name:'Institut français (touring)', url:'https://www.institutfrancais.com/en/programmes/performing-arts/',priority:'Normal', type:'funding'  },

  // ── UK ───────────────────────────────────────────────────
  { name:'The Great Escape',            url:'https://greatescapefestival.com/apply/',                         priority:'High',   type:'festival' },
  { name:'Green Man Festival',          url:'https://www.greenman.net',                                       priority:'High',   type:'festival' },
  { name:'End of the Road',             url:'https://endoftheroadfestival.com',                               priority:'Normal', type:'festival' },
  { name:'Liverpool Sound City',        url:'https://www.liverpoolsoundcity.co.uk',                           priority:'Normal', type:'festival' },
  { name:'BBC Introducing',             url:'https://www.bbc.co.uk/music/introducing',                        priority:'High',   type:'radio'    },

  // ── Scandinavia ──────────────────────────────────────────
  { name:'Iceland Airwaves',            url:'https://icelandairwaves.is',                                     priority:'High',   type:'festival' },
  { name:'by:Larm Oslo',                url:'https://bylarm.no/en/apply/',                                    priority:'High',   type:'festival' },
  { name:'Way Out West Gothenburg',     url:'https://www.wayoutwest.se',                                      priority:'High',   type:'festival' },
  { name:'Roskilde Festival',           url:'https://www.roskilde-festival.dk/da/musik/ansoeg-om-at-spille/', priority:'High',   type:'festival' },
  { name:'Slottsfjell Norway',          url:'https://slottsfjell.no',                                         priority:'Normal', type:'festival' },

  // ── Southern Europe ──────────────────────────────────────
  { name:'MIL Lisboa',                  url:'https://mil-lisboa.pt',                                          priority:'High',   type:'festival' },
  { name:'Primavera Sound Barcelona',   url:'https://www.primaverasound.com',                                  priority:'High',   type:'festival' },
  { name:'Sónar Barcelona',             url:'https://sonar.es/en/artists/apply/',                              priority:'High',   type:'festival' },
  { name:'Mad Cool Madrid',             url:'https://www.madcoolfestival.es',                                  priority:'High',   type:'festival' },
  { name:'Linecheck Milan',             url:'https://www.linecheckmusic.com',                                  priority:'High',   type:'festival' },
  { name:'MI AMI Festival Milan',       url:'https://miamifestival.it',                                       priority:'Normal', type:'festival' },

  // ── International ────────────────────────────────────────
  { name:'Folk Alliance International', url:'https://www.folkalliance.org/conference/',                       priority:'Normal', type:'festival' },
  { name:'SXSW',                        url:'https://www.sxsw.com/music/applications/',                       priority:'High',   type:'festival' },
  { name:'M for Montreal',              url:'https://mformontreal.com/en/artists/',                           priority:'High',   type:'festival' },
  { name:'Womex',                       url:'https://www.womex.com/virtual/womex/apply',                      priority:'High',   type:'festival' },
  { name:'Midem (Cannes)',              url:'https://www.midem.com',                                           priority:'Normal', type:'festival' },
];

const KEYWORDS = [
  'open call','open for','apply now','applications open','submit','submission',
  'jetzt bewerben','bewerbung','förderantrag','einreichung','einreichen',
  'deadline','frist','förderung','stipendium','grant','funding round',
  'pitch your music','new artists','emerging artists','unsigned','demo'
];

function scanSources() {
  const sheet   = getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = headersFromRow(data[0]);
  const existing = data.slice(1).map(r => rowToObject(headers, r));
  const existingURLs = new Set(existing.map(o => o.URL));

  const today = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd');
  let added = 0;

  SCAN_SOURCES.forEach(src => {
    if (existingURLs.has(src.url)) return; // already tracked
    try {
      const resp = UrlFetchApp.fetch(src.url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (compatible; lovelee-scanner/1.0)' }
      });
      if (resp.getResponseCode() !== 200) return;

      const html = resp.getContentText().toLowerCase();
      const hit  = KEYWORDS.some(kw => html.includes(kw));
      if (!hit) return;

      // Try to extract a deadline date
      const dateMatch = html.match(/(\d{4}-\d{2}-\d{2})|(\d{2}\.\d{2}\.\d{4})/);
      let deadline = 'rolling';
      if (dateMatch) {
        const raw = dateMatch[0];
        if (raw.includes('-')) deadline = raw;
        else { const [d,m,y] = raw.split('.'); deadline = `${y}-${m}-${d}`; }
      }

      const row = headers.map(h => {
        switch (h) {
          case 'ID':           return nextId(sheet, headers);
          case 'Title':        return src.name + ' — Open Call (auto-scanned)';
          case 'Organization': return src.name;
          case 'Type':         return src.type || 'funding';
          case 'Status':       return 'watching';
          case 'Deadline':     return deadline;
          case 'Amount':       return 'variable';
          case 'Region':       return 'International';
          case 'Source':       return 'Auto-scan';
          case 'URL':          return src.url;
          case 'Priority':     return src.priority || 'Normal';
          case 'Notes':        return 'Auto-detected. Review and update details.';
          case 'Date Added':   return today;
          case 'Date Updated': return today;
          case 'Applied':      return 'FALSE';
          default:             return '';
        }
      });

      sheet.appendRow(row);
      existingURLs.add(src.url);
      added++;
    } catch (err) {
      Logger.log('Scan error for ' + src.name + ': ' + err.message);
    }
  });

  return { added, scanned: SCAN_SOURCES.length, timestamp: today };
}

// ── Monthly auto-trigger ──────────────────────────────────────
// Run setupMonthlyTrigger() ONCE manually to activate.

function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runMonthlyScan')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runMonthlyScan').timeBased().onMonthDay(1).atHour(9).create();
  Logger.log('Monthly trigger set: 1st of each month at 9am Berlin.');
}

function runMonthlyScan() {
  const result = scanSources();
  Logger.log('Monthly scan: ' + JSON.stringify(result));
}

// ── Test functions ────────────────────────────────────────────

function testGet() {
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();
  const headers = headersFromRow(data[0]);
  Logger.log(JSON.stringify(data.slice(1,4).map(r => rowToObject(headers, r)), null, 2));
}

function testMarkApplied() {
  Logger.log(JSON.stringify(markApplied('OPP002')));
}

function testPost() {
  const fake = { postData: { contents: JSON.stringify({
    Title:'TEST — delete me', Organization:'Test', Type:'funding',
    Status:'watching', Deadline:'2026-12-31', Amount:'€0',
    Region:'Germany', Source:'test', Priority:'Low'
  })}};
  Logger.log(doPost(fake).getContent());
}
