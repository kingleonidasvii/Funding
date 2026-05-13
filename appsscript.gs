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

const SHEET_ID          = '1Si1_GF1DVPq7ZNfH5NPOxQaH6gJDtiDDbqTUYT7kQg4';
const OPPORTUNITIES_TAB = 'Opportunities';
const RADIO_TAB         = 'Radio';

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
    if (action === 'setPriority')  return corsResponse(setPriorityById(e.parameter.id, e.parameter.priority));
    if (action === 'updateField')  return corsResponse(updateField(e.parameter.id, e.parameter.field, e.parameter.value));
    if (action === 'listRadio')    return corsResponse(listRadio());
    if (action === 'updateRadio')  return corsResponse(updateRadio(e.parameter.id, e.parameter.field, e.parameter.value));

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

// ── Set priority on a single row ─────────────────────────────
function setPriorityById(id, priority) {
  if (!id) return { error: 'No ID provided' };
  if (priority !== 'High' && priority !== 'Normal') return { error: 'Priority must be High or Normal' };
  const sheet   = getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = headersFromRow(data[0]);
  const idCol   = headers.indexOf('ID');
  const prioCol = headers.indexOf('Priority');
  const updCol  = headers.indexOf('Date Updated');
  if (idCol < 0 || prioCol < 0) return { error: 'Column not found' };

  const today = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd');
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idCol]).trim() === String(id).trim()) {
      sheet.getRange(r + 1, prioCol + 1).setValue(priority);
      if (updCol >= 0) sheet.getRange(r + 1, updCol + 1).setValue(today);
      return { success: true, id, priority };
    }
  }
  return { error: 'ID not found: ' + id };
}

// ── Update a single field on any row ─────────────────────────
function updateField(id, field, value) {
  if (!id || !field) return { error: 'Missing id or field' };
  const sheet   = getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = headersFromRow(data[0]);
  const idCol   = headers.indexOf('ID');
  const fCol    = headers.indexOf(field);
  const updCol  = headers.indexOf('Date Updated');
  if (idCol < 0 || fCol < 0) return { error: 'Column not found: ' + field };
  const today = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd');
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idCol]).trim() === String(id).trim()) {
      sheet.getRange(r + 1, fCol + 1).setValue(value);
      if (updCol >= 0) sheet.getRange(r + 1, updCol + 1).setValue(today);
      return { success: true, id, field, value };
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

// ── Reset priorities ──────────────────────────────────────────
// Run once to bulk-correct all existing rows using HIGH_PRIORITY_ORGS.
// High = org name matches one of the known career-defining sources.
function resetPriorities() {
  const sheet   = getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = headersFromRow(data[0]);
  const orgCol  = headers.indexOf('Organization');
  const titleCol= headers.indexOf('Title');
  const prioCol = headers.indexOf('Priority');
  if (prioCol < 0) { Logger.log('No Priority column found.'); return; }

  let changed = 0;
  for (let r = 1; r < data.length; r++) {
    const org   = String(data[r][orgCol]   || '').toLowerCase();
    const title = String(data[r][titleCol] || '').toLowerCase();
    const isHigh = Array.from(HIGH_PRIORITY_ORGS).some(k => org.includes(k) || title.includes(k));
    const newPrio = isHigh ? 'High' : 'Normal';
    if (data[r][prioCol] !== newPrio) {
      sheet.getRange(r + 1, prioCol + 1).setValue(newPrio);
      changed++;
    }
  }
  Logger.log('resetPriorities complete: ' + changed + ' rows updated.');
}

// ── Monitor sources ───────────────────────────────────────────
// These are known high-priority sources that get RE-CHECKED every
// weekly scan regardless of whether they're already in the sheet.
// Goal: catch when a funding round or festival open call goes live.
// When detected as open, the existing watching row is upgraded to
// "open" — or a new row is added if none exists.

// High = genuinely career-defining for an indie Berlin artist. Everything else is Normal.
const HIGH_PRIORITY_ORGS = new Set([
  'initiative musik', 'musicboard berlin', 'creative europe',
  'reeperbahn festival', 'eurosonic', 'esns', 'mama paris', 'mama',
  'trans musicales', 'by:larm', 'bylarm', 'iceland airwaves',
  'sxsw', 'the great escape', 'great escape',
  'lollapalooza berlin', 'roskilde'
]);

const MONITOR_SOURCES = [
  // German funding — rounds open/close on a schedule
  { name:'Initiative Musik',           url:'https://www.initiative-musik.de/foerderung/',                 priority:'High',   type:'funding'  },
  { name:'Musicboard Berlin',          url:'https://www.musicboard-berlin.de/foerderung/',                priority:'High',   type:'funding'  },
  { name:'GEMA Kulturförderung',       url:'https://www.gema.de/kulturfoerderung-online',                 priority:'Normal', type:'funding'  },
  { name:'Musikfonds',                 url:'https://www.musikfonds.de/foerderung/',                       priority:'Normal', type:'funding'  },
  { name:'Fonds Darstellende Künste',  url:'https://www.fonds-daku.de/foerderung/',                       priority:'Normal', type:'funding'  },
  { name:'German Music Export',        url:'https://www.german-music-export.de/en/funding/',              priority:'Normal', type:'funding'  },
  { name:'Berlin Music Commission',    url:'https://www.berlin-music-commission.de',                      priority:'Normal', type:'funding'  },
  { name:'Berlin Senate — Musik',      url:'https://www.berlin.de/sen/kultur/foerderung/antragsfristen/',priority:'Normal', type:'funding'  },
  // EU funding
  { name:'Creative Europe',            url:'https://culture.ec.europa.eu/calls',                         priority:'High',   type:'funding'  },
  { name:'PRS Foundation UK',          url:'https://prsfoundation.com/funding-support/funding/',         priority:'Normal', type:'funding'  },
  { name:'CNM France',                 url:'https://cnm.fr/aides/',                                      priority:'Normal', type:'funding'  },
  // Key festivals with annual open calls
  { name:'Eurosonic Noorderslag',      url:'https://www.eurosonic-noorderslag.nl/showcases/',            priority:'High',   type:'festival' },
  { name:'The Great Escape',           url:'https://greatescapefestival.com',                            priority:'High',   type:'festival' },
  { name:'Waves Vienna',               url:'https://wavesvienna.com',                                    priority:'Normal', type:'festival' },
  { name:'MAMA Paris',                 url:'https://www.mama-event.com/en/',                             priority:'High',   type:'festival' },
  { name:'Trans Musicales',            url:'https://www.lestrans.com/les-trans-musicales/',              priority:'High',   type:'festival' },
  { name:'by:Larm Oslo',               url:'https://bylarm.no/en/',                                      priority:'High',   type:'festival' },
  { name:'Iceland Airwaves',           url:'https://icelandairwaves.is',                                 priority:'High',   type:'festival' },
  { name:'SXSW Austin',                url:'https://www.sxsw.com/music/applications/',                  priority:'High',   type:'festival' },
  { name:'MIL Lisboa',                 url:'https://mil-lisboa.pt',                                      priority:'Normal', type:'festival' },
  { name:'MENT Ljubljana',             url:'https://ment.si',                                            priority:'Normal', type:'festival' },
  { name:'Glastonbury Emerging Talent',url:'https://www.glastonburyfestivals.co.uk/information/emerging-talent/', priority:'Normal', type:'festival' },
];

// Strong signals that a round/call is currently open — not just that the
// page mentions funding in general. These go beyond the general KEYWORDS list.
const OPEN_NOW_KEYWORDS = [
  // German — active round language
  'jetzt bewerben', 'jetzt einreichen', 'bewerbungsschluss', 'bewerbungsfrist',
  'einreichfrist', 'einreichungsschluss', 'antragstellung möglich',
  'antragsschluss', 'aktuelle ausschreibung', 'runde ist geöffnet',
  'bewerbungen sind ab', 'bewerbung einreichen', 'förderantrag stellen',
  // English — active round language
  'applications are open', 'apply now', 'submissions are open', 'now accepting',
  'open for applications', 'call is open', 'deadline for applications',
  'submit your application', 'applications close', 'apply before',
  'application deadline', 'open call', 'apply by',
  // French
  'candidatures ouvertes', 'dépôt de candidature', 'date limite de candidature',
  'appel ouvert', 'soumettre votre candidature',
  // Dutch
  'aanvragen mogelijk', 'aanmeldingsdeadline', 'open voor aanmeldingen',
];

// Check if a detected deadline is in the future (within 6 months)
function isFutureDeadline(dateStr) {
  const d = new Date(dateStr);
  const now = new Date();
  const sixMonths = new Date(); sixMonths.setMonth(sixMonths.getMonth() + 6);
  return d > now && d < sixMonths;
}

// Monitor known sources — re-checks every week, upgrades watching→open
function monitorSources() {
  const sheet   = getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = headersFromRow(data[0]);
  const rows    = data.slice(1).map((r, i) => ({ ...rowToObject(headers, r), _row: i + 2 }));

  const today     = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd');
  const todayDate = new Date(); todayDate.setHours(0, 0, 0, 0);
  const twoWeeksAgo = new Date(todayDate); twoWeeksAgo.setDate(twoWeeksAgo.getDate() - 14);
  const twoWeeksAgoStr = Utilities.formatDate(twoWeeksAgo, 'Europe/Berlin', 'yyyy-MM-dd');

  let upgraded = 0, added = 0;

  MONITOR_SOURCES.forEach(src => {
    try {
      const resp = UrlFetchApp.fetch(src.url, {
        muteHttpExceptions: true, followRedirects: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (compatible; lovelee-monitor/1.0)' }
      });
      if (resp.getResponseCode() !== 200) return;

      const html = resp.getContentText().toLowerCase();

      // Must match at least one strong "open now" keyword
      const isOpen = OPEN_NOW_KEYWORDS.some(kw => html.includes(kw));
      if (!isOpen) return;

      // Try to extract a future deadline date
      const allDates = [];
      const patterns = [
        /(\d{4}-\d{2}-\d{2})/g,
        /(\d{2})\.(\d{2})\.(\d{4})/g,
      ];
      let m;
      const p1 = /\d{4}-\d{2}-\d{2}/g;
      while ((m = p1.exec(html)) !== null) { if (isFutureDeadline(m[0])) allDates.push(m[0]); }
      const p2 = /(\d{2})\.(\d{2})\.(\d{4})/g;
      while ((m = p2.exec(html)) !== null) {
        const iso = m[3] + '-' + m[2] + '-' + m[1];
        if (isFutureDeadline(iso)) allDates.push(iso);
      }
      allDates.sort();
      const deadline = allDates.length ? allDates[0] : 'rolling';

      // Skip if we already added/upgraded a row for this source in the last 14 days
      const recentlyHandled = rows.find(o =>
        (o.Organization === src.name || o.Source === src.name) &&
        o.Status === 'open' &&
        o['Date Updated'] >= twoWeeksAgoStr
      );
      if (recentlyHandled) return;

      // Find an existing watching row for this source to upgrade
      const watchingRow = rows.find(o =>
        (o.Organization === src.name || o.Source === src.name) &&
        o.Status === 'watching'
      );

      const idxOf = col => headers.indexOf(col);

      if (watchingRow) {
        // Upgrade: watching → open
        const r = watchingRow._row;
        if (idxOf('Status')       >= 0) sheet.getRange(r, idxOf('Status') + 1).setValue('open');
        if (idxOf('Deadline')     >= 0) sheet.getRange(r, idxOf('Deadline') + 1).setValue(deadline);
        if (idxOf('Date Updated') >= 0) sheet.getRange(r, idxOf('Date Updated') + 1).setValue(today);
        if (idxOf('Notes')        >= 0) {
          const note = 'Auto-detected OPEN ' + today + (deadline !== 'rolling' ? ' · Deadline: ' + deadline : '') + '. Verify on site.';
          sheet.getRange(r, idxOf('Notes') + 1).setValue(note);
        }
        // Also update the title to remove "(watching)"
        if (idxOf('Title') >= 0) {
          const oldTitle = watchingRow.Title || '';
          sheet.getRange(r, idxOf('Title') + 1).setValue(oldTitle.replace(' (watching)', '') + ' — Open Call');
        }
        upgraded++;
      } else {
        // No watching row — add fresh
        const row = headers.map(h => {
          switch (h) {
            case 'ID':           return nextId(sheet, headers);
            case 'Title':        return src.name + ' — Open Call (auto-detected)';
            case 'Organization': return src.name;
            case 'Type':         return src.type || 'funding';
            case 'Status':       return 'open';
            case 'Deadline':     return deadline;
            case 'Priority':     return src.priority || 'Normal';
            case 'Source':       return src.name;
            case 'URL':          return src.url;
            case 'Notes':        return 'Auto-detected OPEN ' + today + (deadline !== 'rolling' ? ' · Deadline: ' + deadline : '') + '. Verify on site.';
            case 'Date Added':   return today;
            case 'Date Updated': return today;
            case 'Applied':      return 'FALSE';
            default:             return '';
          }
        });
        sheet.appendRow(row);
        added++;
      }
    } catch (err) {
      Logger.log('Monitor error for ' + src.name + ': ' + err.message);
    }
  });

  Logger.log('Monitor complete: ' + upgraded + ' upgraded, ' + added + ' added.');
  return { monitored: MONITOR_SOURCES.length, upgraded, added };
}

// ── Scan sources ─────────────────────────────────────────────
// Fetches public source pages, detects active open calls by keyword,
// and appends new "watching" rows. Login-required sources (e.g.
// Music Group Berlin) should be handled by uploading their PDF
// in Claude chat — Claude will extract and add entries via POST.

const SCAN_SOURCES = [
  // ── Germany: Funding ─────────────────────────────────────
  { name:'Musicboard Berlin',           url:'https://www.musicboard-berlin.de/foerderung/',                   priority:'High',   type:'funding'  },
  { name:'Initiative Musik',            url:'https://www.initiative-musik.de/foerderung/',                   priority:'High',   type:'funding'  },
  { name:'kreativkultur.berlin',        url:'https://kreativkultur.berlin',                                  priority:'Normal', type:'funding'  },
  { name:'Berlin Music Commission',     url:'https://www.berlin-music-commission.de',                        priority:'Normal', type:'funding'  },
  { name:'Musikfonds',                  url:'https://www.musikfonds.de/foerderung/',                         priority:'Normal', type:'funding'  },
  { name:'GEMA Kulturförderung',        url:'https://www.gema.de/kulturfoerderung-online',                   priority:'Normal', type:'funding'  },
  { name:'German Music Export',         url:'https://www.german-music-export.de/en/funding/',                priority:'High',   type:'funding'  },
  { name:'Backstage PRO',               url:'https://www.backstage.de/ratgeber/foerderung/',                 priority:'Normal', type:'funding'  },
  { name:'Fonds Darstellende Künste',   url:'https://www.fonds-daku.de/foerderung/',                         priority:'Normal', type:'funding'  },

  // ── Germany: Festivals & Major Stages ────────────────────
  { name:'Reeperbahn Festival',         url:'https://www.reeperbahnfestival.com/artist-submission',         priority:'High',   type:'festival' },
  { name:'c/o pop Convention',          url:'https://copop.de/convention/',                                  priority:'High',   type:'festival' },
  { name:'SWR3 New Pop Festival',       url:'https://www.swr.de/swr3/musik/new-pop-festival/',               priority:'High',   type:'festival' },
  { name:'Lollapalooza Berlin',         url:'https://www.lollapaloozade.com/',                               priority:'High',   type:'festival' },
  { name:'Melt Festival',               url:'https://www.meltfestival.de',                                   priority:'High',   type:'festival' },
  { name:'Hurricane Festival',          url:'https://www.hurricane.de',                                      priority:'High',   type:'festival' },
  { name:'Southside Festival',          url:'https://www.southside.de',                                      priority:'High',   type:'festival' },
  { name:'Rock am Ring / Rock im Park', url:'https://www.rock-am-ring.com',                                  priority:'High',   type:'festival' },
  { name:'Deichbrand Festival',         url:'https://www.deichbrand.de',                                     priority:'High',   type:'festival' },
  { name:'Highfield Festival Leipzig',  url:'https://www.highfield.de',                                      priority:'High',   type:'festival' },
  { name:'Open Flair Eschwege',         url:'https://www.open-flair.de',                                     priority:'Normal', type:'festival' },
  { name:'Taubertal Festival',          url:'https://www.taubertal-festival.de',                             priority:'Normal', type:'festival' },
  { name:'Traumzeit Festival Duisburg', url:'https://www.traumzeit-festival.de',                             priority:'Normal', type:'festival' },
  { name:'Maifeld Derby Mannheim',      url:'https://www.maifeld-derby.de',                                  priority:'Normal', type:'festival' },
  { name:'Elbjazz Hamburg',             url:'https://www.elbjazz.de',                                        priority:'Normal', type:'festival' },
  { name:'MS Dockville Hamburg',        url:'https://msdockville.de',                                        priority:'Normal', type:'festival' },
  { name:'Immergut Festival',           url:'https://www.immergutrocken.de',                                 priority:'Normal', type:'festival' },
  { name:'Haldern Pop',                 url:'https://haldern-pop.de',                                        priority:'Normal', type:'festival' },

  // ── Germany: Radio & Media ───────────────────────────────
  { name:'Fritz (rbb)',                 url:'https://www.fritz.de/programm/aktionen/',                       priority:'High',   type:'radio'    },
  { name:'ByteFM',                      url:'https://www.byte.fm/features/',                                 priority:'High',   type:'radio'    },
  { name:'radioeins (rbb)',             url:'https://www.radioeins.de/musik/',                                priority:'Normal', type:'radio'    },
  { name:'Deutschlandfunk Kultur',      url:'https://www.deutschlandfunkkultur.de/musik/',                   priority:'Normal', type:'radio'    },
  { name:'MDR Kultur',                  url:'https://www.mdr.de/mdr-kultur/radio/index.html',                priority:'Normal', type:'radio'    },

  // ── EU: Funding ──────────────────────────────────────────
  { name:'Creative Europe',             url:'https://culture.ec.europa.eu/calls',                           priority:'High',   type:'funding'  },
  { name:'SHAPE+',                      url:'https://shapeplatform.eu/open-call/',                           priority:'High',   type:'open-call'},
  { name:'Liveurope',                   url:'https://liveurope.eu',                                          priority:'High',   type:'funding'  },
  { name:'Fonds Podiumkunsten NL',      url:'https://www.fondspodiumkunsten.nl/subsidies/',                  priority:'Normal', type:'funding'  },
  { name:'Österreichischer Musikfonds', url:'https://www.musikfonds.at/foerderungen/',                       priority:'Normal', type:'funding'  },
  { name:'Pro Helvetia Switzerland',    url:'https://prohelvetia.ch/en/funding/',                            priority:'Normal', type:'funding'  },
  { name:'Music Finland Export',        url:'https://musicfinland.com/en/funding/',                          priority:'Normal', type:'funding'  },
  { name:'Music Norway Export',         url:'https://musicnorway.no/funding/',                               priority:'Normal', type:'funding'  },
  { name:'Swedish Arts Council',        url:'https://www.kulturradet.se/en/apply-for-grants/',               priority:'Normal', type:'funding'  },
  { name:'PRS Foundation UK',           url:'https://prsfoundation.com/funding-support/funding/',            priority:'High',   type:'funding'  },
  { name:'Help Musicians UK',           url:'https://www.helpmusicians.org.uk/get-support/funding',          priority:'Normal', type:'funding'  },
  { name:'Arts Council England',        url:'https://www.artscouncil.org.uk/projectgrants',                  priority:'Normal', type:'funding'  },
  { name:'CNM France',                  url:'https://cnm.fr/aides/',                                         priority:'High',   type:'funding'  },

  // ── Netherlands ──────────────────────────────────────────
  { name:'Eurosonic Noorderslag',       url:'https://www.eurosonic-noorderslag.nl/showcases/',               priority:'High',   type:'festival' },
  { name:'Lowlands Festival',           url:'https://www.lowlands.nl',                                       priority:'High',   type:'festival' },
  { name:'Down the Rabbit Hole',        url:'https://www.downtherabbithole.nl',                              priority:'High',   type:'festival' },
  { name:'Best Kept Secret NL',         url:'https://www.bestkeptsecret.nl',                                 priority:'High',   type:'festival' },

  // ── Belgium ──────────────────────────────────────────────
  { name:'Pukkelpop Belgium',           url:'https://www.pukkelpop.be',                                      priority:'High',   type:'festival' },
  { name:'Rock Werchter Belgium',       url:'https://www.rockwerchter.be',                                   priority:'High',   type:'festival' },
  { name:'Dour Festival Belgium',       url:'https://www.dourfestival.eu',                                   priority:'Normal', type:'festival' },
  { name:'Glimps Festival Ghent',       url:'https://glimps.be',                                             priority:'Normal', type:'festival' },

  // ── Austria ──────────────────────────────────────────────
  { name:'Waves Vienna',                url:'https://wavesvienna.com',                                       priority:'High',   type:'festival' },

  // ── France ───────────────────────────────────────────────
  { name:'MAMA Paris',                  url:'https://www.mama-event.com/en/',                                priority:'High',   type:'festival' },
  { name:'Trans Musicales Rennes',      url:'https://www.lestrans.com/les-trans-musicales/',                 priority:'High',   type:'festival' },
  { name:'La Route du Rock',            url:'https://www.laroutedurock.com',                                 priority:'High',   type:'festival' },
  { name:'Printemps de Bourges',        url:'https://www.printemps-bourges.com',                             priority:'High',   type:'festival' },
  { name:'Nuits Sonores Lyon',          url:'https://www.nuits-sonores.com',                                 priority:'High',   type:'festival' },
  { name:'Rock en Seine Paris',         url:'https://www.rockenseine.com',                                   priority:'High',   type:'festival' },
  { name:'Pitchfork Music Fest Paris',  url:'https://www.pitchforkparis.com',                                priority:'High',   type:'festival' },

  // ── UK ───────────────────────────────────────────────────
  { name:'The Great Escape Brighton',   url:'https://greatescapefestival.com',                               priority:'High',   type:'festival' },
  { name:'Green Man Festival',          url:'https://www.greenman.net',                                      priority:'High',   type:'festival' },
  { name:'End of the Road',             url:'https://endoftheroadfestival.com',                              priority:'Normal', type:'festival' },
  { name:'Liverpool Sound City',        url:'https://www.liverpoolsoundcity.co.uk',                          priority:'Normal', type:'festival' },
  { name:'Latitude Festival',           url:'https://www.latitudefestival.com',                              priority:'High',   type:'festival' },
  { name:'Glastonbury Emerging Talent', url:'https://www.glastonburyfestivals.co.uk/information/emerging-talent/', priority:'High', type:'festival' },
  { name:'BBC Introducing',             url:'https://www.bbc.co.uk/music/introducing',                       priority:'High',   type:'radio'    },
  { name:'XpoNorth Scotland',           url:'https://xponorth.co.uk',                                        priority:'High',   type:'festival' },
  { name:'Celtic Connections Glasgow',  url:'https://www.celticconnections.com',                             priority:'High',   type:'festival' },
  { name:'Output Belfast',              url:'https://outputbelfast.com',                                     priority:'High',   type:'festival' },

  // ── Ireland ──────────────────────────────────────────────
  { name:'Other Voices Ireland',        url:'https://www.othervoices.ie',                                    priority:'High',   type:'festival' },
  { name:'Hard Working Class Heroes',   url:'https://www.hwch.net',                                          priority:'High',   type:'festival' },
  { name:'Electric Picnic Ireland',     url:'https://www.electricpicnic.ie',                                 priority:'High',   type:'festival' },
  { name:'First Music Contact Ireland', url:'https://fmc.ie/funding/',                                       priority:'Normal', type:'funding'  },

  // ── Scandinavia ──────────────────────────────────────────
  { name:'Iceland Airwaves',            url:'https://icelandairwaves.is',                                    priority:'High',   type:'festival' },
  { name:'by:Larm Oslo',                url:'https://bylarm.no/en/',                                         priority:'High',   type:'festival' },
  { name:'Way Out West Gothenburg',     url:'https://www.wayoutwest.se',                                     priority:'High',   type:'festival' },
  { name:'Roskilde Festival',           url:'https://www.roskilde-festival.dk/en/',                          priority:'High',   type:'festival' },
  { name:'Øya Festival Oslo',           url:'https://www.oyafestivalen.no',                                  priority:'High',   type:'festival' },
  { name:'Flow Festival Helsinki',      url:'https://www.flowfestival.com',                                  priority:'High',   type:'festival' },

  // ── Southern Europe ──────────────────────────────────────
  { name:'MIL Lisboa',                  url:'https://mil-lisboa.pt',                                         priority:'High',   type:'festival' },
  { name:'NOS Alive Portugal',          url:'https://nosalive.com',                                          priority:'High',   type:'festival' },
  { name:'Primavera Sound Barcelona',   url:'https://www.primaverasound.com',                                 priority:'High',   type:'festival' },
  { name:'Sónar Barcelona',             url:'https://sonar.es/en/',                                          priority:'High',   type:'festival' },
  { name:'Mad Cool Madrid',             url:'https://www.madcoolfestival.es',                                priority:'High',   type:'festival' },
  { name:'Bilbao BBK Live',             url:'https://www.bilbaobblive.com',                                  priority:'High',   type:'festival' },
  { name:'Linecheck Milan',             url:'https://www.linecheckmusic.com/en/',                            priority:'High',   type:'festival' },
  { name:'MI AMI Festival Milan',       url:'https://miamifestival.it',                                      priority:'Normal', type:'festival' },
  { name:'Medimex Bari',                url:'https://www.medimex.it',                                        priority:'Normal', type:'festival' },
  { name:'Sziget Festival Budapest',    url:'https://szigetfestival.com',                                    priority:'High',   type:'festival' },

  // ── Switzerland ──────────────────────────────────────────
  { name:'Montreux Jazz Festival',      url:'https://www.montreuxjazzfestival.com',                          priority:'High',   type:'festival' },
  { name:'Paléo Festival Nyon',         url:'https://www.paleo.ch',                                          priority:'High',   type:'festival' },

  // ── USA / Canada ─────────────────────────────────────────
  { name:'SXSW Austin',                 url:'https://www.sxsw.com/music/applications/',                      priority:'High',   type:'festival' },
  { name:'M for Montreal',              url:'https://mformontreal.com/en/artists/',                          priority:'High',   type:'festival' },
  { name:'Folk Alliance International', url:'https://www.folkalliance.org/conference/',                      priority:'Normal', type:'festival' },
  { name:'CMW Toronto',                 url:'https://cmw.net/showcase-applications/',                        priority:'Normal', type:'festival' },

  // ── Global industry ──────────────────────────────────────
  { name:'Womex',                       url:'https://www.womex.com/virtual/womex/apply',                     priority:'High',   type:'festival' },
  { name:'Womad',                       url:'https://womad.co.uk',                                           priority:'High',   type:'festival' },

  // ── Opportunity aggregators ───────────────────────────────────
  { name:'Ditto Music Opportunities',   url:'https://dittomusic.com/en/blog/latest-music-opportunities',    priority:'High',   type:'open-call' },
  { name:'Creatives Unite',             url:'https://creativesunite.eu/open-calls/',                        priority:'Normal', type:'open-call' },
  { name:'On the Move (Music Mobility)',url:'https://on-the-move.org/resources/funding',                    priority:'Normal', type:'funding'   },

  // ── More UK showcases ────────────────────────────────────────
  { name:'Focus Wales',                 url:'https://focuswales.com',                                       priority:'Normal', type:'festival'  },
  { name:'SXSW London',                 url:'https://london.sxsw.com',                                      priority:'High',   type:'festival'  },

  // ── Central / Eastern Europe ─────────────────────────────────
  { name:'MENT Ljubljana',              url:'https://ment.si',                                              priority:'High',   type:'festival'  },

  // ── Streaming platform emerging artist programs ───────────────
  { name:'Spotify RADAR',               url:'https://artists.spotify.com/blog/spotify-radar',               priority:'High',   type:'open-call' },
  { name:'Apple Music Up Next',         url:'https://artists.apple.com',                                    priority:'High',   type:'open-call' },
  { name:'Deezer Next',                 url:'https://www.deezer.com/en/deezer-next',                        priority:'High',   type:'open-call' },
  { name:'Amazon Music Breakthrough',   url:'https://music.amazon.com/breakthrough',                        priority:'Normal', type:'open-call' },

  // ── Awards & prizes ──────────────────────────────────────────
  { name:'EBBA Awards',                 url:'https://www.eurosonic-noorderslag.nl/ebba/',                   priority:'High',   type:'open-call' },

  // ── City government funding pages ───────────────────────────
  { name:'Berlin Senate — Förderung Musik',  url:'https://www.berlin.de/sen/kultur/foerderung/antragsfristen/', priority:'High',   type:'funding'   },
  { name:'Vienna MA7 — Musikförderung',      url:'https://www.wien.gv.at/kultur/foerderungen-musik',           priority:'Normal', type:'funding'   },
];

const KEYWORDS = [
  // English (UK / IE / US / AU)
  'open call','open for','apply now','applications open','submit','submission',
  'deadline','grant','funding round','pitch your music','new artists',
  'emerging artists','unsigned','demo','apply here','applications now open',
  'call for entries','call for artists','now accepting','accepting submissions',
  'apply for','open applications','music grant','touring grant','artist fund',
  // German
  'jetzt bewerben','bewerbung','förderantrag','einreichung','einreichen',
  'frist','förderung','stipendium','bewerbungsschluss','ausschreibung',
  'projektförderung','musikförderung','reisekostenförderung',
  // Dutch
  'aanvragen','subsidie','aanmelding','open voor','aanmelden','indienen',
  'subsidieregeling','aanvraagronde','muzieksubsidie','open oproep',
  // French
  'candidater','candidature','appel à','dossier de candidature','soumettre',
  'bourse','financement','date limite','appel ouvert','appel à projets',
  'dépôt de candidature','appel à candidatures',
  // Spanish
  'convocatoria','solicitud','inscripción','plazo','subvención','beca',
  'abrir convocatoria','presentar candidatura','fecha límite','open call',
  'artistas emergentes','enviar demo','formulario de inscripción',
  // Italian
  'bando','candidatura','iscrizione','scadenza','contributo','borsa',
  'aperto alle candidature','chiama artisti','presentare domanda',
  'bando aperto','sovvenzione','call for artists',
  // Portuguese
  'candidatura','inscrição','prazo','bolsa','financiamento',
  'candidatar','submeter','chamada aberta','apoio',
  // Scandinavian (Swedish / Norwegian / Danish / Finnish)
  'ansökan','søknad','ansök','tilmeld','haku','stipend',
  'open for søknader','söka bidrag','musikstipendium','residens'
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

// ── Radio tab ────────────────────────────────────────────────
// Columns: ID | Station | Country | Format | URL | Submit URL | Notes | Last Submitted | Last Result | Priority
//
// ONE-TIME SETUP: Run → setupRadioTab() once to create the tab and seed the stations.
// Afterwards the Radio tab is a persistent store — update it via the app or directly.

const RADIO_SEED_DATA = [
  // Germany
  ['RAD001','Fritz (rbb) — Lokalmatadoren','Germany','Indie/Pop/Alternative',      'https://www.fritz.de/programm/aktionen/lokalmatadoren/','musik@fritz.de',                                    'PAID live studio session for Berlin artists. Pitch via musik@fritz.de — short bio, Spotify, press folder. Write in German. No deadline — always open.','','','High'],
  ['RAD002','ByteFM',                'Germany',       'Indie/Alternative',          'https://www.byte.fm',                'https://www.byte.fm/redaktion/',                  'Indie-focused taste-maker. Email submissions to redaktion@byte.fm.',                    '','','High'],
  ['RAD003','radioeins (rbb)',        'Germany',       'Pop/Rock/Indie',             'https://www.radioeins.de',           'https://www.radioeins.de/service/kontakt.html',   'RBB flagship pop station. Submit new releases via contact form.',                       '','','Normal'],
  ['RAD004','Deutschlandfunk Kultur', 'Germany',       'Culture/Indie',              'https://www.deutschlandfunkkultur.de','',                                                'Public cultural broadcaster. Music submissions for review shows.',                      '','','Normal'],
  ['RAD005','MDR Kultur',             'Germany',       'Pop/Indie/Culture',          'https://www.mdr.de/mdr-kultur',      '',                                                'Mid-Germany public radio. Submit releases to music editorial.',                         '','','Normal'],
  // UK
  ['RAD006','BBC Introducing',        'UK',            'Unsigned/Emerging',          'https://www.bbc.co.uk/music/introducing','https://www.bbc.co.uk/music/introducing/about','Upload via BBC Sounds. Most important UK radio route for unsigned artists.',            '','','High'],
  ['RAD007','NTS Radio',             'UK',            'Indie/Experimental',          'https://www.nts.live',               'https://www.nts.live/contact',                    'Influential London/Manchester online radio. DM or email shows directly.',               '','','High'],
  ['RAD008','Worldwide FM',          'UK',            'Global/Indie',               'https://worldwidefm.net',             'https://worldwidefm.net/contact',                 'Global sounds, indie-friendly. Email music@worldwidefm.net',                           '','','Normal'],
  // USA
  ['RAD009','KEXP Seattle',          'USA',           'Indie/Alternative',          'https://www.kexp.org',               'https://www.kexp.org/music-submission/',           'Most influential US indie station. Official submission portal.',                        '','','High'],
  ['RAD010','KCRW Los Angeles',      'USA',           'Indie/World/Pop',            'https://www.kcrw.com',               'https://www.kcrw.com/music/music-submissions',    'LA tastemaker. Submit via SubmitHub or music@kcrw.com.',                               '','','High'],
  ['RAD011','The Current (MPR)',      'USA',           'Indie/Alternative',          'https://www.thecurrent.org',         'https://www.thecurrent.org/music-submissions',    'Minnesota Public Radio, national reach. Direct submission portal.',                    '','','High'],
  ['RAD012','WFUV New York',         'USA',           'Indie/Rock/Pop',             'https://wfuv.org',                   'https://wfuv.org/content/music-submissions',      'NYC public radio. Important indie/Americana outlet.',                                  '','','Normal'],
  ['RAD013','KUTX Austin',           'USA',           'Indie/Alt/Singer-songwriter', 'https://kutx.org',                  'https://kutx.org/music-submission/',              'Austin public radio. SXSW-aligned, strong indie reach.',                               '','','Normal'],
  // Belgium / Netherlands
  ['RAD014','Studio Brussel',        'Belgium',       'Pop/Indie',                  'https://stubru.be',                  'https://stubru.be/contact',                       'Belgiums leading indie station. Key for Benelux reach.',                              '','','High'],
  ['RAD015','Klara (VRT)',           'Belgium',       'Culture/Indie',              'https://klara.be',                   'https://klara.be/contact',                        'Flemish cultural public broadcaster.',                                                 '','','Normal'],
  ['RAD016','3voor12 (VPRO)',        'Netherlands',   'Indie/Alternative',          'https://3voor12.vpro.nl',            'https://3voor12.vpro.nl/contact.html',            'Netherlands indie institution. Online + radio. Submit releases.',                       '','','High'],
  // France
  ['RAD017','France Inter',          'France',        'Pop/Indie/Culture',          'https://www.radiofrance.fr/franceinter','',                                              'Major French public broadcaster. Festival tie-ins, live sessions.',                     '','','Normal'],
  // International
  ['RAD018','Radio Paradise',        'International', 'Eclectic/Indie',             'https://www.radioparadise.com',      'https://www.radioparadise.com/cms/info.php?topic=musician_submission','US-based global audience. Artist submission form available.','','','Normal'],
  ['RAD019','Fritz Unsigned',         'Germany',       'Indie/Pop',                  'https://bands.fritz.de',             'unsigned@fritz.de',                               'Fritz Unsigned — unsigned/independent Berlin artists. Upload via bands.fritz.de or email unsigned@fritz.de. Sunday 8–10pm slot.', '','','High'],
];

const RADIO_HEADERS = ['ID','Station','Country','Format','URL','Submit URL','Notes','Last Submitted','Last Result','Priority'];

function getRadioSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(RADIO_TAB);
}

// Run once to create + seed the Radio tab
function setupRadioTab() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let tab = ss.getSheetByName(RADIO_TAB);
  if (!tab) {
    tab = ss.insertSheet(RADIO_TAB);
    Logger.log('Radio tab created.');
  } else {
    Logger.log('Radio tab already exists — seeding missing stations only.');
  }

  // Ensure header row
  if (tab.getLastRow() === 0) {
    tab.appendRow(RADIO_HEADERS);
  }

  // Get existing IDs to avoid duplicates
  const existing = tab.getDataRange().getValues();
  const existingIDs = new Set(existing.slice(1).map(r => String(r[0]).trim()));

  RADIO_SEED_DATA.forEach(row => {
    if (!existingIDs.has(row[0])) {
      tab.appendRow(row);
      existingIDs.add(row[0]);
    }
  });

  Logger.log('Radio tab setup complete. Rows: ' + (tab.getLastRow() - 1));
}

// GET ?action=listRadio — returns all radio stations
function listRadio() {
  const tab = getRadioSheet();
  if (!tab) return { error: 'Radio tab not found. Run setupRadioTab() first.', stations: [] };

  const data = tab.getDataRange().getValues();
  if (data.length < 2) return { stations: [] };

  const headers = headersFromRow(data[0]);
  const stations = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = String(row[i] !== null && row[i] !== undefined ? row[i] : ''); });
    return obj;
  });
  return { stations };
}

// GET ?action=updateRadio&id=RAD001&field=Last+Submitted&value=2026-05-09
function updateRadio(id, field, value) {
  if (!id || !field) return { error: 'id and field are required' };
  const tab = getRadioSheet();
  if (!tab) return { error: 'Radio tab not found' };

  const data    = tab.getDataRange().getValues();
  const headers = headersFromRow(data[0]);
  const idCol   = headers.indexOf('ID');
  const fldCol  = headers.indexOf(field);
  const updCol  = headers.indexOf('Last Submitted');

  if (idCol < 0)  return { error: 'ID column not found' };
  if (fldCol < 0) return { error: 'Field not found: ' + field };

  const today = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd');

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idCol]).trim() === String(id).trim()) {
      tab.getRange(r + 1, fldCol + 1).setValue(value || '');
      // Auto-update Last Submitted when marking result
      if (field === 'Last Result' && updCol >= 0) {
        tab.getRange(r + 1, updCol + 1).setValue(today);
      }
      return { success: true, id, field, value };
    }
  }
  return { error: 'Radio ID not found: ' + id };
}


// ── Seed all sources as watching rows ────────────────────────
// Run seedAllSources() ONCE to pre-populate the sheet with all
// scanner sources as "watching" rows — no keyword detection needed.
// Existing URLs are skipped, so it's safe to re-run.

function seedAllSources() {
  const sheet   = getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = headersFromRow(data[0]);
  const existing = data.slice(1).map(r => rowToObject(headers, r));
  const existingURLs = new Set(existing.map(o => o.URL));

  const today = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd');
  let added = 0;

  SCAN_SOURCES.forEach(src => {
    if (existingURLs.has(src.url)) return;
    const row = headers.map(h => {
      switch (h) {
        case 'ID':           return nextId(sheet, headers);
        case 'Title':        return src.name + ' (watching)';
        case 'Organization': return src.name;
        case 'Type':         return src.type || 'festival';
        case 'Status':       return 'watching';
        case 'Deadline':     return 'rolling';
        case 'Amount':       return '';
        case 'Region':       return 'International';
        case 'Source':       return 'Seeded';
        case 'URL':          return src.url;
        case 'Priority':     return src.priority || 'Normal';
        case 'Notes':        return 'Auto-seeded. Check site for open call dates.';
        case 'Date Added':   return today;
        case 'Date Updated': return today;
        case 'Applied':      return 'FALSE';
        default:             return '';
      }
    });
    sheet.appendRow(row);
    existingURLs.add(src.url);
    added++;
  });

  Logger.log('Seeded ' + added + ' new watching rows.');
  return { added };
}

// ── Weekly scan trigger ───────────────────────────────────────
// Run setupWeeklyTrigger() ONCE to replace the monthly scan with
// a weekly scan every Monday at 8am Berlin time.

function setupWeeklyTrigger() {
  // Remove old monthly and weekly triggers
  ScriptApp.getProjectTriggers()
    .filter(t => ['runMonthlyScan','runWeeklyScan'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runWeeklyScan')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  Logger.log('Weekly trigger set: every Monday at 8am Berlin.');
}

function runWeeklyScan() {
  const scanResult    = scanSources();
  const monitorResult = monitorSources();
  archiveOverdue();
  Logger.log('Weekly scan: ' + JSON.stringify(scanResult));
  Logger.log('Monitor: ' + JSON.stringify(monitorResult));
}

// ── Archive overdue opportunities ─────────────────────────────
// Marks any open/watching opportunity as 'closed' if its deadline
// passed more than 14 days ago. Runs automatically as part of the
// weekly scan, or can be triggered manually from the editor.

function archiveOverdue() {
  const sheet   = getSheet();
  const data    = sheet.getDataRange().getValues();
  const headers = headersFromRow(data[0]);
  const today   = new Date(); today.setHours(0,0,0,0);
  const cutoff  = new Date(today); cutoff.setDate(today.getDate() - 14);

  let archived = 0;
  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const obj    = rowToObject(headers, row);
    const status = (obj.Status || '').toLowerCase();
    if (!['open','watching'].includes(status)) continue;
    const dl = obj.Deadline ? new Date(obj.Deadline) : null;
    if (!dl || isNaN(dl) || dl > cutoff) continue;
    // Deadline was 14+ days ago — mark closed
    const statusCol  = headers.indexOf('Status');
    const updatedCol = headers.indexOf('Date Updated');
    if (statusCol < 0) continue;
    sheet.getRange(i + 1, statusCol + 1).setValue('closed');
    if (updatedCol >= 0) sheet.getRange(i + 1, updatedCol + 1).setValue(Utilities.formatDate(today, 'Europe/Berlin', 'yyyy-MM-dd'));
    archived++;
  }
  Logger.log('archiveOverdue: ' + archived + ' opportunities closed.');
  return archived;
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
