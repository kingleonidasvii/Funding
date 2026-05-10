// ============================================================
// lovelee ‚Äî Opportunity Tracker ‚Äî Google Apps Script Web App
// ============================================================
//
// DEPLOYMENT INSTRUCTIONS
// -----------------------
// 1. Open the Google Sheet: https://docs.google.com/spreadsheets/d/1Si1_GF1DVPq7ZNfH5NPOxQaH6gJDtiDDbqTUYT7kQg4/edit
// 2. Extensions ‚Üí Apps Script
// 3. Delete any existing code. Paste the entire contents of this file.
// 4. Click Save (Cmd+S). Name the project "lovelee-opportunity-tracker".
// 5. Click Deploy ‚Üí New deployment ‚Üí gear icon ‚Üí Web app.
// 6. Execute as: Me | Who has access: Anyone ‚Üí Deploy ‚Üí Authorize ‚Üí Allow.
// 7. Copy the Web app URL (https://script.google.com/macros/s/AKfy.../exec)
// 8. Paste it into SCRIPT_URL in index.html on GitHub. Commit.
//
// AFTER CODE CHANGES: Deploy ‚Üí Manage deployments ‚Üí Edit ‚Üí New version ‚Üí Deploy
// URL stays the same ‚Äî no need to update index.html again.
//
// ONE-TIME SETUP (run these manually in the editor after first deploy):
//   Run ‚Üí setupPriorityColumn   ‚Äî adds Priority column to the Sheet
//   Run ‚Üí setupMonthlyTrigger   ‚Äî sets up automatic scan on 1st of month
//
// SHEET COLUMNS (Opportunities tab):
//   ID | Title | Organization | Type | Status | Deadline | Amount |
//   Region | Eligibility | URL | Notes | Source | Date Added |
//   Date Updated | Applied | Result | Priority
// ============================================================

const SHEET_ID          = '1Si1_GF1DVPq7ZNfH5NPOxQaH6gJDtiDDbqTUYT7kQg4';
const OPPORTUNITIES_TAB = 'Opportunities';
const RADIO_TAB         = 'Radio';

// ‚îÄ‚îÄ Helpers ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ

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
      obj[h] = '‚Ç¨' + v.toLocaleString('de-DE');
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

// ‚îÄ‚îÄ GET ‚Äî list all opportunities, run scan, or mark applied ‚îÄ‚îÄ

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'list';

    if (action === 'scan')         return corsResponse(scanSources());
    if (action === 'markApplied')  return corsResponse(markApplied(e.parameter.id));
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

// ‚îÄ‚îÄ POST ‚Äî add a new opportunity ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ

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

// ‚îÄ‚îÄ Mark Applied (called from app's "Mark Applied" button) ‚îÄ‚îÄ‚îÄ

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

// ‚îÄ‚îÄ One-time setup: add Priority column if missing ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ

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

// ‚îÄ‚îÄ Scan sources ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
// Fetches public source pages, detects active open calls by keyword,
// and appends new "watching" rows. Login-required sources (e.g.
// Music Group Berlin) should be handled by uploading their PDF
// in Claude chat ‚Äî Claude will extract and add entries via POST.

const SCAN_SOURCES = [
  // ‚îÄ‚îÄ Germany: Funding ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  { name:'Musicboard Berlin',           url:'https://www.musicboard-berlin.de/foerderung/',                   priority:'High',   type:'funding'  },
  { name:'Initiative Musik',            url:'https://www.initiative-musik.de/foerderung/',                   priority:'High',   type:'funding'  },
  { name:'kreativkultur.berlin',        url:'https://kreativkultur.berlin',                                  priority:'Normal', type:'funding'  },
  { name:'Berlin Music Commission',     url:'https://www.berlin-music-commission.de',                        priority:'Normal', type:'funding'  },
  { name:'Musikfonds',                  url:'https://www.musikfonds.de/foerderung/',                         priority:'Normal', type:'funding'  },
  { name:'GEMA Kulturf√∂rderung',        url:'https://www.gema.de/kulturfoerderung-online',                   priority:'Normal', type:'funding'  },
  { name:'German Music Export',         url:'https://www.german-music-export.de/en/funding/',                priority:'High',   type:'funding'  },
  { name:'Backstage PRO',               url:'https://www.backstage.de/ratgeber/foerderung/',                 priority:'Normal', type:'funding'  },
  { name:'Fonds Darstellende K√ºnste',   url:'https://www.fonds-daku.de/foerderung/',                         priority:'Normal', type:'funding'  },

  // ‚îÄ‚îÄ Germany: Festivals & Major Stages ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  // Note: reeperbahn-festival.de has TLS issues ‚Äî using HTTP fallback for scanner
  { name:'Reeperbahn Festival',         url:'http://www.reeperbahn-festival.de/en/',                         priority:'High',   type:'festival' },
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

  // ‚îÄ‚îÄ Germany: Radio & Media ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  { name:'Fritz (rbb)',                 url:'https://www.fritz.de/programm/aktionen/',                       priority:'High',   type:'radio'    },
  { name:'ByteFM',                      url:'https://www.byte.fm/features/',                                 priority:'High',   type:'radio'    },
  { name:'radioeins (rbb)',             url:'https://www.radioeins.de/musik/',                                priority:'Normal', type:'radio'    },
  { name:'Deutschlandfunk Kultur',      url:'https://www.deutschlandfunkkultur.de/musik/',                   priority:'Normal', type:'radio'    },
  { name:'MDR Kultur',                  url:'https://www.mdr.de/mdr-kultur/radio/index.html',                priority:'Normal', type:'radio'    },

  // ‚îÄ‚îÄ EU: Funding ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  { name:'Creative Europe',             url:'https://culture.ec.europa.eu/calls',                           priority:'High',   type:'funding'  },
  { name:'SHAPE+',                      url:'https://shapeplatform.eu/open-call/',                           priority:'High',   type:'open-call'},
  { name:'Liveurope',                   url:'https://liveurope.eu',                                          priority:'High',   type:'funding'  },
  { name:'Fonds Podiumkunsten NL',      url:'https://www.fondspodiumkunsten.nl/subsidies/',                  priority:'Normal', type:'funding'  },
  { name:'√ñsterreichischer Musikfonds', url:'https://www.musikfonds.at/foerderungen/',                       priority:'Normal', type:'funding'  },
  { name:'Pro Helvetia Switzerland',    url:'https://prohelvetia.ch/en/funding/',                            priority:'Normal', type:'funding'  },
  { name:'Music Finland Export',        url:'https://musicfinland.com/en/funding/',                          priority:'Normal', type:'funding'  },
  { name:'Music Norway Export',         url:'https://musicnorway.no/funding/',                               priority:'Normal', type:'funding'  },
  { name:'Swedish Arts Council',        url:'https://www.kulturradet.se/en/apply-for-grants/',               priority:'Normal', type:'funding'  },
  { name:'PRS Foundation UK',           url:'https://prsfoundation.com/funding-support/funding/',            priority:'High',   type:'funding'  },
  { name:'Help Musicians UK',           url:'https://www.helpmusicians.org.uk/get-support/funding',          priority:'Normal', type:'funding'  },
  { name:'Arts Council England',        url:'https://www.artscouncil.org.uk/projectgrants',                  priority:'Normal', type:'funding'  },
  { name:'CNM France',                  url:'https://cnm.fr/aides/',                                         priority:'High',   type:'funding'  },

  // ‚îÄ‚îÄ Netherlands ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  { name:'Eurosonic Noorderslag',       url:'https://www.eurosonic-noorderslag.nl/showcases/',               priority:'High',   type:'festival' },
  { name:'Lowlands Festival',           url:'https://www.lowlands.nl',                                       priority:'High',   type:'festival' },
  { name:'Down the Rabbit Hole',        url:'https://www.downtherabbithole.nl',                              priority:'High',   type:'festival' },
  { name:'Best Kept Secret NL',         url:'https://www.bestkeptsecret.nl',                                 priority:'High',   type:'festival' },

  // ‚îÄ‚îÄ Belgium ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  { name:'Pukkelpop Belgium',           url:'https://www.pukkelpop.be',                                      priority:'High',   type:'festival' },
  { name:'Rock Werchter Belgium',       url:'https://www.rockwerchter.be',                                   priority:'High',   type:'festival' },
  { name:'Dour Festival Belgium',       url:'https://www.dourfestival.eu',                                   priority:'Normal', type:'festival' },
  { name:'Glimps Festival Ghent',       url:'https://glimps.be',                                             priority:'Normal', type:'festival' },

  // ‚îÄ‚îÄ Austria ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  { name:'Waves Vienna',                url:'https://wavesvienna.com',                                       priority:'High',   type:'festival' },

  // ‚îÄ‚îÄ France ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  { name:'MAMA Paris',                  url:'https://www.mama-event.com/en/',                                priority:'High',   type:'festival' },
  { name:'Trans Musicales Rennes',      url:'https://www.lestrans.com/les-trans-musicales/',                 priority:'High',   type:'festival' },
  { name:'La Route du Rock',            url:'https://www.laroutedurock.com',                                 priority:'High',   type:'festival' },
  { name:'Printemps de Bourges',        url:'https://www.printemps-bourges.com',                             priority:'High',   type:'festival' },
  { name:'Nuits Sonores Lyon',          url:'https://www.nuits-sonores.com',                                 priority:'High',   type:'festival' },
  { name:'Rock en Seine Paris',         url:'https://www.rockenseine.com',                                   priority:'High',   type:'festival' },
  { name:'Pitchfork Music Fest Paris',  url:'https://www.pitchforkparis.com',                                priority:'High',   type:'festival' },

  // ‚îÄ‚îÄ UK ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
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

  // ‚îÄ‚îÄ Ireland ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  { name:'Other Voices Ireland',        url:'https://www.othervoices.ie',                                    priority:'High',   type:'festival' },
  { name:'Hard Working Class Heroes',   url:'https://www.hwch.net',                                          priority:'High',   type:'festival' },
  { name:'Electric Picnic Ireland',     url:'https://www.electricpicnic.ie',                                 priority:'High',   type:'festival' },
  { name:'First Music Contact Ireland', url:'https://fmc.ie/funding/',                                       priority:'Normal', type:'funding'  },

  // ‚îÄ‚îÄ Scandinavia ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  { name:'Iceland Airwaves',            url:'https://icelandairwaves.is',                                    priority:'High',   type:'festival' },
  { name:'by:Larm Oslo',                url:'https://bylarm.no/en/',                                         priority:'High',   type:'festival' },
  { name:'Way Out West Gothenburg',     url:'https://www.wayoutwest.se',                                     priority:'High',   type:'festival' },
  { name:'Roskilde Festival',           url:'https://www.roskilde-festival.dk/en/',                          priority:'High',   type:'festival' },
  { name:'√òya Festival Oslo',           url:'https://www.oyafestivalen.no',                                  priority:'High',   type:'festival' },
  { name:'Flow Festival Helsinki',      url:'https://www.flowfestival.com',                                  priority:'High',   type:'festival' },

  // ‚îÄ‚îÄ Southern Europe ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  { name:'MIL Lisboa',                  url:'https://mil-lisboa.pt',                                         priority:'High',   type:'festival' },
  { name:'NOS Alive Portugal',          url:'https://nosalive.com',                                          priority:'High',   type:'festival' },
  { name:'Primavera Sound Barcelona',   url:'https://www.primaverasound.com',                                 priority:'High',   type:'festival' },
  { name:'S√≥nar Barcelona',             url:'https://sonar.es/en/',                                          priority:'High',   type:'festival' },
  { name:'Mad Cool Madrid',             url:'https://www.madcoolfestival.es',                                priority:'High',   type:'festival' },
  { name:'Bilbao BBK Live',             url:'https://www.bilbaobblive.com',                                  priority:'High',   type:'festival' },
  { name:'Linecheck Milan',             url:'https://www.linecheckmusic.com/en/',                            priority:'High',   type:'festival' },
  { name:'MI AMI Festival Milan',       url:'https://miamifestival.it',                                      priority:'Normal', type:'festival' },
  { name:'Medimex Bari',                url:'https://www.medimex.it',                                        priority:'Normal', type:'festival' },
  { name:'Sziget Festival Budapest',    url:'https://szigetfestival.com',                                    priority:'High',   type:'festival' },

  // ‚îÄ‚îÄ Switzerland ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  { name:'Montreux Jazz Festival',      url:'https://www.montreuxjazzfestival.com',                          priority:'High',   type:'festival' },
  { name:'Pal√©o Festival Nyon',         url:'https://www.paleo.ch',                                          priority:'High',   type:'festival' },

  // ‚îÄ‚îÄ USA / Canada ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  { name:'SXSW Austin',                 url:'https://www.sxsw.com/music/applications/',                      priority:'High',   type:'festival' },
  { name:'M for Montreal',              url:'https://mformontreal.com/en/artists/',                          priority:'High',   type:'festival' },
  { name:'Folk Alliance International', url:'https://www.folkalliance.org/conference/',                      priority:'Normal', type:'festival' },
  { name:'CMW Toronto',                 url:'https://cmw.net/showcase-applications/',                        priority:'Normal', type:'festival' },

  // ‚îÄ‚îÄ Global industry ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
  { name:'Womex',                       url:'https://www.womex.com/virtual/womex/apply',                     priority:'High',   type:'festival' },
  { name:'Womad',                       url:'https://womad.co.uk',                                           priority:'High',   type:'festival' },
];

const KEYWORDS = [
  // English (UK / IE / US / AU)
  'open call','open for','apply now','applications open','submit','submission',
  'deadline','grant','funding round','pitch your music','new artists',
  'emerging artists','unsigned','demo','apply here','applications now open',
  'call for entries','call for artists','now accepting','accepting submissions',
  'apply for','open applications','music grant','touring grant','artist fund',
  // German
  'jetzt bewerben','bewerbung','f√∂rderantrag','einreichung','einreichen',
  'frist','f√∂rderung','stipendium','bewerbungsschluss','ausschreibung',
  'projektf√∂rderung','musikf√∂rderung','reisekostenf√∂rderung',
  // Dutch
  'aanvragen','subsidie','aanmelding','open voor','aanmelden','indienen',
  'subsidieregeling','aanvraagronde','muzieksubsidie','open oproep',
  // French
  'candidater','candidature','appel √†','dossier de candidature','soumettre',
  'bourse','financement','date limite','appel ouvert','appel √† projets',
  'd√©p√¥t de candidature','appel √† candidatures',
  // Spanish
  'convocatoria','solicitud','inscripci√≥n','plazo','subvenci√≥n','beca',
  'abrir convocatoria','presentar candidatura','fecha l√≠mite','open call',
  'artistas emergentes','enviar demo','formulario de inscripci√≥n',
  // Italian
  'bando','candidatura','iscrizione','scadenza','contributo','borsa',
  'aperto alle candidature','chiama artisti','presentare domanda',
  'bando aperto','sovvenzione','call for artists',
  // Portuguese
  'candidatura','inscri√ß√£o','prazo','bolsa','financiamento',
  'candidatar','submeter','chamada aberta','apoio',
  // Scandinavian (Swedish / Norwegian / Danish / Finnish)
  'ans√∂kan','s√∏knad','ans√∂k','tilmeld','haku','stipend',
  'open for s√∏knader','s√∂ka bidrag','musikstipendium','residens'
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
          case 'Title':        return src.name + ' ‚Äî Open Call (auto-scanned)';
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

// ‚îÄ‚îÄ Monthly auto-trigger ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
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

// ‚îÄ‚îÄ Radio tab ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ
// Columns: ID | Station | Country | Format | URL | Submit URL | Notes | Last Submitted | Last Result | Priority
//
// ONE-TIME SETUP: Run ‚Üí setupRadioTab() once to create the tab and seed the stations.
// Afterwards the Radio tab is a persistent store ‚Äî update it via the app or directly.

const RADIO_SEED_DATA = [
  // Germany
  ['RAD001','Fritz (rbb)',           'Germany',       'Indie/Pop/Alternative',      'https://www.fritz.de',               'https://www.fritz.de/service/kontakt.html',       'Berlin flagship. Live sessions, Lokalmatadoren slot, demo submissions via contact form.','','','High'],
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
  ['RAD019','Spotify Editorial',     'International', 'All genres',                 'https://artists.spotify.com',        'https://artists.spotify.com/help/article/pitch-a-song','Pitch via Spotify for Artists ‚â•7 days pre-release.',                            '','','High'],
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
    Logger.log('Radio tab already exists ‚Äî seeding missing stations only.');
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

// GET ?action=listRadio ‚Äî returns all radio stations
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

// ‚îÄ‚îÄ Test functions ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ‚îÄ

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
    Title:'TEST ‚Äî delete me', Organization:'Test', Type:'funding',
    Status:'watching', Deadline:'2026-12-31', Amount:'‚Ç¨0',
    Region:'Germany', Source:'test', Priority:'Low'
  })}};
  Logger.log(doPost(fake).getContent());
}
