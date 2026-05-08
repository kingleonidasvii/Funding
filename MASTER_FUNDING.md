# lovelee — MASTER FUNDING SYSTEM
_Created: May 8, 2026. Author: Claude (Senior Artist Manager)._
_Load alongside MASTER_CONTEXT.md at the start of any session involving funding, festivals, showcases, or booking._
_This is the single source of truth for the opportunity tracking system._

---

## SYSTEM OVERVIEW

A centralized opportunity tracker covering funding grants, festival showcases, open calls, residencies, and booking targets. Built across three layers:

1. Google Sheet (lovelee_opportunities) — the database. Single source of truth.
2. Web app hosted on GitHub Pages (github.com/kingleonidasvii) — visual interface. Reads from the Sheet.
3. Claude in Chrome scheduled workflow — monthly scraping routine. Writes to the Sheet automatically.

Manual intake via the app's Add Opportunity form for anything found ad hoc (corporate grants, US foundations, tips from Beratung).

Gmail label structure organizes all correspondence. Claude in Chrome processes incoming newsletters and extracts new entries.

---

## GMAIL LABEL STRUCTURE

Create these labels in Gmail exactly as written. Nested labels use the slash syntax.

```
opportunities/funding/open
opportunities/funding/submitted
opportunities/funding/watching
opportunities/funding/closed
opportunities/funding/ineligible
opportunities/festivals/open
opportunities/festivals/submitted
opportunities/festivals/watching
opportunities/festivals/closed
opportunities/booking/prospect
opportunities/booking/pitched
opportunities/booking/confirmed
opportunities/booking/passed
newsletters/musicpool
newsletters/musicboard
newsletters/initiative-musik
newsletters/kreativkultur
newsletters/german-music-export
outreach/radio
outreach/press
outreach/playlists
```

### Gmail filter rules to create

For each sender below, auto-apply the matching label and skip inbox:

| Sender domain / keyword | Label |
|---|---|
| musicpoolberlin.net | newsletters/musicpool |
| musicboard-berlin.de | newsletters/musicboard |
| initiative-musik.de | newsletters/initiative-musik |
| kreativkultur.berlin | newsletters/kreativkultur |
| german-music-export.de | newsletters/german-music-export |
| Subject contains "Förderung" OR "funding" OR "open call" | opportunities/funding/open (review manually) |

---

## GOOGLE SHEET SCHEMA

Sheet name: lovelee_opportunities

### Tab 1: Opportunities

| Column | Field | Notes |
|---|---|---|
| A | ID | Auto-increment. Format: OPP-001 |
| B | Title | Name of the opportunity |
| C | Organization | Who runs it |
| D | Type | funding / festival / showcase / residency / booking / open-call |
| E | Status | open / watching / submitted / closed / ineligible |
| F | Deadline | ISO date YYYY-MM-DD or "ongoing" or "TBD" |
| G | Amount | Max funding amount as string e.g. "€8,750" or "up to €3,500" |
| H | Region | DE / EU / NL / US / international |
| I | Eligibility | One-line note on whether lovelee qualifies |
| J | URL | Direct link to application or info page |
| K | Notes | Any context, requirements, contacts |
| L | Source | How it was found: musicpool / musicboard / scraped / manual / gmail |
| M | Date Added | ISO date |
| N | Date Updated | ISO date |
| O | Applied | TRUE / FALSE |
| P | Result | Pending / Approved / Rejected / Withdrawn |

### Tab 2: Sources

Tracks every URL being scraped and when it was last checked.

| Column | Field |
|---|---|
| A | Source Name |
| B | URL |
| C | Region |
| D | Scrape Frequency |
| E | Last Checked |
| F | Status (active/paused) |
| G | Notes |

### Tab 3: Outreach Pipeline

Mirrors MASTER_OUTREACH.md logic for radio, press, and playlist targets. Same structure as Opportunities tab but focused on contacts not deadlines.

---

## SCRAPING SOURCE LIST

These are the URLs the monthly Claude in Chrome workflow scrapes. Organized by region and reliability.

### Germany (high priority — scrape monthly)

| Source | URL | Notes |
|---|---|---|
| Music Pool Berlin Padlet | https://padlet.com/Eileen_MPB/infoboard-music-pool-berlin-qmv8zjbkhhi0ag7x | Login required. Use saved Musicpool credentials. |
| Musicboard Berlin | https://www.musicboard-berlin.de/en/funding/ | Public. Reliable. |
| kreativkultur.berlin | https://www.kreativkultur.berlin/en/funding-database/ | Public. Comprehensive. Covers Senate + EU programs. |
| Initiative Musik | https://www.initiative-musik.de/en/funding/ | Public. |
| Berlin Music Commission | https://www.berlin-music-commission.de | Public. Watch for Music Ambassador and Listen to Berlin. |
| Backstage PRO | https://www.backstagepro.de/thema/foerderung | Public. Aggregates DE funding news. |
| German Music Export | https://www.german-music-export.de | Public. WOMEX, showcase delegations. |
| GEMA Kulturförderung | https://www.gema.de/musikschaffende/foerderung/ | Public. Basisförderung rounds. |
| Musikfonds | https://www.musikfonds.de/foerderprogramme/ | Public. Quarterly deadlines. |

### Netherlands / EU (scrape quarterly)

| Source | URL | Notes |
|---|---|---|
| Eurosonic Noorderslag | https://www.eurosonic-noorderslag.nl/en/play-esns/ | Showcase application. Sept 1 deadline for Jan 2027. |
| Creative Europe | https://culture.ec.europa.eu/calls | EU flagship. Complex eligibility. Watch for music-relevant calls. |
| SHAPE+ | https://shapeplus.eu/open-calls/ | Experimental/electronic focus. Check eligibility. |
| Liveurope | https://liveurope.eu | Venue network booking emerging EU acts. Not a grant — booking pipeline. |
| On The Move | https://on-the-move.org/grants | Mobility funding for EU artists. Highly relevant for touring. |
| Fonds Podiumkunsten (NL) | https://www.fondspodiumkunsten.nl/en/ | Dutch performing arts fund. |
| Stichting Popcoalitie (NL) | https://www.popcoalitie.nl | Dutch pop music sector body. Open calls and export support. |

### International festivals and showcases (scrape quarterly)

| Source | URL | Typical Deadline | Notes |
|---|---|---|---|
| Reeperbahn Festival | https://www.reeperbahnfestival.com/en/artists/artist-submission | Spring | Hamburg. Right level now. |
| c/o pop | https://www.copop.de | Spring | Cologne. Indie/pop. |
| Waves Vienna | https://www.wavesvienna.com/apply | May | Vienna. |
| M for Montreal | https://www.mformotnreal.com | April | Canada. Industry-facing. |
| Iceland Airwaves | https://icelandairwaves.is | April/May | Reykjavik. |
| MIL Lisboa | https://www.millisboa.com | June | Portugal. |
| Mastering the Music Business | https://www.masteringthemusicbusiness.ro | Rolling | Bucharest. |
| SXSW | https://www.sxsw.com/apply | August | Austin. Long game — 2027 target. |
| Folk Alliance International | https://folkalliance.org/showcases | October | US. Worth monitoring post-Violet. |

### US funding (manual intake only — no reliable public aggregator)

These require manual checking and use of the app's Add Opportunity form. Set a quarterly calendar reminder.

| Organization | URL | Notes |
|---|---|---|
| ASCAP Foundation | https://www.ascapfoundation.org/grants | Multiple grant programs. GEMA equivalent. |
| BMI Foundation | https://bmifoundation.org | Grants for composers. |
| American Music Abroad | https://americanmusicabroad.americanensembles.org | US State Dept. Requires US connection. |
| New Music USA | https://newmusicusa.org/grants | Project grants. Genre-flexible. |
| Meet the Composer | https://meetthecomposer.org | US touring and commissioning. |
| Spotify Sound Up / Creator Equity | https://soundup.spotify.com | Watch for open rounds. |
| YouTube Artist Development | https://artists.youtube.com | Rolling. Watch for program announcements. |

---

## CURRENTLY KNOWN OPPORTUNITIES (as of May 8, 2026)

Pre-loaded into the Sheet on first run.

| Title | Org | Type | Status | Deadline | Amount | Notes |
|---|---|---|---|---|---|---|
| Round 73 Popmusik | Initiative Musik | funding | submitted | 2026-04-30 | €8,750 | K53677/2026. Awaiting jury decision. |
| Support Tour Funding 2026 | Musicboard Berlin | funding | open | 2026-11-30 | up to €3,500 | Requires confirmed support slot with established headliner. 3+ shows outside Berlin. Directly relevant for Dilla touring. Apply at least 3 weeks before tour start. |
| Music Ambassador Programm 2026 | Berlin Music Commission | funding | open | 2026-05-28 | travel costs | For Berlin-based music professionals traveling abroad. Deadline 28 May 23:59. Could apply to international showcase travel. |
| GEMA Basisförderung | GEMA | funding | open | 2026-05-17 | €2,000 | Applications open 04.05–17.05. First 400 approved get funded. Must be GEMA member (confirmed). Apply immediately. |
| Labelförderung Round 2 | Musicboard Berlin | funding | ineligible | 2026-05-15 | variable | Requires GVL label code. Not applicable as solo artist without a label. |
| International Live and Showcase Funding | Initiative Musik | funding | watching | rolling | variable | Updated program. Focus on showcase performances and support tours abroad. Lump-sum funding available. German application only. |
| International Export Funding | Initiative Musik | funding | watching | 2-3 rounds/year | up to €20,000 | For music companies. Not yet eligible — watch post-Violet for label or company formation. |
| WOMEX 2026 Delegate | German Music Export | event | watching | 2026-06-08 | subsidised ticket | Las Palmas, Oct 21-25. Apply as delegate for German delegation. Discounted registration + possible travel subsidy. |
| Twisted Trees Festival 2026 | Kero Productions | festival | open | 2026-05-08 | performance | TODAY. Free open-air. Send 10-15 min of music to info@keroproductions.com. Subject: "Twisted Trees Submission". |
| AltShift Festival 2026 | AltShift | festival | open | 2026-05-10 | food/camping/reimbursement | Harzgerode, July 20-25. Degrowth festival. Send application. |
| Reeperbahn Festival 2026 | Reeperbahn | festival | open | rolling | performance | Hamburg, Sept 16-19. Application open. Submit now. |
| Waves Vienna 2026 | Waves Vienna | festival | open | 2026-05-29 | performance | Vienna, Oct 1-3. Apply via website. |
| MIL Lisboa 2026 | MIL | festival | open | 2026-06-10 | performance | Lisbon, Oct 7-10. Fill out form on millisboa.com. |
| Eurosonic Noorderslag 2027 | ESNS | festival | watching | 2026-09-01 | performance | Groningen, Jan 2027. Deadline September. Priority target for Violet campaign window. |
| Fête de la Musique Berlin | Musicboard Berlin | gig | open | 2026-06-21 | exposure | June 21, Berlin city-wide. Apply via matchmaking tool on fetedelamusique.de. Free visibility gig. |
| Listen to Berlin 2026 | Berlin Music Commission | open-call | closed | 2026-04-28 | compilation/showcase | Deadline passed. Watch for 2027 round. Opens March. |
| Tune In Studio Open Call | Berlin Senate / Tune In | open-call | open | 2026-05-31 | free studio time Oct 2026-Jul 2027 | Competitive. Jury selection. Apply via berlin.de link. Highly relevant for Violet production. |
| Musikfonds Kleine Projektförderung | Musikfonds | funding | watching | 2026-05-29 | up to €3,000 | Next deadline May 29 for July-Sept events. Focus on rural/non-metropolitan now — Berlin projects rarely funded. Watch for exceptions. |
| KOMPASS Soloselbständige | EU / BMAS | funding | watching | rolling | training costs | ESF+ funded. Covers further training for solo self-employed. Relevant for professional development costs. |

---

## IMMEDIATE ACTIONS (priority order as of May 8, 2026)

1. GEMA Basisförderung — deadline May 17. Window closes fast. Apply now at gema.de/kulturfoerderung-online. First 400 approved get €2,000. You are a GEMA member.

2. Twisted Trees Festival — deadline TODAY May 8. Send 10-15 min audio to info@keroproductions.com. Subject: "Twisted Trees Submission". Free gig, good Berlin exposure.

3. AltShift Festival — deadline May 10. Low-barrier. Fill in form at altshiftfestival.org/open-calls.

4. Music Ambassador Programme — deadline May 28. Review eligibility for any international showcase travel planned for H2 2026.

5. Tune In Studio — deadline May 31. Free professional studio time Oct 2026-Jul 2027. Directly useful for Violet production. Apply via berlin.de.

6. Reeperbahn Festival — rolling deadline. Submit artist application now at reeperbahnfestival.com.

7. Waves Vienna — deadline May 29. Apply at wavesvienna.com.

8. WOMEX delegate — deadline June 8. Research whether this makes sense for where lovelee is in October 2026.

---

## CLAUDE IN CHROME WORKFLOW SPEC

### Workflow name: lovelee monthly opportunity scan

### Trigger: First of every month, 9:00 AM Berlin time

### Steps (record once, Claude repeats):

1. Open kreativkultur.berlin/en/funding-database — extract all listings with deadlines in next 90 days. Write to Sheet Tab 1.
2. Open musicboard-berlin.de/en/funding — check for any program deadline updates. Update existing rows.
3. Open initiative-musik.de/en/funding — check for new rounds. Update existing rows.
4. Open berlin-music-commission.de — check for new open calls.
5. Log into Musicpool Berlin Padlet (saved credentials) — extract all new listings since last visit. Parse with Claude API. Write new rows to Sheet.
6. Check Gmail label newsletters/* — for any unprocessed newsletter emails, extract opportunities and write to Sheet.
7. Send Telegram message to @wren_lovelee_bot: "Monthly scan complete. [N] new opportunities added. [M] deadlines in next 30 days: [list]."

### Claude API parsing prompt (used in step 5 and 6):

```
You are parsing music industry content for a Berlin-based indie pop artist named lovelee (Leon Byvanck). Extract all funding opportunities, festival submissions, open calls, residencies, and showcase opportunities. For each one extract: title, organization, type (funding/festival/showcase/residency/open-call/gig), deadline (ISO date or "ongoing"), amount (if mentioned), URL (if mentioned), eligibility notes (is this relevant for a solo Berlin-based indie pop artist?), and a one-line summary. Return as JSON array. Skip jobs, workshops, and irrelevant content (classical, jazz-only, children's programs, heavy metal).
```

---

## APP BUILD SPEC (for Claude Code)

### Repository: github.com/kingleonidasvii

### Tech stack: HTML + vanilla JS + Google Sheets API v4 (read-only public Sheet)

### Pages:

**Page 1: Opportunity Tracker**
- Filter bar: All / Funding / Festival / Showcase / Open Call / Gig
- Filter bar: All / Open / Watching / Submitted / Closed
- Filter bar: All / DE / EU / NL / US / International
- Deadline urgency sort: red (0-14 days) / amber (15-60 days) / green (60+ days)
- Card per opportunity: title, org, type badge, status badge, deadline with urgency color, amount, one-line note, link button
- Add Opportunity button: opens form, submits to Sheet via Google Forms embed or Apps Script webhook
- Search bar: filters cards by keyword

**Page 2: Outreach Pipeline**
- Mirrors existing Wren outreach logic
- Contacts table: name, station/publication, type, status, last contacted, notes
- Add Contact button

**Page 3: Sources**
- Lists all scraping sources with last-checked date and status
- Manual refresh button (triggers Claude in Chrome workflow or shows instructions)

### Data connection:
- Google Sheet must be published as public (read-only) for the app to read without auth
- Write operations (Add Opportunity, Add Contact) go via a Google Apps Script web app endpoint
- Apps Script endpoint URL stored as a constant in the JS

### Design:
- Match lovelee_OS_dashboard.jsx aesthetic: dark background (#0a0a0f), violet accents (#7c3aed), monospace for data, serif for headers
- Mobile responsive
- No frameworks — vanilla JS only to keep it lightweight

---

## ELIGIBILITY RULES (apply when triaging new opportunities)

Lee is eligible for programs requiring:
- Berlin residence: YES
- German GEMA membership: YES (Mitgliedsnummer 2113538)
- KSK registration: YES
- Solo artist: YES
- Pop music genre: YES
- Professional touring experience: YES (Dilla, 150+ concerts)
- DIY/independent: YES (no label, recordJet distribution)

Lee is NOT eligible for:
- Label funding (no GVL label code)
- Programs requiring classical/jazz/new music
- Programs requiring a GbR or registered company (currently)
- US-specific funds without a US address or collaborator
- Programs requiring enrollment in a university

Watch for future eligibility:
- International Export Funding (Initiative Musik) — requires music company. Post-Violet, consider GbR or UG formation.
- Label Funding (Musicboard) — if lovelee ever formalizes as a label structure.

---

## SESSION WORKFLOW

When a session involves funding or opportunities:

1. Load this file + MASTER_CONTEXT.md
2. Run "scan kreativkultur.berlin" to fetch fresh listings if more than 30 days since last scan
3. Check immediate actions list for anything with deadline within 14 days
4. Update this file with any new decisions, new opportunities found, or status changes
5. Flag anything over €200 as per guardrails

Trigger phrase: "funding scan" — immediately check all open deadlines and surface anything expiring within 30 days.

---

## SOURCES NOT YET INTEGRATED (research backlog)

These need manual investigation before adding to the scrape list:

- Hauptstadtkulturfonds (Berlin) — interdisciplinary. Check if pop music projects qualify.
- Fonds Darstellende Künste — performing arts fund. Check pop music eligibility.
- Norsk kulturråd (Norway) — Norwegian Arts Council. Worth monitoring for Nordic showcase connections.
- Pro Helvetia (Switzerland) — Swiss Arts Council. Export funding for international tours.
- British Council Germany — UK-Germany cultural exchange. Post-Brexit eligibility unclear.
- Goethe Institut — cultural diplomacy grants. Usually requires German institutional partner.
- Red Bull Music — dormant. Watch for revival.
- Spotify Sound Up — rolling. Watch for open rounds.
- YouTube Artist Development — watch for announcements.
- Candid / Foundation Center (US) — requires account. Quarterly manual check.

---

_Last updated: May 8, 2026_
_Next scheduled review: June 1, 2026 (monthly scan)_
_File maintained by: Claude (Senior Artist Manager)_
