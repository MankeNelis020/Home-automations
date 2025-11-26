/***** Generieke Taken – Sheets ↔ Calendar Sync (incrementeel + veilig) *****/

const TIMEZONE = 'Europe/Amsterdam';
const SHEET_NAME = 'Taken-overzicht';      // <- tabnaam met taken
const ADMIN_SHEET_NAME = 'Administratie';

// Exacte kopregels zoals in je sheet:
const HEADER = {
  categorie:       'Categorie',
  omschrijving:    'Omschrijving',
  deadline:        'Deadline',
  starttijd:       'Starttijd',
  toegewezen:      'Toegewezen aan',
  duurUren:        'Duur (uren)',
  prioriteit:      'Prioriteit',
  budget:          'Budget',
  status:          'Status',
  doneFlag:        'Done?',
  cancelledFlag:   'Geannuleerd?',
  notities:        'Notities',
  eventId:         'CalEventId',
  reminderDays:    'ReminderOffsetDays',
  calendarIdUsed:  'CalendarIdUsed' // wordt door script gevuld
};

// Optioneel: forceer alles naar één agenda door hier een ID te zetten.
// Leeg laten = per persoon agenda uit Administratie + anders default calendar.
const CALENDAR_ID = '';

// Fallback-duur als er geen "Duur (uren)" staat
const DEFAULT_TIMED_EVENT_DURATION_MIN = 60;

// Als status deze woorden bevat, mag een event ook weg
const DELETE_KEYWORDS = ['geannuleerd'];

// Standaard reminders (dagen vooraf). We doen 14d en 7d + dag zelf (0 min).
const DEFAULT_REMINDERS_DAYS = [14, 7];

// Fallback starttijd (als Deadline geen tijdcomponent heeft)
const FALLBACK_START_HOUR = 20; // 20:00 uur

// Visuele sync-feedback in kolom (bijv. "Sync?")
const SYNC_STATUS_COLUMN = 'I';

// Max agenda-operaties per batch-run (syncAll / 5-min trigger)
const MAX_CALENDAR_OPS_PER_RUN = 50;

// Cache voor administratie-config
let ADMIN_CACHE = null;

/* ========== Menu ========== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const manualMenu = ui.createMenu('Handmatige acties')
    .addItem('Sync alle taken (batch)', 'syncAll')
    .addItem('Push geselecteerde rij', 'syncActiveRow')
    .addItem('Verwijder gekoppelde event (rij)', 'deleteActiveRowEvent');

  const automationMenu = ui.createMenu('Automatisering')
    .addItem('Herstel autosync (on-edit trigger)', 'ensureAutoSyncTrigger')
    .addSubMenu(
      ui.createMenu('5-min autosync')
        .addItem('Inschakelen', 'enableFiveMinSync')
        .addItem('Uitschakelen', 'disableFiveMinSync')
    );

  ui.createMenu('Taken')
    .addSubMenu(manualMenu)
    .addSeparator()
    .addSubMenu(automationMenu)
    .addSeparator()
    .addItem('Taken genereren uit selectie (GPT)', 'generateTasksFromSelection')
    .addToUi();
}

/* ========== INSTALLABLE onEdit (maak trigger in UI of via menu) ========== */
function onEditInstalled(e) {
  try {
    if (!e || !e.range || !e.source) return;
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== SHEET_NAME) return;

    // Check: zit er een relevante kolom in het bewerkte bereik?
    const relevantHeaders = {
      [HEADER.deadline]:      true,
      [HEADER.starttijd]:     true,
      [HEADER.status]:        true,
      [HEADER.omschrijving]:  true,
      [HEADER.categorie]:     true,
      [HEADER.reminderDays]:  true,
      [HEADER.toegewezen]:    true,
      [HEADER.duurUren]:      true,
      [HEADER.doneFlag]:      true,
      [HEADER.cancelledFlag]: true,
      [HEADER.notities]:      true
    };
    const firstCol = e.range.getColumn();
    const lastCol  = firstCol + e.range.getNumColumns() - 1;

    let hasRelevant = false;
    for (let c = firstCol; c <= lastCol; c++) {
      const h = String(sheet.getRange(1, c).getValue()).trim();
      if (relevantHeaders[h]) { hasRelevant = true; break; }
    }
    if (!hasRelevant) return;

    const headers = getHeaderMap(sheet);
    if (!headers) return;

    // Sync elke data-rij in het bewerkte bereik
    const firstRow = Math.max(2, e.range.getRow()); // vanaf rij 2
    const lastRow  = e.range.getRow() + e.range.getNumRows() - 1;

    for (let r = firstRow; r <= lastRow; r++) {
      const rowVals = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getValues()[0];
      const isEmpty = rowVals.every(v => v === '' || v === null);
      if (isEmpty) continue;

      const result = syncRow(sheet, r, headers, { source: 'edit' }); // 'synced' | 'deleted' | 'skipped'
      if (result === 'synced') {
        setSyncStatus(sheet, r, '✅ Gesynchroniseerd op ' +
          Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
      } else if (result === 'deleted') {
        setSyncStatus(sheet, r, '❌ Verwijderd uit agenda op ' +
          Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
      }
    }
  } catch (err) {
    Logger.log(err);
    try { SpreadsheetApp.getActive().toast('Fout bij onEditInstalled: ' + err); } catch (_) {}
  }
}

/* ========== Handmatige acties ========== */
function syncAll() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" niet gevonden.`);
  const headers = getHeaderMap(sheet);
  if (!headers) throw new Error('Kolomkoppen niet gevonden of onvolledig.');

  const lastRow = sheet.getLastRow();
  let ops = 0;

  for (let r = 2; r <= lastRow; r++) {
    if (ops >= MAX_CALENDAR_OPS_PER_RUN) {
      SpreadsheetApp.getActive().toast('Batchlimiet bereikt, volgende taken bij volgende run.');
      break;
    }

    const rowVals = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (rowVals.every(v => v === '' || v === null)) continue;

    const result = syncRow(sheet, r, headers, { source: 'batch' });
    if (result === 'synced' || result === 'deleted') {
      ops++;
    }

    if (result === 'synced') {
      setSyncStatus(sheet, r, '✅ Gesynchroniseerd op ' +
        Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
    } else if (result === 'deleted') {
      setSyncStatus(sheet, r, '❌ Verwijderd uit agenda op ' +
        Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
    }
  }
}

function syncActiveRow() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const headers = getHeaderMap(sheet);
  if (!headers) {
    SpreadsheetApp.getUi().alert('Kolomkoppen niet gevonden of onvolledig.');
    return;
  }
  const r = sheet.getActiveRange().getRow();
  if (r <= 1) {
    SpreadsheetApp.getUi().alert('Selecteer een data-rij (vanaf rij 2).');
    return;
  }
  const result = syncRow(sheet, r, headers, { source: 'manual' });
  if (result === 'synced') {
    setSyncStatus(sheet, r, '✅ Gesynchroniseerd op ' +
      Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
  } else if (result === 'deleted') {
    setSyncStatus(sheet, r, '❌ Verwijderd uit agenda op ' +
      Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
  }
}

function deleteActiveRowEvent() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const headers = getHeaderMap(sheet);
  if (!headers) {
    SpreadsheetApp.getUi().alert('Kolomkoppen niet gevonden of onvolledig.');
    return;
  }
  const r = sheet.getActiveRange().getRow();
  if (r <= 1) {
    SpreadsheetApp.getUi().alert('Selecteer een data-rij (vanaf rij 2).');
    return;
  }

  const toegewezen = getCell(sheet, r, headers[HEADER.toegewezen]);
  const calendarIdUsed = headers[HEADER.calendarIdUsed]
    ? getCell(sheet, r, headers[HEADER.calendarIdUsed])
    : '';

  const cal = getCalendarForRow(toegewezen, calendarIdUsed);
  const eventId = getCell(sheet, r, headers[HEADER.eventId]);

  if (eventId) {
    try {
      const ev = cal.getEventById(normalizeEventId(eventId));
      if (ev) ev.deleteEvent();
    } catch (_) {}
    setCell(sheet, r, headers[HEADER.eventId], '');
  }
  setSyncStatus(sheet, r, '❌ Verwijderd uit agenda op ' +
    Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
  SpreadsheetApp.getUi().alert('Event verwijderd (indien aanwezig) en koppeling gewist.');
}

/* ========== Triggers helpers (menu) ========== */
function ensureAutoSyncTrigger() {
  const f = 'onEditInstalled';
  const triggers = ScriptApp.getProjectTriggers();
  const has = triggers.some(t =>
    t.getHandlerFunction() === f &&
    t.getEventType() === ScriptApp.EventType.ON_EDIT
  );
  if (!has) {
    ScriptApp.newTrigger(f)
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
  }
  SpreadsheetApp.getActive().toast('Autosync (on-edit) staat AAN.');
}

function enableFiveMinSync() {
  const f = 'syncAll';
  const has = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === f &&
    t.getEventType() === ScriptApp.EventType.TIME_DRIVEN
  );
  if (!has) {
    ScriptApp.newTrigger(f)
      .timeBased()
      .everyMinutes(5)
      .create();
  }
  SpreadsheetApp.getActive().toast('5-min autosync staat AAN.');
}

function disableFiveMinSync() {
  const f = 'syncAll';
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === f &&
        t.getEventType() === ScriptApp.EventType.TIME_DRIVEN) {
      ScriptApp.deleteTrigger(t);
    }
  });
  SpreadsheetApp.getActive().toast('5-min autosync is UIT.');
}

/* ========== Kern-sync per rij (incrementeel) ========== */
// return: 'synced' | 'deleted' | 'skipped'
function syncRow(sheet, row, headers, options) {
  options = options || {};
  const source = options.source || 'manual'; // 'edit' | 'batch' | 'manual'

  let categorie    = getCell(sheet, row, headers[HEADER.categorie]);
  const omschrijving = getCell(sheet, row, headers[HEADER.omschrijving]);
  const deadlineVal  = getCell(sheet, row, headers[HEADER.deadline]);
  const startTimeVal = getCell(sheet, row, headers[HEADER.starttijd]);
  const toegewezen   = getCell(sheet, row, headers[HEADER.toegewezen]);
  const duurUrenRaw  = getCell(sheet, row, headers[HEADER.duurUren]);
  const status       = getCell(sheet, row, headers[HEADER.status]);
  const notities     = getCell(sheet, row, headers[HEADER.notities]);
  const existingEventId = getCell(sheet, row, headers[HEADER.eventId]);
  const reminderDaysOverride = getCell(sheet, row, headers[HEADER.reminderDays]); // optioneel
  const doneVal      = getCell(sheet, row, headers[HEADER.doneFlag]);
  const cancelledVal = getCell(sheet, row, headers[HEADER.cancelledFlag]);
  const calendarIdUsed = headers[HEADER.calendarIdUsed]
    ? getCell(sheet, row, headers[HEADER.calendarIdUsed])
    : '';

  // Categorie automatisch invullen als leeg, o.b.v. omschrijving + Administratie
  if (!categorie) {
    categorie = inferCategoryFromDescription_(omschrijving);
    if (categorie) {
      setCell(sheet, row, headers[HEADER.categorie], categorie);
    }
  }

  // Duur in uren → minuten
  let duurUren = parseFloat(duurUrenRaw);
  if (!isFinite(duurUren) || duurUren <= 0) {
    duurUren = DEFAULT_TIMED_EVENT_DURATION_MIN / 60;
  }
  const duurMinuten = Math.max(1, Math.round(duurUren * 60));

  const cal = getCalendarForRow(toegewezen, calendarIdUsed);

  const statusStr = (status || '').toString().toLowerCase();
  const isDone = doneVal === true || String(doneVal).toLowerCase() === 'true';
  const isCancelled = cancelledVal === true || String(cancelledVal).toLowerCase() === 'true';

  const shouldDelete =
    !deadlineVal ||
    isDone ||
    isCancelled ||
    DELETE_KEYWORDS.some(k => statusStr.includes(k));

  // Als deze taak niet meer in de agenda hoeft te staan
  if (shouldDelete) {
    if (existingEventId) {
      try {
        const ev = cal.getEventById(normalizeEventId(existingEventId));
        if (ev) ev.deleteEvent();
      } catch (_) {}
      setCell(sheet, row, headers[HEADER.eventId], '');
    }
    return 'deleted';
  }

  // Geen omschrijving of toegewezen -> we maken niets
  if (!omschrijving || !toegewezen || !deadlineVal) {
    return 'skipped';
  }

  // Titel en description
  const title = buildTitle(categorie, omschrijving, toegewezen);
  const description = buildDescription({
    categorie, omschrijving, toegewezen, status, notities
  });

  // Deadline + duur -> start & end (timed event)
  const { start, end } = parseDeadlineWithDuration_(deadlineVal, startTimeVal, duurMinuten);

  // Reminders: 14d en 7d + dag zelf (0 min) of override
  const customDay = Number(reminderDaysOverride);
  const reminderDays = (isFinite(customDay) && customDay > 0)
    ? [customDay] // per rij override
    : DEFAULT_REMINDERS_DAYS.slice();

  let ev = existingEventId ? getEventSafe(cal, existingEventId) : null;

  // EventId aanwezig maar event bestaat niet meer → id wissen
  if (existingEventId && !ev) {
    setCell(sheet, row, headers[HEADER.eventId], '');
  }

  // Geen event-object → proberen bestaande te vinden op basis van titel + tijd
  if (!ev) {
    ev = findExistingEventByTitleAndTime_(cal, title, start, end);
    if (ev) {
      setCell(sheet, row, headers[HEADER.eventId], ev.getId());
    }
  }

  // Als we nog steeds geen event hebben: nieuw aanmaken (sheet → calendar)
  if (!ev) {
    ev = cal.createEvent(title, start, end, { description });
    setReminders(ev, reminderDays, /*includeSameDay*/ true);
    setCell(sheet, row, headers[HEADER.eventId], ev.getId());
    if (headers[HEADER.calendarIdUsed]) {
      try { setCell(sheet, row, headers[HEADER.calendarIdUsed], cal.getId()); } catch (_) {}
    }
    return 'synced';
  }

  // Vanaf hier: we hebben een bestaand event (met of zonder eerdere koppeling)

  // Sheet → Calendar logica bij onEdit / handmatig / batch
  const calStart = ev.getStartTime();
  const calEnd   = ev.getEndTime();
  const calTitle = ev.getTitle();
  const calDesc  = ev.getDescription() || '';

  const timesEqual = datesNearlyEqual_(calStart, start) && datesNearlyEqual_(calEnd, end);
  const titleEqual = calTitle === title;
  const descEqual  = calDesc === description;

  if (timesEqual && titleEqual && descEqual) {
    // Niets gewijzigd -> geen calendar-operatie
    if (headers[HEADER.calendarIdUsed]) {
      try { setCell(sheet, row, headers[HEADER.calendarIdUsed], cal.getId()); } catch (_) {}
    }
    return 'skipped';
  }

  // Er is wél wat veranderd in relevante velden → updaten
  try {
    if (!timesEqual) ev.setTime(start, end);
    if (!titleEqual) ev.setTitle(title);
    if (!descEqual) ev.setDescription(description);
    setReminders(ev, reminderDays, /*includeSameDay*/ true);
  } catch (e) {
    // Fallback: bij rare fouten opnieuw aanmaken
    try { ev.deleteEvent(); } catch (_) {}
    const newEv = cal.createEvent(title, start, end, { description });
    setReminders(newEv, reminderDays, /*includeSameDay*/ true);
    ev = newEv;
    setCell(sheet, row, headers[HEADER.eventId], ev.getId());
  }

  if (headers[HEADER.calendarIdUsed]) {
    try { setCell(sheet, row, headers[HEADER.calendarIdUsed], cal.getId()); } catch (_) {}
  }

  return 'synced';
}

/* ========== Helpers ========== */

function getHeaderMap(sheet) {
  const lastCol = sheet.getLastColumn();
  const headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const map = {};
  Object.values(HEADER).forEach(name => {
    const idx = headerVals.findIndex(h => h.trim() === name);
    if (idx !== -1) map[name] = idx + 1;
  });
  // Vereist minimaal deze kolommen:
  const required = [HEADER.omschrijving, HEADER.deadline, HEADER.eventId];
  const ok = required.every(k => map[k]);
  return ok ? map : null;
}

function getCell(sheet, row, col) {
  if (!col) return '';
  return sheet.getRange(row, col).getValue();
}

function setCell(sheet, row, col, val) {
  if (!col) return;
  sheet.getRange(row, col).setValue(val);
}

function buildTitle(categorie, omschrijving, toegewezen) {
  const c = (categorie || '').toString().trim();
  const o = (omschrijving || '').toString().trim();
  const t = (toegewezen || '').toString().trim();

  let base;
  if (c && o) base = `${c} – ${o}`;
  else base = o || c || 'Taak';

  if (t) return `[${t}] ${base}`;
  return base;
}

function buildDescription({ categorie, omschrijving, toegewezen, status, notities }) {
  const parts = [];
  if (omschrijving) parts.push(`Omschrijving: ${omschrijving}`);
  if (categorie) parts.push(`Categorie: ${categorie}`);
  if (toegewezen) parts.push(`Toegewezen aan: ${toegewezen}`);
  if (status) parts.push(`Status: ${status}`);
  if (notities) parts.push(`Notities: ${notities}`);
  parts.push(`Bron: ${SpreadsheetApp.getActiveSpreadsheet().getName()}`);
  parts.push(`Link: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`);
  return parts.join('\n');
}

// Deadline (datum) + Starttijd (tijd) + duur → start & end (timed event)
function parseDeadlineWithDuration_(dateVal, timeVal, duurMinuten) {
  if (!dateVal) throw new Error('Geen deadline-datum.');

  const d = (dateVal instanceof Date) ? new Date(dateVal) : new Date(dateVal);
  if (isNaN(d)) throw new Error('Ongeldige deadline-datum.');

  // Default tijd
  let hours = FALLBACK_START_HOUR;
  let minutes = 0;

  // Als er een tijd is ingevuld, neem die over
  if (timeVal) {
    const t = (timeVal instanceof Date) ? new Date(timeVal) : new Date(timeVal);
    if (!isNaN(t)) {
      hours = t.getHours();
      minutes = t.getMinutes();
    }
  }

  d.setHours(hours, minutes, 0, 0);

  const start = new Date(d);
  const end   = new Date(start.getTime() + duurMinuten * 60000);

  return { start, end };
}

function normalizeEventId(eventId) {
  return eventId.includes('@') ? `${eventId}` : `${eventId}@google.com`;
}

function getEventSafe(cal, id) {
  try {
    return cal.getEventById(normalizeEventId(id));
  } catch (_) {
    return null;
  }
}

// Zet reminders op dagen vooraf + dag-zelf popup
function setReminders(ev, reminderDaysArray, includeSameDay) {
  try { ev.removeAllReminders(); } catch (_) {}
  (reminderDaysArray || []).forEach(d => {
    const mins = Math.max(0, Math.round(Number(d) * 24 * 60));
    if (isFinite(mins) && mins > 0) {
      try { ev.addPopupReminder(mins); } catch (_) {}
    }
  });
  if (includeSameDay) {
    try { ev.addPopupReminder(0); } catch (_) {}
  }
}

/* ========== Event matching & Calendar → Sheet helpers ========== */

// Zoek een event met zelfde titel en (ongeveer) dezelfde start/eindtijd
function findExistingEventByTitleAndTime_(cal, title, start, end) {
  const windowMs = 5 * 60 * 1000; // 5 minuten tolerantie
  const searchStart = new Date(start.getTime() - windowMs);
  const searchEnd   = new Date(end.getTime() + windowMs);
  let events;
  try {
    events = cal.getEvents(searchStart, searchEnd, { search: title });
  } catch (e) {
    Logger.log('Fout bij getEvents voor matching: ' + e);
    return null;
  }
  for (const ev of events) {
    const evStart = ev.getStartTime();
    const evEnd   = ev.getEndTime();
    if (ev.getTitle() === title &&
        Math.abs(evStart.getTime() - start.getTime()) <= windowMs &&
        Math.abs(evEnd.getTime() - end.getTime()) <= windowMs) {
      return ev;
    }
  }
  return null;
}

// Bij batch-run: als event-tijden in kalender afwijken, stuur die terug naar de sheet.
// Return: true als sheet is aangepast, anders false.
function updateSheetFromCalendarIfNeeded_(sheet, row, headers, ev, deadlineVal, startTimeVal, duurUren) {
  const calStart = ev.getStartTime();
  const calEnd   = ev.getEndTime();

  // Bepaal "sheetverwachting" op basis van huidige waarden
  const duurMinutenSheet = Math.round((parseFloat(duurUren) || 1) * 60);
  const desired = parseDeadlineWithDuration_(deadlineVal, startTimeVal, duurMinutenSheet);
  const desiredStart = desired.start;
  const desiredEnd   = desired.end;

  const timesEqual = datesNearlyEqual_(calStart, desiredStart) && datesNearlyEqual_(calEnd, desiredEnd);
  if (timesEqual) {
    return false; // niets te corrigeren
  }

  // Kalender wordt hier leidend voor Deadline, Starttijd en Duur(uren)
  const newDeadline = new Date(calStart.getFullYear(), calStart.getMonth(), calStart.getDate());
  const newTimeOnly = new Date(0, 0, 0, calStart.getHours(), calStart.getMinutes(), 0);
  const durationHours = (calEnd.getTime() - calStart.getTime()) / (1000 * 60 * 60);
  const roundedHours = Math.round(durationHours * 100) / 100;

  setCell(sheet, row, headers[HEADER.deadline], newDeadline);
  setCell(sheet, row, headers[HEADER.starttijd], newTimeOnly);
  setCell(sheet, row, headers[HEADER.duurUren], roundedHours);

  return true;
}

function datesNearlyEqual_(a, b) {
  if (!a || !b) return false;
  const diff = Math.abs(a.getTime() - b.getTime());
  return diff <= 60 * 1000; // 1 minuut tolerantie
}

/* ========== Administratie-config ========== */

function loadAdminCache_() {
  if (ADMIN_CACHE) return ADMIN_CACHE;

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(ADMIN_SHEET_NAME);
  const cache = { assigneeToCal: {}, categories: [] };

  if (!sheet) {
    ADMIN_CACHE = cache;
    return cache;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    ADMIN_CACHE = cache;
    return cache;
  }

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  for (let r = 0; r < values.length; r++) {
    const row = values[r].map(v => (v || '').toString().trim());
    if (!row.some(Boolean)) continue;

    // Blok 1: Persoon / CalendarId
    if (row.includes('Persoon') && row.includes('CalendarId')) {
      const idxPers = row.indexOf('Persoon');
      const idxCal  = row.indexOf('CalendarId');
      for (let i = r + 1; i < values.length; i++) {
        const rr = values[i];
        if (rr.every(v => v === '' || v === null)) break;
        const p = (rr[idxPers] || '').toString().trim();
        const c = (rr[idxCal] || '').toString().trim();
        if (p && c) cache.assigneeToCal[p.toLowerCase()] = c;
      }
    }

    // Blok 2: Categorie / Regex
    if (row.includes('Categorie') && row.includes('Regex')) {
      const idxCat = row.indexOf('Categorie');
      const idxRx  = row.indexOf('Regex');
      for (let i = r + 1; i < values.length; i++) {
        const rr = values[i];
        if (rr.every(v => v === '' || v === null)) break;
        const cat = (rr[idxCat] || '').toString().trim();
        const rx  = (rr[idxRx] || '').toString().trim();
        if (cat && rx) {
          try {
            cache.categories.push({ cat, regex: new RegExp(rx, 'i') });
          } catch (e) {
            Logger.log('Ongeldige regex in Administratie: ' + rx + ' (' + e + ')');
          }
        }
      }
    }
  }

  ADMIN_CACHE = cache;
  return cache;
}

function getCalendarForRow(toegewezen, calendarIdUsed) {
  // 1) Hard override via CALENDAR_ID
  if (CALENDAR_ID) {
    return CalendarApp.getCalendarById(CALENDAR_ID);
  }

  // 2) Als we eerder al een calendarId hebben gebruikt voor deze rij: die weer pakken
  if (calendarIdUsed) {
    try {
      const cal = CalendarApp.getCalendarById(calendarIdUsed);
      if (cal) return cal;
    } catch (_) {}
  }

  // 3) Anders via Administratie → Persoon / CalendarId
  const name = (toegewezen || '').toString().trim();
  if (name) {
    const admin = loadAdminCache_();
    const id = admin.assigneeToCal[name.toLowerCase()];
    if (id) {
      try {
        const cal = CalendarApp.getCalendarById(id);
        if (cal) return cal;
      } catch (_) {}
    }
  }

  // 4) Fallback: default calendar
  return CalendarApp.getDefaultCalendar();
}

function inferCategoryFromDescription_(description) {
  const text = (description || '').toString();
  if (!text) return '';
  const admin = loadAdminCache_();
  for (const entry of admin.categories) {
    try {
      if (entry.regex.test(text)) return entry.cat;
    } catch (_) {}
  }
  return '';
}

/* ========== Snelle testfunctie ========== */
function _testCreateEvent() {
  const cal = getCalendarForRow('', '');
  const start = new Date(new Date().getTime() + 60 * 60 * 1000);
  const end   = new Date(start.getTime() + 30 * 60 * 1000);
  const ev = cal.createEvent('Test – Taken', start, end, { description: 'Test event' });
  Logger.log('Gemaakt: ' + ev.getId());
  SpreadsheetApp.getActive().toast('Testevent aangemaakt voor over 1 uur.');
}

/* ========== Visuele sync-feedback helper ========== */
function setSyncStatus(sheet, row, message) {
  try {
    const colIndex = sheet.getRange(`${SYNC_STATUS_COLUMN}1`).getColumn();
    sheet.getRange(row, colIndex).setValue(message);
  } catch (err) {
    Logger.log('Kon sync-status niet schrijven: ' + err);
  }
}

/**
 * Leest de geselecteerde cel (projectomschrijving),
 * stuurt deze naar OpenAI en voegt taken toe aan Taken-overzicht.
 */
function generateTasksFromSelection() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "' + SHEET_NAME + '" niet gevonden.');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const headers = getHeaderMap(sheet);
  if (!headers) {
    ui.alert('Kolomkoppen niet gevonden of onvolledig.');
    return;
  }

  const activeRow = sheet.getActiveCell().getRow();
  if (activeRow <= 1) {
    ui.alert('Selecteer een data-rij (vanaf rij 2).');
    return;
  }

  // Projectomschrijving + einddeadline uit de geselecteerde rij
  const projectTextVal = getCell(sheet, activeRow, headers[HEADER.omschrijving]);
  const projectDeadlineVal = getCell(sheet, activeRow, headers[HEADER.deadline]);

  const projectText = projectTextVal ? projectTextVal.toString().trim() : '';
  if (!projectText) {
    ui.alert('In deze rij staat geen Omschrijving. Vul eerst een projectomschrijving in.');
    return;
  }

  // Einddeadline als YYYY-MM-DD string (mag leeg blijven)
  const projectDeadlineStr = projectDeadlineVal
    ? formatDateForApi_(projectDeadlineVal)
    : '';

  try {
    const tasks = callOpenAIForTasks_(projectText, projectDeadlineStr);
    if (!tasks || !tasks.length) {
      ui.alert('Geen taken teruggekregen van GPT.');
      return;
    }

    // -------- Preview bouwen --------
    let previewLines = tasks.map((t, idx) => {
      const d   = t.deadline_date || '';
      const s   = t.starttijd || '';
      const who = t.toegewezen_aan || '';
      const when = [d, s].filter(Boolean).join(' ');
      return (idx + 1) + '. [' + (who || '?') + '] ' +
             (t.omschrijving || '') +
             (when ? ' (' + when + ')' : '');
    });

    const MAX_LINES = 10;
    let extraNote = '';
    if (previewLines.length > MAX_LINES) {
      extraNote = '\n\n… en nog ' + (previewLines.length - MAX_LINES) + ' extra taken.';
      previewLines = previewLines.slice(0, MAX_LINES);
    }

    const response = ui.alert(
      'GPT-taken controleren',
      'Ik heb ' + tasks.length + ' taken voorgesteld:\n\n' +
        previewLines.join('\n') +
        extraNote +
        '\n\nWil je deze taken toevoegen en de geselecteerde rij overschrijven?',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      ui.alert('Geen taken toegevoegd.');
      return;
    }

    // -------- Ruimte maken: geselecteerde rij + extra rijen eronder --------
    const numTasks = tasks.length;
    if (numTasks > 1) {
      sheet.insertRowsAfter(activeRow, numTasks - 1);
    }

    const rows = [];
    tasks.forEach(t => {
      rows.push([
        '',                                   // Categorie (regex vult later)
        t.omschrijving || '',
        parseDateForSheet_(t.deadline_date), // Deadline (datum)
        t.starttijd || '',                   // Starttijd (tijd)
        t.toegewezen_aan || '',
        t.duur_uren || '',
        t.prioriteit || '',
        t.budget || '',
        t.status || 'Te doen',
        '',                                   // Done?
        '',                                   // Geannuleerd?
        t.notities || '',
        '',                                   // CalEventId
        '',                                   // ReminderOffsetDays
        '',                                   // CalendarIdUsed
        '',                                   // Maand (formule)
        ''                                    // Kwartaal (formule)
      ]);
    });

    // Schrijf de taken vanaf de geselecteerde rij naar beneden
    sheet.getRange(activeRow, 1, rows.length, rows[0].length).setValues(rows);
    ss.toast('GPT-taken toegevoegd: ' + tasks.length);

  } catch (e) {
    Logger.log('Fout in generateTasksFromSelection: ' + e);
    SpreadsheetApp.getUi().alert('Fout bij GPT-taakgeneratie: ' + e);
  }
}

/**
 * Roept OpenAI Chat Completions aan en retourneert een array van taak-objecten.
 */
function callOpenAIForTasks_(projectText) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('OPENAI_API_KEY niet ingesteld in Script properties.');
  }

  const url = 'https://api.openai.com/v1/chat/completions';

  const systemPrompt =
    'Je bent een assistent die Nederlandse huishoudelijke, baby-, klus- en cadeau-projecten ' +
    'opknipt in concrete taken voor een Google Sheet. ' +
    'Je output MOET puur JSON zijn (zonder uitleg) volgens het gevraagde schema.';

  const userPrompt =
    'Projectbeschrijving:\n' + projectText + '\n\n' +
    'Genereer een lijst taken als JSON-array. Elke taak is een object met deze velden:\n' +
    '- omschrijving: korte, duidelijke taakbeschrijving (in het Nederlands)\n' +
    '- deadline_date: datum in formaat YYYY-MM-DD (logische chronologische planning, eventueel weekenden gebruiken)\n' +
    '- starttijd: tijd in 24u-formaat HH:MM (grote klussen overdag, kleinere ook ’s avonds mogelijk)\n' +
    '- toegewezen_aan: één van: "Niel", "Kelsey", "Nielsley"\n' +
    '- duur_uren: duur in uren (mag decimalen, bijv. 1.5)\n' +
    '- prioriteit: bijvoorbeeld "Hoog", "Midden" of "Laag"\n' +
    '- budget: leeg laten of een globale schatting als dat logisch is\n' +
    '- status: standaard "Te doen"\n' +
    '- notities: optionele extra context of afhankelijkheden\n\n' +
    'BELANGRIJK:\n' +
    '- Geef ALLEEN de JSON-array terug, zonder markdown, zonder extra tekst.\n';

  const payload = {
    model: 'gpt-4.1-mini',
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ],
    temperature: 0.3
  };

  const options = {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error('OpenAI API-fout: HTTP ' + status + ' - ' + response.getContentText());
  }

  const data = JSON.parse(response.getContentText());
  const text = data.choices[0].message.content.trim();

  let tasks;
  try {
    tasks = JSON.parse(text);
  } catch (e) {
    throw new Error('Kon JSON van GPT niet parsen. Ruwe output: ' + text);
  }

  if (!Array.isArray(tasks)) {
    throw new Error('GPT-output is geen array.');
  }
  return tasks;
}

/**
 * Hulpfunctie: string "YYYY-MM-DD" → Date object voor in de sheet.
 */
function parseDateForSheet_(s) {
  if (!s) return '';
  if (s instanceof Date) return s;
  const str = s.toString().trim();
  const parts = str.split('-');
  if (parts.length === 3) {
    const year = Number(parts[0]);
    const month = Number(parts[1]) - 1; // 0-based
    const day = Number(parts[2]);
    if (isFinite(year) && isFinite(month) && isFinite(day)) {
      return new Date(year, month, day);
    }
  }
  // fallback: laat Google het zelf proberen
  return new Date(str);
}

/**
 * Formatteer datum naar YYYY-MM-DD string (voor API).
 */
function formatDateForApi_(d) {
  if (!(d instanceof Date)) d = new Date(d);
  const y = d.getFullYear();
  const m = ('0' + (d.getMonth() + 1)).slice(-2);
  const day = ('0' + d.getDate()).slice(-2);
  return `${y}-${m}-${day}`;
}
