const CONFIG = Object.freeze({
  CLUB_NAME: 'Debate Club',
  EMAIL_DOMAIN: 'example.edu',
  LOCALE: 'en-US',
  GRADES: [['9', 3], ['10', 2], ['11', 1], ['12', 0]],
  ASK_SECTION: true,
  UPPERCASE_LAST_NAME: false,
  SCHOOL_YEAR_START_MONTH: 8,
  DEFAULT_DURATION_MIN: 60,
  OPEN_EARLY_MIN: 5,
  GRACE_MIN: 5,
  MAX_MEETINGS: 40
});

const SHEET = Object.freeze({
  DASHBOARD: 'Dashboard',
  MEMBERS: 'Members',
  MEETINGS: 'Meetings',
  ATTENDANCE: 'Attendance',
  REPORT: 'Report',
  RESPONSES: 'Responses',
  REGISTRATIONS: 'Registrations',
  MARKS: 'Marks'
});

const LAYOUT = Object.freeze({
  FIRST_ROW: 2,
  FIRST_MEETING_COL: 5,
  STATUS_CELL: 'B12'
});

const PRESENT = 'PRESENT';
const ABSENT = 'ABSENT';
const DATE_FORMAT = 'yyyy-mm-dd';
const DATETIME_FORMAT = 'yyyy-mm-dd hh:mm:ss';

const REG_QUESTIONS = [
  { key: 'first', title: 'First name', type: 'text' },
  { key: 'last', title: 'Last name', type: 'text' },
  { key: 'id', title: 'Student ID', type: 'id' },
  { key: 'grade', title: 'Grade', type: 'grade' },
  { key: 'section', title: 'Section', type: 'section' }
];

let TIME_ZONE_ = null;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Attendance')
    .addItem('Reprocess responses', 'reprocessResponses')
    .addItem('Check form now', 'updateFormWindow')
    .addItem('Open / close registration', 'toggleRegistration')
    .addSeparator()
    .addItem('Run setup', 'setup')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the Google Sheet and run setup from Extensions > Apps Script.');
  if (CONFIG.EMAIL_DOMAIN === 'example.edu') {
    throw new Error('Edit CONFIG at the top of the script: set EMAIL_DOMAIN to your school email domain ' +
      '(or to an empty string to accept any Google account).');
  }
  const missing = Object.keys(SHEET).map(k => SHEET[k]).filter(name => !ss.getSheetByName(name));
  if (missing.length) {
    throw new Error('Missing sheet(s): ' + missing.join(', ') + '. Upload Club_Attendance_Template.xlsx to Google Drive, ' +
      'open it with Google Sheets, choose File > Save as Google Sheets, and run setup from that copy.');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SHEET_ID', ss.getId());
  prepareSheets_(ss);

  const warnings = [];
  const form = setupAttendanceForm_(props, warnings);
  const regForm = setupRegistrationForm_(props, warnings);

  for (const trigger of ScriptApp.getProjectTriggers()) {
    if (['recordAttendance', 'updateFormWindow', 'recordRegistration'].includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  ScriptApp.newTrigger('recordAttendance').forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger('recordRegistration').forForm(regForm).onFormSubmit().create();
  ScriptApp.newTrigger('updateFormWindow').timeBased().everyMinutes(5).create();

  writeFormLinks_(ss, form, regForm);
  updateFormWindow();

  const ctx = context_(ss);
  const active = ctx.members.filter(m => m.active).length;
  const noEmail = ctx.members.filter(m => m.active && !m.email).length;
  const duplicateIds = Object.keys(ctx.duplicateIds);
  if (noEmail) warnings.push(noEmail + ' active member(s) have no email.');
  if (duplicateIds.length) warnings.push('Duplicate student IDs in Members: ' + duplicateIds.join(', '));
  const all = warnings.concat(ctx.meetings.filter(m => m.error).map(m => m.error));
  notify_(
    'Setup complete. Both forms and their QR codes are on the Dashboard.\n\n' +
    'Required for both forms: open the form > Published > Responders > Manage > choose your organization > Save. ' +
    'Then test with a student account.\n\nActive members: ' + active +
    (all.length ? '\n\nWarnings:\n- ' + all.join('\n- ') : '')
  );
}

function setupAttendanceForm_(props, warnings) {
  let form = openForm_(props.getProperty('ATTENDANCE_FORM_ID'));
  if (!form) {
    form = FormApp.create(CONFIG.CLUB_NAME + ' · Attendance');
    form.addMultipleChoiceItem()
      .setTitle('I am at today\'s meeting.')
      .setChoiceValues(['Present'])
      .setRequired(true);
    props.setProperty('ATTENDANCE_FORM_ID', form.getId());
    props.deleteProperty('CODE_ITEM_ID');
  }
  form.setDescription('Sign in with your school account and submit during the meeting.');
  form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(true);
  tryRun_(() => form.setConfirmationMessage('Your attendance has been received.'));
  verifiedEmail_(form, warnings);
  codeItem_(form, props);
  publish_(form, warnings);
  form.setAcceptingResponses(false);
  tryRun_(() => form.setCustomClosedFormMessage('Attendance is closed. It opens only during club meetings.'));
  return form;
}

function updateFormWindow() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ATTENDANCE_FORM_ID')) throw new Error('Run setup first.');
  const ss = spreadsheet_();
  const form = attendanceForm_();
  let ctx = context_(ss);
  const now = new Date();
  const open = ctx.meetings.find(m => m.valid && now >= opensAt_(m) && now <= m.end) || null;
  const wasOpen = form.isAcceptingResponses();

  if (open) setCodeRequired_(form, props, !!open.code);
  if (wasOpen !== !!open) form.setAcceptingResponses(!!open);
  if (wasOpen && !open) props.setProperty('REPROCESS', '1');

  const lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      if (processPendingRegistrations_(ss, props)) ctx = context_(ss);
      const state = syncColumns_(ss, ctx, props);
      ensureRows_(ss, ctx.members.map(m => m.id));
      const ended = ctx.meetings.filter(m =>
        m.valid && now > acceptsUntil_(m) && state[m.index] && !state[m.index].closed);
      if (!open && (ended.length || props.getProperty('REPROCESS'))) {
        reprocess_(ss, form, ctx);
        if (ended.length) {
          markAbsent_(ss, ctx, ended);
          ended.forEach(m => { state[m.index].closed = true; });
          props.setProperty('COLUMNS', JSON.stringify(state));
        }
        props.deleteProperty('REPROCESS');
      }
    } finally {
      lock.releaseLock();
    }
  }
  writeStatus_(ss, ctx.meetings, open);
}

function recordAttendance(e) {
  if (!e || !e.response) throw new Error('This function runs automatically when the attendance form is submitted.');
  const props = PropertiesService.getScriptProperties();
  const ss = spreadsheet_();
  const ctx = context_(ss);
  const form = e.source || attendanceForm_();
  const r = readResponse_(e.response, findCodeItem_(form, props));

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(120000)) {
    props.setProperty('REPROCESS', '1');
    console.warn('Lock not acquired: ' + r.email);
    return;
  }
  try {
    const marks = ss.getSheetByName(SHEET.MARKS);
    const index = rowIndex_(marks).map;
    const result = evaluate_(r, ctx, (id, mi) => {
      const row = index[key_(id)];
      return !!row && marks.getRange(row, LAYOUT.FIRST_MEETING_COL + mi).getValue() === PRESENT;
    });
    if (result.status === 'ACCEPTED') apply_(ss, [{ id: result.member.id, mi: result.meeting.index, value: PRESENT }]);
    const sh = ss.getSheetByName(SHEET.RESPONSES);
    const row = Math.max(sh.getLastRow(), LAYOUT.FIRST_ROW - 1) + 1;
    fitRows_(sh, row);
    sh.getRange(row, 1, 1, 5).setValues([[r.ts, r.email, result.date, result.status, result.reason]]);
    sh.getRange(row, 1).setNumberFormat(DATETIME_FORMAT);
    sh.getRange(row, 3).setNumberFormat(DATE_FORMAT);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function reprocessResponses() {
  const ss = spreadsheet_();
  const form = attendanceForm_();
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  let count;
  try {
    count = reprocess_(ss, form, context_(ss));
    PropertiesService.getScriptProperties().deleteProperty('REPROCESS');
  } finally {
    lock.releaseLock();
  }
  notify_(count + ' response(s) reprocessed. Manually edited cells were left unchanged.');
}

function reprocess_(ss, form, ctx) {
  const codeItem = findCodeItem_(form, PropertiesService.getScriptProperties());
  const responses = form.getResponses().slice().sort((a, b) => a.getTimestamp() - b.getTimestamp());
  const accepted = {};
  const rows = [];
  const changes = [];
  for (const response of responses) {
    const r = readResponse_(response, codeItem);
    const result = evaluate_(r, ctx, (id, mi) => !!accepted[key_(id) + '|' + mi]);
    if (result.status === 'ACCEPTED') {
      accepted[key_(result.member.id) + '|' + result.meeting.index] = true;
      changes.push({ id: result.member.id, mi: result.meeting.index, value: PRESENT });
    }
    rows.push([r.ts, r.email, result.date, result.status, result.reason]);
  }
  apply_(ss, changes);

  const sh = ss.getSheetByName(SHEET.RESPONSES);
  const last = sh.getLastRow();
  if (last >= LAYOUT.FIRST_ROW) sh.getRange(LAYOUT.FIRST_ROW, 1, last - LAYOUT.FIRST_ROW + 1, 5).clearContent();
  if (rows.length) {
    fitRows_(sh, LAYOUT.FIRST_ROW + rows.length - 1);
    sh.getRange(LAYOUT.FIRST_ROW, 1, rows.length, 5).setValues(rows);
    sh.getRange(LAYOUT.FIRST_ROW, 1, rows.length, 1).setNumberFormat(DATETIME_FORMAT);
    sh.getRange(LAYOUT.FIRST_ROW, 3, rows.length, 1).setNumberFormat(DATE_FORMAT);
  }
  SpreadsheetApp.flush();
  return rows.length;
}

function markAbsent_(ss, ctx, meetings) {
  const changes = [];
  for (const m of meetings) {
    for (const member of ctx.members) {
      if (member.active && key_(member.id)) {
        changes.push({ id: member.id, mi: m.index, value: ABSENT, onlyIfUnset: true });
      }
    }
  }
  apply_(ss, changes);
}

function apply_(ss, changes) {
  if (!changes.length) return;
  const n = CONFIG.MAX_MEETINGS, C = LAYOUT.FIRST_MEETING_COL, F = LAYOUT.FIRST_ROW;
  ensureRows_(ss, changes.map(c => c.id));
  const marks = ss.getSheetByName(SHEET.MARKS), att = ss.getSheetByName(SHEET.ATTENDANCE);
  const mi = rowIndex_(marks), ai = rowIndex_(att);
  const marksLast = mi.next - 1, attLast = ai.next - 1;
  const markValues = marksLast >= F ? marks.getRange(F, C, marksLast - F + 1, n).getValues() : [];
  const attValues = attLast >= F ? att.getRange(F, C, attLast - F + 1, n).getValues() : [];

  const writes = {};
  let marksChanged = false;
  for (const c of changes) {
    const k = key_(c.id);
    const m = mi.map[k] - F, a = ai.map[k] - F;
    if (!(m >= 0) || !(a >= 0)) continue;
    const previous = String(markValues[m][c.mi]).trim();
    if (previous === c.value) continue;
    if (c.onlyIfUnset && previous !== '') continue;
    markValues[m][c.mi] = c.value;
    marksChanged = true;
    const current = String(attValues[a][c.mi]).trim();
    if (current === '' || current === previous) {
      attValues[a][c.mi] = c.value;
      (writes[c.value] = writes[c.value] || []).push(a1_(F + a, C + c.mi));
    }
  }
  if (marksChanged) marks.getRange(F, C, markValues.length, n).setValues(markValues);
  for (const value of Object.keys(writes)) att.getRangeList(writes[value]).setValue(value);
}

function ensureRows_(ss, ids) {
  const marks = ss.getSheetByName(SHEET.MARKS), att = ss.getSheetByName(SHEET.ATTENDANCE);
  const mi = rowIndex_(marks), ai = rowIndex_(att);
  const seen = {};
  for (const id of ids) {
    const k = key_(id);
    if (!k || seen[k]) continue;
    seen[k] = true;
    if (!ai.map[k]) {
      const r = ai.next++;
      fitRows_(att, r);
      att.getRange(r, 1).setValue(id);
      const f = attendanceFormulas_(r);
      att.getRange(r, 2, 1, 3).setFormulas([f.info]);
      att.getRange(r, LAYOUT.FIRST_MEETING_COL + CONFIG.MAX_MEETINGS, 1, 3).setFormulas([f.summary]);
      ai.map[k] = r;
    }
    if (!mi.map[k]) {
      const r = mi.next++;
      fitRows_(marks, r);
      marks.getRange(r, 1).setValue(id);
      mi.map[k] = r;
    }
  }
}

function rowIndex_(sh) {
  const map = {};
  let lastFilled = LAYOUT.FIRST_ROW - 1;
  const last = sh.getLastRow();
  if (last >= LAYOUT.FIRST_ROW) {
    sh.getRange(LAYOUT.FIRST_ROW, 1, last - LAYOUT.FIRST_ROW + 1, 1).getValues().forEach((x, i) => {
      const k = key_(x[0]);
      if (k) {
        if (!map[k]) map[k] = LAYOUT.FIRST_ROW + i;
        lastFilled = LAYOUT.FIRST_ROW + i;
      }
    });
  }
  return { map: map, next: lastFilled + 1 };
}

function attendanceFormulas_(r) {
  const first = LAYOUT.FIRST_MEETING_COL, n = CONFIG.MAX_MEETINGS;
  const a = col_(first), z = col_(first + n - 1), p = col_(first + n), q = col_(first + n + 1);
  const M = sheetRef_(SHEET.MEMBERS);
  const lookup = c => '=IF($A' + r + '="","",IFERROR(INDEX(' + M + '!$' + c + '$2:$' + c + ',MATCH($A' + r +
    ',' + M + '!$A$2:$A,0)),""))';
  return {
    info: [lookup('B'), lookup('C'), lookup('E')],
    summary: [
      '=IF($A' + r + '="","",COUNTIF($' + a + r + ':$' + z + r + ',"' + PRESENT + '"))',
      '=IF($A' + r + '="","",COUNTIF($' + a + r + ':$' + z + r + ',"' + ABSENT + '"))',
      '=IF($A' + r + '="","",IFERROR(' + p + r + '/(' + p + r + '+' + q + r + '),""))'
    ]
  };
}

function syncColumns_(ss, ctx, props) {
  const n = CONFIG.MAX_MEETINGS;
  const state = JSON.parse(props.getProperty('COLUMNS') || '{}');
  const current = {};
  ctx.meetings.forEach(m => { current[m.index] = m.day || null; });

  const source = [];
  const used = {};
  let dataChanged = false, stateChanged = false;
  for (let i = 0; i < n; i++) {
    const before = state[i] ? state[i].day : null;
    const now = current[i] || null;
    if (before === now) { source[i] = i; used[i] = true; continue; }
    source[i] = -1;
    stateChanged = true;
    if (before) dataChanged = true;
  }
  const next = {};
  for (let i = 0; i < n; i++) {
    const now = current[i] || null;
    if (source[i] === i) { if (state[i]) next[i] = state[i]; continue; }
    if (!now) continue;
    for (let j = 0; j < n; j++) {
      if (!used[j] && state[j] && state[j].day === now && current[j] !== now) {
        source[i] = j;
        used[j] = true;
        dataChanged = true;
        next[i] = state[j];
        break;
      }
    }
    if (source[i] === -1) next[i] = { day: now, closed: false };
  }

  if (dataChanged) {
    for (const name of [SHEET.ATTENDANCE, SHEET.MARKS]) {
      const sh = ss.getSheetByName(name);
      const count = sh.getLastRow() - LAYOUT.FIRST_ROW + 1;
      if (count <= 0) continue;
      const range = sh.getRange(LAYOUT.FIRST_ROW, LAYOUT.FIRST_MEETING_COL, count, n);
      const old = range.getValues();
      range.setValues(old.map(row => source.map(j => (j >= 0 ? row[j] : ''))));
    }
  }
  if (stateChanged || dataChanged) props.setProperty('COLUMNS', JSON.stringify(next));
  return next;
}

function evaluate_(r, ctx, isMarked) {
  const day = Utilities.formatDate(r.ts, tz_(), 'yyyy-MM-dd');
  const today = ctx.meetings.find(k => k.day === day);
  const m = ctx.meetings.find(k => k.valid && r.ts >= opensAt_(k) && r.ts <= acceptsUntil_(k)) || null;
  const result = {
    status: 'REJECTED', reason: '', date: m ? m.date : (today ? today.date : ''), member: null, meeting: m
  };

  if (!emailAllowed_(r.email)) { result.reason = 'Not a school account, or no email was collected.'; return result; }
  if (!m) {
    result.reason = !today ? 'No meeting on this date.' : today.error ? today.error : 'Outside meeting hours.';
    return result;
  }
  const matches = ctx.byEmail[r.email] || [];
  if (matches.length > 1) { result.reason = 'This email is on more than one row in Members.'; return result; }
  if (!matches.length) { result.reason = 'Email not found in Members.'; return result; }
  const member = matches[0];
  result.member = member;
  if (!key_(member.id)) { result.reason = 'Student ID is empty in Members.'; return result; }
  if (ctx.duplicateIds[key_(member.id)]) { result.reason = 'Student ID appears more than once in Members.'; return result; }
  if (m.code && normalizeCode_(r.code) !== normalizeCode_(m.code)) {
    result.reason = r.code ? 'Wrong meeting code.' : 'Meeting code missing.';
    return result;
  }
  if (isMarked(member.id, m.index)) {
    result.status = 'DUPLICATE';
    result.reason = 'Already marked present for this meeting.';
    return result;
  }
  result.status = 'ACCEPTED';
  result.reason = member.active ? 'Valid.' : 'Valid (member is marked as Left).';
  return result;
}

function context_(ss) {
  const { members, byEmail } = readMembers_(ss);
  const count = {};
  members.forEach(m => { const k = key_(m.id); if (k) count[k] = (count[k] || 0) + 1; });
  const duplicateIds = {};
  Object.keys(count).forEach(k => { if (count[k] > 1) duplicateIds[k] = true; });
  return { meetings: readMeetings_(ss), members, byEmail, duplicateIds };
}

function readMembers_(ss) {
  const sh = ss.getSheetByName(SHEET.MEMBERS);
  const last = sh.getLastRow();
  const members = [], byEmail = {};
  if (last < LAYOUT.FIRST_ROW) return { members, byEmail };
  sh.getRange(LAYOUT.FIRST_ROW, 1, last - LAYOUT.FIRST_ROW + 1, 5).getValues().forEach(r => {
    const email = String(r[3]).trim().toLowerCase();
    if (!key_(r[0]) && !email) return;
    const member = { id: r[0], name: String(r[1]).trim(), email, active: word_(r[4]) === 'ACTIVE' };
    members.push(member);
    if (email) (byEmail[email] = byEmail[email] || []).push(member);
  });
  return { members, byEmail };
}

function readMeetings_(ss) {
  const sh = ss.getSheetByName(SHEET.MEETINGS);
  const n = CONFIG.MAX_MEETINGS;
  const range = sh.getRange(LAYOUT.FIRST_ROW, 1, n, 4);
  const values = range.getValues();
  const shown = range.getDisplayValues();
  const tz = tz_();
  const out = [];
  for (let i = 0; i < n; i++) {
    const row = LAYOUT.FIRST_ROW + i;
    const day = parseDay_(values[i][0], shown[i][0], tz);
    if (day === null) continue;
    const m = {
      index: i, row: row, day: day || '', date: day ? Utilities.parseDate(day, tz, 'yyyy-MM-dd') : '',
      code: String(shown[i][3]).trim(), valid: false, error: ''
    };
    const start = parseTime_(shown[i][1]);
    const endText = String(shown[i][2]).trim();
    const end = endText ? parseTime_(endText) : '';
    if (!day) m.error = 'Meetings row ' + row + ': date not recognized.';
    else if (!start) m.error = 'Meetings row ' + row + ': start time not recognized (use 16:15).';
    else if (end === null) m.error = 'Meetings row ' + row + ': end time not recognized.';
    else {
      m.start = Utilities.parseDate(day + ' ' + start, tz, 'yyyy-MM-dd HH:mm');
      m.end = end ? Utilities.parseDate(day + ' ' + end, tz, 'yyyy-MM-dd HH:mm')
        : new Date(m.start.getTime() + CONFIG.DEFAULT_DURATION_MIN * 60000);
      if (m.end <= m.start) m.error = 'Meetings row ' + row + ': end time must be after start time.';
    }
    m.valid = !m.error;
    out.push(m);
  }
  return out;
}

function parseDay_(value, shown, tz) {
  if (value instanceof Date) return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  const s = String(shown || value || '').trim();
  if (!s) return null;
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (m) return validDay_(m[1], m[2], m[3]);
  m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(s);
  if (m) return validDay_(m[3], m[2], m[1]);
  return '';
}

function validDay_(y, mo, d) {
  mo = Number(mo); d = Number(d);
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return '';
  return y + '-' + String(mo).padStart(2, '0') + '-' + String(d).padStart(2, '0');
}

function parseTime_(text) {
  const s = String(text || '').trim().toUpperCase().replace(/\s+/g, ' ');
  const m = /^(\d{1,2})[:.](\d{2})(?:[:.]\d{2})?(?: ?(AM|PM))?$/.exec(s);
  if (!m) return null;
  let hour = Number(m[1]);
  const minute = Number(m[2]);
  if (m[3]) {
    if (hour < 1 || hour > 12) return null;
    hour = hour === 12 ? (m[3] === 'PM' ? 12 : 0) : (m[3] === 'PM' ? hour + 12 : hour);
  }
  if (hour > 23 || minute > 59) return null;
  return String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
}

function opensAt_(m) { return new Date(m.start.getTime() - CONFIG.OPEN_EARLY_MIN * 60000); }
function acceptsUntil_(m) { return new Date(m.end.getTime() + CONFIG.GRACE_MIN * 60000); }

function emailAllowed_(email) {
  if (!email) return false;
  const domain = String(CONFIG.EMAIL_DOMAIN || '').trim().toLowerCase().replace(/^@/, '');
  return !domain || email.endsWith('@' + domain);
}

function normalizeCode_(s) {
  return stripAccents_(String(s == null ? '' : s).toLocaleUpperCase(CONFIG.LOCALE)).replace(/\s+/g, '');
}

function stripAccents_(s) {
  return s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function word_(s) { return String(s == null ? '' : s).trim().toUpperCase(); }

function key_(id) {
  const s = String(id == null ? '' : id).trim().toUpperCase();
  return /^\d+$/.test(s) ? s.replace(/^0+(?=\d)/, '') : s;
}

function col_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function a1_(row, column) { return col_(column) + row; }

function sheetRef_(name) { return '\'' + name.replace(/'/g, '\'\'') + '\''; }

function readResponse_(response, codeItem) {
  let code = '';
  if (codeItem) {
    const item = response.getResponseForItem(codeItem);
    if (item) code = String(item.getResponse() || '');
  }
  return {
    ts: response.getTimestamp(),
    email: String(response.getRespondentEmail() || '').trim().toLowerCase(),
    code: code
  };
}

function codeItem_(form, props) {
  const existing = findCodeItem_(form, props);
  if (existing) return existing.asTextItem();
  const item = form.addTextItem()
    .setTitle('Meeting code')
    .setHelpText('No code for this meeting; leave blank.')
    .setRequired(false);
  props.setProperty('CODE_ITEM_ID', String(item.getId()));
  return item;
}

function findCodeItem_(form, props) {
  const id = props.getProperty('CODE_ITEM_ID');
  if (!id) return null;
  try { return form.getItemById(Number(id)) || null; } catch (e) { return null; }
}

function setCodeRequired_(form, props, required) {
  const item = codeItem_(form, props);
  if (item.isRequired() !== required) {
    item.setRequired(required)
      .setHelpText(required ? 'Type the code announced in the meeting.' : 'No code for this meeting; leave blank.');
  }
}

function writeStatus_(ss, meetings, open) {
  const tz = tz_();
  const f = (d, p) => Utilities.formatDate(d, tz, p);
  const errors = meetings.filter(m => m.error).map(m => m.error);
  let text;
  if (errors.length) {
    text = 'ERROR · ' + errors.join(' · ');
  } else if (open) {
    text = 'Form open · ' + f(open.start, 'yyyy-MM-dd HH:mm') + '–' + f(open.end, 'HH:mm') +
      (open.code ? ' · code: ' + open.code : '');
  } else {
    const now = new Date();
    const next = meetings.filter(m => m.valid && m.end > now).sort((a, b) => a.start - b.start)[0];
    text = 'Running · form closed · ' + (next ? 'next: ' + f(next.start, 'yyyy-MM-dd HH:mm') : 'no meetings scheduled');
  }
  const cell = ss.getSheetByName(SHEET.DASHBOARD).getRange(LAYOUT.STATUS_CELL);
  if (cell.getDisplayValue() !== text) cell.setValue(text);
}

function setupRegistrationForm_(props, warnings) {
  let form = openForm_(props.getProperty('REGISTRATION_FORM_ID'));
  if (!form) {
    form = FormApp.create(CONFIG.CLUB_NAME + ' · Registration');
    props.setProperty('REGISTRATION_FORM_ID', form.getId());
    props.deleteProperty('REG_ITEMS');
  }
  form.setDescription('Fill this in with your school Google account to join the club. ' +
    'Your email address is collected automatically.');
  form.setLimitOneResponsePerUser(false);
  tryRun_(() => form.setConfirmationMessage('You are registered. Scan the attendance QR code at meetings.'));
  verifiedEmail_(form, warnings);
  registrationItems_(form, props);
  publish_(form, warnings);
  tryRun_(() => form.setCustomClosedFormMessage('Registration is currently closed.'));
  return form;
}

function regQuestions_() {
  return REG_QUESTIONS.filter(q => q.key !== 'section' || CONFIG.ASK_SECTION);
}

function registrationItems_(form, props) {
  const items = findRegistrationItems_(form, props);
  const ids = {};
  for (const q of regQuestions_()) {
    if (!items[q.key]) items[q.key] = createRegistrationItem_(form, q);
    ids[q.key] = items[q.key].getId();
  }
  props.setProperty('REG_ITEMS', JSON.stringify(ids));
  return items;
}

function findRegistrationItems_(form, props) {
  const ids = JSON.parse(props.getProperty('REG_ITEMS') || '{}');
  const items = {};
  for (const q of regQuestions_()) {
    let item = null;
    if (ids[q.key] != null) {
      try { item = form.getItemById(Number(ids[q.key])); } catch (e) { item = null; }
    }
    items[q.key] = item;
  }
  return items;
}

function createRegistrationItem_(form, q) {
  if (q.type === 'grade') {
    return form.addListItem().setTitle(q.title).setChoiceValues(CONFIG.GRADES.map(g => g[0])).setRequired(true);
  }
  const item = form.addTextItem().setTitle(q.title).setRequired(true);
  if (q.type === 'id') {
    item.setValidation(FormApp.createTextValidation()
      .setHelpText('Letters, digits and hyphens only.')
      .requireTextMatchesPattern('^\\s*[A-Za-z0-9-]{1,20}\\s*$').build());
  } else if (q.type === 'section') {
    item.setHelpText('For example: B');
    item.setValidation(FormApp.createTextValidation()
      .setHelpText('Up to 3 letters or digits.')
      .requireTextMatchesPattern('^\\s*[A-Za-z0-9]{1,3}\\s*$').build());
  }
  return item;
}

function recordRegistration(e) {
  if (!e || !e.response) throw new Error('This function runs automatically when the registration form is submitted.');
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(120000)) {
    const pending = JSON.parse(props.getProperty('PENDING_REGISTRATIONS') || '[]');
    pending.push(e.response.getId());
    props.setProperty('PENDING_REGISTRATIONS', JSON.stringify(pending));
    return;
  }
  try {
    const form = e.source || openForm_(props.getProperty('REGISTRATION_FORM_ID'));
    processRegistration_(spreadsheet_(), e.response, findRegistrationItems_(form, props), props);
  } finally {
    lock.releaseLock();
  }
}

function processPendingRegistrations_(ss, props) {
  const pending = JSON.parse(props.getProperty('PENDING_REGISTRATIONS') || '[]');
  if (!pending.length) return false;
  props.deleteProperty('PENDING_REGISTRATIONS');
  const form = openForm_(props.getProperty('REGISTRATION_FORM_ID'));
  if (!form) return false;
  const items = findRegistrationItems_(form, props);
  for (const id of pending) {
    try { processRegistration_(ss, form.getResponse(id), items, props); }
    catch (err) { console.warn('Pending registration failed: ' + err.message); }
  }
  return true;
}

function processRegistration_(ss, response, items, props) {
  const reg = readRegistration_(response, items);
  const result = applyRegistration_(ss, reg);
  if (result.reprocess) props.setProperty('REPROCESS', '1');
  const sh = ss.getSheetByName(SHEET.REGISTRATIONS);
  if (sh) {
    const row = Math.max(sh.getLastRow(), LAYOUT.FIRST_ROW - 1) + 1;
    fitRows_(sh, row);
    sh.getRange(row, 1, 1, 7).setValues([[reg.ts, reg.email, reg.id, result.status, result.details, reg.name, reg.className]]);
    sh.getRange(row, 1).setNumberFormat(DATETIME_FORMAT);
  }
  SpreadsheetApp.flush();
  return result;
}

function readRegistration_(response, items) {
  const answer = k => {
    const item = items[k];
    if (!item) return '';
    const r = response.getResponseForItem(item);
    return r ? String(r.getResponse() || '').replace(/\s+/g, ' ').trim() : '';
  };
  const first = formatName_(answer('first'), false);
  const last = formatName_(answer('last'), CONFIG.UPPERCASE_LAST_NAME);
  const grade = answer('grade');
  const section = CONFIG.ASK_SECTION ? answer('section').toLocaleUpperCase(CONFIG.LOCALE) : '';
  return {
    ts: response.getTimestamp(),
    email: String(response.getRespondentEmail() || '').trim().toLowerCase(),
    id: key_(answer('id').replace(/\s+/g, '')),
    first: first, last: last, name: (first + ' ' + last).trim(),
    grade: grade, section: section, className: section ? grade + '/' + section : grade
  };
}

function applyRegistration_(ss, reg) {
  const rejected = d => ({ status: 'REJECTED', details: d });
  const review = d => ({ status: 'REVIEW', details: d });
  if (!emailAllowed_(reg.email)) return rejected('Not a school account, or no email was collected.');
  if (!/^[A-Z0-9-]{1,20}$/.test(reg.id)) return rejected('Invalid student ID.');
  if (!reg.first || !reg.last) return rejected('First or last name is missing.');
  if (!CONFIG.GRADES.some(g => g[0] === reg.grade)) return rejected('Invalid grade.');
  if (CONFIG.ASK_SECTION && !/^[\p{L}\p{N}]{1,3}$/u.test(reg.section)) return rejected('Invalid section.');

  const sh = ss.getSheetByName(SHEET.MEMBERS);
  const last = sh.getLastRow();
  const count = Math.max(0, last - LAYOUT.FIRST_ROW + 1);
  const data = count ? sh.getRange(LAYOUT.FIRST_ROW, 1, count, 8).getValues() : [];
  const byId = [], byEmail = [];
  data.forEach((r, i) => {
    if (key_(r[0]) === reg.id) byId.push(i);
    if (String(r[3]).trim().toLowerCase() === reg.email) byEmail.push(i);
  });
  if (byId.length > 1) return review('This student ID is on more than one row in Members.');
  if (byEmail.length > 1) return review('This email is on more than one row in Members.');
  const e = byEmail.length ? byEmail[0] : -1;

  if (!byId.length) {
    if (e >= 0) return review('This email belongs to student ' + data[e][0] + ' (' + data[e][1] + ').');
    const row = Math.max(last, LAYOUT.FIRST_ROW - 1) + 1;
    fitRows_(sh, row);
    const idValue = /^\d{1,15}$/.test(reg.id) ? Number(reg.id) : reg.id;
    const today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
    sh.getRange(row, 1, 1, 8).setValues([[idValue, reg.name, reg.className, reg.email, 'Active', 'Verified',
      graduationYear_(reg.grade), 'Registered via form · ' + today]]);
    ensureRows_(ss, [idValue]);
    return { status: 'ADDED', details: 'New member added.', reprocess: true };
  }

  const i = byId[0], r = data[i], row = LAYOUT.FIRST_ROW + i;
  if (e >= 0 && e !== i) return review('This email belongs to student ' + data[e][0] + ' (' + data[e][1] + ').');
  const oldEmail = String(r[3]).trim().toLowerCase();
  const done = [];
  let reprocess = false;
  if (oldEmail !== reg.email) {
    if (oldEmail && word_(r[5]) === 'VERIFIED') {
      return review('This student ID is linked to a different verified email (' + oldEmail + ').');
    }
    if (!namesMatch_(reg, r[1])) return review('Name does not match Members (listed as: ' + r[1] + ').');
    sh.getRange(row, 4).setValue(reg.email);
    done.push(oldEmail ? 'email corrected (was: ' + oldEmail + ')' : 'email added');
    reprocess = true;
  }
  if (word_(r[5]) !== 'VERIFIED') {
    sh.getRange(row, 6).setValue('Verified');
    if (!reprocess) done.push('email verified');
  }
  if (word_(r[4]) !== 'ACTIVE') {
    sh.getRange(row, 5).setValue('Active');
    done.push('reactivated');
  }
  if (String(r[2]).trim() !== reg.className) {
    sh.getRange(row, 3).setValue(reg.className);
    sh.getRange(row, 7).setValue(graduationYear_(reg.grade));
    done.push('class ' + reg.className);
  }
  if (!done.length) return { status: 'REGISTERED', details: 'Already registered; nothing changed.' };
  return { status: 'UPDATED', details: done.join(', ') + '.', reprocess: reprocess };
}

function namesMatch_(reg, listed) {
  const tokens = nameTokens_(listed);
  const first = nameTokens_(reg.first), last = nameTokens_(reg.last);
  return first.length > 0 && last.length > 0 &&
    last.every(t => tokens.indexOf(t) >= 0) && tokens.indexOf(first[0]) >= 0;
}

function nameTokens_(s) {
  return stripAccents_(String(s == null ? '' : s).toLocaleUpperCase(CONFIG.LOCALE))
    .replace(/[^\p{L}\p{N}]+/gu, ' ').trim().split(' ').filter(Boolean);
}

function formatName_(s, upperAll) {
  const L = CONFIG.LOCALE;
  s = String(s || '').replace(/\s+/g, ' ').trim();
  if (upperAll) return s.toLocaleUpperCase(L);
  if (s !== s.toLocaleLowerCase(L) && s !== s.toLocaleUpperCase(L)) return s;
  return s.toLocaleLowerCase(L).replace(/(^|[\s'-])(\p{L})/gu, (m, p, c) => p + c.toLocaleUpperCase(L));
}

function graduationYear_(grade) {
  const g = CONFIG.GRADES.find(x => x[0] === grade);
  if (!g) return '';
  const now = new Date();
  const year = Number(Utilities.formatDate(now, tz_(), 'yyyy'));
  const month = Number(Utilities.formatDate(now, tz_(), 'MM'));
  return (month >= CONFIG.SCHOOL_YEAR_START_MONTH ? year + 1 : year) + g[1];
}

function toggleRegistration() {
  const form = openForm_(PropertiesService.getScriptProperties().getProperty('REGISTRATION_FORM_ID'));
  if (!form) throw new Error('Run setup first.');
  const open = !form.isAcceptingResponses();
  form.setAcceptingResponses(open);
  if (!open) tryRun_(() => form.setCustomClosedFormMessage('Registration is currently closed.'));
  notify_(open ? 'Registration is now open.' : 'Registration is now closed.');
}

function writeFormLinks_(ss, form, regForm) {
  const d = ss.getSheetByName(SHEET.DASHBOARD);
  const a = form.getPublishedUrl(), b = regForm.getPublishedUrl();
  d.getRange('B8:B10').setValues([[a], [form.getEditUrl()], [qrUrl_(a)]]);
  d.getRange('B11').setFormula('=IMAGE(B10,4,190,190)');
  d.getRange('B15:B17').setValues([[b], [regForm.getEditUrl()], [qrUrl_(b)]]);
  d.getRange('B18').setFormula('=IMAGE(B17,4,190,190)');
}

function prepareSheets_(ss) {
  const first = LAYOUT.FIRST_MEETING_COL, n = CONFIG.MAX_MEETINGS;
  fitColumns_(ss.getSheetByName(SHEET.ATTENDANCE), first + n + 2);
  fitColumns_(ss.getSheetByName(SHEET.MARKS), first + n - 1);
  fitRows_(ss.getSheetByName(SHEET.MEETINGS), LAYOUT.FIRST_ROW + n - 1);
  writeFormulas_(ss);
}

function writeFormulas_(ss) {
  const first = LAYOUT.FIRST_MEETING_COL, n = CONFIG.MAX_MEETINGS;
  const M = sheetRef_(SHEET.MEETINGS), A = sheetRef_(SHEET.ATTENDANCE);

  const date = 'INDEX(' + M + '!$A:$A,COLUMN()-' + (first - LAYOUT.FIRST_ROW) + ')';
  ss.getSheetByName(SHEET.ATTENDANCE).getRange(1, first, 1, n)
    .setFormulas([Array(n).fill('=IF(' + date + '="","",' + date + ')')]);

  const grid = A + '!$' + col_(first) + '$2:$' + col_(first + n - 1);
  const meeting = c => 'INDEX(' + M + '!$' + c + ':$' + c + ',ROW())';
  const column = 'INDEX(' + grid + ',0,ROW()-' + (LAYOUT.FIRST_ROW - 1) + ')';
  const rows = [];
  for (let i = 0; i < n; i++) {
    const r = LAYOUT.FIRST_ROW + i;
    rows.push([
      '=IF(' + meeting('A') + '="","",' + meeting('A') + ')',
      '=IF(' + meeting('A') + '="","",' + meeting('B') + ')',
      '=IF($A' + r + '="","",COUNTIF(' + column + ',"' + PRESENT + '"))',
      '=IF($A' + r + '="","",COUNTIF(' + column + ',"' + ABSENT + '"))',
      '=IF($A' + r + '="","",IFERROR(C' + r + '/(C' + r + '+D' + r + '),""))',
      '=IF(' + meeting('E') + '="","",' + meeting('E') + ')'
    ]);
  }
  ss.getSheetByName(SHEET.REPORT).getRange(LAYOUT.FIRST_ROW, 1, n, 6).setFormulas(rows);

  const d = ss.getSheetByName(SHEET.DASHBOARD);
  const U = sheetRef_(SHEET.MEMBERS), T = M, R = sheetRef_(SHEET.RESPONSES), G = sheetRef_(SHEET.REGISTRATIONS);
  const p = col_(first + n), q = col_(first + n + 1);
  const present = 'SUM(' + A + '!' + p + '2:' + p + ')', absent = 'SUM(' + A + '!' + q + '2:' + q + ')';
  d.getRange('A1').setValue(CONFIG.CLUB_NAME + ' · Attendance');
  d.getRange('B3').setFormula('=COUNTIF(' + U + '!E2:E,"Active")');
  d.getRange('D3').setFormula('=COUNT(' + T + '!A2:A)');
  d.getRange('F3').setFormula('=IFERROR(' + present + '/(' + present + '+' + absent + '),"")');
  d.getRange('B4').setFormula('=' + present);
  d.getRange('D4').setFormula('=COUNTIF(' + R + '!D2:D,"REJECTED")');
  d.getRange('F4').setFormula('=COUNTIF(' + G + '!D2:D,"REVIEW")');
  d.getRange('B5').setFormula('=COUNTIF(' + U + '!F2:F,"Unverified")');
}

function fitRows_(sh, lastRow) {
  const max = sh.getMaxRows();
  if (lastRow > max) sh.insertRowsAfter(max, lastRow - max + 100);
}

function fitColumns_(sh, lastColumn) {
  const max = sh.getMaxColumns();
  if (lastColumn > max) sh.insertColumnsAfter(max, lastColumn - max);
}

function verifiedEmail_(form, warnings) {
  try {
    if (typeof form.setEmailCollectionType === 'function' && FormApp.EmailCollectionType) {
      form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
      return;
    }
  } catch (e) { }
  form.setCollectEmail(true);
  warnings.push('"' + form.getTitle() + '": set Settings > Responses > Collect email addresses to "Verified".');
}

function publish_(form, warnings) {
  if (typeof form.setPublished !== 'function') return;
  try { if (!form.isPublished()) form.setPublished(true); }
  catch (e) { warnings.push('"' + form.getTitle() + '" could not be published: ' + e.message); }
}

function qrUrl_(url) {
  return 'https://quickchart.io/qr?size=400&margin=2&text=' + encodeURIComponent(url);
}

function openForm_(id) {
  if (!id) return null;
  try { return FormApp.openById(id); } catch (e) { return null; }
}

function attendanceForm_() {
  const form = openForm_(PropertiesService.getScriptProperties().getProperty('ATTENDANCE_FORM_ID'));
  if (!form) throw new Error('Attendance form not found. Run setup.');
  return form;
}

function spreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Run setup first.');
  return SpreadsheetApp.openById(id);
}

function tz_() {
  if (!TIME_ZONE_) TIME_ZONE_ = spreadsheet_().getSpreadsheetTimeZone();
  return TIME_ZONE_;
}

function tryRun_(fn) {
  try { fn(); } catch (e) { console.warn('Skipped: ' + e.message); }
}

function notify_(text) {
  try { SpreadsheetApp.getUi().alert(text); } catch (e) { console.log(text); }
}
