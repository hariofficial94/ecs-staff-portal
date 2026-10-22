/**
 * ECS Staff Portal — server (Google Apps Script web app)
 * Client: EC Smart Automations Pvt Ltd
 *
 * Live features: To Do (with due date + in-app reminders), Daily Reflection (mandatory on days
 * the app is used; daily digest email to the owner), Team Reflections with replies,
 * Notice Board (announcements + circulars), Document Library, Admin panel.
 * Built but inert (flagged off): Assigned Tasks, Team Completion.
 *
 * Deployment: Execute as "Me" (deploying account), access "Anyone".
 * After every change to this file: Deploy > Manage deployments > Edit > New version.
 */

// ═══════════════════════════════════════════════════════════════════════════
// FEATURE FLAGS — the single switchboard. Flip a boolean and deploy a new version.
// Every gated feature is checked in doGet (pane not rendered) AND inside its own
// server functions (refuses to run if called directly).
// ═══════════════════════════════════════════════════════════════════════════
const FEATURES_ENABLED = {
  todo: true,
  reflection: true,
  announcements: true,
  library: true,
  teamReflections: true,   // reader view: support requests, by day, by person, replies
  reflectionDigest: true,  // daily email to the owner: written / not written / no activity
  assignedTasks: false,
  teamCompletion: false
};

// ═══════════════════════════════════════════════════════════════════════════
// CONSTANTS — tab names are top-level constants only (no inline literals)
// ═══════════════════════════════════════════════════════════════════════════
const SHEET_ID = '10GJ5INILBMX14bPidnkLHyndGRurdDDYUq6tcDbrhcs';
const TZ = 'Asia/Kolkata';
const APP_NAME = 'ECS Staff Portal';

const TAB_CONFIG = 'Config_EC';
const TAB_TODO = 'Todo_Items_EC';
const TAB_REPORT = 'Daily_Report_EC';
const TAB_ALLOC = 'Time_Alloc_EC';
const TAB_ANN = 'Announcements_EC';
const TAB_ANN_READS = 'Announcement_Reads_EC';
const TAB_LIBRARY = 'Library_EC';
const TAB_TASKS = 'Assigned_Tasks_EC';
const TAB_TASK_COMMENTS = 'Task_Comments_EC';
const TAB_LISTS = 'Lists_EC';
const TAB_ACTIVITY = 'Activity_Log_EC';
const TAB_SUPPORT = 'Support_Responses_EC';
const TAB_DIGEST = 'Digest_Log_EC';

// Daily Reflection rules (SOP 10, revised 22 Oct 2026): mandatory on any day the person used the app.
const REFLECTION_DUE_TIME = '19:00';    // shown to staff as "due by"
const REFLECTION_NUDGE_TIME = '17:00';  // from this time staff see a reminder banner if not written
const DIGEST_HOUR = 19;                 // daily email to the owner, runs ~19:15–19:45 IST
const DIGEST_NEAR_MINUTE = 30;
// Who can read other people's reflections:
//   'owner'    → only the top of the reporting line (empty Reports_To) — ECS setting
//   'managers' → anyone with a team, for their own team (derived from Reports_To)
const REFLECTION_READERS = 'owner';

const SESSION_SECONDS = 21600;      // 6 h — CacheService maximum
const PIN_MAX_FAILS = 5;
const PIN_LOCKOUT_SECONDS = 900;    // 15 min
const DONE_HISTORY_DAYS = 7;        // completed/cancelled to-dos shown for this many days
const MAX_TEXT = 500;
const MAX_LONG_TEXT = 4000;
const STAFF_CODE_ALPHABET = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // no 0/O, 1/I

const RE_DATE = /^\d{4}-\d{2}-\d{2}$/;
const RE_TIME = /^\d{2}:\d{2}$/;

// ═══════════════════════════════════════════════════════════════════════════
// ROUTING
// ═══════════════════════════════════════════════════════════════════════════
function doGet(e) {
  const params = (e && e.parameter) || {};
  const template = HtmlService.createTemplateFromFile('Index');
  template.F = FEATURES_ENABLED;                              // drives which panes render
  template.featuresJson = JSON.stringify(FEATURES_ENABLED);
  template.initialTab = JSON.stringify(resolveInitialTab_(params.page));
  template.appName = APP_NAME;
  return template.evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** A ?page= value pointing at a disabled feature falls back to the first enabled pane. */
function resolveInitialTab_(page) {
  const order = ['todo', 'reflection', 'announcements', 'library', 'assignedTasks', 'teamCompletion'];
  if (page && FEATURES_ENABLED[page] === true) return page;
  for (let i = 0; i < order.length; i++) if (FEATURES_ENABLED[order[i]]) return order[i];
  return '';
}

function assertFeature_(key) {
  if (FEATURES_ENABLED[key] !== true) throw new Error('FEATURE_DISABLED: ' + key);
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET HELPERS — every value leaves readTable_ as a string or number, never a Date
// ═══════════════════════════════════════════════════════════════════════════
let ss_cache_ = null;
function ss_() {
  if (!ss_cache_) ss_cache_ = SpreadsheetApp.openById(SHEET_ID);
  return ss_cache_;
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Missing tab: ' + name);
  return sh;
}

function cellDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v == null ? '' : v).trim();
}

function cellTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm');
  return String(v == null ? '' : v).trim();
}

function cellStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm');
  return String(v == null ? '' : v).trim();
}

/** Normalises by header name: *_Date → date string, *_Time → time string, Hours → number. */
function normaliseCell_(header, v) {
  if (/_Date$/.test(header)) return cellDate_(v);
  if (/_Time$/.test(header)) return cellTime_(v);
  if (header === 'Hours') {
    const n = Number(v);
    return isNaN(n) ? 0 : n;
  }
  return cellStr_(v);
}

/** Reads a tab into {sh, headers, col, rows}. rows[i]._row is the 1-based sheet row. */
function readTable_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  const headers = (values[0] || []).map(function (h) { return String(h).trim(); });
  const col = {};
  headers.forEach(function (h, i) { if (h) col[h] = i + 1; });
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const raw = values[r];
    if (raw.every(function (c) { return c === '' || c === null; })) continue;
    const obj = { _row: r + 1 };
    headers.forEach(function (h, i) { if (h) obj[h] = normaliseCell_(h, raw[i]); });
    rows.push(obj);
  }
  return { sh: sh, headers: headers, col: col, rows: rows };
}

/** Free text that starts with = + - @ is stored with a leading apostrophe so Sheets keeps it as text. */
function safeText_(s) {
  const t = String(s == null ? '' : s);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function appendRecord_(table, obj) {
  const row = table.headers.map(function (h) {
    if (!h || obj[h] === undefined || obj[h] === null) return '';
    return typeof obj[h] === 'number' ? obj[h] : safeText_(obj[h]);
  });
  table.sh.appendRow(row);
}

function updateRecord_(table, rowNum, patch) {
  Object.keys(patch).forEach(function (h) {
    const c = table.col[h];
    if (!c) throw new Error('Unknown column ' + h);
    const v = patch[h];
    table.sh.getRange(rowNum, c).setValue(typeof v === 'number' ? v : safeText_(v));
  });
}

/** Next sequential ID for a prefix, e.g. nextId_(t, 'Ann_ID', 'AN-', 4) → AN-0007. Call inside withLock_. */
function nextId_(table, field, prefix, pad) {
  let max = 0;
  table.rows.forEach(function (r) {
    const id = r[field] || '';
    if (id.indexOf(prefix) === 0) {
      const n = parseInt(id.substring(prefix.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return prefix + String(max + 1).padStart(pad, '0');
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function now_() {
  const d = new Date();
  return {
    date: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'),
    time: Utilities.formatDate(d, TZ, 'HH:mm'),
    key: Utilities.formatDate(d, TZ, 'yyyy-MM-dd HH:mm')
  };
}

function daysAgo_(n) {
  const d = new Date(Date.now() - n * 86400000);
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

// ═══════════════════════════════════════════════════════════════════════════
// INPUT VALIDATION
// ═══════════════════════════════════════════════════════════════════════════
function cleanText_(v, max, label, required) {
  const t = String(v == null ? '' : v).replace(/\r\n/g, '\n').trim();
  if (required && !t) throw new Error(label + ' is required.');
  if (t.length > max) throw new Error(label + ' is too long (max ' + max + ' characters).');
  return t;
}

function cleanDate_(v, label) {
  const t = String(v == null ? '' : v).trim();
  if (t && !RE_DATE.test(t)) throw new Error(label + ' must be a date (YYYY-MM-DD).');
  return t;
}

function cleanTime_(v, label) {
  const t = String(v == null ? '' : v).trim();
  if (t && !RE_TIME.test(t)) throw new Error(label + ' must be a time (HH:MM).');
  return t;
}

function cleanLink_(v) {
  const t = String(v == null ? '' : v).trim();
  if (!t) return '';
  if (!/^https:\/\/[^\s<>"']+$/i.test(t)) throw new Error('Document link must start with https://');
  return t;
}

function linkOut_(v) {
  return /^https:\/\/[^\s<>"']+$/i.test(v || '') ? v : '';
}

function getList_(header) {
  const t = readTable_(TAB_LISTS);
  return t.rows.map(function (r) { return r[header]; }).filter(function (v) { return v; });
}

// ═══════════════════════════════════════════════════════════════════════════
// STAFF + VISIBILITY (derived from Reports_To, never from Type)
// ═══════════════════════════════════════════════════════════════════════════
/** Returns staff without PINs. PINs stay inside verifyPin_. */
function getStaff_() {
  const t = readTable_(TAB_CONFIG);
  const list = t.rows
    .filter(function (r) { return r.Staff_Code; })
    .map(function (r) {
      return {
        code: r.Staff_Code,
        name: r.Name,
        role: r.Role,
        type: r.Type,
        status: r.Staff_Status || 'Active',
        reportsTo: r.Reports_To,
        email: r.Email || ''   // server-side only (digest); never mapped into a client payload
      };
    });
  const byCode = {};
  list.forEach(function (s) { byCode[s.code] = s; });
  return { list: list, byCode: byCode };
}

function isActive_(s) { return s && s.status === 'Active'; }
function isAdmin_(me) { return !me.reportsTo; }

/** Own code + every Active person below me in the Reports_To tree. Roots see all Active staff. */
function visibleCodes_(me, staff) {
  if (isAdmin_(me)) return staff.list.filter(isActive_).map(function (s) { return s.code; });
  const out = [me.code];
  const seenRoles = {};
  const queue = [me.role];
  while (queue.length) {
    const role = queue.shift();
    if (seenRoles[role]) continue;
    seenRoles[role] = true;
    staff.list.forEach(function (s) {
      if (s.reportsTo === role && isActive_(s)) {
        if (out.indexOf(s.code) === -1) out.push(s.code);
        queue.push(s.role);
      }
    });
  }
  return out;
}

function nameOf_(staff, code) {
  const s = staff.byCode[code];
  return s ? s.name : code;
}

// ═══════════════════════════════════════════════════════════════════════════
// SESSIONS — PIN checked server-side only; the browser holds an opaque token
// ═══════════════════════════════════════════════════════════════════════════
function createSession_(code) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('S_' + token, code, SESSION_SECONDS);
  return token;
}

/** Every protected function starts here. Throws SESSION_EXPIRED if the token is unknown. */
function requireUser_(token) {
  if (typeof token !== 'string' || token.length < 32) throw new Error('SESSION_EXPIRED');
  const cache = CacheService.getScriptCache();
  const code = cache.get('S_' + token);
  if (!code) throw new Error('SESSION_EXPIRED');
  const staff = getStaff_();
  const me = staff.byCode[code];
  if (!isActive_(me)) {
    cache.remove('S_' + token);
    throw new Error('SESSION_EXPIRED');
  }
  cache.put('S_' + token, code, SESSION_SECONDS); // sliding expiry
  touchActivity_(code);
  return { me: me, staff: staff };
}

/**
 * Records that this person used the app today (one row per person per day).
 * "Used the app" = any signed-in action, so a session carried over from earlier still counts.
 * Never throws: activity tracking must not break the app.
 */
function touchActivity_(code) {
  try {
    const n = now_();
    const key = 'A_' + code + '_' + n.date;
    const cache = CacheService.getScriptCache();
    if (cache.get(key)) return;
    withLock_(function () {
      const t = readTable_(TAB_ACTIVITY);
      const id = 'AC-' + code + '-' + n.date.replace(/-/g, '');
      if (!t.rows.some(function (r) { return r.Activity_ID === id; })) {
        appendRecord_(t, { Activity_ID: id, Staff_Code: code, Activity_Date: n.date, First_Seen_Time: n.time });
      }
    });
    cache.put(key, '1', SESSION_SECONDS);
  } catch (e) {
    console.error('touchActivity_ failed: ' + e);
  }
}

function requireAdmin_(token) {
  const ctx = requireUser_(token);
  if (!isAdmin_(ctx.me)) throw new Error('Only the top of the reporting line can do this.');
  return ctx;
}

/** Public: names for the login picker. Active staff only; no PINs, no types. */
function getLoginRoster() {
  return getStaff_().list
    .filter(isActive_)
    .map(function (s) { return { code: s.code, name: s.name, role: s.role }; })
    .sort(function (a, b) { return a.name.localeCompare(b.name); });
}

/** Public: verifies PIN, returns {ok, token, data} or {ok:false, message}. PIN never returned. */
function login(code, pin) {
  code = String(code || '').trim();
  pin = String(pin || '').trim();
  if (!code || !pin) return { ok: false, message: 'Choose your name and enter your PIN.' };

  const cache = CacheService.getScriptCache();
  const failKey = 'F_' + code;
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= PIN_MAX_FAILS) {
    return { ok: false, message: 'Too many wrong attempts. Try again in 15 minutes.' };
  }

  const t = readTable_(TAB_CONFIG);
  const row = t.rows.filter(function (r) { return r.Staff_Code === code; })[0];
  const valid = row && (row.Staff_Status || 'Active') === 'Active' && row.PIN && row.PIN === pin;
  if (!valid) {
    cache.put(failKey, String(fails + 1), PIN_LOCKOUT_SECONDS);
    return { ok: false, message: 'Name and PIN do not match.' };
  }
  cache.remove(failKey);
  const token = createSession_(code);
  touchActivity_(code);
  const staff = getStaff_();
  return { ok: true, token: token, data: bootstrap_({ me: staff.byCode[code], staff: staff }) };
}

function logout(token) {
  if (typeof token === 'string') CacheService.getScriptCache().remove('S_' + token);
  return true;
}

/** Everything the page needs after login (or on reopening with a stored token). */
function getBootstrap(token) {
  return bootstrap_(requireUser_(token));
}

function bootstrap_(ctx) {
  const me = ctx.me;
  const n = now_();
  const visible = visibleCodes_(me, ctx.staff);
  const out = {
    me: {
      code: me.code,
      name: me.name,
      role: me.role,
      type: me.type,
      isAdmin: isAdmin_(me),
      hasTeam: visible.length > 1,
      canReadReflections: canReadReflections_(me, ctx.staff),
      reflectionRequired: FEATURES_ENABLED.reflection && !isAdmin_(me)
    },
    reflectionRules: { dueTime: REFLECTION_DUE_TIME, nudgeTime: REFLECTION_NUDGE_TIME },
    features: FEATURES_ENABLED,
    today: n.date,
    nowTime: n.time,
    serverNowMs: Date.now(),
    lists: {},
    todos: [],
    reflection: null,
    announcements: [],
    library: [],
    team: []
  };
  if (FEATURES_ENABLED.todo) out.todos = listTodos_(me.code);
  if (FEATURES_ENABLED.reflection) {
    out.lists.activity = getList_('Activity_Category');
    out.reflection = reflectionFor_(me.code, n.date);
  }
  if (FEATURES_ENABLED.announcements) {
    out.announcements = listAnnouncements_(me, ctx.staff);
    out.lists.annCategory = getList_('Ann_Category');
  }
  if (FEATURES_ENABLED.library) {
    out.library = listLibrary_(me, ctx.staff);
    out.lists.libraryCategory = getList_('Library_Category');
  }
  if (FEATURES_ENABLED.assignedTasks) {
    out.team = visible
      .filter(function (c) { return c !== me.code; })
      .map(function (c) { const s = ctx.staff.byCode[c]; return { code: c, name: s.name, role: s.role }; });
  }
  return out;
}

// ═══════════════════════════════════════════════════════════════════════════
// TO DO — one row per item for its whole life; open items stay until Done/Cancelled
// ═══════════════════════════════════════════════════════════════════════════
function todoOut_(r) {
  return {
    id: r.Item_ID,
    text: r.Text,
    createdDate: r.Created_Date,
    createdTime: r.Created_Time,
    dueDate: r.Due_Date,
    dueTime: r.Due_Time,
    remindDate: r.Remind_Date,
    remindTime: r.Remind_Time,
    status: r.Status || 'Open',
    completedDate: r.Completed_Date,
    completedTime: r.Completed_Time,
    cancelledDate: r.Cancelled_Date
  };
}

function listTodos_(code) {
  const cutoff = daysAgo_(DONE_HISTORY_DAYS);
  return readTable_(TAB_TODO).rows
    .filter(function (r) {
      if (r.Staff_Code !== code) return false;
      const st = r.Status || 'Open';
      if (st === 'Open') return true;
      if (st === 'Done') return r.Completed_Date >= cutoff;
      if (st === 'Cancelled') return r.Cancelled_Date >= cutoff;
      return false;
    })
    .map(todoOut_);
}

function todoInput_(input) {
  input = input || {};
  const v = {
    text: cleanText_(input.text, MAX_TEXT, 'To-do text', true),
    dueDate: cleanDate_(input.dueDate, 'Due date'),
    dueTime: cleanTime_(input.dueTime, 'Due time'),
    remindDate: cleanDate_(input.remindDate, 'Reminder date'),
    remindTime: cleanTime_(input.remindTime, 'Reminder time')
  };
  if (v.dueTime && !v.dueDate) throw new Error('Add a due date for the due time.');
  if (v.remindTime && !v.remindDate) throw new Error('Add a reminder date for the reminder time.');
  return v;
}

function getTodos(token) {
  assertFeature_('todo');
  return listTodos_(requireUser_(token).me.code);
}

function addTodo(token, input) {
  assertFeature_('todo');
  const me = requireUser_(token).me;
  const v = todoInput_(input);
  return withLock_(function () {
    const t = readTable_(TAB_TODO);
    const n = now_();
    appendRecord_(t, {
      Item_ID: nextId_(t, 'Item_ID', 'TD-' + me.code + '-', 4),
      Staff_Code: me.code,
      Created_Date: n.date,
      Created_Time: n.time,
      Text: v.text,
      Due_Date: v.dueDate,
      Due_Time: v.dueTime,
      Remind_Date: v.remindDate,
      Remind_Time: v.remindTime,
      Status: 'Open'
    });
    return listTodos_(me.code);
  });
}

function findOwnTodo_(t, me, itemId) {
  const row = t.rows.filter(function (r) { return r.Item_ID === itemId; })[0];
  if (!row || row.Staff_Code !== me.code) throw new Error('To-do not found.');
  return row;
}

function updateTodo(token, itemId, input) {
  assertFeature_('todo');
  const me = requireUser_(token).me;
  const v = todoInput_(input);
  return withLock_(function () {
    const t = readTable_(TAB_TODO);
    const row = findOwnTodo_(t, me, itemId);
    if ((row.Status || 'Open') !== 'Open') throw new Error('Only open to-dos can be edited.');
    updateRecord_(t, row._row, {
      Text: v.text,
      Due_Date: v.dueDate,
      Due_Time: v.dueTime,
      Remind_Date: v.remindDate,
      Remind_Time: v.remindTime
    });
    return listTodos_(me.code);
  });
}

/** status: 'Done' | 'Open' (undo / restore) | 'Cancelled' (soft cancel, row kept). */
function setTodoStatus(token, itemId, status) {
  assertFeature_('todo');
  const me = requireUser_(token).me;
  if (['Open', 'Done', 'Cancelled'].indexOf(status) === -1) throw new Error('Unknown status.');
  return withLock_(function () {
    const t = readTable_(TAB_TODO);
    const row = findOwnTodo_(t, me, itemId);
    const n = now_();
    const patch = { Status: status };
    if (status === 'Done') { patch.Completed_Date = n.date; patch.Completed_Time = n.time; patch.Cancelled_Date = ''; }
    if (status === 'Cancelled') { patch.Cancelled_Date = n.date; patch.Completed_Date = ''; patch.Completed_Time = ''; }
    if (status === 'Open') { patch.Completed_Date = ''; patch.Completed_Time = ''; patch.Cancelled_Date = ''; }
    updateRecord_(t, row._row, patch);
    return listTodos_(me.code);
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// DAILY REFLECTION — no deadline, no lock, no lateness label. Any past date is editable.
// ═══════════════════════════════════════════════════════════════════════════
function reportId_(code, date) {
  return 'RP-' + code + '-' + date.replace(/-/g, '');
}

function reflectionFor_(code, date) {
  const reports = readTable_(TAB_REPORT).rows.filter(function (r) { return r.Staff_Code === code; });
  const row = reports.filter(function (r) { return r.Report_Date === date; })[0];
  const rid = reportId_(code, date);
  const alloc = readTable_(TAB_ALLOC).rows
    .filter(function (r) { return r.Report_ID === rid && (r.Status || 'Active') === 'Active'; })
    .map(function (r) { return { category: r.Activity_Category, hours: r.Hours, note: r.Note }; });
  const recent = reports
    .map(function (r) { return r.Report_Date; })
    .filter(function (d) { return d && d !== date; })
    .sort()
    .reverse()
    .slice(0, 10);
  return {
    reportId: row ? row.Report_ID : rid,
    date: date,
    exists: !!row,
    challenges: row ? row.Challenges_Blockers : '',
    support: row ? row.Support_Required : '',
    learning: row ? row.Learning_Note : '',
    submittedDate: row ? row.Submitted_Date : '',
    submittedTime: row ? row.Submitted_Time : '',
    lastEditedDate: row ? row.Last_Edited_Date : '',
    lastEditedTime: row ? row.Last_Edited_Time : '',
    alloc: alloc,
    recentDates: recent,
    responses: responsesFor_(row ? row.Report_ID : rid, getStaff_())
  };
}

function getReflection(token, date) {
  assertFeature_('reflection');
  const me = requireUser_(token).me;
  const d = cleanDate_(date, 'Date') || now_().date;
  if (d > now_().date) throw new Error('Reflections can only be written for today or earlier.');
  return reflectionFor_(me.code, d);
}

function saveReflection(token, input) {
  assertFeature_('reflection');
  const me = requireUser_(token).me;
  input = input || {};
  const n = now_();
  const date = cleanDate_(input.date, 'Date') || n.date;
  if (date > n.date) throw new Error('Reflections can only be written for today or earlier.');
  const challenges = cleanText_(input.challenges, MAX_LONG_TEXT, 'Challenges / Blockers', false);
  const support = cleanText_(input.support, MAX_LONG_TEXT, 'Support Required', false);
  const learning = cleanText_(input.learning, MAX_LONG_TEXT, 'Learning Note', false);

  const categories = getList_('Activity_Category');
  const alloc = (Array.isArray(input.alloc) ? input.alloc : [])
    .map(function (a) {
      return {
        category: String((a && a.category) || '').trim(),
        hours: Number(a && a.hours),
        note: cleanText_(a && a.note, MAX_TEXT, 'Time note', false)
      };
    })
    .filter(function (a) { return a.category || a.hours || a.note; });
  let total = 0;
  alloc.forEach(function (a) {
    if (categories.indexOf(a.category) === -1) throw new Error('Pick an activity for every time row.');
    if (!(a.hours > 0 && a.hours <= 24)) throw new Error('Hours must be between 0 and 24.');
    a.hours = Math.round(a.hours * 4) / 4; // quarter-hour precision
    total += a.hours;
  });
  if (total > 24) throw new Error('Time allocation adds up to more than 24 hours.');
  if (!challenges && !support && !learning && !alloc.length) {
    throw new Error('Write something in at least one section before saving.');
  }

  return withLock_(function () {
    const rid = reportId_(me.code, date);
    const reports = readTable_(TAB_REPORT);
    const existing = reports.rows.filter(function (r) { return r.Report_ID === rid; })[0];
    if (existing) {
      updateRecord_(reports, existing._row, {
        Challenges_Blockers: challenges,
        Support_Required: support,
        Learning_Note: learning,
        Last_Edited_Date: n.date,
        Last_Edited_Time: n.time
      });
    } else {
      appendRecord_(reports, {
        Report_ID: rid,
        Staff_Code: me.code,
        Report_Date: date,
        Submitted_Date: n.date,
        Submitted_Time: n.time,
        Challenges_Blockers: challenges,
        Support_Required: support,
        Learning_Note: learning
      });
    }

    // Time allocation: previous rows are marked Replaced (kept for history), new rows appended.
    const allocT = readTable_(TAB_ALLOC);
    allocT.rows.forEach(function (r) {
      if (r.Report_ID === rid && (r.Status || 'Active') === 'Active') {
        updateRecord_(allocT, r._row, { Status: 'Replaced' });
      }
    });
    alloc.forEach(function (a) {
      const allocId = nextId_(allocT, 'Alloc_ID', 'TL-', 5);
      appendRecord_(allocT, {
        Alloc_ID: allocId,
        Report_ID: rid,
        Staff_Code: me.code,
        Report_Date: date,
        Activity_Category: a.category,
        Hours: a.hours,
        Note: a.note,
        Status: 'Active'
      });
      allocT.rows.push({ Alloc_ID: allocId }); // so the next ID in this loop increments
    });
    return reflectionFor_(me.code, date);
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// NOTICE BOARD — announcements and circulars (a circular carries a document link)
// ═══════════════════════════════════════════════════════════════════════════
function annOut_(r, staff) {
  return {
    id: r.Ann_ID,
    category: r.Category || 'Announcement',
    title: r.Title,
    body: r.Body,
    docLink: linkOut_(r.Doc_Link),
    postedBy: nameOf_(staff, r.Posted_By_Code),
    postedDate: r.Posted_Date,
    postedTime: r.Posted_Time,
    expiryDate: r.Expiry_Date,
    pinned: r.Pinned === 'Yes',
    status: r.Status || 'Active'
  };
}

function sortAnn_(a, b) {
  if (a.pinned !== b.pinned) return a.pinned ? -1 : 1;
  const ka = a.postedDate + ' ' + a.postedTime;
  const kb = b.postedDate + ' ' + b.postedTime;
  return ka < kb ? 1 : ka > kb ? -1 : 0;
}

function listAnnouncements_(me, staff) {
  const today = now_().date;
  const readIds = {};
  readTable_(TAB_ANN_READS).rows.forEach(function (r) {
    if (r.Staff_Code === me.code) readIds[r.Ann_ID] = true;
  });
  return readTable_(TAB_ANN).rows
    .filter(function (r) {
      return (r.Status || 'Active') === 'Active' && (!r.Expiry_Date || r.Expiry_Date >= today);
    })
    .map(function (r) {
      const o = annOut_(r, staff);
      o.unread = !readIds[r.Ann_ID] && r.Posted_By_Code !== me.code; // your own posts are never "new" to you
      return o;
    })
    .sort(sortAnn_);
}

function getAnnouncements(token) {
  assertFeature_('announcements');
  const ctx = requireUser_(token);
  return listAnnouncements_(ctx.me, ctx.staff);
}

/** Records reads for the given IDs (skips ones already read). Returns the refreshed list. */
function markAnnouncementsRead(token, annIds) {
  assertFeature_('announcements');
  const ctx = requireUser_(token);
  const ids = (Array.isArray(annIds) ? annIds : []).map(String).slice(0, 200);
  if (!ids.length) return listAnnouncements_(ctx.me, ctx.staff);
  return withLock_(function () {
    const reads = readTable_(TAB_ANN_READS);
    const known = {};
    readTable_(TAB_ANN).rows.forEach(function (r) { known[r.Ann_ID] = true; });
    const already = {};
    reads.rows.forEach(function (r) { if (r.Staff_Code === ctx.me.code) already[r.Ann_ID] = true; });
    const n = now_();
    ids.forEach(function (id) {
      if (!known[id] || already[id]) return;
      const rid = nextId_(reads, 'Read_ID', 'RD-', 6);
      appendRecord_(reads, { Read_ID: rid, Ann_ID: id, Staff_Code: ctx.me.code, Read_Date: n.date, Read_Time: n.time });
      reads.rows.push({ Read_ID: rid });
      already[id] = true;
    });
    return listAnnouncements_(ctx.me, ctx.staff);
  });
}

function postAnnouncement(token, input) {
  assertFeature_('announcements');
  const ctx = requireAdmin_(token);
  input = input || {};
  const categories = getList_('Ann_Category');
  const category = String(input.category || 'Announcement');
  if (categories.indexOf(category) === -1) throw new Error('Unknown category.');
  const title = cleanText_(input.title, 200, 'Title', true);
  const body = cleanText_(input.body, MAX_LONG_TEXT, 'Message', false);
  const docLink = cleanLink_(input.docLink);
  const expiry = cleanDate_(input.expiryDate, 'Expiry date');
  if (category === 'Circular' && !docLink) throw new Error('A circular needs a document link.');
  if (!body && !docLink) throw new Error('Add a message or a document link.');
  return withLock_(function () {
    const t = readTable_(TAB_ANN);
    const n = now_();
    appendRecord_(t, {
      Ann_ID: nextId_(t, 'Ann_ID', 'AN-', 4),
      Category: category,
      Title: title,
      Body: body,
      Doc_Link: docLink,
      Posted_By_Code: ctx.me.code,
      Posted_Date: n.date,
      Posted_Time: n.time,
      Expiry_Date: expiry,
      Pinned: input.pinned ? 'Yes' : 'No',
      Status: 'Active'
    });
    return adminData_(ctx);
  });
}

function setAnnouncementPinned(token, annId, pinned) {
  assertFeature_('announcements');
  const ctx = requireAdmin_(token);
  return withLock_(function () {
    const t = readTable_(TAB_ANN);
    const row = t.rows.filter(function (r) { return r.Ann_ID === annId; })[0];
    if (!row) throw new Error('Announcement not found.');
    updateRecord_(t, row._row, { Pinned: pinned ? 'Yes' : 'No' });
    return adminData_(ctx);
  });
}

function archiveAnnouncement(token, annId) {
  assertFeature_('announcements');
  const ctx = requireAdmin_(token);
  return withLock_(function () {
    const t = readTable_(TAB_ANN);
    const row = t.rows.filter(function (r) { return r.Ann_ID === annId; })[0];
    if (!row) throw new Error('Announcement not found.');
    updateRecord_(t, row._row, { Status: 'Archived' });
    return adminData_(ctx);
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// DOCUMENT LIBRARY — audience-filtered on the server; links only reach the audience
// ═══════════════════════════════════════════════════════════════════════════
function splitAudience_(v) {
  return String(v || '').split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; });
}

/** A document is visible if its audience covers me or anyone in my visible team. */
function docVisible_(r, me, staff, visible) {
  const type = r.Audience_Type || 'All';
  if (type === 'All') return true;
  const values = splitAudience_(r.Audience_Value);
  return visible.some(function (code) {
    const s = staff.byCode[code];
    if (!s) return false;
    if (type === 'Role') return values.indexOf(s.role) !== -1;
    if (type === 'Staff') return values.indexOf(s.code) !== -1;
    return false;
  });
}

function audienceLabel_(r, staff) {
  const type = r.Audience_Type || 'All';
  if (type === 'All') return 'Everyone';
  const values = splitAudience_(r.Audience_Value);
  if (type === 'Staff') return values.map(function (c) { return nameOf_(staff, c); }).join(', ');
  return values.join(', ');
}

function docOut_(r, staff) {
  return {
    id: r.Doc_ID,
    title: r.Title,
    category: r.Category || 'Other',
    docLink: linkOut_(r.Doc_Link),
    version: r.Version,
    addedDate: r.Added_Date,
    audience: audienceLabel_(r, staff),
    personal: (r.Audience_Type || 'All') !== 'All',
    status: r.Status || 'Active'
  };
}

function listLibrary_(me, staff) {
  const visible = visibleCodes_(me, staff);
  return readTable_(TAB_LIBRARY).rows
    .filter(function (r) { return (r.Status || 'Active') === 'Active' && docVisible_(r, me, staff, visible); })
    .map(function (r) { return docOut_(r, staff); })
    .sort(function (a, b) { return a.title.localeCompare(b.title); });
}

function getLibrary(token) {
  assertFeature_('library');
  const ctx = requireUser_(token);
  return listLibrary_(ctx.me, ctx.staff);
}

function addLibraryDoc(token, input) {
  assertFeature_('library');
  const ctx = requireAdmin_(token);
  input = input || {};
  const title = cleanText_(input.title, 200, 'Title', true);
  const category = String(input.category || '');
  if (getList_('Library_Category').indexOf(category) === -1) throw new Error('Pick a category.');
  const audienceType = String(input.audienceType || 'All');
  if (['All', 'Role', 'Staff'].indexOf(audienceType) === -1) throw new Error('Unknown audience.');
  let values = Array.isArray(input.audienceValues) ? input.audienceValues.map(String) : [];
  if (audienceType === 'All') values = [];
  if (audienceType === 'Role') {
    const roles = ctx.staff.list.map(function (s) { return s.role; });
    values = values.filter(function (v) { return roles.indexOf(v) !== -1; });
    if (!values.length) throw new Error('Pick at least one role.');
  }
  if (audienceType === 'Staff') {
    values = values.filter(function (v) { return !!ctx.staff.byCode[v]; });
    if (!values.length) throw new Error('Pick at least one person.');
  }
  const docLink = cleanLink_(input.docLink);
  if (!docLink) throw new Error('Paste the document link.');
  const version = cleanText_(input.version, 40, 'Version', false);
  return withLock_(function () {
    const t = readTable_(TAB_LIBRARY);
    appendRecord_(t, {
      Doc_ID: nextId_(t, 'Doc_ID', 'LIB-', 4),
      Title: title,
      Category: category,
      Audience_Type: audienceType,
      Audience_Value: values.join(', '),
      Doc_Link: docLink,
      Version: version,
      Added_By_Code: ctx.me.code,
      Added_Date: now_().date,
      Status: 'Active'
    });
    return adminData_(ctx);
  });
}

function supersedeLibraryDoc(token, docId) {
  assertFeature_('library');
  const ctx = requireAdmin_(token);
  return withLock_(function () {
    const t = readTable_(TAB_LIBRARY);
    const row = t.rows.filter(function (r) { return r.Doc_ID === docId; })[0];
    if (!row) throw new Error('Document not found.');
    updateRecord_(t, row._row, { Status: 'Superseded' });
    return adminData_(ctx);
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// ADMIN PANEL — root of the reporting line only (derived, not hardcoded to a name)
// ═══════════════════════════════════════════════════════════════════════════
function getAdminData(token) {
  return adminData_(requireAdmin_(token));
}

function adminData_(ctx) {
  const staff = ctx.staff;
  const active = staff.list.filter(isActive_);
  const out = {
    staff: active.map(function (s) { return { code: s.code, name: s.name, role: s.role }; }),
    roles: active.map(function (s) { return s.role; }),
    announcements: [],
    library: []
  };
  if (FEATURES_ENABLED.announcements) {
    const readers = {};
    readTable_(TAB_ANN_READS).rows.forEach(function (r) {
      (readers[r.Ann_ID] = readers[r.Ann_ID] || {})[r.Staff_Code] = true;
    });
    const cutoff = daysAgo_(60);
    out.announcements = readTable_(TAB_ANN).rows
      .filter(function (r) { return (r.Status || 'Active') === 'Active' || r.Posted_Date >= cutoff; })
      .map(function (r) {
        const o = annOut_(r, staff);
        const seen = readers[r.Ann_ID] || {};
        const audience = active.filter(function (s) { return s.code !== r.Posted_By_Code; });
        o.readCount = audience.filter(function (s) { return seen[s.code]; }).length;
        o.audienceCount = audience.length;
        o.notRead = audience.filter(function (s) { return !seen[s.code]; }).map(function (s) { return s.name; });
        o.expired = !!(r.Expiry_Date && r.Expiry_Date < now_().date);
        return o;
      })
      .sort(sortAnn_);
  }
  if (FEATURES_ENABLED.library) {
    out.library = readTable_(TAB_LIBRARY).rows
      .map(function (r) { return docOut_(r, staff); })
      .sort(function (a, b) {
        if (a.status !== b.status) return a.status === 'Active' ? -1 : 1;
        return a.title.localeCompare(b.title);
      });
  }
  return out;
}

// ═══════════════════════════════════════════════════════════════════════════
// TEAM REFLECTIONS — reader view + replies to support requests (flag: teamReflections)
// Who reads is set by REFLECTION_READERS. Readers never see their own entries here.
// ═══════════════════════════════════════════════════════════════════════════
function canReadReflections_(me, staff) {
  if (!FEATURES_ENABLED.teamReflections) return false;
  if (REFLECTION_READERS === 'owner') return isAdmin_(me);
  return visibleCodes_(me, staff).length > 1;
}

/** Staff whose reflections this reader may see (Active only, never themselves, never other owners). */
function readableCodes_(me, staff) {
  return visibleCodes_(me, staff).filter(function (c) {
    const s = staff.byCode[c];
    return c !== me.code && isActive_(s) && !isAdmin_(s);
  });
}

function requireReader_(token) {
  assertFeature_('teamReflections');
  const ctx = requireUser_(token);
  if (!canReadReflections_(ctx.me, ctx.staff)) throw new Error('You do not have access to team reflections.');
  ctx.readable = readableCodes_(ctx.me, ctx.staff);
  return ctx;
}

/** All responses for one report, oldest first. Each reply is its own row (SR-00001…). */
function responsesFor_(reportId, staff) {
  if (!FEATURES_ENABLED.teamReflections) return [];
  return readTable_(TAB_SUPPORT).rows
    .filter(function (r) { return r.Report_ID === reportId; })
    .map(function (r) {
      return {
        id: r.Response_ID,
        by: nameOf_(staff, r.Responder_Code),
        status: r.Status,
        reply: r.Reply,
        date: r.Response_Date,
        time: r.Response_Time
      };
    });
}

/** Report_ID → latest response row. */
function latestResponses_() {
  const out = {};
  readTable_(TAB_SUPPORT).rows.forEach(function (r) { out[r.Report_ID] = r; });
  return out;
}

function tileOut_(r, staff, latest) {
  const s = staff.byCode[r.Staff_Code] || {};
  const last = latest[r.Report_ID];
  return {
    reportId: r.Report_ID,
    code: r.Staff_Code,
    name: s.name || r.Staff_Code,
    role: s.role || '',
    date: r.Report_Date,
    challenges: r.Challenges_Blockers,
    support: r.Support_Required,
    learning: r.Learning_Note,
    submittedDate: r.Submitted_Date,
    submittedTime: r.Submitted_Time,
    lastEditedDate: r.Last_Edited_Date,
    lastEditedTime: r.Last_Edited_Time,
    supportStatus: r.Support_Required ? (last ? last.Status : 'Open') : '',
    lastReply: last ? { by: nameOf_(staff, last.Responder_Code), reply: last.Reply, date: last.Response_Date, time: last.Response_Time } : null
  };
}

function personOut_(s) { return { code: s.code, name: s.name, role: s.role }; }

/** Three groups for one date: Written / Not written (used the app) / No activity. */
function reflectionDayGroups_(date, codes, staff) {
  const reports = readTable_(TAB_REPORT).rows.filter(function (r) { return r.Report_Date === date; });
  const active = {};
  readTable_(TAB_ACTIVITY).rows.forEach(function (r) { if (r.Activity_Date === date) active[r.Staff_Code] = true; });
  const latest = latestResponses_();
  const out = { date: date, written: [], notWritten: [], noActivity: [] };
  codes.forEach(function (code) {
    const s = staff.byCode[code];
    if (!isActive_(s)) return;
    const rep = reports.filter(function (r) { return r.Staff_Code === code; })[0];
    if (rep) out.written.push(tileOut_(rep, staff, latest));
    else if (active[code]) out.notWritten.push(personOut_(s));
    else out.noActivity.push(personOut_(s));
  });
  const byName = function (a, b) { return a.name.localeCompare(b.name); };
  out.written.sort(byName); out.notWritten.sort(byName); out.noActivity.sort(byName);
  return out;
}

function getTeamReflectionDay(token, date) {
  const ctx = requireReader_(token);
  const d = cleanDate_(date, 'Date') || now_().date;
  const out = reflectionDayGroups_(d, ctx.readable, ctx.staff);
  out.dueTime = REFLECTION_DUE_TIME;
  out.isToday = d === now_().date;
  return out;
}

/** Full entry for a reader: sections + time allocation + reply thread. */
function getReflectionDetail(token, reportId) {
  const ctx = requireReader_(token);
  return reflectionDetail_(ctx, reportId);
}

function reflectionDetail_(ctx, reportId) {
  const row = readTable_(TAB_REPORT).rows.filter(function (r) { return r.Report_ID === reportId; })[0];
  if (!row || ctx.readable.indexOf(row.Staff_Code) === -1) throw new Error('Reflection not found.');
  const out = tileOut_(row, ctx.staff, latestResponses_());
  out.alloc = readTable_(TAB_ALLOC).rows
    .filter(function (r) { return r.Report_ID === reportId && (r.Status || 'Active') === 'Active'; })
    .map(function (r) { return { category: r.Activity_Category, hours: r.Hours, note: r.Note }; });
  out.responses = responsesFor_(reportId, ctx.staff);
  return out;
}

/** Support requests from the last 60 days. show: 'open' (not Resolved) | 'all'. */
function getSupportRequests(token, show) {
  const ctx = requireReader_(token);
  const cutoff = daysAgo_(60);
  const latest = latestResponses_();
  const rank = { Open: 0, Acknowledged: 1, Resolved: 2 };
  return readTable_(TAB_REPORT).rows
    .filter(function (r) {
      return r.Support_Required && r.Report_Date >= cutoff && ctx.readable.indexOf(r.Staff_Code) !== -1;
    })
    .map(function (r) { return tileOut_(r, ctx.staff, latest); })
    .filter(function (t) { return show === 'all' || t.supportStatus !== 'Resolved'; })
    .sort(function (a, b) {
      const ra = rank[a.supportStatus] || 0, rb = rank[b.supportStatus] || 0;
      if (ra !== rb) return ra - rb;
      return a.date < b.date ? 1 : a.date > b.date ? -1 : 0;
    });
}

/** Reader replies to a reflection. status: 'Acknowledged' | 'Resolved'. Returns the refreshed detail. */
function respondToSupport(token, reportId, status, reply) {
  const ctx = requireReader_(token);
  if (['Acknowledged', 'Resolved'].indexOf(status) === -1) throw new Error('Unknown status.');
  const text = cleanText_(reply, MAX_LONG_TEXT, 'Reply', false);
  const row = readTable_(TAB_REPORT).rows.filter(function (r) { return r.Report_ID === reportId; })[0];
  if (!row || ctx.readable.indexOf(row.Staff_Code) === -1) throw new Error('Reflection not found.');
  withLock_(function () {
    const t = readTable_(TAB_SUPPORT);
    const n = now_();
    appendRecord_(t, {
      Response_ID: nextId_(t, 'Response_ID', 'SR-', 5),
      Report_ID: reportId,
      Staff_Code: row.Staff_Code,
      Responder_Code: ctx.me.code,
      Status: status,
      Reply: text,
      Response_Date: n.date,
      Response_Time: n.time
    });
  });
  return reflectionDetail_(ctx, reportId);
}

/** One person's recent entries, newest first. */
function getPersonReflections(token, code) {
  const ctx = requireReader_(token);
  if (ctx.readable.indexOf(code) === -1) throw new Error('Person not found.');
  const latest = latestResponses_();
  return readTable_(TAB_REPORT).rows
    .filter(function (r) { return r.Staff_Code === code; })
    .map(function (r) { return tileOut_(r, ctx.staff, latest); })
    .sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; })
    .slice(0, 30);
}

/** People the reader can pick in "By person". */
function getReflectionPeople(token) {
  const ctx = requireReader_(token);
  return ctx.readable.map(function (c) { return personOut_(ctx.staff.byCode[c]); })
    .sort(function (a, b) { return a.name.localeCompare(b.name); });
}

// ═══════════════════════════════════════════════════════════════════════════
// DAILY DIGEST EMAIL — time-driven trigger, ~19:15–19:45 IST (flag: reflectionDigest)
// Sent to every owner (empty Reports_To) with an Email in Config_EC.
// Skipped on days nobody used the app; never sent twice for the same date (Digest_Log_EC).
// ═══════════════════════════════════════════════════════════════════════════
/** Trigger entry point. The trigger passes an event object, so "force" is never read from arguments. */
function sendReflectionDigest() {
  return digest_(false);
}

/** Editor utility: send today's digest now, even if one was already sent. */
function testDigestNow() {
  return digest_(true);
}

function digest_(force) {
  if (!FEATURES_ENABLED.reflectionDigest) { Logger.log('reflectionDigest is switched off.'); return 'off'; }
  const staff = getStaff_();
  const n = now_();
  const digestId = 'DG-' + n.date.replace(/-/g, '');
  const owners = staff.list.filter(function (s) { return isActive_(s) && isAdmin_(s); });
  const recipients = owners.map(function (s) { return s.email; }).filter(function (e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e); });
  if (!recipients.length) { Logger.log('No owner has an Email in ' + TAB_CONFIG + '. Digest not sent.'); return 'no-recipient'; }

  if (!force && readTable_(TAB_DIGEST).rows.some(function (r) { return r.Digest_ID === digestId; })) {
    Logger.log('Digest for ' + n.date + ' already sent.'); return 'duplicate';
  }

  const codes = staff.list.filter(function (s) { return isActive_(s) && !isAdmin_(s); }).map(function (s) { return s.code; });
  const g = reflectionDayGroups_(n.date, codes, staff);
  if (!g.written.length && !g.notWritten.length) {
    Logger.log('Nobody used the app on ' + n.date + '. Digest skipped.'); return 'no-activity';
  }

  let appUrl = '';
  try { appUrl = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  const subject = 'Daily reflections ' + Utilities.formatDate(new Date(), TZ, 'd MMM') + ': ' +
    g.written.length + ' written, ' + g.notWritten.length + ' not written' +
    (g.noActivity.length ? ', ' + g.noActivity.length + ' no activity' : '');
  const html = digestHtml_(g, appUrl);
  const plain = digestPlain_(g, appUrl);
  MailApp.sendEmail({ to: recipients.join(','), subject: subject, htmlBody: html, body: plain, name: APP_NAME });

  withLock_(function () {
    const t = readTable_(TAB_DIGEST);
    appendRecord_(t, {
      Digest_ID: force ? nextId_(t, 'Digest_ID', digestId + '-T', 2) : digestId,
      Digest_Date: n.date,
      Sent_Time: now_().time,
      Recipients: recipients.join(', '),
      Written_Count: String(g.written.length),
      Not_Written_Count: String(g.notWritten.length),
      No_Activity_Count: String(g.noActivity.length)
    });
  });
  Logger.log('Digest sent to ' + recipients.join(', ') + ': ' + subject);
  return 'sent';
}

function h_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function digestHtml_(g, appUrl) {
  const section = function (title, color, body) {
    return '<h3 style="margin:22px 0 8px;font:600 15px Arial,sans-serif;color:' + color + '">' + title + '</h3>' + body;
  };
  const names = function (list) {
    return list.length ? '<p style="margin:0;font:14px Arial,sans-serif;color:#1F2937">' +
      list.map(function (p) { return h_(p.name) + ' <span style="color:#6B7280">(' + h_(p.role) + ')</span>'; }).join('<br>') + '</p>'
      : '<p style="margin:0;font:14px Arial,sans-serif;color:#6B7280">None</p>';
  };
  const written = g.written.length ? g.written.map(function (t) {
    return '<div style="border:1px solid #DDE3EA;border-radius:8px;padding:10px 12px;margin-bottom:8px;font:14px Arial,sans-serif;color:#1F2937">' +
      '<b>' + h_(t.name) + '</b> <span style="color:#6B7280">· ' + h_(t.role) + ' · ' + h_(t.submittedTime) + '</span>' +
      (t.support ? '<div style="margin-top:6px;padding:6px 8px;background:#FFF7E6;border-radius:6px"><b style="color:#92400E">Support required:</b> ' + h_(t.support) + '</div>' : '') +
      (t.challenges ? '<div style="margin-top:6px"><b>Challenges:</b> ' + h_(t.challenges) + '</div>' : '') +
      '</div>';
  }).join('') : '<p style="margin:0;font:14px Arial,sans-serif;color:#6B7280">None</p>';
  return '<div style="max-width:620px">' +
    '<div style="background:#12263A;color:#FFFFFF;padding:14px 16px;border-radius:8px 8px 0 0;font:600 16px Arial,sans-serif">EC Smart · Daily reflections · ' +
    h_(Utilities.formatDate(new Date(), TZ, 'EEE d MMM yyyy')) + '</div>' +
    '<div style="padding:4px 16px 16px;border:1px solid #DDE3EA;border-top:none;border-radius:0 0 8px 8px">' +
    section('Not written (' + g.notWritten.length + ')', '#B42318', names(g.notWritten)) +
    section('Written (' + g.written.length + ')', '#15803D', written) +
    section('No activity today (' + g.noActivity.length + ')', '#6B7280', names(g.noActivity) +
      '<p style="margin:6px 0 0;font:12px Arial,sans-serif;color:#6B7280">Did not open the portal today: on leave, or not using it.</p>') +
    (appUrl ? '<p style="margin:22px 0 0"><a href="' + h_(appUrl) + '" style="display:inline-block;background:#0E6B74;color:#FFFFFF;padding:9px 16px;border-radius:8px;text-decoration:none;font:600 14px Arial,sans-serif">Open Team Reflections</a></p>' : '') +
    '</div></div>';
}

function digestPlain_(g, appUrl) {
  const list = function (l) { return l.length ? l.map(function (p) { return '- ' + p.name + ' (' + p.role + ')'; }).join('\n') : '- None'; };
  return 'NOT WRITTEN (' + g.notWritten.length + ')\n' + list(g.notWritten) +
    '\n\nWRITTEN (' + g.written.length + ')\n' +
    (g.written.length ? g.written.map(function (t) {
      return '- ' + t.name + ' (' + t.submittedTime + ')' + (t.support ? '\n  Support required: ' + t.support : '');
    }).join('\n') : '- None') +
    '\n\nNO ACTIVITY TODAY (' + g.noActivity.length + ')\n' + list(g.noActivity) +
    (appUrl ? '\n\n' + appUrl : '');
}

/** Editor utility: (re)creates the daily digest trigger. Run once after deploying. */
function installDigestTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendReflectionDigest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendReflectionDigest')
    .timeBased().everyDays(1).atHour(DIGEST_HOUR).nearMinute(DIGEST_NEAR_MINUTE).inTimezone(TZ)
    .create();
  Logger.log('Daily digest trigger set for about ' + DIGEST_HOUR + ':' + String(DIGEST_NEAR_MINUTE).padStart(2, '0') + ' IST.');
}

// ═══════════════════════════════════════════════════════════════════════════
// ASSIGNED TASKS — fully built, inert while FEATURES_ENABLED.assignedTasks is false.
// Normal task management: due dates and overdue flags are allowed here (SOP 10).
// ═══════════════════════════════════════════════════════════════════════════
function taskOut_(r, staff, commentCounts, nowKey) {
  const dueKey = r.Due_Date ? r.Due_Date + ' ' + (r.Due_Time || '23:59') : '';
  return {
    id: r.Task_ID,
    text: r.Text,
    byCode: r.Assigned_By_Code,
    byName: nameOf_(staff, r.Assigned_By_Code),
    toCode: r.Assigned_To_Code,
    toName: nameOf_(staff, r.Assigned_To_Code),
    assignedDate: r.Assigned_Date,
    assignedTime: r.Assigned_Time,
    dueDate: r.Due_Date,
    dueTime: r.Due_Time,
    status: r.Status || 'Open',
    completedDate: r.Completed_Date,
    completedTime: r.Completed_Time,
    reopenedBy: r.Reopened_By_Code ? nameOf_(staff, r.Reopened_By_Code) : '',
    overdue: !!dueKey && (r.Status || 'Open') !== 'Done' && dueKey < nowKey,
    comments: commentCounts[r.Task_ID] || 0
  };
}

function listTasks_(me, staff) {
  const counts = {};
  readTable_(TAB_TASK_COMMENTS).rows.forEach(function (r) { counts[r.Task_ID] = (counts[r.Task_ID] || 0) + 1; });
  const nowKey = now_().key;
  const cutoff = daysAgo_(14);
  const rows = readTable_(TAB_TASKS).rows.filter(function (r) {
    return (r.Status || 'Open') !== 'Done' || r.Completed_Date >= cutoff;
  });
  return {
    toMe: rows.filter(function (r) { return r.Assigned_To_Code === me.code; })
      .map(function (r) { return taskOut_(r, staff, counts, nowKey); }),
    byMe: rows.filter(function (r) { return r.Assigned_By_Code === me.code; })
      .map(function (r) { return taskOut_(r, staff, counts, nowKey); })
  };
}

function getAssignedTasks(token) {
  assertFeature_('assignedTasks');
  const ctx = requireUser_(token);
  return listTasks_(ctx.me, ctx.staff);
}

function assignTask(token, input) {
  assertFeature_('assignedTasks');
  const ctx = requireUser_(token);
  input = input || {};
  const toCode = String(input.toCode || '');
  const visible = visibleCodes_(ctx.me, ctx.staff);
  if (toCode === ctx.me.code || visible.indexOf(toCode) === -1) throw new Error('You can only assign to your own team.');
  const text = cleanText_(input.text, MAX_TEXT, 'Task', true);
  const dueDate = cleanDate_(input.dueDate, 'Due date');
  const dueTime = cleanTime_(input.dueTime, 'Due time');
  if (dueTime && !dueDate) throw new Error('Add a due date for the due time.');
  return withLock_(function () {
    const t = readTable_(TAB_TASKS);
    const n = now_();
    appendRecord_(t, {
      Task_ID: nextId_(t, 'Task_ID', 'AT-', 5),
      Assigned_By_Code: ctx.me.code,
      Assigned_To_Code: toCode,
      Text: text,
      Assigned_Date: n.date,
      Assigned_Time: n.time,
      Due_Date: dueDate,
      Due_Time: dueTime,
      Status: 'Open'
    });
    return listTasks_(ctx.me, ctx.staff);
  });
}

/** Assignee moves Open ↔ In Progress → Done. Assigner may reopen a Done task. */
function updateAssignedTaskStatus(token, taskId, status) {
  assertFeature_('assignedTasks');
  const ctx = requireUser_(token);
  if (['Open', 'In Progress', 'Done'].indexOf(status) === -1) throw new Error('Unknown status.');
  return withLock_(function () {
    const t = readTable_(TAB_TASKS);
    const row = t.rows.filter(function (r) { return r.Task_ID === taskId; })[0];
    if (!row) throw new Error('Task not found.');
    const isAssignee = row.Assigned_To_Code === ctx.me.code;
    const isAssigner = row.Assigned_By_Code === ctx.me.code;
    const current = row.Status || 'Open';
    const n = now_();
    if (isAssignee) {
      const patch = { Status: status };
      if (status === 'Done') { patch.Completed_Date = n.date; patch.Completed_Time = n.time; }
      else { patch.Completed_Date = ''; patch.Completed_Time = ''; }
      updateRecord_(t, row._row, patch);
    } else if (isAssigner && current === 'Done' && status === 'Open') {
      updateRecord_(t, row._row, { Status: 'Open', Completed_Date: '', Completed_Time: '', Reopened_By_Code: ctx.me.code });
    } else {
      throw new Error('You cannot change this task.');
    }
    return listTasks_(ctx.me, ctx.staff);
  });
}

function taskForParticipant_(ctx, taskId) {
  const row = readTable_(TAB_TASKS).rows.filter(function (r) { return r.Task_ID === taskId; })[0];
  if (!row || (row.Assigned_To_Code !== ctx.me.code && row.Assigned_By_Code !== ctx.me.code)) {
    throw new Error('Task not found.');
  }
  return row;
}

function getTaskComments(token, taskId) {
  assertFeature_('assignedTasks');
  const ctx = requireUser_(token);
  taskForParticipant_(ctx, taskId);
  return readTable_(TAB_TASK_COMMENTS).rows
    .filter(function (r) { return r.Task_ID === taskId; })
    .map(function (r) {
      return {
        id: r.Comment_ID,
        author: nameOf_(ctx.staff, r.Author_Code),
        mine: r.Author_Code === ctx.me.code,
        text: r.Text,
        date: r.Comment_Date,
        time: r.Comment_Time
      };
    });
}

function addTaskComment(token, taskId, text) {
  assertFeature_('assignedTasks');
  const ctx = requireUser_(token);
  taskForParticipant_(ctx, taskId);
  const body = cleanText_(text, MAX_TEXT * 2, 'Comment', true);
  withLock_(function () {
    const t = readTable_(TAB_TASK_COMMENTS);
    const n = now_();
    appendRecord_(t, {
      Comment_ID: nextId_(t, 'Comment_ID', 'TC-', 5),
      Task_ID: taskId,
      Author_Code: ctx.me.code,
      Text: body,
      Comment_Date: n.date,
      Comment_Time: n.time
    });
  });
  return getTaskComments(token, taskId);
}

// ═══════════════════════════════════════════════════════════════════════════
// TEAM COMPLETION — fully built, inert while FEATURES_ENABLED.teamCompletion is false.
// Placeholder staff excluded. Reflection shows written / not written only — no lateness.
// ═══════════════════════════════════════════════════════════════════════════
function getTeamCompletion(token, date) {
  assertFeature_('teamCompletion');
  const ctx = requireUser_(token);
  const d = cleanDate_(date, 'Date') || now_().date;
  const nowKey = now_().key;
  const team = visibleCodes_(ctx.me, ctx.staff).filter(function (c) { return c !== ctx.me.code; });
  if (!team.length) throw new Error('No one reports to you.');
  const todos = readTable_(TAB_TODO).rows;
  const reports = readTable_(TAB_REPORT).rows;
  return team.map(function (code) {
    const s = ctx.staff.byCode[code];
    const mine = todos.filter(function (r) { return r.Staff_Code === code; });
    const open = mine.filter(function (r) { return (r.Status || 'Open') === 'Open'; });
    return {
      code: code,
      name: s.name,
      role: s.role,
      openTodos: open.length,
      overdueTodos: open.filter(function (r) {
        return r.Due_Date && (r.Due_Date + ' ' + (r.Due_Time || '23:59')) < nowKey;
      }).length,
      doneOnDate: mine.filter(function (r) { return r.Status === 'Done' && r.Completed_Date === d; }).length,
      reflectionWritten: reports.some(function (r) { return r.Staff_Code === code && r.Report_Date === d; })
    };
  });
}

// ═══════════════════════════════════════════════════════════════════════════
// EDITOR UTILITIES — run from the Apps Script editor, not from the web app
// ═══════════════════════════════════════════════════════════════════════════
/** Checks every tab exists with the expected headers. Run after setup or any sheet edit. */
function setupCheck() {
  const expected = {};
  expected[TAB_CONFIG] = ['Role', 'Name', 'PIN', 'Type', 'Staff_Status', 'Staff_Code', 'Reports_To', 'Date_of_Joining', 'Email'];
  expected[TAB_ACTIVITY] = ['Activity_ID', 'Staff_Code', 'Activity_Date', 'First_Seen_Time'];
  expected[TAB_SUPPORT] = ['Response_ID', 'Report_ID', 'Staff_Code', 'Responder_Code', 'Status', 'Reply', 'Response_Date', 'Response_Time'];
  expected[TAB_DIGEST] = ['Digest_ID', 'Digest_Date', 'Sent_Time', 'Recipients', 'Written_Count', 'Not_Written_Count', 'No_Activity_Count'];
  expected[TAB_TODO] = ['Item_ID', 'Staff_Code', 'Created_Date', 'Created_Time', 'Text', 'Due_Date', 'Due_Time', 'Remind_Date', 'Remind_Time', 'Status', 'Completed_Date', 'Completed_Time', 'Cancelled_Date'];
  expected[TAB_REPORT] = ['Report_ID', 'Staff_Code', 'Report_Date', 'Submitted_Date', 'Submitted_Time', 'Challenges_Blockers', 'Support_Required', 'Learning_Note', 'Last_Edited_Date', 'Last_Edited_Time'];
  expected[TAB_ALLOC] = ['Alloc_ID', 'Report_ID', 'Staff_Code', 'Report_Date', 'Activity_Category', 'Hours', 'Note', 'Status'];
  expected[TAB_ANN] = ['Ann_ID', 'Category', 'Title', 'Body', 'Doc_Link', 'Posted_By_Code', 'Posted_Date', 'Posted_Time', 'Expiry_Date', 'Pinned', 'Status'];
  expected[TAB_ANN_READS] = ['Read_ID', 'Ann_ID', 'Staff_Code', 'Read_Date', 'Read_Time'];
  expected[TAB_LIBRARY] = ['Doc_ID', 'Title', 'Category', 'Audience_Type', 'Audience_Value', 'Doc_Link', 'Version', 'Added_By_Code', 'Added_Date', 'Status'];
  expected[TAB_TASKS] = ['Task_ID', 'Assigned_By_Code', 'Assigned_To_Code', 'Text', 'Assigned_Date', 'Assigned_Time', 'Due_Date', 'Due_Time', 'Status', 'Completed_Date', 'Completed_Time', 'Reopened_By_Code'];
  expected[TAB_TASK_COMMENTS] = ['Comment_ID', 'Task_ID', 'Author_Code', 'Text', 'Comment_Date', 'Comment_Time'];
  const problems = [];
  Object.keys(expected).forEach(function (tab) {
    const sh = ss_().getSheetByName(tab);
    if (!sh) { problems.push('Missing tab ' + tab); return; }
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    expected[tab].forEach(function (h) { if (headers.indexOf(h) === -1) problems.push(tab + ' missing column ' + h); });
  });
  const staff = getStaff_();
  const codes = {};
  const roles = {};
  staff.list.forEach(function (s) {
    if (codes[s.code]) problems.push('Duplicate Staff_Code ' + s.code);
    if (roles[s.role]) problems.push('Duplicate Role ' + s.role);
    codes[s.code] = true;
    roles[s.role] = true;
  });
  staff.list.forEach(function (s) {
    if (s.reportsTo && !roles[s.reportsTo]) problems.push(s.name + ' reports to unknown role "' + s.reportsTo + '"');
  });
  if (FEATURES_ENABLED.reflectionDigest &&
      !staff.list.some(function (s) { return isActive_(s) && isAdmin_(s) && s.email; })) {
    problems.push('No owner (empty Reports_To) has an Email in ' + TAB_CONFIG + ', so the daily digest has no recipient.');
  }
  Logger.log(problems.length ? 'PROBLEMS:\n' + problems.join('\n') : 'All tabs and headers OK. Roster: ' + staff.list.length + ' people.');
}

/** Logs a new, unused 6-character Staff_Code for a new joiner. */
function generateStaffCode() {
  const used = getStaff_().byCode;
  let code;
  do {
    code = '';
    for (let i = 0; i < 6; i++) code += STAFF_CODE_ALPHABET.charAt(Math.floor(Math.random() * STAFF_CODE_ALPHABET.length));
  } while (used[code]);
  Logger.log('New Staff_Code: ' + code);
  return code;
}
