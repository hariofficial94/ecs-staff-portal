/**
 * ECS Staff Portal — server (Google Apps Script web app)
 * Client: EC Smart Automations Pvt Ltd
 *
 * Live features (v1): To Do (with due date + in-app reminders), Daily Reflection,
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
        reportsTo: r.Reports_To
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
  return { me: me, staff: staff };
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
      hasTeam: visible.length > 1
    },
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
    recentDates: recent
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
  expected[TAB_CONFIG] = ['Role', 'Name', 'PIN', 'Type', 'Staff_Status', 'Staff_Code', 'Reports_To', 'Date_of_Joining'];
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
