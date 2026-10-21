/**
 * TeamBoard — a free Trello + ClickUp style task manager for small teams.
 * Backend: Google Apps Script. Database: the Google Sheet this script is bound to.
 *
 * Setup: run setup() once from the editor, then Deploy > New deployment > Web app.
 */

const APP_NAME = 'TeamBoard';
const NOTIFY_BY_EMAIL = true; // email people when a task is assigned to them

const SCHEMA = {
  Boards:   ['id', 'name', 'color', 'statuses', 'createdBy', 'createdAt'],
  Tasks:    ['id', 'boardId', 'title', 'description', 'status', 'priority', 'assignee',
             'dueDate', 'tags', 'checklist', 'order', 'createdBy', 'createdAt', 'updatedAt'],
  Comments: ['id', 'taskId', 'author', 'text', 'createdAt'],
  Members:  ['email', 'name', 'role'],
  Activity: ['ts', 'user', 'action', 'taskId', 'detail']
};

// ---------- Web app entry ----------

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // allows embedding in Google Sites
}

// ---------- One-time setup ----------

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange('A:Z').setNumberFormat('@'); // keep everything as plain text (no date auto-conversion)
    if (sh.getLastRow() === 0) {
      sh.appendRow(SCHEMA[name]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, SCHEMA[name].length).setFontWeight('bold').setBackground('#ececfb');
    }
  });
  const blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);

  const me = Session.getEffectiveUser().getEmail();
  if (sheet_('Members').getLastRow() === 1) {
    sheet_('Members').appendRow([me, me.split('@')[0], 'admin']);
  }
  if (sheet_('Boards').getLastRow() === 1) {
    sheet_('Boards').appendRow([uid_(), 'General', '#5b5bd6', 'To Do,In Progress,Review,Done', me, now_()]);
  }
}

// ---------- API (called from Index.html via google.script.run) ----------

function getData() {
  const counts = {};
  readAll_('Comments').forEach(function (c) { counts[c.taskId] = (counts[c.taskId] || 0) + 1; });
  return {
    me: Session.getActiveUser().getEmail().trim().toLowerCase(),
    appUrl: ScriptApp.getService().getUrl(),
    members: readAll_('Members'),
    boards: readAll_('Boards'),
    tasks: readAll_('Tasks'),
    commentCounts: counts
  };
}

function saveTask(task, actor) {
  return withLock_(function () {
    const user = who_(actor);
    const isNew = !task.id;
    const old = isNew ? null : findById_('Tasks', task.id);
    if (isNew) {
      task.id = uid_();
      task.createdBy = user;
      task.createdAt = now_();
    } else if (old) {
      task.createdBy = old.obj.createdBy;
      task.createdAt = old.obj.createdAt;
    }
    task.assignee = String(task.assignee || '').trim().toLowerCase();
    task.updatedAt = now_();
    upsert_('Tasks', task);
    log_(user, isNew ? 'created' : 'updated', task.id, task.title);

    const prevAssignee = old ? old.obj.assignee : '';
    if (task.assignee && task.assignee !== prevAssignee && task.assignee !== user) notifyAssignee_(task, user);
    return task;
  });
}

function moveTask(id, status, order, actor) {
  return withLock_(function () {
    const found = findById_('Tasks', id);
    if (!found) throw new Error('Task not found (it may have been deleted).');
    const t = found.obj;
    const from = t.status;
    t.status = status;
    t.order = String(order);
    t.updatedAt = now_();
    upsert_('Tasks', t);
    if (from !== status) log_(who_(actor), 'moved', id, from + ' → ' + status);
    return t;
  });
}

function deleteTask(id, actor) {
  return withLock_(function () {
    deleteWhere_('Tasks', function (r) { return r.id === id; });
    deleteWhere_('Comments', function (r) { return r.taskId === id; });
    log_(who_(actor), 'deleted', id, '');
    return true;
  });
}

function getComments(taskId) {
  return readAll_('Comments').filter(function (c) { return c.taskId === taskId; });
}

function addComment(taskId, text, actor) {
  return withLock_(function () {
    const c = { id: uid_(), taskId: taskId, author: who_(actor), text: text, createdAt: now_() };
    upsert_('Comments', c);
    log_(c.author, 'commented', taskId, text.slice(0, 80));
    return c;
  });
}

function saveBoard(board, actor) {
  return withLock_(function () {
    if (!board.id) {
      board.id = uid_();
      board.createdBy = who_(actor);
      board.createdAt = now_();
    }
    upsert_('Boards', board);
    return board;
  });
}

function deleteBoard(id, actor) {
  return withLock_(function () {
    const taskIds = readAll_('Tasks').filter(function (t) { return t.boardId === id; }).map(function (t) { return t.id; });
    deleteWhere_('Boards', function (r) { return r.id === id; });
    deleteWhere_('Tasks', function (r) { return r.boardId === id; });
    deleteWhere_('Comments', function (r) { return taskIds.indexOf(r.taskId) !== -1; });
    log_(who_(actor), 'deleted board', id, '');
    return true;
  });
}

function saveMember(member) {
  return withLock_(function () {
    member.email = String(member.email || '').trim().toLowerCase();
    if (!member.email) throw new Error('Email is required.');
    upsert_('Members', member, 'email');
    return member;
  });
}

function deleteMember(email) {
  return withLock_(function () {
    deleteWhere_('Members', function (r) { return r.email === email; });
    return true;
  });
}

// ---------- Sheet helpers ----------

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" is missing. Run setup() in the Apps Script editor.');
  return sh;
}

function readAll_(name) {
  const values = sheet_(name).getDataRange().getValues();
  const head = values.shift();
  return values
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
    .map(function (r) {
      const o = {};
      head.forEach(function (h, i) {
        let v = r[i];
        if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        o[h] = v === null || v === undefined ? '' : String(v);
      });
      return o;
    });
}

function findById_(name, id, key) {
  key = key || SCHEMA[name][0];
  const rows = readAll_(name);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][key] === id) return { row: i + 2, obj: rows[i] };
  }
  return null;
}

function upsert_(name, obj, key) {
  key = key || SCHEMA[name][0];
  const sh = sheet_(name);
  const head = SCHEMA[name];
  const rowValues = head.map(function (h) { return clean_(obj[h]); });
  const found = findById_(name, obj[key], key);
  if (found) sh.getRange(found.row, 1, 1, head.length).setValues([rowValues]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, head.length).setNumberFormat('@').setValues([rowValues]);
}

function deleteWhere_(name, pred) {
  const sh = sheet_(name);
  const rows = readAll_(name);
  for (let i = rows.length - 1; i >= 0; i--) {
    if (pred(rows[i])) sh.deleteRow(i + 2);
  }
}

function clean_(v) {
  v = v === null || v === undefined ? '' : String(v);
  return /^[=+@]/.test(v) ? "'" + v : v; // stop text being treated as a formula
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function who_(actor) {
  return (Session.getActiveUser().getEmail() || actor || 'unknown').trim().toLowerCase();
}

function log_(user, action, taskId, detail) {
  sheet_('Activity').appendRow([now_(), user, action, taskId, clean_(detail)]);
}

function notifyAssignee_(task, by) {
  if (!NOTIFY_BY_EMAIL || task.assignee.indexOf('@') === -1) return;
  try {
    MailApp.sendEmail({
      to: task.assignee,
      subject: '[' + APP_NAME + '] Assigned to you: ' + task.title,
      htmlBody: '<p><b>' + by + '</b> assigned you a task:</p>' +
        '<p style="font-size:16px"><b>' + escape_(task.title) + '</b></p>' +
        (task.dueDate ? '<p>Due: ' + task.dueDate + '</p>' : '') +
        (task.description ? '<p>' + escape_(task.description).replace(/\n/g, '<br>') + '</p>' : '') +
        '<p><a href="' + ScriptApp.getService().getUrl() + '">Open ' + APP_NAME + '</a></p>'
    });
  } catch (e) {
    console.warn('Email failed: ' + e); // daily mail quota reached, etc. Never block the save.
  }
}

function escape_(s) {
  return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
}

function uid_() {
  return Utilities.getUuid().slice(0, 8) + Date.now().toString(36);
}

function now_() {
  return new Date().toISOString();
}
