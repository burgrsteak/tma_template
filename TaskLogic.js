/**
 * Router to safely serve HTML pages.
 */
function API_getPage(pageName) {
  const allowedPages = ['Dashboard', 'ShiftReport', 'Analytics', 'Settings'];
  if (!allowedPages.includes(pageName)) {
    throw new Error('Unauthorized page request or page does not exist.');
  }
  return include('Pages/' + pageName);
}

/**
 * Checks if the user has an active 12-hour session on page load/refresh.
 * Returns { active: false } — never throws — if the user is not in the Users sheet.
 *
 * FIX (2026-05-15): Uses colIdx() helper for safe column lookups so a header
 * mismatch can never silently produce userRecord = null.  When a valid session
 * cookie exists but the email is absent from the Users sheet the function now
 * attempts a one-time auto-registration as Dev before giving up, preventing
 * the developer from being permanently locked out after a fresh DB init.
 */
function API_checkCurrentSession() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return JSON.stringify({ active: false });

  const props = PropertiesService.getUserProperties();
  const sessionStart = props.getProperty('sessionStart');

  if (!sessionStart) return JSON.stringify({ active: false });

  const now = new Date().getTime();
  const start = parseInt(sessionStart, 10);
  const hoursElapsed = (now - start) / (1000 * 60 * 60);

  if (hoursElapsed >= 12) {
    props.deleteProperty('sessionStart');
    return JSON.stringify({ active: false, expired: true });
  }

  try {
    const usersSheet = getTable('Users');
    const data = usersSheet.getDataRange().getValues();
    const headers = data[0];

    // Use colIdx() so a header mismatch throws a clear error instead of
    // silently returning -1 and wiping the session.
    const emailIdx = colIdx(headers, 'USER_EMAIL');
    const nameIdx  = colIdx(headers, 'USER_NAME',  false); // optional — fall back gracefully
    const roleIdx  = colIdx(headers, 'USER_ROLE',  false);

    let userRecord = null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][emailIdx]).trim().toLowerCase() === email.trim().toLowerCase()) {
        userRecord = {
          name:  nameIdx  > -1 ? (data[i][nameIdx]  || email) : email,
          email: email,
          role:  roleIdx  > -1 ? (data[i][roleIdx]  || 'User') : 'User'
        };
        break;
      }
    }

    // ----------------------------------------------------------------
    // DEVELOPER SAFETY NET
    // If the session cookie is valid but the email is missing from the
    // Users sheet (e.g. after a DB re-init or accidental row deletion)
    // auto-register the user as Dev so the owner is never locked out.
    // The client will receive active:true and can fix the Users sheet
    // from inside the app.
    // ----------------------------------------------------------------
    if (!userRecord) {
      try {
        const nameColIdx = nameIdx > -1 ? nameIdx : -1;
        const roleColIdx = roleIdx > -1 ? roleIdx : -1;

        // Build a sparse row aligned to the existing header order
        const newRow = headers.map((h, idx) => {
          if (idx === emailIdx)    return email;
          if (idx === nameColIdx)  return email.split('@')[0];
          if (idx === roleColIdx)  return 'Dev';
          return '';
        });
        usersSheet.appendRow(newRow);
        SpreadsheetApp.flush();

        userRecord = {
          name:  email.split('@')[0],
          email: email,
          role:  'Dev'
        };
        console.warn('API_checkCurrentSession: auto-registered missing user as Dev — ' + email);
      } catch (regErr) {
        // Auto-registration failed (permissions, locked sheet, etc.).
        // Fall back to killing the session gracefully rather than crashing.
        console.error('API_checkCurrentSession: auto-register failed: ' + regErr.message);
        props.deleteProperty('sessionStart');
        props.deleteProperty('currentLogId');
        return JSON.stringify({ active: false, reason: 'user_not_found' });
      }
    }

    return JSON.stringify({ active: true, user: userRecord });
  } catch (e) {
    console.error('API_checkCurrentSession error: ' + e.message);
    return JSON.stringify({ active: false, reason: 'error', detail: e.message });
  }
}

/**
 * Logs the user in, starts the 12-hour clock, and records timeIn.
 * Takes a permanent snapshot of the user's Team and Shift at login.
 */
function API_loginUser() {
  return withLock(() => {
    const email = Session.getActiveUser().getEmail().trim().toLowerCase();

    const usersSheet = getTable('Users');
    const data = usersSheet.getDataRange().getValues();
    const headers = data[0];

    const emailIdx      = colIdx(headers, 'USER_EMAIL');
    const nameIdx       = colIdx(headers, 'USER_NAME',       false);
    const roleIdx       = colIdx(headers, 'USER_ROLE',       false);
    const teamIdx       = colIdx(headers, 'USER_TEAM',       false);
    const shiftStartIdx = colIdx(headers, 'USER_SHIFT_START',false);
    const shiftEndIdx   = colIdx(headers, 'USER_SHIFT_END',  false);

    let userRecord = null;
    let currentTeam  = '--';
    let currentStart = '';
    let currentEnd   = '';

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][emailIdx]).trim().toLowerCase() === email) {
        userRecord   = {
          name:  nameIdx > -1 ? (data[i][nameIdx] || email) : email,
          email: email,
          role:  roleIdx > -1 ? (data[i][roleIdx] || 'User') : 'User'
        };
        currentTeam  = teamIdx       > -1 ? data[i][teamIdx]       : '--';
        currentStart = shiftStartIdx > -1 ? data[i][shiftStartIdx] : '';
        currentEnd   = shiftEndIdx   > -1 ? data[i][shiftEndIdx]   : '';
        break;
      }
    }

    if (userRecord) {
      const props = PropertiesService.getUserProperties();
      const now = new Date();

      props.setProperty('sessionStart', now.getTime().toString());

      const logSheet = getTable('SessionLogs');
      const logData = logSheet.getDataRange().getValues();

      // Auto-close ghost sessions
      if (logData.length > 1) {
        const lHeaders    = logData[0].map(h => String(h).trim().toLowerCase());
        const lEmailIdx   = lHeaders.indexOf(COLUMN_MAP.SESSION_EMAIL.toLowerCase());
        const lTimeOutIdx = lHeaders.indexOf(COLUMN_MAP.SESSION_TIME_OUT.toLowerCase());
        const lRemarkIdx  = lHeaders.indexOf(COLUMN_MAP.SESSION_REMARK.toLowerCase());

        for (let i = 1; i < logData.length; i++) {
          if (String(logData[i][lEmailIdx]).trim().toLowerCase() === email &&
              String(logData[i][lTimeOutIdx]).trim() === '') {
            logSheet.getRange(i + 1, lTimeOutIdx + 1).setValue(now.toISOString());
            logSheet.getRange(i + 1, lRemarkIdx  + 1).setValue('System Auto-Close: Duplicate Session Detected');
          }
        }
      }

      const logId = generateUUID();
      props.setProperty('currentLogId', logId);
      logSheet.appendRow([logId, email, now.toISOString(), '', '', currentTeam, currentStart, currentEnd]);
      SpreadsheetApp.flush();

      return JSON.stringify({ success: true, authorized: true, user: userRecord });
    } else {
      return JSON.stringify({ success: true, authorized: false, email: email });
    }
  });
}

/**
 * Logs the user out and records timeOut.
 */
function API_logoutUser() {
  return withLock(() => {
    const props = PropertiesService.getUserProperties();
    const logId = props.getProperty('currentLogId');

    if (logId) {
      const logSheet = getTable('SessionLogs');
      const data = logSheet.getDataRange().getValues();
      const headers = data[0].map(h => String(h).trim().toLowerCase());
      const idIdx      = headers.indexOf(COLUMN_MAP.SESSION_ID.toLowerCase());
      const timeOutIdx = headers.indexOf(COLUMN_MAP.SESSION_TIME_OUT.toLowerCase());

      for (let i = 1; i < data.length; i++) {
        if (data[i][idIdx] === logId) {
          logSheet.getRange(i + 1, timeOutIdx + 1).setValue(new Date().toISOString());
          break;
        }
      }
      SpreadsheetApp.flush();
    }

    props.deleteProperty('sessionStart');
    props.deleteProperty('currentLogId');
    return JSON.stringify({ success: true });
  });
}

// ---------------------------------------------------------------------------
// SECURITY HELPER
// ---------------------------------------------------------------------------

/**
 * Returns the numeric Tier of the calling user.
 * Tier 0 = highest privilege (Admin/Dev), Tier 3 = lowest (User).
 * Returns Tier 3 safely if the user is not found.
 */
function getUserTier_(email) {
  try {
    const usersData = getTable('Users').getDataRange().getValues();
    const headers   = usersData[0];
    const emailCol  = colIdx(headers, 'USER_EMAIL');
    const roleCol   = colIdx(headers, 'USER_ROLE', false);

    let userRole = '';
    if (emailCol !== -1 && roleCol !== -1) {
      for (let i = 1; i < usersData.length; i++) {
        if (String(usersData[i][emailCol]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
          userRole = String(usersData[i][roleCol] || '').trim();
          break;
        }
      }
    }

    const props = PropertiesService.getScriptProperties();
    const settingsStr = props.getProperty('DYNAMIC_APP_SETTINGS');
    let currentTiers = { '0': ['Admin','Dev'], '1': ['Manager'], '2': ['Lead','QA'], '3': ['User'] };
    if (settingsStr) {
      try {
        const conf = JSON.parse(settingsStr);
        if (conf.roleTiers) currentTiers = conf.roleTiers;
      } catch(e) {}
    }

    let userTier = 3;
    for (const [tierLvl, rolesArray] of Object.entries(currentTiers)) {
      if (Array.isArray(rolesArray) && rolesArray.some(r => r && r.toLowerCase() === userRole.toLowerCase())) {
        userTier = parseInt(tierLvl);
        break;
      }
    }
    return userTier;
  } catch(e) {
    console.error('getUserTier_ error for ' + email + ': ' + e.message);
    return 3; // Safest fallback — treat unknown users as lowest tier
  }
}

// ---------------------------------------------------------------------------
// EMAIL HELPERS
// ---------------------------------------------------------------------------

function sendPriorityEmail(task) {
  if (task.priority !== 'High' || !task.assignedTo) return;
  const subject = '🚨 HIGH PRIORITY: ' + task.title;
  const body = 'Hi,\n\nYou have been assigned a High Priority task: "' + task.title + '".\n\n' +
               'Link: ' + ScriptApp.getService().getUrl() + '\n\n' +
               'Priority: ' + task.priority + '\nDue: ' + (task.deadline || 'N/A');
  try { MailApp.sendEmail(task.assignedTo, subject, body); }
  catch (e) { console.error('Email failed: ' + e.message); }
}

// ---------------------------------------------------------------------------
// NOTIFICATIONS
// ---------------------------------------------------------------------------

/**
 * Writes a notification record for a user.
 */
function pushNotification_(userEmail, message) {
  try {
    const sheet = getTable('Notifications');
    sheet.appendRow([generateUUID(), userEmail, message, false, new Date().toISOString()]);
  } catch(e) {
    console.error('pushNotification_ failed: ' + e.message);
  }
}

/**
 * Returns all unread notifications for the current user.
 */
function API_getNotifications() {
  try {
    const email = Session.getActiveUser().getEmail();
    const sheet = getTable('Notifications');
    const data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return JSON.stringify({ success: true, data: [] });

    const headers    = data[0].map(h => String(h).trim().toLowerCase());
    const idIdx      = headers.indexOf('id');
    const userIdx    = headers.indexOf('user');
    const msgIdx     = headers.indexOf('message');
    const readIdx    = headers.indexOf('isread');
    const createdIdx = headers.indexOf('createdat');

    const notifs = [];
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][userIdx]).trim().toLowerCase() === email.trim().toLowerCase()) {
        notifs.push({
          id:        data[i][idIdx],
          message:   data[i][msgIdx],
          isRead:    data[i][readIdx] === true || data[i][readIdx] === 'TRUE',
          createdAt: data[i][createdIdx]
        });
      }
      if (notifs.length >= 50) break;
    }
    return JSON.stringify({ success: true, data: notifs });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

/**
 * Marks a single notification as read by its ID.
 */
function API_markNotificationRead(notifId) {
  return withLock(() => {
    const sheet = getTable('Notifications');
    const data  = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const idIdx   = headers.indexOf('id');
    const readIdx = headers.indexOf('isread');

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === notifId) {
        sheet.getRange(i + 1, readIdx + 1).setValue(true);
        SpreadsheetApp.flush();
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, error: 'Notification not found.' });
  });
}

/**
 * Marks ALL notifications for the current user as read.
 */
function API_markAllNotificationsRead() {
  return withLock(() => {
    const email = Session.getActiveUser().getEmail();
    const sheet = getTable('Notifications');
    const data  = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const userIdx = headers.indexOf('user');
    const readIdx = headers.indexOf('isread');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][userIdx]).trim().toLowerCase() === email.trim().toLowerCase() &&
          data[i][readIdx] !== true && data[i][readIdx] !== 'TRUE') {
        sheet.getRange(i + 1, readIdx + 1).setValue(true);
      }
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true });
  });
}

// ---------------------------------------------------------------------------
// AUDIT LOG
// ---------------------------------------------------------------------------

function logActivity(action, oldData, newData) {
  const logSheet = getTable('ActivityLog');
  logSheet.appendRow([
    new Date().toISOString(),
    Session.getActiveUser().getEmail(),
    action,
    JSON.stringify(oldData || {}),
    JSON.stringify(newData || {})
  ]);
}

// ---------------------------------------------------------------------------
// TASK CRUD
// ---------------------------------------------------------------------------

function API_createTask(payloadJson) {
  return withLock(() => {
    const payload = parsePayload(payloadJson);
    const sheet   = getTable('Tasks');
    const headers = sheet.getDataRange().getValues()[0];

    const taskId = generateUUID();
    const now    = new Date().toISOString();
    const userEmail = Session.getActiveUser().getEmail();

    const newTask = {
      id:          taskId,
      title:       payload.title       || 'Untitled Task',
      description: payload.description || '',
      taskType:    payload.taskType    || 'General',
      subType:     payload.subType     || '',
      status:      payload.status      || 'New',
      priority:    payload.priority    || 'Medium',
      assignedTo:  payload.assignedTo  || '',
      createdBy:   userEmail,
      createdAt:   now,
      updatedAt:   now,
      deadline:    payload.deadline    || '',
      isCompleted: false,
      isDeleted:   false,
      metadata:    payload.metadata    || ''
    };

    const rowToAppend = headers.map(header => {
      const hLower = String(header).trim().toLowerCase();
      if (hLower === 'createdat' || hLower === 'updatedat') return "'" + now;
      if (hLower === 'metadata') return newTask.metadata;
      const matchingKey = Object.keys(newTask).find(k => k.toLowerCase() === hLower);
      return matchingKey !== undefined ? newTask[matchingKey] : '';
    });

    sheet.appendRow(rowToAppend);
    SpreadsheetApp.flush();
    logActivity('CREATE_TASK', null, newTask);

    // System comment
    const commentsSheet = getTable('Comments');
    const assignText = payload.assignedTo ? ' and assigned to **' + payload.assignedTo.split('@')[0] + '**' : '';
    commentsSheet.appendRow([generateUUID(), taskId, 'System|' + userEmail, 'Task created' + assignText + '.', "'" + now, "'" + now, false]);
    SpreadsheetApp.flush();

    // Notify assignee
    if (payload.assignedTo && payload.assignedTo !== userEmail) {
      pushNotification_(payload.assignedTo, 'You have been assigned a new task: "' + newTask.title + '".');
    }

    sendPriorityEmail(newTask);
    return JSON.stringify({ success: true, data: newTask });
  });
}

function API_getTasks(page, pageSize, showAll, teamFilter) {
  page     = page     || 1;
  pageSize = pageSize || 500;
  showAll  = showAll  || false;

  const sheet = getTable('Tasks');
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return JSON.stringify({ page, pageSize, total: 0, rows: [], isAdmin: false });

  const headers     = data[0];
  const currentUser = Session.getActiveUser().getEmail();
  const userTier    = getUserTier_(currentUser);
  const isAdmin     = userTier < 3;

  // For team-based filtering, resolve the current user's team
  let currentUserTeam = null;
  if (isAdmin && teamFilter === 'my_team') {
    const usersData = getTable('Users').getDataRange().getValues();
    const uHeaders  = usersData[0].map(h => String(h).trim().toLowerCase());
    const uEmailIdx = uHeaders.indexOf(COLUMN_MAP.USER_EMAIL.toLowerCase());
    const uTeamIdx  = uHeaders.indexOf(COLUMN_MAP.USER_TEAM.toLowerCase());
    for (let i = 1; i < usersData.length; i++) {
      if (String(usersData[i][uEmailIdx]).trim().toLowerCase() === currentUser.trim().toLowerCase()) {
        currentUserTeam = usersData[i][uTeamIdx];
        break;
      }
    }

    if (currentUserTeam) {
      const teamEmails = new Set();
      for (let i = 1; i < usersData.length; i++) {
        if (String(usersData[i][uTeamIdx]).trim() === String(currentUserTeam).trim()) {
          teamEmails.add(String(usersData[i][uEmailIdx]).trim().toLowerCase());
        }
      }
      currentUserTeam = teamEmails;
    }
  }

  let tasks = data.slice(1).map(row => {
    let task = {};
    headers.forEach((h, idx) => { task[String(h).trim()] = row[idx]; });
    return task;
  }).filter(t => t.isDeleted !== true && t.isDeleted !== 'TRUE');

  // Access filter
  if (!isAdmin || !showAll) {
    tasks = tasks.filter(t => t.assignedTo === currentUser || t.createdBy === currentUser);
  } else if (teamFilter === 'my_team' && currentUserTeam instanceof Set) {
    tasks = tasks.filter(t =>
      currentUserTeam.has(String(t.assignedTo || '').trim().toLowerCase()) ||
      currentUserTeam.has(String(t.createdBy  || '').trim().toLowerCase())
    );
  }

  tasks.reverse();
  const start = (page - 1) * pageSize;
  return JSON.stringify({ page, pageSize, total: tasks.length, rows: tasks.slice(start, start + pageSize), isAdmin });
}

function API_getTaskById(taskId) {
  const sheet = getTable('Tasks');
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.map(h => String(h).trim().toLowerCase()).indexOf('id');

  let task = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === taskId) {
      task = {};
      headers.forEach((h, idx) => { task[String(h).trim()] = data[i][idx]; });
      break;
    }
  }
  if (!task) throw new Error('Task not found.');

  const commentsData = getTable('Comments').getDataRange().getValues();
  const cHeaders     = commentsData[0].map(h => String(h).trim().toLowerCase());
  const cTaskIdIdx   = cHeaders.indexOf('taskid');
  const cDeletedIdx  = cHeaders.indexOf('isdeleted');
  let comments = [];
  if (commentsData.length > 1 && cTaskIdIdx !== -1) {
    comments = commentsData.slice(1)
      .filter(row => row[cTaskIdIdx] === taskId && row[cDeletedIdx] !== true && row[cDeletedIdx] !== 'TRUE')
      .map(row => {
        let c = {};
        commentsData[0].forEach((h, idx) => { c[String(h).trim()] = row[idx]; });
        return c;
      });
  }
  return JSON.stringify({ success: true, data: { task, comments } });
}

function API_updateTaskStatus(payloadJson) {
  return withLock(() => {
    const payload = parsePayload(payloadJson);
    const sheet   = getTable('Tasks');
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const user    = Session.getActiveUser().getEmail();

    const idIndex        = headers.indexOf('id');
    const statusIndex    = headers.indexOf('status');
    const updatedAtIndex = headers.indexOf('updatedat');
    if (idIndex === -1 || statusIndex === -1) throw new Error('Database missing required columns.');

    let targetRowIndex = -1;
    let oldRecord = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === payload.id) {
        targetRowIndex = i + 1;
        data[0].forEach((h, idx) => { oldRecord[String(h).trim()] = data[i][idx]; });
        break;
      }
    }
    if (targetRowIndex === -1) throw new Error('Task not found.');

    const now = new Date().toISOString();
    sheet.getRange(targetRowIndex, statusIndex + 1).setValue(payload.status);
    sheet.getRange(targetRowIndex, updatedAtIndex + 1).setValue("'" + now);
    SpreadsheetApp.flush();

    const newRecord = Object.assign({}, oldRecord, { status: payload.status, updatedAt: now });
    logActivity('UPDATE_STATUS', oldRecord, newRecord);

    const commentsSheet = getTable('Comments');
    const systemMessage = 'Changed status from **' + oldRecord.status + '** to **' + payload.status + '**.';
    commentsSheet.appendRow([generateUUID(), payload.id, 'System|' + user, systemMessage, "'" + now, "'" + now, false]);

    if (oldRecord.assignedTo && oldRecord.assignedTo !== user) {
      pushNotification_(oldRecord.assignedTo, 'Task "' + (oldRecord.title || payload.id) + '" status changed to ' + payload.status + '.');
    }

    return JSON.stringify({ success: true, data: newRecord });
  });
}

function API_updateTask(payloadJson) {
  return withLock(() => {
    const payload = parsePayload(payloadJson);
    const sheet   = getTable('Tasks');
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const user    = Session.getActiveUser().getEmail();

    const idIndex = headers.map(h => String(h).trim().toLowerCase()).indexOf('id');
    if (idIndex === -1) throw new Error('Database missing required columns.');

    let targetRowIndex = -1;
    let oldRecord = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === payload.id) {
        targetRowIndex = i + 1;
        headers.forEach((h, idx) => { oldRecord[String(h).trim()] = data[i][idx]; });
        break;
      }
    }
    if (targetRowIndex === -1) throw new Error('Task not found.');

    const now = new Date().toISOString();
    payload.updatedAt = now;

    let changes = [];
    if (oldRecord.assignedTo !== payload.assignedTo) changes.push('Assigned to ' + (payload.assignedTo || 'Unassigned'));
    if (oldRecord.priority   !== payload.priority)   changes.push('Priority to ' + payload.priority);
    if (oldRecord.status     !== payload.status)     changes.push('Status to ' + payload.status);
    if (oldRecord.deadline   !== payload.deadline)   changes.push('Deadline updated');

    const mergedRecord = Object.assign({}, oldRecord, payload);
    const rowToUpdate  = headers.map(h => {
      const k = String(h).trim();
      if (k.toLowerCase() === 'updatedat') return "'" + mergedRecord[k];
      return mergedRecord[k] !== undefined ? mergedRecord[k] : '';
    });

    sheet.getRange(targetRowIndex, 1, 1, headers.length).setValues([rowToUpdate]);
    SpreadsheetApp.flush();

    if (oldRecord.status !== payload.status) logActivity('UPDATE_STATUS', oldRecord, mergedRecord);

    if (changes.length > 0) {
      const commentsSheet = getTable('Comments');
      commentsSheet.appendRow([generateUUID(), payload.id, 'System|' + user, 'Task Updated: ' + changes.join(', ') + '.', "'" + now, "'" + now, false]);
    }

    if (payload.assignedTo && payload.assignedTo !== oldRecord.assignedTo && payload.assignedTo !== user) {
      pushNotification_(payload.assignedTo, 'You have been assigned task: "' + (mergedRecord.title || payload.id) + '".');
    }

    return JSON.stringify({ success: true, data: mergedRecord });
  });
}

/**
 * Soft-deletes a task by setting isDeleted = true.
 * Moves a copy to DeletedTasks sheet for recovery audit.
 */
function API_deleteTask(taskId) {
  return withLock(() => {
    const taskSheet = getTable('Tasks');
    const taskData  = taskSheet.getDataRange().getValues();
    const tHeaders  = taskData[0].map(h => String(h).trim().toLowerCase());
    const tIdIdx       = tHeaders.indexOf('id');
    const tTitleIdx    = tHeaders.indexOf('title');
    const tDeletedIdx  = tHeaders.indexOf('isdeleted');
    const tUpdatedIdx  = tHeaders.indexOf('updatedat');
    const userEmail = Session.getActiveUser().getEmail();
    const now       = new Date().toISOString();

    let taskRowIndex = -1;
    let taskTitle = 'Untitled Task';

    for (let i = 1; i < taskData.length; i++) {
      if (taskData[i][tIdIdx] === taskId) {
        taskRowIndex = i + 1;
        taskTitle    = taskData[i][tTitleIdx] || 'Untitled Task';
        break;
      }
    }
    if (taskRowIndex === -1) throw new Error('Task not found.');

    taskSheet.getRange(taskRowIndex, tDeletedIdx + 1).setValue(true);
    taskSheet.getRange(taskRowIndex, tUpdatedIdx + 1).setValue("'" + now);
    SpreadsheetApp.flush();

    try {
      const deletedSheet = getTable('DeletedTasks');
      deletedSheet.appendRow(taskData[taskRowIndex - 1]);
    } catch(e) { console.warn('DeletedTasks sheet missing: ' + e.message); }

    const commentsSheet = getTable('Comments');
    commentsSheet.appendRow([generateUUID(), taskId, 'System|' + userEmail, '🗑️ Task archived: **' + taskTitle + '**', "'" + now, "'" + now, false]);

    logActivity('ARCHIVED_TASK', { id: taskId, title: taskTitle }, null);
    return JSON.stringify({ success: true });
  });
}

function API_bulkUpdate(payloadJson) {
  return withLock(() => {
    const payload = parsePayload(payloadJson);
    const sheet   = getTable('Tasks');
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const user    = Session.getActiveUser().getEmail();
    const now     = new Date().toISOString();

    const idIdx      = headers.indexOf('id');
    const targetIdx  = headers.indexOf(payload.field.toLowerCase());
    const updatedIdx = headers.indexOf('updatedat');
    if (idIdx === -1 || targetIdx === -1) throw new Error('Invalid database columns for bulk update.');

    let updatedCount = 0;
    const commentsSheet = getTable('Comments');

    for (let i = 1; i < data.length; i++) {
      const taskId = data[i][idIdx];
      if (payload.taskIds.includes(taskId)) {
        const oldValue = data[i][targetIdx];
        sheet.getRange(i + 1, targetIdx + 1).setValue(payload.value);
        sheet.getRange(i + 1, updatedIdx  + 1).setValue("'" + now);
        const actionType    = payload.field === 'status' ? 'UPDATE_STATUS' : 'UPDATE_ASSIGNMENT';
        logActivity(actionType, { id: taskId, [payload.field]: oldValue }, { id: taskId, [payload.field]: payload.value });
        const systemMessage = payload.field === 'status'
          ? 'Bulk changed status from **' + (oldValue||'None') + '** to **' + payload.value + '**.'
          : 'Bulk reassigned from **' + (oldValue||'Unassigned') + '** to **' + payload.value + '**.';
        commentsSheet.appendRow([generateUUID(), taskId, 'System|' + user, systemMessage, "'" + now, "'" + now, false]);
        updatedCount++;
      }
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, updatedCount });
  });
}

// ---------------------------------------------------------------------------
// COMMENTS
// ---------------------------------------------------------------------------

function API_addComment(payloadJson) {
  return withLock(() => {
    const payload = parsePayload(payloadJson);
    const sheet   = getTable('Comments');
    const now     = new Date().toISOString();
    const user    = Session.getActiveUser().getEmail();

    const newComment = {
      id:        generateUUID(),
      taskId:    payload.taskId,
      user:      user,
      message:   payload.message,
      createdAt: now,
      updatedAt: now,
      isDeleted: false
    };

    sheet.appendRow([newComment.id, newComment.taskId, newComment.user, newComment.message, "'" + now, "'" + now, false]);
    SpreadsheetApp.flush();
    logActivity('ADD_COMMENT', null, newComment);
    return JSON.stringify({ success: true, data: newComment });
  });
}

/**
 * Edits the message of an existing comment.
 * Only the original author or a Tier 0/1 admin may edit.
 */
function API_editComment(payloadJson) {
  return withLock(() => {
    const payload   = parsePayload(payloadJson);
    const sheet     = getTable('Comments');
    const data      = sheet.getDataRange().getValues();
    const headers   = data[0].map(h => String(h).trim().toLowerCase());
    const user      = Session.getActiveUser().getEmail();
    const userTier  = getUserTier_(user);

    const idIdx      = headers.indexOf('id');
    const userIdx    = headers.indexOf('user');
    const msgIdx     = headers.indexOf('message');
    const updatedIdx = headers.indexOf('updatedat');
    const deletedIdx = headers.indexOf('isdeleted');

    let targetRow  = -1;
    let oldMessage = '';

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === payload.id) {
        if (data[i][deletedIdx] === true || data[i][deletedIdx] === 'TRUE') {
          throw new Error('Cannot edit a deleted comment.');
        }
        if (data[i][userIdx] !== user && userTier > 1) {
          throw new Error('Security Block: You can only edit your own comments.');
        }
        targetRow  = i + 1;
        oldMessage = data[i][msgIdx];
        break;
      }
    }
    if (targetRow === -1) throw new Error('Comment not found.');

    const now = new Date().toISOString();
    sheet.getRange(targetRow, msgIdx     + 1).setValue(payload.message);
    sheet.getRange(targetRow, updatedIdx + 1).setValue("'" + now);
    SpreadsheetApp.flush();

    logActivity('EDIT_COMMENT', { id: payload.id, message: oldMessage }, { id: payload.id, message: payload.message });
    return JSON.stringify({ success: true });
  });
}

/**
 * Soft-deletes a comment by setting isDeleted = true.
 * Only the original author or a Tier 0/1 admin may delete.
 */
function API_deleteComment(payloadJson) {
  return withLock(() => {
    const payload   = parsePayload(payloadJson);
    const sheet     = getTable('Comments');
    const data      = sheet.getDataRange().getValues();
    const headers   = data[0].map(h => String(h).trim().toLowerCase());
    const user      = Session.getActiveUser().getEmail();
    const userTier  = getUserTier_(user);

    const idIdx      = headers.indexOf('id');
    const userIdx    = headers.indexOf('user');
    const deletedIdx = headers.indexOf('isdeleted');
    const updatedIdx = headers.indexOf('updatedat');

    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === payload.id) {
        if (data[i][userIdx] !== user && userTier > 1) {
          throw new Error('Security Block: You can only delete your own comments.');
        }
        targetRow = i + 1;
        break;
      }
    }
    if (targetRow === -1) throw new Error('Comment not found.');

    const now = new Date().toISOString();
    sheet.getRange(targetRow, deletedIdx + 1).setValue(true);
    sheet.getRange(targetRow, updatedIdx + 1).setValue("'" + now);
    SpreadsheetApp.flush();

    logActivity('DELETE_COMMENT', { id: payload.id }, null);
    return JSON.stringify({ success: true });
  });
}

// ---------------------------------------------------------------------------
// TIME TRACKING
// ---------------------------------------------------------------------------

function API_toggleTime(taskId) {
  return withLock(() => {
    const sheet   = getTable('TimeLogs');
    const data    = sheet.getDataRange().getValues();
    const user    = Session.getActiveUser().getEmail();
    const now     = new Date();
    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const taskIdIdx  = headers.indexOf('taskid');
    const userIdx    = headers.indexOf('user');
    const timeOutIdx = headers.indexOf('timeout');
    const timeInIdx  = headers.indexOf('timein');
    const durIdx     = headers.indexOf('duration');

    let openLogIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][taskIdIdx] === taskId && data[i][userIdx] === user && !data[i][timeOutIdx]) {
        openLogIndex = i + 1;
        break;
      }
    }

    if (openLogIndex !== -1) {
      const timeInStr    = sheet.getRange(openLogIndex, timeInIdx + 1).getValue();
      const diffMs       = now - new Date(timeInStr);
      const durationMins = Math.round(diffMs / 60000);
      sheet.getRange(openLogIndex, timeOutIdx + 1).setValue(now.toISOString());
      sheet.getRange(openLogIndex, durIdx     + 1).setValue(durationMins);
      SpreadsheetApp.flush();
      return JSON.stringify({ success: true, action: 'clocked_out', duration: durationMins });
    } else {
      sheet.appendRow([generateUUID(), taskId, user, now.toISOString(), '', '']);
      SpreadsheetApp.flush();
      return JSON.stringify({ success: true, action: 'clocked_in' });
    }
  });
}

// ---------------------------------------------------------------------------
// SHIFT REPORT
// ---------------------------------------------------------------------------

function API_getShiftReport() {
  try {
    const sessionSheet = getTable('SessionLogs');
    const sessionData  = sessionSheet.getDataRange().getValues();

    const usersSheet = getTable('Users');
    const usersData  = usersSheet.getDataRange().getValues();
    const uHeaders   = usersData[0].map(h => String(h).trim().toLowerCase());
    const uEmailIdx  = uHeaders.indexOf(COLUMN_MAP.USER_EMAIL.toLowerCase());
    const uNameIdx   = uHeaders.indexOf(COLUMN_MAP.USER_NAME.toLowerCase());
    const uRoleIdx   = uHeaders.indexOf(COLUMN_MAP.USER_ROLE.toLowerCase());

    const userMap = {};
    if (uEmailIdx > -1) {
      for (let i = 1; i < usersData.length; i++) {
        userMap[usersData[i][uEmailIdx]] = {
          name: usersData[i][uNameIdx] || 'Unknown',
          role: usersData[i][uRoleIdx] || 'User'
        };
      }
    }

    const sHeaders    = sessionData[0].map(h => String(h).trim().toLowerCase());
    const sEmailIdx   = sHeaders.indexOf(COLUMN_MAP.SESSION_EMAIL.toLowerCase());
    const sTimeInIdx  = sHeaders.indexOf(COLUMN_MAP.SESSION_TIME_IN.toLowerCase());
    const sTimeOutIdx = sHeaders.indexOf(COLUMN_MAP.SESSION_TIME_OUT.toLowerCase());
    const sRemarkIdx  = sHeaders.indexOf(COLUMN_MAP.SESSION_REMARK.toLowerCase());
    const sTeamIdx    = sHeaders.indexOf(COLUMN_MAP.SESSION_TEAM.toLowerCase());
    const sStartIdx   = sHeaders.indexOf(COLUMN_MAP.SESSION_START.toLowerCase());
    const sEndIdx     = sHeaders.indexOf(COLUMN_MAP.SESSION_END.toLowerCase());

    const logs = [];
    const oneDayAgo = new Date(new Date().getTime() - (24 * 60 * 60 * 1000));

    for (let i = sessionData.length - 1; i >= 1; i--) {
      const timeInRaw = sessionData[i][sTimeInIdx];
      if (!timeInRaw) continue;
      const timeIn = new Date(timeInRaw);
      if (timeIn >= oneDayAgo) {
        const email   = sessionData[i][sEmailIdx];
        const userObj = userMap[email] || { name: email, role: 'Unknown' };
        logs.push({
          name:       userObj.name,
          role:       userObj.role,
          timeIn:     timeInRaw,
          timeOut:    sessionData[i][sTimeOutIdx] || null,
          isActive:   !sessionData[i][sTimeOutIdx],
          remark:     sessionData[i][sRemarkIdx]  || '',
          team:       sessionData[i][sTeamIdx]    || '--',
          shiftStart: sessionData[i][sStartIdx]   || '',
          shiftEnd:   sessionData[i][sEndIdx]      || ''
        });
      }
    }
    return JSON.stringify({ success: true, data: logs });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ---------------------------------------------------------------------------
// ANALYTICS
// ---------------------------------------------------------------------------

function API_getAnalytics() {
  try {
    const tasksData = getTable('Tasks').getDataRange().getValues();
    const tHeaders  = tasksData[0];
    const rows      = tasksData.slice(1);

    const usersData = getTable('Users').getDataRange().getValues();
    const uHeaders  = usersData[0].map(h => String(h).trim().toLowerCase());
    const uEmailIdx = uHeaders.indexOf(COLUMN_MAP.USER_EMAIL.toLowerCase());
    const uTeamIdx  = uHeaders.indexOf(COLUMN_MAP.USER_TEAM.toLowerCase());
    const uShiftIdx = uHeaders.indexOf(COLUMN_MAP.USER_SHIFT_NAME.toLowerCase());

    let userMap = {};
    if (uEmailIdx > -1) {
      for (let i = 1; i < usersData.length; i++) {
        userMap[usersData[i][uEmailIdx]] = {
          team:  uTeamIdx  > -1 ? (usersData[i][uTeamIdx]  || 'No Team')  : 'No Team',
          shift: uShiftIdx > -1 ? (usersData[i][uShiftIdx] || 'No Shift') : 'No Shift'
        };
      }
    }

    const props = PropertiesService.getScriptProperties();
    const settingsStr = props.getProperty('DYNAMIC_APP_SETTINGS');
    let timers = { high: 2, medium: 24, low: 48 };
    if (settingsStr) {
      try { const conf = JSON.parse(settingsStr); if (conf.timers) timers = conf.timers; } catch(e) {}
    }

    let metrics = {
      total: 0, completed: 0,
      byStatus: {}, byType: {},
      teamUsage: {}, teamOverdue: {}, teamOverdueDetails: {},
      avgHandleTimeHours: 0
    };

    let totalCompletionTimeMs = 0;
    let completionCount = 0;
    const nowMs = new Date().getTime();

    const hLower      = tHeaders.map(h => String(h).trim().toLowerCase());
    const idIdx       = hLower.indexOf('id');
    const statusIdx   = hLower.indexOf('status');
    const assigneeIdx = hLower.indexOf('assignedto');
    const typeIdx     = hLower.indexOf('tasktype');
    const createdIdx  = hLower.indexOf('createdat');
    const updatedIdx  = hLower.indexOf('updatedat');
    const deletedIdx  = hLower.indexOf('isdeleted');
    const deadlineIdx = hLower.indexOf('deadline');
    const titleIdx    = hLower.indexOf('title');
    const priorityIdx = hLower.indexOf('priority');

    rows.forEach(row => {
      if (row[deletedIdx] === true || row[deletedIdx] === 'TRUE') return;
      metrics.total++;

      const status   = row[statusIdx]   || 'Unknown';
      const type     = row[typeIdx]     || 'General';
      const assignee = row[assigneeIdx];
      const deadline = row[deadlineIdx];
      const title    = row[titleIdx] ? String(row[titleIdx]) : '';
      const priority = (row[priorityIdx] || '').toLowerCase();
      const createdAt = row[createdIdx];
      const updatedAt = row[updatedIdx];
      const isComp    = (status === 'Completed' || status === 'Done');

      metrics.byStatus[status] = (metrics.byStatus[status] || 0) + 1;
      metrics.byType[type]     = (metrics.byType[type]     || 0) + 1;

      let team  = 'Unassigned';
      let shift = 'Unassigned';
      if (assignee && userMap[assignee]) {
        team  = userMap[assignee].team;
        shift = userMap[assignee].shift;
      }
      if (!metrics.teamUsage[team]) metrics.teamUsage[team] = {};
      metrics.teamUsage[team][shift] = (metrics.teamUsage[team][shift] || 0) + 1;

      if (isComp) {
        metrics.completed++;
        const start = new Date(createdAt).getTime();
        const end   = new Date(updatedAt).getTime();
        if (start && end && end > start) { totalCompletionTimeMs += (end - start); completionCount++; }
      }

      let isOverdue = false;
      const slaHours = timers[priority] || 0;
      if (deadline) {
        const dlMs = new Date(deadline).getTime();
        if (isComp) { if (updatedAt && new Date(updatedAt).getTime() > dlMs) isOverdue = true; }
        else        { if (nowMs > dlMs) isOverdue = true; }
      }
      if (!isOverdue && slaHours > 0 && createdAt) {
        const expireMs = new Date(createdAt).getTime() + (slaHours * 3600000);
        if (isComp) { if (updatedAt && new Date(updatedAt).getTime() > expireMs) isOverdue = true; }
        else if (status === 'In-Progress') { if (nowMs > expireMs) isOverdue = true; }
      }
      if (!isOverdue && /^\d{4}_\d{2}_\d{2}/.test(title)) {
        const taskDate = new Date(title.substring(0, 10).replace(/_/g, '-'));
        taskDate.setHours(0,0,0,0);
        const today = new Date(nowMs); today.setHours(0,0,0,0);
        if (taskDate < today) {
          if (!isComp) { isOverdue = true; }
          else if (updatedAt) { const cd = new Date(updatedAt); cd.setHours(0,0,0,0); if (cd > taskDate) isOverdue = true; }
        }
      }
      if (isOverdue) {
        metrics.teamOverdue[team] = (metrics.teamOverdue[team] || 0) + 1;
        if (!metrics.teamOverdueDetails[team]) metrics.teamOverdueDetails[team] = [];
        metrics.teamOverdueDetails[team].push({ id: row[idIdx], title, priority: row[priorityIdx]||'Medium', status, assignedTo: assignee||'Unassigned', isCompleted: isComp });
      }
    });

    if (completionCount > 0) {
      metrics.avgHandleTimeHours = (totalCompletionTimeMs / completionCount / 3600000).toFixed(1);
    }
    return JSON.stringify({ success: true, data: { metrics } });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ---------------------------------------------------------------------------
// CONFIG & SETTINGS
// ---------------------------------------------------------------------------

function API_getConfig() {
  try {
    const props = PropertiesService.getScriptProperties();
    const dynamicSettingsStr = props.getProperty('DYNAMIC_APP_SETTINGS');

    let configData = {};
    if (dynamicSettingsStr && dynamicSettingsStr.trim() !== '') {
      try { configData = JSON.parse(dynamicSettingsStr); } catch(e) { console.error('Failed to parse DYNAMIC_APP_SETTINGS:', e); }
    }

    if (!configData || !configData.taskTypes || configData.taskTypes.length === 0) {
      configData = {
        appName: 'Task Manager',
        theme:   { accent: '#01696f', mode: 'light' },
        timers:  { high: 4, medium: 24, low: 48, none: 0 },
        roleTiers: { '0': ['Admin','Dev'], '1': ['Manager'], '2': ['Lead','QA'], '3': ['User'] },
        taskTypes: [{ label: 'General', color: '#E5E7EB', subTypes: ['Inquiry','Other'], statuses: ['Open','In-Progress','Completed'], metadata: [] }]
      };
    }

    const usersSheet = getTable('Users');
    const uData    = usersSheet.getDataRange().getValues();
    const uHeaders = uData[0].map(h => String(h).trim().toLowerCase());
    const emailIdx = uHeaders.indexOf(COLUMN_MAP.USER_EMAIL.toLowerCase());
    const nameIdx  = uHeaders.indexOf(COLUMN_MAP.USER_NAME.toLowerCase());
    const roleIdx  = uHeaders.indexOf(COLUMN_MAP.USER_ROLE.toLowerCase());
    const teamIdx  = uHeaders.indexOf(COLUMN_MAP.USER_TEAM.toLowerCase());

    let users = [];
    if (emailIdx > -1 && nameIdx > -1) {
      for (let i = 1; i < uData.length; i++) {
        if (uData[i][emailIdx]) {
          users.push({
            email: uData[i][emailIdx],
            name:  uData[i][nameIdx]  || uData[i][emailIdx],
            role:  roleIdx > -1 ? (uData[i][roleIdx] || 'User') : 'User',
            team:  teamIdx > -1 ? (uData[i][teamIdx] || '')     : ''
          });
        }
      }
    }
    configData.users = users;
    return JSON.stringify({ success: true, data: configData });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

function API_getAppSettings() {
  try {
    const props = PropertiesService.getScriptProperties();
    const settingsStr = props.getProperty('DYNAMIC_APP_SETTINGS');
    if (settingsStr && settingsStr.trim() !== '') {
      return JSON.stringify({ success: true, data: JSON.parse(settingsStr) });
    }
    const defaultSettings = {
      appName: 'Task Manager',
      theme:   { accent: '#01696f', mode: 'light' },
      timers:  { high: 4, medium: 24, low: 48, none: 0 },
      roleTiers: { '0': ['Admin','Dev'], '1': ['Manager'], '2': ['Lead','QA'], '3': ['User'] },
      taskTypes: [{ id: generateUUID(), label: 'General', color: '#E5E7EB', subTypes: ['Inquiry','Follow-up','Other'], statuses: ['Open','In-Progress','Completed'], metadata: [] }],
      roles: ['Admin','Manager','Lead','QA','User','Dev'],
      frequencies: ['Ad-hoc','Daily','Weekly','Monthly','Quarterly','Annually','Fortnightly','As Needed'],
      shifts: [
        { name: 'Morning', start: '08:00', end: '17:00' },
        { name: 'Night',   start: '22:00', end: '07:00' },
        { name: 'N/a',     start: '',      end: ''       }
      ]
    };
    return JSON.stringify({ success: true, data: defaultSettings });
  } catch (e) {
    return JSON.stringify({ success: false, error: 'Failed to load App Settings: ' + e.message });
  }
}

function API_saveAppSettings(payloadJson) {
  return withLock(() => {
    const currentUser = Session.getActiveUser().getEmail();
    const userTier    = getUserTier_(currentUser);
    if (userTier >= 2) throw new Error('Security Block: Only Tier 0 and Tier 1 administrators can modify system settings.');

    const payload = parsePayload(payloadJson);
    PropertiesService.getScriptProperties().setProperty('DYNAMIC_APP_SETTINGS', JSON.stringify(payload));
    logActivity('UPDATE_SYSTEM_SETTINGS', null, payload);
    return JSON.stringify({ success: true, message: 'Settings saved successfully.' });
  });
}

// ---------------------------------------------------------------------------
// USERS ADMIN
// ---------------------------------------------------------------------------

function API_getUsersAdmin() {
  try {
    const sheet = getTable('Users');
    const data  = sheet.getDataRange().getDisplayValues();
    const headers = data[0];
    const users = data.slice(1).map(row => {
      let obj = {}; headers.forEach((h, i) => obj[h] = row[i]); return obj;
    });
    return JSON.stringify({ success: true, data: users });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

function API_saveUserAdmin(payloadJson) {
  return withLock(() => {
    const currentUser = Session.getActiveUser().getEmail();
    if (getUserTier_(currentUser) >= 2) throw new Error('Security Block: Unauthorized to modify user records.');

    const payload = parsePayload(payloadJson);
    const sheet   = getTable('Users');
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const emailIdx = headers.indexOf(COLUMN_MAP.USER_EMAIL.toLowerCase());
    let rowIndex = -1;
    if (payload.originalEmail) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][emailIdx]).trim().toLowerCase() === payload.originalEmail.trim().toLowerCase()) {
          rowIndex = i + 1; break;
        }
      }
    }
    const rowData = data[0].map(h => {
      const key   = String(h).trim().toLowerCase();
      const match = Object.keys(payload).find(k => k.toLowerCase() === key);
      return match !== undefined ? payload[match] : '';
    });
    if (rowIndex > -1) { sheet.getRange(rowIndex, 1, 1, data[0].length).setValues([rowData]); }
    else               { sheet.appendRow(rowData); }
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true });
  });
}

function API_deleteUserAdmin(email) {
  return withLock(() => {
    if (getUserTier_(Session.getActiveUser().getEmail()) >= 2) throw new Error('Security Block: Unauthorized to delete user records.');
    const sheet   = getTable('Users');
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const emailIdx = headers.indexOf(COLUMN_MAP.USER_EMAIL.toLowerCase());
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][emailIdx]).trim().toLowerCase() === email.trim().toLowerCase()) {
        sheet.deleteRow(i + 1); SpreadsheetApp.flush();
        return JSON.stringify({ success: true });
      }
    }
    throw new Error('User not found.');
  });
}

function API_bulkUpdateUsersAdmin(payloadJson) {
  return withLock(() => {
    if (getUserTier_(Session.getActiveUser().getEmail()) >= 2) throw new Error('Security Block: Unauthorized to modify user records.');
    const payload = parsePayload(payloadJson);
    const sheet   = getTable('Users');
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const emailIdx = headers.indexOf(COLUMN_MAP.USER_EMAIL.toLowerCase());
    if (emailIdx === -1) throw new Error('email column not found.');
    let updatedCount = 0;
    for (let i = 1; i < data.length; i++) {
      if (payload.emails.includes(data[i][emailIdx])) {
        for (const [key, value] of Object.entries(payload.updates)) {
          if (value !== undefined && value !== null && value !== '') {
            const colIdx2 = headers.indexOf(key.toLowerCase());
            if (colIdx2 !== -1) sheet.getRange(i + 1, colIdx2 + 1).setValue(value);
          }
        }
        updatedCount++;
      }
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, count: updatedCount });
  });
}

/**
 * DEV ONLY: Switches the calling user's role.
 * Gated to Tier 0 (Admin/Dev) only.
 */
function API_devSwitchRole(newRole) {
  return withLock(() => {
    const email    = Session.getActiveUser().getEmail();
    const userTier = getUserTier_(email);
    if (userTier > 0) throw new Error('Security Block: API_devSwitchRole is restricted to Tier 0 (Admin/Dev) users only.');

    const sheet   = getTable('Users');
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const emailIdx = headers.indexOf(COLUMN_MAP.USER_EMAIL.toLowerCase());
    const roleIdx  = headers.indexOf(COLUMN_MAP.USER_ROLE.toLowerCase());
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][emailIdx]).trim().toLowerCase() === email.trim().toLowerCase()) {
        sheet.getRange(i + 1, roleIdx + 1).setValue(newRole || 'Dev');
        SpreadsheetApp.flush();
        return JSON.stringify({ success: true, newRole: newRole });
      }
    }
    throw new Error('Dev user not found in the Users table.');
  });
}

// ---------------------------------------------------------------------------
// SHIFT HANDOVER
// ---------------------------------------------------------------------------

function API_submitShiftHandover(payloadStr) {
  try {
    const data  = JSON.parse(payloadStr);
    const sheet = getTable('ShiftLogs');
    const mentionedStr   = (data.mentionedUsers && data.mentionedUsers.length > 0) ? data.mentionedUsers.join(', ') : 'None';
    const mentionedArray = data.mentionedUsers || [];

    sheet.appendRow([
      new Date().toISOString(), data.user, data.completedCount, data.pendingCount,
      data.summary, data.handoff || 'None', data.blockers || 'None', mentionedStr
    ]);
    SpreadsheetApp.flush();

    if (mentionedArray.length > 0) {
      const senderName = data.user.split('@')[0].toUpperCase();
      const subject = '📢 Shift Handover Mention from ' + senderName;
      const body = 'Hi,\n\n' + senderName + ' has mentioned you in their End-of-Shift Handover.\n\n' +
                   '📝 SUMMARY:\n' + data.summary + '\n\n' +
                   '⏳ PENDING ITEMS FOR YOU:\n' + (data.handoff || 'None') + '\n\n' +
                   '🛑 BLOCKERS/ESCALATIONS:\n' + (data.blockers || 'None') + '\n\n' +
                   'Log into the Task Manager to view the full report.';
      mentionedArray.forEach(email => {
        try { MailApp.sendEmail(email.trim(), subject, body); }
        catch(e) { console.error('Handover email failed: ' + e.message); }
      });
    }
    return JSON.stringify({ success: true });
  } catch (error) {
    return JSON.stringify({ success: false, error: error.message });
  }
}

function API_getShiftLogs() {
  try {
    const sheet = getTable('ShiftLogs');
    const data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return JSON.stringify({ success: true, data: [] });
    const headers = data[0];
    const logs = data.slice(1).map(row => {
      let obj = {}; headers.forEach((h, i) => obj[h] = row[i]); return obj;
    }).reverse();
    return JSON.stringify({ success: true, data: logs });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ---------------------------------------------------------------------------
// REAL-TIME CHECKS
// ---------------------------------------------------------------------------

function API_checkNewAssignments(lastCheckIsoString, userEmail) {
  try {
    const serverNow = new Date();
    if (!lastCheckIsoString) return JSON.stringify({ success: true, updates: [], serverTime: serverNow.toISOString() });

    const taskSheet = getTable('Tasks');
    const tasks     = taskSheet.getDataRange().getValues();
    const tHeaders  = tasks[0].map(h => String(h).trim().toLowerCase());
    const idIdx        = tHeaders.indexOf('id');
    const titleIdx     = tHeaders.indexOf('title');
    const assignedIdx  = tHeaders.indexOf('assignedto');
    const createdByIdx = tHeaders.indexOf('createdby');
    const updatedIdx   = tHeaders.indexOf('updatedat');
    const checkTimeMs  = new Date(lastCheckIsoString).getTime() - 2000;
    let updates = [];

    const activitySheet = getTable('Comments');
    const activities    = activitySheet.getDataRange().getValues();
    const aHeaders      = activities[0].map(h => String(h).trim().toLowerCase());
    const aTaskIdx      = aHeaders.indexOf('taskid');
    const aMsgIdx       = aHeaders.indexOf('message');

    for (let i = 1; i < tasks.length; i++) {
      const row = tasks[i];
      const isConnected = row[assignedIdx] === userEmail || row[createdByIdx] === userEmail;
      if (isConnected && row[updatedIdx]) {
        if (new Date(row[updatedIdx]).getTime() > checkTimeMs) {
          let lastAction = 'Task was updated';
          for (let j = activities.length - 1; j >= 1; j--) {
            if (activities[j][aTaskIdx] === row[idIdx]) { lastAction = activities[j][aMsgIdx]; break; }
          }
          updates.push({ taskId: row[idIdx], title: row[titleIdx], message: lastAction });
        }
      }
    }
    return JSON.stringify({ success: true, updates: updates, serverTime: serverNow.toISOString() });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ---------------------------------------------------------------------------
// RECENT ACTIVITY
// ---------------------------------------------------------------------------

function API_getRecentActivity() {
  try {
    const commentsData = getTable('Comments').getDataRange().getValues();
    const tasksData    = getTable('Tasks').getDataRange().getValues();
    const tHeaders     = tasksData[0].map(h => String(h).trim().toLowerCase());
    const tIdIdx       = tHeaders.indexOf('id');
    const tTitleIdx    = tHeaders.indexOf('title');

    let taskMap = {};
    for (let i = 1; i < tasksData.length; i++) taskMap[tasksData[i][tIdIdx]] = tasksData[i][tTitleIdx];

    let activities = [];
    if (commentsData.length > 1) {
      const cHeaders = commentsData[0].map(h => String(h).trim().toLowerCase());
      const cTaskIdx = cHeaders.indexOf('taskid');
      const cUserIdx = cHeaders.indexOf('user');
      const cMsgIdx  = cHeaders.indexOf('message');
      const cTimeIdx = cHeaders.indexOf('createdat');
      const cDelIdx  = cHeaders.indexOf('isdeleted');
      const start = Math.max(1, commentsData.length - 50);
      for (let i = commentsData.length - 1; i >= start; i--) {
        if (commentsData[i][cDelIdx] === true || commentsData[i][cDelIdx] === 'TRUE') continue;
        const taskId = commentsData[i][cTaskIdx];
        let userStr  = commentsData[i][cUserIdx] || 'Unknown';
        let isSystem = false;
        if (userStr.startsWith('System|')) { isSystem = true; userStr = userStr.split('|')[1] || 'System'; }
        activities.push({
          taskId:    taskId,
          taskTitle: taskMap[taskId] || 'Archived/Deleted Task',
          user:      userStr.split('@')[0],
          message:   commentsData[i][cMsgIdx],
          timestamp: commentsData[i][cTimeIdx],
          isSystem:  isSystem
        });
      }
    }
    return JSON.stringify({ success: true, data: activities });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

// ---------------------------------------------------------------------------
// ACCESS REQUESTS
// ---------------------------------------------------------------------------

function API_requestAccess(reason) {
  return withLock(() => {
    const email = Session.getActiveUser().getEmail();
    if (!email) throw new Error('Could not detect Google account.');

    const sheet = getTable('AccessRequests');
    const data  = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const eIdx = headers.indexOf('email');
    const sIdx = headers.indexOf('status');

    const pending = data.slice(1).some(row => String(row[eIdx]).trim().toLowerCase() === email.trim().toLowerCase() && row[sIdx] === 'Pending');
    if (pending) return JSON.stringify({ success: false, message: 'You already have a pending access request.' });

    sheet.appendRow([new Date().toISOString(), email, reason || 'No reason provided', 'Pending']);
    SpreadsheetApp.flush();

    const usersSheet = getTable('Users');
    const uData    = usersSheet.getDataRange().getValues();
    const uHeaders = uData[0].map(h => String(h).trim().toLowerCase());
    const uEmailIdx = uHeaders.indexOf(COLUMN_MAP.USER_EMAIL.toLowerCase());
    const uRoleIdx  = uHeaders.indexOf(COLUMN_MAP.USER_ROLE.toLowerCase());
    const adminEmails = [];
    if (uEmailIdx > -1 && uRoleIdx > -1) {
      const props = PropertiesService.getScriptProperties();
      const settingsStr = props.getProperty('DYNAMIC_APP_SETTINGS');
      let tier0Roles = ['Admin','Dev'];
      if (settingsStr) { try { const c = JSON.parse(settingsStr); if (c.roleTiers && c.roleTiers['0']) tier0Roles = c.roleTiers['0']; } catch(e) {} }
      for (let i = 1; i < uData.length; i++) {
        const role = String(uData[i][uRoleIdx] || '').toLowerCase().trim();
        if (tier0Roles.some(r => r.toLowerCase() === role) && uData[i][uEmailIdx]) adminEmails.push(uData[i][uEmailIdx]);
      }
    }
    if (adminEmails.length > 0) {
      MailApp.sendEmail(adminEmails.join(','), '🔐 TMA Access Request: ' + email,
        'A new user has requested access.\n\nEmail: ' + email + '\nReason: ' + (reason || 'N/A') +
        '\n\nAdd them to the Users sheet to grant access.');
    }
    return JSON.stringify({ success: true });
  });
}

// ---------------------------------------------------------------------------
// JANITOR
// ---------------------------------------------------------------------------

function system_janitorCleanup() {
  const DAYS_TO_KEEP = 7;
  const threshold = new Date();
  threshold.setDate(threshold.getDate() - DAYS_TO_KEEP);

  const taskSheet = getTable('Tasks');
  const taskData  = taskSheet.getDataRange().getValues();
  const headers   = taskData[0].map(h => String(h).trim().toLowerCase());
  const statusIdx  = headers.indexOf('status');
  const updatedIdx = headers.indexOf('updatedat');
  const idIdx      = headers.indexOf('id');

  const toArchive = [];
  for (let i = 1; i < taskData.length; i++) {
    const status    = taskData[i][statusIdx];
    const updatedAt = new Date(taskData[i][updatedIdx]);
    if ((status === 'Completed' || status === 'Done') && updatedAt < threshold) {
      toArchive.push(taskData[i][idIdx]);
    }
  }
  toArchive.forEach(taskId => {
    try { API_deleteTask(taskId); } catch(e) { console.error('Janitor failed on ' + taskId + ': ' + e.message); }
  });
  if (toArchive.length > 0) logActivity('SYSTEM_JANITOR_RUN', null, { archivedCount: toArchive.length });
}
