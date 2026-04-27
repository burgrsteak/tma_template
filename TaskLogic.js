/**
 * Router to safely serve HTML pages.
 */
function API_getPage(pageName) {
  const allowedPages = ['Dashboard', 'ShiftReport', 'Analytics', 'Settings']; 
  if (!allowedPages.includes(pageName)) {
    throw new Error("Unauthorized page request or page does not exist.");
  }
  return include('Pages/' + pageName);
}

/**
 * Checks if the user has an active 12-hour session on page load/refresh.
 */
function API_checkCurrentSession() {
  const email = Session.getActiveUser().getEmail();
  if (!email) return JSON.stringify({ active: false });

  const props = PropertiesService.getUserProperties();
  const sessionStart = props.getProperty('sessionStart');
  
  if (sessionStart) {
    const now = new Date().getTime();
    const start = parseInt(sessionStart, 10);
    const hoursElapsed = (now - start) / (1000 * 60 * 60);
    
    if (hoursElapsed < 12) {
      const usersSheet = getTable('Users');
      const data = usersSheet.getDataRange().getValues();
      const headers = data[0];
      const emailIdx = headers.indexOf('Snap Emails');
      const nameIdx = headers.indexOf('Names');
      const roleIdx = headers.indexOf('Role');
      
      let userRecord = null;
      for (let i = 1; i < data.length; i++) {
        if (data[i][emailIdx] === email) {
          userRecord = { name: data[i][nameIdx], email: email, role: data[i][roleIdx] };
          break;
        }
      }
      return JSON.stringify({ active: true, user: userRecord });
    } else {
      props.deleteProperty('sessionStart');
      return JSON.stringify({ active: false, expired: true });
    }
  }
  return JSON.stringify({ active: false });
}

/**
 * Logs the user in, starts the 12-hour clock, and records 'timeIn'.
 * Takes a permanent snapshot of the user's Team and Shift at the moment of login.
 */
function API_loginUser() {
  return withLock(() => {
    const email = Session.getActiveUser().getEmail().trim().toLowerCase();
    
    // 1. Get the User's Current Profile
    const usersSheet = getTable('Users');
    const data = usersSheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailIdx = headers.indexOf('Snap Emails');
    const nameIdx = headers.indexOf('Names');
    const roleIdx = headers.indexOf('Role');
    const teamIdx = headers.indexOf('Team');
    const shiftStartIdx = headers.indexOf('Start Shift Time');
    const shiftEndIdx = headers.indexOf('End Shift Time');
    
    let userRecord = null;
    let currentTeam = '--';
    let currentStart = '';
    let currentEnd = '';

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][emailIdx]).trim().toLowerCase() === email) {
        userRecord = { name: data[i][nameIdx], email: email, role: data[i][roleIdx] };
        currentTeam = teamIdx > -1 ? data[i][teamIdx] : '--';
        currentStart = shiftStartIdx > -1 ? data[i][shiftStartIdx] : '';
        currentEnd = shiftEndIdx > -1 ? data[i][shiftEndIdx] : '';
        break;
      }
    }
    
    if (userRecord) {
      const props = PropertiesService.getUserProperties();
      const now = new Date();
      
      props.setProperty('sessionStart', now.getTime().toString());
      
      const logSheet = getTable('SessionLogs');
      const logData = logSheet.getDataRange().getValues();
      
      // AUTO-CLOSE GHOST SESSIONS
      if (logData.length > 1) {
        for (let i = 1; i < logData.length; i++) {
          let rowEmail = String(logData[i][1]).trim().toLowerCase();
          let rowTimeOut = String(logData[i][3]).trim();
          
          if (rowEmail === email && rowTimeOut === "") { 
            logSheet.getRange(i + 1, 4).setValue(now.toISOString()); // Set Time Out
            logSheet.getRange(i + 1, 5).setValue("System Auto-Close: Duplicate Session Detected"); // Add Remark
          }
        }
      }
      
      const logId = generateUUID();
      props.setProperty('currentLogId', logId);
      
      // 2. Append new session WITH the permanent snapshot of their schedule
      logSheet.appendRow([logId, email, now.toISOString(), '', '', currentTeam, currentStart, currentEnd]);
      SpreadsheetApp.flush(); // Force immediate save
      
      return JSON.stringify({ success: true, authorized: true, user: userRecord });
    } else {
      return JSON.stringify({ success: true, authorized: false, email: email });
    }
  });
}

/**
 * Logs the user out, clears the 12-hour clock, and records 'timeOut'.
 */
function API_logoutUser() {
  return withLock(() => {
    const props = PropertiesService.getUserProperties();
    const logId = props.getProperty('currentLogId');
    
    if (logId) {
      const logSheet = getTable('SessionLogs');
      const data = logSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === logId) { 
          logSheet.getRange(i + 1, 4).setValue(new Date().toISOString()); 
          break;
        }
      }
      SpreadsheetApp.flush(); // Force immediate save
    }
    
    props.deleteProperty('sessionStart');
    props.deleteProperty('currentLogId');
    
    return JSON.stringify({ success: true });
  });
}

/**
 * Helper: Sends an email notification for High Priority assignments.
 */
function sendPriorityEmail(task) {
  if (task.priority !== 'High' || !task.assignedTo) return;
  
  const subject = `🚨 HIGH PRIORITY: ${task.title}`;
  const body = `Hi,\n\nYou have been assigned a High Priority task: "${task.title}".\n\nLink: ${ScriptApp.getService().getUrl()}\n\nPriority: ${task.priority}\nDue: ${task.deadline || 'N/A'}`;
  
  try {
    MailApp.sendEmail(task.assignedTo, subject, body);
  } catch (e) {
    console.error("Email failed to send: " + e.message);
  }
}

/**
 * Creates a new task atomically and logs it to the Activity Feed.
 */
function API_createTask(payloadJson) {
  return withLock(() => {
    const payload = parsePayload(payloadJson);
    const sheet = getTable('Tasks');
    const headers = sheet.getDataRange().getValues()[0]; 
    
    const taskId = generateUUID();
    const now = new Date().toISOString(); 
    const userEmail = Session.getActiveUser().getEmail();
    
    const newTask = {
      id: taskId,
      title: payload.title || 'Untitled Task',
      description: payload.description || '',
      taskType: payload.taskType || 'General',
      subType: payload.subType || '', 
      status: payload.status || 'New', 
      priority: payload.priority || 'Medium',
      assignedTo: payload.assignedTo || '',
      createdBy: userEmail,
      createdAt: now,
      updatedAt: now,
      deadline: payload.deadline || '',
      isCompleted: false,
      isDeleted: false,
      metadata: payload.metadata || ''
    };
    
    // 🔥 THE FIX: Bulletproof mapping. Checks exact match first, then lowercase match.
    const rowToAppend = headers.map(header => {
      const hStr = header.toString().trim();
      const hLower = hStr.toLowerCase();
      
      if (hLower === 'createdat' || hLower === 'updatedat') return "'" + now;
      if (hLower === 'metadata') return newTask.metadata;
      
      if (newTask[hStr] !== undefined) return newTask[hStr];
      const matchingKey = Object.keys(newTask).find(k => k.toLowerCase() === hLower);
      return matchingKey && newTask[matchingKey] !== undefined ? newTask[matchingKey] : '';
    });
    
    sheet.appendRow(rowToAppend);
    SpreadsheetApp.flush(); 
    
    logActivity('CREATE_TASK', null, newTask);

    const commentsSheet = getTable('Comments');
    const assignText = payload.assignedTo ? ` and assigned to **${payload.assignedTo.split('@')[0]}**` : '';
    commentsSheet.appendRow([generateUUID(), taskId, 'System|' + userEmail, `Task created${assignText}.`, "'" + now, "'" + now]);
    SpreadsheetApp.flush();
    
    sendPriorityEmail(newTask);
    
    return JSON.stringify({ success: true, data: newTask });
  });
}

/**
 * Centralized Audit Log Helper.
 */
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

/**
 * Fetches the session logs for the last 24 hours.
 */
function API_getShiftReport() {
  try {
    const sessionSheet = getTable('SessionLogs');
    const sessionData = sessionSheet.getDataRange().getValues();
    
    const usersSheet = getTable('Users');
    const usersData = usersSheet.getDataRange().getValues();
    const userMap = {};
    const uHeaders = usersData[0];
    const emailIdx = uHeaders.indexOf('Snap Emails');
    const nameIdx = uHeaders.indexOf('Names');
    const roleIdx = uHeaders.indexOf('Role');
    
    if (emailIdx > -1) {
      for (let i = 1; i < usersData.length; i++) {
        userMap[usersData[i][emailIdx]] = {
          name: usersData[i][nameIdx] || 'Unknown',
          role: usersData[i][roleIdx] || 'Agent'
        };
      }
    }
    
    const sHeaders = sessionData[0];
    const sEmailIdx = sHeaders.indexOf('email') > -1 ? sHeaders.indexOf('email') : 1;
    const sTimeInIdx = sHeaders.indexOf('timeIn') > -1 ? sHeaders.indexOf('timeIn') : 2;
    const sTimeOutIdx = sHeaders.indexOf('timeOut') > -1 ? sHeaders.indexOf('timeOut') : 3;
    const sRemarkIdx = 4; // Col E
    const sTeamIdx = 5;   // Col F 
    const sStartIdx = 6;  // Col G 
    const sEndIdx = 7;    // Col H 
    
    let logs = [];
    const oneDayAgo = new Date(new Date().getTime() - (24 * 60 * 60 * 1000));
    
    for (let i = sessionData.length - 1; i >= 1; i--) {
      const timeInRaw = sessionData[i][sTimeInIdx];
      if (!timeInRaw) continue;
      
      const timeIn = new Date(timeInRaw);
      if (timeIn >= oneDayAgo) {
        const email = sessionData[i][sEmailIdx];
        const userObj = userMap[email] || { name: email, role: 'Unknown' };
        
        logs.push({
          name: userObj.name,
          role: userObj.role,
          timeIn: timeInRaw,
          timeOut: sessionData[i][sTimeOutIdx] || null,
          isActive: !sessionData[i][sTimeOutIdx],
          remark: sessionData[i][sRemarkIdx] || '',
          team: sessionData[i][sTeamIdx] || '--',
          shiftStart: sessionData[i][sStartIdx] || '',
          shiftEnd: sessionData[i][sEndIdx] || ''
        });
      }
    }
    
    return JSON.stringify({ success: true, data: logs });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

/**
 * 🔒 SECURITY HELPER: Calculates the user's Tier based on dynamic Global Settings
 */
function getUserTier_(email) {
  const usersData = getTable('Users').getDataRange().getValues();
  let userRole = '';
  
  const emailCol = usersData[0].indexOf('Snap Emails');
  const roleCol = usersData[0].indexOf('Role');
  
  if (emailCol !== -1 && roleCol !== -1) {
    for (let i = 1; i < usersData.length; i++) {
      if (usersData[i][emailCol] === email) { 
        userRole = usersData[i][roleCol]; 
        break; 
      }
    }
  }
  
  const props = PropertiesService.getScriptProperties();
  const settingsStr = props.getProperty('DYNAMIC_APP_SETTINGS');
  
  let currentTiers = { 
    "0": ["SDL", "Dev"], 
    "1": ["TL", "OM", "Ops Lead"], 
    "2": ["SME", "QA"], 
    "3": ["Agent"] 
  };
  
  if (settingsStr) {
    const conf = JSON.parse(settingsStr);
    if (conf.roleTiers) currentTiers = conf.roleTiers;
  }
  
  let userTier = 3; 
  for (const [tierLvl, rolesArray] of Object.entries(currentTiers)) {
    if (rolesArray.some(r => r.toLowerCase() === (userRole || '').toLowerCase())) {
      userTier = parseInt(tierLvl);
      break;
    }
  }
  
  return userTier;
}

function API_getTasks(page = 1, pageSize = 500, showAll = false) {
  const sheet = getTable('Tasks');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return JSON.stringify({ page, pageSize, total: 0, rows: [], isAdmin: false });

  const headers = data[0];
  const rows = data.slice(1);
  const currentUser = Session.getActiveUser().getEmail();

  // Tier Check
  const userTier = getUserTier_(currentUser);
  const isAdmin = userTier < 3; 

  let tasks = rows.map(row => {
    let task = {};
    headers.forEach((h, idx) => {
      let key = h.toString().trim();
      if (key.toLowerCase() === 'metadata') key = 'metadata';
      task[key] = row[idx];
    });
    return task;
  }).filter(t => t.isDeleted !== true && t.isDeleted !== 'TRUE');

  // Logic: If Admin wants "Team View", show all. Otherwise, only show theirs.
  if (!isAdmin || !showAll) {
    tasks = tasks.filter(t => t.assignedTo === currentUser || t.createdBy === currentUser);
  }

  // 🔥 THE FIX: Reverse the array so the newest rows at the bottom of the sheet show up first!
  tasks.reverse();

  const start = (page - 1) * pageSize;
  return JSON.stringify({ page, pageSize, total: tasks.length, rows: tasks.slice(start, start + pageSize), isAdmin });
}

function API_getTaskById(taskId) {
  const sheet = getTable('Tasks');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idIndex = headers.indexOf('id');
  
  let task = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idIndex] === taskId) {
      task = {};
      headers.forEach((h, idx) => {
        let key = h.toString().trim();
        if (key.toLowerCase() === 'metadata') key = 'metadata';
        task[key] = data[i][idx];
      });
      break;
    }
  }
  if (!task) throw new Error('Task not found.');

  const commentsData = getTable('Comments').getDataRange().getValues();
  const cHeaders = commentsData[0];
  const cTaskIdIdx = cHeaders.indexOf('taskId');
  let comments = [];
  if (commentsData.length > 1 && cTaskIdIdx !== -1) {
    comments = commentsData.slice(1)
      .filter(row => row[cTaskIdIdx] === taskId)
      .map(row => {
        let c = {};
        cHeaders.forEach((h, idx) => c[h] = row[idx]);
        return c;
      });
  }

  return JSON.stringify({ success: true, data: { task, comments } }); 
}

function API_addComment(payloadJson) {
  return withLock(() => {
    const payload = parsePayload(payloadJson);
    const sheet = getTable('Comments');
    const now = new Date().toISOString();
    const user = Session.getActiveUser().getEmail();
    
    const newComment = {
      id: generateUUID(),
      taskId: payload.taskId,
      user: user,
      message: payload.message,
      createdAt: now,
      updatedAt: now
    };
    
    sheet.appendRow([newComment.id, newComment.taskId, newComment.user, newComment.message, "'" + now, "'" + now]);
    SpreadsheetApp.flush(); // Force save
    
    logActivity('ADD_COMMENT', null, newComment); 
    
    return JSON.stringify({ success: true, data: newComment });
  });
}

function API_toggleTime(taskId) {
  return withLock(() => {
    const sheet = getTable('TimeLogs');
    const data = sheet.getDataRange().getValues();
    const user = Session.getActiveUser().getEmail();
    const now = new Date();
    
    const tHeaders = data[0];
    const taskIdIdx = tHeaders.indexOf('taskId');
    const userIdx = tHeaders.indexOf('user');
    const timeOutIdx = tHeaders.indexOf('timeOut');
    const timeInIdx = tHeaders.indexOf('timeIn');
    const durationIdx = tHeaders.indexOf('duration');
    
    let openLogIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][taskIdIdx] === taskId && data[i][userIdx] === user && !data[i][timeOutIdx]) {
        openLogIndex = i + 1; 
        break;
      }
    }
    
    if (openLogIndex !== -1) {
      const timeInStr = sheet.getRange(openLogIndex, timeInIdx + 1).getValue();
      const timeInDate = new Date(timeInStr);
      const diffMs = now - timeInDate;
      const durationMins = Math.round(diffMs / 60000);
      
      sheet.getRange(openLogIndex, timeOutIdx + 1).setValue(now.toISOString());
      sheet.getRange(openLogIndex, durationIdx + 1).setValue(durationMins);
      SpreadsheetApp.flush();
      
      return JSON.stringify({ success: true, action: 'clocked_out', duration: durationMins });
    } else {
      const newLogId = generateUUID();
      sheet.appendRow([newLogId, taskId, user, now.toISOString(), '', '']);
      SpreadsheetApp.flush();
      return JSON.stringify({ success: true, action: 'clocked_in' });
    }
  });
}

function API_updateTaskStatus(payloadJson) {
  return withLock(() => {
    const payload = parsePayload(payloadJson);
    const sheet = getTable('Tasks');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const user = Session.getActiveUser().getEmail(); 
    
    const idIndex = headers.indexOf('id');
    const statusIndex = headers.indexOf('status');
    const updatedAtIndex = headers.indexOf('updatedAt');
    
    if (idIndex === -1 || statusIndex === -1) throw new Error('Database missing required columns.');
    
    let targetRowIndex = -1;
    let oldRecord = {};
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === payload.id) {
        targetRowIndex = i + 1;
        headers.forEach((h, idx) => oldRecord[h] = data[i][idx]);
        break;
      }
    }
    
    if (targetRowIndex === -1) throw new Error('Task not found.');
    
    const now = new Date().toISOString(); 
    
    sheet.getRange(targetRowIndex, statusIndex + 1).setValue(payload.status);
    sheet.getRange(targetRowIndex, updatedAtIndex + 1).setValue("'" + now);
    SpreadsheetApp.flush(); 
    
    const newRecord = { ...oldRecord, status: payload.status, updatedAt: now };
    
    logActivity('UPDATE_STATUS', oldRecord, newRecord);
    
    const commentsSheet = getTable('Comments');
    const systemMessage = `Changed status from **${oldRecord.status}** to **${payload.status}**.`;
    commentsSheet.appendRow([generateUUID(), payload.id, 'System|' + user, systemMessage, "'" + now, "'" + now]);
    
    return JSON.stringify({ success: true, data: newRecord });
  });
}

function API_updateTask(payloadJson) {
  return withLock(() => {
    const payload = parsePayload(payloadJson);
    const sheet = getTable('Tasks');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const user = Session.getActiveUser().getEmail();
    
    const idIndex = headers.indexOf('id');
    if (idIndex === -1) throw new Error('Database missing required columns.');
    
    let targetRowIndex = -1;
    let oldRecord = {};
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === payload.id) {
        targetRowIndex = i + 1;
        headers.forEach((h, idx) => oldRecord[h] = data[i][idx]);
        break;
      }
    }
    
    if (targetRowIndex === -1) throw new Error('Task not found.');
    
    const now = new Date().toISOString();
    payload.updatedAt = now;
    
    let changes = [];
    if (oldRecord.assignedTo !== payload.assignedTo) changes.push(`Assigned to ${payload.assignedTo || 'Unassigned'}`);
    if (oldRecord.priority !== payload.priority) changes.push(`Priority to ${payload.priority}`);
    if (oldRecord.status !== payload.status) changes.push(`Status to ${payload.status}`);
    if (oldRecord.deadline !== payload.deadline) changes.push(`Deadline updated`);

    const mergedRecord = { ...oldRecord, ...payload };
    
    const rowToUpdate = headers.map(h => {
      if (h === 'updatedAt') return "'" + mergedRecord[h];
      return mergedRecord[h] !== undefined ? mergedRecord[h] : '';
    });
    
    sheet.getRange(targetRowIndex, 1, 1, headers.length).setValues([rowToUpdate]);
    SpreadsheetApp.flush(); 
    
    if (oldRecord.status !== payload.status) {
      logActivity('UPDATE_STATUS', oldRecord, mergedRecord);
    }
    
    if (changes.length > 0) {
      const commentsSheet = getTable('Comments');
      const systemMessage = `Task Updated: ${changes.join(', ')}.`;
      commentsSheet.appendRow([generateUUID(), payload.id, 'System|' + user, systemMessage, "'" + now, "'" + now]);
    }
    
    return JSON.stringify({ success: true, data: mergedRecord });
  });
}

function API_checkNewAssignments(lastCheckIsoString, userEmail) {
  try {
    const serverNow = new Date();
    let debugLogs = []; 
    
    if (!lastCheckIsoString) {
      return JSON.stringify({ success: true, updates: [], serverTime: serverNow.toISOString(), debug: debugLogs });
    }

    const taskSheet = getTable('Tasks');
    const tasks = taskSheet.getDataRange().getValues();
    const tHeaders = tasks[0];
    
    const idIdx = tHeaders.indexOf('id');
    const titleIdx = tHeaders.indexOf('title');
    const assignedIdx = tHeaders.indexOf('assignedTo');
    const createdByIdx = tHeaders.indexOf('createdBy');
    const updatedIdx = tHeaders.indexOf('updatedAt');

    const checkTimeMs = new Date(lastCheckIsoString).getTime() - 2000;
    let updates = [];

    const activitySheet = getTable('Comments');
    const activities = activitySheet.getDataRange().getValues();

    for (let i = 1; i < tasks.length; i++) {
      const row = tasks[i];
      const isConnected = row[assignedIdx] === userEmail || row[createdByIdx] === userEmail;
      
      if (isConnected && row[updatedIdx]) {
        const rowTimeMs = new Date(row[updatedIdx]).getTime();
        
        if (rowTimeMs > checkTimeMs) {
          let lastAction = "Task was updated";
          for (let j = activities.length - 1; j >= 1; j--) {
            if (activities[j][1] === row[idIdx]) { 
              lastAction = activities[j][3]; 
              break;
            }
          }
          updates.push({ taskId: row[idIdx], title: row[titleIdx], message: lastAction });
        }
      }
    }
    return JSON.stringify({ success: true, updates: updates, serverTime: serverNow.toISOString(), debug: debugLogs });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message, debug: ["CRASHED"] });
  }
}

function API_deleteTask(taskId) {
  return withLock(() => {
    const taskSheet = getTable('Tasks');
    const taskData = taskSheet.getDataRange().getValues();
    const tHeaders = taskData[0];
    const tIdIdx = tHeaders.indexOf('id');
    const tTitleIdx = tHeaders.indexOf('title');

    let taskRowIndex = -1;
    let taskRowData = null;
    let taskTitle = 'Unknown Task'; 

    for (let i = 1; i < taskData.length; i++) {
      if (taskData[i][tIdIdx] === taskId) {
        taskRowIndex = i + 1; 
        taskRowData = taskData[i];
        taskTitle = taskData[i][tTitleIdx] || 'Untitled Task';
        break;
      }
    }

    if (taskRowIndex === -1) throw new Error('Task not found.');

    const deletedTasksSheet = getTable('DeletedTasks');
    deletedTasksSheet.appendRow(taskRowData);
    taskSheet.deleteRow(taskRowIndex);
    SpreadsheetApp.flush(); 

    const commentsSheet = getTable('Comments');
    const commentsData = commentsSheet.getDataRange().getValues();
    
    if (commentsData.length > 1) {
      const cHeaders = commentsData[0];
      const cTaskIdIdx = cHeaders.indexOf('taskId');
      const deletedCommentsSheet = getTable('DeletedComments');

      for (let i = commentsData.length - 1; i >= 1; i--) {
        if (commentsData[i][cTaskIdIdx] === taskId) {
          deletedCommentsSheet.appendRow(commentsData[i]);
          commentsSheet.deleteRow(i + 1);
        }
      }
      SpreadsheetApp.flush(); 
    }

    const activitySheet = getTable('ActivityLog');
    const activityData = activitySheet.getDataRange().getValues();
    
    if (activityData.length > 1) {
      const deletedActivitySheet = getTable('DeletedActivityLog');

      for (let i = activityData.length - 1; i >= 1; i--) {
        const oldDataStr = String(activityData[i][3] || ''); 
        const newDataStr = String(activityData[i][4] || ''); 
        
        if (oldDataStr.includes(taskId) || newDataStr.includes(taskId)) {
          deletedActivitySheet.appendRow(activityData[i]);
          activitySheet.deleteRow(i + 1);
        }
      }
    }

    const userEmail = Session.getActiveUser().getEmail();
    const now = new Date().toISOString();
    
    commentsSheet.appendRow([
      generateUUID(), 
      taskId, 
      'System|' + userEmail, 
      `🗑️ Task archived: **${taskTitle}**`, 
      "'" + now, 
      "'" + now
    ]);

    logActivity('ARCHIVED_TASK', { id: taskId }, null);
    return JSON.stringify({ success: true });
  });
}

/**
 * Crunch the numbers for the Insights Dashboard.
 * UPDATED: Accurately calculates "Completed Late" and packages the specific task details for Drill-Down Analytics.
 */
function API_getAnalytics() {
  try {
    const tasksData = getTable('Tasks').getDataRange().getValues();
    const tHeaders = tasksData[0];
    const rows = tasksData.slice(1);

    // 1. Get Users to map Team & Shift
    const usersData = getTable('Users').getDataRange().getValues();
    const uHeaders = usersData[0];
    const emailIdx = uHeaders.indexOf('Snap Emails');
    const teamIdx = uHeaders.indexOf('Team');
    const shiftIdx = uHeaders.indexOf('Scheduled Shift');

    let userMap = {};
    if (emailIdx > -1) {
      for (let i = 1; i < usersData.length; i++) {
        let email = usersData[i][emailIdx];
        userMap[email] = {
          team: teamIdx > -1 ? (usersData[i][teamIdx] || 'No Team') : 'No Team',
          shift: shiftIdx > -1 ? (usersData[i][shiftIdx] || 'No Shift') : 'No Shift'
        };
      }
    }

    // 2. Fetch Dynamic SLA Timers
    const props = PropertiesService.getScriptProperties();
    const settingsStr = props.getProperty('DYNAMIC_APP_SETTINGS');
    let timers = { high: 2, medium: 24, low: 48 };
    if (settingsStr) {
      const conf = JSON.parse(settingsStr);
      if (conf.timers) timers = conf.timers;
    }

    let metrics = {
      total: 0,
      completed: 0,
      byStatus: {},
      byType: {},
      teamUsage: {}, 
      teamOverdue: {}, 
      teamOverdueDetails: {}, // NEW: Stores actual task metadata for the drill-down modal
      avgHandleTimeHours: 0
    };

    let totalCompletionTimeMs = 0;
    let completionCount = 0;
    const nowMs = new Date().getTime(); 

    const idIdx = tHeaders.indexOf('id');
    const statusIdx = tHeaders.indexOf('status');
    const assigneeIdx = tHeaders.indexOf('assignedTo');
    const typeIdx = tHeaders.indexOf('taskType');
    const createdIdx = tHeaders.indexOf('createdAt');
    const updatedIdx = tHeaders.indexOf('updatedAt');
    const deletedIdx = tHeaders.indexOf('isDeleted');
    const deadlineIdx = tHeaders.indexOf('deadline');
    const titleIdx = tHeaders.indexOf('title');
    const priorityIdx = tHeaders.indexOf('priority');

    // 3. Loop through Tasks and Aggregate in memory
    rows.forEach(row => {
      if (row[deletedIdx] === true || row[deletedIdx] === 'TRUE') return;

      metrics.total++;
      const status = row[statusIdx] || 'Unknown';
      const type = row[typeIdx] || 'General';
      const assignee = row[assigneeIdx];
      const deadline = row[deadlineIdx];
      const title = row[titleIdx] ? String(row[titleIdx]) : '';
      const priority = (row[priorityIdx] || '').toLowerCase();
      const createdAt = row[createdIdx];
      const updatedAt = row[updatedIdx];

      const isComp = (status === 'Completed' || status === 'Done');

      metrics.byStatus[status] = (metrics.byStatus[status] || 0) + 1;
      metrics.byType[type] = (metrics.byType[type] || 0) + 1;

      // Map Assignee to Team and Shift
      let team = 'Unassigned';
      let shift = 'Unassigned';
      if (assignee && userMap[assignee]) {
        team = userMap[assignee].team;
        shift = userMap[assignee].shift;
      }

      if (!metrics.teamUsage[team]) metrics.teamUsage[team] = {};
      metrics.teamUsage[team][shift] = (metrics.teamUsage[team][shift] || 0) + 1;

      if (isComp) {
        metrics.completed++;
        const start = new Date(createdAt).getTime();
        const end = new Date(updatedAt).getTime();
        if (start && end && end > start) {
          totalCompletionTimeMs += (end - start);
          completionCount++;
        }
      }

      // ==========================================
      // OVERDUE & COMPLETED LATE LOGIC
      // ==========================================
      let isOverdue = false;
      const slaHours = timers[priority] || 0;

      // 1. Hard Deadline
      if (deadline) {
        const dlMs = new Date(deadline).getTime();
        if (isComp) {
          if (updatedAt && new Date(updatedAt).getTime() > dlMs) isOverdue = true;
        } else {
          if (nowMs > dlMs) isOverdue = true;
        }
      }

      // 2. SLA Timer
      if (!isOverdue && slaHours > 0 && createdAt) {
         const startTimeMs = new Date(createdAt).getTime();
         const expireTimeMs = startTimeMs + (slaHours * 60 * 60 * 1000);
         
         if (isComp) {
            if (updatedAt && new Date(updatedAt).getTime() > expireTimeMs) isOverdue = true;
         } else if (status === 'In-Progress') {
            if (nowMs > expireTimeMs) isOverdue = true;
         }
      }

      // 3. Daily Routine
      if (!isOverdue && /^\d{4}_\d{2}_\d{2}/.test(title)) {
        const dateStr = title.substring(0, 10).replace(/_/g, '-');
        const taskDate = new Date(dateStr);
        taskDate.setHours(0,0,0,0);
        
        const today = new Date(nowMs);
        today.setHours(0,0,0,0);
        
        if (taskDate < today) {
          if (!isComp) {
            isOverdue = true;
          } else if (updatedAt) {
            const compDate = new Date(updatedAt);
            compDate.setHours(0,0,0,0);
            if (compDate > taskDate) {
              isOverdue = true;
            }
          }
        }
      }

      // Add to Team Overdue Tally AND store Task Details
      if (isOverdue) {
         metrics.teamOverdue[team] = (metrics.teamOverdue[team] || 0) + 1;
         
         if (!metrics.teamOverdueDetails[team]) metrics.teamOverdueDetails[team] = [];
         metrics.teamOverdueDetails[team].push({
           id: row[idIdx],
           title: title,
           priority: row[priorityIdx] || 'Medium',
           status: status,
           assignedTo: assignee || 'Unassigned',
           isCompleted: isComp
         });
      }
    });

    if (completionCount > 0) {
      metrics.avgHandleTimeHours = (totalCompletionTimeMs / completionCount / (1000 * 60 * 60)).toFixed(1);
    }

    return JSON.stringify({ success: true, data: { metrics } });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

function API_bulkUpdate(payloadJson) {
  return withLock(() => {
    const payload = parsePayload(payloadJson);
    const sheet = getTable('Tasks');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const user = Session.getActiveUser().getEmail();
    const now = new Date().toISOString();

    const idIdx = headers.indexOf('id');
    const targetIdx = headers.indexOf(payload.field);
    const updatedIdx = headers.indexOf('updatedAt');

    if (idIdx === -1 || targetIdx === -1) throw new Error('Invalid database columns for bulk update.');

    let updatedCount = 0;
    const commentsSheet = getTable('Comments');

    for (let i = 1; i < data.length; i++) {
      const taskId = data[i][idIdx];
      if (payload.taskIds.includes(taskId)) {
        const rowNum = i + 1; 
        const oldValue = data[i][targetIdx];
        
        sheet.getRange(rowNum, targetIdx + 1).setValue(payload.value);
        sheet.getRange(rowNum, updatedIdx + 1).setValue("'" + now);
        
        const actionType = payload.field === 'status' ? 'UPDATE_STATUS' : 'UPDATE_ASSIGNMENT';
        logActivity(actionType, { id: taskId, [payload.field]: oldValue }, { id: taskId, [payload.field]: payload.value });
        
        const systemMessage = payload.field === 'status' 
          ? `Bulk changed status from **${oldValue || 'None'}** to **${payload.value}**.`
          : `Bulk reassigned from **${oldValue || 'Unassigned'}** to **${payload.value}**.`;
          
        commentsSheet.appendRow([generateUUID(), taskId, 'System|' + user, systemMessage, "'" + now, "'" + now]);
        
        updatedCount++;
      }
    }
    SpreadsheetApp.flush(); 
    return JSON.stringify({ success: true, updatedCount });
  });
}

function system_janitorCleanup() {
  const DAYS_TO_KEEP = 7; 
  const thresholdDate = new Date();
  thresholdDate.setDate(thresholdDate.getDate() - DAYS_TO_KEEP);

  const taskSheet = getTable('Tasks');
  const taskData = taskSheet.getDataRange().getValues();
  const headers = taskData[0];
  
  const statusIdx = headers.indexOf('status');
  const updatedIdx = headers.indexOf('updatedAt');
  const idIdx = headers.indexOf('id');

  let tasksToArchive = [];

  for (let i = 1; i < taskData.length; i++) {
    const status = taskData[i][statusIdx];
    const updatedAt = new Date(taskData[i][updatedIdx]);
    
    if ((status === 'Completed' || status === 'Done') && updatedAt < thresholdDate) {
      tasksToArchive.push(taskData[i][idIdx]);
    }
  }

  tasksToArchive.forEach(taskId => {
    try { API_deleteTask(taskId); } catch (e) { console.error(`Janitor failed: ${e.message}`); }
  });

  if (tasksToArchive.length > 0) {
    logActivity('SYSTEM_JANITOR_RUN', null, { archivedCount: tasksToArchive.length });
  }
}

/**
 * Global App Configuration Loader
 * FIX: Robustly parses the saved JSON and correctly maps user roles based on the saved data.
 */
function API_getConfig() {
  try {
    const props = PropertiesService.getScriptProperties();
    const dynamicSettingsStr = props.getProperty('DYNAMIC_APP_SETTINGS');
    
    let configData = {};
    
    if (dynamicSettingsStr && dynamicSettingsStr.trim() !== "") {
      try {
        configData = JSON.parse(dynamicSettingsStr);
      } catch (parseError) {
        console.error("Failed to parse DYNAMIC_APP_SETTINGS JSON:", parseError);
        // Fallback below will handle it if parsing fails
      }
    } 
    
    // If empty or parsing failed, supply a robust default that matches your structure
    if (!configData || !configData.taskTypes || configData.taskTypes.length === 0) {
      configData = {
        appName: 'Task Manager',
        timers: { high: 4, medium: 24, low: 48, none: 0 },
        roleTiers: {
          "0": ["SDL", "Dev"],
          "1": ["TL", "OM", "Ops Lead"],
          "2": ["SME", "QA"],
          "3": ["Agent", "POC"] // Added POC based on your user list
        },
        taskTypes: [
          { label: 'General', color: '#E5E7EB', subTypes: ['Inquiry', 'Other'], statuses: ['Open', 'In-Progress', 'Completed'], metadata: [] }
        ]
      };
    }
    
    // 1. Get Users from Sheet to ensure we have the absolute latest list
    const usersSheet = getTable('Users');
    const uData = usersSheet.getDataRange().getValues();
    let users = [];
    const emailIdx = uData[0].indexOf('Snap Emails');
    const nameIdx = uData[0].indexOf('Names');
    const roleIdx = uData[0].indexOf('Role'); 
    
    if (emailIdx > -1 && nameIdx > -1) {
      for (let i = 1; i < uData.length; i++) {
        if (uData[i][emailIdx]) {
          users.push({ 
            email: uData[i][emailIdx], 
            name: uData[i][nameIdx],
            role: roleIdx > -1 ? uData[i][roleIdx] : 'Agent' 
          });
        }
      }
    }
    
    // Overwrite the JSON users array with the LIVE sheet users array
    configData.users = users;
    
    return JSON.stringify({ success: true, data: configData });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

/**
 * Fetches the dynamic App Settings for the Settings Page
 * FIX: Ensures the exact structure is returned, or a highly detailed default is provided.
 */
function API_getAppSettings() {
  try {
    const props = PropertiesService.getScriptProperties();
    const settingsStr = props.getProperty('DYNAMIC_APP_SETTINGS');

    Logger.log(settingsStr)
    
    if (settingsStr && settingsStr.trim() !== "") {
      const parsedData = JSON.parse(settingsStr);
      return JSON.stringify({ success: true, data: parsedData });
    } else {
      // If nothing is found in the properties, return a default template
      const defaultSettings = {
        appName: 'HERMES',
        timers: { high: 4, medium: 24, low: 48, none: 0 }, 
        roleTiers: {
          "0": ["SDL", "Dev"],
          "1": ["TL", "OM", "Ops Lead"],
          "2": ["SME", "QA"],
          "3": ["Agent", "POC"]
        },
        taskTypes: [
          { 
            id: generateUUID(), 
            label: 'General', 
            color: '#E5E7EB', 
            subTypes: ['Inquiry', 'Follow-up', 'Other'],
            statuses: ['Open', 'In-Progress', 'Completed'],
            metadata: [] 
          }
        ],
        roles: ["Agent", "TL", "SME", "QA", "SDL", "Dev", "OM", "Ops Lead", "POC"],
        frequencies: ["Ad-hoc", "Daily", "Weekly", "Monthly", "Quarterly", "Annually", "Fortnightly", "As Needed"],
        shifts: [
          { name: "Mid", start: "14:00", end: "23:00" },
          { name: "Night", start: "22:00", end: "07:00" },
          { name: "N/a", start: "", end: "" }
        ]
      };
      return JSON.stringify({ success: true, data: defaultSettings });
    }
  } catch (e) {
    return JSON.stringify({ success: false, error: "Failed to load App Settings: " + e.message });
  }
}

function API_saveAppSettings(payloadJson) {
  return withLock(() => {
    const currentUser = Session.getActiveUser().getEmail();
    const userTier = getUserTier_(currentUser);
    
    if (userTier >= 2) {
      throw new Error("Security Block: Only Tier 0 and Tier 1 administrators can modify system settings.");
    }

    const payload = parsePayload(payloadJson);
    const props = PropertiesService.getScriptProperties();
    props.setProperty('DYNAMIC_APP_SETTINGS', JSON.stringify(payload));
    
    logActivity('UPDATE_SYSTEM_SETTINGS', null, payload);
    
    return JSON.stringify({ success: true, message: 'Settings saved successfully.' });
  });
}

function API_getUsersAdmin() {
  try {
    const sheet = getTable('Users');
    const data = sheet.getDataRange().getDisplayValues(); 
    const headers = data[0];
    const users = data.slice(1).map(row => {
      let obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
    return JSON.stringify({ success: true, data: users });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

function API_saveUserAdmin(payloadJson) {
  return withLock(() => {
    const currentUser = Session.getActiveUser().getEmail();
    const userTier = getUserTier_(currentUser);
    
    if (userTier >= 2) {
      throw new Error("Security Block: Unauthorized to modify user records.");
    }

    const payload = parsePayload(payloadJson);
    const sheet = getTable('Users');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailIdx = headers.indexOf('Snap Emails');
    let rowIndex = -1;
    
    if (payload.originalEmail) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][emailIdx] === payload.originalEmail) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    const rowData = headers.map(h => payload[h] !== undefined ? payload[h] : '');

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true });
  });
}

function API_deleteUserAdmin(email) {
  return withLock(() => {
    const currentUser = Session.getActiveUser().getEmail();
    const userTier = getUserTier_(currentUser);
    
    if (userTier >= 2) {
      throw new Error("Security Block: Unauthorized to delete user records.");
    }

    const sheet = getTable('Users');
    const data = sheet.getDataRange().getValues();
    const emailIdx = data[0].indexOf('Snap Emails');
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][emailIdx] === email) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return JSON.stringify({ success: true });
      }
    }
    throw new Error("User not found.");
  });
}

function API_devSwitchRole(newRole) {
  return withLock(() => {
    const email = Session.getActiveUser().getEmail();
    const sheet = getTable('Users');
    const data = sheet.getDataRange().getValues();
    const emailIdx = data[0].indexOf('Snap Emails');
    const roleIdx = data[0].indexOf('Role');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailIdx] === email) {
        sheet.getRange(i + 1, roleIdx + 1).setValue(newRole || 'Dev');
        SpreadsheetApp.flush();
        return JSON.stringify({ success: true, newRole: newRole });
      }
    }
    throw new Error("Dev user not found in the Users table.");
  });
}

function API_submitShiftHandover(payloadStr) {
  try {
    const data = JSON.parse(payloadStr);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("ShiftLogs");
    
    if (!sheet) {
      sheet = ss.insertSheet("ShiftLogs");
      sheet.appendRow(["Timestamp", "User Email", "Completed Today", "Pending", "Summary", "Handoff Notes", "Blockers", "Mentioned Users"]);
      sheet.getRange("A1:H1").setFontWeight("bold");
    }
    
    let mentionedStr = "None";
    let mentionedArray = [];
    
    if (data.mentionedUsers && Array.isArray(data.mentionedUsers) && data.mentionedUsers.length > 0) {
      mentionedStr = data.mentionedUsers.join(', ');
      mentionedArray = data.mentionedUsers;
    }
    
    sheet.appendRow([
      new Date().toISOString(),       
      data.user,                      
      data.completedCount,            
      data.pendingCount,              
      data.summary,                   
      data.handoff || "None",         
      data.blockers || "None",        
      mentionedStr                    
    ]);
    SpreadsheetApp.flush();
    
    if (mentionedArray.length > 0) {
      const senderName = data.user.split('@')[0].toUpperCase();
      const subject = `📢 Shift Handover Mention from ${senderName}`;
      const body = `Hi,\n\n${senderName} has mentioned you in their End-of-Shift Handover.\n\n` +
                   `📝 SUMMARY:\n${data.summary}\n\n` +
                   `⏳ PENDING ITEMS FOR YOU:\n${data.handoff || 'None'}\n\n` +
                   `🛑 BLOCKERS/ESCALATIONS:\n${data.blockers || 'None'}\n\n` +
                   `Log into the Task Manager to view the full report.`;
      
      mentionedArray.forEach(email => {
        try {
          MailApp.sendEmail(email.trim(), subject, body);
        } catch (e) {
          console.error("Failed to send handover email to " + email + ": " + e.message);
        }
      });
    }
    
    return JSON.stringify({ success: true });
    
  } catch (error) {
    return JSON.stringify({ success: false, error: error.message });
  }
}

function API_getShiftLogs() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("ShiftLogs");
    
    if (!sheet) return JSON.stringify({ success: true, data: [] });

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return JSON.stringify({ success: true, data: [] });

    const headers = data[0];
    const rows = data.slice(1);
    
    let logs = rows.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    }).reverse();

    return JSON.stringify({ success: true, data: logs });
    
  } catch (error) {
    return JSON.stringify({ success: false, error: error.message });
  }
}

function API_bulkUpdateUsersAdmin(payloadJson) {
  return withLock(() => {
    const currentUser = Session.getActiveUser().getEmail();
    const userTier = getUserTier_(currentUser);
    
    if (userTier >= 2) {
      throw new Error("Security Block: Unauthorized to modify user records.");
    }

    const payload = parsePayload(payloadJson); 
    const sheet = getTable('Users');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const emailIdx = headers.indexOf('Snap Emails');
    if (emailIdx === -1) throw new Error("Snap Emails column not found.");

    let updatedCount = 0;

    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][emailIdx];
      
      if (payload.emails.includes(rowEmail)) {
        for (const [key, value] of Object.entries(payload.updates)) {
          if (value !== undefined && value !== null && value !== '') {
            const colIdx = headers.indexOf(key);
            if (colIdx !== -1) {
              sheet.getRange(i + 1, colIdx + 1).setValue(value);
            }
          }
        }
        updatedCount++;
      }
    }
    SpreadsheetApp.flush();
    return JSON.stringify({ success: true, count: updatedCount });
  });
}

function API_getRecentActivity() {
  try {
    const commentsData = getTable('Comments').getDataRange().getValues();
    const tasksData = getTable('Tasks').getDataRange().getValues();
    
    let taskMap = {};
    if (tasksData.length > 1) {
      const tIdIdx = tasksData[0].indexOf('id');
      const tTitleIdx = tasksData[0].indexOf('title');
      for(let i = 1; i < tasksData.length; i++) {
         taskMap[tasksData[i][tIdIdx]] = tasksData[i][tTitleIdx];
      }
    }
    
    let activities = [];
    if (commentsData.length > 1) {
      const cHeaders = commentsData[0];
      const start = Math.max(1, commentsData.length - 50); 
      
      for(let i = commentsData.length - 1; i >= start; i--) {
         const row = commentsData[i];
         const taskId = row[cHeaders.indexOf('taskId')];
         let userStr = row[cHeaders.indexOf('user')] || 'Unknown';
         let isSystem = false;
         
         if(userStr.startsWith('System|')) {
            isSystem = true;
            userStr = userStr.split('|')[1] || 'System';
         }
         
         activities.push({
           taskId: taskId,
           taskTitle: taskMap[taskId] || 'Archived/Deleted Task',
           user: userStr.split('@')[0],
           message: row[cHeaders.indexOf('message')],
           timestamp: row[cHeaders.indexOf('createdAt')],
           isSystem: isSystem
         });
      }
    }
    return JSON.stringify({ success: true, data: activities });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}

/**
 * Handles unauthorized users requesting access to the system.
 * Auto-creates the tracking sheet and emails administrators.
 */
function API_requestAccess(reason) {
  return withLock(() => {
    const email = Session.getActiveUser().getEmail();
    if (!email) throw new Error("Could not detect Google account.");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('AccessRequests');
    
    // Failsafe: Auto-create the tab if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('AccessRequests');
      sheet.appendRow(['Timestamp', 'Email', 'Reason', 'Status']);
      sheet.getRange("A1:D1").setFontWeight("bold");
    }

    // Prevent spam: Check if they already have a pending request
    const data = sheet.getDataRange().getValues();
    const pending = data.some(row => row[1] === email && row[3] === 'Pending');
    if (pending) {
      return JSON.stringify({ success: false, message: "You already have a pending access request. Please wait for an admin to approve it." });
    }

    // 1. Log the Request
    sheet.appendRow([new Date().toISOString(), email, reason || 'No reason provided', 'Pending']);
    SpreadsheetApp.flush();

    // 2. Find Admins to Notify
    const usersSheet = getTable('Users');
    const uData = usersSheet.getDataRange().getValues();
    const adminEmails = [];
    const emailIdx = uData[0].indexOf('Snap Emails');
    const roleIdx = uData[0].indexOf('Role');
    
    if (emailIdx > -1 && roleIdx > -1) {
      for (let i = 1; i < uData.length; i++) {
        const role = String(uData[i][roleIdx]).toLowerCase().trim();
        // Notify top-tier roles
        if (role === 'sdl' || role === 'dev' || role === 'om') {
          if (uData[i][emailIdx]) adminEmails.push(uData[i][emailIdx]);
        }
      }
    }

    // 3. Send Email Alert
    if (adminEmails.length > 0) {
      const subject = `🔐 TMA Access Request: ${email}`;
      const body = `Hello,\n\nA new user has requested access to the Task Manager App.\n\n` +
                   `Email: ${email}\n` +
                   `Reason / Team: ${reason || 'N/A'}\n\n` +
                   `Please add their email and role to the 'Users' tab in the database to grant them access.`;
                   
      MailApp.sendEmail(adminEmails.join(','), subject, body);
    }

    return JSON.stringify({ success: true });
  });
}