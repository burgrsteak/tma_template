/**
 * Retrieves the database ID from script properties.
 * Run setupDatabaseId() once to set this up.
 */
function getDbId() {
  const dbId = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!dbId) throw new Error('DB_ID is not set. Run API_initDatabase() first.');
  return dbId;
}

/**
 * Returns the connected Spreadsheet instance.
 */
function getDb() {
  return SpreadsheetApp.openById(getDbId());
}

/**
 * Helper to get a specific table (sheet) by name.
 */
function getTable(tableName) {
  const sheet = getDb().getSheetByName(tableName);
  if (!sheet) throw new Error('Table "' + tableName + '" not found. Run API_initDatabase() to create it.');
  return sheet;
}

/**
 * TABLE_SCHEMAS — defines every required sheet and its headers.
 * Add new tables here; API_initDatabase() will auto-create them.
 */
const TABLE_SCHEMAS = {
  Tasks: [
    'id','title','description','taskType','subType','status','priority',
    'assignedTo','createdBy','createdAt','updatedAt','deadline',
    'isCompleted','isDeleted','metadata'
  ],
  Comments: ['id','taskId','user','message','createdAt','updatedAt','isDeleted'],
  TimeLogs:    ['id','taskId','user','timeIn','timeOut','duration'],
  ActivityLog: ['timestamp','user','action','oldData','newData'],
  Notifications: ['id','user','message','isRead','createdAt'],
  SessionLogs: [
    'id','email','timeIn','timeOut','remark',
    'team','shiftStart','shiftEnd'
  ],
  Users: [
    'email','name','role','team','scheduledShift','shiftStart','shiftEnd'
  ],
  DeletedTasks:       ['id','title','description','taskType','subType','status','priority','assignedTo','createdBy','createdAt','updatedAt','deadline','isCompleted','isDeleted','metadata'],
  DeletedComments:    ['id','taskId','user','message','createdAt','updatedAt','isDeleted'],
  DeletedActivityLog: ['timestamp','user','action','oldData','newData'],
  ShiftLogs: ['timestamp','user','completedCount','pendingCount','summary','handoff','blockers','mentionedUsers'],
  AccessRequests: ['timestamp','email','reason','status']
};

/**
 * API_initDatabase
 * ─────────────────
 * One-time setup function. Call this once after deploying a fresh copy.
 *
 * What it does:
 *  1. Reads DB_ID from Script Properties (must be set manually first).
 *  2. For every table in TABLE_SCHEMAS, creates the sheet if it doesn't exist
 *     and writes the header row.
 *  3. Saves a default DYNAMIC_APP_SETTINGS blob so the app has a working config.
 *  4. Logs the result to the Apps Script console.
 *
 * How to run:
 *  - Open Apps Script editor → select API_initDatabase → Run.
 *  - Or call it from a setup trigger.
 */
function API_initDatabase() {
  const props = PropertiesService.getScriptProperties();
  const dbId  = props.getProperty('DB_ID');

  if (!dbId) {
    throw new Error(
      'DB_ID is not set. ' +
      'Go to Project Settings → Script Properties and add DB_ID = <your Google Sheet ID>.'
    );
  }

  const ss = SpreadsheetApp.openById(dbId);
  const results = [];

  for (const [tableName, headers] of Object.entries(TABLE_SCHEMAS)) {
    let sheet = ss.getSheetByName(tableName);
    if (!sheet) {
      sheet = ss.insertSheet(tableName);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      results.push('CREATED: ' + tableName);
    } else {
      // Sheet already exists — ensure header row is correct
      const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const missing = headers.filter(h => !existingHeaders.includes(h));
      if (missing.length > 0) {
        // Append missing columns at the end
        missing.forEach(col => {
          const nextCol = sheet.getLastColumn() + 1;
          sheet.getRange(1, nextCol).setValue(col).setFontWeight('bold');
        });
        results.push('PATCHED: ' + tableName + ' (added: ' + missing.join(', ') + ')');
      } else {
        results.push('OK: ' + tableName);
      }
    }
  }

  // Write default app settings if none exist
  const existing = props.getProperty('DYNAMIC_APP_SETTINGS');
  if (!existing) {
    const defaultSettings = {
      appName: 'Task Manager',
      theme: { accent: '#01696f', mode: 'light' },
      timers: { high: 4, medium: 24, low: 48, none: 0 },
      roleTiers: {
        '0': ['Admin', 'Dev'],
        '1': ['Manager'],
        '2': ['Lead', 'QA'],
        '3': ['User']
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
      roles: ['Admin', 'Manager', 'Lead', 'QA', 'User', 'Dev'],
      frequencies: ['Ad-hoc','Daily','Weekly','Monthly','Quarterly','Annually','As Needed'],
      shifts: [
        { name: 'Morning', start: '08:00', end: '17:00' },
        { name: 'Night',   start: '22:00', end: '07:00' },
        { name: 'N/a',     start: '',      end: '' }
      ]
    };
    props.setProperty('DYNAMIC_APP_SETTINGS', JSON.stringify(defaultSettings));
    results.push('CREATED: DYNAMIC_APP_SETTINGS (default config written)');
  } else {
    results.push('OK: DYNAMIC_APP_SETTINGS (already exists, not overwritten)');
  }

  console.log('=== API_initDatabase Results ===\n' + results.join('\n'));
  return JSON.stringify({ success: true, results: results });
}

/**
 * Fetches and structures all dynamic configuration data
 * from the Config and Users sheets.
 */
function API_getConfig_Legacy() {
  // Kept for backward-compat. The main API_getConfig is in TaskLogic.js.
  try {
    const configSheet = getTable('Config');
    const configData  = configSheet.getDataRange().getValues();
    const configRows  = configData.slice(1);

    const config = {
      taskTypes: [{ label: 'All', color: '#E5E7EB' }],
      subTypes: {},
      workflows: {},
      priorities: [],
      users: []
    };

    configRows.forEach(row => {
      const category = String(row[0]).trim();
      const label    = String(row[1]).trim();
      const parent   = String(row[2]).trim();
      const meta1    = row[3];
      if (!category || !label) return;
      switch (category) {
        case 'TaskType':       config.taskTypes.push({ label, color: meta1 || '#F3F4F6' }); break;
        case 'SubType':        if (!config.subTypes[parent]) config.subTypes[parent] = []; config.subTypes[parent].push(label); break;
        case 'StatusWorkflow': if (!config.workflows[parent]) config.workflows[parent] = []; config.workflows[parent].push(label); break;
        case 'Priority':       config.priorities.push({ label, hours: meta1 }); break;
        default: console.warn('Unknown config category: ' + category);
      }
    });

    config.taskTypes.push({ label: 'General', color: '#F3F4F6' });
    config.workflows['General'] = ['New','Open','In-Progress','Completed'];

    try {
      const usersSheet = getTable('Users');
      const usersData  = usersSheet.getDataRange().getValues();
      if (usersData.length >= 2) {
        const headers  = usersData[0].map(h => String(h).trim().toLowerCase());
        const emailIdx = headers.indexOf(COLUMN_MAP.USER_EMAIL.toLowerCase());
        const nameIdx  = headers.indexOf(COLUMN_MAP.USER_NAME.toLowerCase());
        if (emailIdx === -1 || nameIdx === -1) {
          console.error('Users sheet missing required headers.');
        } else {
          for (let i = 1; i < usersData.length; i++) {
            const email = String(usersData[i][emailIdx]).trim();
            const name  = String(usersData[i][nameIdx]).trim();
            if (email) config.users.push({ name, email });
          }
        }
      }
    } catch (e) {
      console.error('Could not load Users: ' + e.message);
    }

    return JSON.stringify({ success: true, data: config });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}
