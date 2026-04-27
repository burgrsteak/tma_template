/**
 * Retrieves the database ID from script properties.
 * Run setupDatabaseId() once to set this up.
 */
function getDbId() {
  const dbId = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!dbId) throw new Error('DB_ID is not set in Script Properties.');
  return dbId;
}

/**
 * Returns the connected Spreadsheet instance.
 */
function getDb() {
  return SpreadsheetApp.openById(getDbId());
}

/**
 * Helper to get a specific table (sheet).
 */
function getTable(tableName) {
  const sheet = getDb().getSheetByName(tableName);
  if (!sheet) throw new Error(`Table "${tableName}" not found in database.`);
  return sheet;
}

/**
 * Fetches and structures all dynamic configuration data
 * from the Config and Users sheets.
 */
function API_getConfig() {
  try {

    // --- 1. Config Sheet ---
    const configSheet = getTable('Config');
    const configData = configSheet.getDataRange().getValues(); // Bulk read
    const configRows = configData.slice(1); // Skip header row

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

      if (!category || !label) return; // Skip blank rows

      switch (category) {
        case 'TaskType':
          config.taskTypes.push({ label: label, color: meta1 || '#F3F4F6' });
          break;
        case 'SubType':
          if (!config.subTypes[parent]) config.subTypes[parent] = [];
          config.subTypes[parent].push(label);
          break;
        case 'StatusWorkflow':
          if (!config.workflows[parent]) config.workflows[parent] = [];
          config.workflows[parent].push(label);
          break;
        case 'Priority':
          config.priorities.push({ label: label, hours: meta1 });
          break;
        default:
          // Unknown category — log but don't crash
          console.warn(`API_getConfig: Unknown category "${category}" in Config sheet.`);
      }
    });

    // Append built-in General defaults
    config.taskTypes.push({ label: 'General', color: '#F3F4F6' });
    config.workflows['General'] = ['New', 'Open', 'In-Progress', 'Completed'];


    // --- 2. Users Sheet ---
    try {
      const usersSheet = getTable('Users');
      const usersData  = usersSheet.getDataRange().getValues();

      if (usersData.length < 2) {
        console.warn('API_getConfig: Users sheet has no data rows.');
      } else {
        const headers  = usersData[0];
        const nameIdx  = headers.indexOf('Names');
        const emailIdx = headers.indexOf('Snap Emails');

        // FIX: Guard against missing required headers instead of silently skipping
        if (nameIdx === -1 || emailIdx === -1) {
          console.error(
            'API_getConfig: Users sheet is missing required headers. ' +
            'Expected "Names" and "Snap Emails". Found: ' + JSON.stringify(headers)
          );
        } else {
          for (let i = 1; i < usersData.length; i++) {
            const name  = String(usersData[i][nameIdx]).trim();
            const email = String(usersData[i][emailIdx]).trim();
            if (email) {
              config.users.push({ name: name, email: email });
            }
          }
        }
      }
    } catch (usersErr) {
      // Non-fatal: app can still run without Users sheet
      console.error('API_getConfig: Could not load Users sheet — ' + usersErr.message);
    }


    return JSON.stringify({ success: true, data: config });

  } catch (e) {
    // Top-level catch: return a safe error payload to the client
    console.error('API_getConfig failed: ' + e.message);
    return JSON.stringify({ success: false, error: e.message });
  }
}