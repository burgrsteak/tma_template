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
  if (!sheet) throw new Error(`Table ${tableName} not found in database.`);
  return sheet;
}

/**
 * Fetches and structures all dynamic configuration data from the Config sheet.
 */
/**
 * Fetches and structures all dynamic configuration data from the Config and Users sheets.
 */
function API_getConfig() {
  // 1. Get standard Config
  const configSheet = getTable('Config');
  const configData = configSheet.getDataRange().getValues(); // Bulk read
  const configRows = configData.slice(1);
  
  let config = {
    taskTypes: [{ label: 'All', color: '#E5E7EB' }],
    subTypes: {},
    workflows: {},
    priorities: [],
    users: [] // NEW: Array to hold our team members
  };

  configRows.forEach(row => {
    const category = row[0];
    const label = row[1];
    const parent = row[2];
    const meta1 = row[3];
    
    if (category === 'TaskType') {
      config.taskTypes.push({ label: label, color: meta1 || '#F3F4F6' });
    } else if (category === 'SubType') {
      if (!config.subTypes[parent]) config.subTypes[parent] = [];
      config.subTypes[parent].push(label);
    } else if (category === 'StatusWorkflow') {
      if (!config.workflows[parent]) config.workflows[parent] = [];
      config.workflows[parent].push(label);
    } else if (category === 'Priority') {
      config.priorities.push({ label: label, hours: meta1 });
    }
  });

  config.taskTypes.push({ label: 'General', color: '#F3F4F6' });
  config.workflows['General'] = ['New', 'Open', 'In-Progress', 'Completed'];

  // 2. Get Users List for Dropdowns
  try {
    const usersSheet = getTable('Users');
    const usersData = usersSheet.getDataRange().getValues();
    const nameIdx = usersData[0].indexOf('Names');
    const emailIdx = usersData[0].indexOf('Snap Emails');
    
    if (nameIdx !== -1 && emailIdx !== -1) {
      for(let i = 1; i < usersData.length; i++) {
        if(usersData[i][emailIdx]) {
          config.users.push({ name: usersData[i][nameIdx], email: usersData[i][emailIdx] });
        }
      }
    }
  } catch(e) {
    // Failsafe in case Users tab isn't set up yet
    console.error('Users table not found or formatted incorrectly.');
  }

  return JSON.stringify({ success: true, data: config }); // JSON Transport
}