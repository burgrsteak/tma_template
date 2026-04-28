/**
 * Wraps a database write operation in LockService.
 * @param {Function} callback - The write logic to execute.
 */
function withLock(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return callback();
  } catch (e) {
    throw new Error('Database error: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Generates a standard UUID v4.
 */
function generateUUID() {
  return Utilities.getUuid();
}

/**
 * Safely parses incoming JSON payloads.
 */
function parsePayload(jsonString) {
  try {
    return JSON.parse(jsonString);
  } catch (e) {
    throw new Error('Invalid JSON payload received.');
  }
}

/**
 * COLUMN_MAP — centralised column-name registry.
 * Change a value here to rename a column across the whole app.
 * Keys are the logical names used in code; values are the actual
 * sheet column headers.
 */
const COLUMN_MAP = {
  // Users sheet
  USER_EMAIL:       'email',
  USER_NAME:        'name',
  USER_ROLE:        'role',
  USER_TEAM:        'team',
  USER_SHIFT_START: 'shiftStart',
  USER_SHIFT_END:   'shiftEnd',
  USER_SHIFT_NAME:  'scheduledShift',

  // SessionLogs sheet
  SESSION_ID:       'id',
  SESSION_EMAIL:    'email',
  SESSION_TIME_IN:  'timeIn',
  SESSION_TIME_OUT: 'timeOut',
  SESSION_REMARK:   'remark',
  SESSION_TEAM:     'team',
  SESSION_START:    'shiftStart',
  SESSION_END:      'shiftEnd'
};

/**
 * Returns the 0-based index of a column in a header row.
 * Throws a descriptive error if the column is not found, so
 * callers never silently work on the wrong column.
 *
 * @param {Array}  headers   - The first row of a sheet (array of strings).
 * @param {string} colKey    - A key from COLUMN_MAP.
 * @param {boolean} required - If true, throws when not found. Default true.
 * @returns {number} Zero-based column index, or -1 when not required & missing.
 */
function colIdx(headers, colKey, required) {
  if (required === undefined) required = true;
  const colName = COLUMN_MAP[colKey];
  if (!colName) throw new Error('Unknown COLUMN_MAP key: ' + colKey);
  const idx = headers.map(h => String(h).trim().toLowerCase()).indexOf(colName.toLowerCase());
  if (idx === -1 && required) {
    throw new Error('Column "' + colName + '" not found in sheet. Check COLUMN_MAP.');
  }
  return idx;
}
