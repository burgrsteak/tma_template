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
 *
 * Matching strategy (most-to-least strict):
 *   1. Exact case-insensitive + trimmed match against COLUMN_MAP value.
 *   2. Fuzzy match: strip all non-alphanumeric chars and compare — handles
 *      headers like "Email Address", "e mail", "emailAddress", etc.
 *   3. If still not found AND required=true → throws a descriptive error so
 *      the developer sees the real column name and can update COLUMN_MAP.
 *
 * @param {Array}   headers  - The first row of a sheet (array of strings).
 * @param {string}  colKey   - A key from COLUMN_MAP.
 * @param {boolean} required - Throw when not found. Default true.
 * @returns {number} Zero-based column index, or -1 when not required & missing.
 */
function colIdx(headers, colKey, required) {
  if (required === undefined) required = true;

  const colName = COLUMN_MAP[colKey];
  if (!colName) throw new Error('Unknown COLUMN_MAP key: ' + colKey);

  const normalised = headers.map(h => String(h).trim().toLowerCase());

  // 1. Exact (case-insensitive, trimmed) match
  let idx = normalised.indexOf(colName.toLowerCase());

  // 2. Fuzzy match — strip everything except a-z0-9
  if (idx === -1) {
    const fuzzyTarget = colName.toLowerCase().replace(/[^a-z0-9]/g, '');
    idx = normalised.findIndex(h => h.replace(/[^a-z0-9]/g, '') === fuzzyTarget);
  }

  if (idx === -1 && required) {
    // Log the actual headers so the developer knows exactly what to fix.
    console.error(
      'colIdx: Column "' + colName + '" not found. ' +
      'Actual headers: [' + headers.map(h => '"' + String(h).trim() + '"').join(', ') + ']. ' +
      'Update COLUMN_MAP["' + colKey + '"] to match your sheet.'
    );
    throw new Error(
      'Column "' + colName + '" not found in sheet. ' +
      'Actual headers: ' + headers.map(h => String(h).trim()).join(', ') + '. ' +
      'Check COLUMN_MAP.'
    );
  }

  return idx;
}
