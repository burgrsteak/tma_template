/**
 * Wraps a database write operation in LockService.
 * @param {Function} callback - The write logic to execute.
 */
function withLock(callback) {
  const lock = LockService.getScriptLock();
  // Wait up to 10 seconds for other processes to finish
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