// =============================================================================
// Auth.gs — ICO Center Portal | Authentication & Session Management
// =============================================================================

function _sessionKey(token) { return 'sess_' + token; }

function _writeSession(token, data) {
  const json = JSON.stringify(data);
  CacheService.getScriptCache().put(_sessionKey(token), json, Math.min(CONFIG.SESSION_TTL_SEC, 21600));
  PropertiesService.getScriptProperties().setProperty(_sessionKey(token), json);
}

function _readSession(token) {
  if (!token) return null;
  const key = _sessionKey(token);
  let raw = CacheService.getScriptCache().get(key);
  if (!raw) {
    raw = PropertiesService.getScriptProperties().getProperty(key);
    if (raw) CacheService.getScriptCache().put(key, raw, Math.min(CONFIG.SESSION_TTL_SEC, 21600));
  }
  return raw ? JSON.parse(raw) : null;
}

function _deleteSession(token) {
  const key = _sessionKey(token);
  CacheService.getScriptCache().remove(key);
  PropertiesService.getScriptProperties().deleteProperty(key);
}

// ─── PUBLIC AUTH API ─────────────────────────────────────────────────────────

function login(username, password) {
  try {
    if (!username || !password) return { success: false, error: 'Username and password are required.' };

    const sheet  = getCrmSheet(SHEETS.USERS);
    const rows   = sheet.getDataRange().getValues();
    const hashed = hashPassword(password.trim());
    const uname  = username.trim().toLowerCase();

    let found = null;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toLowerCase() === uname && String(rows[i][1]) === hashed) {
        found = { username: rows[i][0], role: String(rows[i][2]).trim().toLowerCase(), status: rows[i][3] };
        break;
      }
    }

    if (!found) {
      _logActivity('UNKNOWN', 'LOGIN_FAILED', `Bad credentials for: ${username}`);
      return { success: false, error: 'Invalid username or password.' };
    }
    if (String(found.status).toLowerCase() !== 'active') {
      return { success: false, error: 'Account is inactive. Contact your administrator.' };
    }

    // ── Look up center name from source data ──────────────────────────────
    let centerName = '';
    if (found.role === 'center') {
      try {
        const srcSheet = getSourceSpreadsheet().getSheetByName(CONFIG.SOURCE_SHEET_NAME);
        if (srcSheet && srcSheet.getLastRow() > 1) {
          const srcData = srcSheet.getRange(2, 1, srcSheet.getLastRow() - 1, 3).getValues();
          for (const row of srcData) {
            if (String(row[COL.CENTER_CODE]).trim().toLowerCase() === found.username.trim().toLowerCase()) {
              centerName = String(row[COL.CENTER_NAME] || '').trim();
              break;
            }
          }
        }
      } catch (_) { /* non-fatal */ }
    }

    const token = generateToken();
    const sess  = {
      token,
      username  : found.username,
      role      : found.role,
      centerCode: String(found.username).trim(),
      centerName: centerName || found.username,
      createdAt : Date.now(),
      expiresAt : Date.now() + CONFIG.SESSION_TTL_SEC * 1000,
    };

    _writeSession(token, sess);
    _logActivity(found.username, 'LOGIN', 'User logged in.');

    return {
      success   : true,
      token,
      role      : found.role,
      username  : found.username,
      centerCode: found.username,
      centerName: sess.centerName,
    };
  } catch (err) {
    console.error('login():', err);
    return { success: false, error: 'Server error during login.' };
  }
}

function logout(token) {
  try {
    const sess = _readSession(token);
    if (sess) _logActivity(sess.username, 'LOGOUT', 'User logged out.');
    _deleteSession(token);
  } catch (_) {}
  return { success: true };
}

function validateSession(token) {
  try {
    if (!token) return { valid: false };
    const sess = _readSession(token);
    if (!sess)  return { valid: false, reason: 'not_found' };
    if (Date.now() > sess.expiresAt) {
      _deleteSession(token);
      return { valid: false, reason: 'expired' };
    }
    sess.expiresAt = Date.now() + CONFIG.SESSION_TTL_SEC * 1000;
    _writeSession(token, sess);
    return { valid: true, username: sess.username, role: sess.role, centerCode: sess.centerCode, centerName: sess.centerName };
  } catch (err) { return { valid: false }; }
}

function requireAuth(token) {
  const sess = _readSession(token);
  if (!sess) throw new Error('AUTH_REQUIRED');
  if (Date.now() > sess.expiresAt) { _deleteSession(token); throw new Error('SESSION_EXPIRED'); }
  sess.expiresAt = Date.now() + CONFIG.SESSION_TTL_SEC * 1000;
  _writeSession(token, sess);
  return sess;
}

function requireAdmin(token) {
  const sess = requireAuth(token);
  if (sess.role !== 'admin') throw new Error('ADMIN_REQUIRED');
  return sess;
}

function requireCloser(token) {
  const sess = requireAuth(token);
  if (!['admin', 'closer'].includes(sess.role)) throw new Error('CLOSER_REQUIRED');
  return sess;
}

// ─── USER MANAGEMENT (admin-only) ────────────────────────────────────────────

function getUsers(token) {
  try {
    requireAdmin(token);
    const rows  = getCrmSheet(SHEETS.USERS).getDataRange().getValues();
    const users = rows.slice(1).map((r, i) => ({
      rowIndex : i + 2,
      username : r[0],
      role     : r[2],
      status   : r[3],
      createdAt: fmtDate(r[4], 'yyyy-MM-dd HH:mm'),
    }));
    return { success: true, users };
  } catch (e) { return { success: false, error: e.message }; }
}

function createUser(token, userData) {
  try {
    requireAdmin(token);
    const { username, password, role } = userData || {};
    if (!username || !password || !role) return { success: false, error: 'Username, password, and role are required.' };
    if (!['admin','closer','center'].includes(role)) return { success: false, error: 'Role must be "admin", "closer", or "center".' };


    const sheet = getCrmSheet(SHEETS.USERS);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toLowerCase() === username.toLowerCase()) return { success: false, error: 'Username already exists.' };
    }
    sheet.appendRow([username.trim(), hashPassword(password), role, 'active', now()]);
    const sess = _readSession(token);
    _logActivity(sess.username, 'CREATE_USER', `Created user: ${username} (${role})`);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function updateUserStatus(token, username, status) {
  try {
    const sess  = requireAdmin(token);
    const sheet = getCrmSheet(SHEETS.USERS);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === username) {
        sheet.getRange(i + 1, 4).setValue(status);
        _logActivity(sess.username, 'UPDATE_USER', `Set ${username} → ${status}`);
        return { success: true };
      }
    }
    return { success: false, error: 'User not found.' };
  } catch (e) { return { success: false, error: e.message }; }
}

function resetUserPassword(token, username, newPassword) {
  try {
    const sess  = requireAdmin(token);
    const sheet = getCrmSheet(SHEETS.USERS);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === username) {
        sheet.getRange(i + 1, 2).setValue(hashPassword(newPassword));
        _logActivity(sess.username, 'RESET_PASSWORD', `Password reset for: ${username}`);
        return { success: true };
      }
    }
    return { success: false, error: 'User not found.' };
  } catch (e) { return { success: false, error: e.message }; }
}

function deleteUser(token, username) {
  try {
    const sess = requireAdmin(token);
    if (username === sess.username) return { success: false, error: 'Cannot delete your own account.' };
    const sheet = getCrmSheet(SHEETS.USERS);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === username) {
        sheet.deleteRow(i + 1);
        _logActivity(sess.username, 'DELETE_USER', `Deleted user: ${username}`);
        return { success: true };
      }
    }
    return { success: false, error: 'User not found.' };
  } catch (e) { return { success: false, error: e.message }; }
}

// ─── ACTIVITY LOGGING ────────────────────────────────────────────────────────

function _logActivity(user, action, details) {
  try { getCrmSheet(SHEETS.ACTIVITY_LOG).appendRow([now(), user, action, details]); }
  catch (e) { console.error('_logActivity failed:', e); }
}

function getActivityLog(token) {
  try {
    requireAdmin(token);
    const rows = getCrmSheet(SHEETS.ACTIVITY_LOG).getDataRange().getValues();
    const logs = rows.slice(1).reverse().slice(0, 500).map(r => ({
      timestamp: r[0], user: r[1], action: r[2], details: r[3],
    }));
    return { success: true, logs };
  } catch (e) { return { success: false, error: e.message }; }
}