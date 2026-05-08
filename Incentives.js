// =============================================================================
// Incentives.gs — ICO Center Portal | Incentives & Chat Management
// =============================================================================

// ─── INCENTIVES ───────────────────────────────────────────────────────────────

function getIncentives(token) {
  try {
    requireAdmin(token);
    const sheet = getCrmSheet(SHEETS.INCENTIVES);
    const rows  = sheet.getDataRange().getValues();
    const incentives = rows.slice(1).map((r, i) => ({
      id          : String(r[0]),
      prize       : String(r[1]),
      weeklyTarget: String(r[2]),
      rules       : String(r[3]),
      whoIncluded : String(r[4]),
      active      : String(r[5]).toLowerCase() === 'true',
      createdBy   : String(r[6]),
      createdAt   : fmtDate(r[7], 'yyyy-MM-dd HH:mm'),
      rowIndex    : i + 2,
    }));
    return { success: true, incentives };
  } catch (e) { return { success: false, error: e.message }; }
}

function createIncentive(token, data) {
  try {
    const sess = requireAdmin(token);
    const { prize, weeklyTarget, rules, whoIncluded } = data || {};
    if (!prize || !prize.trim())         return { success: false, error: 'Prize label is required.' };
    if (!weeklyTarget || !weeklyTarget.trim()) return { success: false, error: 'Weekly target is required.' };

    const id = generateId('INC');
    getCrmSheet(SHEETS.INCENTIVES).appendRow([
      id,
      prize.trim(),
      weeklyTarget.trim(),
      (rules || '').trim(),
      (whoIncluded || 'All Centers').trim(),
      'false',
      sess.username,
      now(),
    ]);
    _logActivity(sess.username, 'CREATE_INCENTIVE', `Prize: ${prize} | Target: ${weeklyTarget}`);
    return { success: true, id };
  } catch (e) { return { success: false, error: e.message }; }
}

/**
 * Activate or deactivate an incentive.
 * Only ONE incentive can be active at a time.
 */
function updateIncentiveStatus(token, id, active) {
  try {
    const sess  = requireAdmin(token);
    const sheet = getCrmSheet(SHEETS.INCENTIVES);
    const rows  = sheet.getDataRange().getValues();
    let found   = false;

    if (active) {
      for (let i = 1; i < rows.length; i++) {
        sheet.getRange(i + 1, 6).setValue('false');
      }
    }

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.getRange(i + 1, 6).setValue(active ? 'true' : 'false');
        _logActivity(sess.username, 'UPDATE_INCENTIVE', `${id} → ${active ? 'ACTIVATED' : 'DEACTIVATED'}`);
        found = true;
        break;
      }
    }

    if (!found) return { success: false, error: 'Incentive not found.' };
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function deleteIncentive(token, id) {
  try {
    const sess  = requireAdmin(token);
    const sheet = getCrmSheet(SHEETS.INCENTIVES);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        _logActivity(sess.username, 'DELETE_INCENTIVE', `Deleted incentive: ${id}`);
        return { success: true };
      }
    }
    return { success: false, error: 'Incentive not found.' };
  } catch (e) { return { success: false, error: e.message }; }
}

function getActiveIncentive(token) {
  try {
    requireAuth(token);
    const sheet = getCrmSheet(SHEETS.INCENTIVES);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][5]).toLowerCase() === 'true') {
        return {
          success   : true,
          incentive : {
            id          : String(rows[i][0]),
            prize       : String(rows[i][1]),
            weeklyTarget: String(rows[i][2]),
            rules       : String(rows[i][3]),
            whoIncluded : String(rows[i][4]),
          },
        };
      }
    }
    return { success: true, incentive: null };
  } catch (e) { return { success: false, error: e.message }; }
}

// ─── LEGACY CHAT (V1) — kept for backward compatibility, not used by new UI ──

function getChatMessages(token, centerCode, limit) {
  try {
    const sess  = requireAuth(token);
    const sheet = getCrmSheet(SHEETS.CHAT_MESSAGES);
    const rows  = sheet.getDataRange().getValues();
    const lim   = parseInt(limit) || 100;

    const targetCC = sess.role === 'admin'
      ? (centerCode ? String(centerCode).trim().toLowerCase() : null)
      : String(sess.centerCode).trim().toLowerCase();

    const messages = [];
    for (let i = 1; i < rows.length; i++) {
      const rowCC = String(rows[i][3]).trim().toLowerCase();
      if (targetCC && rowCC !== targetCC) continue;
      messages.push({
        id           : String(rows[i][0]),
        sender       : String(rows[i][1]),
        role         : String(rows[i][2]),
        centerCode   : String(rows[i][3]),
        message      : String(rows[i][4]),
        timestamp    : rows[i][5],
        readByAdmin  : String(rows[i][6]) === 'true',
        readByCenter : String(rows[i][7]) === 'true',
      });
    }

    const sliced = messages.slice(-lim);
    let unread = 0;
    sliced.forEach(function(m) {
      if (sess.role === 'admin'   && m.role !== 'admin' && !m.readByAdmin)  unread++;
      if (sess.role !== 'admin'   && m.role === 'admin' && !m.readByCenter) unread++;
    });

    return { success: true, messages: sliced, unread };
  } catch (e) { return { success: false, error: e.message }; }
}

function getChatCenters(token) {
  try {
    requireAdmin(token);
    const sheet = getCrmSheet(SHEETS.CHAT_MESSAGES);
    const rows  = sheet.getDataRange().getValues();

    const centerMap = {};
    for (let i = 1; i < rows.length; i++) {
      const cc     = String(rows[i][3]).trim();
      const role   = String(rows[i][2]);
      const unread = role !== 'admin' && String(rows[i][6]) !== 'true';
      if (!cc) continue;
      if (!centerMap[cc]) centerMap[cc] = { centerCode: cc, total: 0, unread: 0 };
      centerMap[cc].total++;
      if (unread) centerMap[cc].unread++;
    }

    const centers = Object.values(centerMap).sort((a, b) => b.unread - a.unread || a.centerCode.localeCompare(b.centerCode));
    return { success: true, centers };
  } catch (e) { return { success: false, error: e.message }; }
}

function sendChatMessage(token, message, targetCenterCode) {
  try {
    const sess = requireAuth(token);
    if (!message || !message.trim()) return { success: false, error: 'Message cannot be empty.' };

    let cc;
    if (sess.role === 'admin') {
      if (!targetCenterCode) return { success: false, error: 'Admin must specify a target center.' };
      cc = String(targetCenterCode).trim();
    } else {
      cc = String(sess.centerCode).trim();
    }

    const id = generateId('MSG');
    getCrmSheet(SHEETS.CHAT_MESSAGES).appendRow([
      id,
      sess.username,
      sess.role,
      cc,
      message.trim(),
      now(),
      sess.role === 'admin' ? 'true' : 'false',
      sess.role !== 'admin' ? 'true' : 'false',
    ]);

    _logActivity(sess.username, 'CHAT_MSG', `To center: ${cc} | ${message.trim().substring(0, 60)}`);
    return { success: true, id };
  } catch (e) { return { success: false, error: e.message }; }
}

function markChatRead(token, centerCode) {
  try {
    const sess  = requireAuth(token);
    const sheet = getCrmSheet(SHEETS.CHAT_MESSAGES);
    const rows  = sheet.getDataRange().getValues();

    const targetCC = sess.role === 'admin'
      ? String(centerCode || '').trim().toLowerCase()
      : String(sess.centerCode).trim().toLowerCase();

    for (let i = 1; i < rows.length; i++) {
      const rowCC = String(rows[i][3]).trim().toLowerCase();
      if (targetCC && rowCC !== targetCC) continue;

      if (sess.role === 'admin' && String(rows[i][2]) !== 'admin' && String(rows[i][6]) !== 'true') {
        sheet.getRange(i + 1, 7).setValue('true');
      }
      if (sess.role !== 'admin' && String(rows[i][2]) === 'admin' && String(rows[i][7]) !== 'true') {
        sheet.getRange(i + 1, 8).setValue('true');
      }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ─── LEGACY unread counter — kept but superseded by getNewChatUnread ──────────

function getChatUnread(token) {
  try {
    const sess  = requireAuth(token);
    const sheet = getCrmSheet(SHEETS.CHAT_MESSAGES);
    const rows  = sheet.getDataRange().getValues();
    let   unread = 0;

    for (let i = 1; i < rows.length; i++) {
      const role   = String(rows[i][2]);
      const rowCC  = String(rows[i][3]).trim().toLowerCase();
      const readByAdmin  = String(rows[i][6]) === 'true';
      const readByCenter = String(rows[i][7]) === 'true';

      if (sess.role === 'admin' && role !== 'admin' && !readByAdmin)  unread++;
      if (sess.role !== 'admin' && role === 'admin' && !readByCenter
          && rowCC === String(sess.centerCode).trim().toLowerCase())  unread++;
    }
    return { success: true, unread };
  } catch (e) { return { success: false, error: e.message }; }
}

// ─── NEW CHAT UNREAD COUNTER (V2) ─────────────────────────────────────────────
// Reads from CHAT_MSGS_V2 (the new multi-team chat sheet).
// For closers/admins: counts messages from centers received in the last 5 minutes.
// For centers: counts messages from Closing Team / ICO Admin received in last 5 min.
// This is a time-window heuristic used for the nav badge only (no per-user read flag
// in V2). A full read-receipt column can be added later if needed.

function getNewChatUnread(token) {
  try {
    const sess  = requireAuth(token);
    const sheet = getCrmSheet(SHEETS.CHAT_MSGS_V2);
    const last  = sheet.getLastRow();
    if (last <= 1) return { success: true, unread: 0 };

    const data       = sheet.getRange(2, 1, last - 1, 9).getValues();
    const fiveMinAgo = Date.now() - 5 * 60 * 1000;
    let   unread     = 0;

    const isCloserOrAdmin = ['admin', 'closer'].includes(sess.role);

    data.forEach(function(row) {
      const ts     = new Date(row[1]);
      const from   = String(row[2] || '');
      const toTeam = String(row[3] || '');
      const type   = String(row[8] || 'message');

      if (isNaN(ts) || ts.getTime() <= fiveMinAgo) return;
      if (type === 'broadcast') return; // broadcasts are announcements, not unread signals

      if (isCloserOrAdmin) {
        // Closer/Admin: unread = messages FROM any center TO Closing Team
        if (from !== 'Closing Team' && from !== 'ICO Admin' && toTeam === 'Closing Team') {
          unread++;
        }
      } else {
        // Center: unread = messages from Closing Team or ICO Admin directed at this center
        const myTeam = String(sess.centerCode || sess.username);
        if ((from === 'Closing Team' || from === 'ICO Admin') && toTeam === myTeam) {
          unread++;
        }
      }
    });

    return { success: true, unread };
  } catch (e) { return { success: false, unread: 0, error: e.message }; }
}