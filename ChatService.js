// =============================================================================
// ChatService.gs — ICO Center Portal | Integrated Chat Backend
// =============================================================================
// All backend logic for the new multi-team chat system.
// Auth is fully delegated to CRM's requireAuth(token) — no separate token system.
// Data lives in the CRM spreadsheet under CHAT_MSGS_V2, CHAT_RING, CHAT_STATUS.
// =============================================================================

const CHAT_MSG_COLS = 11;

// ─── Auth Bridge ─────────────────────────────────────────────────────────────
// Maps a CRM session token → chat identity.
// Returns null if token is invalid.

function _chatId(token) {
  try {
    const sess = requireAuth(token);
    const role = sess.role; // 'admin' | 'closer' | 'center'
    return {
      team    : role === 'admin'  ? 'ICO Admin'
              : role === 'closer' ? 'Closing Team'
              : String(sess.centerCode || sess.username),
      isAdmin : role === 'admin',
      isCloser: role === 'closer' || role === 'admin',
      role    : role,
      username: sess.username,
    };
  } catch(e) { return null; }
}

// ─── Sheet accessors ──────────────────────────────────────────────────────────

function _chatMsgSheet()    { return getCrmSheet(SHEETS.CHAT_MSGS_V2); }
function _chatRingSheet()   { return getCrmSheet(SHEETS.CHAT_RING);    }
function _chatStatusSheet() { return getCrmSheet(SHEETS.CHAT_STATUS);  }

function _rowToChatMsg(row) {
  return {
    id         : Number(row[0]),
    timestamp  : row[1],
    team       : String(row[2] || ''),
    toTeam     : String(row[3] || ''),
    message    : String(row[4] || ''),
    replyTo    : row[5] || null,
    forwarded  : row[6] === true,
    forwardedAt: row[7] || '',
    type       : String(row[8] || 'message'),
    pinned     : row[9] === true,
    pinnedAt   : row[10] || '',
  };
}

// ─── Identity (called on chat init) ──────────────────────────────────────────

function chatGetIdentity(token) {
  const id = _chatId(token);
  if (!id) return { valid: false };
  return { valid: true, team: id.team, isAdmin: id.isAdmin, isCloser: id.isCloser, role: id.role };
}

// ─── Centers list (for Closer/Admin sidebar) ──────────────────────────────────

function chatGetCenters(token) {
  const id = _chatId(token);
  if (!id || !id.isCloser) return { success: false, error: 'Not authorised.' };

  const rows = getCrmSheet(SHEETS.USERS).getDataRange().getValues();
  const centers = [];
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).toLowerCase() === 'center' &&
        String(rows[i][3]).toLowerCase() === 'active') {
      centers.push({ centerCode: String(rows[i][0]) });
    }
  }
  return { success: true, centers };
}

// ─── Send message ─────────────────────────────────────────────────────────────

function chatSendMessage(token, toTeam, message, replyTo) {
  const id = _chatId(token);
  if (!id) return { success: false, error: 'Not authorised.' };

  message = (message || '').trim();
  if (!message)              return { success: false, error: 'Message is empty.' };
  if (message.length > 2000) return { success: false, error: 'Max 2000 characters.' };


  const actualTo = id.isCloser ? (toTeam || 'Closing Team') : 'Closing Team';

  const sheet = _chatMsgSheet();
  const newId = sheet.getLastRow();

  sheet.appendRow([
    newId,
    new Date().toISOString(),
    id.team,
    actualTo,
    message,
    replyTo || '',
    false,
    '',
    'message',
    false,
    ''
  ]);

  return { success: true, id: newId };
}

// ─── Broadcast (admin only) ───────────────────────────────────────────────────

function chatBroadcast(token, message) {
  const id = _chatId(token);
  if (!id || !id.isAdmin) return { success: false, error: 'Not authorised.' };

  message = (message || '').trim();
  if (!message) return { success: false, error: 'Empty message.' };

  const sheet = _chatMsgSheet();
  const newId = sheet.getLastRow();
  sheet.appendRow([newId, new Date().toISOString(), 'ADMIN', 'ALL',
                   message, '', false, '', 'broadcast', false, '']);
  return { success: true, id: newId };
}

// ─── Get messages for a conversation thread ───────────────────────────────────
// convTeam:
//   '__broadcast__'  → broadcast-type messages only
//   a centerCode     → messages between that center and Closing Team
//   'Closing Team'   → messages for a center user's own thread

function chatGetMessages(token, convTeam, afterId) {
  const id = _chatId(token);
  if (!id) return { success: false, error: 'Not authorised.', messages: [] };
  afterId = Number(afterId) || 0;

  const sheet = _chatMsgSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { success: true, messages: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, CHAT_MSG_COLS).getValues();
  const out  = [];

  for (const row of data) {
    const msgId   = Number(row[0]);
    if (msgId <= afterId) continue;
    const msgTeam = String(row[2]);
    const toTeam  = String(row[3]);
    const type    = String(row[8] || 'message');

    if (convTeam === '__broadcast__') {
      if (type === 'broadcast') out.push(_rowToChatMsg(row));
      continue;
    }

    if (id.isCloser) {
      // Closer/Admin sees: messages between Closing Team and the specified center
      const inChannel =
        (msgTeam === convTeam    && toTeam === 'Closing Team') ||
        (msgTeam === 'Closing Team' && toTeam === convTeam)    ||
        (msgTeam === 'ICO Admin'    && toTeam === convTeam)    ||
        (type === 'broadcast');
      if (inChannel) out.push(_rowToChatMsg(row));
    } else {
      // Center sees their own thread with Closing Team
      const inChannel =
        (msgTeam === id.team          && toTeam === 'Closing Team') ||
        (msgTeam === 'Closing Team'   && toTeam === id.team)        ||
        (msgTeam === 'ICO Admin'      && toTeam === id.team)        ||
        (type === 'broadcast');
      if (inChannel) out.push(_rowToChatMsg(row));
    }
  }
  return { success: true, messages: out.slice(-200) };
}

// ─── Pinned messages ──────────────────────────────────────────────────────────

function chatGetPinned(token) {
  const id = _chatId(token);
  if (!id) return { success: false, messages: [] };

  const sheet = _chatMsgSheet();
  const last  = sheet.getLastRow();
  if (last <= 1) return { success: true, messages: [] };

  const data   = sheet.getRange(2, 1, last - 1, CHAT_MSG_COLS).getValues();
  const pinned = data.filter(r => r[9] === true).map(_rowToChatMsg).slice(-50);
  return { success: true, messages: pinned };
}

// ─── Ring alert system ────────────────────────────────────────────────────────

function chatSendRing(token) {
  const id = _chatId(token);
  if (!id)            return { success: false, error: 'Not authorised.' };
  if (id.isCloser)    return { success: false, error: 'Closing team cannot ring itself.' };

  const sheet = _chatRingSheet();
  sheet.appendRow([sheet.getLastRow(), id.team, new Date().toISOString(), false]);
  return { success: true };
}

function chatCheckRing(token) {
  const id = _chatId(token);
  if (!id || !id.isCloser) return { pending: false };

  const sheet = _chatRingSheet();
  const last  = sheet.getLastRow();
  if (last <= 1) return { pending: false };

  const data = sheet.getRange(2, 1, last - 1, 4).getValues();
  for (const row of data) {
    if (row[3] !== true) {
      return { pending: true, id: Number(row[0]), fromTeam: String(row[1]), timestamp: row[2] };
    }
  }
  return { pending: false };
}

function chatClearRing(token) {
  const id = _chatId(token);
  if (!id) return { success: false };

  const sheet = _chatRingSheet();
  const last  = sheet.getLastRow();
  if (last <= 1) return { success: true };

  const data = sheet.getRange(2, 1, last - 1, 4).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][3] !== true) sheet.getRange(i + 2, 4).setValue(true);
  }
  return { success: true };
}

// ─── Team status ──────────────────────────────────────────────────────────────

function chatUpdateStatus(token, status) {
  const id = _chatId(token);
  if (!id) return { success: false };
  if (!['online', 'away', 'busy'].includes(status)) return { success: false };

  const sheet = _chatStatusSheet();
  const last  = sheet.getLastRow();
  if (last > 1) {
    const teams = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < teams.length; i++) {
      if (teams[i][0] === id.team) {
        sheet.getRange(i + 2, 2).setValue(status);
        sheet.getRange(i + 2, 3).setValue(new Date().toISOString());
        return { success: true };
      }
    }
  }
  sheet.appendRow([id.team, status, new Date().toISOString()]);
  return { success: true };
}

function chatGetStatuses(token) {
  const id = _chatId(token);
  if (!id) return { success: false, statuses: [] };

  const sheet = _chatStatusSheet();
  const last  = sheet.getLastRow();
  if (last <= 1) return { success: true, statuses: [] };

  const data = sheet.getRange(2, 1, last - 1, 3).getValues();
  return {
    success  : true,
    statuses : data.filter(r => r[0]).map(r => ({
      team: String(r[0]), status: String(r[1] || 'offline'), updatedAt: r[2],
    })),
  };
}

// ─── Admin message actions ────────────────────────────────────────────────────

function chatPinMessage(token, msgId, shouldPin) {
  const id = _chatId(token);
  if (!id || !id.isAdmin) return { success: false, error: 'Not authorised.' };
  return _mutateChatMsg(msgId, row => {
    row[9]  = shouldPin !== false;
    row[10] = shouldPin !== false ? new Date().toISOString() : '';
    return row;
  });
}

function chatForwardMessage(token, msgId) {
  const id = _chatId(token);
  if (!id || !id.isAdmin) return { success: false, error: 'Not authorised.' };
  return _mutateChatMsg(msgId, row => {
    row[6] = true;
    row[7] = new Date().toISOString();
    return row;
  });
}

function chatDeleteMessage(token, msgId) {
  const id = _chatId(token);
  if (!id || !id.isAdmin) return { success: false, error: 'Not authorised.' };

  msgId = Number(msgId);
  const sheet = _chatMsgSheet();
  const last  = sheet.getLastRow();
  if (last <= 1) return { success: false, error: 'No messages.' };

  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (Number(ids[i][0]) === msgId) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  return { success: false, error: 'Message not found.' };
}

function _mutateChatMsg(msgId, fn) {
  msgId = Number(msgId);
  const sheet = _chatMsgSheet();
  const last  = sheet.getLastRow();
  if (last <= 1) return { success: false, error: 'No messages.' };

  const data = sheet.getRange(2, 1, last - 1, CHAT_MSG_COLS).getValues();
  for (let i = 0; i < data.length; i++) {
    if (Number(data[i][0]) === msgId) {
      sheet.getRange(i + 2, 1, 1, CHAT_MSG_COLS).setValues([fn([...data[i]])]);
      return { success: true };
    }
  }
  return { success: false, error: 'Message not found.' };
}

// ─── Image upload ─────────────────────────────────────────────────────────────
// Uploads a base64-encoded image to Drive, makes it publicly viewable,
// then stores it as a type='image' message row.

function chatUploadImage(token, toTeam, base64Data, mimeType, filename, replyTo) {
  const id = _chatId(token);
  if (!id) return { success: false, error: 'Not authorised.' };

  try {
    // Get or create the shared chat-images folder
    const FOLDER_NAME = 'ICO_Chat_Images';
    let folder;
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);

    // Decode and save to Drive
    const bytes = Utilities.base64Decode(base64Data);
    const blob  = Utilities.newBlob(bytes, mimeType, filename || 'image.jpg');
    const file  = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const url      = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1400-h1400';
    const actualTo = id.isCloser ? (toTeam || 'Closing Team') : 'Closing Team';

    const sheet = _chatMsgSheet();
    const newId = sheet.getLastRow();
    sheet.appendRow([
      newId,
      new Date().toISOString(),
      id.team,
      actualTo,
      url,           // message column holds the Drive URL
      replyTo || '',
      false, '', 'image', false, ''
    ]);

    return { success: true, id: newId, url: url };
  } catch(e) {
    return { success: false, error: e.message };
  }
}