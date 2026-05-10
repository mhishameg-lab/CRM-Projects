// =============================================================================
// DataService.gs — ICO Center Portal | Data Access & Business Logic
// =============================================================================

// ─── SOURCE DATA RETRIEVAL (cached) ──────────────────────────────────────────

function _getSourceData() {
  const cacheKey = 'src_data_v3';
  const cache    = CacheService.getScriptCache();
  const hit      = cache.get(cacheKey);
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }

  const sheet   = getSourceSpreadsheet().getSheetByName(CONFIG.SOURCE_SHEET_NAME);
  if (!sheet) throw new Error('Source sheet "' + CONFIG.SOURCE_SHEET_NAME + '" not found.');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const raw  = sheet.getRange(2, 1, lastRow - 1, 32).getValues();
  const data = raw.filter(r => r.some(c => c !== ''));
  const json = JSON.stringify(data);
  if (json.length < 90000) cache.put(cacheKey, json, CONFIG.CACHE_TTL_SEC);
  return data;
}

function _invalidateSourceCache() {
  CacheService.getScriptCache().remove('src_data_v3');
}

// ─── LEAD-ID MAPPING ─────────────────────────────────────────────────────────

function _getIdMap() {
  const sheet = getCrmSheet(SHEETS.LEAD_ID_MAP);
  const rows  = sheet.getDataRange().getValues();
  const map   = {};
  for (let i = 1; i < rows.length; i++) {
    map[rows[i][0]] = {
      leadId         : rows[i][1],
      centerCode     : rows[i][2],
      watchHash      : rows[i][3],
      watchValuesJson: rows[i][4],
      mapRow         : i + 1,
    };
  }
  return map;
}

function _ensureLeadId(sourceRow, rowData, idMap) {
  if (idMap[sourceRow]) return idMap[sourceRow].leadId;
  const leadId     = generateLeadId(rowData[COL.DATE]);
  const centerCode = String(rowData[COL.CENTER_CODE] || '');
  const watchHash  = _watchHash(rowData);
  const watchVals  = JSON.stringify(_watchValues(rowData));
  getCrmSheet(SHEETS.LEAD_ID_MAP).appendRow([sourceRow, leadId, centerCode, watchHash, watchVals, now()]);
  idMap[sourceRow] = { leadId, centerCode, watchHash, watchValuesJson: watchVals, mapRow: null };
  return leadId;
}

function _watchHash(rowData) {
  return md5(WATCHED_FIELDS.map(f => String(rowData[f.col] || '')).join('|'));
}

function _watchValues(rowData) {
  const obj = {};
  WATCHED_FIELDS.forEach(f => { obj[f.name] = String(rowData[f.col] || ''); });
  return obj;
}

// =============================================================================
// PCP Data Service
// =============================================================================

function _getPcpSourceData() {
  const cacheKey = 'pcp_data_v1';
  const cache    = CacheService.getScriptCache();
  const hit      = cache.get(cacheKey);
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }

  const sheet = getSourceSpreadsheet().getSheetByName(PCP_SOURCE_SHEET_NAME);
  if (!sheet) throw new Error('PCP sheet "' + PCP_SOURCE_SHEET_NAME + '" not found.');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const raw  = sheet.getRange(2, 1, lastRow - 1, 40).getValues();
  const data = raw.filter(r => r.some(c => c !== ''));
  const json = JSON.stringify(data);
  if (json.length < 90000) cache.put(cacheKey, json, CONFIG.CACHE_TTL_SEC);
  return data;
}

function getPcpStats(token) {
  try {
    const sess = requireAuth(token);
    const data = _getPcpSourceData();
    let total = 0, approved = 0, inProcess = 0, denied = 0;
    data.forEach(row => {
      if (!['admin','closer'].includes(sess.role)) {
        if (String(row[PCP_COL.CENTER_CODE]).trim().toLowerCase() !==
            String(sess.centerCode).trim().toLowerCase()) return;
      }
      total++;
      const ps = String(row[PCP_COL.PROC_STATUS_CENTERS] || '').trim().toUpperCase();
      if (ps === 'APPROVED')   approved++;
      if (ps === 'IN PROCESS') inProcess++;
      if (ps === 'DENIED')     denied++;
    });
    return { success: true, total, approved, inProcess, denied };
  } catch(e) { return { success: false, error: e.message }; }
}

function getPcpLeads(token, options) {
  try {
    const sess = requireAuth(token);
    const { page = 1, search = '', filters = {} } = options || {};
    const srch = search.trim().toLowerCase();
    const data = _getPcpSourceData();

    const matched = [];
    data.forEach((row, idx) => {
      if (!['admin','closer'].includes(sess.role)) {
        if (String(row[PCP_COL.CENTER_CODE]).trim().toLowerCase() !==
            String(sess.centerCode).trim().toLowerCase()) return;
      }
      if (srch) {
        const hay = [
          row[PCP_COL.FIRST_NAME], row[PCP_COL.LAST_NAME],
          row[PCP_COL.PHONE],      row[PCP_COL.CENTER_CODE],
          row[PCP_COL.LEAD_TYPE],
        ].join(' ').toLowerCase();
        if (!hay.includes(srch)) return;
      }
      if (filters.procStatus && String(row[PCP_COL.PROC_STATUS_CENTERS]) !== filters.procStatus) return;
      if (filters.centerCode && ['admin','closer'].includes(sess.role) &&
          String(row[PCP_COL.CENTER_CODE]) !== filters.centerCode) return;
      matched.push({ row, sRow: idx + 2 });
    });

    // Sort newest first
    matched.sort((a, b) => {
      const da = new Date(a.row[PCP_COL.TIMESTAMP]);
      const db = new Date(b.row[PCP_COL.TIMESTAMP]);
      if (isNaN(da) && isNaN(db)) return 0;
      if (isNaN(da)) return 1;
      if (isNaN(db)) return -1;
      return db - da;
    });
// How the lead appear in the main sheet in the CRM
    const total      = matched.length;
    const totalPages = Math.max(1, Math.ceil(total / CONFIG.PAGE_SIZE));
    const start      = (page - 1) * CONFIG.PAGE_SIZE;
    const leads      = matched.slice(start, start + CONFIG.PAGE_SIZE).map(({ row, sRow }) => ({
      sRow,
      timestamp        : fmtDate(row[PCP_COL.TIMESTAMP], 'MM/dd/yyyy'),
      centerCode       : String(row[PCP_COL.CENTER_CODE]          || ''),
      fullName         : [row[PCP_COL.FIRST_NAME], row[PCP_COL.LAST_NAME]].filter(Boolean).join(' '),
      phone            : String(row[PCP_COL.PHONE]                || ''),
      leadType         : String(row[PCP_COL.LEAD_TYPE]            || ''),
      procStatusCenters: String(row[PCP_COL.PROC_STATUS_CENTERS]  || ''),
      snsResult        : String(row[PCP_COL.SNS_RESULT]           || ''),
      doNote           : String(row[PCP_COL.DO_NOTE]              || ''),
      procStatusAN     : String(row[PCP_COL.PROC_STATUS_AN]       || ''),

    }));

    _logActivity(sess.username, 'VIEW_PCP', 'Page ' + page + ' search="' + search + '"');
    return { success: true, leads, total, page, totalPages, pageSize: CONFIG.PAGE_SIZE };
  } catch(e) { return { success: false, error: e.message }; }
}

// Detailed info on the lead
function getPcpLeadDetails(token, sRow) {
  try {
    const sess = requireAuth(token);
    const data = _getPcpSourceData();
    const idx  = parseInt(sRow) - 2;
    if (idx < 0 || idx >= data.length) return { success: false, error: 'Record not found.' };
    const row  = data[idx];

    if (!['admin','closer'].includes(sess.role)) {
      if (String(row[PCP_COL.CENTER_CODE]).trim().toLowerCase() !==
          String(sess.centerCode).trim().toLowerCase())
        return { success: false, error: 'Access denied.' };
    }

    return {
      success: true,
      lead: {
        sRow,
        timestamp        : fmtDate(row[PCP_COL.TIMESTAMP], 'yyyy-MM-dd HH:mm'),
        centerCode       : String(row[PCP_COL.CENTER_CODE]          || ''),
        centerName       : String(row[PCP_COL.CENTER_NAME]          || ''),
        closerName       : String(row[PCP_COL.CLOSER_NAME]          || ''),
        doNote           : String(row[PCP_COL.DO_NOTE]              || ''),
        procStatusCenters: String(row[PCP_COL.PROC_STATUS_CENTERS]  || ''),
        snsResult        : String(row[PCP_COL.SNS_RESULT]           || ''),
        leadType         : String(row[PCP_COL.LEAD_TYPE]            || ''),
        requestedProducts: String(row[PCP_COL.REQUESTED_PRODUCTS]   || ''),
        firstName        : String(row[PCP_COL.FIRST_NAME]           || ''),
        lastName         : String(row[PCP_COL.LAST_NAME]            || ''),
        phone            : String(row[PCP_COL.PHONE]                || ''),
        gender           : String(row[PCP_COL.GENDER]               || ''),
        address          : String(row[PCP_COL.ADDRESS]              || ''),
        city             : String(row[PCP_COL.CITY]                 || ''),
        state            : String(row[PCP_COL.STATE]                || ''),
        zip              : String(row[PCP_COL.ZIP]                  || ''),
        dob              : fmtDate(row[PCP_COL.DOB]),
        medId            : String(row[PCP_COL.MED_ID]               || ''),
        height           : String(row[PCP_COL.HEIGHT]               || ''),
        weight           : String(row[PCP_COL.WEIGHT]               || ''),
        shoeSize         : String(row[PCP_COL.SHOE_SIZE]            || ''),
        waistSize        : String(row[PCP_COL.WAIST_SIZE]           || ''),
        doctorName       : String(row[PCP_COL.DOCTOR_NAME]          || ''),
        doctorNpi        : String(row[PCP_COL.DOCTOR_NPI]           || ''),
        doctorPhone      : String(row[PCP_COL.DOCTOR_PHONE]         || ''),
        doctorFax        : String(row[PCP_COL.DOCTOR_FAX]           || ''),
        doctorAddress    : String(row[PCP_COL.DOCTOR_ADDRESS]       || ''),
        doLink           : String(row[PCP_COL.DO_LINK]              || ''),
        cnLink           : String(row[PCP_COL.CN_LINK]              || ''),
        recordLink       : String(row[PCP_COL.RECORD_LINK]          || ''),
        note             : String(row[PCP_COL.NOTE]                 || ''),
        procStatusAN     : String(row[PCP_COL.PROC_STATUS_AN]       || ''),
      }
    };
  } catch(e) { return { success: false, error: e.message }; }
}


function _rowToLead(rowData, sourceRow, idMap) {
  const leadId = _ensureLeadId(sourceRow, rowData, idMap);
  return {
    leadId,
    sourceRow,
    date                    : fmtDate(rowData[COL.DATE]),
    centerCode              : String(rowData[COL.CENTER_CODE]               || ''),
    centerName              : String(rowData[COL.CENTER_NAME]               || ''),
    closerName              : String(rowData[COL.CLOSER_NAME]               || ''),
    leadStatus              : String(rowData[COL.LEAD_STATUS]               || ''),
    closingNotes            : String(rowData[COL.CLOSING_NOTES]             || ''),
    chaserName              : String(rowData[COL.CHASER_NAME]               || ''),
    chaserStatus            : String(rowData[COL.CHASER_STATUS]             || ''),
    chaserNote              : String(rowData[COL.CHASER_NOTE]               || ''),
    processingStatusCenters : String(rowData[COL.PROCESSING_STATUS_CENTERS] || ''),
    processingStatusICO     : String(rowData[COL.PROCESSING_STATUS_ICO]     || ''),
    snsResult               : String(rowData[COL.SNS_RESULT]                || ''),
    leadType                : String(rowData[COL.LEAD_TYPE]                 || ''),
    requestedProducts       : String(rowData[COL.REQUESTED_PRODUCTS]        || ''),
    firstName               : String(rowData[COL.FIRST_NAME]                || ''),
    lastName                : String(rowData[COL.LAST_NAME]                 || ''),
    fullName                : [rowData[COL.FIRST_NAME], rowData[COL.LAST_NAME]].filter(Boolean).join(' '),
    phone                   : String(rowData[COL.PHONE]                     || ''),
    address                 : String(rowData[COL.ADDRESS]                   || ''),
    city                    : String(rowData[COL.CITY]                      || ''),
    state                   : String(rowData[COL.STATE]                     || ''),
    zip                     : String(rowData[COL.ZIP]                       || ''),
    dob                     : fmtDate(rowData[COL.DOB]),
    medId                   : String(rowData[COL.MED_ID]                    || ''),
    height                  : String(rowData[COL.HEIGHT]                    || ''),
    weight                  : String(rowData[COL.WEIGHT]                    || ''),
    shoeSize                : String(rowData[COL.SHOE_SIZE]                 || ''),
    waistSize               : String(rowData[COL.WAIST_SIZE]                || ''),
    gender                  : String(rowData[COL.GENDER]                    || ''),
    doctorName              : String(rowData[COL.DOCTOR_NAME]               || ''),
    doctorPhone             : String(rowData[COL.DOCTOR_PHONE]              || ''),
    doctorFax               : String(rowData[COL.DOCTOR_FAX]                || ''),
    doctorNpi               : String(rowData[COL.DOCTOR_NPI]                || ''),
  };
}

// ─── DASHBOARD STATS ─────────────────────────────────────────────────────────

function getDashboardStats(token) {
  try {
    const sess       = requireAuth(token);
    const sourceData = _getSourceData();
    let total = 0, verified = 0, orderSigned = 0, approved = 0;

    sourceData.forEach(row => {
      if (!['admin','closer'].includes(sess.role)) {
        if (String(row[COL.CENTER_CODE]).trim().toLowerCase() !== String(sess.centerCode).trim().toLowerCase()) return;
      }
      total++;
      const ls = String(row[COL.LEAD_STATUS]               || '').trim();
      const cs = String(row[COL.CHASER_STATUS]             || '').trim();
      const ps = String(row[COL.PROCESSING_STATUS_CENTERS] || '').trim().toUpperCase();
      if (DISPOSITIONS.LEAD_PRODUCTION.includes(ls))    verified++;
      if (DISPOSITIONS.CHASER_SIGNED.includes(cs))      orderSigned++;
      if (DISPOSITIONS.PROC_APPROVED.map(v=>v.toUpperCase()).includes(ps)) approved++;
    });

    return { success: true, total, verified, orderSigned, approved };
  } catch (e) { return { success: false, error: e.message }; }
}

// ─── LEADS (paginated, filtered) ─────────────────────────────────────────────

function getLeads(token, options) {
  try {
    const sess    = requireAuth(token);
    const { page = 1, search = '', filters = {} } = options || {};
    const srch    = search.trim().toLowerCase();
    const sourceData = _getSourceData();
    const idMap      = _getIdMap();

    const matched = [];
    sourceData.forEach((row, idx) => {
      const sRow = idx + 2;

      if (!['admin','closer'].includes(sess.role)) {
        if (String(row[COL.CENTER_CODE]).trim().toLowerCase() !== String(sess.centerCode).trim().toLowerCase()) return;
      }

      if (srch) {
        const haystack = [
          row[COL.FIRST_NAME], row[COL.LAST_NAME], row[COL.PHONE],
          row[COL.CENTER_CODE], row[COL.LEAD_STATUS], row[COL.CHASER_STATUS], row[COL.LEAD_TYPE],
        ].join(' ').toLowerCase();
        if (!haystack.includes(srch)) return;
      }

      if (filters.leadStatus       && String(row[COL.LEAD_STATUS])               !== filters.leadStatus)       return;
      if (filters.chaserStatus     && String(row[COL.CHASER_STATUS])             !== filters.chaserStatus)     return;
      if (filters.processingStatus && String(row[COL.PROCESSING_STATUS_CENTERS]) !== filters.processingStatus) return;
      if (filters.centerCode && sess.role === 'admin' && String(row[COL.CENTER_CODE]) !== filters.centerCode)  return;

      if (filters.dateFrom || filters.dateTo) {
        const d = new Date(row[COL.DATE]);
        if (!isNaN(d)) {
          if (filters.dateFrom && d < new Date(filters.dateFrom)) return;
          if (filters.dateTo   && d > new Date(filters.dateTo))   return;
        }
      }
      matched.push({ row, sRow });
    });

    matched.sort((a, b) => {
      const da = new Date(a.row[COL.DATE]);
      const db = new Date(b.row[COL.DATE]);
      if (isNaN(da) && isNaN(db)) return 0;
      if (isNaN(da)) return 1;
      if (isNaN(db)) return -1;
      return db - da;
    });
    
    const total      = matched.length;
    const totalPages = Math.max(1, Math.ceil(total / CONFIG.PAGE_SIZE));
    const start      = (page - 1) * CONFIG.PAGE_SIZE;
    const slice      = matched.slice(start, start + CONFIG.PAGE_SIZE);

    const leads = slice.map(({ row, sRow }) => {
      const leadId = _ensureLeadId(sRow, row, idMap);
      return {
        leadId,
        date                   : fmtDate(row[COL.DATE]),
        fullName               : [row[COL.FIRST_NAME], row[COL.LAST_NAME]].filter(Boolean).join(' '),
        centerCode             : String(row[COL.CENTER_CODE]               || ''),
        phone                  : String(row[COL.PHONE]                     || ''),
        leadStatus             : String(row[COL.LEAD_STATUS]               || ''),
        chaserStatus           : String(row[COL.CHASER_STATUS]             || ''),
        processingStatusCenters: String(row[COL.PROCESSING_STATUS_CENTERS] || ''),
        leadType               : String(row[COL.LEAD_TYPE]                 || ''),
      };
    });

    _logActivity(sess.username, 'VIEW_LEADS', `Page ${page} | search="${search}"`);
    return { success: true, leads, total, page, totalPages, pageSize: CONFIG.PAGE_SIZE };
  } catch (e) { return { success: false, error: e.message }; }
}

// ─── LEAD DETAILS ────────────────────────────────────────────────────────────

function getLeadDetails(token, leadId) {
  try {
    const sess = requireAuth(token);

    const sourceData = _getSourceData();
    const idMap      = _getIdMap();

    let foundRowData   = null;
    let foundSourceRow = null;

    for (const [sourceRowKey, entry] of Object.entries(idMap)) {
      if (entry.leadId === leadId) {
        const idx = parseInt(sourceRowKey) - 2;
        if (idx >= 0 && idx < sourceData.length) {
          foundRowData   = sourceData[idx];
          foundSourceRow = parseInt(sourceRowKey);
        }
        break;
      }
    }

    if (!foundRowData) return { success: false, error: 'Lead not found.' };

    const lead  = _rowToLead(foundRowData, foundSourceRow, idMap);
    const notes = _getLeadNotes(leadId);

    _logActivity(sess.username, 'VIEW_LEAD', `Lead: ${leadId}`);
    return { success: true, lead, notes };

  } catch (e) { return { success: false, error: e.message }; }
}

// ─── NOTES ───────────────────────────────────────────────────────────────────

function addNote(token, leadId, content) {
  try {
    const sess = requireAuth(token);
    if (!content || !content.trim()) return { success: false, error: 'Note content cannot be empty.' };

    const mapSheet = getCrmSheet(SHEETS.LEAD_ID_MAP);
    const mapRows  = mapSheet.getDataRange().getValues();
    let authorised = false;
    for (let i = 1; i < mapRows.length; i++) {
      if (mapRows[i][1] === leadId) {
        // FIX: closer role can add notes to any lead (same access as viewing)
        authorised = sess.role === 'admin' ||
          sess.role === 'closer' ||
          String(mapRows[i][2]).trim().toLowerCase() === String(sess.centerCode).trim().toLowerCase();
        break;
      }
    }
    if (!authorised) return { success: false, error: 'Access denied.' };

    getCrmSheet(SHEETS.NOTES_LOG).appendRow([now(), leadId, sess.centerCode, content.trim(), sess.username]);
    _logActivity(sess.username, 'ADD_NOTE', `Note on lead: ${leadId}`);
    return { success: true, timestamp: now() };
  } catch (e) { return { success: false, error: e.message }; }
}

function _getLeadNotes(leadId) {
  const rows  = getCrmSheet(SHEETS.NOTES_LOG).getDataRange().getValues();
  const notes = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === leadId) {
      notes.push({ timestamp: rows[i][0], leadId: rows[i][1], centerCode: rows[i][2], content: rows[i][3], createdBy: rows[i][4] });
    }
  }
  return notes.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
}

// ─── NOTIFICATIONS ────────────────────────────────────────────────────────────

function getNotifications(token) {
  try {
    const sess  = requireAuth(token);
    const rows  = getCrmSheet(SHEETS.NOTIFICATIONS).getDataRange().getValues();
    const notifs = [];
    for (let i = 1; i < rows.length; i++) {
      const centerCode = String(rows[i][1]);
      if (sess.role !== 'admin' && centerCode.trim().toLowerCase() !== String(sess.centerCode).trim().toLowerCase()) continue;
      notifs.push({
        id        : String(rows[i][0]),
        centerCode,
        leadId    : rows[i][2],
        messageRaw: rows[i][3],
        status    : rows[i][4],
        timestamp : rows[i][5],
        rowIndex  : i + 1,
      });
    }
    notifs.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    return { success: true, notifications: notifs.slice(0, 100), unreadCount: notifs.filter(n => n.status === 'Unread').length };
  } catch (e) { return { success: false, error: e.message }; }
}

function markNotificationRead(token, notificationId) {
  try {
    const sess  = requireAuth(token);
    const sheet = getCrmSheet(SHEETS.NOTIFICATIONS);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(notificationId)) {
        if (sess.role !== 'admin' && String(rows[i][1]).trim().toLowerCase() !== String(sess.centerCode).trim().toLowerCase())
          return { success: false, error: 'Access denied.' };
        sheet.getRange(i + 1, 5).setValue('Read');
        return { success: true };
      }
    }
    return { success: false, error: 'Notification not found.' };
  } catch (e) { return { success: false, error: e.message }; }
}

function markAllNotificationsRead(token) {
  try {
    const sess  = requireAuth(token);
    const sheet = getCrmSheet(SHEETS.NOTIFICATIONS);
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][4] === 'Unread') {
        const cc = String(rows[i][1]);
        if (sess.role === 'admin' || cc.trim().toLowerCase() === String(sess.centerCode).trim().toLowerCase()) {
          sheet.getRange(i + 1, 5).setValue('Read');
        }
      }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

// ─── INSIGHTS DATA ────────────────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════
// BACKEND ADDITIONS — Google Apps Script
// Add / merge these into your existing getInsightsData()
// ══════════════════════════════════════════════════════════
//
// COLUMN MAP — LG LEADS sheet
//   Col A  = Submission Date
//   Col E  = Lead Status  ("Verified Med B", "Verified PPO", etc.)
//   Col J  = Processing Status  ("SHIPPED RTS", "APPROVED", "DENIED", ...)
//   Chaser Status = whichever column you already use for PROD_CHASER_SIGNED
//                   (the one containing "Order Signed")
//
// COLUMN MAP — PCP LEADS sheet
//   Col A  = Submission Date
//   Col B  = Center Code
//   Col C  = Center Name
//   Col G  = Order Status  ("APPROVED", "DENIED", "SHIPPED RTS")
//
// ══════════════════════════════════════════════════════════

// ── Helper: get Monday of the week for a given date ───────
function getWeekStart_(date) {
  var d   = new Date(date);
  var day = d.getDay();              // 0=Sun … 6=Sat
  var diff = (day === 0) ? -6 : 1 - day;  // shift to Monday
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

// ── Helper: format a date as "MMM D" ─────────────────────
function fmtDate_(d) {
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return months[d.getMonth()] + ' ' + d.getDate();
}

// ── Helper: build last-3-weeks bucket labels ──────────────
function buildWeekBuckets_() {
  var today  = new Date();
  today.setHours(0, 0, 0, 0);
  var buckets = [];
  for (var i = 2; i >= 0; i--) {
    var start = getWeekStart_(today);
    start.setDate(start.getDate() - i * 7);
    var end   = new Date(start);
    end.setDate(end.getDate() + 6);
    buckets.push({
      start : start,
      end   : end,
      label : fmtDate_(start) + '–' + fmtDate_(end),
      count : 0
    });
  }
  return buckets;
}

// ══════════════════════════════════════════════════════════
// MERGE THIS LOGIC INTO YOUR getInsightsData() function
// ══════════════════════════════════════════════════════════
function _buildWeekBuckets() {
  const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const today  = new Date();
  today.setHours(0, 0, 0, 0);
  const day = today.getDay(); // 0=Sun
  const thisMonday = new Date(today);
  thisMonday.setDate(today.getDate() - (day === 0 ? 6 : day - 1));

  return [2, 1, 0].map(function(weeksBack) {
    const start = new Date(thisMonday);
    start.setDate(thisMonday.getDate() - weeksBack * 7);
    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    end.setHours(23, 59, 59, 999);
    const label = MONTHS[start.getMonth()] + ' ' + start.getDate() +
                  '–' + MONTHS[end.getMonth()] + ' ' + end.getDate();
    return { start, end, label, count: 0 };
  });
}
function getInsightsData(token, dateFrom, dateTo) {
  try {
    const sess       = requireAuth(token);
    const sourceData = _getSourceData();

    const leadStatusDist      = {};
    const chaserStatusDist    = {};
    const processingStatusDist = {};
    const leadsOverTime       = {};
    const centerStats         = {};

    // ── LG counters ──────────────────────────────────────────────────
    let total = 0, verified = 0, trash_ls = 0;
    let orderSigned = 0, highPotential = 0, trash_cs = 0;
    let approved = 0, inProcess = 0, denied = 0, rts = 0;

    // Metric 1 — Production Rate  : verified / total
    // Metric 3 — Approval Rate    : approvedFromReturned / returnedDOs
    // Metric 4 — RTS Rate         : rtsFromReturned / returnedDOs
    // Metric 5 — DO Return Rate   : returnedDOs / verified
    let returnedDOs          = 0;   // chaserStatus = "Order Signed"
    let rtsFromReturned      = 0;   // Order Signed AND proc = SHIPPED RTS
    let approvedFromReturned = 0;   // Order Signed AND proc = APPROVED

    // Metric 2 — 45-Day Conversion:
    //   verified leads submitted < 45 days ago that are NOT yet returned
    //   (still in pipeline → high potential to convert)
    let verifiedWithin45     = 0;
    let notReturnedWithin45  = 0;

    // Legacy 45-day order-conversion window (kept for renderConversionDetail)
    let leadsOlderThan45         = 0;
    let orderSignedOlderThan45   = 0;

    const cutoff45 = new Date();
    cutoff45.setDate(cutoff45.getDate() - 45);
    cutoff45.setHours(0, 0, 0, 0);

    // Metric 6 — Weekly LG tracking
    const lgWeekBuckets = _buildWeekBuckets();

    sourceData.forEach(row => {
      if (!['admin', 'closer'].includes(sess.role)) {
        if (String(row[COL.CENTER_CODE]).trim().toLowerCase() !==
            String(sess.centerCode).trim().toLowerCase()) return;
      }

      const rowDate = row[COL.DATE] ? new Date(row[COL.DATE]) : null;

      if (dateFrom || dateTo) {
        if (rowDate && !isNaN(rowDate)) {
          if (dateFrom && rowDate < new Date(dateFrom)) return;
          if (dateTo   && rowDate > new Date(dateTo))   return;
        }
      }

      total++;

      const ls  = String(row[COL.LEAD_STATUS]               || '').trim();
      const cs  = String(row[COL.CHASER_STATUS]             || '').trim();
      const ps  = String(row[COL.PROCESSING_STATUS_CENTERS] || '').trim();
      const cc  = String(row[COL.CENTER_CODE]               || 'Unknown');
      const psU = ps.toUpperCase();

      leadStatusDist[ls || 'Unknown']       = (leadStatusDist[ls || 'Unknown']       || 0) + 1;
      chaserStatusDist[cs || 'Unknown']     = (chaserStatusDist[cs || 'Unknown']     || 0) + 1;
      processingStatusDist[ps || 'Unknown'] = (processingStatusDist[ps || 'Unknown'] || 0) + 1;
      centerStats[cc] = (centerStats[cc] || 0) + 1;

      const isVerified  = DISPOSITIONS.LEAD_PRODUCTION.includes(ls);
      const isReturned  = DISPOSITIONS.CHASER_SIGNED.includes(cs);   // "Order Signed"
      const isRTS       = psU === 'SHIPPED RTS';
      const isApproved  = psU === 'APPROVED';

      // Lead Status bucket
      if (isVerified) verified++; else trash_ls++;

      // Chaser Status bucket
      if (isReturned) {
        orderSigned++;
        returnedDOs++;
        if (isRTS)      rtsFromReturned++;
        if (isApproved) approvedFromReturned++;
      } else if (DISPOSITIONS.CHASER_POTENTIAL.includes(cs)) {
        highPotential++;
      } else if (cs) {
        trash_cs++;
      }

      // Processing Status bucket
      if (isApproved)           approved++;
      else if (psU === 'IN PROCESS') inProcess++;
      else if (psU === 'DENIED')     denied++;
      else if (isRTS)                rts++;

      // Metric 2: verified leads submitted within last 45 days
      if (isVerified && rowDate && !isNaN(rowDate) && rowDate >= cutoff45) {
        verifiedWithin45++;
        if (!isReturned) notReturnedWithin45++;
      }

      // Legacy: leads older than 45 days (for 45-day order conversion card)
      if (rowDate && !isNaN(rowDate) && rowDate < cutoff45) {
        leadsOlderThan45++;
        if (isReturned) orderSignedOlderThan45++;
      }

      // Weekly LG bucket
      if (rowDate && !isNaN(rowDate)) {
        for (const b of lgWeekBuckets) {
          if (rowDate >= b.start && rowDate <= b.end) { b.count++; break; }
        }
        try {
          const key = Utilities.formatDate(rowDate, _tz(), 'yyyy-MM');
          leadsOverTime[key] = (leadsOverTime[key] || 0) + 1;
        } catch (_) {}
      }
    });

    const timeLabels = Object.keys(leadsOverTime).sort();
    const rtsRate    = returnedDOs > 0  ? Math.round((rtsFromReturned / returnedDOs) * 100) : 0;
    const prodRate   = total > 0        ? Math.round((verified  / total)             * 100) : 0;
    const orderConversionRate = leadsOlderThan45 > 0
      ? Math.round((orderSignedOlderThan45 / leadsOlderThan45) * 100)
      : null;

    // ── PCP Leads (Metrics 7–12) ──────────────────────────────────────
    const pcpData       = _getPcpSourceData();
    const pcpWeekBuckets = _buildWeekBuckets();
    let pcpTotal = 0, pcpApproved = 0, pcpDenied = 0, pcpRts = 0;

    pcpData.forEach(row => {
      if (!['admin', 'closer'].includes(sess.role)) {
        if (String(row[PCP_COL.CENTER_CODE]).trim().toLowerCase() !==
            String(sess.centerCode).trim().toLowerCase()) return;
      }

      const pDate = row[PCP_COL.TIMESTAMP] ? new Date(row[PCP_COL.TIMESTAMP]) : null;

      if (dateFrom || dateTo) {
        if (pDate && !isNaN(pDate)) {
          if (dateFrom && pDate < new Date(dateFrom)) return;
          if (dateTo   && pDate > new Date(dateTo))   return;
        }
      }

      pcpTotal++;
      const pStatus = String(row[PCP_COL.PROC_STATUS_CENTERS] || '').trim().toUpperCase();
      if (pStatus === 'APPROVED')    pcpApproved++;
      if (pStatus === 'DENIED')      pcpDenied++;
      if (pStatus === 'SHIPPED RTS') pcpRts++;

      if (pDate && !isNaN(pDate)) {
        for (const b of pcpWeekBuckets) {
          if (pDate >= b.start && pDate <= b.end) { b.count++; break; }
        }
      }
    });

    return {
      success: true,
      totalLeads: total,
      leadStatusDist, chaserStatusDist, processingStatusDist, centerStats,
      leadsOverTime: { labels: timeLabels, data: timeLabels.map(k => leadsOverTime[k]) },

      production: {
        // Core counts
        verified, trash: trash_ls, prodRate,
        orderSigned, highPotential, trash_cs,
        approved, inProcess, denied, rts,

        // New fields for quality cards
        returnedDOs,
        rtsFromReturned,
        approvedFromReturned,
        verifiedWithin45,
        notReturnedWithin45,

        // Rates
        rtsRate,
        rtsWarning: rtsRate > RTS_WARNING_PCT,

        // Legacy conversion window
        orderConversionRate,
        leadsOlderThan45,
        orderSignedOlderThan45,
        orderTargetLow : ORDER_TARGET_LOW,
        orderTargetHigh: ORDER_TARGET_HIGH,
        rtsTargetLow   : RTS_TARGET_LOW,
        rtsTargetHigh  : RTS_TARGET_HIGH,
      },

      // Metric 6 — LG weekly
      weeklyLG: lgWeekBuckets.map(b => ({ label: b.label, count: b.count })),

      // Metrics 7–11 — PCP
      pcp: { total: pcpTotal, approved: pcpApproved, denied: pcpDenied, rts: pcpRts },

      // Metric 12 — PCP weekly
      weeklyPCP: pcpWeekBuckets.map(b => ({ label: b.label, count: b.count })),
    };
  } catch (e) { return { success: false, error: e.message }; }
}

// ─── FILTER OPTIONS ───────────────────────────────────────────────────────────

function getFilterOptions(token) {
  try {
    const sess       = requireAuth(token);
    const sourceData = _getSourceData();
    const lsSet = new Set(), csSet = new Set(), psSet = new Set(), ccSet = new Set();

    sourceData.forEach(row => {
      if (!['admin','closer'].includes(sess.role)) {
        if (String(row[COL.CENTER_CODE]).trim().toLowerCase() !== String(sess.centerCode).trim().toLowerCase()) return;
      }
      if (row[COL.LEAD_STATUS])                lsSet.add(String(row[COL.LEAD_STATUS]));
      if (row[COL.CHASER_STATUS])              csSet.add(String(row[COL.CHASER_STATUS]));
      if (row[COL.PROCESSING_STATUS_CENTERS])  psSet.add(String(row[COL.PROCESSING_STATUS_CENTERS]));
      if (['admin','closer'].includes(sess.role) && row[COL.CENTER_CODE]) ccSet.add(String(row[COL.CENTER_CODE]));
    });

    return {
      success            : true,
      leadStatuses       : [...lsSet].filter(Boolean).sort(),
      chaserStatuses     : [...csSet].filter(Boolean).sort(),
      processingStatuses : [...psSet].filter(Boolean).sort(),
      centerCodes        : ['admin','closer'].includes(sess.role) ? [...ccSet].filter(Boolean).sort() : [],
    };
  } catch (e) { return { success: false, error: e.message }; }
}

// ─── CHANGE DETECTION ────────────────────────────────────────────────────────

function runChangeDetectionTrigger() {
  try { _runChangeDetectionInternal(); }
  catch (e) { console.error('runChangeDetectionTrigger:', e); }
}

function runChangeDetection(token) {
  try { requireAuth(token); return _runChangeDetectionInternal(); }
  catch (e) { return { success: false, error: e.message }; }
}

function _runChangeDetectionInternal() {
  const sourceData = _getSourceData();
  const idMap      = _getIdMap();
  const mapSheet   = getCrmSheet(SHEETS.LEAD_ID_MAP);
  const auditSheet = getCrmSheet(SHEETS.AUDIT_LOG);
  const notifSheet = getCrmSheet(SHEETS.NOTIFICATIONS);
  let newNotifications = 0;

  sourceData.forEach((row, idx) => {
    const sourceRow   = idx + 2;
    const currentHash = _watchHash(row);

    if (idMap[sourceRow]) {
      const existing = idMap[sourceRow];
      if (existing.watchHash !== currentHash) {
        const leadId     = existing.leadId;
        const centerCode = String(row[COL.CENTER_CODE] || '');
        const prevVals   = existing.watchValuesJson ? JSON.parse(existing.watchValuesJson) : {};
        const currVals   = _watchValues(row);
        const changes    = [];

        WATCHED_FIELDS.forEach(f => {
          const oldV = prevVals[f.name] || '';
          const newV = currVals[f.name] || '';
          if (oldV !== newV) {
            changes.push({ field: f.name, from: oldV, to: newV });
            auditSheet.appendRow([now(), leadId, f.name, oldV, newV, 'SYSTEM']);
          }
        });

        if (changes.length) {
          const msgObj  = { changes: changes.filter(c => c.to) };
          const notifId = generateId('N');
          notifSheet.appendRow([notifId, centerCode, leadId, JSON.stringify(msgObj), 'Unread', now()]);
          newNotifications++;
        }

        if (existing.mapRow) {
          mapSheet.getRange(existing.mapRow, 4, 1, 2).setValues([[currentHash, JSON.stringify(currVals)]]);
        }
      }
    } else {
      _ensureLeadId(sourceRow, row, idMap);
    }
  });

  return { success: true, newNotifications };
}

// ─── AUDIT LOG ────────────────────────────────────────────────────────────────

function getAuditLog(token, leadId) {
  try {
    requireAdmin(token);
    const rows = getCrmSheet(SHEETS.AUDIT_LOG).getDataRange().getValues();
    const logs = rows.slice(1)
      .filter(r => !leadId || r[1] === leadId)
      .reverse().slice(0, 200)
      .map(r => ({ timestamp: r[0], leadId: r[1], fieldName: r[2], oldValue: r[3], newValue: r[4], changedBy: r[5] }));
    return { success: true, logs };
  } catch (e) { return { success: false, error: e.message }; }
}

// ─── TRIGGER SETUP ───────────────────────────────────────────────────────────

function installChangeDetectionTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runChangeDetectionTrigger') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runChangeDetectionTrigger').timeBased().everyMinutes(10).create();
  return { success: true, message: 'Change-detection trigger installed (every 10 min).' };
}


