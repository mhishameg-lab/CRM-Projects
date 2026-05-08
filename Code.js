// =============================================================================
// Code.gs — ICO Center Portal | Main Entry Point & Shared Utilities
// =============================================================================

const CONFIG = {
  SOURCE_SHEET_ID   : '15ux5w8d15WFFFes5XeGgNZw7ljg4sPWMV54qPr7Umzo',
  SOURCE_SHEET_NAME : 'Outsourcing Leads Generation',
  CRM_SHEET_ID      : '1-o2vpphU2kvykz63jjes8Afuha-Fol_UoBQt5nmPtco',
  SESSION_TTL_SEC   : 8 * 3600,
  CACHE_TTL_SEC     : 300,
  PAGE_SIZE         : 50,
  APP_TITLE         : 'ICO Center Portal',
  WEBFORM_URL       : 'https://script.google.com/macros/s/AKfycbxecKhXkD7fddhw0YoOleY6DDHYP7gFaqKd65onZjOnr7GqJpcthquXVfO8tqlv01kKPg/exec',
};

// ─── SHEET NAMES ─────────────────────────────────────────────────────────────

const SHEETS = {
  USERS        : 'USERS',
  NOTES_LOG    : 'NOTES_LOG',
  AUDIT_LOG    : 'AUDIT_LOG',
  NOTIFICATIONS: 'NOTIFICATIONS',
  ACTIVITY_LOG : 'ACTIVITY_LOG',
  LEAD_ID_MAP  : 'LEAD_ID_MAP',
  INCENTIVES   : 'INCENTIVES',
  CHAT_MESSAGES: 'CHAT_MESSAGES',
  CHAT_MSGS_V2 : 'ICO_Chat_Messages',   // ← ADD: new chat (multi-team)
  CHAT_RING    : 'ICO_Chat_Ring',        // ← ADD: ring alerts
  CHAT_STATUS  : 'ICO_Chat_Status',      // ← ADD: team presence};
};


// =============================================================================
// PCP MODULE — Column indices & config
// =============================================================================

const PCP_SOURCE_SHEET_NAME = 'Outsourcing PCP Processing';

const PCP_COL = {
  TIMESTAMP          : 0,
  CENTER_CODE        : 1,
  CENTER_NAME        : 2,
  CLOSER_NAME        : 3,
  DOCa_REVIEW        : 4,
  NOTE               : 5,   // ✅ new column added here
  PROC_STATUS_CENTERS: 6,
  PROC_STATUS_ICO    : 7,
  SNS_RESULT         : 8,
  LEAD_TYPE          : 9,
  REQUESTED_PRODUCTS : 10,
  FIRST_NAME         : 11,
  LAST_NAME          : 12,
  PHONE              : 13,
  GENDER             : 14,
  ADDRESS            : 15,
  CITY               : 16,
  STATE              : 17,
  ZIP                : 18,
  DOB                : 19,
  MED_ID             : 20,
  HEIGHT             : 21,
  WEIGHT             : 22,
  SHOE_SIZE          : 23,
  WAIST_SIZE         : 24,
  DOCTOR_NAME        : 25,
  DOCTOR_NPI         : 26,
  DOCTOR_PHONE       : 27,
  DOCTOR_FAX         : 28,
  DOCTOR_ADDRESS     : 29,
  DO_LINK            : 30,
  CN_LINK            : 31,
  RECORD_LINK        : 32,
};


// ─── SOURCE COLUMN INDICES (0-based) ─────────────────────────────────────────

const COL = {
  DATE                      : 0,
  CENTER_CODE               : 1,
  CENTER_NAME               : 2,
  CLOSER_NAME               : 3,
  LEAD_STATUS               : 4,
  CLOSING_NOTES             : 5,
  CHASER_NAME               : 6,
  CHASER_STATUS             : 7,
  CHASER_NOTE               : 8,
  PROCESSING_STATUS_CENTERS : 9,
  PROCESSING_STATUS_ICO     : 10,
  SNS_RESULT                : 11,
  LEAD_TYPE                 : 12,
  REQUESTED_PRODUCTS        : 13,
  FIRST_NAME                : 14,
  LAST_NAME                 : 15,
  PHONE                     : 16,
  ADDRESS                   : 17,
  CITY                      : 18,
  STATE                     : 19,
  ZIP                       : 20,
  DOB                       : 21,
  MED_ID                    : 22,
  HEIGHT                    : 23,
  WEIGHT                    : 24,
  SHOE_SIZE                 : 25,
  WAIST_SIZE                : 26,
  GENDER                    : 27,
  DOCTOR_NAME               : 28,
  DOCTOR_PHONE              : 29,
  DOCTOR_FAX                : 30,
  DOCTOR_NPI                : 31,
};

// ─── DISPOSITION CLASSIFICATION ───────────────────────────────────────────────
// These constants drive color coding and psychological framing across the UI.

const DISPOSITIONS = {
  LEAD_PRODUCTION : ['Verified ppo', 'Verified Med b'],   // Real production
  CHASER_SIGNED   : ['Order Signed'],                     // Ready to process
  CHASER_POTENTIAL: ['Trial', 'Missing CN', 'Missing DO'],// High potential - ICO working on these
  PROC_APPROVED   : ['APPROVED'],                         // Payable to center
  PROC_INPROCESS  : ['IN PROCESS'],                       // Being processed
  PROC_DENIED     : ['DENIED'],                           // Dropped in processing
  PROC_RTS        : ['SHIPPED RTS'],                      // Returned product - monitor closely
};

// RTS thresholds (percentage of total leads)
const RTS_WARNING_PCT = 15;  // Alert if above this
const RTS_TARGET_LOW  = 7;
const RTS_TARGET_HIGH = 15;

// Order signed targets (% of total leads after 45 days)
const ORDER_TARGET_LOW  = 25;
const ORDER_TARGET_HIGH = 30;

// Watched fields for change detection
const WATCHED_FIELDS = [
  { col: COL.LEAD_STATUS,                name: 'Lead Status'               },
  { col: COL.CLOSING_NOTES,             name: 'Closing Notes'             },
  { col: COL.CHASER_NAME,               name: 'Chaser Name'               },
  { col: COL.CHASER_STATUS,             name: 'Chaser Status'             },
  { col: COL.CHASER_NOTE,               name: 'Chaser Note'               },
  { col: COL.PROCESSING_STATUS_CENTERS, name: 'Processing Status'         },
];

// ─── WEB APP ENTRY POINT ──────────────────────────────────────────────────────

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── SPREADSHEET ACCESSORS ───────────────────────────────────────────────────

function getSourceSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SOURCE_SHEET_ID);
}

function getCrmSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.CRM_SHEET_ID);
}

function getCrmSheet(sheetName) {
  const ss    = getCrmSpreadsheet();
  let   sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = _bootstrapSheet(ss, sheetName);
  return sheet;
}

function _bootstrapSheet(ss, sheetName) {
  const HEADERS = {
    [SHEETS.USERS]        : ['Username','Password','Role','Status','Created At'],
    [SHEETS.NOTES_LOG]    : ['Timestamp','Lead ID','Center Code','Note Content','Created By'],
    [SHEETS.AUDIT_LOG]    : ['Timestamp','Lead ID','Field Name','Old Value','New Value','Changed By'],
    [SHEETS.NOTIFICATIONS]: ['Notification ID','Center Code','Lead ID','Message JSON','Status','Timestamp'],
    [SHEETS.ACTIVITY_LOG] : ['Timestamp','User','Action','Details'],
    [SHEETS.LEAD_ID_MAP]  : ['Source Row','Lead ID','Center Code','Watch Hash','Watch Values JSON','Created At'],
    [SHEETS.INCENTIVES]   : ['ID','Prize Label','Weekly Target','Rules','Who Included','Active','Created By','Created At'],
    [SHEETS.CHAT_MESSAGES]: ['ID','Sender','Role','Center Code','Message','Timestamp','Read By Admin','Read By Center'],
    [SHEETS.CHAT_MSGS_V2] : ['ID','Timestamp','Team','ToTeam','Message','ReplyTo','Forwarded','ForwardedAt','Type','Pinned','PinnedAt'], // ← ADD
    [SHEETS.CHAT_RING]    : ['ID','FromTeam','Timestamp','Acknowledged'],                                                               // ← ADD
    [SHEETS.CHAT_STATUS]  : ['Team','Status','UpdatedAt'],      
  };
  const sheet = ss.insertSheet(sheetName);
  const hdrs  = HEADERS[sheetName] || [];
  if (hdrs.length) {
    const hdrRange = sheet.getRange(1, 1, 1, hdrs.length);
    hdrRange.setValues([hdrs])
            .setFontWeight('bold')
            .setBackground('#1e3a5f')
            .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ─── SYSTEM BOOTSTRAP ────────────────────────────────────────────────────────

function initializeSystem() {
  Object.values(SHEETS).forEach(name => getCrmSheet(name));
  const usersSheet = getCrmSheet(SHEETS.USERS);
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.appendRow(['admin', hashPassword('Admin@1234'), 'admin', 'active', now()]);
    Logger.log('Default admin created. Username: admin | Password: Admin@1234');
  }
  return { success: true, message: 'System initialised.' };
}

function getAppConfig() {
  return { success: true, webformUrl: CONFIG.WEBFORM_URL };
}

// ─── PURE UTILITY FUNCTIONS ───────────────────────────────────────────────────

function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password));
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function generateToken() { return Utilities.getUuid(); }

function generateId(prefix) {
  return `${prefix || 'ID'}-${Date.now()}-${Math.random().toString(36).substr(2,6).toUpperCase()}`;
}

function generateLeadId(dateValue) {
  const d  = dateValue instanceof Date ? dateValue : new Date(dateValue || Date.now());
  const ds = isNaN(d) ? _todayStr() : Utilities.formatDate(d, _tz(), 'yyyyMMdd');
  return `LEAD-${ds}-${String(Math.floor(1000 + Math.random() * 9000))}`;
}

function now()       { return new Date().toISOString(); }
function _tz()       { return Session.getScriptTimeZone(); }
function _todayStr() { return Utilities.formatDate(new Date(), _tz(), 'yyyyMMdd'); }

function md5(str) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(str));
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function fmtDate(val, fmt) {
  if (!val) return '';
  try {
    const d = val instanceof Date ? val : new Date(val);
    return isNaN(d) ? String(val) : Utilities.formatDate(d, _tz(), fmt || 'yyyy-MM-dd');
  } catch (_) { return String(val); }
}