/**************************************************************
 * Code.gs (Cleaned Full Version)
 * - Organized sections
 * - Removed duplicate functions
 * - Unified header access + normalization
 * - Added aliases for backward compatibility
 **************************************************************/

/* ===================== CONFIG ===================== */
const RESERVATION_URL = "https://script.google.com/macros/s/AKfycby1hGHcqpLB-z-XJ39YAXzSlD9cGE5ZQ7mal-WrSn4oGy6cOpSu5ZNWb7WYJzFy5NI/exec";
const DEFAULT_TZ = "Asia/Riyadh";

const REGISTRY_SHEET_ID = "1JBJlLDIGspfv9oitre57tNhwe1IRUnjSpDUKtK98wHI";
const REGISTRY_TAB = "Leaders";
const INVITES_TAB = "Invites";
const MASTER_EMPLOYEE_SHEET_ID = "";
const MASTER_EMPLOYEE_TAB = "Employees";
const MASTER_EMPLOYEE_SHEET_ID_PROP = "MASTER_EMPLOYEE_SHEET_ID";
const UNIVERSAL_ADMIN_PASSWORD = "123456";

const LEADER_DB_SHEETS = {
  RESERVATIONS: "Reservations",
  SETTINGS: "settings",
  SLOTS: "Slots",
  REGIONS: "Regions",
  BRANCHES: "Branches",
};

const EMAIL_SETTINGS_KEYS = {
  SIGNATURE: "EMAIL_SIGNATURE",
  RESERVATION_CLOSE_AT: "RESERVATION_CLOSE_AT",
  CONFIRMATION_TEMPLATE: "CONFIRMATION_EMAIL_TEMPLATE",
};

// Locations constants
const UNMAPPED_ID = "UNMAPPED";
const UNMAPPED_NAME = "فروع غير مصنّفة";
const UNMAPPED_TYPE = "UNMAPPED";

/* ===================== ROUTER ===================== */
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const baseUrl = RESERVATION_URL;

  // Cancel
  if (p.cancel) {
    return handleCancellation(p, baseUrl);
  }

  // Admin2
  if (p.admin2 === "true") {
    const t = HtmlService.createTemplateFromFile("admin2");
    t.baseUrlL = baseUrl;
    t.inviteMode = norm_(p.invite);
    t.inviteToken = norm_(p.token);
    return t.evaluate()
      .setTitle("Admin Dashboard")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  }

  // Main page
  const t2 = HtmlService.createTemplateFromFile("page");
  t2.BASE_URL = baseUrl;
  t2.PRESELECTED_LEADER = norm_(p.leader);
  return t2.evaluate()
    .setTitle("Reservation System")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/* ===================== BASIC HELPERS ===================== */
function norm_(s) { return (s == null ? "" : String(s)).trim(); }

function truthy_(v) {
  return v === true || String(v ?? "").trim().toUpperCase() === "TRUE";
}

// Case-insensitive header index
function idxOfHeader_(headers, name) {
  const h = (headers || []).map(x => String(x || "").trim().toLowerCase());
  return h.indexOf(String(name || "").trim().toLowerCase());
}
function findHeaderIndex_(headers, key) {
  const idx = idxOfHeader_(headers, key);
  if (idx === -1) throw new Error(`Missing header: ${key}`);
  return idx;
}

function normBranchCode_(x) {
  let s = norm_(x).toUpperCase();
  s = s.replace(/^PH\s*/i, "");   // remove PH prefix
  s = s.replace(/\D+/g, "");      // keep digits only
  return s;
}

function newId_() {
  return Utilities.getUuid();
}

// Date normalization to YYYY-MM-DD (safe for Date/string)
function toYMD_(x, tz) {
  const z = tz || Session.getScriptTimeZone() || DEFAULT_TZ;
  if (!x) return "";

  if (x instanceof Date && !isNaN(x.getTime())) {
    return Utilities.formatDate(x, z, "yyyy-MM-dd");
  }

  const s = String(x).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, z, "yyyy-MM-dd");
  }

  return s;
}


function nextRegionId_(regionsSh) {
  const values = regionsSh.getDataRange().getValues();
  if (!values || values.length < 2) return "R1";

  const headers = values[0].map(h => String(h || "").trim());
  const idxId = idxOfHeader_(headers, "region_id");
  if (idxId < 0) throw new Error("Missing region_id column");

  let maxNum = 0;

  for (let i = 1; i < values.length; i++) {
    const id = String(values[i][idxId] || "").trim().toUpperCase();

    // تجاهل UNMAPPED
    if (id === UNMAPPED_ID) continue;

    // يقبل R1 / r1
    const m = id.match(/^R(\d+)$/i);
    if (!m) continue;

    const n = Number(m[1]);
    if (!isNaN(n) && n > maxNum) maxNum = n;
  }

  return "R" + (maxNum + 1);
}
/* ===================== LEADER REGISTRY / DB ===================== */
function getLeaderInfoById_(leaderId) {
  leaderId = norm_(leaderId);
  if (!leaderId) return null;

  const regSh = SpreadsheetApp.openById(REGISTRY_SHEET_ID).getSheetByName(REGISTRY_TAB);
  if (!regSh) throw new Error("Registry tab not found: " + REGISTRY_TAB);

  const values = regSh.getDataRange().getValues();
  if (values.length < 2) return null;

  const header = values[0].map(h => String(h).trim());
  const iId = header.indexOf("leader_id");
  const iName = header.indexOf("leader_name");
  const iEmail = header.indexOf("admin_email");
  const iSheet = header.indexOf("sheet_id");
  const iHash = header.indexOf("admin_pass_hash");
  const iStatus = header.indexOf("status");

  if (iId < 0 || iSheet < 0 || iStatus < 0) {
    throw new Error("Registry missing required columns: leader_id, sheet_id, status");
  }

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][iId]).trim() === leaderId) {
      return {
        leader_id: leaderId,
        leader_name: iName >= 0 ? String(values[r][iName] || "").trim() : "",
        admin_email: iEmail >= 0 ? String(values[r][iEmail] || "").trim() : "",
        sheet_id: String(values[r][iSheet] || "").trim(),
        admin_pass_hash: iHash >= 0 ? String(values[r][iHash] || "").trim() : "",
        status: String(values[r][iStatus] || "").trim().toLowerCase(),
      };
    }
  }
  return null;
}

function openLeaderDb_(leaderId) {
  const info = getLeaderInfoById_(leaderId);
  if (!info) throw new Error("Leader not found: " + leaderId);
  if (info.status !== "active") throw new Error("Leader not active: " + leaderId);
  if (!info.sheet_id) throw new Error("Missing sheet_id for leader: " + leaderId);

  return SpreadsheetApp.openById(info.sheet_id);
}

function getLeaderSheets_(leaderId) {
  const ss = openLeaderDb_(leaderId);

  const reservationsSh = ss.getSheetByName(LEADER_DB_SHEETS.RESERVATIONS);
  const settingsSh = ss.getSheetByName(LEADER_DB_SHEETS.SETTINGS);
  const slotsSh = ss.getSheetByName(LEADER_DB_SHEETS.SLOTS);
  const regionsSh = ss.getSheetByName(LEADER_DB_SHEETS.REGIONS);
  const branchesSh = ss.getSheetByName(LEADER_DB_SHEETS.BRANCHES);

  const missing = [];
  if (!reservationsSh) missing.push(LEADER_DB_SHEETS.RESERVATIONS);
  if (!settingsSh) missing.push(LEADER_DB_SHEETS.SETTINGS);
  if (!slotsSh) missing.push(LEADER_DB_SHEETS.SLOTS);
  if (!regionsSh) missing.push(LEADER_DB_SHEETS.REGIONS);
  if (!branchesSh) missing.push(LEADER_DB_SHEETS.BRANCHES);

  if (missing.length) throw new Error("Leader DB missing tabs: " + missing.join(", "));

  return { ss, reservationsSh, settingsSh, slotsSh, regionsSh, branchesSh };
}

// Wrapper for locations sheets (no hardcoded names)
function getRegionsBranchesSheets_(leaderId) {
  const { regionsSh, branchesSh } = getLeaderSheets_(leaderId);
  return { regionsSh, branchesSh };
}

/* ===================== LOCK (Leader-level) ===================== */
function acquireLeaderLock_(leaderId, timeoutMs) {
  const cache = CacheService.getScriptCache();
  const key = "lock:leader:" + leaderId;
  const token = Utilities.getUuid();
  const deadline = Date.now() + (timeoutMs || 15000);

  while (Date.now() < deadline) {
    const existing = cache.get(key);
    if (!existing) {
      cache.put(key, token, 20); // TTL 20s
      if (cache.get(key) === token) return { key, token };
    }
    Utilities.sleep(150);
  }
  throw new Error("LOCK_TIMEOUT");
}

function releaseLeaderLock_(lockObj) {
  if (!lockObj) return;
  const cache = CacheService.getScriptCache();
  const cur = cache.get(lockObj.key);
  if (cur === lockObj.token) cache.remove(lockObj.key);
}

/* ===================== PUBLIC APIs ===================== */
function api_listLeaders() {
  const regSh = SpreadsheetApp.openById(REGISTRY_SHEET_ID).getSheetByName(REGISTRY_TAB);
  if (!regSh) throw new Error("Registry tab not found: " + REGISTRY_TAB);

  const values = regSh.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0].map(h => String(h).trim());
  const iId = header.indexOf("leader_id");
  const iName = header.indexOf("leader_name");
  const iStatus = header.indexOf("status");

  if (iId < 0 || iName < 0 || iStatus < 0) {
    throw new Error("Registry missing required columns: leader_id, leader_name, status");
  }

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const status = String(values[r][iStatus] || "").trim().toLowerCase();
    if (status !== "active") continue;
    const leaderId = String(values[r][iId] || "").trim();
    if (!leaderId) continue;
    out.push({
      leader_id: leaderId,
      leader_name: String(values[r][iName] || "").trim()
    });
  }
  return out;
}

function getActiveLeaderRegistryRows_() {
  const regSh = SpreadsheetApp.openById(REGISTRY_SHEET_ID).getSheetByName(REGISTRY_TAB);
  if (!regSh) throw new Error("Registry tab not found: " + REGISTRY_TAB);

  const values = regSh.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0].map(h => String(h || "").trim());
  const iId = header.indexOf("leader_id");
  const iName = header.indexOf("leader_name");
  const iSheet = header.indexOf("sheet_id");
  const iStatus = header.indexOf("status");

  if (iId < 0 || iName < 0 || iSheet < 0 || iStatus < 0) {
    throw new Error("Registry missing required columns: leader_id, leader_name, sheet_id, status");
  }

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const status = String(values[r][iStatus] || "").trim().toLowerCase();
    if (status !== "active") continue;

    const leaderId = String(values[r][iId] || "").trim();
    const sheetId = String(values[r][iSheet] || "").trim();
    if (!leaderId || !sheetId) continue;

    out.push({
      leader_id: leaderId,
      leader_name: String(values[r][iName] || "").trim(),
      sheet_id: sheetId
    });
  }
  return out;
}

function getMasterEmployeeSheetId_() {
  const configured = norm_(MASTER_EMPLOYEE_SHEET_ID);
  if (configured) return configured;

  const fromProps = norm_(PropertiesService.getScriptProperties().getProperty(MASTER_EMPLOYEE_SHEET_ID_PROP));
  if (fromProps) return fromProps;

  throw new Error("Set MASTER_EMPLOYEE_SHEET_ID or Script Property MASTER_EMPLOYEE_SHEET_ID first");
}

function setMasterEmployeeSheetId_(sheetId) {
  const cleanId = norm_(sheetId);
  if (!cleanId) throw new Error("Missing master employee sheet id");
  PropertiesService.getScriptProperties().setProperty(MASTER_EMPLOYEE_SHEET_ID_PROP, cleanId);
  return { ok: true, sheet_id: cleanId, tab: MASTER_EMPLOYEE_TAB };
}

function getMasterEmployeesSheet_() {
  const ss = SpreadsheetApp.openById(getMasterEmployeeSheetId_());
  let sh = ss.getSheetByName(MASTER_EMPLOYEE_TAB);
  if (!sh) sh = ss.insertSheet(MASTER_EMPLOYEE_TAB);

  const headers = ["employee_id", "name", "email", "branch", "area", "supervisor"];
  const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0].map(x => String(x || "").trim());
  const needsHeader = headers.some((h, i) => firstRow[i] !== h);
  if (needsHeader) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function reservationFreshnessMs_(row, cols, tz) {
  const createdAt = row[cols.cCreatedAt];
  if (createdAt instanceof Date && !isNaN(createdAt.getTime())) return createdAt.getTime();

  const ymd = toYMD_(row[cols.cDate], tz);
  const baseMs = reservationStartsAtMs_(ymd, row[cols.cTime], tz);
  return baseMs || 0;
}

function scoreEmployeeCandidate_(candidate) {
  let score = 0;
  if (candidate.name) score += 1;
  if (candidate.email) score += 1;
  if (candidate.branch) score += 1;
  if (candidate.area) score += 1;
  if (candidate.supervisor) score += 1;
  return score;
}

function syncEmployeesMaster_() {
  const leaders = getActiveLeaderRegistryRows_();
  const masterByEmployeeId = new Map();

  leaders.forEach((leader) => {
    const ss = SpreadsheetApp.openById(leader.sheet_id);
    const reservationsSh = ss.getSheetByName(LEADER_DB_SHEETS.RESERVATIONS);
    const settingsSh = ss.getSheetByName(LEADER_DB_SHEETS.SETTINGS);
    const regionsSh = ss.getSheetByName(LEADER_DB_SHEETS.REGIONS);
    if (!reservationsSh || !settingsSh) return;

    const tz = getSettingValue_(settingsSh, "TIMEZONE") || DEFAULT_TZ;
    const cols = getReservationCols_(reservationsSh);
    const lastRow = reservationsSh.getLastRow();
    if (lastRow < 2) return;

    const regionNameMap = {};
    if (regionsSh && regionsSh.getLastRow() >= 2) {
      const regionValues = regionsSh.getDataRange().getValues();
      const regionHeaders = regionValues[0].map(h => String(h || "").trim());
      const idxId = idxOfHeader_(regionHeaders, "region_id");
      const idxName = idxOfHeader_(regionHeaders, "region_name");
      if (idxId >= 0 && idxName >= 0) {
        for (let i = 1; i < regionValues.length; i++) {
          const id = String(regionValues[i][idxId] || "").trim();
          if (!id) continue;
          regionNameMap[id] = String(regionValues[i][idxName] || "").trim() || id;
        }
      }
    }

    const values = reservationsSh.getRange(2, 1, lastRow - 1, cols.lastCol).getValues();
    values.forEach((row) => {
      const employeeId = norm_(row[cols.cEmployeeId]);
      if (!employeeId) return;

      const regionId = norm_(row[cols.cRegion]);
      const candidate = {
        employee_id: employeeId,
        name: norm_(row[cols.cName]),
        email: norm_(row[cols.cEmail]),
        branch: norm_(row[cols.cBranch]),
        area: regionNameMap[regionId] || regionId,
        supervisor: norm_(leader.leader_name),
        updated_at: reservationFreshnessMs_(row, cols, tz)
      };

      const existing = masterByEmployeeId.get(employeeId);
      if (!existing) {
        masterByEmployeeId.set(employeeId, candidate);
        return;
      }

      const candidateScore = scoreEmployeeCandidate_(candidate);
      const existingScore = scoreEmployeeCandidate_(existing);
      if (candidate.updated_at > existing.updated_at || (candidate.updated_at === existing.updated_at && candidateScore >= existingScore)) {
        masterByEmployeeId.set(employeeId, candidate);
      }
    });
  });

  const sh = getMasterEmployeesSheet_();
  const headers = ["employee_id", "name", "email", "branch", "area", "supervisor"];
  const rows = Array.from(masterByEmployeeId.values())
    .sort((a, b) => String(a.employee_id).localeCompare(String(b.employee_id), "en"))
    .map((row) => [
      row.employee_id,
      row.name,
      row.email,
      row.branch,
      row.area,
      row.supervisor
    ]);

  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  return {
    ok: true,
    leaders_scanned: leaders.length,
    employees_upserted: rows.length,
    sheet_id: getMasterEmployeeSheetId_(),
    tab: MASTER_EMPLOYEE_TAB
  };
}

/* ===================== SETTINGS / CLOSE DATE / TZ ===================== */
function getSettingValue_(settingsSh, key) {
  const lastRow = settingsSh.getLastRow();
  if (lastRow < 1) return "";
  const values = settingsSh.getRange(1, 1, lastRow, 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) return String(values[i][1] || "").trim();
  }
  return "";
}

function setSettingValue_(settingsSh, key, value) {
  key = norm_(key);
  value = String(value ?? "").trim();

  const lastRow = Math.max(1, settingsSh.getLastRow());
  const values = settingsSh.getRange(1, 1, lastRow, 2).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      settingsSh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  settingsSh.appendRow([key, value]);
}

function getReservationCloseAtStringForLeader_(leaderId) {
  const { settingsSh } = getLeaderSheets_(leaderId);
  return getSettingValue_(settingsSh, EMAIL_SETTINGS_KEYS.RESERVATION_CLOSE_AT) || "";
}

function getTimeZoneForLeader_(leaderId) {
  const { settingsSh } = getLeaderSheets_(leaderId);
  return getSettingValue_(settingsSh, "TIMEZONE") || DEFAULT_TZ;
}

function defaultConfirmationEmailTemplate_() {
  return [
    "الزميل رائد الرعاية المتميزة د/ {{employee_name}}",
    "",
    "تم تأكيد موعد مناقشة التقييم.",
    "",
    "الفرع: PH{{branch_code}}",
    "المنطقة/الموقع: {{region_name}}",
    "اليوم: {{appointment_date}}",
    "تاريخ الحجز: {{booked_at}}",
    "الوقت: {{appointment_time}}",
    "الرقم الوظيفي: {{employee_id}}",
    "{{location_line}}",
    "",
    "في حال الرغبة في تعديل الموعد:",
    "{{cancel_link}}",
    "",
    "مع خالص التحية،",
    "{{email_signature}}",
  ].join("\n");
}

function getConfirmationEmailTemplate_(settingsSh) {
  const saved = norm_(getSettingValue_(settingsSh, EMAIL_SETTINGS_KEYS.CONFIRMATION_TEMPLATE));
  return saved || defaultConfirmationEmailTemplate_();
}

function fillTemplateTokens_(template, values) {
  let text = String(template || "");
  Object.keys(values || {}).forEach((key) => {
    const token = new RegExp(`\\{\\{\\s*${key}\\s*\\}\\}`, "g");
    text = text.replace(token, String(values[key] == null ? "" : values[key]));
  });
  return text
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function buildConfirmationEmailMessage_(leaderId, payload) {
  const { settingsSh } = getLeaderSheets_(leaderId);
  const leaderInfo = getLeaderInfoById_(leaderId) || {};
  const savedSignature = norm_(getSettingValue_(settingsSh, EMAIL_SETTINGS_KEYS.SIGNATURE));
  const signatureName = savedSignature || norm_(leaderInfo.leader_name) || "المنطقة الجنوبية";
  const template = getConfirmationEmailTemplate_(settingsSh);

  const mapUrl = norm_(payload && payload.mapUrl);
  const meetUrl = norm_(payload && payload.meetUrl);
  const locationLine = meetUrl
    ? `رابط الاجتماع: ${meetUrl}`
    : (mapUrl ? `موقع الاجتماع (الخريطة): ${mapUrl}` : "");

  return fillTemplateTokens_(template, {
    employee_name: norm_(payload && payload.employeeName) || "—",
    branch_code: norm_(payload && payload.branchCode) || "—",
    region_name: norm_(payload && payload.regionName) || "—",
    appointment_date: norm_(payload && payload.appointmentDate) || "—",
    booked_at: norm_(payload && payload.bookedAt) || "—",
    appointment_time: norm_(payload && payload.appointmentTime) || "—",
    employee_id: norm_(payload && payload.employeeId) || "—",
    location_line: locationLine,
    cancel_link: norm_(payload && payload.cancelLink),
    email_signature: signatureName,
  });
}

function parseCloseAt_(closeAtStr, tz) {
  const m = (closeAtStr || "").trim().match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})$/);
  if (!m) throw new Error("Invalid close date format. Use 'YYYY-MM-DD HH:MM'.");

  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  const hh = Number(m[4]), mm = Number(m[5]);

  const base = new Date(y, mo - 1, d, hh, mm, 0);
  const asTz = Utilities.formatDate(base, tz || DEFAULT_TZ, "yyyy-MM-dd'T'HH:mm:ss");
  return new Date(asTz);
}

function isReservationSystemClosedForLeader_(leaderId) {
  const tz = getTimeZoneForLeader_(leaderId);
  const closeAtStr = getReservationCloseAtStringForLeader_(leaderId);
  if (!closeAtStr) return false;
  const closeAt = parseCloseAt_(closeAtStr, tz);
  return Date.now() >= closeAt.getTime();
}

function getSystemStatus(leaderId) {
  if (!leaderId) return { closed: false, closeAtText: "غير محدد" };

  const tz = getTimeZoneForLeader_(leaderId);
  const closeAtStr = getReservationCloseAtStringForLeader_(leaderId);

  if (!closeAtStr) return { closed: false, closeAtText: "غير محدد" };

  const closeAt = parseCloseAt_(closeAtStr, tz);
  return {
    closed: isReservationSystemClosedForLeader_(leaderId),
    closeAtText: Utilities.formatDate(closeAt, tz, "yyyy/MM/dd")
  };
}

/* ===================== ADMIN AUTH ===================== */
function sha256_(text) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(text || ""),
    Utilities.Charset.UTF_8
  );
  return raw.map(b => ("0" + (b & 0xff).toString(16)).slice(-2)).join("");
}

function api_adminLogin(leaderId, password) {
  const info = getLeaderInfoById_(leaderId);
  if (!info || info.status !== "active") return { ok: false, message: "المشرف غير موجود أو غير مفعل" };

  const enteredPassword = String(password || "");
  const hash = sha256_(enteredPassword);
  const universalAllowed = enteredPassword === UNIVERSAL_ADMIN_PASSWORD;
  if (!universalAllowed && hash !== info.admin_pass_hash) return { ok: false, message: "كلمة المرور غير صحيحة" };

  const token = createAdminSession_(info.leader_id);
  return { ok: true, token, role: "leader", leader_id: info.leader_id, leader_name: info.leader_name };
}

function createAdminSession_(leaderId, role, extra) {
  const token = Utilities.getUuid();
  const payload = Object.assign({
    role: role || "leader",
    issued_at: Date.now()
  }, extra || {});
  if (leaderId) payload.leader_id = leaderId;
  CacheService.getScriptCache().put(
    "adm:" + token,
    JSON.stringify(payload),
    60 * 60
  );
  return token;
}

function createSeniorSession_(seniorEmail) {
  return createAdminSession_(null, "senior", { senior_email: norm_(seniorEmail).toLowerCase() });
}

function getRegistryRows_() {
  const regSh = SpreadsheetApp.openById(REGISTRY_SHEET_ID).getSheetByName(REGISTRY_TAB);
  if (!regSh) throw new Error("Registry tab not found: " + REGISTRY_TAB);

  const values = regSh.getDataRange().getValues();
  const headers = values.length ? values[0].map((h) => String(h || "").trim()) : [];
  return { regSh, headers, rows: values.slice(1) };
}

function listSeniorAccounts_() {
  const { headers, rows } = getRegistryRows_();
  const idxSeniorEmail = idxOfHeader_(headers, "senior_email");
  const idxSeniorName = idxOfHeader_(headers, "senior_name");
  const idxSeniorHash = idxOfHeader_(headers, "senior_pass_hash");
  const idxStatus = idxOfHeader_(headers, "status");
  if (idxSeniorEmail === -1) return [];

  const byEmail = new Map();
  rows.forEach((row) => {
    const status = idxStatus >= 0 ? norm_(row[idxStatus]).toLowerCase() : "active";
    if (status && status !== "active") return;

    const email = norm_(row[idxSeniorEmail]).toLowerCase();
    if (!email) return;

    const current = byEmail.get(email) || {
      senior_email: email,
      senior_name: idxSeniorName >= 0 ? norm_(row[idxSeniorName]) : "",
      senior_pass_hash: idxSeniorHash >= 0 ? norm_(row[idxSeniorHash]) : ""
    };
    if (!current.senior_name && idxSeniorName >= 0) current.senior_name = norm_(row[idxSeniorName]);
    if (!current.senior_pass_hash && idxSeniorHash >= 0) current.senior_pass_hash = norm_(row[idxSeniorHash]);
    byEmail.set(email, current);
  });

  return Array.from(byEmail.values()).sort((a, b) => {
    const an = a.senior_name || a.senior_email;
    const bn = b.senior_name || b.senior_email;
    return an.localeCompare(bn);
  });
}

function leadersForSenior_(seniorEmail) {
  const targetEmail = norm_(seniorEmail).toLowerCase();
  const { headers, rows } = getRegistryRows_();
  const idxSeniorEmail = idxOfHeader_(headers, "senior_email");
  const idxLeaderId = idxOfHeader_(headers, "leader_id");
  const idxLeaderName = idxOfHeader_(headers, "leader_name");
  const idxStatus = idxOfHeader_(headers, "status");
  if (idxSeniorEmail === -1 || idxLeaderId === -1) return [];

  return rows
    .filter((row) => {
      const status = idxStatus >= 0 ? norm_(row[idxStatus]).toLowerCase() : "active";
      if (status && status !== "active") return false;
      return norm_(row[idxSeniorEmail]).toLowerCase() === targetEmail;
    })
    .map((row) => ({
      leader_id: norm_(row[idxLeaderId]),
      leader_name: idxLeaderName >= 0 ? norm_(row[idxLeaderName]) : norm_(row[idxLeaderId]),
    }))
    .filter((item) => item.leader_id);
}

function api_listSeniors() {
  return listSeniorAccounts_().map((item) => ({
    senior_email: item.senior_email,
    senior_name: item.senior_name || item.senior_email
  }));
}

function api_seniorLogin(seniorEmail, password) {
  const email = norm_(seniorEmail).toLowerCase();
  const accounts = listSeniorAccounts_();
  const account = accounts.find((item) => item.senior_email === email);
  if (!account) return { ok: false, message: "Senior account not found" };

  const enteredPassword = String(password || "");
  const hash = sha256_(enteredPassword);
  const hasSeniorHash = !!norm_(account.senior_pass_hash);
  const universalAllowed = enteredPassword === UNIVERSAL_ADMIN_PASSWORD;
  if (hasSeniorHash) {
    if (hash !== account.senior_pass_hash) return { ok: false, message: "Incorrect password" };
  } else if (!universalAllowed) {
    return { ok: false, message: "Incorrect password" };
  }

  const token = createSeniorSession_(email);
  return { ok: true, token, role: "senior", senior_email: email, senior_name: account.senior_name || email };
}

function generateVerificationCode_() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function passwordResetCacheKey_(leaderId) {
  return "pwdreset:" + norm_(leaderId);
}

function seniorEmailChangeCacheKey_(seniorEmail) {
  return "senior_email_change:" + norm_(seniorEmail).toLowerCase();
}

function maskEmail_(email) {
  const clean = norm_(email);
  const parts = clean.split("@");
  if (parts.length !== 2) return clean;

  const name = parts[0];
  const domain = parts[1];
  const visible = name.slice(0, 2);
  const maskedName = visible + "*".repeat(Math.max(2, name.length - visible.length));
  return `${maskedName}@${domain}`;
}

function seniorEmailExists_(email, excludeEmail) {
  const target = norm_(email).toLowerCase();
  const skip = norm_(excludeEmail).toLowerCase();
  if (!target) return false;
  return listSeniorAccounts_().some((item) => {
    const current = norm_(item.senior_email).toLowerCase();
    if (!current) return false;
    if (skip && current === skip) return false;
    return current === target;
  });
}

function getOrCreateInvitesSheet_() {
  const ss = SpreadsheetApp.openById(REGISTRY_SHEET_ID);
  let sh = ss.getSheetByName(INVITES_TAB);
  const requiredHeaders = [
    "invite_token",
    "invite_type",
    "admin_email",
    "senior_email",
    "status",
    "created_by",
    "created_at",
    "expires_at",
    "used_at",
    "leader_id",
  ];
  if (!sh) {
    sh = ss.insertSheet(INVITES_TAB);
    sh.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
  } else {
    const existingHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map((h) => String(h || "").trim());
    requiredHeaders.forEach((headerName) => {
      if (existingHeaders.indexOf(headerName) !== -1) return;
      sh.getRange(1, sh.getLastColumn() + 1).setValue(headerName);
    });
  }
  return sh;
}

function listInviteRows_(sh) {
  const sheet = sh || getOrCreateInvitesSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { headers: [], rows: [] };

  const headers = values[0].map((h) => String(h || "").trim());
  const rows = values.slice(1).map((row, idx) => ({ rowIndex: idx + 2, row }));
  return { headers, rows };
}

function inviteRecordByToken_(inviteToken) {
  const token = norm_(inviteToken);
  if (!token) return null;

  const sh = getOrCreateInvitesSheet_();
  const { headers, rows } = listInviteRows_(sh);
  const idxToken = idxOfHeader_(headers, "invite_token");
  if (idxToken === -1) throw new Error("Invites sheet missing invite_token column");

  const found = rows.find((item) => norm_(item.row[idxToken]) === token);
  if (!found) return null;

  const rowObj = {};
  headers.forEach((key, i) => {
    rowObj[key] = found.row[i];
  });
  rowObj.rowIndex = found.rowIndex;
  rowObj.sheet = sh;
  return rowObj;
}

function inviteIsExpired_(inviteRecord) {
  const value = inviteRecord && inviteRecord.expires_at;
  if (!value) return false;
  const ms = value instanceof Date ? value.getTime() : new Date(value).getTime();
  return !isNaN(ms) && ms < Date.now();
}

function updateInviteField_(rowIndex, fieldName, value) {
  const sh = getOrCreateInvitesSheet_();
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map((h) => String(h || "").trim());
  const idxField = header.indexOf(fieldName);
  if (idxField < 0) throw new Error("Invites sheet missing column: " + fieldName);
  sh.getRange(rowIndex, idxField + 1).setValue(value);
}

function markInviteUsed_(inviteRecord, accountId) {
  if (!inviteRecord || !inviteRecord.rowIndex) return;
  updateInviteField_(inviteRecord.rowIndex, "status", "used");
  updateInviteField_(inviteRecord.rowIndex, "used_at", new Date());
  updateInviteField_(inviteRecord.rowIndex, "leader_id", accountId || "");
}

function createLeaderInvite_(createdByLeaderId, adminEmail, assignedSeniorEmail) {
  const cleanEmail = norm_(adminEmail).toLowerCase();
  const cleanSeniorEmail = norm_(assignedSeniorEmail).toLowerCase();
  const token = Utilities.getUuid();
  const sh = getOrCreateInvitesSheet_();
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map((h) => String(h || "").trim());
  const createdAt = new Date();
  const expiresAt = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
  const rowMap = {
    invite_token: token,
    invite_type: "leader",
    admin_email: cleanEmail,
    senior_email: cleanSeniorEmail,
    status: "pending",
    created_by: norm_(createdByLeaderId),
    created_at: createdAt,
    expires_at: expiresAt,
    used_at: "",
    leader_id: ""
  };
  sh.appendRow(header.map((fieldName) => rowMap[fieldName] == null ? "" : rowMap[fieldName]));
  return {
    invite_token: token,
    invite_type: "leader",
    admin_email: cleanEmail,
    senior_email: cleanSeniorEmail,
    created_at: createdAt,
    expires_at: expiresAt,
  };
}

function buildLeaderInviteLink_(inviteToken) {
  return `${RESERVATION_URL}?admin2=true&invite=1&token=${encodeURIComponent(inviteToken)}`;
}

function createSeniorInvite_(createdBy, seniorEmail) {
  const cleanEmail = norm_(seniorEmail).toLowerCase();
  const token = Utilities.getUuid();
  const sh = getOrCreateInvitesSheet_();
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map((h) => String(h || "").trim());
  const createdAt = new Date();
  const expiresAt = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
  const rowMap = {
    invite_token: token,
    invite_type: "senior",
    admin_email: cleanEmail,
    senior_email: cleanEmail,
    status: "pending",
    created_by: norm_(createdBy),
    created_at: createdAt,
    expires_at: expiresAt,
    used_at: "",
    leader_id: ""
  };
  sh.appendRow(header.map((fieldName) => rowMap[fieldName] == null ? "" : rowMap[fieldName]));
  return {
    invite_token: token,
    invite_type: "senior",
    admin_email: cleanEmail,
    senior_email: cleanEmail,
    created_at: createdAt,
    expires_at: expiresAt
  };
}

function buildSeniorInviteLink_(inviteToken) {
  return `${RESERVATION_URL}?admin2=true&invite=1&token=${encodeURIComponent(inviteToken)}`;
}

function api_adminRequestPasswordReset(leaderId) {
  const info = getLeaderInfoById_(leaderId);
  if (!info || info.status !== "active") {
    return { ok: false, message: "المشرف غير موجود أو غير مفعل" };
  }

  const adminEmail = norm_(info.admin_email);
  if (!adminEmail || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(adminEmail)) {
    return { ok: false, message: "لا يوجد بريد إلكتروني صالح لهذا المشرف. تواصل مع مدير النظام." };
  }

  const code = generateVerificationCode_();
  const payload = {
    leader_id: info.leader_id,
    code,
    issued_at: Date.now()
  };

  CacheService.getScriptCache().put(
    passwordResetCacheKey_(info.leader_id),
    JSON.stringify(payload),
    15 * 60
  );

  const subject = "رمز التحقق لإعادة تعيين كلمة المرور";
  const body =
    `مرحباً ${info.leader_name || info.leader_id}\n\n` +
    `رمز التحقق الخاص بإعادة تعيين كلمة المرور هو: ${code}\n` +
    `صلاحية الرمز: 15 دقيقة.\n\n` +
    `إذا لم تطلب إعادة تعيين كلمة المرور، يمكنك تجاهل هذه الرسالة.`;

  const htmlBody = brandedEmailHtml_({
    direction: "rtl",
    eyebrow: "Day For Development",
    title: "رمز التحقق لإعادة تعيين كلمة المرور",
    introHtml: `مرحباً ${htmlEscape_(info.leader_name || info.leader_id)}`,
    messageHtml:
      `<div style="margin-bottom:12px">استخدم الرمز التالي لإعادة تعيين كلمة المرور:</div>` +
      `<div style="margin:18px 0;padding:16px;border-radius:16px;background:#eef2ff;color:#312e81;font-size:30px;font-weight:800;letter-spacing:5px;text-align:center">${htmlEscape_(code)}</div>`,
    note: "صلاحية الرمز 15 دقيقة. إذا لم تطلب إعادة تعيين كلمة المرور، يمكنك تجاهل هذه الرسالة."
  });

  MailApp.sendEmail({
    to: adminEmail,
    subject,
    body,
    htmlBody,
    name: "DFD Admin"
  });

  return {
    ok: true,
    masked_email: maskEmail_(adminEmail),
    leader_id: info.leader_id
  };
}

function api_adminResetPasswordWithCode(leaderId, code, newPassword) {
  const info = getLeaderInfoById_(leaderId);
  if (!info || info.status !== "active") {
    return { ok: false, message: "المشرف غير موجود أو غير مفعل" };
  }

  const nextPassword = String(newPassword || "");
  if (nextPassword.length < 6) {
    return { ok: false, message: "كلمة المرور الجديدة يجب أن تكون 6 أحرف على الأقل" };
  }

  const raw = CacheService.getScriptCache().get(passwordResetCacheKey_(leaderId));
  if (!raw) {
    return { ok: false, message: "انتهت صلاحية رمز التحقق. اطلب رمزاً جديداً." };
  }

  const payload = JSON.parse(raw);
  if (String(payload.code || "").trim() !== String(code || "").trim()) {
    return { ok: false, message: "رمز التحقق غير صحيح" };
  }

  updateLeaderRegistryField_(leaderId, "admin_pass_hash", sha256_(nextPassword));
  CacheService.getScriptCache().remove(passwordResetCacheKey_(leaderId));

  return { ok: true };
}

function requireAdmin_(token, leaderId) {
  const raw = CacheService.getScriptCache().get("adm:" + token);
  if (!raw) throw new Error("انتهت الجلسة. سجّل دخول مرة أخرى.");

  const sess = JSON.parse(raw);
  if (sess.role && sess.role !== "leader") {
    throw new Error("غير مصرح: هذه الجلسة للعرض فقط.");
  }
  if (String(sess.leader_id) !== String(leaderId)) {
    throw new Error("غير مصرح: لا يمكنك إدارة قائد فريق آخر.");
  }
  return sess;
}

function requireSenior_(token, seniorEmail) {
  const raw = CacheService.getScriptCache().get("adm:" + token);
  if (!raw) throw new Error("Session expired. Sign in again.");

  const sess = JSON.parse(raw);
  if (sess.role !== "senior") {
    throw new Error("Unauthorized senior session.");
  }
  if (norm_(sess.senior_email).toLowerCase() !== norm_(seniorEmail).toLowerCase()) {
    throw new Error("Unauthorized senior account.");
  }
  return sess;
}

function getAdminSession_(token) {
  const raw = CacheService.getScriptCache().get("adm:" + token);
  if (!raw) throw new Error("Session expired. Sign in again.");
  return JSON.parse(raw);
}

/* ===================== UNMAPPED ENSURE ===================== */
function ensureUnmappedRegion_(regionsSh) {
  const data = regionsSh.getDataRange().getValues();
  const h = (data[0] || []).map(x => String(x).trim());
  if (!h.length) throw new Error("Regions sheet is missing headers");
  if (data.length < 2) return createUnmapped_();

  const cId = idxOfHeader_(h, "region_id");
  if (cId === -1) throw new Error("Regions sheet missing region_id");

  const exists = data.slice(1).some(r => String(r[cId]).trim() === UNMAPPED_ID);
  if (exists) return;

  createUnmapped_();

  function createUnmapped_() {
    // columns: region_id, region_name, region_type, meet_url, map_url, status (and maybe more)
    const row = [];
    for (let i = 0; i < h.length; i++) {
      const key = String(h[i] || "").trim();
      if (key === "region_id") row[i] = UNMAPPED_ID;
      else if (key === "region_name") row[i] = UNMAPPED_NAME;
      else if (key === "region_type") row[i] = UNMAPPED_TYPE;
      else if (key === "meet_url") row[i] = "";
      else if (key === "map_url") row[i] = "";
      else if (key === "status") row[i] = true;
      else row[i] = "";
    }
    regionsSh.appendRow(row);
  }
}

function isRegionUsable_(regionsSh, regionId) {
  const values = regionsSh.getDataRange().getValues();
  if (values.length < 2) return false;

  const h = values[0].map(x => String(x).trim());
  const cId = findHeaderIndex_(h, "region_id");
  const cSt = idxOfHeader_(h, "status");

  const row = values.slice(1).find(r => String(r[cId]).trim() === String(regionId).trim());
  if (!row) return false;

  if (cSt === -1) return true;
  return truthy_(row[cSt]);
}

/* ===================== REGION INFO + BRANCH -> REGION ===================== */
function resolveRegionByBranch_(leaderId, branchCode) {
  const { branchesSh } = getLeaderSheets_(leaderId);
  const codeNorm = normBranchCode_(branchCode);
  if (!codeNorm) return "";

  const values = branchesSh.getDataRange().getValues();
  if (values.length < 2) return "";

  const headers = values[0].map(x => String(x || "").trim());
  const idxCode = idxOfHeader_(headers, "branch_code");
  const idxReg = idxOfHeader_(headers, "region_id");
  const idxSt = idxOfHeader_(headers, "status");

  if (idxCode < 0 || idxReg < 0) return "";

  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const rowCodeNorm = normBranchCode_(row[idxCode]);
    if (rowCodeNorm !== codeNorm) continue;

    if (idxSt >= 0) {
      const ok = truthy_(row[idxSt]);
      if (!ok) return "";
    }

    return String(row[idxReg] || "").trim();
  }

  return "";
}

function getRegionInfo_(leaderId, regionId) {
  const { regionsSh } = getLeaderSheets_(leaderId);
  regionId = norm_(regionId);
  if (!regionId) return null;

  const values = regionsSh.getDataRange().getValues();
  if (values.length < 2) return null;

  const headers = values[0].map(x => String(x || "").trim());
  const idxId = idxOfHeader_(headers, "region_id");
  if (idxId < 0) return null;

  const idxName = idxOfHeader_(headers, "region_name");
  const idxType = idxOfHeader_(headers, "region_type");
  const idxMeet = idxOfHeader_(headers, "meet_url");
  const idxMap = idxOfHeader_(headers, "map_url");
  const idxSt = idxOfHeader_(headers, "status");

  for (let i = 1; i < values.length; i++) {
    const r = values[i] || [];
    if (String(r[idxId] || "").trim() !== regionId) continue;

    const status = (idxSt >= 0) ? truthy_(r[idxSt]) : true;

    return {
      region_id: regionId,
      region_name: idxName >= 0 ? norm_(r[idxName]) : "",
      region_type: idxType >= 0 ? norm_(r[idxType]) : "",
      meet_url: idxMeet >= 0 ? norm_(r[idxMeet]) : "",
      map_url: idxMap >= 0 ? norm_(r[idxMap]) : "",
      status
    };
  }
  return null;
}

/* ===================== SLOTS (Public) ===================== */
// Legacy (day/time) kept as-is for older page flows
function getAvailableSlots(leaderId, branchCode) {
  if (!leaderId) return {};
  if (isReservationSystemClosedForLeader_(leaderId)) return {};

  const { slotsSh, reservationsSh } = getLeaderSheets_(leaderId);

  const regionId = resolveRegionByBranch_(leaderId, branchCode);
  if (!regionId) return {};

  // columns: region_id, Day, Time, Active
  const slots = slotsSh.getDataRange().getValues();
  const booked = reservationsSh.getDataRange().getValues();

  const slotsByDay = {};

  for (let i = 1; i < slots.length; i++) {
    const rId = String(slots[i][0] || "").trim();
    const day = slots[i][1];
    const time = slots[i][2];
    const active = slots[i][3];

    if (rId !== regionId) continue;

    const isActive = truthy_(active);
    if (day && time && isActive) {
      if (!slotsByDay[day]) slotsByDay[day] = [];
      slotsByDay[day].push(String(time).trim());
    }
  }

  for (let j = 1; j < booked.length; j++) {
    const bookedDay = booked[j][1];
    const bookedTime = booked[j][2];

    if (slotsByDay[bookedDay]) {
      slotsByDay[bookedDay] = slotsByDay[bookedDay]
        .filter(t => t !== String(bookedTime).trim());
    }
  }

  Object.keys(slotsByDay).forEach(d => {
    if (!slotsByDay[d] || slotsByDay[d].length === 0) delete slotsByDay[d];
  });

  return slotsByDay;
}

// V2 slots by dateStr + slot_id
function getAvailableSlotsV2(leaderId, branchCode) {
  if (!leaderId) return {};
  if (isReservationSystemClosedForLeader_(leaderId)) return {};

  const { slotsSh, reservationsSh } = getLeaderSheets_(leaderId);

  const regionId = resolveRegionByBranch_(leaderId, branchCode);
  if (!regionId) return {};

  const tz = getTimeZoneForLeader_(leaderId) || DEFAULT_TZ;

  const slots = slotsSh.getDataRange().getValues();
  const booked = reservationsSh.getDataRange().getValues();

  const bookedSlotIds = new Set();
  const bookedDateTime = new Set();

  const bookedHeaders = booked[0] || [];
  function pickBookedIdx_(candidates, fallbackIdx) {
    for (let i = 0; i < candidates.length; i++) {
      const idx = idxOfHeader_(bookedHeaders, candidates[i]);
      if (idx !== -1) return idx;
    }
    return fallbackIdx;
  }

  const cBookedDate = pickBookedIdx_(["date", "day"], 1);
  const cBookedTime = pickBookedIdx_(["timeText", "time_text", "apptText", "appointment_text"], 2);
  const cBookedSlotId = pickBookedIdx_(["slot_id", "slotId"], 9);
  const cBookedStatus = pickBookedIdx_(["status"], 10);
  const cOldSlotReleased = pickBookedIdx_(["old_slot_released", "oldSlotReleased"], 14);

  function blocksSlot_(row) {
    const status = norm_(row[cBookedStatus] || "").toLowerCase() || "upcoming";
    const oldSlotReleased = truthy_(row[cOldSlotReleased]);

    if (status === "upcoming") return true;
    if (status === "cancelled" || status === "rescheduled") return !oldSlotReleased;

    return true;
  }

  for (let r = 1; r < booked.length; r++) {
    const row = booked[r] || [];
    if (!blocksSlot_(row)) continue;

    const bDate = toYMD_(row[cBookedDate], tz);
    const bTime = String(row[cBookedTime] || "").trim();
    const bSlotId = String(row[cBookedSlotId] || "").trim();

    if (bSlotId) bookedSlotIds.add(bSlotId);
    if (bDate && bTime) bookedDateTime.add(`${bDate}|${bTime}`);
  }

  const out = {};

  const lastRow = slotsSh.getLastRow();
  if (lastRow < 2) return out;

  const values = slotsSh.getRange(2, 1, lastRow - 1, 5).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i] || [];
    const slotId = String(row[0] || "").trim();
    const rId = String(row[1] || "").trim();
    const dateStr = toYMD_(row[2], tz);
    const timeTxt = String(row[3] || "").trim();
    const active = row[4];

    if (!slotId || !rId || !dateStr || !timeTxt) continue;
    if (rId !== regionId) continue;
    if (!truthy_(active)) continue;
    if (isSlotExpiredAtNow_(dateStr, timeTxt, tz)) continue;

    if (bookedSlotIds.has(slotId)) continue;
    if (bookedDateTime.has(`${dateStr}|${timeTxt}`)) continue;

    if (!out[dateStr]) out[dateStr] = [];
    out[dateStr].push({ slot_id: slotId, timeText: timeTxt });
  }

  return out;
}

/* ===================== BOOKING V2 ===================== */
function bookSlotV2(leaderId, name, email, dateStr, slotId, employeeId, branchCode) {
  let lock;
  try {
    lock = acquireLeaderLock_(leaderId, 15000);

    if (!leaderId) return { success: false, message: "⚠️ اختر المشرف أولاً." };
    if (isReservationSystemClosedForLeader_(leaderId)) {
      return { success: false, message: "⛔ تم إغلاق الحجز. انتهت فترة استقبال المواعيد." };
    }

    const { reservationsSh } = getLeaderSheets_(leaderId);

    employeeId = norm_(employeeId);
    if (employeeId.length < 4) {
      return { success: false, message: "⚠️ رقم الموظف يجب أن يكون 4 أرقام على الأقل." };
    }

    const codeNorm = normBranchCode_(branchCode);
    if (!codeNorm) return { success: false, message: "⚠️ اختر فرعك أولاً." };

    const regionId = resolveRegionByBranch_(leaderId, codeNorm);
    if (!regionId) return { success: false, message: "⚠️ هذا الفرع غير مربوط بمنطقة. تواصل مع المشرف." };

    dateStr = norm_(dateStr);
    slotId = norm_(slotId);
    if (!dateStr) return { success: false, message: "⚠️ اختر اليوم أولاً." };
    if (!slotId) return { success: false, message: "⚠️ اختر الوقت أولاً." };

    // prevent duplicate employeeId only for active future appointments
    const lastRow = reservationsSh.getLastRow();
    if (lastRow >= 2) {
      const cols = getReservationCols_(reservationsSh);
      const tz = getTimeZoneForLeader_(leaderId) || DEFAULT_TZ;
      const rows = reservationsSh.getRange(2, 1, lastRow - 1, cols.lastCol).getValues();

      for (let i = 0; i < rows.length; i++) {
        const existingRow = rows[i] || [];
        const existingEmployeeId = norm_(existingRow[cols.cEmployeeId]);
        if (existingEmployeeId !== employeeId) continue;

        const existingStatus = norm_(existingRow[cols.cStatus] || '').toLowerCase() || 'upcoming';
        const existingIsPast = isPastReservationRow_(existingRow, cols, tz);
        if (existingStatus === 'upcoming' && !existingIsPast) {
          return {
            success: false,
            message: "⚠️ تم الحجز مسبقًا بهذا الرقم، في حال رغبتك بتعديل الموعد نرجو إلغاء الحجز عبر رابط الإلغاء في البريد الإلكتروني."
          };
        }
      }
    }

    // re-check availability & pick timeText
    const avail = getAvailableSlotsV2(leaderId, codeNorm);
    const list = avail[dateStr] || [];
    const picked = list.find(x => String(x.slot_id || "").trim() === slotId);
    if (!picked) return { success: false, message: "⚠️ الموعد تم حجزه بالفعل، اختر موعدًا آخر." };

    const timeText = String(picked.timeText || "").trim();
    const id = Utilities.getUuid();

    // A name, B date, C timeText, D createdAt, E email, F reservationId, G employeeId, H branch, I region, J slot_id
    reservationsSh.appendRow([name, dateStr, timeText, new Date(), email, id, employeeId, "PH" + codeNorm, regionId, slotId]);

    // Email (as per your stable notes)
    try {
      const baseUrl = RESERVATION_URL;
      const cancelLink = `${baseUrl}?cancel=${encodeURIComponent(id)}&leader=${encodeURIComponent(leaderId)}`;

      const tz = getTimeZoneForLeader_(leaderId) || DEFAULT_TZ;
      const bookedAt = Utilities.formatDate(new Date(), tz, "yyyy/MM/dd");

      const region = getRegionInfo_(leaderId, regionId);
      const regionName = (region && region.region_name) ? region.region_name : regionId;
      const meetUrl = (region && region.meet_url) ? region.meet_url : "";
      const mapUrl = (region && region.map_url) ? region.map_url : "";

      const subject = "✅ تأكيد موعد التقييم";
      const body = buildConfirmationEmailMessage_(leaderId, {
        employeeName: name,
        branchCode: codeNorm,
        regionName,
        appointmentDate: dateStr,
        bookedAt,
        appointmentTime: timeText,
        employeeId,
        meetUrl,
        mapUrl,
        cancelLink,
      });
      const htmlBody = confirmationEmailHtml_(subject, body);

      const leaderInfo = getLeaderInfoById_(leaderId);
      const supervisorEmail = leaderInfo && leaderInfo.admin_email ? leaderInfo.admin_email : "";

      const mailOptions = {
        to: email,
        subject,
        body,
        htmlBody,
        name: "ادارة المنطقة الجنوبية | Day For Development "
      };
      if (supervisorEmail) mailOptions.cc = supervisorEmail;

      MailApp.sendEmail(mailOptions);
    } catch (err) {
      Logger.log("Email Error: " + err);
    }

    return { success: true, message: "تم الحجز وإرسال تأكيد بالبريد الإلكتروني ✅" };

  } finally {
    releaseLeaderLock_(lock);
  }
}

/* ===================== CANCELLATION ===================== */
function handleCancellation(params, baseUrl) {
  const safeBaseUrl = baseUrl || ScriptApp.getService().getUrl();
  try {
    const p = params || {};
    const cancelId = norm_(p.cancel);
    const leaderId = norm_(p.leader);
    const confirmCancel = norm_(p.confirm_cancel) === "1";
    const cancelReasonChoice = norm_(p.reason);
    const customReason = norm_(p.other_reason);

    if (!cancelId) {
      return HtmlService.createHtmlOutput("<h3 style='text-align:center;color:red'>❌ رابط الإلغاء غير صحيح</h3>");
    }
    if (!leaderId) {
      return HtmlService.createHtmlOutput("<h3 style='text-align:center;color:red'>❌ رابط الإلغاء ناقص (leader)</h3>");
    }

    const { reservationsSh } = getLeaderSheets_(leaderId);
    const cols = getReservationCols_(reservationsSh);
    const found = findReservationRowById_(reservationsSh, cols, cancelId);
    let cancelled = false;

    if (!found) {
      return HtmlService.createHtmlOutput("<h3 style='text-align:center;color:red'>❌ لم يتم العثور على هذا الحجز</h3>");
    }

    const row = found.row;
    const rowIndex = found.rowIndex;
    const status = norm_(row[cols.cStatus] || '').toLowerCase() || 'upcoming';

    if (!confirmCancel) {
      const html = `
        <html lang="ar" dir="rtl">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>إلغاء الموعد</title>
          <style>
            body{font-family:'Tajawal',sans-serif;background:#f4f7fb;margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px;color:#1f2937}
            .card{width:min(560px,100%);background:#fff;border:1px solid #dbe5f0;border-radius:22px;box-shadow:0 20px 50px rgba(15,23,42,.12);padding:28px}
            h2{margin:0 0 12px;color:#0b3f77}
            p{margin:0 0 16px;line-height:1.7;color:#475569}
            .warning{padding:14px 16px;border-radius:16px;background:#fff7ed;border:1px solid #fdba74;color:#9a3412;font-weight:700;line-height:1.8}
            .label{display:block;margin:18px 0 8px;font-weight:800;color:#334155}
            select,textarea{width:100%;padding:13px 14px;border-radius:14px;border:1px solid #cbd5e1;font-size:15px;box-sizing:border-box}
            textarea{min-height:110px;resize:vertical;display:none;margin-top:10px}
            .actions{display:flex;gap:12px;flex-wrap:wrap;margin-top:22px}
            .btn{border:none;border-radius:14px;padding:13px 18px;font-weight:800;font-size:15px;cursor:pointer;text-decoration:none}
            .btn-primary{background:#b91c1c;color:#fff}
            .btn-secondary{background:#fff;border:1px solid #cbd5e1;color:#0f172a}
          </style>
        </head>
        <body>
          <div class="card">
            <h2>إلغاء الموعد</h2>
            <div class="warning">أنت على وشك إلغاء الموعد. سيصبح هذا الموعد متاحاً للآخرين للاختيار. هل أنت متأكد؟</div>
            <label class="label" for="reason">سبب الإلغاء</label>
            <form method="get" action="${safeBaseUrl}" target="_top" onsubmit="return validateCancelForm_()">
              <input type="hidden" name="cancel" value="${cancelId}">
              <input type="hidden" name="leader" value="${leaderId}">
              <input type="hidden" name="confirm_cancel" value="1">
              <select id="reason" name="reason" required onchange="toggleOtherReason_()">
                <option value="">اختر السبب</option>
                <option value="need_another_appointment">أحتاج موعداً آخر</option>
                <option value="other_reason">سبب آخر</option>
              </select>
              <textarea id="other_reason" name="other_reason" placeholder="اكتب سبب الإلغاء هنا"></textarea>
              <div class="actions">
                <a class="btn btn-secondary" href="${safeBaseUrl}" target="_top" rel="noopener noreferrer">عدم الإلغاء</a>
                <button type="submit" class="btn btn-primary">تأكيد الإلغاء</button>
              </div>
            </form>
          </div>
          <script>
            function toggleOtherReason_(){
              const sel = document.getElementById('reason');
              const box = document.getElementById('other_reason');
              const show = sel && sel.value === 'other_reason';
              if (box) {
                box.style.display = show ? 'block' : 'none';
                box.required = !!show;
              }
            }
            function validateCancelForm_(){
              const sel = document.getElementById('reason');
              const box = document.getElementById('other_reason');
              if (!sel || !sel.value) {
                alert('يرجى اختيار سبب الإلغاء');
                return false;
              }
              if (sel.value === 'other_reason' && (!box || !String(box.value || '').trim())) {
                alert('يرجى كتابة سبب الإلغاء');
                return false;
              }
              return true;
            }
          </script>
        </body>
        </html>`;

      return HtmlService.createHtmlOutput(html)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (!cancelReasonChoice) {
      return HtmlService.createHtmlOutput("<h3 style='text-align:center;color:red'>❌ يرجى اختيار سبب الإلغاء</h3>");
    }

    const finalReason = cancelReasonChoice === 'need_another_appointment'
      ? 'أحتاج موعداً آخر'
      : customReason;

    if (cancelReasonChoice === 'other_reason' && !finalReason) {
      return HtmlService.createHtmlOutput("<h3 style='text-align:center;color:red'>❌ يرجى كتابة سبب الإلغاء</h3>");
    }

    if (status !== 'cancelled') {
      reservationsSh.getRange(rowIndex, cols.cStatus + 1).setValue('cancelled');
      reservationsSh.getRange(rowIndex, cols.cCancelledAt + 1).setValue(new Date());
      reservationsSh.getRange(rowIndex, cols.cCancelledBy + 1).setValue('self');
      reservationsSh.getRange(rowIndex, cols.cOldSlotReleased + 1).setValue(true);
      reservationsSh.getRange(rowIndex, cols.cCancellationReason + 1).setValue(finalReason);
      cancelled = true;
    }

    const msg = cancelled
      ? '✅ تم إلغاء الحجز بنجاح'
      : '⚠️ لم يتم العثور على هذا الحجز أو تم إلغاؤه مسبقًا';

    const html = `
      <html lang="ar" dir="rtl">
      <head>
        <meta charset="UTF-8">
        <title>إلغاء الحجز</title>
        <style>
          body { font-family:'Tajawal', sans-serif; text-align:center; padding-top:80px; background:#f4f6f8; color:#333; }
          h3 { color: ${cancelled ? '#1b5e20' : '#b71c1c'}; margin-bottom:10px; font-size:1.3em; }
          p { color:#555; font-size:15px; margin-bottom:25px; }
          .button { display:inline-block; margin-top:15px; background:#004c97; color:#fff; padding:10px 20px; border-radius:8px; text-decoration:none; font-weight:bold; border:none; cursor:pointer; }
          .button:hover { background:#0360c0; }
          .spinner { margin:25px auto; width:45px; height:45px; border:4px solid #cfd9e0; border-top-color:#004c97; border-radius:50%; animation: spin 1s linear infinite; }
          @keyframes spin { to { transform: rotate(360deg); } }
          .fade-in { animation: fadeIn 0.8s ease-in-out; }
          @keyframes fadeIn { from { opacity:0; transform: translateY(10px);} to { opacity:1; transform: translateY(0);} }
        </style>
      </head>
      <body>
        <h3 class="fade-in">${msg}</h3>
        <div class="spinner"></div>
        <p>تم حفظ الإلغاء. للعودة إلى صفحة الحجز اضغط الزر التالي:</p>
        <form method="get" action="${safeBaseUrl}" target="_top" style="margin-top:18px">
          <button type="submit" class="button" autofocus>العودة إلى صفحة الحجز</button>
        </form>
      </body>
      </html>`;

    return HtmlService.createHtmlOutput(html)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    Logger.log(err);
    return HtmlService.createHtmlOutput("<h3 style='text-align:center;color:red'>❌ حدث خطأ أثناء الإلغاء</h3>");
  }
}

/* ===================== ADMIN: CLOSE DATE ===================== */
function api_adminGetCloseDate(leaderId, token) {
  requireAdmin_(token, leaderId);
  const closeAtStr = getReservationCloseAtStringForLeader_(leaderId) || "";
  return { ok: true, closeAtStr };
}

function api_adminSetCloseDate(leaderId, token, closeAtStr) {
  requireAdmin_(token, leaderId);

  closeAtStr = norm_(closeAtStr);

  // Accept:
  // 1) YYYY-MM-DD
  // 2) YYYY-MM-DD HH:MM
  const m = closeAtStr.match(/^(\d{4})-(\d{2})-(\d{2})(?:\s+(\d{2}):(\d{2}))?$/);
  if (!m) return { ok: false, message: "صيغة غير صحيحة. استخدم YYYY-MM-DD أو YYYY-MM-DD HH:MM" };

  const ymd = `${m[1]}-${m[2]}-${m[3]}`;
  const hh = (m[4] ?? "00");
  const mm = (m[5] ?? "00");
  const normalized = `${ymd} ${hh}:${mm}`;

  const { settingsSh } = getLeaderSheets_(leaderId);
  setSettingValue_(settingsSh, "RESERVATION_CLOSE_AT", normalized);

  return { ok: true };
}

function api_adminSendLeaderInvite(leaderId, token, payload) {
  requireAdmin_(token, leaderId);

  const adminEmail = norm_(payload && payload.admin_email).toLowerCase();
  if (!adminEmail) return { ok: false, message: "Leader email is required" };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(adminEmail)) {
    return { ok: false, message: "Invalid email format" };
  }
  if (isAdminEmailTaken_(adminEmail)) {
    return { ok: false, message: "This email is already used by an existing leader" };
  }

  const invite = createLeaderInvite_(leaderId, adminEmail);
  const inviteLink = buildLeaderInviteLink_(invite.invite_token);
  const senderInfo = getLeaderInfoById_(leaderId) || {};
  const senderName = norm_(senderInfo.leader_name) || leaderId;

  MailApp.sendEmail({
    to: adminEmail,
    subject: "Create your leader account",
    body:
      `Hello,\n\n` +
      `${senderName} has invited you to create a leader account.\n\n` +
      `Open this link to create your account:\n${inviteLink}\n\n` +
      `This link expires in 7 days.\n`,
    htmlBody: brandedEmailHtml_({
      direction: "ltr",
      eyebrow: "Day For Development",
      title: "Create your leader account",
      introHtml: `<strong>${htmlEscape_(senderName)}</strong> invited you to create a leader account.`,
      messageHtml: `Open the link below to finish creating your account.`,
      actionLabel: "Create account",
      actionUrl: inviteLink,
      note: "This invitation link expires in 7 days."
    }),
    name: "Day For Development"
  });

  return { ok: true, invite_link: inviteLink, admin_email: adminEmail };
}

function api_sendLeaderInviteBySession(token, payload) {
  const sess = getAdminSession_(token);
  const adminEmail = norm_(payload && payload.admin_email).toLowerCase();
  let assignedSeniorEmail = norm_(payload && payload.senior_email).toLowerCase();
  if (!adminEmail) return { ok: false, message: "Leader email is required" };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(adminEmail)) {
    return { ok: false, message: "Invalid email format" };
  }
  if (isAdminEmailTaken_(adminEmail)) {
    return { ok: false, message: "This email is already used by an existing leader" };
  }

  let createdBy = "";
  let senderName = "";
  if (sess.role === "senior") {
    const seniorEmail = norm_(sess.senior_email).toLowerCase();
    const seniorAccount = listSeniorAccounts_().find((item) => item.senior_email === seniorEmail) || null;
    if (!assignedSeniorEmail) assignedSeniorEmail = seniorEmail;
    createdBy = "senior:" + seniorEmail;
    senderName = (seniorAccount && norm_(seniorAccount.senior_name)) || seniorEmail;
  } else {
    const leaderId = norm_(sess.leader_id);
    if (!leaderId) throw new Error("Missing leader session.");
    const senderInfo = getLeaderInfoById_(leaderId) || {};
    createdBy = leaderId;
    senderName = norm_(senderInfo.leader_name) || leaderId;
  }

  if (assignedSeniorEmail) {
    const seniorAccount = listSeniorAccounts_().find((item) => item.senior_email === assignedSeniorEmail) || null;
    if (!seniorAccount) return { ok: false, message: "Selected senior account was not found" };
  }

  const invite = createLeaderInvite_(createdBy, adminEmail, assignedSeniorEmail);
  const inviteLink = buildLeaderInviteLink_(invite.invite_token);

  MailApp.sendEmail({
    to: adminEmail,
    subject: "Create your leader account",
    body:
      "Hello,\n\n" +
      senderName + " has invited you to create a leader account.\n\n" +
      "Open this link to create your account:\n" + inviteLink + "\n\n" +
      "This link expires in 7 days.\n",
    htmlBody: brandedEmailHtml_({
      direction: "ltr",
      eyebrow: "Day For Development",
      title: "Create your leader account",
      introHtml: `<strong>${htmlEscape_(senderName)}</strong> invited you to create a leader account.`,
      messageHtml: `Open the link below to finish creating your account.`,
      actionLabel: "Create account",
      actionUrl: inviteLink,
      note: "This invitation link expires in 7 days."
    }),
    name: "Day For Development"
  });

  return { ok: true, invite_link: inviteLink, admin_email: adminEmail, senior_email: assignedSeniorEmail };
}

function api_sendSeniorInviteBySession(token, payload) {
  const sess = getAdminSession_(token);
  if (!sess || sess.role !== "senior") throw new Error("Senior session required.");

  const seniorEmail = norm_(payload && payload.senior_email).toLowerCase();
  if (!seniorEmail) return { ok: false, message: "Senior email is required" };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(seniorEmail)) {
    return { ok: false, message: "Invalid email format" };
  }
  if (seniorEmailExists_(seniorEmail)) {
    return { ok: false, message: "This email is already used by an existing senior" };
  }

  const currentSeniorEmail = norm_(sess.senior_email).toLowerCase();
  const seniorAccount = listSeniorAccounts_().find((item) => item.senior_email === currentSeniorEmail) || null;
  const senderName = (seniorAccount && norm_(seniorAccount.senior_name)) || currentSeniorEmail;
  const invite = createSeniorInvite_("senior:" + currentSeniorEmail, seniorEmail);
  const inviteLink = buildSeniorInviteLink_(invite.invite_token);

  MailApp.sendEmail({
    to: seniorEmail,
    subject: "Create your senior account",
    body:
      "Hello,\n\n" +
      senderName + " has invited you to create a senior account.\n\n" +
      "Open this link to create your account:\n" + inviteLink + "\n\n" +
      "This link expires in 7 days.\n",
    htmlBody: brandedEmailHtml_({
      direction: "ltr",
      eyebrow: "Day For Development",
      title: "Create your senior account",
      introHtml: `<strong>${htmlEscape_(senderName)}</strong> invited you to create a senior account.`,
      messageHtml: `Open the link below to finish creating your account.`,
      actionLabel: "Create account",
      actionUrl: inviteLink,
      note: "This invitation link expires in 7 days."
    }),
    name: "Day For Development"
  });

  return { ok: true, invite_link: inviteLink, senior_email: seniorEmail };
}

function api_adminGetInviteInfo(inviteToken) {
  const invite = inviteRecordByToken_(inviteToken);
  if (!invite) return { ok: false, message: "Invite link is invalid" };

  const status = norm_(invite.status).toLowerCase();
  if (status && status !== "pending") {
    return { ok: false, message: "This invite link is no longer available" };
  }
  if (inviteIsExpired_(invite)) {
    updateInviteField_(invite.rowIndex, "status", "expired");
    return { ok: false, message: "This invite link has expired" };
  }

  return {
    ok: true,
    invite: {
      invite_type: norm_(invite.invite_type).toLowerCase() || "leader",
      admin_email: norm_(invite.admin_email),
      masked_email: maskEmail_(invite.admin_email),
      senior_email: norm_(invite.senior_email).toLowerCase()
    }
  };
}

function updateLeaderRegistryField_(leaderId, fieldName, value) {
  const regSh = SpreadsheetApp.openById(REGISTRY_SHEET_ID).getSheetByName(REGISTRY_TAB);
  if (!regSh) throw new Error("Registry tab not found: " + REGISTRY_TAB);

  const values = regSh.getDataRange().getValues();
  if (values.length < 2) throw new Error("Registry is empty");

  const header = values[0].map(h => String(h || "").trim());
  const iId = header.indexOf("leader_id");
  const iField = header.indexOf(fieldName);
  if (iId < 0 || iField < 0) throw new Error("Registry missing required columns");

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][iId]).trim() === String(leaderId || "").trim()) {
      regSh.getRange(r + 1, iField + 1).setValue(value);
      return true;
    }
  }
  throw new Error("Leader not found in registry");
}

function splitLeaderName_(leaderName) {
  const rawName = norm_(leaderName).replace(/^د(?:كتور)?\.?\s*\/?\s*/i, "").trim();
  const parts = rawName ? rawName.split(/\s+/) : [];
  return {
    first_name_ar: parts.shift() || "",
    second_name: parts.join(" ")
  };
}

function nextLeaderId_() {
  const regSh = SpreadsheetApp.openById(REGISTRY_SHEET_ID).getSheetByName(REGISTRY_TAB);
  if (!regSh) throw new Error("Registry tab not found: " + REGISTRY_TAB);

  const values = regSh.getDataRange().getValues();
  if (values.length < 2) return "TL001";

  const header = values[0].map(h => String(h || "").trim());
  const iId = header.indexOf("leader_id");
  if (iId < 0) throw new Error("Registry missing leader_id column");

  let maxNum = 0;
  for (let r = 1; r < values.length; r++) {
    const rawId = String(values[r][iId] || "").trim().toUpperCase();
    const match = rawId.match(/^TL(\d+)$/);
    if (!match) continue;

    const n = Number(match[1]);
    if (!isNaN(n) && n > maxNum) maxNum = n;
  }

  return "TL" + String(maxNum + 1).padStart(3, "0");
}

function isAdminEmailTaken_(email) {
  const cleanEmail = norm_(email).toLowerCase();
  if (!cleanEmail) return false;

  const regSh = SpreadsheetApp.openById(REGISTRY_SHEET_ID).getSheetByName(REGISTRY_TAB);
  if (!regSh) throw new Error("Registry tab not found: " + REGISTRY_TAB);

  const values = regSh.getDataRange().getValues();
  if (values.length < 2) return false;

  const header = values[0].map(h => String(h || "").trim());
  const iEmail = header.indexOf("admin_email");
  const iStatus = header.indexOf("status");
  if (iEmail < 0) return false;

  for (let r = 1; r < values.length; r++) {
    const rowEmail = norm_(values[r][iEmail]).toLowerCase();
    const rowStatus = iStatus >= 0 ? norm_(values[r][iStatus]).toLowerCase() : "active";
    if (rowEmail && rowEmail === cleanEmail && rowStatus !== "deleted") return true;
  }

  return false;
}

function initializeLeaderDb_(leaderId, leaderName) {
  const ss = SpreadsheetApp.create(`DFD - ${leaderName || leaderId}`);
  const defaultSheet = ss.getSheets()[0];

  defaultSheet.setName(LEADER_DB_SHEETS.RESERVATIONS);
  defaultSheet.clear();
  defaultSheet.getRange(1, 1, 1, 18).setValues([[
    "name",
    "date",
    "timeText",
    "createdAt",
    "email",
    "reservationId",
    "employeeId",
    "branch",
    "region_id",
    "slot_id",
    "status",
    "cancelled_at",
    "cancelled_by",
    "rescheduled_from",
    "old_slot_released",
    "cancellation_reason",
    "feedback_sent_at",
    "feedback_message"
  ]]);

  const settingsSh = ss.insertSheet(LEADER_DB_SHEETS.SETTINGS);
  settingsSh.getRange(1, 1, 4, 2).setValues([
    ["TIMEZONE", DEFAULT_TZ],
    [EMAIL_SETTINGS_KEYS.SIGNATURE, ""],
    [EMAIL_SETTINGS_KEYS.RESERVATION_CLOSE_AT, ""],
    [EMAIL_SETTINGS_KEYS.CONFIRMATION_TEMPLATE, defaultConfirmationEmailTemplate_()]
  ]);

  const slotsSh = ss.insertSheet(LEADER_DB_SHEETS.SLOTS);
  slotsSh.getRange(1, 1, 1, 5).setValues([["slot_id", "region_id", "date", "timeText", "active"]]);

  const regionsSh = ss.insertSheet(LEADER_DB_SHEETS.REGIONS);
  regionsSh.getRange(1, 1, 1, 6).setValues([["region_id", "region_name", "region_type", "meet_url", "map_url", "status"]]);
  ensureUnmappedRegion_(regionsSh);

  const branchesSh = ss.insertSheet(LEADER_DB_SHEETS.BRANCHES);
  branchesSh.getRange(1, 1, 1, 4).setValues([["branch_code", "branch_name", "region_id", "status"]]);

  return ss.getId();
}

function appendLeaderRegistryRow_(payload) {
  return appendRegistryRowWithMap_(payload);
}

function appendRegistryRowWithMap_(payload) {
  const regSh = SpreadsheetApp.openById(REGISTRY_SHEET_ID).getSheetByName(REGISTRY_TAB);
  if (!regSh) throw new Error("Registry tab not found: " + REGISTRY_TAB);

  const values = regSh.getDataRange().getValues();
  if (!values.length) throw new Error("Registry header row is missing");

  const header = values[0].map(h => String(h || "").trim());
  const row = header.map((fieldName) => {
    if (fieldName === "status") return payload.status || "active";
    return payload[fieldName] == null ? "" : payload[fieldName];
  });

  regSh.appendRow(row);
}

function appendSeniorRegistryRow_(payload) {
  return appendRegistryRowWithMap_(payload);
}

function api_adminCreateInvitedAccount(payload) {
  const inviteToken = norm_(payload && payload.invite_token);
  const firstNameAr = norm_(payload && payload.first_name_ar);
  const secondName = norm_(payload && payload.second_name);
  const adminEmail = norm_(payload && payload.admin_email).toLowerCase();
  const selectedSeniorEmail = norm_(payload && payload.senior_email).toLowerCase();
  const password = String((payload && payload.password) || "");

  const invite = inviteRecordByToken_(inviteToken);
  if (!invite) return { ok: false, message: "Invite link is invalid" };
  const inviteStatus = norm_(invite.status).toLowerCase();
  if (inviteStatus && inviteStatus !== "pending") {
    return { ok: false, message: "This invite link is no longer available" };
  }
  if (inviteIsExpired_(invite)) {
    updateInviteField_(invite.rowIndex, "status", "expired");
    return { ok: false, message: "This invite link has expired" };
  }

  if (!firstNameAr) return { ok: false, message: "الاسم الأول مطلوب" };
  if (!adminEmail) return { ok: false, message: "البريد الإلكتروني مطلوب" };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(adminEmail)) {
    return { ok: false, message: "صيغة البريد الإلكتروني غير صحيحة" };
  }
  if (!password || password.length < 6) {
    return { ok: false, message: "كلمة المرور يجب أن تكون 6 أحرف على الأقل" };
  }
  if (adminEmail !== norm_(invite.admin_email).toLowerCase()) {
    return { ok: false, message: "Use the same email address that received the invite" };
  }
  if (isAdminEmailTaken_(adminEmail)) {
    return { ok: false, message: "هذا البريد الإلكتروني مستخدم بالفعل" };
  }

  const inviteType = norm_(invite.invite_type).toLowerCase() || "leader";
  const fullName = [firstNameAr, secondName].filter(Boolean).join(" ").trim();

  if (inviteType === "senior") {
    if (seniorEmailExists_(adminEmail)) {
      return { ok: false, message: "This email is already used by an existing senior" };
    }

    const seniorName = fullName || adminEmail;
    appendSeniorRegistryRow_({
      leader_id: "",
      leader_name: "",
      admin_email: "",
      sheet_id: "",
      admin_pass_hash: "",
      senior_email: adminEmail,
      senior_name: seniorName,
      senior_pass_hash: sha256_(password),
      status: "active"
    });
    markInviteUsed_(invite, "senior:" + adminEmail);

    const token = createSeniorSession_(adminEmail);
    return {
      ok: true,
      role: "senior",
      token,
      senior_email: adminEmail,
      senior_name: seniorName
    };
  }

  const effectiveSeniorEmail = selectedSeniorEmail || norm_(invite.senior_email).toLowerCase();
  if (!effectiveSeniorEmail) {
    return { ok: false, message: "Choose a senior account before creating the leader" };
  }
  const seniorAccount = listSeniorAccounts_().find((item) => item.senior_email === effectiveSeniorEmail) || null;
  if (!seniorAccount) {
    return { ok: false, message: "Selected senior account was not found" };
  }

  const leaderName = `د. ${fullName}`.trim();
  const leaderId = nextLeaderId_();
  const sheetId = initializeLeaderDb_(leaderId, leaderName);

  appendLeaderRegistryRow_({
    leader_id: leaderId,
    leader_name: leaderName,
    admin_email: adminEmail,
    sheet_id: sheetId,
    admin_pass_hash: sha256_(password),
    senior_email: effectiveSeniorEmail,
    senior_name: norm_(seniorAccount.senior_name) || effectiveSeniorEmail,
    status: "active"
  });
  markInviteUsed_(invite, leaderId);

  const token = createAdminSession_(leaderId);
  return {
    ok: true,
    role: "leader",
    token,
    leader_id: leaderId,
    leader_name: leaderName
  };
}

function api_adminGetProfile(leaderId, token) {
  requireAdmin_(token, leaderId);
  const info = getLeaderInfoById_(leaderId) || {};
  const { settingsSh } = getLeaderSheets_(leaderId);
  const signature = getSettingValue_(settingsSh, EMAIL_SETTINGS_KEYS.SIGNATURE) || "";
  const nameParts = splitLeaderName_(info.leader_name);

  return {
    ok: true,
    profile: {
      leader_id: norm_(info.leader_id),
      leader_name: norm_(info.leader_name),
      first_name_ar: nameParts.first_name_ar,
      second_name: nameParts.second_name,
      admin_email: norm_(info.admin_email),
      email_signature: norm_(signature)
    }
  };
}

function updateSeniorRegistryField_(seniorEmail, fieldName, value) {
  const targetEmail = norm_(seniorEmail).toLowerCase();
  const { regSh, headers, rows } = getRegistryRows_();
  const idxSeniorEmail = idxOfHeader_(headers, "senior_email");
  const idxField = idxOfHeader_(headers, fieldName);
  if (idxSeniorEmail === -1 || idxField === -1) {
    throw new Error("Missing required registry column: " + fieldName);
  }

  rows.forEach((row, i) => {
    if (norm_(row[idxSeniorEmail]).toLowerCase() !== targetEmail) return;
    regSh.getRange(i + 2, idxField + 1).setValue(value);
  });
}

function api_seniorGetProfile(seniorEmail, token) {
  requireSenior_(token, seniorEmail);
  const email = norm_(seniorEmail).toLowerCase();
  const account = listSeniorAccounts_().find((item) => item.senior_email === email) || {};

  return {
    ok: true,
    profile: {
      senior_email: email,
      senior_name: norm_(account.senior_name) || email
    }
  };
}

function api_seniorRequestEmailChangeCode(seniorEmail, token, newEmail) {
  requireSenior_(token, seniorEmail);
  const currentEmail = norm_(seniorEmail).toLowerCase();
  const targetEmail = norm_(newEmail).toLowerCase();
  if (!targetEmail) return { ok: false, message: "New email is required" };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(targetEmail)) {
    return { ok: false, message: "Invalid email format" };
  }
  if (targetEmail === currentEmail) {
    return { ok: false, message: "Enter a different email address" };
  }
  if (seniorEmailExists_(targetEmail, currentEmail)) {
    return { ok: false, message: "This email is already used by another senior" };
  }

  const account = listSeniorAccounts_().find((item) => item.senior_email === currentEmail) || {};
  const code = generateVerificationCode_();
  CacheService.getScriptCache().put(
    seniorEmailChangeCacheKey_(currentEmail),
    JSON.stringify({
      senior_email: currentEmail,
      new_email: targetEmail,
      code: code,
      issued_at: Date.now()
    }),
    15 * 60
  );

  MailApp.sendEmail({
    to: targetEmail,
    subject: "Senior email verification code",
    body:
      "Hello,\n\n" +
      "Use this verification code to confirm your new senior email address:\n" +
      code + "\n\n" +
      "This code expires in 15 minutes.\n",
    htmlBody:
      '<div style="font-family:Arial,sans-serif;line-height:1.7;color:#1f2937">' +
        "<p>Hello,</p>" +
        "<p>Use this verification code to confirm your new senior email address.</p>" +
        '<div style="margin:18px 0;padding:16px;border-radius:14px;background:#eef2ff;color:#312e81;font-size:28px;font-weight:800;letter-spacing:4px;text-align:center">' +
          htmlEscape_(code) +
        "</div>" +
        "<p>This code expires in 15 minutes.</p>" +
      "</div>",
    name: "DFD Admin"
  });

  return {
    ok: true,
    masked_email: maskEmail_(targetEmail),
    senior_name: norm_(account.senior_name) || currentEmail
  };
}

function api_seniorConfirmEmailChange(seniorEmail, token, newEmail, code) {
  requireSenior_(token, seniorEmail);
  const currentEmail = norm_(seniorEmail).toLowerCase();
  const targetEmail = norm_(newEmail).toLowerCase();
  const enteredCode = norm_(code);
  if (!targetEmail) return { ok: false, message: "New email is required" };
  if (!enteredCode) return { ok: false, message: "Verification code is required" };

  const raw = CacheService.getScriptCache().get(seniorEmailChangeCacheKey_(currentEmail));
  if (!raw) {
    return { ok: false, message: "Verification code expired. Request a new one." };
  }

  const payload = JSON.parse(raw);
  if (norm_(payload.new_email).toLowerCase() !== targetEmail) {
    return { ok: false, message: "Use the same email that received the code" };
  }
  if (norm_(payload.code) !== enteredCode) {
    return { ok: false, message: "Verification code is incorrect" };
  }
  if (seniorEmailExists_(targetEmail, currentEmail)) {
    return { ok: false, message: "This email is already used by another senior" };
  }

  updateSeniorRegistryField_(currentEmail, "senior_email", targetEmail);
  CacheService.getScriptCache().remove(seniorEmailChangeCacheKey_(currentEmail));
  const nextToken = createSeniorSession_(targetEmail);
  const account = listSeniorAccounts_().find((item) => item.senior_email === targetEmail) || {};

  return {
    ok: true,
    token: nextToken,
    senior_email: targetEmail,
    senior_name: norm_(account.senior_name) || targetEmail
  };
}

function api_seniorUpdateProfile(seniorEmail, token, payload) {
  requireSenior_(token, seniorEmail);
  const email = norm_(seniorEmail).toLowerCase();
  const nextName = norm_(payload && payload.senior_name);
  if (!nextName) return { ok: false, message: "Senior name is required" };

  updateSeniorRegistryField_(email, "senior_name", nextName);
  return { ok: true, senior_name: nextName, senior_email: email };
}

function api_seniorChangePassword(seniorEmail, token, newPassword) {
  requireSenior_(token, seniorEmail);
  const nextPassword = String(newPassword || "");
  if (nextPassword.length < 6) {
    return { ok: false, message: "New password must be at least 6 characters" };
  }

  updateSeniorRegistryField_(seniorEmail, "senior_pass_hash", sha256_(nextPassword));
  return { ok: true };
}

function api_adminUpdateProfile(leaderId, token, payload) {
  requireAdmin_(token, leaderId);
  const firstNameAr = norm_(payload && payload.first_name_ar);
  const secondName = norm_(payload && payload.second_name);
  const nextEmail = norm_(payload && payload.admin_email);
  const nextSignature = norm_(payload && payload.email_signature);

  if (!firstNameAr) return { ok: false, message: "الاسم الأول مطلوب" };
  if (nextEmail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(nextEmail)) {
    return { ok: false, message: "صيغة البريد الإلكتروني غير صحيحة" };
  }

  const fullName = [firstNameAr, secondName].filter(Boolean).join(" ").trim();
  const nextName = `د. ${fullName}`.trim();

  updateLeaderRegistryField_(leaderId, "leader_name", nextName);
  updateLeaderRegistryField_(leaderId, "admin_email", nextEmail);

  const { settingsSh } = getLeaderSheets_(leaderId);
  setSettingValue_(settingsSh, EMAIL_SETTINGS_KEYS.SIGNATURE, nextSignature);

  return {
    ok: true,
    leader_name: nextName,
    admin_email: nextEmail,
    email_signature: nextSignature
  };
}

function api_adminChangePassword(leaderId, token, newPassword) {
  requireAdmin_(token, leaderId);
  const nextPassword = String(newPassword || "");
  if (nextPassword.length < 6) {
    return { ok: false, message: "كلمة المرور الجديدة يجب أن تكون 6 أحرف على الأقل" };
  }

  updateLeaderRegistryField_(leaderId, "admin_pass_hash", sha256_(nextPassword));
  return { ok: true };
}

function api_adminGetConfirmationEmailTemplate(leaderId, token) {
  requireAdmin_(token, leaderId);
  const { settingsSh } = getLeaderSheets_(leaderId);
  return {
    ok: true,
    template: getConfirmationEmailTemplate_(settingsSh),
    placeholders: [
      "{{employee_name}}",
      "{{branch_code}}",
      "{{region_name}}",
      "{{appointment_date}}",
      "{{booked_at}}",
      "{{appointment_time}}",
      "{{employee_id}}",
      "{{location_line}}",
      "{{cancel_link}}",
      "{{email_signature}}",
    ],
  };
}

function api_adminUpdateConfirmationEmailTemplate(leaderId, token, templateText) {
  requireAdmin_(token, leaderId);
  const template = String(templateText || "").trim();
  if (!template) return { ok: false, message: "نص رسالة التأكيد مطلوب" };

  const { settingsSh } = getLeaderSheets_(leaderId);
  setSettingValue_(settingsSh, EMAIL_SETTINGS_KEYS.CONFIRMATION_TEMPLATE, template);
  return { ok: true, template };
}

/* ===================== ADMIN: MEET URL (Regions) ===================== */
function api_adminSetMeetingLocation(leaderId, token, regionId, locationValue) {
  requireAdmin_(token, leaderId);

  regionId = norm_(regionId);
  if (!regionId) return { ok: false, message: "regionId مطلوب" };

  locationValue = norm_(locationValue);

  const { regionsSh } = getLeaderSheets_(leaderId);
  if (!regionsSh) return { ok: false, message: "Sheet 'Regions' غير موجودة" };

  const data = regionsSh.getDataRange().getValues();
  if (data.length < 2) return { ok: false, message: "Regions sheet فاضية" };

  const headers = data[0].map(h => String(h).trim());
  const idxId = idxOfHeader_(headers, "region_id");
  const idxMeet = idxOfHeader_(headers, "meet_url");

  if (idxId === -1 || idxMeet === -1) {
    return { ok: false, message: "الأعمدة المطلوبة: region_id و meet_url" };
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxId]).trim() === regionId) {
      regionsSh.getRange(i + 1, idxMeet + 1).setValue(locationValue);
      return { ok: true };
    }
  }

  return { ok: false, message: "region_id غير موجود: " + regionId };
}

/* ===================== ADMIN: SLOTS ===================== */
function api_adminGetSlots(leaderId, token, regionId) {
  requireAdmin_(token, leaderId);

  const { slotsSh } = getLeaderSheets_(leaderId);
  const lastRow = slotsSh.getLastRow();
  if (lastRow < 2) return { ok: true, slots: [] };

  regionId = norm_(regionId);
  if (regionId === "ALL") regionId = "";

  const tz = getTimeZoneForLeader_(leaderId) || DEFAULT_TZ;

  // columns: A slot_id, B region_id, C date, D timeText, E active
  const values = slotsSh.getRange(2, 1, lastRow - 1, 5).getValues();

  let slots = values.map(r => ({
    slot_id: norm_(r[0]),
    region_id: norm_(r[1]),
    date: toYMD_(r[2], tz),
    time: norm_(r[3]),
    active: truthy_(r[4]),
  }));

  if (regionId) slots = slots.filter(s => s.region_id === regionId);

  return { ok: true, slots };
}

function api_adminAddSlot(leaderId, token, regionId, dateStr, timeText) {
  requireAdmin_(token, leaderId);

  const { slotsSh } = getLeaderSheets_(leaderId);
  const slotId = newId_();

  regionId = norm_(regionId);
  dateStr = norm_(dateStr);     // expected YYYY-MM-DD
  timeText = norm_(timeText);

  if (!regionId) return { ok: false, message: "regionId required" };
  if (!dateStr) return { ok: false, message: "date required" };
  if (!timeText) return { ok: false, message: "time required" };

  const slotDate = toYMD_(dateStr, DEFAULT_TZ);
  const todayYmd = toYMD_(new Date(), DEFAULT_TZ);
  if (slotDate < todayYmd) return { ok: false, message: "Past dates are not allowed" };

  const conflict = findActiveSlotConflict_(slotsSh, slotDate, timeText, "");
  if (conflict) {
    return {
      ok: false,
      message: `يوجد تعارض: هناك موعد نشط آخر متداخل مع الوقت ${normalizeSlotTimeText_(timeText)} بتاريخ ${slotDate}`,
      conflict
    };
  }

  slotsSh.appendRow([slotId, regionId, dateStr, timeText, true]);
  return { ok: true, slot_id: slotId };
}

function normalizeSlotTimeText_(timeText) {
  let s = norm_(timeText);
  s = s.replace(/[\u200e\u200f\u202a-\u202e\u2066-\u2069]/g, "");
  s = s.replace(/\s+/g, " ");
  s = s.replace(/\s*([صم])\s*/g, " $1");
  s = s.replace(/\s*(AM|PM)\s*/gi, " $1");
  s = s.replace(/\s*-\s*/g, " - ");
  return s.trim();
}

function parseArabicTimePartToMinutes_(timePart) {
  const s = normalizeSlotTimeText_(timePart).replace(/\s+/g, "");
  const m = s.match(/^(\d{1,2}):(\d{2})(ص|م|AM|PM)$/i);
  if (!m) return null;

  let hh = Number(m[1]) % 12;
  const mm = Number(m[2]);
  const marker = String(m[3] || "").toUpperCase();
  if (marker === "م" || marker === "PM") hh += 12;
  return hh * 60 + mm;
}

function parseSlotRangeMinutes_(timeText) {
  const parts = normalizeSlotTimeText_(timeText).split(" - ");
  if (parts.length !== 2) return null;

  const startMin = parseArabicTimePartToMinutes_(parts[0]);
  const endMin = parseArabicTimePartToMinutes_(parts[1]);
  if (startMin == null || endMin == null) return null;

  return { startMin, endMin };
}

function slotRangesOverlap_(timeA, timeB) {
  const a = parseSlotRangeMinutes_(timeA);
  const b = parseSlotRangeMinutes_(timeB);
  if (!a || !b) return false;
  return a.startMin < b.endMin && b.startMin < a.endMin;
}

function isSlotExpiredAtNow_(dateStr, timeText, tz) {
  const slotDate = toYMD_(dateStr, tz || DEFAULT_TZ);
  if (!slotDate) return false;

  const now = new Date();
  const nowYmd = Utilities.formatDate(now, tz || DEFAULT_TZ, "yyyy-MM-dd");
  if (slotDate < nowYmd) return true;
  if (slotDate > nowYmd) return false;

  const range = parseSlotRangeMinutes_(timeText);
  if (!range) return false;

  const nowMinutes = Number(Utilities.formatDate(now, tz || DEFAULT_TZ, "H")) * 60
    + Number(Utilities.formatDate(now, tz || DEFAULT_TZ, "m"));

  return range.endMin <= nowMinutes;
}

function findActiveSlotConflict_(slotsSh, dateStr, timeText, excludeSlotId) {
  const lastRow = slotsSh.getLastRow();
  if (lastRow < 2) return null;

  const targetDate = toYMD_(dateStr, DEFAULT_TZ);
  const targetTime = normalizeSlotTimeText_(timeText);
  const skipSlotId = norm_(excludeSlotId);
  const values = slotsSh.getRange(2, 1, lastRow - 1, 5).getValues();

  for (let i = 0; i < values.length; i++) {
    const slotId = norm_(values[i][0]);
    const regionId = norm_(values[i][1]);
    const slotDate = toYMD_(values[i][2], DEFAULT_TZ);
    const slotTime = normalizeSlotTimeText_(values[i][3]);
    const active = truthy_(values[i][4]);

    if (!active) continue;
    if (skipSlotId && slotId === skipSlotId) continue;
    if (slotDate !== targetDate) continue;
    if (!slotRangesOverlap_(slotTime, targetTime)) continue;

    return {
      slot_id: slotId,
      region_id: regionId,
      date: slotDate,
      time: slotTime
    };
  }

  return null;
}

function minutesToArabicTimeText_(mins) {
  mins = Number(mins);
  if (isNaN(mins)) return "";
  mins = ((mins % (24 * 60)) + (24 * 60)) % (24 * 60);
  let hh = Math.floor(mins / 60);
  const mm = mins % 60;
  const suffix = hh >= 12 ? "م" : "ص";
  let hh12 = hh % 12;
  if (hh12 === 0) hh12 = 12;
  return `${hh12}:${String(mm).padStart(2, "0")} ${suffix}`;
}

function api_adminAddSlotsBulk(leaderId, token, regionId, dateStr, startMinutes, endMinutes, durationMinutes, restMinutes) {
  requireAdmin_(token, leaderId);

  const { slotsSh } = getLeaderSheets_(leaderId);
  regionId = norm_(regionId);
  dateStr = norm_(dateStr);
  startMinutes = Number(startMinutes);
  endMinutes = Number(endMinutes);
  durationMinutes = Number(durationMinutes);
  restMinutes = Number(restMinutes || 0);

  if (!regionId) return { ok: false, message: "regionId required" };
  if (!dateStr) return { ok: false, message: "date required" };
  if (isNaN(startMinutes) || isNaN(endMinutes)) return { ok: false, message: "Invalid time range" };
  if (isNaN(durationMinutes) || durationMinutes <= 0) return { ok: false, message: "Invalid duration" };
  if (isNaN(restMinutes) || restMinutes < 0) return { ok: false, message: "Invalid rest minutes" };
  if (endMinutes <= startMinutes) return { ok: false, message: "End time must be after begin time" };

  const lastRow = slotsSh.getLastRow();
  const existing = lastRow >= 2
    ? slotsSh.getRange(2, 1, lastRow - 1, 5).getValues()
    : [];

  const out = [];
  const slotDate = toYMD_(dateStr, DEFAULT_TZ);
  const todayYmd = toYMD_(new Date(), DEFAULT_TZ);
  if (slotDate < todayYmd) return { ok: false, message: "Past dates are not allowed" };

  let cur = startMinutes;
  let guard = 0;
  let skipped = 0;
  while ((cur + durationMinutes) <= endMinutes && guard < 500) {
    const timeText = `${minutesToArabicTimeText_(cur)} - ${minutesToArabicTimeText_(cur + durationMinutes)}`;
    const conflict = existing.find((r) => {
      const active = truthy_(r[4]);
      const sameDate = toYMD_(r[2], DEFAULT_TZ) === slotDate;
      const slotTime = norm_(r[3]);
      return active && sameDate && slotRangesOverlap_(slotTime, timeText);
    });
    const duplicatePending = out.find((r) => slotRangesOverlap_(r[3], timeText));

    if (!conflict && !duplicatePending) {
      out.push([newId_(), regionId, dateStr, timeText, true]);
    } else {
      skipped++;
    }
    cur += durationMinutes + restMinutes;
    guard++;
  }

  if (guard >= 500) return { ok: false, message: "Too many generated slots" };
  if (!out.length) return { ok: true, added: 0, skipped };

  slotsSh.getRange(slotsSh.getLastRow() + 1, 1, out.length, 5).setValues(out);
  return { ok: true, added: out.length, skipped };
}

function api_adminUpdateSlotById(leaderId, token, slotId, patch) {
  requireAdmin_(token, leaderId);

  const { slotsSh } = getLeaderSheets_(leaderId);
  const lastRow = slotsSh.getLastRow();
  if (lastRow < 2) return { ok: false, message: "No slots" };

  slotId = norm_(slotId);
  if (!slotId) return { ok: false, message: "slotId required" };

  const ids = slotsSh.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(x => norm_(x));
  const idx = ids.indexOf(slotId);
  if (idx === -1) return { ok: false, message: "Slot not found" };

  const row = idx + 2;
  patch = patch || {};

  const current = slotsSh.getRange(row, 1, 1, 5).getValues()[0];
  const nextRegionId = ("region_id" in patch) ? norm_(patch.region_id) : norm_(current[1]);
  const nextDate = ("date" in patch) ? toYMD_(patch.date, DEFAULT_TZ) : toYMD_(current[2], DEFAULT_TZ);
  const nextTime = ("time" in patch || "timeText" in patch)
    ? normalizeSlotTimeText_(patch.time ?? patch.timeText)
    : normalizeSlotTimeText_(current[3]);
  const nextActive = ("active" in patch) ? !!patch.active : truthy_(current[4]);

  if (nextActive) {
    const conflict = findActiveSlotConflict_(slotsSh, nextDate, nextTime, slotId);
    if (conflict) {
      return {
        ok: false,
        message: `يوجد تعارض: هناك موعد نشط آخر متداخل مع الوقت ${nextTime} بتاريخ ${nextDate}`,
        conflict
      };
    }
  }

  slotsSh.getRange(row, 1, 1, 5).setValues([[
    slotId,
    nextRegionId,
    nextDate,
    nextTime,
    nextActive
  ]]);

  return { ok: true };
}

function api_adminDeleteSlotById(leaderId, token, slotId) {
  requireAdmin_(token, leaderId);

  const { slotsSh } = getLeaderSheets_(leaderId);
  const lastRow = slotsSh.getLastRow();
  if (lastRow < 2) return { ok: false, message: "No slots" };

  slotId = norm_(slotId);
  if (!slotId) return { ok: false, message: "slotId required" };

  const ids = slotsSh.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(x => norm_(x));
  const idx = ids.indexOf(slotId);
  if (idx === -1) return { ok: false, message: "Slot not found" };

  slotsSh.deleteRow(idx + 2);
  return { ok: true };
}

// Convenience: active-only update
function api_adminUpdateSlotActiveById(leaderId, token, slotId, active) {
  requireAdmin_(token, leaderId);
  return api_adminUpdateSlotById(leaderId, token, slotId, { active: !!active });
}

// Backward compatibility aliases (if older admin2 uses these names)
function api_adminUpdateSlot(leaderId, token, slotId, patch) {
  return api_adminUpdateSlotById(leaderId, token, slotId, patch);
}
function api_adminDeleteSlot(leaderId, token, slotId) {
  return api_adminDeleteSlotById(leaderId, token, slotId);
}

function api_adminDeleteSlotsBulk(leaderId, token, slotIds) {
  requireAdmin_(token, leaderId);

  const idsToDelete = Array.isArray(slotIds)
    ? Array.from(new Set(slotIds.map(norm_).filter(Boolean)))
    : [];
  if (!idsToDelete.length) return { ok: false, message: "slotIds required" };

  const { slotsSh } = getLeaderSheets_(leaderId);
  const lastRow = slotsSh.getLastRow();
  if (lastRow < 2) return { ok: false, message: "No slots" };

  const ids = slotsSh.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(x => norm_(x));
  const rowsToDelete = [];

  idsToDelete.forEach((slotId) => {
    const idx = ids.indexOf(slotId);
    if (idx !== -1) rowsToDelete.push(idx + 2);
  });

  if (!rowsToDelete.length) return { ok: false, message: "No matching slots found" };

  rowsToDelete.sort((a, b) => b - a).forEach((rowIndex) => slotsSh.deleteRow(rowIndex));
  return { ok: true, deleted: rowsToDelete.length };
}

function api_adminUpdateSlotsBulk(leaderId, token, slotIds, patch) {
  requireAdmin_(token, leaderId);

  const idsToUpdate = Array.isArray(slotIds)
    ? Array.from(new Set(slotIds.map(norm_).filter(Boolean)))
    : [];
  if (!idsToUpdate.length) return { ok: false, message: "slotIds required" };

  patch = patch || {};
  if (!("active" in patch)) return { ok: false, message: "No supported patch provided" };

  const { slotsSh } = getLeaderSheets_(leaderId);
  const lastRow = slotsSh.getLastRow();
  if (lastRow < 2) return { ok: false, message: "No slots" };

  const ids = slotsSh.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(x => norm_(x));
  const nextActive = !!patch.active;
  let updated = 0;

  idsToUpdate.forEach((slotId) => {
    const idx = ids.indexOf(slotId);
    if (idx === -1) return;
    slotsSh.getRange(idx + 2, 5).setValue(nextActive);
    updated++;
  });

  return { ok: true, updated };
}

function getReservationCols_(reservationsSh) {
  const lastCol = Math.max(18, reservationsSh.getLastColumn());
  const headerRow = reservationsSh.getRange(1, 1, 1, lastCol).getValues()[0];

  function pickIdx_(candidates, fallbackIdx) {
    for (let i = 0; i < candidates.length; i++) {
      const idx = idxOfHeader_(headerRow, candidates[i]);
      if (idx !== -1) return idx;
    }
    return fallbackIdx;
  }

  return {
    lastCol,
    cName: pickIdx_(["name"], 0),
    cDate: pickIdx_(["date", "day"], 1),
    cTime: pickIdx_(["timeText", "time_text", "apptText", "appointment_text"], 2),
    cCreatedAt: pickIdx_(["createdAt", "created_at"], 3),
    cEmail: pickIdx_(["email"], 4),
    cReservationId: pickIdx_(["reservationId", "reservation_id"], 5),
    cEmployeeId: pickIdx_(["employeeId", "employee_id"], 6),
    cBranch: pickIdx_(["branch", "pharmacy", "branch_code"], 7),
    cRegion: pickIdx_(["region", "location", "region_id"], 8),
    cSlotId: pickIdx_(["slot_id", "slotId"], 9),
    cStatus: pickIdx_(["status"], 10),
    cCancelledAt: pickIdx_(["cancelledAt", "cancelled_at"], 11),
    cCancelledBy: pickIdx_(["cancelled_by", "cancelledBy"], 12),
    cRescheduledFrom: pickIdx_(["rescheduled_from", "rescheduledFrom"], 13),
    cOldSlotReleased: pickIdx_(["old_slot_released", "oldSlotReleased"], 14),
    cCancellationReason: pickIdx_(["cancellation_reason", "cancel_reason", "reason"], 15),
    cFeedbackSentAt: pickIdx_(["feedback_sent_at", "feedbackSentAt"], 16),
    cFeedbackMessage: pickIdx_(["feedback_message", "feedbackMessage"], 17)
  };
}

function findReservationRowById_(reservationsSh, cols, reservationId) {
  const rid = norm_(reservationId);
  const lastRow = reservationsSh.getLastRow();
  if (!rid || lastRow < 2) return null;

  const values = reservationsSh.getRange(2, 1, lastRow - 1, cols.lastCol).getValues();
  for (let i = 0; i < values.length; i++) {
    if (norm_(values[i][cols.cReservationId]) === rid) {
      return { rowIndex: i + 2, row: values[i] };
    }
  }
  return null;
}

function normalizeDigits_(s) {
  return String(s || "")
    .replace(/[\u0660-\u0669]/g, d => String(d.charCodeAt(0) - 0x0660))
    .replace(/[\u06F0-\u06F9]/g, d => String(d.charCodeAt(0) - 0x06F0));
}

function timeTextToMinutes_(timeText) {
  let s = normalizeDigits_(timeText).trim().toLowerCase();
  if (!s) return null;

  s = s.replace(/\u0635/g, 'am').replace(/\u0645/g, 'pm');
  s = s.replace(/a\.?m\.?/g, 'am').replace(/p\.?m\.?/g, 'pm');

  const rangeParts = s.split(/\s*[-\u2013\u2014]\s*/);
  s = (rangeParts[0] || '').trim();

  const m = s.match(/(\d{1,2})(?::|\.)(\d{2})\s*(am|pm)?/i);
  if (!m) return null;

  let hh = Number(m[1]);
  const mm = Number(m[2]);
  const ampm = (m[3] || '').toLowerCase();

  if (ampm === 'pm' && hh < 12) hh += 12;
  if (ampm === 'am' && hh === 12) hh = 0;
  if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;

  return hh * 60 + mm;
}


function reservationStartsAtMs_(dateStr, timeText, tz) {
  const ymd = toYMD_(dateStr, tz);
  const m = String(ymd || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;

  const mins = timeTextToMinutes_(timeText);
  const year = Number(m[1]);
  const month = Number(m[2]) - 1;
  const day = Number(m[3]);

  if (mins == null) {
    return new Date(year, month, day, 0, 0, 0, 0).getTime();
  }

  const hh = Math.floor(mins / 60);
  const mm = mins % 60;
  return new Date(year, month, day, hh, mm, 0, 0).getTime();
}

function isPastReservationRow_(row, cols, tz) {
  const status = norm_(row[cols.cStatus] || '').toLowerCase() || 'upcoming';
  if (status !== 'upcoming') return false;

  const startsAt = reservationStartsAtMs_(row[cols.cDate], row[cols.cTime], tz);
  if (startsAt == null) return false;

  return startsAt < Date.now();
}

function buildReservationRow_(cols, data) {
  const row = new Array(cols.lastCol).fill('');
  row[cols.cName] = norm_(data.name);
  row[cols.cDate] = norm_(data.date);
  row[cols.cTime] = norm_(data.timeText);
  row[cols.cCreatedAt] = data.createdAt || new Date();
  row[cols.cEmail] = norm_(data.email);
  row[cols.cReservationId] = norm_(data.reservationId);
  row[cols.cEmployeeId] = norm_(data.employeeId);
  row[cols.cBranch] = norm_(data.branch);
  row[cols.cRegion] = norm_(data.region);
  row[cols.cSlotId] = norm_(data.slotId);
  row[cols.cStatus] = norm_(data.status || 'upcoming') || 'upcoming';
  row[cols.cCancelledAt] = data.cancelledAt || '';
  row[cols.cCancelledBy] = norm_(data.cancelled_by);
  row[cols.cRescheduledFrom] = norm_(data.rescheduled_from);
  row[cols.cOldSlotReleased] = !!data.old_slot_released;
  row[cols.cCancellationReason] = norm_(data.cancellation_reason);
  row[cols.cFeedbackSentAt] = data.feedback_sent_at || '';
  row[cols.cFeedbackMessage] = norm_(data.feedback_message);
  return row;
}

function htmlEscape_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function brandedEmailHtml_(opts) {
  const direction = String((opts && opts.direction) || "rtl").toLowerCase() === "ltr" ? "ltr" : "rtl";
  const align = direction === "ltr" ? "left" : "right";
  const title = htmlEscape_(opts && opts.title);
  const eyebrow = htmlEscape_((opts && opts.eyebrow) || "Day For Development");
  const introHtml = opts && opts.introHtml ? `<div style="font-size:15px;line-height:1.9;color:#1f2937;margin-bottom:14px">${opts.introHtml}</div>` : "";
  const messageHtml = opts && opts.messageHtml ? `<div style="font-size:14px;line-height:1.95;color:#334155">${opts.messageHtml}</div>` : "";
  const noteHtml = opts && opts.note ? `<div style="margin-top:18px;padding:12px 14px;border-radius:14px;background:#f8fafc;border:1px solid #e2e8f0;font-size:12px;line-height:1.8;color:#64748b">${htmlEscape_(opts.note)}</div>` : "";
  const actionHtml = opts && opts.actionUrl
    ? `<div style="text-align:center;margin-top:22px">
         <a href="${htmlEscape_(opts.actionUrl)}" target="_blank" rel="noopener noreferrer"
           style="display:inline-block;padding:12px 22px;border-radius:12px;background:#0b3f77;color:#ffffff;text-decoration:none;font-weight:800;font-size:14px">
           ${htmlEscape_((opts && opts.actionLabel) || "Open")}
         </a>
       </div>`
    : "";

  return `
    <div style="margin:0;padding:24px;background:#eef4fb;font-family:'Tajawal',Arial,sans-serif;direction:${direction};text-align:${align}">
      <div style="max-width:640px;margin:0 auto;background:#ffffff;border:1px solid #dbe7f3;border-radius:22px;overflow:hidden;box-shadow:0 18px 40px rgba(15,23,42,.10)">
        <div style="padding:18px 24px;background:linear-gradient(135deg,#0b3f77,#0f5ca8);color:#ffffff">
          <div style="font-size:12px;letter-spacing:.4px;opacity:.9">${eyebrow}</div>
          <div style="font-size:24px;font-weight:800;margin-top:6px">${title}</div>
        </div>
        <div style="padding:24px">
          ${introHtml}
          ${messageHtml}
          ${actionHtml}
          ${noteHtml}
        </div>
      </div>
    </div>
  `.trim();
}

function confirmationEmailHtml_(subject, bodyText) {
  return brandedEmailHtml_({
    direction: "rtl",
    eyebrow: "Day For Development",
    title: subject,
    messageHtml: htmlEscape_(bodyText).replace(/\n/g, "<br>")
  });
}

function sendAdminAppointmentEmail_(leaderId, payload) {
  const email = norm_(payload && payload.email);
  if (!email) return;

  const type = norm_(payload && payload.type).toLowerCase();
  if (type !== 'cancelled' && type !== 'rescheduled') return;

  const name = norm_(payload && payload.name) || '—';
  const employeeId = norm_(payload && payload.employeeId) || '—';
  const branch = norm_(payload && payload.branch) || '—';
  const dateStr = norm_(payload && payload.dateStr) || '—';
  const timeText = norm_(payload && payload.timeText) || '—';
  const regionId = norm_(payload && payload.regionId);

  const region = regionId ? getRegionInfo_(leaderId, regionId) : null;
  const regionName = region && region.region_name ? region.region_name : (regionId || '—');

  const leaderInfo = getLeaderInfoById_(leaderId);
  const supervisorEmail = leaderInfo && leaderInfo.admin_email ? leaderInfo.admin_email : '';

  const subject = type === 'cancelled'
    ? 'إشعار بإلغاء موعد التقييم'
    : 'إشعار بتغيير موعد التقييم';

  const intro = `الزميل رائد الرعاية المتميزة د/ ${name}`;
  const baseLine = type === 'cancelled'
    ? 'نفيدكم علماً بأنه تم إلغاء موعد التقييم، ونأمل منكم التواصل مع المشرف في حال وجود أي استفسار.'
    : 'نفيدكم علماً بأنه تم تغيير موعد التقييم، ونأمل منكم التواصل مع المشرف في حال وجود أي استفسار.';

  let body = [intro, '', baseLine, ''];
  let htmlExtra = '';
  let actionLabel = '';
  let actionUrl = '';

  if (type === 'cancelled') {
    body = body.concat([
      'يمكنكم حجز موعد آخر من خلال صفحة الحجز التالية:',
      RESERVATION_URL,
      '',
      'مع خالص الشكر والتقدير'
    ]);

    htmlExtra = `<div style="margin-top:14px">يمكنكم حجز موعد آخر من خلال صفحة الحجز التالية.</div>`;
    actionLabel = 'حجز موعد جديد';
    actionUrl = RESERVATION_URL;
  } else {
    const newBranch = norm_(payload && payload.newBranch) || branch;
    const newDateStr = norm_(payload && payload.newDateStr) || dateStr;
    const newTimeText = norm_(payload && payload.newTimeText) || timeText;
    const newRegionId = norm_(payload && payload.newRegionId) || regionId;
    const newRegion = newRegionId ? getRegionInfo_(leaderId, newRegionId) : null;
    const newRegionName = newRegion && newRegion.region_name ? newRegion.region_name : (newRegionId || '—');

    body = body.concat([
      'الموعد الجديد كما يلي:',
      `الفرع: ${newBranch}`,
      `الموقع: ${newRegionName}`,
      `اليوم: ${newDateStr}`,
      `الوقت: ${newTimeText}`,
      `الرقم الوظيفي: ${employeeId}`,
      '',
      'مع خالص الشكر والتقدير'
    ]);

    htmlExtra = `
      <div style="background:#f8fbff;border-radius:16px;padding:16px;margin:16px 0;font-size:14px;line-height:1.9;border:1px solid #dbe7f3">
        <b>الموعد الجديد كما يلي:</b><br>
        🏪 <b>الفرع:</b> ${htmlEscape_(newBranch)}<br>
        📍 <b>الموقع:</b> ${htmlEscape_(newRegionName)}<br>
        🗓 <b>اليوم:</b> ${htmlEscape_(newDateStr)}<br>
        ⏰ <b>الوقت:</b> ${htmlEscape_(newTimeText)}<br>
        🆔 <b>الرقم الوظيفي:</b> ${htmlEscape_(employeeId)}<br>
      </div>`;
  }

  const bodyText = body.join('\n').trim();
  const htmlBody = brandedEmailHtml_({
    direction: 'rtl',
    eyebrow: 'Day For Development',
    title: type === 'cancelled' ? 'إلغاء موعد التقييم' : 'تغيير موعد التقييم',
    introHtml: htmlEscape_(intro),
    messageHtml: `<div>${htmlEscape_(baseLine)}</div>${htmlExtra}<div style="margin-top:18px;color:#475569">مع خالص الشكر والتقدير</div>`,
    actionLabel,
    actionUrl
  });

  const mailOptions = {
    to: email,
    subject,
    body: bodyText,
    htmlBody,
    name: 'ادارة المنطقة الجنوبية | Day For Development '
  };
  if (supervisorEmail) mailOptions.cc = supervisorEmail;

  try {
    MailApp.sendEmail(mailOptions);
  } catch (richErr) {
    Logger.log('Rich admin appointment email failed: ' + richErr);
    MailApp.sendEmail(email, subject, bodyText);
  }
}

function api_adminSendFeedback(leaderId, token, reservationIds, messageText) {
  requireAdmin_(token, leaderId);

  const ids = Array.isArray(reservationIds)
    ? Array.from(new Set(reservationIds.map(norm_).filter(Boolean)))
    : [];
  const message = norm_(messageText);

  if (!ids.length) return { ok: false, message: "reservationIds required" };
  if (!message) return { ok: false, message: "messageText required" };

  const { reservationsSh } = getLeaderSheets_(leaderId);
  const leaderInfo = getLeaderInfoById_(leaderId) || {};
  const { settingsSh } = getLeaderSheets_(leaderId);
  const savedSignature = norm_(getSettingValue_(settingsSh, "EMAIL_SIGNATURE"));
  const supervisorName = norm_(leaderInfo.leader_name) || "د. اسم المشرف";
  const signatureName = savedSignature || supervisorName;
  const cols = getReservationCols_(reservationsSh);
  const lastRow = reservationsSh.getLastRow();
  if (lastRow < 2) return { ok: false, message: "No reservations found" };

  const values = reservationsSh.getRange(2, 1, lastRow - 1, cols.lastCol).getValues();
  const rowById = new Map();
  values.forEach((row, idx) => {
    const rid = norm_(row[cols.cReservationId]);
    if (rid) rowById.set(rid, { row, rowIndex: idx + 2 });
  });

  const subject = "رسالة شكر بعد الموعد";
  let sent = 0;
  const skipped = [];

  ids.forEach((rid) => {
    const found = rowById.get(rid);
    if (!found) {
      skipped.push({ reservationId: rid, reason: "Reservation not found" });
      return;
    }
    const row = found.row;

    const email = norm_(row[cols.cEmail]);
    if (!email) {
      skipped.push({ reservationId: rid, reason: "Missing email" });
      return;
    }

    const name = norm_(row[cols.cName]) || "الزميل/ة";
    const intro = `الزميل رائد الرعاية المتميزة د/ ${name}،`;
    const closing = `مع خالص التحية،\n${signatureName}\nالمنطقة الجنوبية`;
    const bodyText = `${intro}\n\n${message}\n\n${closing}`;
    const htmlBody = `
      <div style="font-family:'Tajawal',Arial,sans-serif;direction:rtl;text-align:right;line-height:1.9;color:#1f2937">
        <p>${htmlEscape_(intro)}</p>
        <p>${htmlEscape_(message).replace(/\n/g, "<br>")}</p>
        <p>${htmlEscape_(closing).replace(/\n/g, "<br>")}</p>
      </div>
    `;

    try {
      MailApp.sendEmail({
        to: email,
        subject,
        body: bodyText,
        htmlBody,
        name: `${signatureName} - المنطقة الجنوبية`
      });
      reservationsSh.getRange(found.rowIndex, cols.cFeedbackSentAt + 1).setValue(new Date());
      reservationsSh.getRange(found.rowIndex, cols.cFeedbackMessage + 1).setValue(message);
      sent++;
    } catch (err) {
      skipped.push({ reservationId: rid, reason: String(err && err.message || err) });
    }
  });

  return { ok: true, sent, skipped };
}

function api_adminListFeedbackHistory(leaderId, token, limit) {
  requireAdmin_(token, leaderId);

  const { reservationsSh } = getLeaderSheets_(leaderId);
  const cols = getReservationCols_(reservationsSh);
  const lastRow = reservationsSh.getLastRow();
  if (lastRow < 2) return { ok: true, rows: [] };

  const tz = getTimeZoneForLeader_(leaderId) || DEFAULT_TZ;
  const maxItems = Math.max(1, Math.min(Number(limit || 50), 200));
  const values = reservationsSh.getRange(2, 1, lastRow - 1, cols.lastCol).getValues();
  const rows = [];

  values.forEach((r) => {
    const sentAt = r[cols.cFeedbackSentAt];
    const sentDate = sentAt ? new Date(sentAt) : null;
    if (!sentDate || isNaN(sentDate.getTime())) return;

    const regionId = norm_(r[cols.cRegion]);
    const region = regionId ? getRegionInfo_(leaderId, regionId) : null;

    rows.push({
      reservationId: norm_(r[cols.cReservationId]),
      name: norm_(r[cols.cName]) || '—',
      employeeId: norm_(r[cols.cEmployeeId]) || '—',
      branch: norm_(r[cols.cBranch]) || '—',
      location: region && region.region_name ? region.region_name : (regionId || '—'),
      date: toYMD_(r[cols.cDate], tz) || norm_(r[cols.cDate]),
      timeText: norm_(r[cols.cTime]) || '—',
      message: norm_(r[cols.cFeedbackMessage]),
      sentAtText: Utilities.formatDate(sentDate, tz, 'yyyy-MM-dd HH:mm'),
      sentAtMs: sentDate.getTime()
    });
  });

  rows.sort((a, b) => b.sentAtMs - a.sentAtMs);
  return { ok: true, rows: rows.slice(0, maxItems) };
}

function api_adminExportReservations(leaderId, token) {
  requireAdmin_(token, leaderId);

  const { reservationsSh } = getLeaderSheets_(leaderId);
  const data = reservationsSh.getDataRange().getValues();
  if (!data || !data.length) {
    return { ok: true, filename: `reservations_${leaderId}.csv`, headers: [], rows: [] };
  }

  const tz = getTimeZoneForLeader_(leaderId) || DEFAULT_TZ;
  const headers = (data[0] || []).map(v => String(v || "").trim());
  const rows = data.slice(1).map((row) => row.map((cell) => {
    if (cell instanceof Date && !isNaN(cell.getTime())) {
      return Utilities.formatDate(cell, tz, "yyyy-MM-dd HH:mm:ss");
    }
    return String(cell ?? "");
  }));

  const stamp = Utilities.formatDate(new Date(), tz, "yyyyMMdd_HHmm");
  return {
    ok: true,
    filename: `reservations_${leaderId}_${stamp}.csv`,
    headers,
    rows
  };
}

function api_adminBootstrap(leaderId, token) {
  requireAdmin_(token, leaderId);
  const { reservationsSh, settingsSh, slotsSh, regionsSh, branchesSh } = getLeaderSheets_(leaderId);
  const tz = getSettingValue_(settingsSh, "TIMEZONE") || DEFAULT_TZ;

  const regionsData = regionsSh.getDataRange().getValues();
  const regionRows = [];
  const regionNameMap = {};
  if (regionsData && regionsData.length >= 2) {
    const headers = regionsData[0].map(h => String(h || "").trim());
    const hNorm = headers.map(h => h.replace(/\s+/g, " ").toLowerCase());
    const idxId = hNorm.indexOf("region_id");
    const idxName = hNorm.indexOf("region_name");
    const idxType = hNorm.indexOf("region_type");
    const idxMeet = hNorm.indexOf("meet_url");
    const idxMap = hNorm.indexOf("map_url");
    const idxStat = hNorm.indexOf("status");

    const toBoolStatus_ = (v) => {
      if (v === true) return true;
      if (v === false) return false;
      const s = String(v ?? "").trim().toUpperCase();
      if (s === "TRUE" || s === "1" || s === "YES") return true;
      if (s === "FALSE" || s === "0" || s === "NO" || s === "") return false;
      return false;
    };

    const normType_ = (v) => {
      const t = String(v ?? "").trim().toUpperCase();
      if (t === "ONLINE") return "ONLINE";
      if (t === "IN_PERSON" || t === "INPERSON") return "IN_PERSON";
      return "";
    };

    for (let i = 1; i < regionsData.length; i++) {
      const r = regionsData[i] || [];
      const id = idxId >= 0 ? String(r[idxId] || "").trim() : "";
      if (!id) continue;
      const statusRaw = idxStat >= 0 ? r[idxStat] : true;
      const statusStr = String(statusRaw ?? "").trim().toUpperCase();
      if (statusStr === "DELETED") continue;

      const row = {
        region_id: id,
        region_name: idxName >= 0 ? String(r[idxName] || "").trim() : "",
        region_type: idxType >= 0 ? normType_(r[idxType]) : "",
        meet_url: idxMeet >= 0 ? String(r[idxMeet] || "").trim() : "",
        map_url: idxMap >= 0 ? String(r[idxMap] || "").trim() : "",
        status: toBoolStatus_(statusRaw)
      };
      regionRows.push(row);
      regionNameMap[id] = row.region_name || id;
    }
  }

  const branchesData = branchesSh.getDataRange().getValues();
  let branchRows = [];
  if (branchesData && branchesData.length >= 2) {
    const headers = branchesData[0].map(h => String(h || "").trim());
    branchRows = branchesData.slice(1)
      .map(r => {
        const o = {};
        headers.forEach((k, i) => o[k] = r[i]);
        return o;
      })
      .filter(x => String(x.branch_code || "").trim());
  }

  const slotsLastRow = slotsSh.getLastRow();
  let slotRows = [];
  if (slotsLastRow >= 2) {
    const values = slotsSh.getRange(2, 1, slotsLastRow - 1, 5).getValues();
    slotRows = values.map(r => ({
      slot_id: norm_(r[0]),
      region_id: norm_(r[1]),
      date: toYMD_(r[2], tz),
      time: norm_(r[3]),
      active: truthy_(r[4]),
    }));
  }

  const cols = getReservationCols_(reservationsSh);
  const reservationsLastRow = reservationsSh.getLastRow();
  let reservationValues = [];
  if (reservationsLastRow >= 2) {
    reservationValues = reservationsSh.getRange(2, 1, reservationsLastRow - 1, cols.lastCol).getValues();
  }

  const appointmentRows = reservationValues
    .filter(r => (
      r[cols.cName] || r[cols.cDate] || r[cols.cTime] || r[cols.cEmail] ||
      r[cols.cEmployeeId] || r[cols.cBranch] || r[cols.cRegion] || r[cols.cReservationId]
    ))
    .map(r => ({
      reservationId: norm_(r[cols.cReservationId]),
      day: toYMD_(r[cols.cDate], tz),
      apptText: norm_(r[cols.cTime]),
      name: norm_(r[cols.cName]),
      employeeId: norm_(r[cols.cEmployeeId]),
      pharmacy: norm_(r[cols.cBranch]),
      location: norm_(r[cols.cRegion]),
      email: norm_(r[cols.cEmail]),
      status: norm_(r[cols.cStatus]) || "upcoming",
      cancelledAt: toYMD_(r[cols.cCancelledAt], tz) || norm_(r[cols.cCancelledAt]),
      cancelled_by: norm_(r[cols.cCancelledBy]),
      rescheduled_from: norm_(r[cols.cRescheduledFrom]),
      old_slot_released: truthy_(r[cols.cOldSlotReleased])
    }))
    .slice(-200)
    .reverse();

  const notificationRows = [];
  const feedbackRows = [];
  reservationValues.forEach((r) => {
    const name = norm_(r[cols.cName]) || '—';
    const employeeId = norm_(r[cols.cEmployeeId]) || '—';
    const branch = norm_(r[cols.cBranch]) || '—';
    const regionId = norm_(r[cols.cRegion]);
    const location = regionNameMap[regionId] || regionId || '—';
    const dateStr = toYMD_(r[cols.cDate], tz) || norm_(r[cols.cDate]);
    const timeText = norm_(r[cols.cTime]) || '—';
    const reservationId = norm_(r[cols.cReservationId]);

    const createdAt = r[cols.cCreatedAt];
    const createdDate = createdAt ? new Date(createdAt) : null;
    if (createdDate && !isNaN(createdDate.getTime())) {
      notificationRows.push({
        type: 'booking',
        atMs: createdDate.getTime(),
        atText: Utilities.formatDate(createdDate, tz, 'yyyy-MM-dd HH:mm'),
        reservationId,
        name,
        employeeId,
        branch,
        location,
        date: dateStr,
        timeText
      });
    }

    const status = norm_(r[cols.cStatus] || '').toLowerCase();
    const cancelledAt = r[cols.cCancelledAt];
    const cancelledDate = cancelledAt ? new Date(cancelledAt) : null;
    if (status === 'cancelled' && cancelledDate && !isNaN(cancelledDate.getTime())) {
      notificationRows.push({
        type: 'cancellation',
        atMs: cancelledDate.getTime(),
        atText: Utilities.formatDate(cancelledDate, tz, 'yyyy-MM-dd HH:mm'),
        reservationId,
        name,
        employeeId,
        branch,
        location,
        date: dateStr,
        timeText,
        cancellation_reason: norm_(r[cols.cCancellationReason])
      });
    }

    const sentAt = r[cols.cFeedbackSentAt];
    const sentDate = sentAt ? new Date(sentAt) : null;
    if (sentDate && !isNaN(sentDate.getTime())) {
      feedbackRows.push({
        reservationId,
        name,
        employeeId,
        branch,
        location,
        date: dateStr,
        timeText,
        message: norm_(r[cols.cFeedbackMessage]),
        sentAtText: Utilities.formatDate(sentDate, tz, 'yyyy-MM-dd HH:mm'),
        sentAtMs: sentDate.getTime()
      });
    }
  });

  notificationRows.sort((a, b) => b.atMs - a.atMs);
  feedbackRows.sort((a, b) => b.sentAtMs - a.sentAtMs);

  return {
    ok: true,
    appointments: { ok: true, rows: appointmentRows },
    slots: { ok: true, slots: slotRows },
    regions: { ok: true, rows: regionRows },
    branches: { ok: true, rows: branchRows },
    notifications: { ok: true, rows: notificationRows.slice(0, 100) },
    feedback_history: { ok: true, rows: feedbackRows.slice(0, 50) }
  };
}

function api_adminCancelAppointment(leaderId, token, reservationId, releaseOldSlot) {
  requireAdmin_(token, leaderId);

  let lock;
  try {
    lock = acquireLeaderLock_(leaderId, 15000);

    const { reservationsSh } = getLeaderSheets_(leaderId);
    const cols = getReservationCols_(reservationsSh);
    const found = findReservationRowById_(reservationsSh, cols, reservationId);
    if (!found) return { ok: false, message: 'Reservation not found' };

    const tz = getTimeZoneForLeader_(leaderId) || DEFAULT_TZ;
    const row = found.row;
    const rowIndex = found.rowIndex;
    const status = norm_(row[cols.cStatus] || '').toLowerCase() || 'upcoming';

    if (status === 'cancelled') return { ok: false, message: 'Appointment already cancelled' };
    if (status === 'rescheduled') return { ok: false, message: 'Rescheduled appointments cannot be cancelled' };
    if (isPastReservationRow_(row, cols, tz)) return { ok: false, message: 'Past appointments cannot be cancelled' };

    reservationsSh.getRange(rowIndex, cols.cStatus + 1).setValue('cancelled');
    reservationsSh.getRange(rowIndex, cols.cCancelledAt + 1).setValue(new Date());
    reservationsSh.getRange(rowIndex, cols.cCancelledBy + 1).setValue(norm_(leaderId));
    reservationsSh.getRange(rowIndex, cols.cOldSlotReleased + 1).setValue(!!releaseOldSlot);

    try {
      sendAdminAppointmentEmail_(leaderId, {
        type: 'cancelled',
        name: row[cols.cName],
        email: row[cols.cEmail],
        employeeId: row[cols.cEmployeeId],
        branch: row[cols.cBranch],
        regionId: row[cols.cRegion],
        dateStr: row[cols.cDate],
        timeText: row[cols.cTime]
      });
    } catch (err) {
      Logger.log('Admin cancel email error: ' + err);
    }

    return {
      ok: true,
      reservationId: norm_(row[cols.cReservationId]),
      status: 'cancelled',
      old_slot_released: !!releaseOldSlot
    };
  } finally {
    releaseLeaderLock_(lock);
  }
}

function api_adminRescheduleAppointment(leaderId, token, reservationId, newDateStr, newSlotId, releaseOldSlot) {
  requireAdmin_(token, leaderId);

  let lock;
  try {
    lock = acquireLeaderLock_(leaderId, 15000);

    newDateStr = norm_(newDateStr);
    newSlotId = norm_(newSlotId);
    if (!newDateStr) return { ok: false, message: 'newDateStr required' };
    if (!newSlotId) return { ok: false, message: 'newSlotId required' };

    const { reservationsSh } = getLeaderSheets_(leaderId);
    const cols = getReservationCols_(reservationsSh);
    const found = findReservationRowById_(reservationsSh, cols, reservationId);
    if (!found) return { ok: false, message: 'Reservation not found' };

    const tz = getTimeZoneForLeader_(leaderId) || DEFAULT_TZ;
    const row = found.row;
    const rowIndex = found.rowIndex;
    const currentStatus = norm_(row[cols.cStatus] || '').toLowerCase() || 'upcoming';
    if (currentStatus === 'rescheduled') return { ok: false, message: 'This appointment is already rescheduled' };

    const branchValue = norm_(row[cols.cBranch]);
    const branchCode = normBranchCode_(branchValue);
    if (!branchCode) return { ok: false, message: 'Original branch is missing' };

    const available = getAvailableSlotsV2(leaderId, branchCode);
    const picked = (available[newDateStr] || []).find(x => norm_(x.slot_id) === newSlotId);
    if (!picked) return { ok: false, message: 'Selected slot is no longer available' };

    const oldReservationId = norm_(row[cols.cReservationId]);
    const originalIsPast = isPastReservationRow_(row, cols, tz);
    const targetRegionId = resolveRegionByBranch_(leaderId, branchCode) || norm_(row[cols.cRegion]);

    if (currentStatus === 'upcoming' && !originalIsPast) {
      reservationsSh.getRange(rowIndex, cols.cStatus + 1).setValue('rescheduled');
      reservationsSh.getRange(rowIndex, cols.cOldSlotReleased + 1).setValue(!!releaseOldSlot);
    }

    const newReservationId = Utilities.getUuid();
    const newRow = buildReservationRow_(cols, {
      name: row[cols.cName],
      date: newDateStr,
      timeText: picked.timeText,
      createdAt: new Date(),
      email: row[cols.cEmail],
      reservationId: newReservationId,
      employeeId: row[cols.cEmployeeId],
      branch: branchValue,
      region: targetRegionId,
      slotId: newSlotId,
      status: 'upcoming',
      cancelledAt: '',
      cancelled_by: '',
      rescheduled_from: oldReservationId,
      old_slot_released: false
    });

    reservationsSh.appendRow(newRow);

    try {
      sendAdminAppointmentEmail_(leaderId, {
        type: 'rescheduled',
        name: row[cols.cName],
        email: row[cols.cEmail],
        employeeId: row[cols.cEmployeeId],
        branch: row[cols.cBranch],
        regionId: row[cols.cRegion],
        dateStr: row[cols.cDate],
        timeText: row[cols.cTime],
        newBranch: branchValue,
        newRegionId: targetRegionId,
        newDateStr: newDateStr,
        newTimeText: picked.timeText
      });
    } catch (err) {
      Logger.log('Admin reschedule email error: ' + err);
    }

    return {
      ok: true,
      reservationId: newReservationId,
      status: 'upcoming',
      rescheduled_from: oldReservationId,
      old_slot_released: false
    };
  } finally {
    releaseLeaderLock_(lock);
  }
}
/* ===================== ADMIN: APPOINTMENTS LIST ===================== */
function appointmentRowsForLeader_(leaderId, limit) {
  limit = Number(limit || 200);
  if (isNaN(limit) || limit < 1) limit = 200;
  if (limit > 1000) limit = 1000;

  const { reservationsSh, regionsSh } = getLeaderSheets_(leaderId);
  const lastRow = reservationsSh.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = Math.max(15, reservationsSh.getLastColumn());
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;

  const headerRow = reservationsSh.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = reservationsSh.getRange(startRow, 1, numRows, lastCol).getValues();
  const tz = getTimeZoneForLeader_(leaderId) || DEFAULT_TZ;
  const leaderInfo = getLeaderInfoById_(leaderId) || {};
  const regionNameMap = {};
  const regionData = regionsSh.getDataRange().getValues();
  if (regionData.length >= 2) {
    const regionHeaders = regionData[0].map((h) => String(h || "").trim());
    const idxId = idxOfHeader_(regionHeaders, "region_id");
    const idxName = idxOfHeader_(regionHeaders, "region_name");
    if (idxId !== -1) {
      regionData.slice(1).forEach((row) => {
        const id = norm_(row[idxId]);
        if (!id) return;
        regionNameMap[id] = idxName >= 0 ? norm_(row[idxName]) : id;
      });
    }
  }

  function pickIdx_(candidates, fallbackIdx) {
    for (let i = 0; i < candidates.length; i++) {
      const idx = idxOfHeader_(headerRow, candidates[i]);
      if (idx !== -1) return idx;
    }
    return fallbackIdx;
  }

  const cName = pickIdx_(["name"], 0);
  const cDate = pickIdx_(["date", "day"], 1);
  const cTime = pickIdx_(["timeText", "time_text", "apptText", "appointment_text"], 2);
  const cEmail = pickIdx_(["email"], 4);
  const cReservationId = pickIdx_(["reservationId", "reservation_id"], 5);
  const cEmployeeId = pickIdx_(["employeeId", "employee_id"], 6);
  const cBranch = pickIdx_(["branch", "pharmacy", "branch_code"], 7);
  const cRegion = pickIdx_(["region", "location", "region_id"], 8);
  const cStatus = pickIdx_(["status"], 10);
  const cCancelledAt = pickIdx_(["cancelledAt", "cancelled_at"], 11);
  const cCancelledBy = pickIdx_(["cancelled_by", "cancelledBy"], 12);
  const cRescheduledFrom = pickIdx_(["rescheduled_from", "rescheduledFrom"], 13);
  const cOldSlotReleased = pickIdx_(["old_slot_released", "oldSlotReleased"], 14);

  return values
    .filter(r => (
      r[cName] || r[cDate] || r[cTime] || r[cEmail] ||
      r[cEmployeeId] || r[cBranch] || r[cRegion] || r[cReservationId]
    ))
    .map(r => ({
      reservationId: norm_(r[cReservationId]),
      day: toYMD_(r[cDate], tz),
      apptText: norm_(r[cTime]),
      name: norm_(r[cName]),
      employeeId: norm_(r[cEmployeeId]),
      pharmacy: norm_(r[cBranch]),
      location: norm_(r[cRegion]),
      locationName: regionNameMap[norm_(r[cRegion])] || norm_(r[cRegion]),
      email: norm_(r[cEmail]),
      status: norm_(r[cStatus]) || "upcoming",
      cancelledAt: toYMD_(r[cCancelledAt], tz) || norm_(r[cCancelledAt]),
      cancelled_by: norm_(r[cCancelledBy]),
      rescheduled_from: norm_(r[cRescheduledFrom]),
      old_slot_released: truthy_(r[cOldSlotReleased]),
      leaderId,
      leaderName: norm_(leaderInfo.leader_name) || leaderId
    }))
    .reverse();
}

function api_adminListAppointments(leaderId, token, limit) {
  requireAdmin_(token, leaderId);
  return { ok: true, rows: appointmentRowsForLeader_(leaderId, limit) };
}

function api_seniorBootstrap(seniorEmail, token) {
  requireSenior_(token, seniorEmail);

  const leaders = leadersForSenior_(seniorEmail);
  let rows = [];
  leaders.forEach((leader) => {
    try {
      rows = rows.concat(appointmentRowsForLeader_(leader.leader_id, 120));
    } catch (err) {
      Logger.log(`Senior bootstrap skipped ${leader.leader_id}: ${err}`);
    }
  });

  rows.sort((a, b) => {
    const ad = appointmentKeyForSort_(a);
    const bd = appointmentKeyForSort_(b);
    return bd - ad;
  });

  return {
    ok: true,
    role: "senior",
    senior_email: norm_(seniorEmail).toLowerCase(),
    leaders: leaders,
    appointments: { ok: true, rows: rows.slice(0, 800) },
    slots: { ok: true, slots: [] },
    regions: { ok: true, rows: [] },
    branches: { ok: true, rows: [] },
    notifications: { ok: true, rows: [] },
    feedback_history: { ok: true, rows: [] }
  };
}

function appointmentKeyForSort_(row) {
  const day = norm_(row && row.day);
  const time = norm_(row && row.apptText);
  const stamp = `${day} ${time}`.trim();
  const ms = new Date(stamp).getTime();
  return isNaN(ms) ? 0 : ms;
}

function api_adminListNotifications(leaderId, token, limit) {
  requireAdmin_(token, leaderId);

  const { reservationsSh } = getLeaderSheets_(leaderId);
  const cols = getReservationCols_(reservationsSh);
  const lastRow = reservationsSh.getLastRow();
  if (lastRow < 2) return { ok: true, rows: [] };

  const tz = getTimeZoneForLeader_(leaderId) || DEFAULT_TZ;
  const maxItems = Math.max(1, Math.min(Number(limit || 20), 100));
  const values = reservationsSh.getRange(2, 1, lastRow - 1, cols.lastCol).getValues();
  const rows = [];

  values.forEach((r) => {
    const name = norm_(r[cols.cName]) || '—';
    const employeeId = norm_(r[cols.cEmployeeId]) || '—';
    const branch = norm_(r[cols.cBranch]) || '—';
    const regionId = norm_(r[cols.cRegion]);
    const region = regionId ? getRegionInfo_(leaderId, regionId) : null;
    const location = region && region.region_name ? region.region_name : (regionId || '—');
    const dateStr = toYMD_(r[cols.cDate], tz) || norm_(r[cols.cDate]);
    const timeText = norm_(r[cols.cTime]) || '—';
    const reservationId = norm_(r[cols.cReservationId]);
    const cancelReason = norm_(r[cols.cCancellationReason]);

    const createdAt = r[cols.cCreatedAt];
    const createdDate = createdAt ? new Date(createdAt) : null;
    if (createdDate && !isNaN(createdDate.getTime())) {
      rows.push({
        type: 'booking',
        atMs: createdDate.getTime(),
        atText: Utilities.formatDate(createdDate, tz, 'yyyy-MM-dd HH:mm'),
        reservationId,
        name,
        employeeId,
        branch,
        location,
        date: dateStr,
        timeText
      });
    }

    const status = norm_(r[cols.cStatus] || '').toLowerCase();
    const cancelledAt = r[cols.cCancelledAt];
    const cancelledDate = cancelledAt ? new Date(cancelledAt) : null;
    if (status === 'cancelled' && cancelledDate && !isNaN(cancelledDate.getTime())) {
      rows.push({
        type: 'cancellation',
        atMs: cancelledDate.getTime(),
        atText: Utilities.formatDate(cancelledDate, tz, 'yyyy-MM-dd HH:mm'),
        reservationId,
        name,
        employeeId,
        branch,
        location,
        date: dateStr,
        timeText,
        cancellation_reason: cancelReason
      });
    }
  });

  rows.sort((a, b) => b.atMs - a.atMs);
  return { ok: true, rows: rows.slice(0, maxItems) };
}

/* ===================== ADMIN: LOCATIONS (REGIONS) ===================== */
function api_adminCreateRegion(leaderId, token, payload) {
  requireAdmin_(token, leaderId);

  const { regionsSh } = getRegionsBranchesSheets_(leaderId);
  ensureUnmappedRegion_(regionsSh);

  const name = norm_(payload?.region_name);
  const type = norm_(payload?.region_type || "IN_PERSON").toUpperCase();
  const meet = norm_(payload?.meet_url);
  const map  = norm_(payload?.map_url);

  if (!name) return { ok:false, message:"region_name مطلوب" };
  if (!["IN_PERSON","ONLINE"].includes(type)) return { ok:false, message:"region_type لازم IN_PERSON أو ONLINE" };

  const values = regionsSh.getDataRange().getValues();
  const h = values[0].map(x=>String(x).trim());

  const cId   = findHeaderIndex_(h, "region_id");
  const cName = findHeaderIndex_(h, "region_name");
  const cType = findHeaderIndex_(h, "region_type");
  const cMeet = idxOfHeader_(h, "meet_url");
  const cMap  = idxOfHeader_(h, "map_url");
  const cSt   = idxOfHeader_(h, "status");

  const newId = nextRegionId_(regionsSh);

  const row = new Array(h.length).fill("");
  row[cId] = newId;
  row[cName] = name;
  row[cType] = type;

  if (cMeet !== -1) row[cMeet] = (type === "ONLINE" ? meet : "");
  if (cMap  !== -1) row[cMap]  = (type === "IN_PERSON" ? map : "");
  if (cSt   !== -1) row[cSt]   = true;

  regionsSh.appendRow(row);
  return { ok:true, region_id:newId };
}

function api_adminUpdateRegion(leaderId, token, regionId, payload) {
  requireAdmin_(token, leaderId);

  const rid = norm_(regionId);
  if (!rid) return { ok: false, message: "region_id مطلوب" };
  if (rid === UNMAPPED_ID) return { ok: false, message: "مينفعش تعديل UNMAPPED" };

  const { regionsSh } = getRegionsBranchesSheets_(leaderId);
  ensureUnmappedRegion_(regionsSh);

  const values = regionsSh.getDataRange().getValues();
  if (values.length < 2) return { ok: false, message: "Regions sheet empty" };

  const h = values[0].map(x => String(x).trim());
  const cId = findHeaderIndex_(h, "region_id");
  const cName = findHeaderIndex_(h, "region_name");
  const cType = findHeaderIndex_(h, "region_type");
  const cMeet = idxOfHeader_(h, "meet_url");
  const cMap = idxOfHeader_(h, "map_url");
  const cSt = idxOfHeader_(h, "status");

  let rowIndex = -1;
  for (let i = 2; i <= values.length; i++) {
    if (String(values[i - 1][cId]).trim() === rid) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex === -1) return { ok: false, message: "Region not found" };

  const name = (payload && "region_name" in payload) ? norm_(payload.region_name) : null;
  const type = (payload && "region_type" in payload) ? norm_(payload.region_type).toUpperCase() : null;
  const meet = (payload && "meet_url" in payload) ? norm_(payload.meet_url) : null;
  const map = (payload && "map_url" in payload) ? norm_(payload.map_url) : null;

  if (name !== null && !name) return { ok: false, message: "region_name مينفعش فاضي" };
  if (type !== null && !["IN_PERSON", "ONLINE"].includes(type)) return { ok: false, message: "region_type غلط" };

  const cur = values[rowIndex - 1];
  const curType = String(cur[cType] || "").trim().toUpperCase() || "IN_PERSON";
  const newType = type || curType;

  const finalMeet = (newType === "ONLINE")
    ? (meet !== null ? meet : (cMeet !== -1 ? norm_(cur[cMeet]) : ""))
    : "";
  const finalMap = (newType === "IN_PERSON")
    ? (map !== null ? map : (cMap !== -1 ? norm_(cur[cMap]) : ""))
    : "";

  if (name !== null) regionsSh.getRange(rowIndex, cName + 1).setValue(name);
  if (type !== null) regionsSh.getRange(rowIndex, cType + 1).setValue(newType);

  if (cMeet !== -1) regionsSh.getRange(rowIndex, cMeet + 1).setValue(finalMeet);
  if (cMap !== -1) regionsSh.getRange(rowIndex, cMap + 1).setValue(finalMap);

  if (cSt !== -1 && payload && "status" in payload) regionsSh.getRange(rowIndex, cSt + 1).setValue(!!payload.status);

  return { ok: true };
}

// Canonical regions list: excludes DELETED, returns boolean status
function api_adminListRegions(leaderId, token) {
  requireAdmin_(token, leaderId);

  const { regionsSh } = getLeaderSheets_(leaderId);
  if (!regionsSh) return { ok: true, rows: [] };

  const values = regionsSh.getDataRange().getValues();
  if (!values || values.length < 2) return { ok: true, rows: [] };

  const headers = values[0].map(h => String(h || "").trim());
  const hNorm = headers.map(h => h.replace(/\s+/g, " ").toLowerCase());

  const idxId = hNorm.indexOf("region_id");
  const idxName = hNorm.indexOf("region_name");
  const idxType = hNorm.indexOf("region_type");
  const idxMeet = hNorm.indexOf("meet_url");
  const idxMap = hNorm.indexOf("map_url");
  const idxStat = hNorm.indexOf("status");

  if (idxId === -1) return { ok: true, rows: [] };

  const toBoolStatus_ = (v) => {
    if (v === true) return true;
    if (v === false) return false;
    const s = String(v ?? "").trim().toUpperCase();
    if (s === "TRUE" || s === "1" || s === "YES") return true;
    if (s === "FALSE" || s === "0" || s === "NO" || s === "") return false;
    return false;
  };

  const normType_ = (v) => {
    const t = String(v ?? "").trim().toUpperCase();
    if (t === "ONLINE") return "ONLINE";
    if (t === "IN_PERSON" || t === "INPERSON") return "IN_PERSON";
    return "";
  };

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i] || [];
    const id = String(r[idxId] || "").trim();
    if (!id) continue;

    const statusRaw = (idxStat >= 0) ? r[idxStat] : true;
    const statusStr = String(statusRaw ?? "").trim().toUpperCase();

    if (statusStr === "DELETED") continue;

    const regionType = (idxType >= 0) ? normType_(r[idxType]) : "";

    rows.push({
      region_id: id,
      region_name: (idxName >= 0) ? String(r[idxName] || "").trim() : "",
      region_type: regionType,
      meet_url: (idxMeet >= 0) ? String(r[idxMeet] || "").trim() : "",
      map_url: (idxMap >= 0) ? String(r[idxMap] || "").trim() : "",
      status: toBoolStatus_(statusRaw)
    });
  }

  return { ok: true, rows };
}

// Canonical region status setter (cascade branches enable/disable)
function api_adminSetRegionStatus(leaderId, token, regionId, newStatus) {
  requireAdmin_(token, leaderId);

  const { regionsSh, branchesSh } = getLeaderSheets_(leaderId);
  if (!regionsSh) return { ok: false, message: "Regions sheet not found" };
  if (!branchesSh) return { ok: false, message: "Branches sheet not found" };

  ensureUnmappedRegion_(regionsSh);

  const rid = norm_(regionId);
  if (!rid) return { ok: false, message: "region_id مطلوب" };
  if (rid === UNMAPPED_ID) return { ok: false, message: "مينفعش تعطيل/تفعيل UNMAPPED" };

  const values = regionsSh.getDataRange().getValues();
  if (!values || values.length < 2) return { ok: false, message: "No regions data" };

  const headers = values[0].map(h => String(h || "").trim());
  const hNorm = headers.map(h => h.replace(/\s+/g, " ").toLowerCase());

  const idxId = hNorm.indexOf("region_id");
  const idxStat = hNorm.indexOf("status");
  if (idxId < 0 || idxStat < 0) return { ok: false, message: "Missing columns (region_id/status)" };

  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idxId] || "").trim() === rid) {
      rowIndex = i + 1; // 1-based
      break;
    }
  }
  if (rowIndex === -1) return { ok: false, message: "Region not found" };

  const curStatusRaw = values[rowIndex - 1][idxStat];
  const curStatusStr = String(curStatusRaw ?? "").trim().toUpperCase();
  if (curStatusStr === "DELETED") return { ok: false, message: "Region is DELETED" };

  const finalStatus = !!newStatus;
  regionsSh.getRange(rowIndex, idxStat + 1).setValue(finalStatus);

  // Cascade all branches in this region
  const bVals = branchesSh.getDataRange().getValues();
  if (bVals && bVals.length >= 2) {
    const bh = bVals[0].map(h => String(h || "").trim());
    const bhNorm = bh.map(h => h.replace(/\s+/g, " ").toLowerCase());
    const bIdxRegion = bhNorm.indexOf("region_id");
    const bIdxStatus = bhNorm.indexOf("status");

    if (bIdxRegion >= 0 && bIdxStatus >= 0) {
      for (let r = 1; r < bVals.length; r++) {
        const brRid = String(bVals[r][bIdxRegion] || "").trim();
        if (brRid === rid) {
          branchesSh.getRange(r + 1, bIdxStatus + 1).setValue(finalStatus);
        }
      }
    }
  }

  return { ok: true, region_id: rid, status: finalStatus };
}

// Backward compatibility: old name
function api_adminToggleRegionStatus(leaderId, token, regionId, newStatus) {
  return api_adminSetRegionStatus(leaderId, token, regionId, newStatus);
}

function api_adminDeleteRegion(leaderId, token, regionId, deleteBranchesAlso) {
  requireAdmin_(token, leaderId);

  const { regionsSh, branchesSh, slotsSh } = getLeaderSheets_(leaderId);
  if (!regionsSh) return { ok:false, message:"Regions sheet not found" };
  if (!branchesSh) return { ok:false, message:"Branches sheet not found" };
  if (!slotsSh) return { ok:false, message:"Slots sheet not found" };

  ensureUnmappedRegion_(regionsSh);

  const rid = norm_(regionId);
  if (!rid) return { ok:false, message:"Missing region id" };
  if (rid === UNMAPPED_ID) return { ok:false, message:"مينفعش حذف غير مُحدد" };

  const regData = regionsSh.getDataRange().getValues();
  if (regData.length < 2) return { ok:false, message:"No regions data" };

  const headers = regData[0].map(h=>String(h||"").trim());
  const idxId = idxOfHeader_(headers, "region_id");
  const idxStatus = idxOfHeader_(headers, "status");
  if (idxId < 0) return { ok:false, message:"Missing region_id column" };
  if (idxStatus < 0) return { ok:false, message:"Missing status column" };

  // 1) حذف/نقل الفروع المرتبطة
  const bData = branchesSh.getDataRange().getValues();
  if (bData.length >= 2) {
    const bh = bData[0].map(h=>String(h||"").trim());
    const bIdxRegion = idxOfHeader_(bh, "region_id");
    const bIdxStatus = idxOfHeader_(bh, "status");

    if (bIdxRegion >= 0) {
      for (let i = bData.length - 1; i >= 1; i--) {
        const row = bData[i] || [];
        if (String(row[bIdxRegion] || "").trim() !== rid) continue;

        if (deleteBranchesAlso) {
          branchesSh.deleteRow(i + 1);
        } else {
          branchesSh.getRange(i + 1, bIdxRegion + 1).setValue(UNMAPPED_ID);
          if (bIdxStatus >= 0 && (row[bIdxStatus] == null || row[bIdxStatus] === "")) {
            branchesSh.getRange(i + 1, bIdxStatus + 1).setValue(true);
          }
        }
      }
    }
  }

  // 2) حذف الـ slots المرتبطة بالمنطقة
  const sData = slotsSh.getDataRange().getValues();
  if (sData.length >= 2) {
    const sh = sData[0].map(h => String(h || "").trim());
    const sIdxRegion = idxOfHeader_(sh, "region_id");

    if (sIdxRegion >= 0) {
      for (let i = sData.length - 1; i >= 1; i--) {
        const row = sData[i] || [];
        if (String(row[sIdxRegion] || "").trim() === rid) {
          slotsSh.deleteRow(i + 1);
        }
      }
    }
  }

  // 3) Soft delete للمنطقة
  for (let i = 1; i < regData.length; i++) {
    const row = regData[i] || [];
    if (String(row[idxId] || "").trim() === rid) {
      regionsSh.getRange(i + 1, idxStatus + 1).setValue("DELETED");
      return { ok:true };
    }
  }

  return { ok:false, message:"Region not found" };
}

/* ===================== ADMIN: LOCATIONS (BRANCHES) ===================== */
function api_adminCreateBranch(leaderId, token, payload) {
  requireAdmin_(token, leaderId);

  const { branchesSh, regionsSh } = getLeaderSheets_(leaderId);
  if (!branchesSh) return { ok: false, message: "Branches sheet not found" };
  if (!regionsSh) return { ok: false, message: "Regions sheet not found" };

  ensureUnmappedRegion_(regionsSh);

  const codeNorm = normBranchCode_(payload?.branch_code);
  if (!codeNorm) return { ok: false, message: "اكتب رقم الفرع فقط (مثال: 1200)" };

  const regionId = norm_(payload?.region_id || UNMAPPED_ID) || UNMAPPED_ID;

  if (regionId !== UNMAPPED_ID && !isRegionUsable_(regionsSh, regionId)) {
    return { ok: false, message: "المنطقة غير موجودة أو غير مفعّلة" };
  }

  const name = "PH" + codeNorm;
  const status = true;

  const data = branchesSh.getDataRange().getValues();
  if (!data || data.length < 1) return { ok: false, message: "Branches sheet has no header" };

  const headers = data[0].map(x => String(x || "").trim());
  const idxCode = idxOfHeader_(headers, "branch_code");
  const idxName = idxOfHeader_(headers, "branch_name");
  const idxRegion = idxOfHeader_(headers, "region_id");
  const idxStatus = idxOfHeader_(headers, "status");

  if (idxCode < 0) return { ok: false, message: "Missing column: branch_code" };
  if (idxRegion < 0) return { ok: false, message: "Missing column: region_id" };
  if (idxStatus < 0) return { ok: false, message: "Missing column: status" };

  const exists = data.slice(1).some(r => normBranchCode_(r[idxCode]) === codeNorm);
  if (exists) return { ok: false, message: "رقم الفرع موجود بالفعل" };

  const newRow = new Array(headers.length).fill("");
  newRow[idxCode] = codeNorm;
  if (idxName >= 0) newRow[idxName] = name;
  newRow[idxRegion] = regionId;
  newRow[idxStatus] = status;

  branchesSh.appendRow(newRow);

  return {
    ok: true,
    row: { branch_code: codeNorm, branch_name: name, region_id: regionId, status }
  };
}

function api_adminUpdateBranch(leaderId, token, branchCode, payload) {
  requireAdmin_(token, leaderId);

  const codeNorm = normBranchCode_(branchCode);
  if (!codeNorm) return { ok: false, message: "branch_code مطلوب" };

  const { branchesSh, regionsSh } = getRegionsBranchesSheets_(leaderId);
  ensureUnmappedRegion_(regionsSh);

  const values = branchesSh.getDataRange().getValues();
  if (values.length < 2) return { ok: false, message: "Branches sheet empty" };

  const h = values[0].map(x => String(x).trim());
  const cCode = findHeaderIndex_(h, "branch_code");
  const cName = idxOfHeader_(h, "branch_name");
  const cRid = idxOfHeader_(h, "region_id");
  const cSt = idxOfHeader_(h, "status");

  let rowIndex = -1;
  for (let i = 2; i <= values.length; i++) {
    if (normBranchCode_(values[i - 1][cCode]) === codeNorm) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex === -1) return { ok: false, message: "Branch not found" };

  if (payload && "branch_name" in payload && cName !== -1) {
    branchesSh.getRange(rowIndex, cName + 1).setValue(norm_(payload.branch_name));
  }

  if (payload && "region_id" in payload && cRid !== -1) {
    const rid = norm_(payload.region_id) || UNMAPPED_ID;
    if (rid !== UNMAPPED_ID && !isRegionUsable_(regionsSh, rid)) {
      return { ok: false, message: "المنطقة غير موجودة أو غير مفعّلة" };
    }
    branchesSh.getRange(rowIndex, cRid + 1).setValue(rid);
  }

  if (payload && "status" in payload && cSt !== -1) {
    branchesSh.getRange(rowIndex, cSt + 1).setValue(!!payload.status);
  }

  return { ok: true };
}

// Canonical smart delete (delete if UNMAPPED, else move to UNMAPPED)
function api_adminDeleteBranchSmart(leaderId, token, branchCode) {
  return api_adminDeleteBranch(leaderId, token, branchCode);
}

// Canonical branch status setter
function api_adminToggleBranchStatus(leaderId, token, branchCode, newStatus) {
  requireAdmin_(token, leaderId);

  const { branchesSh } = getLeaderSheets_(leaderId);
  if (!branchesSh) return { ok: false, message: "Branches sheet not found" };

  const values = branchesSh.getDataRange().getValues();
  if (values.length < 2) return { ok: false, message: "No data" };

  const headers = values[0].map(h => String(h || "").trim());
  const idxCode = idxOfHeader_(headers, "branch_code");
  const idxStatus = idxOfHeader_(headers, "status");
  const idxName = idxOfHeader_(headers, "branch_name");
  const idxRegion = idxOfHeader_(headers, "region_id");

  if (idxCode === -1 || idxStatus === -1)
    return { ok: false, message: "Missing columns (branch_code/status)" };

  const target = normBranchCode_(branchCode);
  if (!target) return { ok: false, message: "Invalid branch code" };

  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const rowCodeNorm = normBranchCode_(row[idxCode]);

    if (rowCodeNorm === target) {
      branchesSh.getRange(i + 1, idxCode + 1).setValue(target);

      const st = !!newStatus;
      branchesSh.getRange(i + 1, idxStatus + 1).setValue(st);

      if (idxName >= 0) {
        const curName = norm_(row[idxName]);
        if (!curName) branchesSh.getRange(i + 1, idxName + 1).setValue("PH" + target);
      }

      const out = {
        branch_code: target,
        branch_name: (idxName >= 0 ? (norm_(row[idxName]) || ("PH" + target)) : ("PH" + target)),
        region_id: (idxRegion >= 0 ? norm_(row[idxRegion]) : ""),
        status: st
      };

      return { ok: true, row: out };
    }
  }

  return { ok: false, message: "Branch not found" };
}

/* ===== Branch backward-compatibility aliases ===== */
function api_adminDeleteBranch(leaderId, token, branchCode) {
  requireAdmin_(token, leaderId);

  const { branchesSh, regionsSh } = getLeaderSheets_(leaderId);
  if (!branchesSh) return { ok:false, message:"Branches sheet not found" };
  if (regionsSh) ensureUnmappedRegion_(regionsSh);

  const values = branchesSh.getDataRange().getValues();
  if (!values || values.length < 2) return { ok:false, message:"No data" };

  const headers = values[0].map(h => String(h || "").trim());
  const idxCode   = idxOfHeader_(headers, "branch_code");
  const idxRegion = idxOfHeader_(headers, "region_id");
  const idxName   = idxOfHeader_(headers, "branch_name");
  const idxStatus = idxOfHeader_(headers, "status");

  if (idxCode < 0 || idxRegion < 0) {
    return { ok:false, message:"Missing columns (branch_code/region_id)" };
  }

  const target = normBranchCode_(branchCode);
  if (!target) return { ok:false, message:"Invalid branch code" };

  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const rowCodeNorm = normBranchCode_(row[idxCode]);

    if (rowCodeNorm !== target) continue;

    const currentRegion = String(row[idxRegion] || "").trim();

    // لو الفرع داخل UNMAPPED => حذف نهائي
    if (currentRegion === UNMAPPED_ID) {
      branchesSh.deleteRow(i + 1);
      return {
        ok: true,
        action: "deleted",
        branch_code: target
      };
    }

    // غير كده => انقله إلى UNMAPPED
    branchesSh.getRange(i + 1, idxRegion + 1).setValue(UNMAPPED_ID);

    if (idxStatus >= 0) {
      const st = row[idxStatus];
      if (st == null || st === "") {
        branchesSh.getRange(i + 1, idxStatus + 1).setValue(true);
      }
    }

    const name = (idxName >= 0 ? (norm_(row[idxName]) || ("PH" + target)) : ("PH" + target));

    return {
      ok: true,
      action: "moved_to_unmapped",
      row: {
        branch_code: target,
        branch_name: name,
        region_id: UNMAPPED_ID,
        status: (idxStatus >= 0 ? truthy_(row[idxStatus]) : true)
      }
    };
  }

  return { ok:false, message:"Branch not found" };
}
function api_adminUpdateBranchStatus(leaderId, token, branchCode, status) {
  return api_adminToggleBranchStatus(leaderId, token, branchCode, status);
}
function api_adminMoveBranchToUnmapped(leaderId, token, branchCode) {
  return api_adminUpdateBranch(leaderId, token, branchCode, { region_id: UNMAPPED_ID });
}

/* ===================== ADMIN: LIST BRANCHES (simple) ===================== */
function api_adminListBranches(leaderId, token) {
  requireAdmin_(token, leaderId);

  const { branchesSh, regionsSh } = getLeaderSheets_(leaderId);
  if (regionsSh) ensureUnmappedRegion_(regionsSh);

  const values = branchesSh.getDataRange().getValues();
  if (values.length < 2) return { ok: true, rows: [] };

  const h = values[0].map(x => String(x).trim());
  const rows = values.slice(1)
    .map(r => {
      const o = {};
      h.forEach((k, i) => o[k] = r[i]);
      return o;
    })
    .filter(x => String(x.branch_code || "").trim());

  return { ok: true, rows };
}

/* ===================== ADMIN: AREAS LIST (for filters) ===================== */
function api_adminListAreas(leaderId, token) {
  requireAdmin_(token, leaderId);

  const { regionsSh } = getLeaderSheets_(leaderId);
  const areas = [{ id: "ALL", name: "All" }];

  if (!regionsSh) return { ok: true, areas };

  const values = regionsSh.getDataRange().getValues();
  if (!values || values.length < 2) return { ok: true, areas };

  const headers = values[0].map(h => String(h || "").trim());
  const hNorm = headers.map(h => h.replace(/\s+/g, " ").toLowerCase());

  const idxId = hNorm.indexOf("region_id");
  const idxName = hNorm.indexOf("region_name");
  const idxStat = hNorm.indexOf("status");

  if (idxId === -1) return { ok: true, areas };

  for (let i = 1; i < values.length; i++) {
    const r = values[i] || [];
    const id = String(r[idxId] || "").trim();
    if (!id) continue;

    const statusRaw = (idxStat >= 0) ? r[idxStat] : true;
    const statusStr = String(statusRaw ?? "").trim().toUpperCase();

    if (statusStr === "DELETED") continue;
    if (statusRaw === false || statusStr === "FALSE") continue;

    const name = (idxName >= 0) ? String(r[idxName] || "").trim() : id;
    areas.push({ id, name: name || id });
  }

  return { ok: true, areas };

}  

/* ===================== PUBLIC: LIST BRANCHES FOR PAGE ===================== */
function api_listBranches(leaderId) {
  if (!leaderId) return [];

  const { branchesSh } = getLeaderSheets_(leaderId);
  const values = branchesSh.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  const headers = values[0].map(h => String(h || "").trim());
  const idxCode   = idxOfHeader_(headers, "branch_code");
  const idxName   = idxOfHeader_(headers, "branch_name");
  const idxRegion = idxOfHeader_(headers, "region_id");
  const idxStatus = idxOfHeader_(headers, "status");

  if (idxCode < 0) return [];

  return values.slice(1)
    .filter(r => {
      const code = normBranchCode_(r[idxCode]);
      if (!code) return false;

      // استبعد الفروع غير المفعلة من صفحة الحجز
      if (idxStatus >= 0 && !truthy_(r[idxStatus])) return false;

      return true;
    })
    .map(r => ({
      branch_code: "PH" + normBranchCode_(r[idxCode]),
      branch_name: idxName >= 0 ? norm_(r[idxName]) : "",
      region_id: idxRegion >= 0 ? norm_(r[idxRegion]) : ""
    }));
}

// -------------------------------------------------------------------------------------------

  function makeHash_TL010() {
    const pass = "12345";        // الباسورد اللي إنت عايزه
    const hash = sha256_(pass);  // نحوله لهاش
    Logger.log(hash);
  }

  function createPasswordForLeader() {
    const password = generatePassword_(12); // 12 حرف
    const hash = sha256_(password);
    
    Logger.log("كلمة المرور: " + password);
    Logger.log("الهاش: " + hash);
    
    // انسخ الهاش وضعه في خانة admin_pass_hash للمشرف في sheet "Leaders"
  }

  function generatePassword_(len) {
    len = len || 12;
    const upper = "ABCDEFGHJKLMNPQRSTUVWXYZ";
    const lower = "abcdefghijkmnpqrstuvwxyz";
    const digits = "23456789";
    const symbols = "@#$%_-+";
    const all = upper + lower + digits + symbols;

    // ضمان وجود نوع من كل فئة
    let pass = [
      upper[Math.floor(Math.random() * upper.length)],
      lower[Math.floor(Math.random() * lower.length)],
      digits[Math.floor(Math.random() * digits.length)],
      symbols[Math.floor(Math.random() * symbols.length)],
    ];

    while (pass.length < len) {
      pass.push(all[Math.floor(Math.random() * all.length)]);
    }

    // shuffle
    for (let i = pass.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [pass[i], pass[j]] = [pass[j], pass[i]];
    }
    return pass.join("");
  }

  function admin_generatePasswordsAndEmailMe() {
    const TO_EMAIL = "omar0452007@gmail.com"; 
    const ss = SpreadsheetApp.openById(REGISTRY_SHEET_ID);
    const sh = ss.getSheetByName(REGISTRY_TAB);
    const values = sh.getDataRange().getValues();
    if (values.length < 2) throw new Error("Leaders sheet is empty.");

    const headers = values[0].map(h => String(h).trim());
    const idxId = headers.indexOf("leader_id");
    const idxName = headers.indexOf("leader_name");
    const idxStatus = headers.indexOf("status");
    const idxHash = headers.indexOf("admin_pass_hash");

    if (idxId === -1 || idxName === -1 || idxStatus === -1 || idxHash === -1) {
      throw new Error("Missing required columns: leader_id, leader_name, status, admin_pass_hash");
    }

    const out = [];
    const updates = []; // {row, hash}

    for (let r = 1; r < values.length; r++) {
      const leaderId = String(values[r][idxId] || "").trim();
      const leaderName = String(values[r][idxName] || "").trim();
      const status = String(values[r][idxStatus] || "").trim().toLowerCase();
      const existingHash = String(values[r][idxHash] || "").trim();

      if (!leaderId) continue;
      if (status !== "active") continue;

      // if (existingHash) continue;

      const pass = generatePassword_(12);
      const hash = sha256_(pass);

      updates.push({ row: r + 1, hash }); // sheet rows are 1-based
      out.push([leaderId, leaderName, pass]);
    }

    // اكتب الهاشات دفعة واحدة
    updates.forEach(u => sh.getRange(u.row, idxHash + 1).setValue(u.hash));

    // لو مفيش قادة محتاجين باسورد
    if (out.length === 0) {
      MailApp.sendEmail({
        to: TO_EMAIL,
        subject: "Leaders Passwords - No Updates",
        body: "No active leaders were missing admin_pass_hash. No passwords were generated."
      });
      return { ok: true, generated: 0 };
    }

    // جهّز CSV في الإيميل
    const lines = [];
    lines.push("leader_id,leader_name,password");
    out.forEach(row => {
      // escape commas/quotes
      const safe = row.map(v => `"${String(v).replace(/"/g, '""')}"`);
      lines.push(safe.join(","));
    });

    const body =
      "Hi,\n\nAttached below are the generated admin passwords for leaders (NEW ONLY).\n" +
      "Please store them securely and delete this email after saving.\n\n" +
      lines.join("\n");

    MailApp.sendEmail({
      to: TO_EMAIL,
      subject: "✅ Generated Admin Passwords for Leaders",
      body: body,
      name: "Reservation System"
    });

    return { ok: true, generated: out.length };
  }

  function runSlotsMigration() {
    const leaderId = "TL009"; // 👈 حط Leader ID 
    const res = migrateSlots_AddSlotIdAndDateCols_(leaderId);
    Logger.log(res);
  }

  function migrateSlots_AddSlotIdAndDateCols_(leaderId) {
    const { slotsSh } = getLeaderSheets_(leaderId);

    const a1 = String(slotsSh.getRange(1,1).getValue() || "").toLowerCase().trim();
    if (a1 === "slot_id") return { ok:true, message:"Already migrated" };

    // Insert new col A for slot_id
    slotsSh.insertColumnBefore(1);

    // Write headers (نفترض أول صف headers)
    slotsSh.getRange(1,1,1,5).setValues([["slot_id","region_id","date","timeText","active"]]);

    const lastRow = slotsSh.getLastRow();
    if (lastRow < 2) return { ok:true, message:"Migrated (no data)" };

    const old = slotsSh.getRange(2,2,lastRow-1,4).getValues(); // B:E القديمة
    const out = old.map(r => {
      const regionId = r[0];
      const dayOld = r[1];      // كان اسم يوم أو تاريخ قديم
      const timeText = r[2];
      const active = r[3];

      // نحاول نحول "dayOld" لتاريخ نصي YYYY-MM-DD لو هو Date object
      let dateStr = "";
      if (dayOld instanceof Date) {
        dateStr = Utilities.formatDate(dayOld, DEFAULT_TZ, "yyyy-MM-dd");
      } else {
        dateStr = String(dayOld || "").trim(); // لو نص، سيبه زي ما هو
      }

      return [newId_(), String(regionId||"").trim(), dateStr, String(timeText||"").trim(), active];
    });

    slotsSh.getRange(2,1,out.length,5).setValues(out);

    return { ok:true, message:"Migrated with slot_id + date" };
  }


  function createTestPasswordForTL010() {
    const plainPassword = "12345";
    const hash = sha256_(plainPassword);
    
    Logger.log("الهاش الناتج: " + hash);
    
    // ابحث عن leader_id = TL010
    const regSh = SpreadsheetApp.openById(REGISTRY_SHEET_ID).getSheetByName(REGISTRY_TAB);
    const data = regSh.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    
    const colId = headers.indexOf("leader_id");
    const colHash = headers.indexOf("admin_pass_hash");
    
    if (colId === -1 || colHash === -1) {
      Logger.log("الأعمدة المطلوبة غير موجودة");
      return;
    }
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colId]).trim() === "TL010") {
        // كتابة الهاش في الخانة المناسبة
        regSh.getRange(i + 1, colHash + 1).setValue(hash);
        Logger.log("تم تحديث الهاش للصف " + (i + 1));
        break;
      }
    }
  }
  