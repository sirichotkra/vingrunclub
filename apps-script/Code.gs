/**
 * VING RUN CLUB — ระบบเช็คอินสะสมครั้งด้วย QR
 * Google Apps Script + Google Sheets (ฟรี ไม่ต้องมีเซิร์ฟเวอร์)
 *
 * วิธีติดตั้ง: ดูไฟล์ SETUP.md
 */

var SS_KEY = 'VING_SHEET_ID';
var SH = {
  RUNNERS: 'Runners',
  EVENTS: 'Events',
  CHECKINS: 'Checkins',
  REDEMPTIONS: 'Redemptions',
  SETTINGS: 'Settings',
  DELETED: 'Deleted',
  TIER_IMAGES: 'TierImages',
  PHOTOS: 'Photos'
};
var HEADERS = {
  Runners: ['id', 'name', 'phone', 'createdAt', 'code', 'ชื่อ-นามสกุล', 'วันเกิด', 'updatedAt'],
  Events: ['id', 'code', 'name', 'date', 'active', 'createdAt', 'limit'],
  Checkins: ['id', 'runnerId', 'eventId', 'ts', 'วันเวลา', 'ชื่อ', 'เบอร์โทร', 'รหัสสมาชิก', 'กิจกรรม', 'เพิ่มโดยทีมงาน'],
  Redemptions: ['runnerId', 'tierN', 'ts', 'วันเวลา', 'ชื่อ', 'เบอร์โทร', 'รหัสสมาชิก', 'ของรางวัล'],
  Settings: ['key', 'value'],
  Deleted: ['runnerId', 'name', 'phone', 'checkins', 'deletedAt', 'รหัสสมาชิก'],
  TierImages: ['tierN', 'image', 'updatedAt'],
  Photos: ['runnerId', 'image', 'updatedAt']
};
var DEFAULT_TIERS = [
  { n: 3, reward: 'เหรียญที่ระลึก' },
  { n: 6, reward: 'เสื้อวิ่ง VING' },
  { n: 9, reward: 'กระบอกน้ำ' },
  { n: 12, reward: 'รางวัลใหญ่ประจำซีซั่น' }
];

/* ============================ Web entry ============================ */

/**
 * รายชื่อฟังก์ชันที่ยอมให้เรียกจากหน้าเว็บภายนอก (Cloudflare Pages)
 * อะไรที่ไม่อยู่ในนี้ เรียกจากข้างนอกไม่ได้เด็ดขาด
 */
function apiTable_() {
  return {
    apiRunnerState: apiRunnerState, apiCheckPhone: apiCheckPhone, apiRegister: apiRegister,
    apiRecover: apiRecover, apiCheckin: apiCheckin,
    apiGetProfile: apiGetProfile, apiUpdateProfile: apiUpdateProfile, apiSetPhoto: apiSetPhoto,
    apiGetRunnerPhoto: apiGetRunnerPhoto, apiManualCheckin: apiManualCheckin,
    apiRemoveCheckin: apiRemoveCheckin, apiEventCheckins: apiEventCheckins,
    apiStaffState: apiStaffState, apiCreateEvent: apiCreateEvent,
    apiSetEventLimit: apiSetEventLimit, apiSetEventActive: apiSetEventActive,
    apiDeleteEvent: apiDeleteEvent,
    apiSetTiers: apiSetTiers, apiSetSettings: apiSetSettings, apiDeleteRunner: apiDeleteRunner,
    apiSetTierImage: apiSetTierImage, apiFulfill: apiFulfill, apiUnfulfill: apiUnfulfill
  };
}

function dispatch_(fn, args) {
  var f = apiTable_()[fn];
  if (typeof f !== 'function') throw new Error('ไม่รู้จักคำสั่ง: ' + fn);
  return f.apply(null, args || []);
}

function jsonOut_(obj, callback) {
  var body = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function runApi_(fn, args, callback) {
  try {
    return jsonOut_({ ok: true, data: dispatch_(fn, args) }, callback);
  } catch (err) {
    return jsonOut_({ ok: false, error: String((err && err.message) || err) }, callback);
  }
}

/**
 * POST — ทางหลักที่หน้าเว็บบน Cloudflare ใช้คุยกับที่นี่
 * ส่งมาเป็น text/plain เพื่อเลี่ยง CORS preflight (Apps Script ตอบ preflight ไม่ได้)
 * body: {"fn":"apiRunnerState","args":["abc",""]}
 */
function doPost(e) {
  var fn = '', args = [];
  try {
    var raw = (e && e.postData && e.postData.contents) || '';
    var req = JSON.parse(raw || '{}');
    fn = req.fn; args = req.args || [];
  } catch (err) {
    return jsonOut_({ ok: false, error: 'อ่านคำขอไม่ได้' }, null);
  }
  return runApi_(fn, args, null);
}

function doGet(e) {
  var p = (e && e.parameter) || {};

  // โหมด API (สำรอง) — ใช้ตอน POST ใช้ไม่ได้ เช่นเน็ตบางที่บล็อกไว้
  // เรียกแบบ: ?fn=apiRunnerState&args=[...]&callback=cb
  if (p.fn) {
    var args = [];
    try { args = p.args ? JSON.parse(p.args) : []; } catch (err) { args = []; }
    return runApi_(p.fn, args, p.callback || null);
  }

  // โหมดหน้าเว็บ — ยังใช้ได้เหมือนเดิมทุกอย่าง (ลิงก์ /exec เดิมไม่พัง)
  var t = HtmlService.createTemplateFromFile('Index');
  t.eventCode = p.event ? String(p.event) : '';
  t.webAppUrl = getWebAppUrl_();
  return t.evaluate()
    .setTitle('VING RUN CLUB')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl_() {
  try {
    var u = ScriptApp.getService().getUrl();
    return u ? u.replace(/\/dev$/, '/exec') : '';
  } catch (err) {
    return '';
  }
}

/* ============================ Spreadsheet ============================ */

function getSS_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(SS_KEY);
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (err) { /* recreate below */ }
  }
  var ss = SpreadsheetApp.create('VING RUN CLUB — ข้อมูลเช็คอิน');
  props.setProperty(SS_KEY, ss.getId());
  initSheets_(ss);
  return ss;
}

function initSheets_(ss) {
  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    // เขียนหัวตารางทุกครั้ง (idempotent) เพื่อให้ชีตเก่าที่สร้างไว้ก่อน
    // ได้คอลัมน์ใหม่ที่เพิ่มมาภายหลัง เช่น limit ของ Events
    sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
    if (sh.getLastRow() <= 1) sh.setFrozenRows(1);
  });
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (def && def.getLastRow() === 0) ss.deleteSheet(def);

  // บังคับคอลัมน์ที่เป็น "รหัส" ให้เป็นข้อความล้วน ไม่ให้ Sheets แปลงเป็นตัวเลข
  // (ไม่งั้นเบอร์ 0812345678 จะกลายเป็น 812345678 แล้วหาสมาชิกเดิมไม่เจอ)
  try {
    ss.getSheetByName(SH.RUNNERS).getRange('A:E').setNumberFormat('@');
    ss.getSheetByName(SH.EVENTS).getRange('A:B').setNumberFormat('@');
    ss.getSheetByName(SH.CHECKINS).getRange('A:C').setNumberFormat('@');
    ss.getSheetByName(SH.REDEMPTIONS).getRange('A:A').setNumberFormat('@');
  } catch (err) { /* ไม่ critical — phoneKey_ รองรับกรณีถูกแปลงไปแล้วอยู่ดี */ }

  var st = ss.getSheetByName(SH.SETTINGS);
  if (st.getLastRow() < 2) {
    st.getRange(2, 1, 4, 2).setValues([
      ['adminPin', '2026'],
      ['clubName', 'VING RUN CLUB'],
      ['tiers', JSON.stringify(DEFAULT_TIERS)],
      ['memberSeq', 0]
    ]);
  }
}

function sheet_(name) {
  var ss = getSS_();
  var sh = ss.getSheetByName(name);
  if (!sh) { initSheets_(ss); sh = ss.getSheetByName(name); }
  return sh;
}

function rows_(name) {
  var sh = sheet_(name);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var head = HEADERS[name];
  var vals = sh.getRange(2, 1, last - 1, head.length).getValues();
  return vals.map(function (r) {
    var o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  }).filter(function (o) { return String(o[HEADERS[name][0]]).length > 0; });
}

function append_(name, obj) {
  var sh = sheet_(name);
  var head = HEADERS[name];
  sh.appendRow(head.map(function (h) { return obj[h] === undefined ? '' : obj[h]; }));
}

/* ============================ Settings ============================ */

function getSettings_() {
  var out = { adminPin: '2026', clubName: 'VING RUN CLUB', tiers: DEFAULT_TIERS };
  rows_(SH.SETTINGS).forEach(function (r) {
    var k = String(r.key), v = r.value;
    if (k === 'tiers') {
      try { out.tiers = JSON.parse(v); } catch (e) { /* keep default */ }
    } else if (k) {
      out[k] = String(v);
    }
  });
  if (!Array.isArray(out.tiers) || !out.tiers.length) out.tiers = DEFAULT_TIERS;
  return out;
}

function setSetting_(key, value) {
  var sh = sheet_(SH.SETTINGS);
  var last = sh.getLastRow();
  if (last >= 2) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === key) {
        sh.getRange(i + 2, 2).setValue(value);
        return;
      }
    }
  }
  sh.appendRow([key, value]);
}

function checkPin_(pin) {
  var s = getSettings_();
  // trim ทั้งสองฝั่ง กันช่องว่างที่ติดมาจากการก๊อปวางหรือคีย์บอร์ดมือถือ
  // ทำให้ล็อกอินไม่ผ่านทั้งที่พิมพ์ถูก
  if (String(pin == null ? '' : pin).trim() !== String(s.adminPin).trim()) {
    throw new Error('รหัสผ่านไม่ถูกต้อง');
  }
  return s;
}

/* ============================ Helpers ============================ */

// ขึ้นต้นด้วยตัวอักษรเสมอ กัน Google Sheets แปลง id เป็นตัวเลข
function uid_() {
  return 'r' + Utilities.getUuid().replace(/-/g, '').slice(0, 15);
}

/**
 * คีย์สำหรับเทียบเบอร์โทร
 * Google Sheets จะแปลง "0812345678" เป็นตัวเลข 812345678 (ศูนย์นำหน้าหาย)
 * จึงต้องตัดศูนย์นำหน้าออกทั้งสองฝั่งก่อนเทียบ ไม่งั้นจะหาสมาชิกเดิมไม่เจอ
 */
function phoneKey_(p) {
  var d = String(p == null ? '' : p).replace(/[^0-9]/g, '').replace(/^0+/, '');
  // รองรับเบอร์ที่บันทึกแบบ +66 ให้ตรงกับ 0xx เดิม
  // ตัด 66 ออกเฉพาะตอนที่ตัดแล้วเหลือ 9 หลักพอดี (เบอร์มือถือไทยไม่รวม 0 นำหน้า)
  // ไม่งั้นเบอร์จริงอย่าง 066-123-4567 จะถูกตัดผิด
  if (d.length === 11 && d.indexOf('66') === 0) d = d.slice(2);
  return d;
}

/**
 * แปลงเบอร์กลับเป็นรูปแบบที่คนอ่านเข้าใจ: 081-234-5678
 * ชีตเก่าอาจเก็บเป็นตัวเลข 812345678 (เลข 0 หาย) จึงเติมกลับให้
 */
function phoneDisplay_(p) {
  var d = String(p == null ? '' : p).replace(/[^0-9]/g, '');
  if (!d) return '';
  if (d.length === 11 && d.indexOf('66') === 0) d = d.slice(2);   // +66
  if (d.length === 9) d = '0' + d;                                // 0 นำหน้าหายไป
  if (d.length === 10) return d.slice(0, 3) + '-' + d.slice(3, 6) + '-' + d.slice(6);
  return d;
}

/**
 * เวลาไทยแบบคนอ่านเข้าใจ: 25/08/2026 16:30
 * ในชีตเก็บ ts เป็น UTC (มาตรฐาน) แต่คนอ่านต้องเห็นเวลาไทย
 * ไม่งั้นเช็คอินตอน 5 ทุ่มจะกลายเป็นวันถัดไปในชีต
 */
function thaiTime_(iso) {
  try {
    var tz = Session.getScriptTimeZone() || 'Asia/Bangkok';
    return Utilities.formatDate(iso ? new Date(iso) : new Date(), tz, 'dd/MM/yyyy HH:mm');
  } catch (err) {
    return String(iso || '');
  }
}

/** เติม 0 ข้างหน้าให้ครบ 4 หลัก: 7 -> "0007" */
function pad4_(n) {
  var t = String(n);
  while (t.length < 4) t = '0' + t;
  return t;
}

/** อ่านรหัสสมาชิกจากชีต (ตัด ' ที่ใช้บังคับเป็นข้อความออก) */
function codeOf_(r) {
  return String(r && r.code != null ? r.code : '').replace(/^'/, '').trim();
}

/**
 * เลขสมาชิกถัดไป — เดินหน้าอย่างเดียว ไม่วนกลับมาใช้เลขเดิม
 * ถึงจะลบสมาชิกออกไป เลขที่เคยใช้แล้วก็จะไม่ถูกแจกซ้ำ
 * กันสับสนตอนตรวจย้อนหลังในชีต
 */
/** รหัสสมาชิกที่ยังว่างและน้อยที่สุด
 *
 *  เดิมใช้ตัวนับเดินหน้าอย่างเดียว ลบคนออกแล้วเลขก็ยังวิ่งต่อ เกิดช่องโหว่
 *  (เช่น ลบหมดทุกคน คนใหม่ยังได้ 0007) ตอนนี้เปลี่ยนเป็น "เติมเลขที่ว่าง"
 *  ลบ 0002 ออก คนถัดไปจะได้ 0002 เลขจึงเรียงติดกันเสมอ
 *
 *  หมายเหตุ: เลขที่ถูกนำกลับมาใช้ จะไปตรงกับเลขของคนเก่าในประวัติเช็คอิน
 *  ถ้าต้องการสอบย้อนหลังว่าใครคือใคร ให้ดูที่ชีต Deleted ซึ่งเก็บไว้ครบ
 */
function usedMemberNos_() {
  var used = {};
  rows_(SH.RUNNERS).forEach(function (r) {
    var n = parseInt(codeOf_(r), 10);
    if (!isNaN(n) && n > 0) used[n] = true;
  });
  return used;
}

function firstFreeNo_(used, from) {
  var n = from || 1;
  while (used[n]) n++;
  return n;
}

function nextMemberNo_() {
  var n = firstFreeNo_(usedMemberNos_(), 1);
  setSetting_('memberSeq', n);   // เก็บไว้ดูเฉยๆ ไม่ได้ใช้ตัดสินใจแล้ว
  return pad4_(n);
}

/**
 * เติมรหัสให้คนที่ยังไม่มี และแปลงรหัสรูปแบบเก่า (V1234) ให้เป็นเลขลำดับ
 * เรียงตามวันที่สมัคร คนสมัครก่อนได้เลขน้อยกว่า
 */
function backfillCodes_() {
  return withLock_(function () {
    var sh = sheet_(SH.RUNNERS);
    var last = sh.getLastRow();
    if (last < 2) return 0;

    var codeCol = HEADERS.Runners.indexOf('code') + 1;
    var createdCol = HEADERS.Runners.indexOf('createdAt') + 1;
    var codes = sh.getRange(2, codeCol, last - 1, 1).getValues();
    var created = sh.getRange(2, createdCol, last - 1, 1).getValues();

    // แถวไหนยังไม่มีเลขลำดับ (ว่าง หรือเป็นรหัสแบบเก่า) และเลขไหนถูกใช้ไปแล้ว
    var need = [], used = {};
    for (var i = 0; i < codes.length; i++) {
      var c = String(codes[i][0] == null ? '' : codes[i][0]).replace(/^'/, '').trim();
      if (/^[0-9]+$/.test(c)) {
        used[parseInt(c, 10)] = true;
      } else {
        need.push({ row: i, created: String(created[i][0] || '') });
      }
    }
    if (!need.length) return 0;

    // คนที่สมัครก่อน ได้เลขน้อยกว่า
    need.sort(function (a, b) { return a.created.localeCompare(b.created); });

    // เติมเลขที่ว่างจากน้อยไปมาก ให้เรียงติดกัน ไม่ข้ามเลข
    var n = 1, last = 0;
    need.forEach(function (item) {
      n = firstFreeNo_(used, n);
      used[n] = true;
      codes[item.row][0] = "'" + pad4_(n);   // ' = บังคับเก็บเป็นข้อความ เลข 0 นำหน้าไม่หาย
      last = n;
    });

    sh.getRange(2, codeCol, codes.length, 1).setValues(codes);
    setSetting_('memberSeq', last);
    return need.length;
  });
}

function shortCode_() {
  var chars = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ', s = '';
  for (var i = 0; i < 6; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

function nowIso_() {
  return new Date().toISOString();
}

function normPhone_(p) {
  return String(p || '').replace(/[^0-9]/g, '');
}

/** ลบทุกแถวที่คอลัมน์ที่ระบุมีค่าตรงกับ value — คืนจำนวนแถวที่ลบ */
function deleteRowsWhere_(sheetName, colIndex, value) {
  var sh = sheet_(sheetName);
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, colIndex, last - 1, 1).getValues();
  var targets = [];
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === String(value)) targets.push(i + 2);
  }
  // ลบจากล่างขึ้นบน ไม่งั้นเลขแถวจะเลื่อนหลังลบแถวแรก
  for (var j = targets.length - 1; j >= 0; j--) sh.deleteRow(targets[j]);
  return targets.length;
}

/** ลิงก์ตรงไปยังชีต Deleted (ประวัติการลบสมาชิก) */
function deletedSheetUrl_() {
  try {
    return getSS_().getUrl() + '#gid=' + sheet_(SH.DELETED).getSheetId();
  } catch (err) {
    return getSS_().getUrl();
  }
}

var MAX_IMG_CHARS = 45000;  // เซลล์ Google Sheets เก็บได้ ~50,000 ตัวอักษร

/** อ่านรูปของรางวัลทั้งหมด -> { "3": "data:image/...", ... } */
function tierImages_() {
  var out = {};
  rows_(SH.TIER_IMAGES).forEach(function (r) {
    var k = String(r.tierN), v = String(r.image || '');
    if (k && v) out[k] = v;
  });
  return out;
}

/** บันทึก/ลบรูปของรางวัลระดับหนึ่ง (dataUri ว่าง = ลบรูป) */
function setTierImage_(tierN, dataUri) {
  var sh = sheet_(SH.TIER_IMAGES);
  var last = sh.getLastRow();
  var target = 0;
  if (last >= 2) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === String(tierN)) { target = i + 2; break; }
    }
  }
  if (!dataUri) {
    if (target) sh.deleteRow(target);
    return;
  }
  if (target) {
    sh.getRange(target, 2).setValue(dataUri);
    sh.getRange(target, 3).setValue(nowIso_());
  } else {
    append_(SH.TIER_IMAGES, { tierN: String(tierN), image: dataUri, updatedAt: nowIso_() });
  }
}

/**
 * รูปโปรไฟล์เก็บแยกชีต ไม่ปนกับ Runners
 * เพราะถ้าเก็บรวม เวลาดึงรายชื่อทีมงานจะลากรูปทุกคนมาด้วย = ช้ามาก
 * ดึงทีละคนเฉพาะตอนที่ต้องใช้จริง
 */
function photoOf_(runnerId) {
  var r = rows_(SH.PHOTOS).filter(function (x) {
    return String(x.runnerId) === String(runnerId);
  })[0];
  return r ? String(r.image || '') : '';
}

/** คนไหนมีรูปแล้วบ้าง (ไม่ดึงตัวรูป — แค่บอกว่ามี) */
function photoFlags_() {
  var out = {};
  rows_(SH.PHOTOS).forEach(function (r) {
    if (r.image) out[String(r.runnerId)] = true;
  });
  return out;
}

function setPhoto_(runnerId, dataUri) {
  var sh = sheet_(SH.PHOTOS);
  var last = sh.getLastRow();
  var target = 0;
  if (last >= 2) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]) === String(runnerId)) { target = i + 2; break; }
    }
  }
  if (!dataUri) {
    if (target) sh.deleteRow(target);
    return;
  }
  if (target) {
    sh.getRange(target, 2).setValue(dataUri);
    sh.getRange(target, 3).setValue(thaiTime_());
  } else {
    append_(SH.PHOTOS, { runnerId: String(runnerId), image: dataUri, updatedAt: thaiTime_() });
  }
}

/** จำนวนที่รับได้ของกิจกรรม — 0 หรือว่าง = ไม่จำกัด */
function eventLimit_(ev) {
  var n = parseInt(ev.limit, 10);
  return (isNaN(n) || n < 0) ? 0 : n;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); }
  finally { lock.releaseLock(); }
}

/* ============================ Public API ============================ */

/** ข้อมูลสำหรับนักวิ่ง — ไม่มี PIN, ไม่คืนข้อมูลคนอื่น */
function apiRunnerState(runnerId, eventCode) {
  var s = getSettings_();
  var out = {
    clubName: s.clubName,
    tiers: s.tiers,
    runner: null,
    count: 0,
    history: [],
    redeemed: {},
    event: null,
    alreadyCheckedIn: false,
    images: tierImages_()
  };

  if (eventCode) {
    var ev = rows_(SH.EVENTS).filter(function (e) {
      return String(e.code).toUpperCase() === String(eventCode).toUpperCase();
    })[0];
    if (ev && String(ev.active) !== 'false' && ev.active !== false) {
      var lim = eventLimit_(ev);
      var used = rows_(SH.CHECKINS).filter(function (c) {
        return String(c.eventId) === String(ev.id);
      }).length;
      out.event = {
        id: ev.id, code: ev.code, name: ev.name, date: formatDate_(ev.date),
        limit: lim, used: used, full: (lim > 0 && used >= lim)
      };
    }
  }

  if (!runnerId) return out;

  var runner = rows_(SH.RUNNERS).filter(function (r) { return String(r.id) === String(runnerId); })[0];
  if (!runner) return out;
  out.runner = {
    id: runner.id, name: runner.name, phone: phoneDisplay_(runner.phone),
    code: codeOf_(runner),
    photo: photoOf_(runnerId)
  };

  var evMap = {};
  rows_(SH.EVENTS).forEach(function (e) { evMap[String(e.id)] = e; });

  var mine = rows_(SH.CHECKINS).filter(function (c) { return String(c.runnerId) === String(runnerId); });
  out.count = mine.length;
  out.history = mine.map(function (c) {
    var e = evMap[String(c.eventId)];
    return { name: e ? e.name : '(กิจกรรมถูกลบ)', ts: String(c.ts) };
  }).sort(function (a, b) { return String(b.ts).localeCompare(String(a.ts)); }).slice(0, 20);

  if (out.event) {
    out.alreadyCheckedIn = mine.some(function (c) { return String(c.eventId) === String(out.event.id); });
  }

  rows_(SH.REDEMPTIONS).forEach(function (r) {
    if (String(r.runnerId) === String(runnerId)) out.redeemed[String(r.tierN)] = true;
  });

  return out;
}

/** เช็คว่าเบอร์นี้เคยสมัครไว้หรือยัง — ใช้เป็นหน้าจอแรกของนักวิ่ง */
function apiCheckPhone(phone) {
  var key = phoneKey_(phone);
  if (key.length < 8) throw new Error('กรุณากรอกเบอร์โทรให้ครบ');
  var r = rows_(SH.RUNNERS).filter(function (x) { return phoneKey_(x.phone) === key; })[0];
  if (!r) return { found: false };
  return { found: true, runnerId: String(r.id), name: String(r.name) };
}

/** สมัครสมาชิกใหม่ — ถ้าเบอร์ซ้ำจะคืนสมาชิกเดิม ไม่สร้างใหม่ */
function apiRegister(name, phone) {
  name = String(name || '').trim().slice(0, 40);
  if (!name) throw new Error('กรุณากรอกชื่อ');
  var key = phoneKey_(phone);
  var ph = normPhone_(phone).slice(0, 15);

  return withLock_(function () {
    // 1 เบอร์ = 1 บัญชี — เบอร์ที่มีคนใช้แล้ว สมัครใหม่ไม่ได้
    // (ถ้าเป็นเจ้าของเบอร์เอง ให้กลับไปหน้าใส่เบอร์ ระบบจะพากลับเข้าบัญชีเดิมให้)
    if (key) {
      var existing = rows_(SH.RUNNERS).filter(function (r) { return phoneKey_(r.phone) === key; })[0];
      if (existing) {
        throw new Error('PHONE_TAKEN|เบอร์นี้สมัครไว้แล้วในชื่อ "' + String(existing.name) +
          '" (รหัส ' + codeOf_(existing) + ') — 1 เบอร์ใช้ได้ 1 บัญชี' +
          ' ถ้าเป็นคุณเอง กดถัดไปอีกครั้งเพื่อเข้าบัญชีเดิม');
      }
    }
    var id = uid_(), code = nextMemberNo_();
    // ' นำหน้า = บังคับ Sheets เก็บเป็นข้อความ เลข 0 จะไม่หาย
    append_(SH.RUNNERS, { id: id, name: name, phone: "'" + ph, createdAt: nowIso_(), code: "'" + code });
    return { runnerId: id, name: name, code: code, recovered: false };
  });
}

/** กู้บัตรคืนด้วยเบอร์โทร (เปลี่ยนเครื่อง / ล้างเบราว์เซอร์) */
function apiRecover(phone) {
  var key = phoneKey_(phone);
  if (key.length < 8) throw new Error('กรุณากรอกเบอร์โทรให้ครบ');
  var r = rows_(SH.RUNNERS).filter(function (x) { return phoneKey_(x.phone) === key; })[0];
  if (!r) throw new Error('ไม่พบเบอร์นี้ในระบบ — ถ้ายังไม่เคยสมัคร ให้กดสมัครใหม่');
  return { runnerId: String(r.id), name: String(r.name) };
}

/** เช็คอิน — กันซ้ำต่อ 1 คน 1 กิจกรรม */
function apiCheckin(runnerId, eventCode) {
  if (!runnerId) throw new Error('ยังไม่ได้ลงทะเบียน');

  return withLock_(function () {
    var runner = rows_(SH.RUNNERS).filter(function (r) { return String(r.id) === String(runnerId); })[0];
    if (!runner) throw new Error('ไม่พบสมาชิก กรุณาลงทะเบียนใหม่');

    var ev = rows_(SH.EVENTS).filter(function (e) {
      return String(e.code).toUpperCase() === String(eventCode || '').toUpperCase();
    })[0];
    if (!ev) throw new Error('ไม่พบกิจกรรมนี้');
    if (String(ev.active) === 'false' || ev.active === false) throw new Error('กิจกรรมนี้ปิดรับเช็คอินแล้ว');

    var allCheckins = rows_(SH.CHECKINS);
    var mine = allCheckins.filter(function (c) { return String(c.runnerId) === String(runnerId); });
    var dup = mine.some(function (c) { return String(c.eventId) === String(ev.id); });

    // โควตาของกิจกรรม — คนที่เช็คอินไปแล้วยังเข้าซ้ำได้ ไม่นับเพิ่ม
    var lim = eventLimit_(ev);
    if (!dup && lim > 0) {
      var usedSlots = allCheckins.filter(function (c) {
        return String(c.eventId) === String(ev.id);
      }).length;
      if (usedSlots >= lim) {
        throw new Error('กิจกรรมนี้เต็มแล้ว (รับ ' + lim + ' คน) — แจ้งทีมงานหน้างาน');
      }
    }

    if (!dup) {
      append_(SH.CHECKINS, {
        id: uid_(), runnerId: runnerId, eventId: ev.id, ts: nowIso_(),
        'วันเวลา': thaiTime_(),
        'ชื่อ': String(runner.name),
        'เบอร์โทร': "'" + phoneDisplay_(runner.phone),
        'รหัสสมาชิก': "'" + codeOf_(runner),
        'กิจกรรม': String(ev.name)
      });
      mine.push({});
    }

    var count = mine.length;
    var tiers = getSettings_().tiers.slice().sort(function (a, b) { return a.n - b.n; });
    var hit = tiers.filter(function (t) { return t.n === count; })[0] || null;
    var next = tiers.filter(function (t) { return t.n > count; })[0] || null;

    var totalHere = rows_(SH.CHECKINS).filter(function (c) {
      return String(c.eventId) === String(ev.id);
    }).length;

    return {
      duplicate: dup,
      count: count,
      eventName: String(ev.name),
      hitTier: hit,
      nextTier: next,
      hitImage: hit ? (tierImages_()[String(hit.n)] || '') : '',
      eventUsed: totalHere,
      eventLimit: lim
    };
  });
}

/** ดึงข้อมูลส่วนตัวของนักวิ่ง (เจ้าของเครื่องเท่านั้น — ใช้ runnerId ที่เก็บในเครื่อง) */
function apiGetProfile(runnerId) {
  if (!runnerId) throw new Error('ยังไม่ได้ลงทะเบียน');
  var r = rows_(SH.RUNNERS).filter(function (x) { return String(x.id) === String(runnerId); })[0];
  if (!r) throw new Error('ไม่พบสมาชิก');
  return {
    code: codeOf_(r),
    name: String(r.name || ''),
    fullName: String(r['ชื่อ-นามสกุล'] || ''),
    phone: phoneDisplay_(r.phone),
    birthDate: String(r['วันเกิด'] || '').replace(/^'/, ''),
    photo: photoOf_(runnerId)
  };
}

/**
 * บันทึกข้อมูลส่วนตัว
 * เบอร์โทรเป็นตัวยืนยันตัวตนหลัก จึงห้ามซ้ำกับสมาชิกคนอื่น
 */
function apiUpdateProfile(runnerId, data) {
  if (!runnerId) throw new Error('ยังไม่ได้ลงทะเบียน');
  data = data || {};

  var name = String(data.name || '').trim().slice(0, 40);
  if (!name) throw new Error('กรุณากรอกชื่อเล่น');

  // ส่งมาเฉพาะช่องไหน แก้เฉพาะช่องนั้น — ไม่ส่งมา = เก็บของเดิมไว้
  // (กันข้อมูลหายเวลาเรียกแบบไม่ครบทุกช่อง)
  var hasFullName = Object.prototype.hasOwnProperty.call(data, 'fullName');
  var hasBirth = Object.prototype.hasOwnProperty.call(data, 'birthDate');

  var fullName = String(data.fullName || '').trim().slice(0, 80);
  var birth = String(data.birthDate || '').trim().slice(0, 10);
  if (birth && !/^\d{4}-\d{2}-\d{2}$/.test(birth)) throw new Error('วันเกิดไม่ถูกต้อง');
  if (birth) {
    var y = parseInt(birth.slice(0, 4), 10);
    var nowY = new Date().getFullYear();
    if (y < 1900 || y > nowY) throw new Error('ปีเกิดไม่ถูกต้อง');
  }

  var newKey = phoneKey_(data.phone);
  if (!newKey || newKey.length < 8) throw new Error('กรุณากรอกเบอร์โทรให้ครบ');

  return withLock_(function () {
    var all = rows_(SH.RUNNERS);
    var meIdx = -1;
    for (var i = 0; i < all.length; i++) {
      if (String(all[i].id) === String(runnerId)) { meIdx = i; break; }
    }
    if (meIdx < 0) throw new Error('ไม่พบสมาชิก');

    // เบอร์นี้มีคนอื่นใช้อยู่หรือยัง
    for (var j = 0; j < all.length; j++) {
      if (j !== meIdx && phoneKey_(all[j].phone) === newKey) {
        throw new Error('เบอร์นี้มีสมาชิกคนอื่นใช้อยู่แล้ว');
      }
    }

    var sh = sheet_(SH.RUNNERS);
    var row = meIdx + 2;
    var H = HEADERS.Runners;
    sh.getRange(row, H.indexOf('name') + 1).setValue(name);
    sh.getRange(row, H.indexOf('phone') + 1).setValue("'" + normPhone_(data.phone));
    if (hasFullName) sh.getRange(row, H.indexOf('ชื่อ-นามสกุล') + 1).setValue(fullName);
    if (hasBirth) sh.getRange(row, H.indexOf('วันเกิด') + 1).setValue(birth ? "'" + birth : '');
    sh.getRange(row, H.indexOf('updatedAt') + 1).setValue(thaiTime_());

    var me = all[meIdx];
    return {
      ok: true, name: name,
      fullName: hasFullName ? fullName : String(me['ชื่อ-นามสกุล'] || ''),
      phone: phoneDisplay_(data.phone),
      birthDate: hasBirth ? birth : String(me['วันเกิด'] || '').replace(/^'/, '')
    };
  });
}

/**
 * อัปโหลด/ลบรูปโปรไฟล์ของตัวเอง (ส่งค่าว่าง = ลบรูป)
 * นักวิ่งเป็นคนตัดสินใจเองว่าจะใส่หรือไม่ใส่ — ไม่บังคับ
 */
function apiSetPhoto(runnerId, dataUri) {
  if (!runnerId) throw new Error('ยังไม่ได้ลงทะเบียน');
  var r = rows_(SH.RUNNERS).filter(function (x) { return String(x.id) === String(runnerId); })[0];
  if (!r) throw new Error('ไม่พบสมาชิก');

  var img = String(dataUri || '');
  if (img) {
    if (img.indexOf('data:image/') !== 0) throw new Error('ไฟล์ไม่ใช่รูปภาพ');
    if (img.length > MAX_IMG_CHARS) throw new Error('รูปใหญ่เกินไป ลองถ่ายใหม่');
  }
  return withLock_(function () {
    setPhoto_(runnerId, img);
    return { ok: true, hasPhoto: !!img };
  });
}

/** ทีมงานขอดูรูปทีละคน ตอนยืนยันตัวตนก่อนแจกของ */
function apiGetRunnerPhoto(pin, runnerId) {
  checkPin_(pin);
  return { photo: photoOf_(runnerId) };
}

/**
 * ทีมงานเพิ่มเช็คอินย้อนหลังให้คนที่ลืม / มาแล้วแต่สแกนไม่ทัน
 * ตั้งใจให้ทำได้เฉพาะทีมงาน และบันทึกไว้ในชีตว่าเป็นการเพิ่มเอง
 * เพื่อให้ตรวจย้อนหลังได้ว่าแถวไหนมาจาก QR แถวไหนทีมงานเพิ่มให้
 */
function apiManualCheckin(pin, runnerId, eventId) {
  checkPin_(pin);
  if (!runnerId || !eventId) throw new Error('เลือกสมาชิกและกิจกรรมก่อน');

  return withLock_(function () {
    var runner = rows_(SH.RUNNERS).filter(function (x) { return String(x.id) === String(runnerId); })[0];
    if (!runner) throw new Error('ไม่พบสมาชิก');

    var ev = rows_(SH.EVENTS).filter(function (e) { return String(e.id) === String(eventId); })[0];
    if (!ev) throw new Error('ไม่พบกิจกรรม');

    var all = rows_(SH.CHECKINS);
    var dup = all.some(function (c) {
      return String(c.runnerId) === String(runnerId) && String(c.eventId) === String(eventId);
    });
    if (dup) throw new Error(String(runner.name) + ' เช็คอินกิจกรรมนี้ไปแล้ว');

    // กิจกรรมที่ปิดแล้วยังเพิ่มย้อนหลังได้ (นั่นคือจุดประสงค์)
    // แต่โควตายังต้องเคารพ ไม่งั้นตัวเลขจะเพี้ยนกับที่ประกาศไว้
    var lim = eventLimit_(ev);
    if (lim > 0) {
      var used = all.filter(function (c) { return String(c.eventId) === String(eventId); }).length;
      if (used >= lim) throw new Error('กิจกรรมนี้เต็มโควตาแล้ว (' + lim + ' คน) — ขยายจำนวนก่อนถ้าต้องการเพิ่ม');
    }

    append_(SH.CHECKINS, {
      id: uid_(), runnerId: runnerId, eventId: ev.id, ts: nowIso_(),
      'วันเวลา': thaiTime_(),
      'ชื่อ': String(runner.name),
      'เบอร์โทร': "'" + phoneDisplay_(runner.phone),
      'รหัสสมาชิก': "'" + codeOf_(runner),
      'กิจกรรม': String(ev.name),
      'เพิ่มโดยทีมงาน': 'ใช่'
    });

    var count = rows_(SH.CHECKINS).filter(function (c) {
      return String(c.runnerId) === String(runnerId);
    }).length;

    return { ok: true, name: String(runner.name), eventName: String(ev.name), count: count };
  });
}

/** ลบการเช็คอินรายครั้ง (เผลอเพิ่มผิดคน / เช็คอินมั่ว) */
function apiRemoveCheckin(pin, runnerId, eventId) {
  checkPin_(pin);
  return withLock_(function () {
    var sh = sheet_(SH.CHECKINS);
    var last = sh.getLastRow();
    if (last < 2) throw new Error('ไม่พบรายการ');
    var vals = sh.getRange(2, 2, last - 1, 2).getValues();   // runnerId, eventId
    for (var i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]) === String(runnerId) && String(vals[i][1]) === String(eventId)) {
        sh.deleteRow(i + 2);
        return { ok: true };
      }
    }
    throw new Error('ไม่พบรายการเช็คอินนี้');
  });
}

/** รายชื่อคนที่เช็คอินกิจกรรมหนึ่งๆ — ใช้ตอนทีมงานตรวจสอบ/แก้ไข */
function apiEventCheckins(pin, eventId) {
  checkPin_(pin);
  var runners = {};
  rows_(SH.RUNNERS).forEach(function (r) { runners[String(r.id)] = r; });
  return rows_(SH.CHECKINS)
    .filter(function (c) { return String(c.eventId) === String(eventId); })
    .map(function (c) {
      var r = runners[String(c.runnerId)];
      return {
        runnerId: String(c.runnerId),
        name: r ? String(r.name) : String(c['ชื่อ'] || '(ลบแล้ว)'),
        code: r ? codeOf_(r) : String(c['รหัสสมาชิก'] || '').replace(/^'/, ''),
        phone: r ? phoneDisplay_(r.phone) : '',
        when: String(c['วันเวลา'] || ''),
        manual: String(c['เพิ่มโดยทีมงาน'] || '') === 'ใช่'
      };
    })
    .sort(function (a, b) { return String(a.when).localeCompare(String(b.when)); });
}

/* ============================ Staff API (ต้องใส่ PIN ทุกครั้ง) ============================ */

function apiStaffState(pin) {
  var s = checkPin_(pin);
  backfillCodes_();                 // สมาชิกเก่าที่ยังไม่มีรหัส ให้เติมให้ครบก่อน
  var runners = rows_(SH.RUNNERS);
  var checkins = rows_(SH.CHECKINS);
  var redemptions = rows_(SH.REDEMPTIONS);

  var countBy = {};
  checkins.forEach(function (c) {
    var k = String(c.runnerId);
    countBy[k] = (countBy[k] || 0) + 1;
  });
  var redeemed = {};
  redemptions.forEach(function (r) { redeemed[String(r.runnerId) + '|' + String(r.tierN)] = true; });

  var perEvent = {};
  checkins.forEach(function (c) {
    var k = String(c.eventId);
    perEvent[k] = (perEvent[k] || 0) + 1;
  });

  var hasPhoto = photoFlags_();
  var tiers = s.tiers.slice().sort(function (a, b) { return a.n - b.n; });

  // สต็อกของรางวัล: แจกไปแล้วกี่ชิ้น / ค้างรอแจกกี่ชิ้น / เหลือเท่าไหร่
  var givenByTier = {};
  redemptions.forEach(function (x) {
    var k = String(x.tierN);
    givenByTier[k] = (givenByTier[k] || 0) + 1;
  });

  // ชื่อไหนซ้ำกันบ้าง — หน้าจอจะได้เตือนให้เช็ครหัส/เบอร์ก่อนแจกของ
  var nameCount = {};
  runners.forEach(function (r) {
    var k = String(r.name).trim();
    nameCount[k] = (nameCount[k] || 0) + 1;
  });

  var pending = [];
  runners.forEach(function (r) {
    var n = countBy[String(r.id)] || 0;
    tiers.forEach(function (t) {
      if (n >= t.n && !redeemed[String(r.id) + '|' + t.n]) {
        pending.push({
          runnerId: String(r.id), name: String(r.name), code: codeOf_(r),
          phone: phoneDisplay_(r.phone), tierN: t.n, reward: t.reward,
          hasPhoto: !!hasPhoto[String(r.id)],
          dupName: nameCount[String(r.name).trim()] > 1
        });
      }
    });
  });

  // ประวัติการแจกของรางวัล — ล่าสุดอยู่บนสุด
  var tierName = {};
  tiers.forEach(function (t) { tierName[String(t.n)] = t.reward; });
  var given = redemptions.map(function (x) {
    var rr = runners.filter(function (r) { return String(r.id) === String(x.runnerId); })[0];
    return {
      runnerId: String(x.runnerId),
      tierN: parseInt(x.tierN, 10) || 0,
      name: rr ? String(rr.name) : String(x['ชื่อ'] || '(ลบแล้ว)'),
      code: rr ? codeOf_(rr) : String(x['รหัสสมาชิก'] || '').replace(/^'/, ''),
      phone: rr ? phoneDisplay_(rr.phone) : phoneDisplay_(x['เบอร์โทร']),
      reward: String(x['ของรางวัล'] || tierName[String(x.tierN)] || ('ระดับ ' + x.tierN)),
      ts: String(x.ts)
    };
  }).sort(function (a, b) { return String(b.ts).localeCompare(String(a.ts)); });

  return {
    webAppUrl: getWebAppUrl_(),
    sheetUrl: getSS_().getUrl(),
    deletedUrl: deletedSheetUrl_(),
    deletedCount: rows_(SH.DELETED).length,
    clubName: s.clubName,
    adminPin: s.adminPin,
    tiers: tiers.map(function (t) {
      var given = givenByTier[String(t.n)] || 0;
      var waiting = runners.filter(function (r) {
        var n = countBy[String(r.id)] || 0;
        return n >= t.n && !redeemed[String(r.id) + '|' + t.n];
      }).length;
      var stock = parseInt(t.stock, 10) || 0;
      return {
        n: t.n, reward: t.reward, stock: stock,
        given: given, waiting: waiting,
        left: stock > 0 ? (stock - given) : null,
        short: stock > 0 ? Math.max(0, (given + waiting) - stock) : 0
      };
    }),
    images: tierImages_(),
    events: rows_(SH.EVENTS).map(function (e) {
      var lim = eventLimit_(e);
      var used = perEvent[String(e.id)] || 0;
      return {
        id: String(e.id), code: String(e.code), name: String(e.name),
        date: formatDate_(e.date),
        active: !(String(e.active) === 'false' || e.active === false),
        count: used, limit: lim, full: (lim > 0 && used >= lim),
        createdAt: String(e.createdAt)
      };
    }).sort(function (a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); }),
    runners: runners.map(function (r) {
      return {
        id: String(r.id), name: String(r.name), phone: phoneDisplay_(r.phone),
        code: codeOf_(r), count: countBy[String(r.id)] || 0,
        fullName: String(r['ชื่อ-นามสกุล'] || ''),
        birthDate: String(r['วันเกิด'] || '').replace(/^'/, ''),
        hasPhoto: !!hasPhoto[String(r.id)],
        dupName: nameCount[String(r.name).trim()] > 1
      };
    }).sort(function (a, b) { return b.count - a.count; }),
    pending: pending,
    given: given
  };
}

function apiCreateEvent(pin, name, date, limit) {
  checkPin_(pin);
  name = String(name || '').trim().slice(0, 80);
  if (!name) throw new Error('กรุณาใส่ชื่อกิจกรรม');
  var lim = parseInt(limit, 10);
  if (isNaN(lim) || lim < 0) lim = 0;   // 0 = ไม่จำกัดจำนวน
  return withLock_(function () {
    var used = {};
    rows_(SH.EVENTS).forEach(function (e) { used[String(e.code).toUpperCase()] = true; });
    var code = shortCode_(), guard = 0;
    while (used[code] && guard++ < 50) code = shortCode_();
    var id = uid_();
    append_(SH.EVENTS, {
      id: id, code: code, name: name,
      date: String(date || '').slice(0, 10),
      active: true, createdAt: nowIso_(), limit: lim
    });
    return { id: id, code: code, limit: lim };
  });
}

/** แก้จำนวนที่รับได้ของกิจกรรมที่สร้างไปแล้ว */
function apiSetEventLimit(pin, eventId, limit) {
  checkPin_(pin);
  var lim = parseInt(limit, 10);
  if (isNaN(lim) || lim < 0) lim = 0;
  return withLock_(function () {
    var used = rows_(SH.CHECKINS).filter(function (c) {
      return String(c.eventId) === String(eventId);
    }).length;
    if (lim > 0 && lim < used) {
      throw new Error('ตั้งได้ไม่ต่ำกว่าจำนวนที่เช็คอินไปแล้ว (' + used + ' คน)');
    }
    var sh = sheet_(SH.EVENTS);
    var last = sh.getLastRow();
    if (last < 2) throw new Error('ไม่พบกิจกรรม');
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    var col = HEADERS.Events.indexOf('limit') + 1;
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(eventId)) {
        sh.getRange(i + 2, col).setValue(lim);
        return { ok: true, limit: lim };
      }
    }
    throw new Error('ไม่พบกิจกรรม');
  });
}

function apiSetEventActive(pin, eventId, active) {
  checkPin_(pin);
  return withLock_(function () {
    var sh = sheet_(SH.EVENTS);
    var last = sh.getLastRow();
    if (last < 2) throw new Error('ไม่พบกิจกรรม');
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(eventId)) {
        sh.getRange(i + 2, 5).setValue(active ? true : false);
        return { ok: true };
      }
    }
    throw new Error('ไม่พบกิจกรรม');
  });
}

/**
 * ลบกิจกรรม (และ QR ของกิจกรรมนั้น) ออกจากระบบ
 *
 * ตั้งใจให้ "ลบยาก" ถ้ามีคนเช็คอินไปแล้ว เพราะการลบจะทำให้แต้มของสมาชิกหายไปด้วย
 * ต้องส่ง confirmCount ที่ตรงกับจำนวนคนที่เช็คอินจริง เป็นการยืนยันว่ารู้ตัว
 * ถ้าแค่ต้องการหยุดใช้ QR ให้ "ปิดกิจกรรม" แทน แต้มจะยังอยู่ครบ
 */
function apiDeleteEvent(pin, eventId, confirmCount) {
  checkPin_(pin);
  return withLock_(function () {
    var ev = rows_(SH.EVENTS).filter(function (e) { return String(e.id) === String(eventId); })[0];
    if (!ev) throw new Error('ไม่พบกิจกรรมนี้');

    var used = rows_(SH.CHECKINS).filter(function (c) {
      return String(c.eventId) === String(eventId);
    }).length;

    if (used > 0 && parseInt(confirmCount, 10) !== used) {
      throw new Error('กิจกรรมนี้มีคนเช็คอินแล้ว ' + used + ' คน — ลบแล้วแต้มของทุกคนจะหายไปด้วย' +
        ' ถ้าแค่ไม่อยากให้สแกนได้อีก ให้กด "ปิดกิจกรรม" แทน');
    }

    var removedCheckins = deleteRowsWhere_(SH.CHECKINS, HEADERS.Checkins.indexOf('eventId') + 1, eventId);
    deleteRowsWhere_(SH.EVENTS, 1, eventId);
    return { ok: true, name: String(ev.name || ''), removedCheckins: removedCheckins };
  });
}

function apiSetTiers(pin, tiers) {
  checkPin_(pin);
  var clean = (tiers || [])
    .map(function (t) {
      var st = parseInt(t.stock, 10);
      return {
        n: parseInt(t.n, 10) || 0,
        reward: String(t.reward || '').slice(0, 60),
        stock: (isNaN(st) || st < 0) ? 0 : st      // 0 = ไม่จำกัด/ไม่ได้นับ
      };
    })
    .filter(function (t) { return t.n > 0; })
    .sort(function (a, b) { return a.n - b.n; });
  if (!clean.length) throw new Error('ต้องมีอย่างน้อย 1 ระดับ');
  setSetting_('tiers', JSON.stringify(clean));
  return { ok: true, tiers: clean };
}

function apiSetSettings(pin, clubName, newPin) {
  checkPin_(pin);
  if (clubName) setSetting_('clubName', String(clubName).slice(0, 40));
  if (newPin) {
    // รับได้ทั้งตัวอักษร ตัวเลข ไทย อังกฤษ ยาวได้ถึง 24 ตัว
    var np = String(newPin).trim().slice(0, 24);
    if (!np) throw new Error('รหัสผ่านว่างไม่ได้');
    setSetting_('adminPin', np);
  }
  return { ok: true };
}

/**
 * ลบสมาชิกออกจากระบบ (ใช้ล้างข้อมูลตอนทดสอบ)
 * ลบทั้งตัวสมาชิก การเช็คอิน และประวัติการรับรางวัลของคนนั้น
 * ก่อนลบจะบันทึกลงชีต "Deleted" ไว้เสมอ ว่าลบใคร เบอร์อะไร มีกี่ครั้ง เมื่อไหร่
 */
function apiDeleteRunner(pin, runnerId) {
  checkPin_(pin);
  if (!runnerId) throw new Error('ไม่ได้ระบุสมาชิกที่จะลบ');

  return withLock_(function () {
    var r = rows_(SH.RUNNERS).filter(function (x) { return String(x.id) === String(runnerId); })[0];
    if (!r) throw new Error('ไม่พบสมาชิกคนนี้ (อาจถูกลบไปแล้ว)');

    var n = rows_(SH.CHECKINS).filter(function (c) {
      return String(c.runnerId) === String(runnerId);
    }).length;

    // บันทึกประวัติก่อนลบ เพื่อให้ตรวจย้อนหลังได้ว่าใครหายไปบ้าง
    append_(SH.DELETED, {
      runnerId: String(runnerId),
      name: String(r.name),
      phone: "'" + phoneDisplay_(r.phone),
      checkins: n,
      deletedAt: nowIso_(),
      'รหัสสมาชิก': "'" + codeOf_(r)
    });

    var removedCheckins   = deleteRowsWhere_(SH.CHECKINS, 2, runnerId);     // คอลัมน์ runnerId
    var removedRedemptions = deleteRowsWhere_(SH.REDEMPTIONS, 1, runnerId); // คอลัมน์ runnerId
    deleteRowsWhere_(SH.PHOTOS, 1, runnerId);                               // รูปโปรไฟล์
    deleteRowsWhere_(SH.RUNNERS, 1, runnerId);                              // คอลัมน์ id

    return {
      ok: true,
      name: String(r.name),
      removedCheckins: removedCheckins,
      removedRedemptions: removedRedemptions
    };
  });
}

/**
 * แนบรูปของรางวัลให้ระดับหนึ่ง
 * dataUri: รูปที่ย่อ+บีบอัดมาแล้วจากฝั่งมือถือ (ส่งค่าว่างมาเพื่อลบรูป)
 */
function apiSetTierImage(pin, tierN, dataUri) {
  checkPin_(pin);
  var n = parseInt(tierN, 10);
  if (isNaN(n) || n <= 0) throw new Error('ระดับรางวัลไม่ถูกต้อง');

  var img = String(dataUri || '');
  if (img) {
    if (img.indexOf('data:image/') !== 0) throw new Error('ไฟล์ไม่ใช่รูปภาพ');
    if (img.length > MAX_IMG_CHARS) {
      throw new Error('รูปใหญ่เกินไป ลองเลือกรูปที่เล็กลง');
    }
  }
  return withLock_(function () {
    setTierImage_(n, img);
    return { ok: true, tierN: n, hasImage: !!img };
  });
}

function apiFulfill(pin, runnerId, tierN) {
  var st = checkPin_(pin);
  return withLock_(function () {
    var dup = rows_(SH.REDEMPTIONS).some(function (r) {
      return String(r.runnerId) === String(runnerId) && String(r.tierN) === String(tierN);
    });
    if (dup) return { ok: true, already: true };

    var r = rows_(SH.RUNNERS).filter(function (x) { return String(x.id) === String(runnerId); })[0];
    if (!r) throw new Error('ไม่พบสมาชิกคนนี้');

    var tier = st.tiers.filter(function (t) { return String(t.n) === String(tierN); })[0];
    var rewardName = (tier && tier.reward) ? String(tier.reward) : ('ระดับ ' + tierN);

    // ของหมดสต็อก — กันแจกเกินโดยไม่รู้ตัว (ตั้ง 0 = ไม่นับสต็อก)
    var stock = tier ? (parseInt(tier.stock, 10) || 0) : 0;
    if (stock > 0) {
      var alreadyGiven = rows_(SH.REDEMPTIONS).filter(function (x) {
        return String(x.tierN) === String(tierN);
      }).length;
      if (alreadyGiven >= stock) {
        throw new Error(rewardName + ' หมดสต็อกแล้ว (' + stock + ' ชิ้น) — เพิ่มจำนวนในแท็บรางวัลก่อน');
      }
    }

    append_(SH.REDEMPTIONS, {
      runnerId: runnerId, tierN: tierN, ts: nowIso_(),
      'วันเวลา': thaiTime_(),
      'ชื่อ': String(r.name),
      'เบอร์โทร': "'" + phoneDisplay_(r.phone),
      'รหัสสมาชิก': "'" + codeOf_(r),
      'ของรางวัล': rewardName
    });
    return { ok: true, name: String(r.name), reward: rewardName };
  });
}

/** ยกเลิกการแจก (กดผิดคน) */
function apiUnfulfill(pin, runnerId, tierN) {
  checkPin_(pin);
  return withLock_(function () {
    var sh = sheet_(SH.REDEMPTIONS);
    var last = sh.getLastRow();
    if (last < 2) throw new Error('ไม่พบรายการนี้');
    var vals = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][0]) === String(runnerId) && String(vals[i][1]) === String(tierN)) {
        sh.deleteRow(i + 2);
        return { ok: true };
      }
    }
    throw new Error('ไม่พบรายการนี้');
  });
}

/* ============================ Utils ============================ */

function formatDate_(d) {
  if (d instanceof Date) {
    return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
  }
  return String(d || '').slice(0, 10);
}

/** รันครั้งเดียวตอนติดตั้ง เพื่อสร้าง Google Sheet และให้สิทธิ์ */
function setup() {
  var ss = getSS_();
  initSheets_(ss);
  var filled = backfillCodes_();
  Logger.log('เติมรหัสสมาชิกให้คนเก่า ' + filled + ' คน');
  Logger.log('พร้อมใช้งานแล้ว');
  Logger.log('Google Sheet เก็บข้อมูล: ' + ss.getUrl());
  return ss.getUrl();
}
