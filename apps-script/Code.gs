/**
 * 단어정원 - 구글 시트 연동 스크립트
 *
 * 이 스크립트는 구글 시트를 단어정원 앱의 저장소로 만들어 줍니다.
 *  - 이름이 '_'로 시작하지 않는 시트 탭 하나가 단어장 하나입니다. (A열 단어, B열 뜻, C열 예문, D열 예문 해석)
 *  - '_프로필', '_주간기록', '_학습기록', '_백업' 탭은 앱이 자동으로 관리합니다.
 *
 * 설치 방법은 README.md 의 "2단계"를 보세요.
 */

var PROFILE_SHEET = '_프로필';
var WEEK_SHEET = '_주간기록';
var LOG_SHEET = '_학습기록';
var BACKUP_SHEET = '_백업';
var CHUNK = 40000;

var PROFILE_HEAD = ['ID', '이름', '아바타', '학년', '레벨', '총XP', '이번주시작', '이번주XP', '연속일', '배운단어', '마스터단어', '정답률(%)', '마지막학습일', '갱신시각'];
var WEEK_HEAD = ['주시작(월)', 'ID', '이름', 'XP'];
var LOG_HEAD = ['일시', '이름', '단어장', '종류', '문제수', '정답수', 'XP', '소요시간(초)'];
var BACKUP_HEAD = ['ID', '갱신시각', '조각수', '데이터'];

/** 처음 한 번 실행: 관리용 탭을 만들고 권한을 승인합니다. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  metaSheet_(ss, PROFILE_SHEET, PROFILE_HEAD, [1, 7, 13]);
  metaSheet_(ss, WEEK_SHEET, WEEK_HEAD, [1, 2]);
  metaSheet_(ss, LOG_SHEET, LOG_HEAD, []);
  metaSheet_(ss, BACKUP_SHEET, BACKUP_HEAD, [1]);
  return '준비 완료';
}

function doGet(e) {
  return out_(handle_((e && e.parameter) || {}));
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return out_({ ok: false, error: '요청 형식이 올바르지 않습니다' });
  }
  return out_(handle_(body));
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(b) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
  } catch (err) {
    return { ok: false, error: '잠시 후 다시 시도해 주세요' };
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    switch (b.a) {
      case 'ping':
        return { ok: true, name: ss.getName() };
      case 'sync':
        return sync_(ss, b);
      case 'deck':
        return { ok: true, name: b.name, rows: deckRows_(ss, b.name) };
      case 'putDeck':
        return putDeck_(ss, b.name, b.rows || []);
      case 'delDeck':
        return delDeck_(ss, b.name);
      case 'profiles':
        return { ok: true, profiles: profiles_(ss) };
      case 'loadBackup':
        return { ok: true, data: loadBackup_(ss, b.id) };
      case 'delProfile':
        return delProfile_(ss, b.id);
      default:
        return { ok: false, error: '알 수 없는 요청입니다' };
    }
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 동기화 ---------- */

function sync_(ss, b) {
  if (b.profile && b.profile.id) {
    saveProfile_(ss, b.profile);
    saveWeek_(ss, b.profile);
  }
  if (b.logs && b.logs.length) appendLogs_(ss, b.logs);
  if (b.backup && b.profile && b.profile.id) saveBackup_(ss, b.profile.id, b.backup);
  return { ok: true, decks: deckList_(ss), profiles: profiles_(ss), weeks: weeks_(ss) };
}

/* ---------- 단어장 ---------- */

function isDeck_(name) {
  return name.charAt(0) !== '_';
}

function deckList_(ss) {
  var list = [];
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (!isDeck_(name)) continue;
    if (sheets[i].getLastRow() < 1) continue;
    list.push({ name: name, ver: version_(sheets[i]) });
  }
  return list;
}

function version_(sheet) {
  var values = sheet.getDataRange().getDisplayValues();
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(values), Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var v = (bytes[i] + 256) % 256;
    hex += (v < 16 ? '0' : '') + v.toString(16);
  }
  return hex;
}

function deckRows_(ss, name) {
  var sheet = ss.getSheetByName(String(name));
  if (!sheet || !isDeck_(String(name)) || sheet.getLastRow() < 1) return [];
  var values = sheet.getDataRange().getDisplayValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) rows.push(values[i].slice(0, 8));
  return rows;
}

function cleanName_(name) {
  var s = String(name || '').replace(/[:\\\/\?\*\[\]]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
  if (!s) s = '단어장';
  if (s.charAt(0) === '_') s = '단어장 ' + s;
  return s;
}

function putDeck_(ss, name, rows) {
  name = cleanName_(name);
  var data = [['단어', '뜻', '예문', '예문 해석']];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i] || [];
    if (!r[0] || !r[1]) continue;
    data.push([String(r[0]), String(r[1]), String(r[2] || ''), String(r[3] || '')]);
  }
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear();
  var range = sheet.getRange(1, 1, data.length, 4);
  range.setNumberFormat('@');
  range.setValues(data);
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return { ok: true, name: name, ver: version_(sheet), count: data.length - 1 };
}

function delDeck_(ss, name) {
  var sheet = ss.getSheetByName(String(name));
  if (sheet && isDeck_(String(name))) {
    if (ss.getSheets().length <= 1) metaSheet_(ss, PROFILE_SHEET, PROFILE_HEAD, [1, 7, 13]);
    ss.deleteSheet(sheet);
  }
  return { ok: true };
}

/* ---------- 프로필 / 주간 기록 / 학습 기록 ---------- */

function metaSheet_(ss, name, head, textCols) {
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  sheet = ss.insertSheet(name);
  for (var i = 0; i < textCols.length; i++) {
    sheet.getRange(1, textCols[i], sheet.getMaxRows(), 1).setNumberFormat('@');
  }
  sheet.getRange(1, 1, 1, head.length).setValues([head]);
  sheet.getRange(1, 1, 1, head.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function text_(v, ss) {
  if (v instanceof Date) return Utilities.formatDate(v, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v === null || v === undefined ? '' : v);
}

function num_(v) {
  var n = Number(v);
  return isFinite(n) ? n : 0;
}

function findRow_(sheet, matchFn) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var values = sheet.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    if (matchFn(values[i])) return i + 2;
  }
  return -1;
}

function saveProfile_(ss, p) {
  var sheet = metaSheet_(ss, PROFILE_SHEET, PROFILE_HEAD, [1, 7, 13]);
  var id = String(p.id);
  var row = findRow_(sheet, function (r) { return String(r[0]) === id; });
  if (row < 0) row = sheet.getLastRow() + 1;
  var values = [[
    id, String(p.name || ''), String(p.avatar || ''), String(p.grade || ''),
    num_(p.level), num_(p.xp), String(p.weekKey || ''), num_(p.weekXp), num_(p.streak),
    num_(p.learned), num_(p.mastered), num_(p.acc), String(p.lastDay || ''), new Date()
  ]];
  sheet.getRange(row, 1, 1, PROFILE_HEAD.length).setValues(values);
}

function profiles_(ss) {
  var sheet = ss.getSheetByName(PROFILE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  var list = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    list.push({
      id: String(r[0]), name: String(r[1]), avatar: String(r[2]), grade: String(r[3]),
      level: num_(r[4]), xp: num_(r[5]), weekKey: text_(r[6], ss), weekXp: num_(r[7]), streak: num_(r[8]),
      learned: num_(r[9]), mastered: num_(r[10]), acc: num_(r[11]), lastDay: text_(r[12], ss)
    });
  }
  return list;
}

function delProfile_(ss, id) {
  id = String(id);
  var names = [PROFILE_SHEET, BACKUP_SHEET];
  for (var n = 0; n < names.length; n++) {
    var sheet = ss.getSheetByName(names[n]);
    if (!sheet) continue;
    var row = findRow_(sheet, function (r) { return String(r[0]) === id; });
    if (row > 0) sheet.deleteRow(row);
  }
  return { ok: true };
}

function saveWeek_(ss, p) {
  if (!p.weekKey) return;
  var sheet = metaSheet_(ss, WEEK_SHEET, WEEK_HEAD, [1, 2]);
  var id = String(p.id);
  var week = String(p.weekKey);
  var row = findRow_(sheet, function (r) { return text_(r[0], ss) === week && String(r[1]) === id; });
  if (row < 0) row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 4).setValues([[week, id, String(p.name || ''), num_(p.weekXp)]]);
}

function weeks_(ss) {
  var sheet = ss.getSheetByName(WEEK_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var last = sheet.getLastRow();
  var start = Math.max(2, last - 79);
  var values = sheet.getRange(start, 1, last - start + 1, 4).getValues();
  var list = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[0] || !r[1]) continue;
    list.push({ week: text_(r[0], ss), id: String(r[1]), name: String(r[2]), xp: num_(r[3]) });
  }
  return list;
}

function appendLogs_(ss, logs) {
  var sheet = metaSheet_(ss, LOG_SHEET, LOG_HEAD, []);
  var rows = [];
  for (var i = 0; i < logs.length && i < 200; i++) {
    var l = logs[i] || {};
    rows.push([
      l.at ? new Date(l.at) : new Date(), String(l.name || ''), String(l.deck || ''), String(l.kind || ''),
      num_(l.n), num_(l.ok), num_(l.xp), num_(l.sec)
    ]);
  }
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, LOG_HEAD.length).setValues(rows);
}

/* ---------- 학습 진행 백업 (기기 변경 대비) ---------- */

function saveBackup_(ss, id, json) {
  var sheet = metaSheet_(ss, BACKUP_SHEET, BACKUP_HEAD, [1]);
  id = String(id);
  json = String(json);
  var chunks = [];
  for (var i = 0; i < json.length; i += CHUNK) chunks.push('J' + json.substr(i, CHUNK));
  var need = 3 + chunks.length;
  var maxCols = sheet.getMaxColumns();
  if (need > maxCols) sheet.insertColumnsAfter(maxCols, need - maxCols);
  var row = findRow_(sheet, function (r) { return String(r[0]) === id; });
  if (row < 0) row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, sheet.getMaxColumns()).clearContent();
  sheet.getRange(row, 1, 1, need).setNumberFormat('@');
  sheet.getRange(row, 1, 1, need).setValues([[id, new Date().toISOString(), String(chunks.length)].concat(chunks)]);
}

function loadBackup_(ss, id) {
  var sheet = ss.getSheetByName(BACKUP_SHEET);
  if (!sheet) return null;
  id = String(id);
  var row = findRow_(sheet, function (r) { return String(r[0]) === id; });
  if (row < 0) return null;
  var n = num_(sheet.getRange(row, 3).getValue());
  if (n < 1) return null;
  var cells = sheet.getRange(row, 4, 1, n).getValues()[0];
  var json = '';
  for (var i = 0; i < cells.length; i++) json += String(cells[i]).substr(1);
  return json;
}
