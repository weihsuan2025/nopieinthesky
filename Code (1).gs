/**
 * 情緒導航員｜後端（Google Apps Script + Google 試算表）
 *
 * 這份程式只放在你自己的 Google 帳號裡，【不要】放上 GitHub。
 * 教師帳號與密碼請設定在「專案設定 → 指令碼屬性」，不要寫在程式裡。
 */

const TZ = 'Asia/Taipei';
const DEFAULT_TEACHER_EMAIL = 'g111444011@grad.ntue.edu.tw'; // 也可在指令碼屬性設定 TEACHER_EMAIL 覆蓋
const TESTER_GROUP = '6年0班', TESTER_SEAT = '0'; // 測試帳號（任何學校）：不受次數與間隔限制
const DAILY_LIMIT = 2;               // 每天最多使用次數
const MIN_GAP_MS = 40 * 60 * 1000;   // 兩次使用至少間隔 40 分鐘
const HELP_COOLDOWN_SEC = 60;        // 同一位學生 60 秒內重複按「找大人」不重複寄信
const TOKEN_TTL_SEC = 6 * 60 * 60;   // 教師登入有效 6 小時
const STATES = ['噴火火箭', '發光火箭', '沒電火箭', '飄飄火箭', '尚未選擇']; // 「尚未選擇」＝時間到時還沒選火箭

const SHEETS = {
  records: { name: '紀錄', headers: ['時間戳', '時間', '學生ID', '班級', '座號', '火箭類型', '情緒', '前測', '後測', '求助次數', '方法與前後測(JSON)', '學校', '勳章(JSON)'] },
  starts: { name: '使用', headers: ['學生ID', '時間戳', '時間'] },
  teachers: { name: '教師', headers: ['帳號', 'Email', '鹽', '密碼雜湊', '任教班級(JSON)', '註冊時間', '任教學校'] },
  help: { name: '求助', headers: ['時間戳', '時間', '學生ID', '學生', '情緒', '強度', '已寄信', '學校', '原因'] },
};

/* ---------- 進入點 ---------- */

function doPost(e) {
  let req;
  try {
    const raw = e && e.postData && e.postData.contents;
    if (!raw || raw.length > 30000) return json_({ ok: false, error: 'bad_request' });
    req = JSON.parse(raw);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }
  try {
    const handler = HANDLERS[req.action];
    if (!handler) return json_({ ok: false, error: 'unknown_action' });
    return json_(handler(req));
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server_error' });
  }
}

function doGet() {
  return json_({ ok: true, service: '情緒導航員後端' });
}

const HANDLERS = {
  getStudent: getStudent_,
  startSession: startSession_,
  saveRecord: saveRecord_,
  help: help_,
  teacherRegister: teacherRegister_,
  teacherUpdateProfile: teacherUpdateProfile_,
  teacherLogin: teacherLogin_,
  teacherData: teacherData_,
  teacherLogout: teacherLogout_,
};

/* ---------- 學生端 ---------- */

// 只回傳「使用時間、勳章、方法與前後測」，不回傳情緒與火箭類型，避免別人猜到班級座號就能看到敏感內容。
function getStudent_(req) {
  if (!validId_(req.id)) return { ok: false, error: 'bad_id' };
  const starts = rows_('starts').filter((r) => r[0] === req.id).map((r) => Number(r[1]));
  const mine = rows_('records').filter((r) => r[2] === req.id);
  const badges = {};
  mine.forEach((r) => {
    const b = parse_(r[12], {});
    Object.keys(b).forEach((k) => { badges[k] = Math.max(badges[k] || 0, Number(b[k]) || 0); });
  });
  const records = mine.slice(-100).map((r) => ({ at: Number(r[0]), attempts: parse_(r[10], []) }));
  return { ok: true, starts: starts, badges: badges, records: records };
}

// 後端判斷每日次數與間隔，學生清除瀏覽器資料也繞不過去。
function startSession_(req) {
  if (!validId_(req.id)) return { ok: false, error: 'bad_id', message: '班級或座號格式不正確。' };
  const pid = parseId_(req.id);
  if (pid.group === TESTER_GROUP && pid.seat === TESTER_SEAT) return { ok: true, tester: true, serverNow: Date.now() };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = Date.now();
    const mine = rows_('starts').filter((r) => r[0] === req.id).map((r) => Number(r[1]));
    const todayCount = mine.filter((t) => day_(t) === day_(now)).length;
    if (todayCount >= DAILY_LIMIT) return { ok: false, message: '今天已經使用 ' + DAILY_LIMIT + ' 次囉，明天再來吧！' };
    const last = mine.length ? Math.max.apply(null, mine) : 0;
    const wait = MIN_GAP_MS - (now - last);
    if (wait > 0) return { ok: false, message: '距離上次使用還不到 40 分鐘，請再等 ' + Math.ceil(wait / 60000) + ' 分鐘。' };
    getSheet_('starts').appendRow([req.id, now, fmt_(now)]);
    return { ok: true, serverNow: now };
  } finally {
    lock.releaseLock();
  }
}

function saveRecord_(req) {
  if (!validId_(req.id)) return { ok: false, error: 'bad_id' };
  const record = cleanRecord_(req.record);
  if (!record) return { ok: false, error: 'bad_record' };
  if (!rateOk_('save:' + req.id, 12, 60)) return { ok: false, error: 'rate_limited' };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const pid = parseId_(req.id);
    const sh = getSheet_('records');
    const n = sh.getLastRow();
    // 只讀「時間戳」「時間」「學生ID」三欄來檢查重複，人多同時存檔時比較快
    const keys = n < 2 ? [] : sh.getRange(2, 1, n - 1, 3).getValues();
    const duplicate = keys.some((r) => r[2] === req.id && Number(r[0]) === record.at);
    if (!duplicate) {
      sh.appendRow([
        record.at, fmt_(record.at), safe_(req.id), safe_(pid.group), safe_(pid.seat),
        safe_(record.state), safe_(record.emotion), record.initial, record.final, record.helpCount,
        JSON.stringify(record.attempts), safe_(pid.school), JSON.stringify(cleanBadges_(req.badges)),
      ]);
      SpreadsheetApp.flush(); // 立刻寫入，避免下一筆同時進來時讀不到
    }
    return { ok: true, duplicate: duplicate };
  } finally {
    lock.releaseLock();
  }
}

// 只有「真的寄出信」才回報 emailed: true，前端才會顯示「已通知老師」。
// 前端只在兩種情況自動呼叫：前後測都是 5 級（two-fives）；前測 5 級但時間內沒做完調適方式（timeout）。
const HELP_REASONS = {
  'two-fives': '前測與後測都是 5 級（感受很強），系統已自動通知。',
  timeout: '前測是 5 級，但在使用時間內沒有完成調適方式，系統已自動通知。',
};
function help_(req) {
  const id = validId_(req.id) ? req.id : '';
  const who = oneLine_(req.who || id || '未登入', 40);
  const emotion = oneLine_(req.emotion || '', 10);
  const rating = Math.max(0, Math.min(5, Math.round(Number(req.rating) || 0)));
  const reason = HELP_REASONS[req.reason] ? req.reason : '';
  const cache = CacheService.getScriptCache();
  const key = 'help:' + (id || 'anon');
  if (cache.get(key)) return { ok: true, emailed: true, repeat: true };
  const now = Date.now();
  const pid = parseId_(id);
  const classTeachers = id ? teacherEmailsFor_(pid.school, pid.group) : [];
  const recipients = classTeachers.length ? classTeachers : [teacherEmail_()]; // 該班沒有註冊老師時，寄給管理者
  let emailed = false;
  try {
    const lines = ['學生：' + who, '時間：' + fmt_(now)];
    if (emotion) lines.push('目前情緒：' + emotion + (rating ? '（' + rating + ' 級）' : ''));
    if (reason) lines.push('原因：' + HELP_REASONS[reason]);
    lines.push('', '請盡快到學生身邊關心。');
    MailApp.sendEmail({ to: recipients.join(','), subject: '【情緒導航員】' + who + ' 需要老師關心', body: lines.join('\n') });
    emailed = true;
  } catch (err) {
    console.error(err);
  }
  getSheet_('help').appendRow([now, fmt_(now), safe_(id), safe_(who), safe_(emotion), rating, emailed ? '是' : '否', safe_(pid.school), reason]);
  if (emailed) cache.put(key, '1', HELP_COOLDOWN_SEC);
  return { ok: true, emailed: emailed };
}

/* ---------- 教師端 ---------- */

// 老師註冊：不需要邀請碼，但限制註冊頻率，避免被大量亂註冊。
function teacherRegister_(req) {
  const props = PropertiesService.getScriptProperties();
  if (!rateOk_('register', 20, 3600)) return { ok: false, message: '註冊的人太多，請稍後再試。' };
  const account = String(req.account || '').trim();
  const password = String(req.password || '');
  const email = String(req.email || '').trim();
  const school = normSchool_(req.school);
  const classes = (Array.isArray(req.classes) ? req.classes : []).map(normClass_);
  if (!school) return { ok: false, message: '請輸入學校名稱（2 到 20 個中文字、英文或數字）。' };
  if (!/^[A-Za-z0-9_]{3,20}$/.test(account)) return { ok: false, message: '帳號請用 3–20 個英文字母、數字或底線。' };
  if (password.length < 8 || password.length > 64) return { ok: false, message: '密碼長度需為 8–64 個字。' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 100) return { ok: false, message: 'Email 格式不正確。' };
  if (!classes.length || classes.length > 12 || classes.some((c) => !c)) return { ok: false, message: '任教班級格式不正確，例如：6年丙班。' };
  const adminAccount = props.getProperty('TEACHER_ACCOUNT');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const taken = rows_('teachers').some((r) => String(r[0]).toLowerCase() === account.toLowerCase())
      || (adminAccount && adminAccount.toLowerCase() === account.toLowerCase());
    if (taken) return { ok: false, message: '這個帳號已經有人使用。' };
    const salt = Utilities.getUuid();
    getSheet_('teachers').appendRow([safe_(account), safe_(email), salt, hash_(salt, password), JSON.stringify(classes), fmt_(Date.now()), safe_(school)]);
  } finally {
    lock.releaseLock();
  }
  let emailed = false;
  try {
    MailApp.sendEmail(email, '【情緒導航員】教師帳號註冊成功', '您好，帳號「' + account + '」已註冊成功。\n學校：' + school + '\n任教班級：' + classes.join('、') + '\n學生的情緒強度很高、需要老師關心時，通知信會寄到這個信箱。');
    emailed = true;
  } catch (err) {
    console.error(err);
  }
  return { ok: true, emailed: emailed };
}

// 老師登入後修改自己的任教學校、班級與 Email。
function teacherUpdateProfile_(req) {
  const info = tokenInfo_(req.token);
  if (!info) return { ok: false, error: 'unauthorized' };
  if (info.admin) return { ok: false, message: '管理者帳號的資料請在指令碼屬性修改。' };
  const school = normSchool_(req.school);
  const email = String(req.email || '').trim();
  const classes = (Array.isArray(req.classes) ? req.classes : []).map(normClass_);
  if (!school) return { ok: false, message: '請輸入學校名稱（2 到 20 個中文字、英文或數字）。' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 100) return { ok: false, message: 'Email 格式不正確。' };
  if (!classes.length || classes.length > 12 || classes.some((c) => !c)) return { ok: false, message: '任教班級格式不正確，例如：6年丙班。' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_('teachers');
    const n = sh.getLastRow();
    const accounts = n < 2 ? [] : sh.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < accounts.length; i++) {
      if (String(accounts[i][0]).toLowerCase() === info.account.toLowerCase()) {
        sh.getRange(i + 2, 2).setValue(safe_(email));
        sh.getRange(i + 2, 5).setValue(JSON.stringify(classes));
        sh.getRange(i + 2, 7).setValue(safe_(school));
        const me = { account: info.account, admin: false, school: school, classes: classes, email: email };
        CacheService.getScriptCache().put('tok:' + req.token, JSON.stringify(me), TOKEN_TTL_SEC);
        return { ok: true, me: me };
      }
    }
    return { ok: false, message: '找不到這個教師帳號。' };
  } finally {
    lock.releaseLock();
  }
}

// 管理者（指令碼屬性裡的帳密）看得到全部班級；一般老師只看得到自己註冊的班級。
function teacherLogin_(req) {
  const props = PropertiesService.getScriptProperties();
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('login_fails') || 0);
  if (fails >= 5) return { ok: false, message: '嘗試次數太多，請 10 分鐘後再試。' };
  const account = String(req.account || '').trim();
  const password = String(req.password || '');
  const adminAccount = props.getProperty('TEACHER_ACCOUNT');
  const adminPassword = props.getProperty('TEACHER_PASSWORD');
  let info = null;
  if (adminAccount && adminPassword && same_(account, adminAccount) && same_(password, adminPassword)) {
    info = { account: adminAccount, admin: true, school: '', classes: [], email: teacherEmail_() };
  } else {
    const t = rows_('teachers').find((r) => String(r[0]).toLowerCase() === account.toLowerCase());
    if (t && same_(hash_(String(t[2]), password), String(t[3]))) info = { account: String(t[0]), admin: false, school: String(t[6] || ''), classes: parse_(t[4], []), email: String(t[1]) };
  }
  if (!info) {
    cache.put('login_fails', String(fails + 1), 600);
    return { ok: false, message: '帳號或密碼錯誤。' };
  }
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('tok:' + token, JSON.stringify(info), TOKEN_TTL_SEC);
  cache.remove('login_fails');
  return { ok: true, token: token, me: info };
}

function teacherData_(req) {
  const info = tokenInfo_(req.token);
  if (!info) return { ok: false, error: 'unauthorized' };
  // 舊資料沒有學校欄位時（空字串），只比對班級。
  const mine = (school, group) => info.admin || (info.classes.indexOf(group) >= 0 && (!school || !info.school || school === info.school));
  const byId = {};
  rows_('records').forEach((r) => {
    if (!mine(String(r[11] || ''), String(r[3]))) return;
    const id = r[2];
    if (!byId[id]) byId[id] = { id: id, school: String(r[11] || ''), group: String(r[3]), seat: String(r[4]), records: [] };
    byId[id].records.push({
      at: Number(r[0]), state: r[5], emotion: r[6], initial: Number(r[7]), final: Number(r[8]),
      helpCount: Number(r[9]) || 0, attempts: parse_(r[10], []),
    });
  });
  const help = rows_('help').filter((r) => { const p = parseId_(r[2]); return mine(p.school, p.group); }).slice(-30).reverse().map((r) => ({
    at: Number(r[0]), who: r[3], emotion: r[4], rating: Number(r[5]) || 0,
  }));
  return { ok: true, records: Object.keys(byId).map((k) => byId[k]), help: help, me: info };
}

function teacherLogout_(req) {
  if (req.token) CacheService.getScriptCache().remove('tok:' + req.token);
  return { ok: true };
}

/* ---------- 工具函式 ---------- */

function getSheet_(key) {
  const def = SHEETS[key];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.appendRow(def.headers);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < def.headers.length) {
    sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]); // 舊版工作表補上新欄位標題
  }
  return sh;
}

function rows_(key) {
  const sh = getSheet_(key);
  const n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, SHEETS[key].headers.length).getValues();
}

function cleanRecord_(r) {
  if (!r || typeof r !== 'object') return null;
  const at = Number(r.at);
  if (!isFinite(at) || at <= 0) return null;
  const initial = int_(r.initial, 1, 5);
  const final = int_(r.final, 1, 5);
  if (initial === null || final === null || STATES.indexOf(r.state) < 0) return null;
  const attempts = (Array.isArray(r.attempts) ? r.attempts : []).slice(0, 20).map((a) => ({
    name: oneLine_(a && a.name, 30),
    level: int_(a && a.level, 1, 3),
    pre: int_(a && a.pre, 1, 5),
    post: a && a.post != null ? int_(a.post, 1, 5) : null,
    status: ['done', 'unfinished', 'nopost', 'pending'].indexOf(a && a.status) >= 0 ? a.status : 'unfinished',
  }));
  return { at: at, state: r.state, emotion: oneLine_(r.emotion, 10), initial: initial, final: final, helpCount: int_(r.helpCount || 0, 0, 99) || 0, attempts: attempts };
}

function cleanBadges_(b) {
  const out = {};
  if (b && typeof b === 'object') {
    Object.keys(b).forEach((k) => {
      const v = int_(b[k], 0, 5);
      if (/^(sprint|connection|cruise)-[0-3]$/.test(k) && v !== null) out[k] = v;
    });
  }
  return out;
}

function int_(v, lo, hi) {
  const n = Math.round(Number(v));
  return isFinite(n) && n >= lo && n <= hi ? n : null;
}

function oneLine_(v, max) { return String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').slice(0, max); }
function safe_(v) { const s = String(v == null ? '' : v); return /^[=+\-@]/.test(s) ? "'" + s : s; } // 避免試算表公式注入
function parse_(text, fallback) { try { return JSON.parse(text); } catch (err) { return fallback; } }
// 學生 ID 格式：「學校|6年丙班-8」；舊資料沒有學校，格式為「6年丙班-8」。
const ID_RE = /^(?:([\u4e00-\u9fa5A-Za-z0-9]{2,20})\|)?([1-6]年[\u4e00-\u9fa50-9]{1,4}班)-(\d{1,3})$/;
function validId_(id) { return typeof id === 'string' && ID_RE.test(id); }
function parseId_(id) {
  const m = ID_RE.exec(String(id || ''));
  return m ? { school: m[1] || '', group: m[2], seat: m[3] } : { school: '', group: '', seat: '' };
}
function normSchool_(s) {
  const v = String(s == null ? '' : s).replace(/\s+/g, '');
  return /^[\u4e00-\u9fa5A-Za-z0-9]{2,20}$/.test(v) ? v : '';
}
function fmt_(ms) { return Utilities.formatDate(new Date(ms), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function day_(ms) { return Utilities.formatDate(new Date(ms), TZ, 'yyyy-MM-dd'); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function teacherEmail_() { return PropertiesService.getScriptProperties().getProperty('TEACHER_EMAIL') || DEFAULT_TEACHER_EMAIL; }
function tokenInfo_(token) {
  if (typeof token !== 'string' || token.length < 20) return null;
  const raw = CacheService.getScriptCache().get('tok:' + token);
  return raw ? parse_(raw, null) : null;
}

// 通知對象：同學校、同班級的已註冊老師（舊資料沒有學校時只比對班級；老師沒填學校時也視為符合）。
function teacherEmailsFor_(school, group) {
  const seen = {};
  rows_('teachers').forEach((r) => {
    const tSchool = String(r[6] || '');
    if (parse_(r[4], []).indexOf(group) >= 0 && (!school || !tSchool || school === tSchool)) seen[String(r[1])] = true;
  });
  return Object.keys(seen);
}

function normClass_(s) {
  const m = String(s || '').trim().match(/^([1-6])\s*年?\s*([\u4e00-\u9fa50-9]{1,4}?)\s*班?$/);
  return m ? m[1] + '年' + m[2] + '班' : '';
}

function hash_(salt, password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + password, Utilities.Charset.UTF_8);
  return bytes.map((b) => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function same_(a, b) {
  const x = String(a == null ? '' : a);
  const y = String(b == null ? '' : b);
  let diff = x.length ^ y.length;
  for (let i = 0; i < Math.max(x.length, y.length); i++) diff |= (x.charCodeAt(i) || 0) ^ (y.charCodeAt(i) || 0);
  return diff === 0;
}

function rateOk_(key, max, seconds) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(key) || 0);
  if (n >= max) return false;
  cache.put(key, String(n + 1), seconds);
  return true;
}

/* ---------- 第一次設定時手動執行 ---------- */

// 在編輯器選 authorize 後按「執行」，同意授權（存取試算表、寄信）。
function authorize() {
  SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('今天還能寄出 ' + MailApp.getRemainingDailyQuota() + ' 封信');
  Object.keys(SHEETS).forEach(getSheet_); // 順便建立需要的工作表（紀錄、使用、教師、求助）
}

// 選 sendTestMail 後按「執行」，確認老師信箱收得到信。
function sendTestMail() {
  MailApp.sendEmail(teacherEmail_(), '【情緒導航員】測試信', '收到這封信，代表求助通知功能正常。');
}

// 【只需執行一次】把舊「學生」分頁裡的勳章，搬到「紀錄」分頁（寫在該學生最新一筆紀錄的「勳章(JSON)」欄）。
// 執行完成、確認沒問題後，就可以手動刪除試算表裡的「學生」分頁；程式之後不會再建立或讀取它。
function migrateStudentSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = ss.getSheetByName('學生');
  if (!old) { Logger.log('沒有「學生」分頁，不需要搬移。'); return; }
  const sh = getSheet_('records');
  const n = sh.getLastRow();
  if (n < 2) { Logger.log('「紀錄」還沒有資料，沒有可搬移的對象。'); return; }
  const ids = sh.getRange(2, 3, n - 1, 1).getValues();
  const lastRowOf = {};
  ids.forEach((r, i) => { lastRowOf[r[0]] = i + 2; });
  const m = old.getLastRow();
  const students = m < 2 ? [] : old.getRange(2, 1, m - 1, 4).getValues();
  let moved = 0;
  students.forEach((r) => {
    const row = lastRowOf[r[0]];
    if (!row) return;
    const merged = parse_(r[3], {});
    const cur = parse_(sh.getRange(row, 13).getValue(), {});
    Object.keys(cur).forEach((k) => { merged[k] = Math.max(merged[k] || 0, Number(cur[k]) || 0); });
    sh.getRange(row, 13).setValue(JSON.stringify(merged));
    moved++;
  });
  Logger.log('已搬移 ' + moved + ' 位學生的勳章。確認後可手動刪除「學生」分頁。');
}
