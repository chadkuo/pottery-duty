/**
 * 陶藝班「值日生」線上排班系統 — 後端
 * 執行環境：Google Apps Script（附掛在一份 Google 試算表上）
 *
 * 設計原則：所有「會變的東西」都放在試算表，程式碼不需要為了新學期而修改。
 *   ├─ 設定      學期名稱、第一堂課日期、上課星期時間、總堂數、每週人數、每人最少次數…
 *   ├─ 學員名單   成員（可在試算表新增，也可由同學在網頁自行加入）
 *   ├─ 特殊安排   公民週、國定假日、期末聚餐…哪幾天不用排值日生
 *   ├─ 排班表    由上面三張表自動產生，也可以隨時手改
 *   ├─ 代班需求   當天不能到的人在此徵求代班、記錄誰代誰
 *   └─ 操作紀錄   所有異動的稽核軌跡
 *
 * 換學期只要：改「設定」+「特殊安排」→ 執行選單「值日生 → 依設定重建排班表」
 */

// ---------------------------------------------------------------------------
// 常數
// ---------------------------------------------------------------------------
var TZ = 'Asia/Taipei';

var SHEET_CONFIG   = '設定';
var SHEET_ROSTER   = '學員名單';
var SHEET_SPECIAL  = '特殊安排';
var SHEET_SCHEDULE = '排班表';
var SHEET_SUBS     = '代班需求';
var SHEET_LOG      = '操作紀錄';

// 排班表固定欄位（1-based），第 5 欄之後是值日生欄位（數量由「每週值日生人數」決定）
var COL_WEEK_NO = 1;
var COL_DATE    = 2;
var COL_LABEL   = 3;
var COL_NEEDS   = 4;
var COL_SLOT1   = 5;

// 代班需求狀態
var SUB_PENDING = '徵求中';
var SUB_DONE    = '已代班';
var SUB_VOID    = '已失效';

var WEEKDAY_CH = ['日', '一', '二', '三', '四', '五', '六'];

// LINE Messaging API
var LINE_API            = 'https://api.line.me/v2/bot';
var LINE_TOKEN_KEY      = 'LINE_CHANNEL_TOKEN';    // 放 PropertiesService，不放試算表
var LINE_LAST_PUSH      = 'LINE_LAST_PUSH_DATE';   // 上次推播日期，用來算間隔
var LINE_SECRET_KEY     = 'LINE_WEBHOOK_SECRET';   // webhook 網址上的通行碼，見 doPost 的說明
var LINE_QUOTA_BUFFER   = 20;                      // 額度安全邊際：剩餘不足「群組人數 + 這個數」就不推
var LINE_SIZE_FALLBACK  = 60;                      // 查不到群組人數時的保守估計

// 「設定」工作表的 LINE 欄位。建表時寫入，既有的表用 ensureLineConfigRows() 補。
var LINE_CONFIG_ROWS = [
  ['LINE 群組 ID', '', '（選填）填了才會主動推播。把 bot 加進群組後，在群組打一次關鍵字，ID 會記在「操作紀錄」'],
  ['啟用 LINE 推播', 'FALSE', 'TRUE = 定期主動推播到群組。推播按群組人數計費，免費方案每月 200 則'],
  ['LINE 推播間隔週數', 2, '每隔幾週推一次。58 人的群組推一次 58 則，設 2 約等於每月 116 則'],
  ['LINE 查詢關鍵字', '值日生', '群組裡打到這幾個字，bot 就回覆最新名單。回覆不計費，次數無上限'],
  ['LINE 關鍵字比對方式', '完全相同', '完全相同 / 開頭 / 包含。群組裡「值日生」本來就是日常用語，設成「包含」會一直跳出來'],
  ['LINE 回覆顯示堂數', 2, '群組查詢時回覆接下來幾堂。停課的那幾週照樣會列出來，但不佔名額'],
  ['排班網頁網址', '', '（選填）填了的話，LINE 回覆最後會附上這個連結。⚠ 要填「正式」那個部署對應的網頁，不要填測試的']
];

// ---------------------------------------------------------------------------
// 日期工具（一律用 yyyy-MM-dd 純字串運算，避開時區位移）
// ---------------------------------------------------------------------------

function normalizeDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  if (!s) return '';
  // 容許 2026/9/1 這種輸入
  var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) return m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);
  return s;
}

function pad2_(n) { return ('0' + n).slice(-2); }

function toUTC_(dateStr) {
  var p = dateStr.split('-');
  return new Date(Date.UTC(Number(p[0]), Number(p[1]) - 1, Number(p[2])));
}

function fromUTC_(d) {
  return d.getUTCFullYear() + '-' + pad2_(d.getUTCMonth() + 1) + '-' + pad2_(d.getUTCDate());
}

function addDays_(dateStr, n) {
  var d = toUTC_(dateStr);
  d.setUTCDate(d.getUTCDate() + n);
  return fromUTC_(d);
}

function weekdayCh_(dateStr) { return WEEKDAY_CH[toUTC_(dateStr).getUTCDay()]; }

function todayStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function nowStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }

function isTrue_(v) {
  if (v === true) return true;
  var s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === 'Y' || s === 'YES' || s === '1' || s === '是' || s === '✓';
}

// ---------------------------------------------------------------------------
// 初始化
// ---------------------------------------------------------------------------

/**
 * 只需要執行一次。已存在的工作表不會被覆蓋。
 * 想整份重來：先手動刪掉那幾張工作表，再執行一次。
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);

  setupConfigSheet_(ss);
  setupRosterSheet_(ss);
  setupSpecialSheet_(ss);
  setupSubsSheet_(ss);
  setupLogSheet_(ss);

  if (!ss.getSheetByName(SHEET_SCHEDULE)) {
    var sh = ss.insertSheet(SHEET_SCHEDULE);
    sh.setFrozenRows(1);
  }
  var result = generateSchedule_();
  seedInitialAssignments_();

  Logger.log('setup 完成：' + result.summary);
  Logger.log('接著請執行 installTriggers() 開啟自動提醒。');
  return result.summary;
}

function setupConfigSheet_(ss) {
  if (ss.getSheetByName(SHEET_CONFIG)) return;
  var sh = ss.insertSheet(SHEET_CONFIG);
  var rows = [
    ['設定項目', '值', '說明'],
    ['班級名稱', '週二陶藝班', '顯示在網頁最上方'],
    ['學期名稱', '2026 秋季班', '★ 換學期改這裡。顯示在網頁最上方'],
    ['第一堂課日期', '2026-09-01', '★ 換學期改這裡。格式 yyyy-MM-dd，之後每 7 天一堂'],
    ['總堂數', 18, '★ 換學期改這裡。整學期共幾堂課'],
    ['上課時間', '14:00–17:00', '（選填）顯示在網頁與提醒信上'],
    ['每週值日生人數', 3, '每一堂課要幾位值日生'],
    ['每人最少次數', 2, '每位同學整學期至少要排幾次'],
    ['允許自行輸入姓名', 'TRUE', 'TRUE = 名單上沒有的人可以在網頁自己輸入名字加入'],
    ['新名字自動加入名單', 'TRUE', 'TRUE = 自行輸入的名字會自動寫進「學員名單」，之後大家都選得到'],
    ['開放報名', 'TRUE', 'FALSE = 鎖定，網頁變唯讀（排班確定後可鎖起來）'],
    ['開放代班', 'TRUE', 'FALSE = 關閉「徵求代班 / 我來代」功能'],
    ['代班通知全班', 'FALSE', 'TRUE = 有人徵求代班時寄信通知全班（注意 Gmail 每日 100 封上限）'],
    ['班長信箱', '', '（選填）每週進度信、代班通知的收件人'],
    ['提醒信主旨前綴', '【陶藝班值日生】', '（選填）']
  ].concat(LINE_CONFIG_ROWS);
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8eaed');
  sh.getRange(2, 2, rows.length - 1, 1).setNumberFormat('@'); // 純文字，避免日期被轉型
  sh.getRange(3, 1, 3, 3).setBackground('#fff4e5');           // 標出換學期要改的三列
  sh.getRange(3, 1).setNote('橘色這三列就是換學期唯一要改的地方：\n學期名稱、第一堂課日期、總堂數。\n改完後執行選單「值日生 → 依設定重建排班表」。');
  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 460);
  sh.setFrozenRows(1);
}

function setupRosterSheet_(ss) {
  if (ss.getSheetByName(SHEET_ROSTER)) return;
  var sh = ss.insertSheet(SHEET_ROSTER);
  sh.getRange(1, 1, 1, 3).setValues([['姓名', 'Email（選填）', '備註']])
    .setFontWeight('bold').setBackground('#e8eaed');
  sh.getRange(1, 2).setNote('有填 Email 的人，才會在值日當天早上收到提醒信。');

  var names = ['王小明', '陳美玲', '林志豪', '張淑芬', '黃俊傑', '吳雅婷', '劉建宏', '蔡佩珊'];
  sh.getRange(2, 1, names.length, 1).setValues(names.map(function (n) { return [n]; }));

  sh.setColumnWidth(1, 140); sh.setColumnWidth(2, 240); sh.setColumnWidth(3, 240);
  sh.setFrozenRows(1);
}

function setupSpecialSheet_(ss) {
  if (ss.getSheetByName(SHEET_SPECIAL)) return;
  var sh = ss.insertSheet(SHEET_SPECIAL);
  sh.getRange(1, 1, 1, 3).setValues([['日期', '說明', '需排值日生']])
    .setFontWeight('bold').setBackground('#e8eaed');
  sh.getRange(1, 1).setNote(
    '這裡列出「不照常態排值日生」的日期。\n' +
    '重建排班表時，這些設定會覆蓋到對應的那一堂課上。\n' +
    '日期必須落在學期範圍內，否則會被忽略。');

  var rows = [
    ['2026-10-27', '社大公民週', false],
    ['2026-12-29', '吃好料的時光', false]
  ];
  sh.getRange(2, 1, rows.length, 3).setValues(rows);
  sh.getRange(2, 1, 200, 1).setNumberFormat('@');
  sh.getRange(2, 3, 200, 1).insertCheckboxes();
  sh.getRange(2, 3, rows.length, 1).setValues(rows.map(function (r) { return [r[2]]; }));

  sh.setColumnWidth(1, 120); sh.setColumnWidth(2, 240); sh.setColumnWidth(3, 110);
  sh.setFrozenRows(1);
}

function setupSubsSheet_(ss) {
  if (ss.getSheetByName(SHEET_SUBS)) return;
  var sh = ss.insertSheet(SHEET_SUBS);
  sh.getRange(1, 1, 1, 7).setValues([['建立時間', '日期', '週次', '原值日生', '狀態', '代班人', '完成時間']])
    .setFontWeight('bold').setBackground('#e8eaed');
  sh.getRange(1, 5).setNote('徵求中 / 已代班 / 已失效');
  sh.setColumnWidth(1, 150); sh.setColumnWidth(7, 150);
  sh.setFrozenRows(1);
}

function setupLogSheet_(ss) {
  if (ss.getSheetByName(SHEET_LOG)) return;
  var sh = ss.insertSheet(SHEET_LOG);
  sh.getRange(1, 1, 1, 6).setValues([['時間', '動作', '姓名', '週次', '日期', '備註']])
    .setFontWeight('bold').setBackground('#e8eaed');
  sh.setColumnWidth(1, 150); sh.setColumnWidth(6, 260);
  sh.setFrozenRows(1);
}

// ---------------------------------------------------------------------------
// 讀取各張工作表
// ---------------------------------------------------------------------------

function getConfig_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  if (!sh) throw new Error('找不到「' + SHEET_CONFIG + '」工作表，請先執行 setup()');
  var values = sh.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var k = String(values[i][0]).trim();
    if (k) map[k] = values[i][1];
  }

  var firstDate = normalizeDate_(map['第一堂課日期']);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(firstDate)) {
    throw new Error('「設定」的第一堂課日期格式不正確（需為 yyyy-MM-dd），目前是：' + firstDate);
  }
  var total = Number(map['總堂數']);
  if (!(total > 0 && total <= 100)) throw new Error('「設定」的總堂數必須是 1～100 的數字');

  var perWeek = Number(map['每週值日生人數']);
  if (!(perWeek > 0 && perWeek <= 10)) throw new Error('「設定」的每週值日生人數必須是 1～10');

  return {
    className:     String(map['班級名稱'] || '陶藝班'),
    semester:      String(map['學期名稱'] || ''),
    firstDate:     firstDate,
    totalLessons:  total,
    classTime:     String(map['上課時間'] || '').trim(),
    weekdayCh:     weekdayCh_(firstDate),
    perWeek:       perWeek,
    minPerPerson:  Number(map['每人最少次數']) || 0,
    allowFreeName: isTrue_(map['允許自行輸入姓名']),
    autoAddRoster: isTrue_(map['新名字自動加入名單']),
    open:          isTrue_(map['開放報名']),
    subsOpen:      isTrue_(map['開放代班']),
    notifyAllSubs: isTrue_(map['代班通知全班']),
    leaderEmail:   String(map['班長信箱'] || '').trim(),
    subjectPrefix: String(map['提醒信主旨前綴'] || '【值日生】'),

    // LINE。欄位不存在時一律視為停用，所以舊的試算表不會因此壞掉。
    lineGroupId:   String(map['LINE 群組 ID'] || '').trim(),
    linePushOn:    isTrue_(map['啟用 LINE 推播']),
    linePushEvery: Math.max(1, Number(map['LINE 推播間隔週數']) || 2),
    lineKeyword:   String(map['LINE 查詢關鍵字'] || '值日生').trim(),
    lineMatch:     String(map['LINE 關鍵字比對方式'] || '完全相同').trim(),
    lineReplyAhead: Math.max(1, Number(map['LINE 回覆顯示堂數']) || 2),
    siteUrl:       String(map['排班網頁網址'] || '').trim()
  };
}

function getRoster_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  if (!sh) throw new Error('找不到「' + SHEET_ROSTER + '」工作表，請先執行 setup()');
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, 2).getValues();
  var out = [], seen = {};
  for (var i = 0; i < values.length; i++) {
    var name = String(values[i][0]).trim();
    if (!name || seen[name]) continue;
    seen[name] = true;
    out.push({ name: name, email: String(values[i][1] || '').trim() });
  }
  return out;
}

function addToRoster_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  if (!sh) return false;
  var exists = getRoster_().some(function (p) { return p.name === name; });
  if (exists) return false;
  sh.appendRow([name, '', '由網頁自行加入 ' + nowStr_()]);
  return true;
}

function getSpecial_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SPECIAL);
  var map = {};
  if (!sh) return map;
  var last = sh.getLastRow();
  if (last < 2) return map;
  var values = sh.getRange(2, 1, last - 1, 3).getValues();
  for (var i = 0; i < values.length; i++) {
    var date = normalizeDate_(values[i][0]);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) continue;
    map[date] = { label: String(values[i][1] || '').trim(), needsDuty: isTrue_(values[i][2]) };
  }
  return map;
}

function getScheduleSheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SCHEDULE);
  if (!sh) throw new Error('找不到「' + SHEET_SCHEDULE + '」工作表，請先執行 setup()');
  return sh;
}

/** 讀排班表。不指定 perWeek 時，依實際欄數推算（重建排班表時會用到）。 */
function readSchedule_(perWeek) {
  var sh = getScheduleSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var slotCount = perWeek || Math.max(0, sh.getLastColumn() - (COL_SLOT1 - 1));
  var width = COL_SLOT1 - 1 + slotCount;
  var values = sh.getRange(2, 1, last - 1, width).getValues();
  var weeks = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var date = normalizeDate_(row[COL_DATE - 1]);
    if (!date) continue;
    var slots = [];
    for (var s = 0; s < slotCount; s++) slots.push(String(row[COL_SLOT1 - 1 + s] || '').trim());
    weeks.push({
      row: i + 2,
      no: Number(row[COL_WEEK_NO - 1]) || (i + 1),
      date: date,
      label: String(row[COL_LABEL - 1] || '').trim(),
      needsDuty: isTrue_(row[COL_NEEDS - 1]),
      slots: slots
    });
  }
  return weeks;
}

// ---------------------------------------------------------------------------
// 依「設定 + 特殊安排」重建排班表
// ---------------------------------------------------------------------------

/** 選單用的包裝：跑完跳出結果對話框 */
/**
 * 先試算一次「如果現在重建，會發生什麼事」，但不寫入任何東西。
 * 給重建前的確認視窗用，讓班長按下去之前就看得到後果。
 */
function previewRebuild_() {
  var cfg = getConfig_();
  var special = getSpecial_();

  var old = {};
  readSchedule_().forEach(function (w) {
    var names = w.slots.filter(function (n) { return n; });
    if (names.length) old[w.date] = names;
  });

  var dates = [], inRange = {};
  for (var i = 0; i < cfg.totalLessons; i++) {
    var d = addDays_(cfg.firstDate, 7 * i);
    dates.push(d);
    inRange[d] = true;
  }

  var kept = [], cleared = [];
  dates.forEach(function (d) {
    if (!old[d]) return;
    var sp = special[d] || { needsDuty: true };
    if (!sp.needsDuty) {
      cleared.push(d + '（' + old[d].join('、') + '）改為不排值日生');
    } else if (old[d].length > cfg.perWeek) {
      kept.push(d);
      cleared.push(d + '（' + old[d].slice(cfg.perWeek).join('、') + '）每週人數調少');
    } else {
      kept.push(d);
    }
  });
  Object.keys(old).forEach(function (d) {
    if (!inRange[d]) cleared.push(d + '（' + old[d].join('、') + '）不在新學期範圍內');
  });

  // 「特殊安排」的日期必須和某一堂課完全相同才對得上，打錯了會整列失效，
  // 所以先挑出來提醒，否則班長會以為設定好了、其實沒生效。
  var ignoredSpecial = [];
  Object.keys(special).forEach(function (d) {
    if (!inRange[d]) ignoredSpecial.push(d + '（' + (special[d].label || '無說明') + '）');
  });

  return {
    cfg: cfg, dates: dates, kept: kept, cleared: cleared,
    ignoredSpecial: ignoredSpecial, hadAny: Object.keys(old).length > 0
  };
}

/** 把目前的排班表另存一份，命名為「排班表 備份 MMDD-HHmm」 */
function backupScheduleSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName(SHEET_SCHEDULE);
  if (!src || src.getLastRow() < 2) return '';
  var name = '排班表 備份 ' + Utilities.formatDate(new Date(), TZ, 'MMdd-HHmm');
  src.copyTo(ss).setName(name);
  return name;
}

/**
 * 選單「依設定重建排班表」。
 * 三道保險：先試算 → 跳出確認視窗說明後果 → 確定後才動手，而且動手前自動備份。
 */
function rebuildSchedule() {
  var ui, p;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }

  try {
    p = previewRebuild_();
  } catch (err) {
    if (ui) ui.alert('設定有問題，還沒有動到排班表', err.message, ui.ButtonSet.OK);
    throw err;
  }

  var msg = [
    '將依「設定」重新產生排班表：',
    '',
    '　學期　' + p.cfg.semester,
    '　課程　' + p.dates[0] + ' 起，每週' + p.cfg.weekdayCh + '，共 ' + p.cfg.totalLessons + ' 堂',
    '　最後一堂　' + p.dates[p.dates.length - 1],
    '　每週值日生　' + p.cfg.perWeek + ' 人',
    ''
  ];

  if (!p.hadAny) {
    msg.push('目前排班表沒有任何人，不會有資料損失。');
  } else {
    msg.push('日期相同的 ' + p.kept.length + ' 堂，排班會原樣保留。');
    if (p.cleared.length) {
      msg.push('');
      msg.push('⚠ 以下 ' + p.cleared.length + ' 筆排班會被清空：');
      p.cleared.slice(0, 12).forEach(function (c) { msg.push('　' + c); });
      if (p.cleared.length > 12) msg.push('　…以及其他 ' + (p.cleared.length - 12) + ' 筆');
    } else {
      msg.push('沒有任何排班會被清空。');
    }
    msg.push('');
    msg.push('動手前會自動把目前的排班表另存一份備份，按錯了也救得回來。');
  }
  if (p.ignoredSpecial.length) {
    msg.push('');
    msg.push('⚠「特殊安排」有 ' + p.ignoredSpecial.length + ' 筆日期對不到任何一堂課，會被忽略：');
    p.ignoredSpecial.slice(0, 8).forEach(function (c) { msg.push('　' + c); });
    msg.push('　（日期要和排班表上的某一堂完全相同才會生效）');
  }

  msg.push('');
  msg.push('確定要重建嗎？');

  if (ui) {
    var answer = ui.alert('重建排班表', msg.join('\n'), ui.ButtonSet.OK_CANCEL);
    if (answer !== ui.Button.OK) return;
  }

  var backup = backupScheduleSheet_();
  var res = generateSchedule_();
  var done = res.summary + (backup ? '\n\n舊的排班已備份為工作表「' + backup + '」。\n確認新的排班沒問題後，可以自行刪掉那張備份。' : '');

  if (ui) ui.alert('排班表已重建', done, ui.ButtonSet.OK);
  else Logger.log(done);
  return done;
}

/**
 * 從第一堂課日期起、每 7 天一堂，產生「總堂數」列。
 * 已經排好的值日生會依「日期」保留下來；日期若因設定改變而消失，該日的排班會被丟棄並列在回報中。
 */
function generateSchedule_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('系統忙碌中，請稍後再試');

  try {
    var cfg = getConfig_();
    var sh = getScheduleSheet_();

    // 1. 記下目前排好的人。日期與堂次都留著，等一下兩種方式都試著對回去。
    var oldEntries = [];
    readSchedule_().forEach(function (w) {
      var names = w.slots.filter(function (n) { return n; });
      if (names.length) oldEntries.push({ date: w.date, no: w.no, names: names, used: false });
    });
    var byDate = {};
    oldEntries.forEach(function (e) { byDate[e.date] = e; });

    // 2. 產生新的日期清單
    var special = getSpecial_();
    var plan = [], inRange2 = {};
    for (var i = 0; i < cfg.totalLessons; i++) {
      var date = addDays_(cfg.firstDate, 7 * i);
      inRange2[date] = true;
      plan.push({
        no: i + 1,
        date: date,
        sp: special[date] || { label: '', needsDuty: true },
        names: null
      });
    }

    // 2a. 第一輪：日期完全相同的，直接對回去（新增/刪減堂數、插入停課日都走這條）
    plan.forEach(function (p) {
      var e = byDate[p.date];
      if (e && !e.used) { e.used = true; p.names = e.names.slice(); }
    });

    // 註：這裡刻意「只認日期」。曾經試過日期對不上時改以「第幾堂」對應，
    //     但換學期時新舊日期完全不同，那樣會把上學期整份排班默默複製到新學期，
    //     班長不一定看得出來。規則單純可預期，比聰明重要。

    // 2c. 套用特殊安排、依每週人數裁切
    var rows = [], movedToOff = [], truncated = [];
    plan.forEach(function (p) {
      var names = p.names || [];
      if (!p.sp.needsDuty && names.length) {
        movedToOff.push(p.date + '（' + names.join('、') + '）');
        names = [];
      }
      if (names.length > cfg.perWeek) {
        truncated.push(p.date + '（' + names.slice(cfg.perWeek).join('、') + '）');
        names = names.slice(0, cfg.perWeek);
      }
      var row = [p.no, p.date, p.sp.label, p.sp.needsDuty];
      for (var s2 = 0; s2 < cfg.perWeek; s2++) row.push(names[s2] || '');
      rows.push(row);
    });

    var dropped = [];
    oldEntries.forEach(function (e) {
      if (!e.used) dropped.push(e.date + '（' + e.names.join('、') + '）');
    });

    var dates = plan.map(function (p) { return p.date; });

    // 3. 清空重寫
    sh.clear();
    var header = ['週次', '日期', '說明', '需排值日生'];
    for (var k = 1; k <= cfg.perWeek; k++) header.push('值日生' + k);
    sh.getRange(1, 1, 1, header.length).setValues([header])
      .setFontWeight('bold').setBackground('#e8eaed');
    sh.getRange(2, 1, rows.length, header.length).setValues(rows);
    sh.getRange(2, COL_DATE, rows.length, 1).setNumberFormat('@');
    sh.getRange(2, COL_NEEDS, rows.length, 1).insertCheckboxes();
    sh.getRange(2, COL_NEEDS, rows.length, 1)
      .setValues(rows.map(function (r) { return [r[COL_NEEDS - 1]]; }));

    sh.setColumnWidth(COL_WEEK_NO, 60);
    sh.setColumnWidth(COL_DATE, 110);
    sh.setColumnWidth(COL_LABEL, 170);
    sh.setColumnWidth(COL_NEEDS, 100);
    sh.setFrozenRows(1);
    if (sh.getMaxColumns() > header.length) {
      sh.deleteColumns(header.length + 1, sh.getMaxColumns() - header.length);
    }

    // 4. 讓已失效的代班需求歸零
    voidStaleSubRequests_(cfg);

    var lines = [
      '學期：' + cfg.semester,
      '課程：' + cfg.firstDate + ' 起，每週' + cfg.weekdayCh + '，共 ' + cfg.totalLessons + ' 堂',
      '最後一堂：' + dates[dates.length - 1],
      '每週值日生：' + cfg.perWeek + ' 人',
      '不排值日生的日期：' + (dates.filter(function (d) { return special[d] && !special[d].needsDuty; }).join('、') || '無')
    ];
    var ignoredSpecial = [];
    Object.keys(special).forEach(function (d) {
      if (!inRange2[d]) ignoredSpecial.push(d + '（' + (special[d].label || '無說明') + '）');
    });
    if (ignoredSpecial.length) lines.push('⚠「特殊安排」有 ' + ignoredSpecial.length + ' 筆日期對不到課，已忽略：' + ignoredSpecial.join('、'));
    if (movedToOff.length) lines.push('⚠ 因改為「不排值日生」而清空的排班：' + movedToOff.join('、'));
    if (truncated.length)  lines.push('⚠ 因每週人數調少而移除的排班：' + truncated.join('、'));
    if (dropped.length)    lines.push('⚠ 已不在學期內、無法對應而遺失的排班：' + dropped.join('、'));
    if (!movedToOff.length && !dropped.length && !truncated.length) lines.push('原有排班全部保留。');

    appendLog_('重建排班表', '', '', '', lines.join(' / '));
    SpreadsheetApp.flush();
    return { summary: lines.join('\n'), dropped: dropped, movedToOff: movedToOff,
             truncated: truncated, ignoredSpecial: ignoredSpecial };

  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// 一次性：把 LINE 接龍上「已經排好」的人搬進系統
// ---------------------------------------------------------------------------

/**
 * 這是 2026 秋季班從 LINE 群組抄過來的既有排班，只在第一次 setup() 時使用。
 * 下個學期用不到，可以整段連同 INITIAL_ASSIGNMENTS 一起刪掉。
 */
var INITIAL_ASSIGNMENTS = {
  '2026-09-01': ['王小明', '陳美玲', '林志豪'],
  '2026-09-08': ['王小明', '張淑芬', '黃俊傑'],
  '2026-09-15': ['王小明', '張淑芬', '黃俊傑'],
  '2026-09-22': ['張淑芬', '黃俊傑', '吳雅婷'],
  '2026-10-06': ['劉建宏', '吳雅婷', '蔡佩珊'],
  '2026-11-03': ['陳美玲', '劉建宏', '吳雅婷']
};

/** 只有在排班表「一個名字都沒有」時才寫入，避免蓋掉大家已經排好的資料。 */
function seedInitialAssignments_() {
  var cfg = getConfig_();
  var sh = getScheduleSheet_();
  var weeks = readSchedule_(cfg.perWeek);

  var hasAnyone = weeks.some(function (w) {
    return w.slots.some(function (n) { return n; });
  });
  if (hasAnyone) return 0;

  var written = 0;
  weeks.forEach(function (w) {
    if (!w.needsDuty) return;
    var names = INITIAL_ASSIGNMENTS[w.date];
    if (!names) return;
    for (var i = 0; i < cfg.perWeek && i < names.length; i++) {
      sh.getRange(w.row, COL_SLOT1 + i).setValue(names[i]);
      written++;
    }
  });
  if (written) {
    appendLog_('匯入既有排班', '', '', '', '共 ' + written + ' 個名額');
    SpreadsheetApp.flush();
  }
  return written;
}

// ---------------------------------------------------------------------------
// 代班需求
// ---------------------------------------------------------------------------

function getSubsSheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SUBS);
  if (!sh) throw new Error('找不到「' + SHEET_SUBS + '」工作表，請先執行 setup()');
  return sh;
}

/** 讀出所有「徵求中」的代班需求 */
function readPendingSubs_() {
  var sh = getSubsSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, 7).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][4]).trim() !== SUB_PENDING) continue;
    out.push({
      row: i + 2,
      createdAt: String(values[i][0]),
      date: normalizeDate_(values[i][1]),
      weekNo: Number(values[i][2]),
      original: String(values[i][3]).trim()
    });
  }
  return out;
}

/** 若原值日生已不在該週名單上（班長手改、或已被代掉），把徵求標記為失效 */
function voidStaleSubRequests_(cfg) {
  var weeks = {};
  readSchedule_(cfg.perWeek).forEach(function (w) { weeks[w.no] = w; });
  var sh = getSubsSheet_();
  var changed = 0;
  readPendingSubs_().forEach(function (s) {
    var w = weeks[s.weekNo];
    var stillThere = w && w.needsDuty && w.slots.indexOf(s.original) >= 0;
    if (!stillThere) {
      sh.getRange(s.row, 5).setValue(SUB_VOID);
      sh.getRange(s.row, 7).setValue(nowStr_());
      changed++;
    }
  });
  return changed;
}

// ---------------------------------------------------------------------------
// 對外 API
// ---------------------------------------------------------------------------

function doGet(e) {
  try { return json_(buildState_()); }
  catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: '請求格式錯誤' }); }

  // LINE webhook 跟網頁前端共用同一個網址，用 payload 的形狀分流。
  // 不論處理成功與否都要回 200 —— 回非 2xx 的話 LINE 會重試，連續失敗會把 webhook 關掉。
  if (body && body.events instanceof Array) {
    // Apps Script 的 doPost 拿不到 HTTP 標頭，所以**沒辦法驗證 X-Line-Signature**。
    // 退而求其次：webhook 網址後面掛 ?k=通行碼，從 e.parameter 讀得到。
    // 沒設通行碼時照常運作（方便初次設定），設了就一定要對。
    var secret = PropertiesService.getScriptProperties().getProperty(LINE_SECRET_KEY);
    if (secret && String((e && e.parameter && e.parameter.k) || '') !== String(secret)) {
      Logger.log('LINE webhook 通行碼不符，已略過');
      return json_({ ok: true });
    }
    try { handleLineWebhook_(body); }
    catch (err) { Logger.log('LINE webhook 失敗：' + (err && err.message || err)); }
    return json_({ ok: true });
  }

  try {
    switch (body.action) {
      case 'read':      return json_(buildState_());
      case 'signup':    return json_(mutate_('signup', body.name, Number(body.week)));
      case 'cancel':    return json_(mutate_('cancel', body.name, Number(body.week)));
      case 'requestSub':return json_(requestSub_(body.name, Number(body.week), body.reason));
      case 'cancelSub': return json_(cancelSub_(body.name, Number(body.week)));
      case 'claimSub':  return json_(claimSub_(body.name, Number(body.week), body.original));
      case 'addMember': return json_(addMember_(body.name));
      default:          return json_({ ok: false, error: '未知的動作：' + body.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildState_() {
  var cfg = getConfig_();
  var roster = getRoster_();
  var weeks = readSchedule_(cfg.perWeek);
  var pending = readPendingSubs_();
  var today = todayStr_();

  var counts = {};
  roster.forEach(function (p) { counts[p.name] = 0; });
  weeks.forEach(function (w) {
    if (!w.needsDuty) return;
    w.slots.forEach(function (n) { if (n) counts[n] = (counts[n] || 0) + 1; });
  });

  var subsByWeek = {};
  pending.forEach(function (s) {
    (subsByWeek[s.weekNo] = subsByWeek[s.weekNo] || []).push(s.original);
  });

  var currentWeekNo = null;
  for (var i = 0; i < weeks.length; i++) {
    if (weeks[i].date >= today) { currentWeekNo = weeks[i].no; break; }
  }

  return {
    ok: true,
    className: cfg.className,
    semester: cfg.semester,
    classTime: cfg.classTime,
    weekdayCh: cfg.weekdayCh,
    perWeek: cfg.perWeek,
    minPerPerson: cfg.minPerPerson,
    allowFreeName: cfg.allowFreeName,
    open: cfg.open,
    subsOpen: cfg.subsOpen,
    today: today,
    currentWeekNo: currentWeekNo,
    roster: roster.map(function (p) { return p.name; }),
    counts: counts,
    weeks: weeks.map(function (w) {
      return {
        no: w.no, date: w.date, label: w.label, needsDuty: w.needsDuty,
        slots: w.slots,
        subs: subsByWeek[w.no] || []
      };
    })
  };
}

function findWeek_(weeks, no) {
  for (var i = 0; i < weeks.length; i++) if (weeks[i].no === no) return weeks[i];
  return null;
}

/** 報名 / 取消 */
function mutate_(action, rawName, weekNo) {
  var name = String(rawName || '').trim();
  if (!name) return { ok: false, error: '請先選擇你的名字' };
  if (!weekNo) return { ok: false, error: '缺少週次' };

  var cfg = getConfig_();
  if (!cfg.open) return { ok: false, error: '班長已鎖定排班，如需異動請直接聯絡班長' };

  if (action === 'signup' && !cfg.allowFreeName) {
    var known = getRoster_().some(function (p) { return p.name === name; });
    if (!known) return { ok: false, error: '「' + name + '」不在學員名單中，請聯絡班長加入名單' };
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: '系統忙碌中，請稍後再試一次' };

  try {
    var sh = getScheduleSheet_();
    var target = findWeek_(readSchedule_(cfg.perWeek), weekNo);
    if (!target) return { ok: false, error: '找不到第 ' + weekNo + ' 週' };

    if (action === 'signup') {
      if (!target.needsDuty) return { ok: false, error: (target.label || '這一週') + '不用排值日生' };
      if (target.slots.indexOf(name) >= 0) return { ok: false, error: '你已經排在 ' + target.date + ' 了' };
      var free = target.slots.indexOf('');
      if (free < 0) {
        return { ok: false, error: target.date + ' 名額已滿（' + target.slots.join('、') + '），請改選其他日期' };
      }
      sh.getRange(target.row, COL_SLOT1 + free).setValue(name);
      if (cfg.autoAddRoster) addToRoster_(name);
      appendLog_('報名', name, target.no, target.date, '');

    } else {
      var idx = target.slots.indexOf(name);
      if (idx < 0) return { ok: false, error: '你原本就沒有排在 ' + target.date };
      sh.getRange(target.row, COL_SLOT1 + idx).setValue('');
      appendLog_('取消', name, target.no, target.date, '');
      voidStaleSubRequests_(cfg);
    }

    SpreadsheetApp.flush();
    var state = buildState_();
    state.message = (action === 'signup' ? '已排定 ' : '已取消 ') + mdText_(target.date);
    return state;

  } finally {
    lock.releaseLock();
  }
}

/** 值日生表示當天不能到，公開徵求代班 */
function requestSub_(rawName, weekNo, reason) {
  var name = String(rawName || '').trim();
  if (!name) return { ok: false, error: '請先選擇你的名字' };

  var cfg = getConfig_();
  if (!cfg.subsOpen) return { ok: false, error: '班長已關閉代班功能，請直接聯絡班長' };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: '系統忙碌中，請稍後再試一次' };

  try {
    var target = findWeek_(readSchedule_(cfg.perWeek), weekNo);
    if (!target) return { ok: false, error: '找不到第 ' + weekNo + ' 週' };
    if (target.slots.indexOf(name) < 0) return { ok: false, error: '你不是 ' + mdText_(target.date) + ' 的值日生' };

    var already = readPendingSubs_().some(function (s) {
      return s.weekNo === weekNo && s.original === name;
    });
    if (already) return { ok: false, error: '你已經在徵求 ' + mdText_(target.date) + ' 的代班了' };

    getSubsSheet_().appendRow([
      nowStr_(), target.date, target.no, name, SUB_PENDING, '', ''
    ]);
    appendLog_('徵求代班', name, target.no, target.date, String(reason || '').slice(0, 200));
    SpreadsheetApp.flush();

    notifySubRequest_(cfg, target, name, reason);

    var state = buildState_();
    state.message = '已公開徵求 ' + mdText_(target.date) + ' 的代班';
    return state;

  } finally {
    lock.releaseLock();
  }
}

/** 原值日生取消徵求（例如又可以到了） */
function cancelSub_(rawName, weekNo) {
  var name = String(rawName || '').trim();
  if (!name) return { ok: false, error: '請先選擇你的名字' };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: '系統忙碌中，請稍後再試一次' };

  try {
    var sh = getSubsSheet_();
    var hit = null;
    readPendingSubs_().forEach(function (s) {
      if (s.weekNo === weekNo && s.original === name) hit = s;
    });
    if (!hit) return { ok: false, error: '找不到你的代班徵求（可能已經有人接手了）' };

    sh.getRange(hit.row, 5).setValue(SUB_VOID);
    sh.getRange(hit.row, 7).setValue(nowStr_());
    appendLog_('取消徵求代班', name, hit.weekNo, hit.date, '');
    SpreadsheetApp.flush();

    var state = buildState_();
    state.message = '已取消 ' + mdText_(hit.date) + ' 的代班徵求';
    return state;

  } finally {
    lock.releaseLock();
  }
}

/** 其他同學認領代班：把排班表上的原值日生換成代班人 */
function claimSub_(rawName, weekNo, rawOriginal) {
  var name = String(rawName || '').trim();
  var original = String(rawOriginal || '').trim();
  if (!name) return { ok: false, error: '請先選擇你的名字' };
  if (!original) return { ok: false, error: '缺少原值日生' };
  if (name === original) return { ok: false, error: '不能代自己的班，若可以到請按「取消徵求」' };

  var cfg = getConfig_();
  if (!cfg.subsOpen) return { ok: false, error: '班長已關閉代班功能，請直接聯絡班長' };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: '系統忙碌中，請稍後再試一次' };

  try {
    var sh = getScheduleSheet_();
    var target = findWeek_(readSchedule_(cfg.perWeek), weekNo);
    if (!target) return { ok: false, error: '找不到第 ' + weekNo + ' 週' };
    if (target.slots.indexOf(name) >= 0) {
      return { ok: false, error: '你本來就排在 ' + mdText_(target.date) + '，不需要代班' };
    }

    var hit = null;
    readPendingSubs_().forEach(function (s) {
      if (s.weekNo === weekNo && s.original === original) hit = s;
    });
    if (!hit) return { ok: false, error: '這筆代班已經被別人接走了，請重新整理看看' };

    var idx = target.slots.indexOf(original);
    if (idx < 0) {
      // 排班表已被改動，徵求作廢
      getSubsSheet_().getRange(hit.row, 5).setValue(SUB_VOID);
      getSubsSheet_().getRange(hit.row, 7).setValue(nowStr_());
      return { ok: false, error: original + ' 已經不在 ' + mdText_(target.date) + ' 的名單上了' };
    }

    sh.getRange(target.row, COL_SLOT1 + idx).setValue(name);
    var subSh = getSubsSheet_();
    subSh.getRange(hit.row, 5).setValue(SUB_DONE);
    subSh.getRange(hit.row, 6).setValue(name);
    subSh.getRange(hit.row, 7).setValue(nowStr_());

    if (cfg.autoAddRoster) addToRoster_(name);
    appendLog_('代班', name, target.no, target.date, '代替 ' + original);
    SpreadsheetApp.flush();

    notifySubClaimed_(cfg, target, original, name);

    var state = buildState_();
    state.message = '感謝！你代了 ' + original + ' 的 ' + mdText_(target.date);
    return state;

  } finally {
    lock.releaseLock();
  }
}

/** 同學在網頁上把自己加進名單 */
function addMember_(rawName) {
  var name = String(rawName || '').trim();
  if (!name) return { ok: false, error: '請輸入名字' };
  if (name.length > 20) return { ok: false, error: '名字太長了' };

  var cfg = getConfig_();
  if (!cfg.allowFreeName) return { ok: false, error: '班長已關閉自行加入，請聯絡班長' };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: '系統忙碌中，請稍後再試一次' };
  try {
    var added = addToRoster_(name);
    appendLog_(added ? '加入名單' : '加入名單（已存在）', name, '', '', '');
    SpreadsheetApp.flush();
    var state = buildState_();
    state.message = added ? '已把「' + name + '」加入名單' : '「' + name + '」本來就在名單裡';
    return state;
  } finally {
    lock.releaseLock();
  }
}

function mdText_(dateStr) {
  var p = dateStr.split('-');
  return Number(p[1]) + '/' + Number(p[2]);
}

function appendLog_(action, name, weekNo, date, note) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (!sh) return;
  sh.appendRow([nowStr_(), action, name, weekNo, date, note || '']);
}

// ---------------------------------------------------------------------------
// 通知信
// ---------------------------------------------------------------------------

function emailsByName_() {
  var map = {};
  getRoster_().forEach(function (p) { if (p.email) map[p.name] = p.email; });
  return map;
}

function dedupe_(arr) {
  var seen = {}, out = [];
  arr.forEach(function (v) { if (v && !seen[v]) { seen[v] = 1; out.push(v); } });
  return out;
}

function notifySubRequest_(cfg, week, original, reason) {
  var to = [];
  if (cfg.leaderEmail) to.push(cfg.leaderEmail);
  if (cfg.notifyAllSubs) {
    var map = emailsByName_();
    Object.keys(map).forEach(function (n) { if (n !== original) to.push(map[n]); });
  }
  to = dedupe_(to);
  if (!to.length) return;

  var body = [
    cfg.className + ' ' + cfg.semester,
    '',
    original + ' 無法出席 ' + week.date + '（' + weekdayCh_(week.date) + '）第 ' + week.no + ' 堂的值日生，正在徵求代班。',
    reason ? '原因：' + reason : '',
    '',
    '目前該日值日生：' + (week.slots.filter(function (n) { return n; }).join('、') || '無'),
    '',
    '可以幫忙的話，請到排班網頁按「我來代」。'
  ].filter(function (l) { return l !== ''; }).join('\n');

  try {
    MailApp.sendEmail({
      to: to.join(','),
      subject: cfg.subjectPrefix + week.date + ' ' + original + ' 徵求代班',
      body: body
    });
  } catch (e) {
    Logger.log('代班徵求通知寄送失敗：' + e.message);
  }
}

function notifySubClaimed_(cfg, week, original, claimer) {
  var map = emailsByName_();
  var to = dedupe_([map[original], map[claimer], cfg.leaderEmail]);
  if (!to.length) return;
  try {
    MailApp.sendEmail({
      to: to.join(','),
      subject: cfg.subjectPrefix + week.date + ' 代班成立：' + claimer + ' 代 ' + original,
      body: [
        cfg.className + ' ' + cfg.semester,
        '',
        week.date + '（' + weekdayCh_(week.date) + '）第 ' + week.no + ' 堂的值日生已更換：',
        original + ' → ' + claimer,
        '',
        '該日值日生：' + week.slots.map(function (n) {
          return n === original ? claimer : n;
        }).filter(function (n) { return n; }).join('、')
      ].join('\n')
    });
  } catch (e) {
    Logger.log('代班成立通知寄送失敗：' + e.message);
  }
}

// ---------------------------------------------------------------------------
// 自動提醒（時間觸發器）
// ---------------------------------------------------------------------------

/**
 * 只需執行一次。會依「設定」的上課星期，把提醒排在上課當天早上 07:00。
 * 重複執行會先清掉舊的觸發器，不會產生重複排程。
 */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });

  var cfg = getConfig_();
  var dow = toUTC_(cfg.firstDate).getUTCDay(); // 0=日
  var WD = [ScriptApp.WeekDay.SUNDAY, ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY,
            ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY,
            ScriptApp.WeekDay.SATURDAY];

  ScriptApp.newTrigger('sendTodayDutyReminder')
    .timeBased().onWeekDay(WD[dow]).atHour(7).inTimezone(TZ).create();

  var eve = WD[(dow + 6) % 7]; // 上課前一天晚上
  ScriptApp.newTrigger('sendWeeklyStatusToLeader')
    .timeBased().onWeekDay(eve).atHour(20).inTimezone(TZ).create();

  // LINE 推播掛每週，但 pushRosterToGroup() 自己會依「推播間隔週數」決定要不要真的推
  if (cfg.linePushOn) {
    ScriptApp.newTrigger('pushRosterToGroup')
      .timeBased().onWeekDay(eve).atHour(19).inTimezone(TZ).create();
  }

  var msg = '已建立提醒：每週' + WEEKDAY_CH[dow] + ' 07:00 通知當天值日生、' +
            '每週' + WEEKDAY_CH[(dow + 6) % 7] + ' 20:00 寄排班進度給班長。' +
            (cfg.linePushOn ? '另外每週' + WEEKDAY_CH[(dow + 6) % 7] + ' 19:00 檢查是否該推播 LINE（每 ' +
                              cfg.linePushEvery + ' 週一次）。' : '');
  Logger.log(msg);
  return msg;
}

function sendTodayDutyReminder() {
  var cfg = getConfig_();
  var weeks = readSchedule_(cfg.perWeek);
  var today = todayStr_();

  var week = null;
  for (var i = 0; i < weeks.length; i++) if (weeks[i].date === today) week = weeks[i];
  if (!week || !week.needsDuty) return;

  var members = week.slots.filter(function (n) { return n; });
  var pendingNames = readPendingSubs_()
    .filter(function (s) { return s.weekNo === week.no; })
    .map(function (s) { return s.original; });

  var map = emailsByName_();
  var to = [];
  members.forEach(function (n) { if (map[n]) to.push(map[n]); });
  if (cfg.leaderEmail) to.push(cfg.leaderEmail);
  to = dedupe_(to);
  if (!to.length) return;

  var body = [
    cfg.className + ' ' + cfg.semester,
    '',
    '今天 ' + today + '（' + weekdayCh_(today) + '）' + (cfg.classTime ? ' ' + cfg.classTime : '') +
      '，第 ' + week.no + '/' + weeks.length + ' 堂。',
    '',
    '值日生：' + (members.length ? members.join('、') : '（尚無人登記，請班長現場找人）'),
    pendingNames.length ? '⚠ 徵求代班中：' + pendingNames.join('、') + '（還沒有人接手）' : '',
    '',
    '臨時不能到的話，請到排班網頁按「我不能到，找人代班」，其他同學就看得到了。',
    '',
    '今天辛苦了，謝謝 🙏'
  ].filter(function (l) { return l !== ''; }).join('\n');

  MailApp.sendEmail({
    to: to.join(','),
    subject: cfg.subjectPrefix + today + ' 第 ' + week.no + ' 堂 值日生',
    body: body
  });
}

function sendWeeklyStatusToLeader() {
  var cfg = getConfig_();
  if (!cfg.leaderEmail) return;

  var state = buildState_();
  var today = state.today;

  var shortWeeks = state.weeks.filter(function (w) {
    return w.needsDuty && w.date >= today &&
           w.slots.filter(function (n) { return n; }).length < state.perWeek;
  });
  var shortPeople = state.roster.filter(function (n) {
    return (state.counts[n] || 0) < state.minPerPerson;
  });
  var pending = readPendingSubs_();

  var lines = [cfg.className + ' ' + cfg.semester, ''];

  if (pending.length) {
    lines.push('▍徵求代班中（' + pending.length + ' 筆）');
    pending.forEach(function (s) {
      lines.push('  ' + s.date + '　' + s.original + '　尚無人接手');
    });
    lines.push('');
  }
  if (shortWeeks.length) {
    lines.push('▍還缺人的日期（' + shortWeeks.length + ' 天）');
    shortWeeks.forEach(function (w) {
      var filled = w.slots.filter(function (n) { return n; });
      lines.push('  ' + w.date + '（第 ' + w.no + ' 堂）  ' + filled.length + '/' + state.perWeek +
                 '  ' + (filled.length ? filled.join('、') : '尚無人登記'));
    });
    lines.push('');
  }
  if (shortPeople.length) {
    lines.push('▍還沒排滿 ' + state.minPerPerson + ' 次的同學（' + shortPeople.length + ' 位）');
    shortPeople.forEach(function (n) {
      lines.push('  ' + n + '：目前 ' + (state.counts[n] || 0) + ' 次');
    });
    lines.push('');
  }

  if (!pending.length && !shortWeeks.length && !shortPeople.length) {
    MailApp.sendEmail({
      to: cfg.leaderEmail,
      subject: cfg.subjectPrefix + '排班已全部完成 ✅',
      body: cfg.className + ' ' + cfg.semester + '\n\n所有日期都排滿了，每位同學也都達到 ' +
            state.minPerPerson + ' 次，沒有待處理的代班。'
    });
    return;
  }

  MailApp.sendEmail({
    to: cfg.leaderEmail,
    subject: cfg.subjectPrefix + '排班進度（缺 ' + shortWeeks.length + ' 天 / 待補 ' +
             shortPeople.length + ' 人 / 代班 ' + pending.length + ' 筆）',
    body: lines.join('\n')
  });
}

// ---------------------------------------------------------------------------
// 班長工具
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// LINE Messaging API
//
// 計費的只有「主動推播」(push)，而且是按群組人數計算 —— 58 人的群組推一次就是
// 58 則，免費方案一個月只有 200 則。「被動回覆」(reply) 不計費也沒有上限。
//
// 所以這裡的策略是：
//   日常查詢 → 群組裡有人打關鍵字，bot 回覆（免費、無限次）
//   主動通知 → 每隔幾週推一次，推之前先問額度，不夠就改寄信給班長
//
// Channel access token 放 PropertiesService，不放試算表 —— 試算表會被複製、
// 匯出、截圖，任何有檢視權的人都看得到。
// ---------------------------------------------------------------------------

/** @return {string} 空字串代表尚未設定 */
function lineToken_() {
  var t = PropertiesService.getScriptProperties().getProperty(LINE_TOKEN_KEY);
  return t ? String(t).trim() : '';
}

/** 所有 LINE API 呼叫的共用出口。自己判斷狀態碼，才看得到錯誤訊息。 */
function lineFetch_(method, path, payload) {
  var token = lineToken_();
  if (!token) throw new Error('尚未設定 ' + LINE_TOKEN_KEY + '（專案設定 → 指令碼屬性）');

  var opt = { method: method, headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true };
  if (payload) { opt.contentType = 'application/json'; opt.payload = JSON.stringify(payload); }

  var res  = UrlFetchApp.fetch(LINE_API + path, opt);
  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code < 200 || code >= 300) throw new Error('LINE API ' + path + ' 回 ' + code + '：' + text);
  return text ? JSON.parse(text) : {};
}

/** 回覆。不計費、無上限，所以查詢功能想用幾次都可以。 */
function lineReply_(replyToken, text) {
  return lineFetch_('post', '/message/reply', {
    replyToken: replyToken, messages: [{ type: 'text', text: text }]
  });
}

/** 主動推播。計費，而且推到群組是按人數算。 */
function linePush_(to, text) {
  return lineFetch_('post', '/message/push', {
    to: to, messages: [{ type: 'text', text: text }]
  });
}

/** 本月還剩幾則可推。Infinity = 無上限或查不到，查不到時不要因此擋住推播。 */
function lineQuotaLeft_() {
  var q = lineFetch_('get', '/message/quota');
  if (!q || q.type !== 'limited') return Infinity;
  var used = lineFetch_('get', '/message/quota/consumption');
  return Number(q.value || 0) - Number((used && used.totalUsage) || 0);
}

/** 群組實際人數。加人、退群都會變，所以不要寫死。 */
function lineGroupSize_(groupId) {
  var r = lineFetch_('get', '/group/' + encodeURIComponent(groupId) + '/members/count');
  return Number((r && r.count) || 0);
}

/**
 * 給 LINE 看的簡短名單：只列接下來幾堂，不是整學期。
 * 整學期 18 行在手機上要滑很久，而且代班換人之後舊訊息就過期了。
 */
function buildLineDigest_(ahead) {
  var state = buildState_();
  var today = todayStr_();
  // 停課週照樣列出來（同學需要知道那天不用來），但不佔「接下來幾堂」的名額，
  // 否則碰到連假時，問一次只看得到一堂真正要排班的課。
  var want = Math.max(1, ahead);
  var upcoming = [], counted = 0;
  state.weeks.forEach(function (w) {
    if (w.date < today || counted >= want) return;
    upcoming.push(w);
    if (w.needsDuty) counted++;
  });

  // 連結放最後一行。LINE 會自動變成可點的，不需要額外處理。
  var cfg  = getConfig_();
  var foot = cfg.siteUrl ? ['', '想報名或找人代班：', cfg.siteUrl] : [];

  if (!upcoming.length) {
    return [state.className + ' ' + state.semester + ' 已經沒有未來的課程了。']
           .concat(foot).join('\n');
  }

  var lines = [state.className + ' ' + state.semester + ' 值日生'];
  upcoming.forEach(function (w) {
    var head = w.date.slice(5).replace('-', '/') + '（' + weekdayCh_(w.date) + '）' +
               '第 ' + w.no + '/' + state.weeks.length + ' 堂';
    if (!w.needsDuty) { lines.push(head + '　' + (w.label || '不用排值日生')); return; }

    var filled = w.slots.filter(function (n) { return n; });
    var line = head + '　' + (filled.length ? filled.join('、') : '（還沒有人）');
    var short = state.perWeek - filled.length;
    if (short > 0) line += '　還缺 ' + short + ' 位';
    if (w.subs.length) line += '\n　⚠ ' + w.subs.join('、') + ' 徵求代班中';
    lines.push(line);
  });
  return lines.concat(foot).join('\n');
}

/**
 * 這句話算不算在問值日生。
 *
 * 預設「完全相同」—— 群組裡「值日生」本來就是日常用語（「今天值日生很不OK」），
 * 只要「包含」就回的話，bot 會變成洗版機器。安靜比聰明重要。
 * 結尾的問號、波浪號這類語氣符號先去掉，免得打了「值日生？」反而沒反應。
 */
function lineKeywordHit_(text, keyword, mode) {
  if (!keyword) return true;   // 沒設關鍵字 = 來者不拒，自我診斷會對此提出警告
  var t = String(text || '').trim().replace(/[\s?？!！。.~～、，,]+$/, '');
  if (mode === '包含') return t.indexOf(keyword) !== -1;
  if (mode === '開頭') return t.indexOf(keyword) === 0;
  return t === keyword;        // 完全相同
}

/**
 * 處理 LINE 傳進來的事件。
 * 只回覆、不推播，所以這條路徑完全不花額度。
 */
function handleLineWebhook_(body) {
  var cfg = getConfig_();
  var events = (body && body.events) || [];

  events.forEach(function (ev) {
    if (!ev || ev.type !== 'message' || !ev.message || ev.message.type !== 'text') return;
    if (!ev.replyToken) return;

    // 還沒設定群組 ID 時，把它記進「操作紀錄」，班長才找得到要填什麼
    var gid = ev.source && ev.source.groupId;
    if (gid && !cfg.lineGroupId) appendLog_('LINE 群組 ID', '', '', '', gid);

    if (!lineKeywordHit_(ev.message.text, cfg.lineKeyword, cfg.lineMatch)) return;

    lineReply_(ev.replyToken, buildLineDigest_(cfg.lineReplyAhead));
  });
}

/**
 * 主動推播給群組。由時間觸發器每週呼叫，但自己判斷「這次該不該推」：
 *   1. 沒啟用、沒群組 ID、沒 token  → 不推
 *   2. 距離上次推播不滿設定的週數    → 不推
 *   3. 本月額度不夠                → 不推，改寄信給班長
 * 額度用完時 push 會回 429 而且不會有人知道，所以第 3 點不是加分項。
 */
function pushRosterToGroup() {
  var cfg = getConfig_();
  if (!cfg.linePushOn)    return '未啟用 LINE 推播';
  if (!cfg.lineGroupId)   return '「設定」裡還沒填 LINE 群組 ID';
  if (!lineToken_())      return '尚未設定 ' + LINE_TOKEN_KEY;

  var props = PropertiesService.getScriptProperties();
  var last  = props.getProperty(LINE_LAST_PUSH) || '';
  var today = todayStr_();
  if (last) {
    var days = Math.round((toUTC_(today) - toUTC_(last)) / 86400000);
    if (days < cfg.linePushEvery * 7 - 1) {
      return '距離上次推播（' + last + '）未滿 ' + cfg.linePushEvery + ' 週，略過';
    }
  }

  // 查不到就不要因此擋住推播 —— 寧可推出去，也不要安靜地什麼都沒發生
  var size = 0, left = Infinity;
  try { size = lineGroupSize_(cfg.lineGroupId); } catch (err) { size = 0; }
  try { left = lineQuotaLeft_(); }               catch (err) { left = Infinity; }

  var need = (size || LINE_SIZE_FALLBACK) + LINE_QUOTA_BUFFER;
  if (left < need) {
    notifyLeaderQuotaLow_(cfg, left, size, need);
    return '本月額度剩 ' + left + ' 則，不足 ' + need + ' 則，已改寄信通知班長';
  }

  linePush_(cfg.lineGroupId, buildLineDigest_(cfg.linePushEvery + 1));
  props.setProperty(LINE_LAST_PUSH, today);
  appendLog_('LINE 推播', '', '', today, '群組 ' + (size || '?') + ' 人，推播後剩約 ' + (left - (size || 0)) + ' 則');
  return '已推播給群組（' + (size || '?') + ' 人）';
}

/** 額度不夠時改走 email —— 這是整條路上唯一不依賴 LINE 額度的通道。 */
function notifyLeaderQuotaLow_(cfg, left, size, need) {
  if (!cfg.leaderEmail) return;
  MailApp.sendEmail({
    to: cfg.leaderEmail,
    subject: cfg.subjectPrefix + 'LINE 推播額度不足，本次未發送',
    body: [
      '這次的 LINE 群組推播沒有送出，因為本月額度不夠。',
      '',
      '　本月剩餘　' + left + ' 則',
      '　這次需要　' + need + ' 則（群組 ' + (size || '未知') + ' 人 + 安全邊際 ' + LINE_QUOTA_BUFFER + '）',
      '',
      '群組裡的同學仍然可以打「' + cfg.lineKeyword + '」查詢，那是回覆訊息，不計費。',
      '想手動貼一次的話，用試算表選單「值日生 → 產生 LINE 公告文字」。',
      '',
      '額度每月一號重置。'
    ].join('\n')
  });
}

/** 選單用：跑一次推播並把結果顯示出來（觸發器呼叫時沒有 UI，所以要容錯）。 */
function pushLineNow() {
  var msg;
  try { msg = pushRosterToGroup(); }
  catch (err) { msg = '失敗：' + (err && err.message || err); }
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }
  if (ui) ui.alert('LINE 推播', msg, ui.ButtonSet.OK);
  Logger.log(msg);
  return msg;
}

/**
 * 在既有的「設定」工作表補上 LINE 相關欄位。
 * setupConfigSheet_() 只在建表時跑一次，已經在用的試算表不會自動長出新欄位，
 * 所以要有這個。可以重複執行，已經存在的不會重複加。
 */
function ensureLineConfigRows() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
  if (!sh) throw new Error('找不到「' + SHEET_CONFIG + '」工作表，請先執行 setup()');

  var have = {};
  sh.getDataRange().getValues().forEach(function (r) { have[String(r[0]).trim()] = true; });

  var added = LINE_CONFIG_ROWS.filter(function (r) { return !have[r[0]]; });
  if (added.length) {
    sh.getRange(sh.getLastRow() + 1, 1, added.length, 3).setValues(added);
    sh.getRange(sh.getLastRow() - added.length + 1, 2, added.length, 1).setNumberFormat('@');
  }

  var msg = added.length
    ? '已在「設定」補上 ' + added.length + ' 個欄位：' + added.map(function (r) { return r[0]; }).join('、')
    : '「設定」裡的 LINE 欄位都已存在，沒有變更。';
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }
  if (ui) ui.alert('LINE 設定', msg + '\n\n提醒：Channel access token 不放在這裡，\n要填在「專案設定 → 指令碼屬性 → ' + LINE_TOKEN_KEY + '」。', ui.ButtonSet.OK);
  return msg;
}

/**
 * 一鍵自我診斷：bot 在群組裡沒反應時先跑這個。
 * 從 token 一路驗到 webhook，直接告訴你斷在哪一關 ——
 * 「執行項目」的記錄只看得到我們自己這邊的錯誤，
 * 看不到「LINE 根本沒把請求送過來」這種情況，那才是最常見的原因。
 */
function lineSelfTest() {
  var out  = [];
  var pass = function (t, d) { out.push('✅ ' + t + (d ? '\n　　' + d : '')); };
  var fail = function (t, d) { out.push('❌ ' + t + (d ? '\n　　' + d : '')); };
  var warn = function (t, d) { out.push('⚠️ ' + t + (d ? '\n　　' + d : '')); };
  var show = function () {
    var msg = out.join('\n');
    var ui = null;
    try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }
    if (ui) ui.alert('LINE 連線診斷', msg, ui.ButtonSet.OK);
    Logger.log(msg);
    return msg;
  };

  // 1. token 有沒有設。後面每一關都要用，沒有就不必往下驗了。
  var token = lineToken_();
  if (!token) {
    fail('Channel access token 尚未設定',
         '專案設定 → 指令碼屬性 → 新增 ' + LINE_TOKEN_KEY);
    return show();
  }
  pass('Channel access token 已設定', '長度 ' + token.length + ' 字');

  // 2. token 是不是真的能用
  try {
    var info = lineFetch_('get', '/info');
    pass('token 有效', '官方帳號：' + (info.displayName || '?') + '（' + (info.basicId || '') + '）');
    if (String(info.chatMode) === 'chat') {
      warn('目前是「聊天模式」',
           '回應模式要切到「Bot」，否則訊息會被聊天功能接走。\n　　OA Manager → 設定 → 回應設定');
    }
  } catch (err) {
    fail('token 無效或已被撤銷', String(err && err.message || err));
    return show();
  }

  // 3. webhook 網址登記了沒
  var endpoint = '';
  try {
    var ep = lineFetch_('get', '/channel/webhook/endpoint');
    endpoint = String(ep.endpoint || '');
    if (!endpoint)       fail('還沒登記 webhook 網址', 'Developers Console → Messaging API → Webhook URL');
    else if (!ep.active) fail('webhook 網址已登記，但沒啟用', '把下面的「Use webhook」打開');
    else                 pass('webhook 已啟用', endpoint);

    if (endpoint && endpoint.indexOf('k=') === -1) {
      warn('webhook 網址沒有帶通行碼',
           'Apps Script 讀不到 HTTP 標頭、驗不了簽章，所以改用網址上的 ?k=。\n　　沒帶的話，任何人知道網址就能觸發。');
    }
  } catch (err) {
    fail('查不到 webhook 設定', String(err && err.message || err));
  }

  // 4. 請 LINE 實際打一次，確認它真的碰得到我們
  if (endpoint) {
    try {
      var t = lineFetch_('post', '/channel/webhook/test', {});
      if (t.success) pass('LINE 送得到我們', 'HTTP ' + t.statusCode);
      else fail('LINE 送不到我們',
                'HTTP ' + t.statusCode + ' ' + (t.reason || '') + ' ' + (t.detail || '') +
                '\n　　多半是部署沒更新，或存取權不是「所有人」');
    } catch (err) {
      fail('webhook 測試失敗', String(err && err.message || err));
    }
    warn('這一關只測「打得到」',
         '通行碼不對時我們一樣回 200（回非 2xx 會讓 LINE 以為壞掉而關掉 webhook），\n　　所以這關過了，不代表 ?k= 填對了。');
  }

  // 5. 試算表這邊的設定
  var cfg = getConfig_();
  if (cfg.lineKeyword) pass('查詢關鍵字', '「' + cfg.lineKeyword + '」（比對方式：' + cfg.lineMatch + '）');
  else warn('沒設查詢關鍵字', '群組裡任何一句話都會觸發回覆');

  if (!cfg.lineGroupId) {
    warn('還沒填 LINE 群組 ID',
         '這只影響「主動推播」，不影響關鍵字回覆。\n　　在群組打一次關鍵字，ID 就會記進「操作紀錄」。');
  } else {
    try { pass('群組 ID 可用', '目前 ' + lineGroupSize_(cfg.lineGroupId) + ' 人'); }
    catch (err) { fail('群組 ID 有問題', String(err && err.message || err)); }
  }

  // 6. 額度
  try {
    var left = lineQuotaLeft_();
    if (left === Infinity) pass('本月推播額度', '無上限');
    else out.push((left < LINE_SIZE_FALLBACK ? '⚠️ ' : '✅ ') +
                  '本月推播額度\n　　還剩 ' + left + ' 則（回覆不計費，只有主動推播要算）');
  } catch (err) {
    warn('查不到額度', String(err && err.message || err));
  }

  out.push('');
  out.push(cfg.linePushOn
    ? '主動推播：已啟用，每 ' + cfg.linePushEvery + ' 週一次'
    : '主動推播：未啟用（「設定」的「啟用 LINE 推播」改成 TRUE 才會推）');
  return show();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('值日生')
    .addItem('依設定重建排班表', 'rebuildSchedule')
    .addItem('清空所有值日生（重新排班）', 'clearAllDuties')
    .addSeparator()
    .addItem('產生 LINE 公告文字', 'showLineText')
    .addItem('立即推播到 LINE 群組', 'pushLineNow')
    .addItem('在「設定」補上 LINE 欄位', 'ensureLineConfigRows')
    .addItem('LINE 連線自我診斷', 'lineSelfTest')
    .addSeparator()
    .addItem('立即寄出今日提醒', 'sendTodayDutyReminder')
    .addItem('立即寄出進度給班長', 'sendWeeklyStatusToLeader')
    .addItem('重新建立自動提醒排程', 'installTriggers')
    .addToUi();
}

/**
 * 選單「清空所有值日生」。
 * 只清掉值日生欄位，日期、說明、需排值日生的勾選都保留。
 * 一樣是先問過再動手，動手前自動備份。
 */
function clearAllDuties() {
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }

  var cfg = getConfig_();
  var weeks = readSchedule_(cfg.perWeek);

  var filled = [];
  weeks.forEach(function (w) {
    var names = w.slots.filter(function (n) { return n; });
    if (names.length) filled.push({ date: w.date, no: w.no, names: names });
  });
  var totalNames = filled.reduce(function (a, f) { return a + f.names.length; }, 0);

  if (!totalNames) {
    var none = '排班表上目前沒有任何值日生，不需要清空。';
    if (ui) ui.alert('清空所有值日生', none, ui.ButtonSet.OK); else Logger.log(none);
    return none;
  }

  var msg = [
    '將清掉排班表上所有的值日生，讓全班重新排一次。',
    '',
    '　會清掉　' + totalNames + ' 個名額（' + filled.length + ' 堂）',
    '　保留　　日期、說明、需排值日生的勾選、學員名單',
    ''
  ];
  filled.slice(0, 10).forEach(function (f) {
    msg.push('　' + f.date + '　' + f.names.join('、'));
  });
  if (filled.length > 10) msg.push('　…以及其他 ' + (filled.length - 10) + ' 堂');
  msg.push('');
  msg.push('動手前會自動備份，按錯了也救得回來。');
  msg.push('');
  msg.push('確定要清空嗎？');

  if (ui) {
    var answer = ui.alert('清空所有值日生', msg.join('\n'), ui.ButtonSet.OK_CANCEL);
    if (answer !== ui.Button.OK) return;
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('系統忙碌中，請稍後再試');

  try {
    var backup = backupScheduleSheet_();
    var sh = getScheduleSheet_();
    weeks.forEach(function (w) {
      sh.getRange(w.row, COL_SLOT1, 1, cfg.perWeek)
        .setValues([w.slots.map(function () { return ''; })]);
    });
    SpreadsheetApp.flush();
    voidStaleSubRequests_(cfg);
    appendLog_('清空所有值日生', '', '', '', '共 ' + totalNames + ' 個名額，備份於「' + backup + '」');
    SpreadsheetApp.flush();

    var done = '已清空 ' + totalNames + ' 個名額，全班可以重新排了。' +
      (backup ? '\n\n舊的排班已備份為工作表「' + backup + '」。\n確認沒問題後可以自行刪掉。' : '');
    if (ui) ui.alert('清空完成', done, ui.ButtonSet.OK); else Logger.log(done);
    return done;
  } finally {
    lock.releaseLock();
  }
}

function buildLineText_() {
  var state = buildState_();
  var total = state.weeks.length;
  var head = state.semester + '值日生每週 ' + state.perWeek + ' 人' +
             (state.classTime ? '（' + state.classTime + '）' : '');
  var lines = [head, ''];
  state.weeks.forEach(function (w) {
    var no = pad2_(w.no);
    var md = w.date.slice(5).replace('-', '/');
    var line = no + '、' + md + '(' + weekdayCh_(w.date) + ') 第 ' + w.no + '/' + total + ' 堂';
    if (!w.needsDuty) {
      lines.push(line + (w.label ? '(' + w.label + ')' : '') + ' --> 不用排值日生');
    } else {
      var filled = w.slots.filter(function (n) { return n; });
      var mark = w.subs.length ? '  ⚠' + w.subs.join('、') + ' 徵求代班中' : '';
      lines.push(line + '~' + filled.join('、') + mark);
    }
  });
  return lines.join('\n');
}

function showLineText() {
  var text = buildLineText_();
  var html = HtmlService.createHtmlOutput(
    '<p style="font:13px/1.5 -apple-system,sans-serif">全選複製後貼到 LINE 群組：</p>' +
    '<textarea style="width:100%;height:420px;font:13px/1.6 monospace" onclick="this.select()">' +
    text.replace(/&/g, '&amp;').replace(/</g, '&lt;') +
    '</textarea>'
  ).setWidth(560).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '貼回 LINE 的公告文字');
}
