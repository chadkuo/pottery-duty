/**
 * 陶藝班「值日生」線上排班系統 — 後端
 * 執行環境：Google Apps Script（附掛在一份 Google 試算表上）
 *
 * 部署方式請見專案 README.md
 * 一句話版本：貼上這段程式 → 執行一次 setup() → 執行一次 installTriggers()
 *            → 部署為「網頁應用程式」(執行身分：我 / 存取權：所有人) → 把網址貼進 docs/config.js
 */

// ---------------------------------------------------------------------------
// 常數
// ---------------------------------------------------------------------------
var TZ = 'Asia/Taipei';

var SHEET_CONFIG   = '設定';
var SHEET_ROSTER   = '學員名單';
var SHEET_SCHEDULE = '排班表';
var SHEET_LOG      = '操作紀錄';

// 排班表欄位位置（1-based）
var COL_WEEK_NO = 1; // 週次
var COL_DATE    = 2; // 日期
var COL_LABEL   = 3; // 說明
var COL_NEEDS   = 4; // 需排值日生
var COL_SLOT1   = 5; // 值日生 1（後面連續 N 欄）

// ---------------------------------------------------------------------------
// 初始化：建立所有工作表並填入這學期的資料
// ---------------------------------------------------------------------------

/**
 * 只需要執行一次。重複執行不會刪掉已有的資料表（會直接略過）。
 * 若想整份重來，先手動把 4 張工作表刪掉再跑一次。
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);

  setupConfigSheet_(ss);
  setupRosterSheet_(ss);
  setupScheduleSheet_(ss);
  setupLogSheet_(ss);

  SpreadsheetApp.getUi === undefined
    ? Logger.log('setup 完成')
    : Logger.log('setup 完成，請接著執行 installTriggers()');
}

function setupConfigSheet_(ss) {
  if (ss.getSheetByName(SHEET_CONFIG)) return;
  var sh = ss.insertSheet(SHEET_CONFIG);
  var rows = [
    ['設定項目', '值', '說明'],
    ['班級名稱', '大同週二拉坏班', '顯示在網頁最上方'],
    ['學期名稱', '2026 秋季班', '顯示在網頁最上方'],
    ['每週值日生人數', 3, '每一堂課要幾位值日生'],
    ['每人最少次數', 2, '每位同學整學期至少要排幾次'],
    ['允許自行輸入姓名', 'TRUE', 'TRUE = 名單上沒有的人也能自己打名字報名；FALSE = 只能從名單挑'],
    ['開放報名', 'TRUE', 'FALSE = 鎖定，網頁變唯讀（排班完成後可鎖起來）'],
    ['班長信箱', '', '（選填）每週提醒信的副本收件人、系統異常通知'],
    ['提醒信主旨前綴', '【陶藝班值日生】', '（選填）']
  ];
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8eaed');
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 420);
  sh.setFrozenRows(1);
}

function setupRosterSheet_(ss) {
  if (ss.getSheetByName(SHEET_ROSTER)) return;
  var sh = ss.insertSheet(SHEET_ROSTER);
  sh.getRange(1, 1, 1, 3).setValues([['姓名', 'Email（選填）', '備註']])
    .setFontWeight('bold').setBackground('#e8eaed');

  // 目前已知的同學（從 LINE 接龍抓下來的）。班長請把其餘同學補上，一列一位。
  var names = ['錢芸惠', '俞光彥', '李築善', '志平', '宓敦', '高淑梅', '石惠禎', '顏淑琦'];
  var rows = names.map(function (n) { return [n, '', '']; });
  sh.getRange(2, 1, rows.length, 3).setValues(rows);

  sh.setColumnWidth(1, 140);
  sh.setColumnWidth(2, 240);
  sh.setColumnWidth(3, 240);
  sh.setFrozenRows(1);
  sh.getRange(11, 1).setNote('往下繼續加同學即可，系統會自動讀取。');
}

function setupScheduleSheet_(ss) {
  if (ss.getSheetByName(SHEET_SCHEDULE)) return;
  var sh = ss.insertSheet(SHEET_SCHEDULE);
  var perWeek = 3;

  var header = ['週次', '日期', '說明', '需排值日生'];
  for (var i = 1; i <= perWeek; i++) header.push('值日生' + i);
  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#e8eaed');

  // 2026 秋季班：09/01 起，每週二一堂，共 18 堂
  var seed = [
    [1,  '2026-09-01', '', true,  ['錢芸惠', '俞光彥', '李築善']],
    [2,  '2026-09-08', '', true,  ['錢芸惠', '志平', '宓敦']],
    [3,  '2026-09-15', '', true,  ['錢芸惠', '志平', '宓敦']],
    [4,  '2026-09-22', '', true,  ['志平', '宓敦', '高淑梅']],
    [5,  '2026-09-29', '', true,  []],
    [6,  '2026-10-06', '', true,  ['石惠禎', '高淑梅', '顏淑琦']],
    [7,  '2026-10-13', '', true,  []],
    [8,  '2026-10-20', '', true,  []],
    [9,  '2026-10-27', '社大公民週', false, []],
    [10, '2026-11-03', '', true,  ['俞光彥', '石惠禎', '高淑梅']],
    [11, '2026-11-10', '', true,  []],
    [12, '2026-11-17', '', true,  []],
    [13, '2026-11-24', '', true,  []],
    [14, '2026-12-01', '', true,  []],
    [15, '2026-12-08', '', true,  []],
    [16, '2026-12-15', '', true,  []],
    [17, '2026-12-22', '', true,  []],
    [18, '2026-12-29', '吃好料的時光', false, []]
  ];

  var values = seed.map(function (r) {
    var row = [r[0], r[1], r[2], r[3]];
    for (var i = 0; i < perWeek; i++) row.push(r[4][i] || '');
    return row;
  });
  sh.getRange(2, 1, values.length, header.length).setValues(values);
  sh.getRange(2, COL_DATE, values.length, 1).setNumberFormat('@'); // 日期以純文字保存，避免時區位移

  sh.setColumnWidth(COL_WEEK_NO, 60);
  sh.setColumnWidth(COL_DATE, 110);
  sh.setColumnWidth(COL_LABEL, 160);
  sh.setColumnWidth(COL_NEEDS, 100);
  sh.setFrozenRows(1);
  sh.getRange(2, COL_NEEDS, values.length, 1).insertCheckboxes();
}

function setupLogSheet_(ss) {
  if (ss.getSheetByName(SHEET_LOG)) return;
  var sh = ss.insertSheet(SHEET_LOG);
  sh.getRange(1, 1, 1, 5).setValues([['時間', '動作', '姓名', '週次', '日期']])
    .setFontWeight('bold').setBackground('#e8eaed');
  sh.setColumnWidth(1, 160);
  sh.setFrozenRows(1);
}

// ---------------------------------------------------------------------------
// 讀取設定 / 名單 / 排班表
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
  return {
    className:     String(map['班級名稱'] || '陶藝班'),
    semester:      String(map['學期名稱'] || ''),
    perWeek:       Number(map['每週值日生人數']) || 3,
    minPerPerson:  Number(map['每人最少次數']) || 0,
    allowFreeName: isTrue_(map['允許自行輸入姓名']),
    open:          isTrue_(map['開放報名']),
    leaderEmail:   String(map['班長信箱'] || '').trim(),
    subjectPrefix: String(map['提醒信主旨前綴'] || '【值日生】')
  };
}

function isTrue_(v) {
  if (v === true) return true;
  var s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === 'Y' || s === 'YES' || s === '1' || s === '是';
}

function getRoster_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);
  if (!sh) throw new Error('找不到「' + SHEET_ROSTER + '」工作表，請先執行 setup()');
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, 2).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var name = String(values[i][0]).trim();
    if (!name) continue;
    out.push({ name: name, email: String(values[i][1] || '').trim() });
  }
  return out;
}

function getScheduleSheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SCHEDULE);
  if (!sh) throw new Error('找不到「' + SHEET_SCHEDULE + '」工作表，請先執行 setup()');
  return sh;
}

/** 讀出整份排班表（含每列在試算表中的實際 row index，供寫入用） */
function readSchedule_(perWeek) {
  var sh = getScheduleSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var width = COL_SLOT1 - 1 + perWeek;
  var values = sh.getRange(2, 1, last - 1, width).getValues();
  var weeks = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[COL_WEEK_NO - 1] && !row[COL_DATE - 1]) continue;
    var slots = [];
    for (var s = 0; s < perWeek; s++) slots.push(String(row[COL_SLOT1 - 1 + s] || '').trim());
    weeks.push({
      row:       i + 2,
      no:        Number(row[COL_WEEK_NO - 1]) || (i + 1),
      date:      normalizeDate_(row[COL_DATE - 1]),
      label:     String(row[COL_LABEL - 1] || '').trim(),
      needsDuty: isTrue_(row[COL_NEEDS - 1]),
      slots:     slots
    });
  }
  return weeks;
}

/** 不論儲存格是文字還是日期，都轉成 yyyy-MM-dd */
function normalizeDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '').trim();
}

function todayStr_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

// ---------------------------------------------------------------------------
// 對外 API
// ---------------------------------------------------------------------------

function doGet(e) {
  try {
    return json_(buildState_());
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  var body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: '請求格式錯誤' });
  }

  try {
    switch (body.action) {
      case 'read':   return json_(buildState_());
      case 'signup': return json_(mutate_('signup', body.name, Number(body.week)));
      case 'cancel': return json_(mutate_('cancel', body.name, Number(body.week)));
      default:       return json_({ ok: false, error: '未知的動作：' + body.action });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 組出前端需要的完整狀態 */
function buildState_() {
  var cfg = getConfig_();
  var roster = getRoster_();
  var weeks = readSchedule_(cfg.perWeek);
  var today = todayStr_();

  var counts = {};
  roster.forEach(function (p) { counts[p.name] = 0; });
  weeks.forEach(function (w) {
    if (!w.needsDuty) return;
    w.slots.forEach(function (n) {
      if (!n) return;
      counts[n] = (counts[n] || 0) + 1;
    });
  });

  // 找出「今天或之後最近的一堂課」當作目前這一週
  var currentWeekNo = null;
  for (var i = 0; i < weeks.length; i++) {
    if (weeks[i].date >= today) { currentWeekNo = weeks[i].no; break; }
  }

  return {
    ok: true,
    className:     cfg.className,
    semester:      cfg.semester,
    perWeek:       cfg.perWeek,
    minPerPerson:  cfg.minPerPerson,
    allowFreeName: cfg.allowFreeName,
    open:          cfg.open,
    today:         today,
    currentWeekNo: currentWeekNo,
    roster:        roster.map(function (p) { return p.name; }),
    counts:        counts,
    weeks: weeks.map(function (w) {
      return {
        no: w.no, date: w.date, label: w.label,
        needsDuty: w.needsDuty,
        slots: w.slots
      };
    })
  };
}

/**
 * 報名 / 取消。整段用 LockService 包起來，避免兩個人同時搶同一個空位。
 */
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
  if (!lock.tryLock(15000)) {
    return { ok: false, error: '系統忙碌中，請稍後再試一次' };
  }

  try {
    var sh = getScheduleSheet_();
    var weeks = readSchedule_(cfg.perWeek);
    var target = null;
    for (var i = 0; i < weeks.length; i++) {
      if (weeks[i].no === weekNo) { target = weeks[i]; break; }
    }
    if (!target) return { ok: false, error: '找不到第 ' + weekNo + ' 週' };

    if (action === 'signup') {
      if (!target.needsDuty) {
        return { ok: false, error: (target.label || '這一週') + '不用排值日生' };
      }
      if (target.slots.indexOf(name) >= 0) {
        return { ok: false, error: '你已經排在 ' + target.date + ' 了' };
      }
      var free = target.slots.indexOf('');
      if (free < 0) {
        return { ok: false, error: target.date + ' 名額已滿（' + target.slots.join('、') + '），請改選其他日期' };
      }
      sh.getRange(target.row, COL_SLOT1 + free).setValue(name);
      appendLog_('報名', name, target);

    } else { // cancel
      var idx = target.slots.indexOf(name);
      if (idx < 0) return { ok: false, error: '你原本就沒有排在 ' + target.date };
      sh.getRange(target.row, COL_SLOT1 + idx).setValue('');
      appendLog_('取消', name, target);
    }

    SpreadsheetApp.flush();
    var state = buildState_();
    state.message = (action === 'signup' ? '已排定 ' : '已取消 ') + target.date;
    return state;

  } finally {
    lock.releaseLock();
  }
}

function appendLog_(action, name, week) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  if (!sh) return;
  sh.appendRow([
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    action, name, week.no, week.date
  ]);
}

// ---------------------------------------------------------------------------
// 自動提醒（時間觸發器）
// ---------------------------------------------------------------------------

/**
 * 只需要執行一次。會建立兩個排程：
 *   - 每週二 07:00：提醒今天的值日生
 *   - 每週日 20:00：提醒還沒排滿 / 還沒排足次數的人（寄給班長彙整）
 * 重複執行會先清掉本專案舊的觸發器，不會產生重複排程。
 */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('sendTodayDutyReminder')
    .timeBased().onWeekDay(ScriptApp.WeekDay.TUESDAY).atHour(7).inTimezone(TZ).create();

  ScriptApp.newTrigger('sendWeeklyStatusToLeader')
    .timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(20).inTimezone(TZ).create();

  Logger.log('已建立每週二 07:00 與每週日 20:00 的自動提醒');
}

/** 每週二早上：寄信給今天的值日生 */
function sendTodayDutyReminder() {
  var cfg = getConfig_();
  var weeks = readSchedule_(cfg.perWeek);
  var today = todayStr_();

  var week = null;
  for (var i = 0; i < weeks.length; i++) {
    if (weeks[i].date === today) { week = weeks[i]; break; }
  }
  if (!week) return;                 // 今天沒課
  if (!week.needsDuty) return;       // 公民週 / 最後一堂

  var members = week.slots.filter(function (n) { return n; });
  var emailByName = {};
  getRoster_().forEach(function (p) { if (p.email) emailByName[p.name] = p.email; });

  var listText = members.length
    ? members.join('、')
    : '（還沒有人登記，請班長現場找人）';

  var subject = cfg.subjectPrefix + today + ' 第 ' + week.no + ' 堂 值日生';
  var body = [
    cfg.className + ' ' + cfg.semester,
    '',
    '今天（' + today + '，第 ' + week.no + '/' + weeks.length + ' 堂）的值日生是：',
    listText,
    '',
    '麻煩今天辛苦了，謝謝 🙏'
  ].join('\n');

  var to = [];
  members.forEach(function (n) { if (emailByName[n]) to.push(emailByName[n]); });
  if (cfg.leaderEmail) to.push(cfg.leaderEmail);
  to = dedupe_(to);
  if (!to.length) return;            // 沒人留 email 就不寄

  MailApp.sendEmail({ to: to.join(','), subject: subject, body: body });
}

/** 每週日晚上：把「還缺人的日期」和「還沒排足次數的人」彙整寄給班長 */
function sendWeeklyStatusToLeader() {
  var cfg = getConfig_();
  if (!cfg.leaderEmail) return;

  var state = buildState_();
  var today = state.today;

  var shortWeeks = state.weeks.filter(function (w) {
    if (!w.needsDuty) return false;
    if (w.date < today) return false;
    return w.slots.filter(function (n) { return n; }).length < state.perWeek;
  });

  var shortPeople = state.roster.filter(function (n) {
    return (state.counts[n] || 0) < state.minPerPerson;
  });

  if (!shortWeeks.length && !shortPeople.length) {
    MailApp.sendEmail({
      to: cfg.leaderEmail,
      subject: cfg.subjectPrefix + '排班已全部完成 ✅',
      body: cfg.className + ' ' + cfg.semester + '\n\n所有日期都排滿了，每位同學也都達到 ' +
            state.minPerPerson + ' 次，不需要再催了。'
    });
    return;
  }

  var lines = [cfg.className + ' ' + cfg.semester, ''];

  if (shortWeeks.length) {
    lines.push('▍還缺人的日期（' + shortWeeks.length + ' 天）');
    shortWeeks.forEach(function (w) {
      var filled = w.slots.filter(function (n) { return n; });
      lines.push('  ' + w.date + '（第 ' + w.no + ' 堂）  ' +
                 filled.length + '/' + state.perWeek + '  ' +
                 (filled.length ? filled.join('、') : '尚無人登記'));
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

  MailApp.sendEmail({
    to: cfg.leaderEmail,
    subject: cfg.subjectPrefix + '本週排班進度（缺 ' + shortWeeks.length + ' 天 / 待補 ' + shortPeople.length + ' 人）',
    body: lines.join('\n')
  });
}

function dedupe_(arr) {
  var seen = {}, out = [];
  arr.forEach(function (v) { if (v && !seen[v]) { seen[v] = 1; out.push(v); } });
  return out;
}

// ---------------------------------------------------------------------------
// 班長工具：把排班結果輸出成可以直接貼回 LINE 的文字
// ---------------------------------------------------------------------------

/**
 * 在試算表選單「值日生」→「產生 LINE 公告文字」使用。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('值日生')
    .addItem('產生 LINE 公告文字', 'showLineText')
    .addItem('立即寄出今日提醒', 'sendTodayDutyReminder')
    .addItem('立即寄出進度給班長', 'sendWeeklyStatusToLeader')
    .addToUi();
}

function buildLineText_() {
  var state = buildState_();
  var total = state.weeks.length;
  var lines = [state.semester + '值日生每週 ' + state.perWeek + ' 人（公民週不用值日生）'];
  lines.push('');
  state.weeks.forEach(function (w) {
    var no = ('0' + w.no).slice(-2);
    var md = w.date.slice(5).replace('-', '/');
    var head = no + '、' + md + '(二) 陶藝課 ' + w.no + '/' + total;
    if (!w.needsDuty) {
      lines.push(head + (w.label ? '(' + w.label + ')' : '') + ' --> 不用排值日生');
    } else {
      var filled = w.slots.filter(function (n) { return n; });
      lines.push(head + '~' + filled.join('、'));
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
  ).setWidth(520).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '貼回 LINE 的公告文字');
}
