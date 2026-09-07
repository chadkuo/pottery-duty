/**
 * 後端邏輯測試。不需要網路，也不需要 Google 帳號。
 *   node test/run.js
 */
const path = require('path');
const fs = require('fs');
const { ss, mails, dialogs, ui } = require('./mock.js');
eval(fs.readFileSync(path.join(__dirname, '..', 'apps-script', 'Code.gs'), 'utf8'));

let pass = 0, fail = 0, group = '';
const section = n => { group = n; console.log('\n▍' + n); };
const t = (name, fn) => {
  try { fn(); console.log('  ✅', name); pass++; }
  catch (e) { console.log('  ❌', name, '→', e.message); fail++; }
};
const eq = (a, b, m) => {
  if (JSON.stringify(a) !== JSON.stringify(b)) throw new Error(`${m ? m + ' ' : ''}期望 ${JSON.stringify(b)}，實得 ${JSON.stringify(a)}`);
};
const ok = (c, m) => { if (!c) throw new Error(m || '不成立'); };
const errIncludes = (res, frag) => {
  ok(res.ok === false, '預期失敗但成功了');
  ok(String(res.error).includes(frag), `錯誤訊息應含「${frag}」，實得「${res.error}」`);
};

/** 依「設定項目」名稱改設定值，不必記列號 */
function setCfg(key, value) {
  const sh = ss.getSheetByName('設定');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  }
  throw new Error('找不到設定項目：' + key);
}
const sheetRows = name => ss.getSheetByName(name).getDataRange().getValues().slice(1);
const weekOf = (state, no) => state.weeks.find(w => w.no === no);

// ===========================================================================
section('setup()：依設定自動產生排班表');
setup();
t('建立 6 張工作表', () => eq(Object.keys(ss.sheets).sort(),
  ['代班需求', '學員名單', '排班表', '操作紀錄', 'особ'].slice(0, 4).concat(['特殊安排', '設定']).sort()));

let s = buildState_();
t('產生 18 堂', () => eq(s.weeks.length, 18));
t('第一堂 = 設定的第一堂課日期', () => eq(s.weeks[0].date, '2026-09-01'));
t('每 7 天一堂', () => eq([s.weeks[1].date, s.weeks[2].date], ['2026-09-08', '2026-09-15']));
t('最後一堂 = 12/29', () => eq(s.weeks[17].date, '2026-12-29'));
t('自動判斷上課星期為「二」', () => eq(s.weekdayCh, '二'));
t('帶出上課時間', () => eq(s.classTime, '14:00–17:00'));

section('特殊安排覆蓋到對應堂次');
t('10/27 標為公民週且不排班', () => {
  const w = s.weeks.find(w => w.date === '2026-10-27');
  eq([w.label, w.needsDuty], ['社大公民週', false]);
});
t('12/29 標為吃好料且不排班', () => {
  const w = s.weeks.find(w => w.date === '2026-12-29');
  eq([w.label, w.needsDuty], ['吃好料的時光', false]);
});
t('其餘 16 堂都要排班', () => eq(s.weeks.filter(w => w.needsDuty).length, 16));

section('一次性匯入 LINE 上既有排班');
t('第 1 堂三人到位', () => eq(s.weeks[0].slots, ['王小明', '陳美玲', '林志豪']));
t('王小明 3 次、蔡佩珊 1 次', () => eq([s.counts['王小明'], s.counts['蔡佩珊']], [3, 1]));
t('共匯入 18 個名額', () => eq(s.weeks.reduce((a, w) => a + w.slots.filter(Boolean).length, 0), 18));
t('重跑不會重複匯入', () => { seedInitialAssignments_(); eq(buildState_().weeks[0].slots, ['王小明', '陳美玲', '林志豪']); });

// ===========================================================================
section('報名 / 取消');
let r = mutate_('signup', '蔡佩珊', 7);
t('報名成功並寫入第一個空位', () => eq(weekOf(r, 7).slots, ['蔡佩珊', '', '']));
t('次數 +1', () => eq(r.counts['蔡佩珊'], 2));
t('同一週不能重複報名', () => errIncludes(mutate_('signup', '蔡佩珊', 7), '已經排在'));
t('停課週不能報名', () => errIncludes(mutate_('signup', '蔡佩珊', 9), '不用排值日生'));
t('不存在的週次擋下', () => errIncludes(mutate_('signup', '蔡佩珊', 99), '找不到'));
t('空白姓名擋下', () => errIncludes(mutate_('signup', '   ', 7), '選擇你的名字'));

mutate_('signup', '劉建宏', 7);
mutate_('signup', '吳雅婷', 7);
t('額滿後第 4 人被拒', () => errIncludes(mutate_('signup', '黃俊傑', 7), '名額已滿'));
t('取消後留下空位', () => eq(weekOf(mutate_('cancel', '劉建宏', 7), 7).slots, ['蔡佩珊', '', '吳雅婷']));
t('沒排的人取消被擋', () => errIncludes(mutate_('cancel', '黃俊傑', 7), '原本就沒有排'));
t('取消後別人可補位', () => eq(weekOf(mutate_('signup', '黃俊傑', 7), 7).slots, ['蔡佩珊', '黃俊傑', '吳雅婷']));

// ===========================================================================
section('成員可自行新增');
t('網頁自行加入名單', () => {
  const res = addMember_('新同學甲');
  ok(res.ok); ok(res.roster.includes('新同學甲'), '名單應含新同學甲');
});
t('重複加入不會產生兩筆', () => {
  addMember_('新同學甲');
  eq(sheetRows('學員名單').filter(row => row[0] === '新同學甲').length, 1);
});
t('報名時自動把新名字寫進名單', () => {
  const res = mutate_('signup', '新同學乙', 11);
  ok(res.ok); ok(res.roster.includes('新同學乙'), '名單應含新同學乙');
});
t('關閉自行輸入後，陌生名字被擋', () => {
  setCfg('允許自行輸入姓名', 'FALSE');
  errIncludes(mutate_('signup', '路人甲', 12), '不在學員名單');
  setCfg('允許自行輸入姓名', 'TRUE');
});

// ===========================================================================
section('代班：徵求');
t('不是值日生不能徵求代班', () => errIncludes(requestSub_('林志豪', 7), '不是'));
r = requestSub_('黃俊傑', 7, '臨時出差');
t('徵求成功', () => ok(r.ok, r.error));
t('該週出現徵求中名單', () => eq(weekOf(r, 7).subs, ['黃俊傑']));
t('排班表仍掛原值日生', () => eq(weekOf(r, 7).slots, ['蔡佩珊', '黃俊傑', '吳雅婷']));
t('不能重複徵求同一天', () => errIncludes(requestSub_('黃俊傑', 7), '已經在徵求'));
t('代班需求表留下徵求中紀錄', () => {
  const row = sheetRows('代班需求').find(x => x[3] === '黃俊傑');
  eq([row[1], row[2], row[4]], ['2026-10-13', 7, '徵求中']);
});

section('代班：認領');
t('不能代自己的班', () => errIncludes(claimSub_('黃俊傑', 7, '黃俊傑'), '不能代自己'));
t('已排在同一天的人不能代', () => errIncludes(claimSub_('吳雅婷', 7, '黃俊傑'), '本來就排在'));
t('原值日生名字對不上會擋下', () => errIncludes(claimSub_('林志豪', 7, '王小明'), '已經被別人接走'));

r = claimSub_('林志豪', 7, '黃俊傑');
t('認領成功', () => ok(r.ok, r.error));
t('排班表換成代班人', () => eq(weekOf(r, 7).slots, ['蔡佩珊', '林志豪', '吳雅婷']));
t('徵求中名單清空', () => eq(weekOf(r, 7).subs, []));
t('次數自動轉移：黃俊傑 -1、林志豪 +1', () => { eq(r.counts['黃俊傑'], 3); eq(r.counts['林志豪'], 2); });
t('代班需求表記錄誰代誰', () => {
  const row = sheetRows('代班需求').find(x => x[3] === '黃俊傑' && x[4] === '已代班');
  eq(row[5], '林志豪');
});
t('同一筆不能被認領兩次', () => errIncludes(claimSub_('王小明', 7, '黃俊傑'), '已經被別人接走'));

section('代班：取消徵求與失效');
requestSub_('吳雅婷', 7);
r = cancelSub_('吳雅婷', 7);
t('原值日生可自行取消徵求', () => { ok(r.ok, r.error); eq(weekOf(r, 7).subs, []); });
t('沒有徵求時取消會擋下', () => errIncludes(cancelSub_('吳雅婷', 7), '找不到你的代班徵求'));
t('原值日生退掉排班後，徵求自動失效', () => {
  requestSub_('蔡佩珊', 7);
  ok(weekOf(buildState_(), 7).subs.includes('蔡佩珊'), '應先處於徵求中');
  mutate_('cancel', '蔡佩珊', 7);
  eq(weekOf(buildState_(), 7).subs, []);
});
t('關閉代班功能後不能徵求', () => {
  setCfg('開放代班', 'FALSE');
  errIncludes(requestSub_('吳雅婷', 7), '關閉代班');
  setCfg('開放代班', 'TRUE');
});

// ===========================================================================
section('鎖定排班');
t('鎖定後不能報名', () => {
  setCfg('開放報名', 'FALSE');
  errIncludes(mutate_('signup', '黃俊傑', 12), '鎖定');
  setCfg('開放報名', 'TRUE');
});

// ===========================================================================
section('換學期：只改設定就重建排班表');
const before7 = weekOf(buildState_(), 7).slots.filter(Boolean);

t('加開兩堂課（18 → 20）', () => {
  setCfg('總堂數', 20);
  generateSchedule_();
  const st = buildState_();
  eq(st.weeks.length, 20);
  eq(st.weeks[19].date, '2027-01-12');
});
t('加開後原有排班完整保留', () => eq(weekOf(buildState_(), 7).slots.filter(Boolean), before7));

t('縮短堂數會回報被丟棄的排班', () => {
  mutate_('signup', '吳雅婷', 20);
  const res = generateSchedule_.length, out = (setCfg('總堂數', 18), generateSchedule_());
  eq(buildState_().weeks.length, 18);
  ok(out.dropped.some(d => d.includes('2027-01-12') && d.includes('吳雅婷')),
     '應回報 2027-01-12 吳雅婷，實得 ' + JSON.stringify(out.dropped));
});

t('改上課星期：整學期日期跟著移動，舊排班一律清空並列出', () => {
  setCfg('第一堂課日期', '2026-09-02');   // 改成週三
  const out = generateSchedule_();
  const st = buildState_();
  eq(st.weeks[0].date, '2026-09-02');
  eq(st.weekdayCh, '三');
  eq(st.weeks[17].date, '2026-12-30');
  eq(st.weeks[0].slots, ['', '', '']);
  ok(out.dropped.some(d => d.includes('2026-09-01') && d.includes('王小明')),
     '應列出被清空的 9/1，實得 ' + JSON.stringify(out.dropped.slice(0, 3)));
});
t('不會把舊排班默默搬到新日期', () => {
  const st = buildState_();
  eq(st.weeks.reduce((a, w) => a + w.slots.filter(Boolean).length, 0), 0);
});

t('改回原本的週二並重新填入排班', () => {
  setCfg('第一堂課日期', '2026-09-01');
  generateSchedule_();
  seedInitialAssignments_();
  eq(buildState_().weeks[0].slots, ['王小明', '陳美玲', '林志豪']);
});
t('把某天加進特殊安排 → 該堂變成不排班', () => {
  mutate_('signup', '蔡佩珊', 11);
  const sp = ss.getSheetByName('特殊安排');
  sp.appendRow(['2026-11-10', '國定假日', false]);
  const out = generateSchedule_();
  const w = buildState_().weeks.find(w => w.date === '2026-11-10');
  eq([w.label, w.needsDuty, w.slots.filter(Boolean)], ['國定假日', false, []]);
  ok(out.movedToOff.some(d => d.includes('2026-11-10') && d.includes('蔡佩珊')),
     '應回報被清空的排班，實得 ' + JSON.stringify(out.movedToOff));
});

t('調整每週人數 3 → 4，欄位跟著增加', () => {
  setCfg('每週值日生人數', 4);
  generateSchedule_();
  const st = buildState_();
  eq(st.perWeek, 4);
  eq(st.weeks[0].slots.length, 4);
  eq(st.weeks[0].slots, ['王小明', '陳美玲', '林志豪', '']);
});
t('調整每週人數 4 → 2，多餘的人會被截掉並保留前 2 位', () => {
  setCfg('每週值日生人數', 2);
  const out = generateSchedule_();
  ok(out.truncated.some(d => d.includes('林志豪')), '應回報被移除的林志豪，實得 ' + JSON.stringify(out.truncated));
  const st = buildState_();
  eq(st.weeks[0].slots, ['王小明', '陳美玲']);
  eq(st.weeks[0].slots.length, 2);
  setCfg('每週值日生人數', 3);
  generateSchedule_();
});

section('設定值防呆');
t('日期格式錯誤會明確報錯', () => {
  setCfg('第一堂課日期', '九月一日');
  let msg = '';
  try { getConfig_(); } catch (e) { msg = e.message; }
  ok(msg.includes('第一堂課日期格式不正確'), msg);
  setCfg('第一堂課日期', '2026-09-01');
});
t('接受 2026/9/1 這種寫法', () => {
  setCfg('第一堂課日期', '2026/9/1');
  eq(getConfig_().firstDate, '2026-09-01');
  setCfg('第一堂課日期', '2026-09-01');
});
t('總堂數為 0 會報錯', () => {
  setCfg('總堂數', 0);
  let msg = ''; try { getConfig_(); } catch (e) { msg = e.message; }
  ok(msg.includes('總堂數'), msg);
  setCfg('總堂數', 18);
  generateSchedule_();
});

// ===========================================================================
section('重建前的確認視窗與備份');
t('試算不會動到任何資料', () => {
  const before = JSON.stringify(buildState_().weeks);
  const p = previewRebuild_();
  eq(JSON.stringify(buildState_().weeks), before, '試算後排班表不該改變');
  ok(p.dates.length === 18 && p.kept.length > 0, JSON.stringify({ d: p.dates.length, k: p.kept.length }));
});
t('按「取消」時完全不動作', () => {
  const before = JSON.stringify(buildState_().weeks);
  const sheetsBefore = Object.keys(ss.sheets).length;
  ui.answer = 'CANCEL';
  setCfg('總堂數', 10);
  rebuildSchedule();
  ui.answer = 'OK';
  eq(JSON.stringify(buildState_().weeks), before, '取消後排班表不該改變');
  eq(Object.keys(ss.sheets).length, sheetsBefore, '取消後不該產生備份工作表');
  setCfg('總堂數', 18);
});
t('確認視窗會先說明將被清空的排班', () => {
  dialogs.length = 0;
  setCfg('總堂數', 12);
  ui.answer = 'CANCEL';
  rebuildSchedule();
  ui.answer = 'OK';
  const d = dialogs[0];
  eq(d.title, '重建排班表');
  ok(d.buttons === 'OK_CANCEL', '應該是可取消的視窗，實得 ' + d.buttons);
  ok(d.msg.includes('共 12 堂'), d.msg);
  ok(d.msg.includes('會被清空'), d.msg);
  setCfg('總堂數', 18);
});
t('設定填錯時會先擋下並說明，不會動到排班表', () => {
  const before = JSON.stringify(buildState_().weeks);
  dialogs.length = 0;
  setCfg('第一堂課日期', '下週二');
  let threw = false;
  try { rebuildSchedule(); } catch (e) { threw = true; }
  setCfg('第一堂課日期', '2026-09-01');          // 先修回設定，才讀得回狀態
  ok(threw, '應該要拋出錯誤');
  ok(dialogs.length === 1 && dialogs[0].title.includes('還沒有動到排班表'),
     JSON.stringify(dialogs.map(d => d.title)));
  eq(JSON.stringify(buildState_().weeks), before, '排班表不該被動到');
});
t('重建前自動備份，且備份內容是舊的排班', () => {
  const before = buildState_().weeks[0].slots.slice();
  setCfg('總堂數', 16);
  rebuildSchedule();
  const backup = Object.keys(ss.sheets).find(n => n.indexOf('排班表 備份') === 0);
  ok(backup, '找不到備份工作表，現有：' + Object.keys(ss.sheets).join('、'));
  const rows = ss.getSheetByName(backup).getDataRange().getValues();
  eq(rows[1].slice(4, 7), before, '備份的第 1 堂應與重建前相同');
  eq(rows.length - 1, 18, '備份應保有重建前的 18 堂');
  eq(buildState_().weeks.length, 16, '新的排班表應為 16 堂');
  setCfg('總堂數', 18);
  rebuildSchedule();
});

section('換學期：只改三格 + 按一次重建');
t('學期名稱、第一堂課日期、總堂數改完即完成換學期', () => {
  setCfg('學期名稱', '2027 春季班');
  setCfg('第一堂課日期', '2027-02-16');
  setCfg('總堂數', 18);
  // 特殊安排換成新學期的
  const sp = ss.getSheetByName('特殊安排');
  sp.clear();
  sp.getRange(1, 1, 1, 3).setValues([['日期', '說明', '需排值日生']]);
  sp.appendRow(['2027-04-06', '社大公民週', false]);
  rebuildSchedule();

  const st = buildState_();
  eq(st.semester, '2027 春季班');
  eq(st.weeks.length, 18);
  eq(st.weeks[0].date, '2027-02-16');
  eq(st.weekdayCh, '二');
  eq(st.weeks[17].date, '2027-06-15');
});
t('新學期的排班表是乾淨的，沒有上學期的人', () => {
  const st = buildState_();
  eq(st.weeks.reduce((a, w) => a + w.slots.filter(Boolean).length, 0), 0);
  Object.keys(st.counts).forEach(n => eq(st.counts[n], 0, n + ' 的次數應歸零'));
});
t('新學期的停課日生效', () => {
  const w = buildState_().weeks.find(w => w.date === '2027-04-06');
  eq([w.label, w.needsDuty], ['社大公民週', false]);
});
t('學員名單延續到新學期', () => {
  ok(buildState_().roster.includes('王小明'), '名單不該被清掉');
});
t('上學期的排班仍留在備份工作表裡', () => {
  const backups = Object.keys(ss.sheets).filter(n => n.indexOf('排班表 備份') === 0);
  ok(backups.length >= 1, '應該至少有一張備份');
  const rows = ss.getSheetByName(backups[backups.length - 1]).getDataRange().getValues();
  ok(rows.some(r => String(r[1]).indexOf('2026-') === 0), '備份裡應找得到 2026 年的日期');
});
t('新學期可以正常報名', () => {
  const res = mutate_('signup', '王小明', 1);
  ok(res.ok, res.error);
  eq(weekOf(res, 1).slots, ['王小明', '', '']);
});

// ===========================================================================
section('直接在「排班表」上手改（不重建，網頁立刻生效）');

const sched = () => ss.getSheetByName('排班表');
const rowOf = no => no + 1;                       // 第 no 堂在試算表的第幾列

t('手改「說明」欄，網頁立刻看得到', () => {
  sched().getRange(rowOf(3), 3).setValue('校外教學');
  eq(weekOf(buildState_(), 3).label, '校外教學');
});

t('手動取消勾選「需排值日生」，該堂就變成不用排', () => {
  mutate_('signup', '吳雅婷', 3);
  mutate_('signup', '黃俊傑', 3);
  eq(weekOf(buildState_(), 3).slots.filter(Boolean), ['吳雅婷', '黃俊傑']);

  sched().getRange(rowOf(3), 4).setValue(false);
  const w = weekOf(buildState_(), 3);
  eq(w.needsDuty, false);
  eq(buildState_().counts['吳雅婷'], 0, '不排班的堂次不該計入次數');
});

t('取消勾選不會刪掉原本排的人', () => {
  eq(sched().getRange(rowOf(3), 5, 1, 3).getValues()[0], ['吳雅婷', '黃俊傑', ''],
     '名字應該還留在儲存格裡');
});

t('重新勾選回來，人和次數都回來', () => {
  sched().getRange(rowOf(3), 4).setValue(true);
  const st = buildState_();
  eq(weekOf(st, 3).slots.filter(Boolean), ['吳雅婷', '黃俊傑']);
  eq(st.counts['吳雅婷'], 1);
});

t('不排班的堂次擋下報名', () => {
  sched().getRange(rowOf(4), 4).setValue(false);
  sched().getRange(rowOf(4), 3).setValue('停課');
  errIncludes(mutate_('signup', '林志豪', 4), '不用排值日生');
  sched().getRange(rowOf(4), 4).setValue(true);
  sched().getRange(rowOf(4), 3).setValue('');
});

t('手改日期（補課調到別天）立刻生效', () => {
  sched().getRange(rowOf(5), 2).setValue('2027-03-18');
  const w = weekOf(buildState_(), 5);
  eq(w.date, '2027-03-18');
  ok(mutate_('signup', '林志豪', 5).ok, '改過日期的堂次仍可報名');
});

t('手動加一列就多一堂課', () => {
  const last = buildState_().weeks.length;
  sched().appendRow([last + 1, '2027-06-22', '補課', true, '', '', '']);
  const st = buildState_();
  eq(st.weeks.length, last + 1);
  eq(st.weeks[last].date, '2027-06-22');
  ok(mutate_('signup', '蔡佩珊', last + 1).ok, '新加的一堂應該可以報名');
});

section('手改 vs. 特殊安排：重建後誰說了算');

t('手改「排班表」的說明，重建後會被蓋掉', () => {
  sched().getRange(rowOf(3), 3).setValue('校外教學');
  eq(weekOf(buildState_(), 3).label, '校外教學');
  rebuildSchedule();
  eq(weekOf(buildState_(), 3).label, '', '重建是依「特殊安排」重寫，手改的說明不會留下');
});

t('寫進「特殊安排」才會永久生效', () => {
  const date = buildState_().weeks[2].date;          // 第 3 堂的日期
  ss.getSheetByName('特殊安排').appendRow([date, '校外教學', false]);
  rebuildSchedule();
  const w = weekOf(buildState_(), 3);
  eq([w.label, w.needsDuty], ['校外教學', false]);
  rebuildSchedule();
  eq(weekOf(buildState_(), 3).label, '校外教學', '再重建幾次都還在');
});

// ===========================================================================
section('API 進入點');
const post = b => JSON.parse(doPost({ postData: { contents: JSON.stringify(b) } }));
t('read 可用', () => eq(post({ action: 'read' }).weeks.length, 18));
t('未知動作回錯誤', () => eq(post({ action: 'boom' }).ok, false));
t('壞掉的 JSON 不會讓服務掛掉', () => eq(JSON.parse(doPost({ postData: { contents: '{{{' } })).ok, false));
t('doGet 回完整狀態', () => eq(JSON.parse(doGet({})).ok, true));

section('提醒信');
t('installTriggers 依上課星期設定排程', () => ok(installTriggers().includes('每週二 07:00')));
t('提醒信與進度信不會拋錯', () => { sendTodayDutyReminder(); setCfg('班長信箱', 'leader@example.com'); sendWeeklyStatusToLeader(); });
t('進度信內容含缺人日期', () => {
  const m = mails[mails.length - 1];
  ok(m.to === 'leader@example.com', '收件人 ' + m.to);
  ok(/還缺人的日期/.test(m.body), m.body.slice(0, 120));
});

section('LINE 公告文字');
const line = buildLineText_();
t('每一堂都有一行', () => eq(line.split('\n').filter(l => /^\d\d、/.test(l)).length, 18));
t('標出不用排值日生的日期', () => ok(line.includes('(社大公民週) --> 不用排值日生'), line));
t('標出徵求代班中', () => {
  ok(requestSub_('王小明', 1).ok, '應能徵求代班');
  ok(buildLineText_().includes('⚠王小明 徵求代班中'), buildLineText_().split('\n')[2]);
  cancelSub_('王小明', 1);
});

section('稽核');
t('每筆異動都有紀錄', () => ok(sheetRows('操作紀錄').length > 20, '僅 ' + sheetRows('操作紀錄').length + ' 筆'));
t('代班紀錄含「代替」備註', () => ok(sheetRows('操作紀錄').some(r => r[1] === '代班' && String(r[5]).includes('代替 黃俊傑'))));

console.log(`\n=== ${pass} 通過 / ${fail} 失敗 ===`);
process.exit(fail ? 1 : 0);
