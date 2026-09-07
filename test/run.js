require('./mock.js');
const fs = require('fs');
// 直接把 Code.gs 當成腳本載入到全域
eval(fs.readFileSync(require('path').join(__dirname, '..', 'apps-script', 'Code.gs'), 'utf8'));

let pass = 0, fail = 0;
const t = (name, fn) => { try { fn(); console.log('  ✅', name); pass++; } catch (e) { console.log('  ❌', name, '→', e.message); fail++; } };
const eq = (a, b, m) => { if (JSON.stringify(a) !== JSON.stringify(b)) throw new Error(`${m || ''} 期望 ${JSON.stringify(b)}，實得 ${JSON.stringify(a)}`); };

console.log('\n▍setup()');
setup();
t('建立 4 張工作表', () => eq(Object.keys(require('./mock.js').ss.sheets).sort(), ['learn'].slice(0, 0).concat(['學員名單', '排班表', '操作紀錄', '設定'].sort())));

console.log('\n▍buildState_()');
let s = buildState_();
t('18 週', () => eq(s.weeks.length, 18));
t('每週 3 人', () => eq(s.perWeek, 3));
t('公民週不排班', () => eq(s.weeks.find(w => w.no === 9).needsDuty, false));
t('12/29 不排班', () => eq(s.weeks.find(w => w.no === 18).needsDuty, false));
t('第 1 週名單正確', () => eq(s.weeks[0].slots, ['錢芸惠', '俞光彥', '李築善']));
t('日期為 yyyy-MM-dd 純文字', () => eq(s.weeks[0].date, '2026-09-01'));
t('錢芸惠計 3 次', () => eq(s.counts['錢芸惠'], 3));
t('顏淑琦計 1 次', () => eq(s.counts['顏淑琦'], 1));
t('李築善計 1 次', () => eq(s.counts['李築善'], 1));

console.log('\n▍報名 signup');
let r = mutate_('signup', '顏淑琦', 7);
t('報名成功', () => eq(r.ok, true));
t('寫進第 7 週第 1 格', () => eq(r.weeks.find(w => w.no === 7).slots, ['顏淑琦', '', '']));
t('次數 +1', () => eq(r.counts['顏淑琦'], 2));

t('同一週不能重複報名', () => eq(mutate_('signup', '顏淑琦', 7).error.includes('已經排在'), true));
t('停課週不能報名', () => eq(mutate_('signup', '顏淑琦', 9).error.includes('不用排值日生'), true));
t('不存在的週次會擋下', () => eq(mutate_('signup', '顏淑琦', 99).error.includes('找不到'), true));
t('空白姓名會擋下', () => eq(mutate_('signup', '  ', 7).error.includes('選擇你的名字'), true));

console.log('\n▍額滿保護');
mutate_('signup', '石惠禎', 7);
mutate_('signup', '高淑梅', 7);
t('第 7 週已滿 3 人', () => eq(buildState_().weeks.find(w => w.no === 7).slots.filter(Boolean).length, 3));
t('第 4 人報名被拒', () => eq(mutate_('signup', '宓敦', 7).error.includes('名額已滿'), true));

console.log('\n▍取消 cancel');
r = mutate_('cancel', '石惠禎', 7);
t('取消成功且留下空位', () => eq(r.weeks.find(w => w.no === 7).slots, ['顏淑琦', '', '高淑梅']));
t('沒排的人取消會擋下', () => eq(mutate_('cancel', '宓敦', 7).error.includes('原本就沒有排'), true));
t('取消後別人可以補位', () => eq(mutate_('signup', '宓敦', 7).weeks.find(w => w.no === 7).slots, ['顏淑琦', '宓敦', '高淑梅']));

console.log('\n▍名單外的人');
require('./mock.js').ss.getSheetByName('設定').getRange(6, 2).setValue('FALSE'); // 允許自行輸入姓名 = FALSE
t('關閉自由輸入時，陌生名字被擋', () => eq(mutate_('signup', '路人甲', 8).error.includes('不在學員名單'), true));
require('./mock.js').ss.getSheetByName('設定').getRange(6, 2).setValue('TRUE');
t('開啟自由輸入時，陌生名字可報名', () => eq(mutate_('signup', '路人甲', 8).ok, true));

console.log('\n▍鎖定');
require('./mock.js').ss.getSheetByName('設定').getRange(7, 2).setValue('FALSE'); // 開放報名 = FALSE
t('鎖定後不能異動', () => eq(mutate_('signup', '宓敦', 11).error.includes('鎖定'), true));
require('./mock.js').ss.getSheetByName('設定').getRange(7, 2).setValue('TRUE');

console.log('\n▍操作紀錄');
t('每次異動都有記錄', () => { const n = require('./mock.js').ss.getSheetByName('操作紀錄').getLastRow(); if (n < 5) throw new Error('紀錄僅 ' + n + ' 列'); });

console.log('\n▍doPost 進入點');
const post = b => JSON.parse(doPost({ postData: { contents: JSON.stringify(b) } }));
t('read 可用', () => eq(post({ action: 'read' }).weeks.length, 18));
t('未知動作回錯誤', () => eq(post({ action: 'boom' }).ok, false));
t('壞掉的 JSON 不會炸', () => eq(JSON.parse(doPost({ postData: { contents: '{{{' } })).ok, false));

console.log('\n▍LINE 公告文字');
const line = buildLineText_();
t('包含公民週標註', () => eq(line.includes('09、10/27(二) 陶藝課 9/18(社大公民週) --> 不用排值日生'), true));
t('包含第一週名單', () => eq(line.includes('01、09/01(二) 陶藝課 1/18~錢芸惠、俞光彥、李築善'), true));

console.log('\n▍每週提醒信');
t('sendTodayDutyReminder 不會拋錯', () => sendTodayDutyReminder());
require('./mock.js').ss.getSheetByName('設定').getRange(8, 2).setValue('leader@example.com');
t('sendWeeklyStatusToLeader 不會拋錯', () => sendWeeklyStatusToLeader());

console.log(`\n=== ${pass} 通過 / ${fail} 失敗 ===`);
process.exit(fail ? 1 : 0);
