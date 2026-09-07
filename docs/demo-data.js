/**
 * 展示模式：把後端的行為在瀏覽器裡重現一份，用假資料跑。
 * 當 config.js 尚未填入後端網址，或網址後面加上 ?demo 時啟用。
 * 所有操作只存在記憶體裡，重新整理就還原，不會寫進任何試算表。
 */
window.DUTY_DEMO = function () {
  // ── 對應試算表「設定」 ────────────────────────────────
  var cfg = {
    className: '大同週二拉坏班',
    semester: '2026 秋季班（展示模式）',
    firstDate: '2026-09-01',
    totalLessons: 18,
    classTime: '14:00–17:00',
    perWeek: 3,
    minPerPerson: 2,
    allowFreeName: true,
    open: true,
    subsOpen: true
  };
  // ── 對應試算表「特殊安排」 ────────────────────────────
  var special = {
    '2026-10-27': { label: '社大公民週', needsDuty: false },
    '2026-12-29': { label: '吃好料的時光', needsDuty: false }
  };
  // ── 對應試算表「學員名單」 ────────────────────────────
  var roster = ['王小明', '陳美玲', '林志豪', '張淑芬', '黃俊傑', '吳雅婷', '劉建宏', '蔡佩珊'];
  // ── LINE 上已排好的 ──────────────────────────────────
  var initial = {
    '2026-09-01': ['王小明', '陳美玲', '林志豪'],
    '2026-09-08': ['王小明', '張淑芬', '黃俊傑'],
    '2026-09-15': ['王小明', '張淑芬', '黃俊傑'],
    '2026-09-22': ['張淑芬', '黃俊傑', '吳雅婷'],
    '2026-10-06': ['劉建宏', '吳雅婷', '蔡佩珊'],
    '2026-11-03': ['陳美玲', '劉建宏', '吳雅婷']
  };

  function pad2(n) { return ('0' + n).slice(-2); }
  function addDays(dateStr, n) {
    var p = dateStr.split('-');
    var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
    d.setUTCDate(d.getUTCDate() + n);
    return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
  }
  function weekdayCh(dateStr) {
    var p = dateStr.split('-');
    return '日一二三四五六'[new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay()];
  }
  function today() { return new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Taipei' }).format(new Date()); }
  function mdText(d) { var p = d.split('-'); return +p[1] + '/' + +p[2]; }

  // 依設定產生排班表，與後端 generateSchedule_ 同樣的規則
  var weeks = [];
  for (var i = 0; i < cfg.totalLessons; i++) {
    var date = addDays(cfg.firstDate, 7 * i);
    var sp = special[date] || { label: '', needsDuty: true };
    var names = sp.needsDuty ? (initial[date] || []) : [];
    var slots = [];
    for (var s = 0; s < cfg.perWeek; s++) slots.push(names[s] || '');
    weeks.push({ no: i + 1, date: date, label: sp.label, needsDuty: sp.needsDuty, slots: slots, subs: [] });
  }

  function byNo(no) {
    for (var i = 0; i < weeks.length; i++) if (weeks[i].no === Number(no)) return weeks[i];
    return null;
  }
  function fail(msg) { return Promise.resolve({ ok: false, error: msg }); }

  function state(message) {
    var counts = {};
    roster.forEach(function (n) { counts[n] = 0; });
    weeks.forEach(function (w) {
      if (!w.needsDuty) return;
      w.slots.forEach(function (n) { if (n) counts[n] = (counts[n] || 0) + 1; });
    });
    var t = today(), cur = null;
    for (var i = 0; i < weeks.length; i++) { if (weeks[i].date >= t) { cur = weeks[i].no; break; } }
    return {
      ok: true, demo: true, message: message,
      className: cfg.className, semester: cfg.semester,
      classTime: cfg.classTime, weekdayCh: weekdayCh(cfg.firstDate),
      perWeek: cfg.perWeek, minPerPerson: cfg.minPerPerson,
      allowFreeName: cfg.allowFreeName, open: cfg.open, subsOpen: cfg.subsOpen,
      today: t, currentWeekNo: cur,
      roster: roster.slice(), counts: counts,
      weeks: JSON.parse(JSON.stringify(weeks))
    };
  }

  function addToRoster(name) { if (roster.indexOf(name) < 0) roster.push(name); }

  function call(action, p) {
    p = p || {};
    var name = String(p.name || '').trim();
    if (action === 'read') return Promise.resolve(state());

    if (action === 'addMember') {
      if (!name) return fail('請輸入名字');
      var isNew = roster.indexOf(name) < 0;
      addToRoster(name);
      return Promise.resolve(state(isNew ? '已把「' + name + '」加入名單' : '「' + name + '」本來就在名單裡'));
    }

    var w = byNo(p.week);
    if (!w) return fail('找不到第 ' + p.week + ' 週');
    if (!name) return fail('請先選擇你的名字');

    if (action === 'signup') {
      if (!w.needsDuty) return fail((w.label || '這一週') + '不用排值日生');
      if (w.slots.indexOf(name) >= 0) return fail('你已經排在 ' + mdText(w.date) + ' 了');
      var free = w.slots.indexOf('');
      if (free < 0) return fail(mdText(w.date) + ' 名額已滿（' + w.slots.join('、') + '），請改選其他日期');
      w.slots[free] = name;
      addToRoster(name);
      return Promise.resolve(state('已排定 ' + mdText(w.date)));
    }

    if (action === 'cancel') {
      var idx = w.slots.indexOf(name);
      if (idx < 0) return fail('你原本就沒有排在 ' + mdText(w.date));
      w.slots[idx] = '';
      w.subs = w.subs.filter(function (n) { return n !== name; });
      return Promise.resolve(state('已取消 ' + mdText(w.date)));
    }

    if (action === 'requestSub') {
      if (!cfg.subsOpen) return fail('班長已關閉代班功能，請直接聯絡班長');
      if (w.slots.indexOf(name) < 0) return fail('你不是 ' + mdText(w.date) + ' 的值日生');
      if (w.subs.indexOf(name) >= 0) return fail('你已經在徵求 ' + mdText(w.date) + ' 的代班了');
      w.subs.push(name);
      return Promise.resolve(state('已公開徵求 ' + mdText(w.date) + ' 的代班'));
    }

    if (action === 'cancelSub') {
      if (w.subs.indexOf(name) < 0) return fail('找不到你的代班徵求（可能已經有人接手了）');
      w.subs = w.subs.filter(function (n) { return n !== name; });
      return Promise.resolve(state('已取消 ' + mdText(w.date) + ' 的代班徵求'));
    }

    if (action === 'claimSub') {
      var original = String(p.original || '').trim();
      if (!cfg.subsOpen) return fail('班長已關閉代班功能，請直接聯絡班長');
      if (!original) return fail('缺少原值日生');
      if (name === original) return fail('不能代自己的班，若可以到請按「取消徵求」');
      if (w.slots.indexOf(name) >= 0) return fail('你本來就排在 ' + mdText(w.date) + '，不需要代班');
      if (w.subs.indexOf(original) < 0) return fail('這筆代班已經被別人接走了，請重新整理看看');
      var at = w.slots.indexOf(original);
      if (at < 0) return fail(original + ' 已經不在 ' + mdText(w.date) + ' 的名單上了');
      w.slots[at] = name;
      w.subs = w.subs.filter(function (n) { return n !== original; });
      addToRoster(name);
      return Promise.resolve(state('感謝！你代了 ' + original + ' 的 ' + mdText(w.date)));
    }

    return fail('未知的動作');
  }

  return { call: call };
};
