/**
 * 展示模式用的假資料。
 * 當 config.js 尚未填入後端網址，或網址後面加上 ?demo 時啟用。
 * 這個模式下所有操作只存在瀏覽器記憶體裡，重新整理就會還原，不會寫進試算表。
 */
window.DUTY_DEMO = function () {
  var perWeek = 3, minPerPerson = 2;
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
  var roster = ['錢芸惠', '俞光彥', '李築善', '志平', '宓敦', '高淑梅', '石惠禎', '顏淑琦'];

  var weeks = seed.map(function (r) {
    var slots = [];
    for (var i = 0; i < perWeek; i++) slots.push(r[4][i] || '');
    return { no: r[0], date: r[1], label: r[2], needsDuty: r[3], slots: slots };
  });

  function today() {
    var d = new Date();
    return new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Taipei' }).format(d);
  }

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
      className: '大同週二拉坏班', semester: '2026 秋季班（展示模式）',
      perWeek: perWeek, minPerPerson: minPerPerson,
      allowFreeName: true, open: true,
      today: t, currentWeekNo: cur,
      roster: roster.slice(),
      counts: counts,
      weeks: JSON.parse(JSON.stringify(weeks))
    };
  }

  function call(action, p) {
    p = p || {};
    if (action === 'read') return Promise.resolve(state());
    var w = null;
    for (var i = 0; i < weeks.length; i++) if (weeks[i].no === Number(p.week)) w = weeks[i];
    if (!w) return Promise.resolve({ ok: false, error: '找不到第 ' + p.week + ' 週' });

    if (action === 'signup') {
      if (!w.needsDuty) return Promise.resolve({ ok: false, error: (w.label || '這一週') + '不用排值日生' });
      if (w.slots.indexOf(p.name) >= 0) return Promise.resolve({ ok: false, error: '你已經排在 ' + w.date + ' 了' });
      var free = w.slots.indexOf('');
      if (free < 0) return Promise.resolve({ ok: false, error: w.date + ' 名額已滿，請改選其他日期' });
      w.slots[free] = p.name;
      if (roster.indexOf(p.name) < 0) roster.push(p.name);
      return Promise.resolve(state('已排定 ' + w.date));
    }
    if (action === 'cancel') {
      var idx = w.slots.indexOf(p.name);
      if (idx < 0) return Promise.resolve({ ok: false, error: '你原本就沒有排在 ' + w.date });
      w.slots[idx] = '';
      return Promise.resolve(state('已取消 ' + w.date));
    }
    return Promise.resolve({ ok: false, error: '未知的動作' });
  }

  return { call: call };
};
