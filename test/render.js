/**
 * 前端渲染與互動測試（jsdom，跑在展示模式下）。
 *   npm install && node test/render.js
 */
const fs = require('fs'), path = require('path'), http = require('http');
const { JSDOM, VirtualConsole } = require('jsdom');
const DOCS = path.join(__dirname, '..', 'docs');

const vc = new VirtualConsole();
vc.on('jsdomError', e => { if (!/Could not parse CSS/.test(e.message)) console.log('[jsdomError]', e.message); });

let pass = 0, fail = 0;
const section = n => console.log('\n▍' + n);
const t = (name, fn) => {
  try { fn(); console.log('  ✅', name); pass++; }
  catch (e) { console.log('  ❌', name, '→', e.message); fail++; }
};
const ok = (c, m) => { if (!c) throw new Error(m || '不成立'); };
const eq = (a, b, m) => {
  if (JSON.stringify(a) !== JSON.stringify(b)) throw new Error(`${m ? m + ' ' : ''}期望 ${JSON.stringify(b)}，實得 ${JSON.stringify(a)}`);
};

const MIME = { '.html': 'text/html', '.js': 'text/javascript' };
const server = http.createServer((req, res) => {
  const f = path.join(DOCS, req.url === '/' ? 'index.html' : decodeURIComponent(req.url.split('?')[0]));
  if (!f.startsWith(DOCS) || !fs.existsSync(f)) { res.writeHead(404).end(); return; }
  res.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'text/plain' });
  res.end(fs.readFileSync(f));
}).listen(0);

const dom = new JSDOM(fs.readFileSync(path.join(DOCS, 'index.html'), 'utf8'), {
  // 明確帶 ?demo：config.js 已指向正式後端，不加的話會去打真的 API
  url: 'http://localhost:' + server.address().port + '/?demo',
  runScripts: 'dangerously', resources: 'usable', pretendToBeVisual: true, virtualConsole: vc,
});
const { window } = dom;
const $ = s => window.document.querySelector(s);
const $$ = s => [...window.document.querySelectorAll(s)];
const card = n => $$('.week')[n - 1];                       // 第 n 堂
const btns = el => [...el.querySelectorAll('button[data-act]')].map(b => b.dataset.act + ':' + b.textContent.trim());
const acts = el => btns(el.querySelector('.acts') || el.ownerDocument.createElement('div'));
const txt = el => el.textContent.replace(/\s+/g, ' ').trim();   // 比對時忽略排版空白
const click = el => { ok(el, '按鈕不存在'); el.dispatchEvent(new window.MouseEvent('click', { bubbles: true })); };
const pick = name => { $('#me').value = name; $('#me').dispatchEvent(new window.Event('change')); };
const wait = ms => new Promise(r => setTimeout(r, ms));

window.confirm = () => true;
let promptAnswer = '';
window.prompt = () => promptAnswer;

(async () => {
  await wait(700);

  section('載入與版面');
  t('標題與副標來自設定', () => {
    eq($('#title').textContent, '週二陶藝班');
    const sub = $('#subtitle').textContent;
    ok(sub.includes('每週二 14:00–17:00'), sub);
    ok(sub.includes('共 18 堂'), sub);
    ok(sub.includes('值日生 3 人／每人至少 2 次'), sub);
  });
  t('渲染 18 張卡片', () => eq($$('.week').length, 18));
  t('下一堂課橫幅指向 9/8', () => {
    const b = $$('.banner').find(b => b.textContent.includes('下一堂課'));
    ok(b, '找不到下一堂課橫幅');
    ok(b.textContent.includes('王小明、張淑芬、黃俊傑'), b.textContent);
    ok(b.textContent.includes('2026-09-08'), b.textContent);
  });
  t('停課週標示且無按鈕', () => {
    ok(card(9).textContent.includes('社大公民週'));
    ok(card(9).textContent.includes('不用排值日生'));
    eq(btns(card(9)), []);
    eq(btns(card(18)), []);
  });
  t('已過去的第 1 堂不可操作', () => eq(btns(card(1)), []));
  t('未選名字時卡片沒有操作列', () => eq(card(5).querySelectorAll('.acts').length, 0));

  section('報名');
  pick('蔡佩珊');
  t('顯示個人狀態', () => ok(/蔡佩珊.*還差 1 次/.test($('#mine').textContent), $('#mine').textContent));
  t('第 7 堂有 3 個報名鈕', () => eq(card(7).querySelectorAll('button[data-act="signup"]').length, 3));
  click(card(7).querySelector('button[data-act="signup"]'));
  await wait(60);
  t('名字寫入第 7 堂', () => ok(card(7).textContent.includes('蔡佩珊'), card(7).textContent.trim()));
  t('計數變 1/3', () => eq(card(7).querySelector('.tag').textContent, '1/3'));
  t('個人狀態變已達標', () => ok(/已達標 2\/2/.test($('#mine').textContent), $('#mine').textContent));
  t('出現取消與找代班按鈕', () => eq(acts(card(7)), ['cancel:取消排班', 'requestSub:我不能到，找人代班']));
  t('名字記進 localStorage', () => eq(window.localStorage.getItem('pottery-duty:me'), '蔡佩珊'));

  section('額滿保護');
  pick('劉建宏'); click(card(7).querySelector('button[data-act="signup"]')); await wait(60);
  pick('林志豪'); click(card(7).querySelector('button[data-act="signup"]')); await wait(60);
  t('第 7 堂已滿 3/3', () => eq(card(7).querySelector('.tag').textContent, '3/3'));
  t('滿了就沒有報名鈕', () => eq(card(7).querySelectorAll('button[data-act="signup"]').length, 0));

  section('代班：徵求');
  pick('吳雅婷');
  t('第 4 堂顯示自己可操作', () => eq(acts(card(4)), ['cancel:取消排班', 'requestSub:我不能到，找人代班']));
  promptAnswer = '臨時要看醫生';
  click(card(4).querySelector('button[data-act="requestSub"]'));
  await wait(60);
  t('格子標為徵求代班中', () => {
    const slot = [...card(4).querySelectorAll('.slot')].find(s => s.textContent.includes('吳雅婷'));
    ok(slot.classList.contains('pending'), slot.className);
    ok(slot.textContent.includes('徵求代班中'), slot.textContent);
  });
  t('自己看到的是「取消徵求」', () => eq(acts(card(4)), ['cancelSub:我可以到了，取消徵求']));
  t('頂部出現求救橫幅', () => {
    const b = $$('.banner').find(b => b.textContent.includes('徵求代班'));
    ok(b, '找不到徵求代班橫幅');
    ok(txt(b).includes('有 1 個班需要人接手'), txt(b));
    ok(txt(b).includes('吳雅婷（你） 無法出席'), txt(b));
    ok(txt(b).includes('9/22（二）'), txt(b));
  });

  section('代班：認領');
  pick('黃俊傑');
  t('已排在同一天的人看不到「我來代」', () => eq(acts(card(4)).filter(b => b.startsWith('claimSub')), []));
  pick('蔡佩珊');
  t('其他同學看得到「我來代 吳雅婷」', () => ok(acts(card(4)).includes('claimSub:我來代 吳雅婷'), JSON.stringify(acts(card(4)))));
  t('橫幅也有認領按鈕', () => {
    const b = $$('.banner').find(b => b.textContent.includes('徵求代班'));
    ok(b.querySelector('button[data-act="claimSub"]'), '橫幅缺少認領按鈕');
  });
  click($$('.banner').find(b => b.textContent.includes('徵求代班')).querySelector('button[data-act="claimSub"]'));
  await wait(60);
  t('第 4 堂的吳雅婷換成蔡佩珊', () => {
    const names = [...card(4).querySelectorAll('.slot')].map(s => s.textContent.trim());
    ok(names.includes('蔡佩珊'), JSON.stringify(names));
    ok(!names.some(n => n.includes('吳雅婷')), JSON.stringify(names));
  });
  t('求救橫幅消失', () => ok(!$$('.banner').some(b => b.textContent.includes('有 1 個班需要人接手'))));
  t('顯示成功訊息', () => ok($('#toast').textContent.includes('你代了 吳雅婷'), $('#toast').textContent));

  section('代班：取消徵求');
  pick('劉建宏');
  promptAnswer = '';
  click(card(6).querySelector('button[data-act="requestSub"]'));
  await wait(60);
  t('徵求成立', () => ok($$('.banner').some(b => txt(b).includes('劉建宏（你） 無法出席')),
      $$('.banner').map(txt).join(' || ')));
  click(card(6).querySelector('button[data-act="cancelSub"]'));
  await wait(60);
  t('取消後回到一般狀態', () => eq(acts(card(6)), ['cancel:取消排班', 'requestSub:我不能到，找人代班']));
  t('求救橫幅消失', () => ok(!$$('.banner').some(b => txt(b).includes('無法出席'))));

  section('自行加入名單');
  promptAnswer = '新同學丙';
  pick('__other__');
  await wait(60);
  t('新名字加入下拉選單', () => ok($$('#me option').some(o => o.value === '新同學丙'), '選單無新同學丙'));
  t('自動選為目前身分', () => eq($('#me').value, '新同學丙'));
  t('可以直接報名', () => {
    click(card(11).querySelector('button[data-act="signup"]'));
  });
  await wait(60);
  t('名字出現在第 11 堂', () => ok(card(11).textContent.includes('新同學丙'), card(11).textContent.trim()));

  section('防呆');
  pick('');
  click(card(12).querySelector('button[data-act="signup"]'));
  await wait(60);
  t('未選名字時提示', () => ok($('#toast').textContent.includes('請先在上方選擇你的名字'), $('#toast').textContent));
  t('沒有誤寫入', () => eq(card(12).querySelector('.tag').textContent, '0/3'));

  section('進度統計');
  t('顯示名額比例', () => {
    const h = $$('.stat h2').map(e => e.textContent).join(' | ');
    ok(/已排 \d+ \/ 48 個名額/.test(h), h);   // 16 個上課週 × 3 = 48 格
  });
  t('列出未達 2 次的同學', () => {
    const names = $$('.name').map(e => e.textContent.trim());
    ok(names.length > 0, '未達標名單不該是空的');
    ok(!names.some(n => n.startsWith('王小明')), '王小明已 3 次不該出現');
  });

  console.log(`\n=== ${pass} 通過 / ${fail} 失敗 ===`);
  server.close();
  process.exit(fail ? 1 : 0);
})();
