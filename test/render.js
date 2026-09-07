const fs = require('fs'), path = require('path');
const { JSDOM, VirtualConsole } = require('jsdom');
const DOCS = path.join(__dirname, '..', 'docs');

const vc = new VirtualConsole();
vc.on('jsdomError', e => { if (!/Could not parse CSS/.test(e.message)) console.log('[jsdomError]', e.message); });
vc.on('error', (...a) => console.log('[err]', ...a));

let pass = 0, fail = 0;
const t = (n, fn) => { try { fn(); console.log('  ✅', n); pass++; } catch (e) { console.log('  ❌', n, '→', e.message); fail++; } };
const assert = (c, m) => { if (!c) throw new Error(m || 'assertion failed'); };

const http = require('http');
const MIME = { '.html': 'text/html', '.js': 'text/javascript' };
const server = http.createServer((req, res) => {
  const f = path.join(DOCS, req.url === '/' ? 'index.html' : decodeURIComponent(req.url.split('?')[0]));
  if (!f.startsWith(DOCS) || !fs.existsSync(f)) { res.writeHead(404).end(); return; }
  res.writeHead(200, { 'Content-Type': MIME[path.extname(f)] || 'text/plain' });
  res.end(fs.readFileSync(f));
}).listen(0);
const PORT = server.address().port;

const dom = new JSDOM(fs.readFileSync(path.join(DOCS, 'index.html'), 'utf8'), {
  url: 'http://localhost:' + PORT + '/',
  runScripts: 'dangerously', resources: 'usable', pretendToBeVisual: true, virtualConsole: vc,
});
const { window } = dom;
const $ = s => window.document.querySelector(s);
const $$ = s => [...window.document.querySelectorAll(s)];

setTimeout(() => {
  console.log('\n▍展示模式啟動');
  t('標題來自假資料', () => assert($('#title').textContent === '大同週二拉坏班', $('#title').textContent));
  t('副標顯示規則', () => assert(/每週 3 人.*至少 2 次/.test($('#subtitle').textContent), $('#subtitle').textContent));
  t('顯示展示模式橫幅', () => assert($$('.banner').some(b => b.textContent.includes('這是假資料，可以隨意點')),
      '橫幅：' + $$('.banner').map(b => b.textContent.trim()).join(' | ')));

  console.log('\n▍身分選單');
  t('選單有 8 位同學 + 提示 + 自行輸入', () => assert($$('#me option').length === 10, '實得 ' + $$('#me option').length));
  t('包含錢芸惠', () => assert($$('#me option').some(o => o.value === '錢芸惠')));

  console.log('\n▍週次卡片');
  t('渲染 18 張卡片', () => assert($$('.week').length === 18, '實得 ' + $$('.week').length));
  t('公民週標為不用排', () => {
    const c = $$('.week')[8];
    assert(c.classList.contains('off') && c.textContent.includes('不用排值日生'), c.textContent.trim());
  });
  t('第 1 週顯示三個名字', () => {
    const c = $$('.week')[0];
    ['錢芸惠', '俞光彥', '李築善'].forEach(n => assert(c.textContent.includes(n), '缺 ' + n));
  });
  t('第 1 週標 3/3', () => assert($$('.week')[0].querySelector('.tag').textContent === '3/3'));
  t('第 5 週標 0/3 且有 3 個報名鈕', () => {
    const c = $$('.week')[4];
    assert(c.querySelector('.tag').textContent === '0/3');
    assert(c.querySelectorAll('button[data-act="signup"]').length === 3);
  });
  t('停課週沒有任何按鈕', () => {
    assert($$('.week')[8].querySelectorAll('button').length === 0);
    assert($$('.week')[17].querySelectorAll('button').length === 0);
  });

  console.log('\n▍進度區');
  t('顯示已排名額比例', () => {
    const h = $$('.stat h2').map(e => e.textContent).join(' | ');
    assert(/已排 18 \/ 48 個名額（38%）/.test(h), h);   // 16 個上課週 × 3 = 48 格，截圖已填 18 格
  });
  t('列出未達 2 次的同學', () => {
    const names = $$('.name').map(e => e.textContent.trim());
    assert(names.some(n => n.startsWith('李築善')), JSON.stringify(names));
    assert(!names.some(n => n.startsWith('錢芸惠')), '錢芸惠已 3 次不該出現');
  });

  console.log('\n▍互動：報名');
  const me = $('#me');
  me.value = '李築善';
  me.dispatchEvent(new window.Event('change'));
  t('選好名字後顯示個人狀態', () => assert(/李築善.*還差 1 次/.test($('#mine').textContent), $('#mine').textContent));

  const btn = $$('.week')[4].querySelector('button[data-act="signup"]');
  btn.dispatchEvent(new window.MouseEvent('click', { bubbles: true }));

  setTimeout(() => {
    t('點擊後名字出現在第 5 週', () => assert($$('.week')[4].textContent.includes('李築善'), $$('.week')[4].textContent.trim()));
    t('第 5 週變成 1/3', () => assert($$('.week')[4].querySelector('.tag').textContent === '1/3'));
    t('個人狀態更新為已達標', () => assert(/已達標 2\/2/.test($('#mine').textContent), $('#mine').textContent));
    t('自己的格子變成可取消', () => assert($$('.week')[4].querySelector('button[data-act="cancel"]') !== null));
    t('未達標名單少一人', () => assert(!$$('.name').map(e => e.textContent).some(n => n.startsWith('李築善'))));
    t('顯示成功提示', () => assert($('#toast').textContent.includes('已排定'), $('#toast').textContent));
    t('名字記進 localStorage', () => assert(window.localStorage.getItem('pottery-duty:me') === '李築善'));

    console.log('\n▍互動：未選名字就報名');
    me.value = ''; me.dispatchEvent(new window.Event('change'));
    setTimeout(() => {
      $$('.week')[6].querySelector('button[data-act="signup"]').dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
      setTimeout(() => {
        t('提示先選名字', () => assert($('#toast').textContent.includes('請先在上方選擇你的名字'), $('#toast').textContent));
        t('沒有誤寫入任何人', () => assert($$('.week')[6].querySelector('.tag').textContent === '0/3'));
        console.log(`\n=== ${pass} 通過 / ${fail} 失敗 ===`);
        server.close(); process.exit(fail ? 1 : 0);
      }, 80);
    }, 80);
  }, 120);
}, 700);
