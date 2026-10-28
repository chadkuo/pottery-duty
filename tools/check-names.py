#!/usr/bin/env python3
"""
擋下含有真實姓名的 commit。

為什麼需要這個 ——
這個 repo 是公開的，示範資料一律用化名，真名只存在試算表裡。但真名會從
意想不到的地方滲進來：從截圖抄一個名字當註解的例子、commit message 裡貼
一段 log、README 舉例。local/sync.py 只保護 apps-script/Code.gs，
其他地方沒有人看著，而那些地方都出過事。

名單從哪裡來 ——
**不在本機存任何名字。** 名單是班長每學期維護的東西，真正的來源是試算表；
在這裡再養一份副本只會過期，而且讓個資多出現一個地方。
所以每次 commit 直接跟正式環境的 API 要一次，用完就丟。

local/api-url.txt 放正式部署的 /exec 網址（那是網址，不是個資，而且不進 git）。
檔案不存在、或連不上時會警告並放行 —— 不能因為沒網路就不讓人 commit。

用法（由 git hook 呼叫）：
  check-names.py staged          檢查已 stage 的檔案內容與檔名
  check-names.py message <檔案>   檢查 commit message
"""
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(subprocess.run(['git', 'rev-parse', '--show-toplevel'],
                           capture_output=True, text=True).stdout.strip())
URL_FILE = ROOT / 'local' / 'api-url.txt'


def warn(msg):
    print(f'\033[33m⚠ 姓名檢查略過：{msg}\033[0m', file=sys.stderr)


def load_names():
    """跟正式環境要一次名單。回傳空清單 = 這次不檢查。"""
    if not URL_FILE.exists():
        warn(f'找不到 {URL_FILE.relative_to(ROOT)}')
        return []
    url = ''
    for line in URL_FILE.read_text(encoding='utf-8').splitlines():
        line = line.split('#', 1)[0].strip()
        if line:
            url = line
            break
    if not url:
        warn(f'{URL_FILE.relative_to(ROOT)} 是空的')
        return []
    # 用 curl 而不是 urllib：macOS 上的 Python 不吃系統憑證庫，
    # urlopen 會以 CERTIFICATE_VERIFY_FAILED 收場。curl 沒這個問題。
    r = subprocess.run(['curl', '-sSL', '--max-time', '20', url],
                       capture_output=True, text=True)
    if r.returncode != 0:                       # 沒網路、Google 掛了、網址過期
        warn(f'讀不到名單（curl {r.returncode}：{r.stderr.strip()[:80]}）')
        return []
    try:
        data = json.loads(r.stdout)
    except ValueError:
        warn('回應不是 JSON（網址可能不對，或部署已失效）')
        return []
    if not data.get('ok'):
        warn(f'API 回錯誤：{data.get("error")}')
        return []
    # 兩個字以下的「名字」當成誤判來源，不拿來比對
    return [n for n in (data.get('roster') or []) if isinstance(n, str) and len(n) >= 2]


def scan(text, names):
    hits = []
    for i, line in enumerate(text.splitlines(), 1):
        for n in names:
            if n in line:
                hits.append((i, n, line.strip()))
    return hits


def mask(name):
    """報告本身不該變成另一份外洩紀錄 —— 終端機捲動紀錄會留下來。"""
    return name[0] + '○' * (len(name) - 1)


def fail(found, total):
    print('\n\033[31m✗ commit 被擋下：內容含有真實姓名\033[0m\n', file=sys.stderr)
    for source, hits in found:
        for lineno, name, line in hits:
            where = f'{source}:{lineno}' if lineno else source
            snippet = line[:60] + ('…' if len(line) > 60 else '')
            print(f'  {where}\n    命中「{mask(name)}」→ {snippet}', file=sys.stderr)
    print(f'\n  （比對正式環境名單上的 {total} 位）', file=sys.stderr)
    print('\n  這個 repo 是公開的。請改用化名，或把那段內容留在本機。', file=sys.stderr)
    print('  確定是誤判要放行：git commit --no-verify\n', file=sys.stderr)
    sys.exit(1)


def check_staged(names):
    files = subprocess.run(
        ['git', 'diff', '--cached', '--name-only', '--diff-filter=ACMR'],
        capture_output=True, text=True).stdout.split()
    found = []
    for f in files:
        name_hits = scan(f, names)                      # 檔名本身也可能含名字
        if name_hits:
            found.append((f'檔名 {f}', [(0, name_hits[0][1], f)]))
        blob = subprocess.run(['git', 'show', f':{f}'], capture_output=True)
        try:
            text = blob.stdout.decode('utf-8')
        except UnicodeDecodeError:
            continue                                    # 二進位檔跳過
        hits = scan(text, names)
        if hits:
            found.append((f, hits))
    if found:
        fail(found, len(names))


def check_message(path, names):
    text = '\n'.join(l for l in Path(path).read_text(encoding='utf-8').splitlines()
                     if not l.startswith('#'))
    hits = scan(text, names)
    if hits:
        fail([('commit message', hits)], len(names))


if __name__ == '__main__':
    names = load_names()
    if not names:
        sys.exit(0)
    if len(sys.argv) > 1 and sys.argv[1] == 'message':
        check_message(sys.argv[2], names)
    else:
        check_staged(names)
