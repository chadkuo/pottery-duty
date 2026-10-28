#!/usr/bin/env python3
"""
擋下含有真實姓名的 commit。

為什麼需要這個 ——
這個 repo 是公開的，示範資料一律用化名，真名只存在本機。但真名會從意想不到的
地方滲進來：從截圖抄一個名字當註解的例子、commit message 裡貼一段 log、
README 舉例。local/sync.py 只保護 apps-script/Code.gs，其他檔案沒有人看著。

名單放 local/real-names.txt（已在 .gitignore 裡，不會進 git）。
檔案不存在時直接放行 —— 別人 clone 這個 repo 之後不該被陌生的檢查擋住。

用法（由 git hook 呼叫）：
  check-names.py staged          檢查已 stage 的檔案內容與檔名
  check-names.py message <檔案>   檢查 commit message
"""
import subprocess
import sys
from pathlib import Path

ROOT = Path(subprocess.run(['git', 'rev-parse', '--show-toplevel'],
                           capture_output=True, text=True).stdout.strip())
NAMES_FILE = ROOT / 'local' / 'real-names.txt'


def load_names():
    if not NAMES_FILE.exists():
        return []
    out = []
    for line in NAMES_FILE.read_text(encoding='utf-8').splitlines():
        line = line.split('#', 1)[0].strip()
        if line:
            out.append(line)
    return out


def scan(text, names):
    """回傳 [(行號, 名字, 那一行)]。一行命中多個名字就多筆。"""
    hits = []
    for i, line in enumerate(text.splitlines(), 1):
        for n in names:
            if n in line:
                hits.append((i, n, line.strip()))
    return hits


def mask(name):
    """報告裡不要把完整的名字再印一次 —— 終端機紀錄也會留下。"""
    return name[0] + '○' * (len(name) - 1)


def fail(hits_by_source, names_count):
    print('\n\033[31m✗ commit 被擋下：內容含有真實姓名\033[0m\n', file=sys.stderr)
    for source, hits in hits_by_source:
        for lineno, name, line in hits:
            where = f'{source}:{lineno}' if lineno else source
            snippet = line[:60] + ('…' if len(line) > 60 else '')
            print(f'  {where}', file=sys.stderr)
            print(f'    命中「{mask(name)}」→ {snippet}', file=sys.stderr)
    print(f'\n  （比對 local/real-names.txt 的 {names_count} 個名字）', file=sys.stderr)
    print('\n  這個 repo 是公開的。請改用化名，或把那段內容移到 local/。', file=sys.stderr)
    print('  確定是誤判要放行：git commit --no-verify\n', file=sys.stderr)
    sys.exit(1)


def check_staged(names):
    files = subprocess.run(
        ['git', 'diff', '--cached', '--name-only', '--diff-filter=ACMR'],
        capture_output=True, text=True).stdout.split()

    found = []
    for f in files:
        for lineno, name, _ in scan(f, names):   # 檔名本身也可能含名字
            found.append((f'檔名 {f}', [(0, name, f)]))
            break
        blob = subprocess.run(['git', 'show', f':{f}'], capture_output=True)
        try:
            text = blob.stdout.decode('utf-8')
        except UnicodeDecodeError:
            continue                              # 二進位檔跳過
        hits = scan(text, names)
        if hits:
            found.append((f, hits))
    if found:
        fail(found, len(names))


def check_message(path, names):
    text = Path(path).read_text(encoding='utf-8')
    # 去掉 git 自動加上的註解行
    text = '\n'.join(l for l in text.splitlines() if not l.startswith('#'))
    hits = scan(text, names)
    if hits:
        fail([('commit message', hits)], len(names))


if __name__ == '__main__':
    names = load_names()
    if not names:
        sys.exit(0)
    mode = sys.argv[1] if len(sys.argv) > 1 else 'staged'
    if mode == 'message':
        check_message(sys.argv[2], names)
    else:
        check_staged(names)
