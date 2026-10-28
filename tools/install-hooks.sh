#!/bin/sh
# 安裝 git hooks。clone 之後跑一次 —— hooks 不會跟著 git 走，所以這個步驟省不掉。
#
#   sh tools/install-hooks.sh
#
set -e
root=$(git rev-parse --show-toplevel)
hooks=$(git rev-parse --git-path hooks)
mkdir -p "$hooks"

cat > "$hooks/pre-commit" <<'HOOK'
#!/bin/sh
exec python3 -I "$(git rev-parse --show-toplevel)/tools/check-names.py" staged
HOOK

cat > "$hooks/commit-msg" <<'HOOK'
#!/bin/sh
exec python3 -I "$(git rev-parse --show-toplevel)/tools/check-names.py" message "$1"
HOOK

chmod +x "$hooks/pre-commit" "$hooks/commit-msg"
echo "已安裝 pre-commit 與 commit-msg hook"

if [ -f "$root/local/real-names.txt" ]; then
  n=$(grep -cve '^\s*$' -e '^\s*#' "$root/local/real-names.txt" || true)
  echo "名單：local/real-names.txt（$n 個名字）"
else
  echo "⚠ 還沒有 local/real-names.txt —— 名單是空的，檢查不會擋下任何東西"
fi
