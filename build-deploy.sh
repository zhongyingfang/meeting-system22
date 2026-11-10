#!/usr/bin/env bash
# 生成 seat-finder 部署包（tar.gz + sha256）
# 用法：./build-deploy.sh [日期后缀，默认今天]
# 产物：seat-finder-deploy-<YYYYMMDD>.tar.gz 与同名 .sha256
set -euo pipefail

cd "$(dirname "$0")"

STAMP="${1:-$(date +%Y%m%d)}"
NAME="seat-finder-deploy-${STAMP}"
OUT="${NAME}.tar.gz"
OUT_SHA="${OUT}.sha256"

# 需要打包的文件（保持与历史部署包一致的结构）
FILES=(
  .dockerignore
  .env.example
  .streamlit
  DEPLOY.md
  Dockerfile
  INSTALL.md
  README.md
  app.py
  config.json
  docker-compose.yml
  fonts
  install.bat
  install.sh
  nginx.conf
  package-lock.json
  package.json
  parse-excel.js
  public
  requirements.txt
  server.js
  svg-export.js
  supervisord.conf
  更新说明-20261110.md
)

# 空目录（用 .gitkeep 占位，首次启动会自动生成内容）
EMPTY_DIRS=(
  data
  data/backups
  output
  templates
  uploads
)

STAGE="$(mktemp -d)"
trap 'rm -rf "$STAGE" 2>/dev/null || true' EXIT

DEST="${STAGE}/${NAME}"
mkdir -p "$DEST"

for f in "${FILES[@]}"; do
  if [ ! -e "$f" ]; then
    echo "缺少文件：$f" >&2
    exit 1
  fi
  cp -R "$f" "$DEST/"
done

for d in "${EMPTY_DIRS[@]}"; do
  mkdir -p "$DEST/$d"
  : > "$DEST/$d/.gitkeep"
done

# 打包前自检：关键文件必须在
for key in server.js svg-export.js public/admin.html public/index.html package.json; do
  [ -f "$DEST/$key" ] || { echo "打包内容缺少 $key" >&2; exit 1; }
done

rm -f "$OUT" "$OUT_SHA"
tar -czf "$OUT" -C "$STAGE" "$NAME"
sha256sum "$OUT" > "$OUT_SHA"

echo "已生成：$OUT ($(du -h "$OUT" | cut -f1))"
echo "校验值：$(cat "$OUT_SHA")"
echo "文件数：$(tar -tzf "$OUT" | wc -l)"
