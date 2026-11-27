#!/usr/bin/env bash
# 使い方: deploy-gas.sh <フォルダ（log か records）> <デプロイID>
# Apps Script 上の appsscript.json を取ってきて、リポジトリのコードと一緒に push し、デプロイを新しい版にする。
set -euo pipefail
dir="$1"; deployment_id="${2:-}"
if [ -z "$deployment_id" ]; then echo "$dir のデプロイID（Variables）が設定されていません"; exit 1; fi

work="$(mktemp -d)"
cp "$dir/.clasp.json" "$work/.clasp.json"
cd "$work"
clasp pull                                   # 今の appsscript.json を取ってくる
find . -maxdepth 1 -type f \( -name '*.js' -o -name '*.gs' -o -name '*.html' \) -delete
cd - > /dev/null
cp "$dir"/*.gs "$dir"/*.html "$work"/
cd "$work"
ls -la
clasp push -f
clasp deploy -i "$deployment_id" -d "GitHub ${GITHUB_SHA:0:7}"
