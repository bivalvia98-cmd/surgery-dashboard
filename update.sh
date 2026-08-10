#!/bin/bash
# 대시보드 갱신: 엑셀 읽기 → data.js 재생성 → 변경 있으면 GitHub 배포
# 수동 실행:  ~/Projects/surgery-dashboard/update.sh
set -uo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$REPO" || exit 1
LOG="$REPO/update.log"

log(){ printf '%s  %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$*" | tee -a "$LOG"; }

log "── 갱신 시작 ──"

if [ ! -x "$REPO/.venv/bin/python" ]; then
  log "ERROR: .venv 가 없습니다.  python3 -m venv .venv && .venv/bin/pip install openpyxl pyyaml"
  exit 1
fi

OUT=$("$REPO/.venv/bin/python" "$REPO/build.py" 2>&1)
STATUS=$?
printf '%s\n' "$OUT" | tee -a "$LOG"
if [ $STATUS -ne 0 ]; then
  log "ERROR: 빌드 실패 — 배포 중단"
  exit 1
fi

# 미분류 값이 생기면 눈에 띄게 남겨둔다 (배포는 계속 진행)
if printf '%s' "$OUT" | grep -q '\[미분류\]'; then
  log "주의: 매핑에 없는 새 값이 있습니다. mapping.yaml 을 확인하세요."
fi

if [ -z "$(git status --porcelain docs)" ]; then
  log "변경 없음 — 배포 생략"
  exit 0
fi

N=$(printf '%s' "$OUT" | sed -n 's/.*docs\/data.json 생성 — \([0-9,]*\)건.*/\1/p')
git add docs
git commit -q -m "데이터 갱신: ${N:-?}건 ($(date '+%Y-%m-%d %H:%M'))" || { log "ERROR: 커밋 실패"; exit 1; }

if git remote get-url origin >/dev/null 2>&1; then
  if git push -q origin HEAD 2>>"$LOG"; then
    log "배포 완료 — $(git remote get-url origin)"
  else
    log "ERROR: push 실패 (커밋은 로컬에 저장됨)"
    exit 1
  fi
else
  log "커밋 완료 (origin 미설정 — push 생략)"
fi

log "── 갱신 끝 ──"
