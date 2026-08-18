#!/usr/bin/env python3
"""
수술통계 대시보드 빌드 스크립트
  원본 엑셀(iCloud) → 정규화 → 비식별 집계 JSON(docs/data.json)

* 환자 성명, 등록번호, 수술일 전체날짜, 자유기술 remark 는 출력에 절대 포함되지 않습니다.
* 분류 규칙은 mapping.yaml 에서만 수정합니다. 원본 엑셀은 읽기만 합니다.
"""
from __future__ import annotations

import datetime as dt
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl
import yaml

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "docs" / "data.json"

# 엑셀 컬럼 인덱스 (0-based). 원본 시트의 컬럼 순서가 바뀌면 여기만 고치면 됩니다.
COL = {
    "date": 3, "sex": 6, "age": 7,
    "category": 11, "presentation": 12, "dx": 13, "dx2": 14,
    "location": 15, "side": 16,
    "treatment": 17, "opname": 18, "procedure": 19,
    "attending": 28, "operator": 29,
    "complication": 37,
}
AGE_BANDS = ["~29", "30대", "40대", "50대", "60대", "70대", "80+"]


def s(v) -> str:
    return str(v).strip() if v is not None else ""


def norm(v) -> str:
    """대소문자·공백 무시 비교용 키"""
    return re.sub(r"\s+", " ", s(v)).lower()


def age_band(a) -> str | None:
    if not isinstance(a, (int, float)):
        return None
    a = int(a)
    if a < 30:
        return "~29"
    if a >= 80:
        return "80+"
    return f"{a // 10 * 10}대"


class Lookup:
    """대소문자/공백을 무시하는 매핑 조회 + 미매칭 값 수집"""

    def __init__(self, table: dict, name: str):
        self.name = name
        self.map = {norm(k): v for k, v in (table or {}).items()}
        self.missing: Counter = Counter()

    def get(self, raw, default="미분류"):
        key = norm(raw)
        if not key:
            return None
        if key in self.map:
            return self.map[key]
        self.missing[s(raw)] += 1
        return default


def load_rows(path: Path, sheet: str):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    rows = []
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if all(x is None for x in r):
            continue
        rows.append((i, list(r) + [None] * (44 - len(r))))
    wb.close()
    return rows


def main() -> int:
    cfg = yaml.safe_load((ROOT / "mapping.yaml").read_text(encoding="utf-8"))
    src = Path(cfg["source_file"])
    if not src.exists():
        print(f"[ERROR] 원본 엑셀을 찾을 수 없습니다: {src}", file=sys.stderr)
        return 1

    rows = load_rows(src, cfg["source_sheet"])
    print(f"[info] {src.name}: {len(rows)}행 읽음")

    lk_cat = Lookup(cfg["category"], "category")
    lk_pres = Lookup(cfg["presentation"], "category2")
    lk_tx = Lookup(cfg["treatment"], "Op")
    lk_proc = Lookup(cfg["procedure"], "Op name")
    lk_loc = Lookup(cfg["location"], "location")
    lk_coilt = Lookup(cfg["coil_technique_text"], "coil technique")
    coilt_contains = [(norm(k), v) for k, v in (cfg.get("coil_technique_contains") or {}).items()]

    def coil_from_text(raw):
        """정확히 일치하는 값이 없으면 부분문자열 규칙으로 재시도"""
        key = norm(raw)
        if key in lk_coilt.map:
            return lk_coilt.map[key]
        for frag, val in coilt_contains:
            if frag in key:
                return val
        lk_coilt.missing[s(raw)] += 1
        return "기타"

    # 질환군: {정규화Dx -> 그룹명}
    dx2grp = {}
    for grp, dxs in cfg["disease_group"].items():
        for d in dxs:
            dx2grp[norm(d)] = grp
    grp_missing: Counter = Counter()

    an_cfg = cfg["aneurysm"]
    an_dx = {norm(d) for d in an_cfg["aneurysm_dx"]}
    coil_num = {int(k): v for k, v in cfg["coil_technique_numeric"].items()}
    opname_num = {int(k): v for k, v in cfg["numeric_opname"].items()}

    cases = []
    bad_date = 0
    ages_by_year = defaultdict(list)
    all_ages = []

    for rownum, v in rows:
        d = v[COL["date"]]
        if not isinstance(d, dt.datetime):
            bad_date += 1
            print(f"[warn] row {rownum}: OP Date 없음/형식오류 → 제외")
            continue

        raw_dx = s(v[COL["dx"]])
        raw_op = v[COL["opname"]]

        # ---- 질환군 ----
        if raw_dx:
            grp = dx2grp.get(norm(raw_dx))
            if grp is None:
                grp = "미분류"
                grp_missing[raw_dx] += 1
        else:
            grp = "미분류"

        # ---- 세부 술기 ----
        if isinstance(raw_op, (int, float)) and int(raw_op) in opname_num:
            proc = opname_num[int(raw_op)]
        else:
            proc = lk_proc.get(raw_op) if s(raw_op) else None

        # ---- 동맥류 치료 판정 (mapping.yaml 의 aneurysm 정의) ----
        is_an = 0
        if norm(raw_dx) in an_dx:
            if isinstance(raw_op, (int, float)) and int(raw_op) in (0, 1):
                is_an = 1
            elif an_cfg["include_angio_only"] and norm(raw_op) == "angio only":
                is_an = 1
            elif an_cfg["include_failure"] and norm(raw_op) == "failure of procedure":
                is_an = 1

        # ---- 코일 세부기법 (동맥류 코일 건에 한함) ----
        coilt = None
        if is_an and raw_op == 1:
            p = v[COL["procedure"]]
            if isinstance(p, (int, float)) and int(p) in coil_num:
                coilt = coil_num[int(p)]
            elif s(p):
                coilt = coil_from_text(p)
            else:
                coilt = "기타"

        # ---- 동맥류 위치 (동맥류 건에 한함) ----
        loc = None
        if is_an and s(v[COL["location"]]):
            loc = lk_loc.get(v[COL["location"]], default="기타")

        age = v[COL["age"]]
        if isinstance(age, (int, float)):
            ages_by_year[d.year].append(float(age))
            all_ages.append(float(age))

        cx = v[COL["complication"]]
        cx = int(cx) if isinstance(cx, (int, float)) and int(cx) in (0, 1) else None

        sex = v[COL["sex"]]
        sex = int(sex) if isinstance(sex, (int, float)) and int(sex) in (0, 1) else None

        cases.append({
            "y": d.year, "m": d.month,
            "c": lk_cat.get(v[COL["category"]]) or "미분류",
            "g": grp,
            "d": raw_dx or None,
            "p": lk_pres.get(v[COL["presentation"]]),
            "t": lk_tx.get(v[COL["treatment"]]) or "미분류",
            "o": proc,
            "k": coilt,
            "l": loc,
            "a": s(v[COL["attending"]]) or None,
            "s": sex,
            "b": age_band(age),
            "x": cx,
            "an": is_an,
        })

    # ---------------- 차원 사전으로 압축 ----------------
    order_cfg = {
        "g": cfg["disease_group_order"],
        "t": cfg["treatment_order"],
        "k": cfg["coil_technique_order"],
        "b": AGE_BANDS,
    }
    dim_keys = ["c", "g", "d", "p", "t", "o", "k", "l", "a", "b"]
    dims = {}
    for key in dim_keys:
        vals = {c[key] for c in cases if c[key] is not None}
        pref = order_cfg.get(key)
        if pref:
            ordered = [x for x in pref if x in vals] + sorted(vals - set(pref))
        else:  # 빈도 내림차순
            cnt = Counter(c[key] for c in cases if c[key] is not None)
            ordered = [k for k, _ in cnt.most_common()]
        dims[key] = ordered

    index = {key: {v: i for i, v in enumerate(vals)} for key, vals in dims.items()}
    fields = ["y", "m", "c", "g", "d", "p", "t", "o", "k", "l", "a", "s", "b", "x", "an"]
    packed = []
    for c in cases:
        row = []
        for f in fields:
            v = c[f]
            if v is None:
                row.append(-1)
            elif f in index:
                row.append(index[f][v])
            else:
                row.append(v)
        packed.append(row)

    an_total = sum(c["an"] for c in cases)
    clip = sum(1 for c in cases if c["an"] and c["o"] == opname_num[0])
    coil = sum(1 for c in cases if c["an"] and c["o"] == opname_num[1])

    data = {
        "meta": {
            "generated": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "source": src.name,
            "source_mtime": dt.datetime.fromtimestamp(src.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
            "total": len(cases),
            "date_min": min(f"{c['y']}-{c['m']:02d}" for c in cases),
            "date_max": max(f"{c['y']}-{c['m']:02d}" for c in cases),
            "gaps": cfg.get("gap_periods", []),
            "age_mean": round(sum(all_ages) / len(all_ages), 1) if all_ages else None,
            "age_mean_by_year": {str(y): round(sum(a) / len(a), 1) for y, a in sorted(ages_by_year.items())},
            "aneurysm": {"total": an_total, "clip": clip, "coil": coil,
                         "definition": "Dx=UIA/SAH & (clip/coil"
                                       + (" or angio only" if an_cfg["include_angio_only"] else "")
                                       + (" or failure" if an_cfg["include_failure"] else "") + ")"},
        },
        "fields": fields,
        "dims": dims,
        "cases": packed,
    }

    OUT.parent.mkdir(parents=True, exist_ok=True)

    rc = cfg.get("recent_changes") or {}
    MAX_ITEMS = int(rc.get("max_items", 30))
    KEEP_UPDATES = int(rc.get("keep_updates", 12))

    # 집계 내용이 그대로면 파일을 다시 쓰지 않는다.
    # (생성시각·변경이력만 바뀌어 빈 커밋이 쌓이는 것을 막기 위함)
    VOLATILE = ("generated", "source_mtime", "last_change", "updates")

    def substantive(obj):
        m = {k: v for k, v in obj["meta"].items() if k not in VOLATILE}
        return {"meta": m, "fields": obj["fields"], "dims": obj["dims"], "cases": obj["cases"]}

    def readable(obj):
        """압축된 케이스를 사람이 읽는 dict 로 복원 (빌드 간 대조용)"""
        f, dm = obj["fields"], obj["dims"]
        out = []
        for row in obj["cases"]:
            rec = {}
            for i, name in enumerate(f):
                v = row[i]
                rec[name] = None if v == -1 else (dm[name][v] if name in dm else v)
            out.append(rec)
        return out

    old = None
    if OUT.exists():
        try:
            old = json.loads(OUT.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            old = None
    # 이전 결과가 같은 스키마이면 증감 대조 가능 (이력 유무와 무관)
    comparable = bool(old) and old.get("fields") == data["fields"]
    unchanged = comparable and substantive(old) == substantive(data)

    if unchanged:
        data["meta"]["last_change"] = old["meta"]["last_change"]
        data["meta"]["updates"] = old["meta"]["updates"]
        print("[ok] 집계 결과 동일 — 파일 갱신 생략")
    else:
        now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
        added, removed = [], 0
        if comparable:
            key = lambda r: tuple(str(r[k]) for k in data["fields"])
            new_recs, old_recs = readable(data), readable(old)
            oldc, seen = Counter(map(key, old_recs)), Counter()
            for r in new_recs:
                k = key(r)
                seen[k] += 1
                if seen[k] > oldc[k]:
                    added.append(r)
            removed = sum((oldc - Counter(map(key, new_recs))).values())

        def brief(r):
            # 공개 페이지이므로 날짜는 연·월까지만 (일 단위는 식별 위험)
            return {"날짜": f"{r['y']}-{r['m']:02d}", "질환군": r["g"], "진단": r["d"],
                    "위치": r["l"], "치료": r["t"], "술기": r["o"]}

        entry = {
            "at": now,
            "total": len(cases),
            "added": len(added) if comparable else None,
            "removed": removed,
            "aneurysm": an_total,
            "cases": [brief(r) for r in added[-MAX_ITEMS:]] if rc.get("show", True) else [],
        }
        data["meta"]["last_change"] = now
        prev_hist = (old["meta"].get("updates") or []) if old else []
        data["meta"]["updates"] = (prev_hist + [entry])[-KEEP_UPDATES:]

        payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
        OUT.write_text(payload, encoding="utf-8")
        # file:// 로 열어도 동작하도록 JS 형태로도 함께 출력 (index.html 이 이걸 읽습니다)
        (OUT.parent / "data.js").write_text("window.SURGERY_DATA=" + payload + ";", encoding="utf-8")

        if comparable:
            print(f"[변경] 추가 {len(added)}건" + (f", 삭제 {removed}건" if removed else ""))
            for r in added[-10:]:
                b = brief(r)
                print(f"       + {b['날짜']}  {b['진단']}"
                      + (f" / {b['위치']}" if b["위치"] else "")
                      + f"  {b['술기'] or '-'}")
        else:
            print("[변경] 기준선 생성 — 다음 빌드부터 증감 내역을 기록합니다")

    # ---------------- 리포트 ----------------
    print(f"[ok] {OUT.relative_to(ROOT)} {'확인' if unchanged else '생성'} — "
          f"{len(cases)}건, {OUT.stat().st_size / 1024:.0f} KB")
    print(f"     기간 {data['meta']['date_min']} ~ {data['meta']['date_max']} · "
          f"동맥류 {an_total}건 (clip {clip} / coil {coil})")
    if bad_date:
        print(f"[warn] 날짜 오류로 제외된 행: {bad_date}건")

    problems = 0
    for label, miss in [("질환군(Dx)", grp_missing), ("대분류", lk_cat.missing),
                        ("임상양상", lk_pres.missing), ("치료", lk_tx.missing),
                        ("세부술기", lk_proc.missing), ("동맥류 위치", lk_loc.missing),
                        ("코일기법", lk_coilt.missing)]:
        if miss:
            problems += sum(miss.values())
            print(f"[미분류] {label}: " + ", ".join(f"{k!r}×{n}" for k, n in miss.most_common(20)))
    if problems:
        print(f"\n>>> 매핑에 없는 값 {problems}건이 '미분류'로 집계되었습니다.")
        print(f">>> mapping.yaml 에 추가하면 바로 반영됩니다.")
    else:
        print("[ok] 미분류 값 없음 — 모든 값이 매핑되었습니다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
