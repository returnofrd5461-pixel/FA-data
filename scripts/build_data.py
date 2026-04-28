"""
build_data.py
raw/ 폴더의 Excel 파일 3종을 읽어 data/data.json을 갱신한다.

입력 (raw/ 안에 있어야 함):
  - 손생보합산_*.xlsx  → D  (통산유지율 by FA·월)
  - 통산유지율_*.xlsx  → LOST (실효·해지 건별 목록)
  - 건별실적_*.xlsx   → PERF (월별 건·보험료 실적)

출력:
  - data/data.json  {"D": {...}, "LOST": {...}, "PERF": {...}}
"""

from __future__ import annotations

import glob
import json
import os
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl 이 설치되지 않았습니다: pip install openpyxl")

ROOT = Path(__file__).parent.parent
RAW  = ROOT / "raw"
OUT  = ROOT / "data" / "data.json"


# ── helpers ──────────────────────────────────────────────────────────────────

def find_file(pattern: str) -> Path:
    matches = sorted(RAW.glob(pattern))
    if not matches:
        sys.exit(f"[ERROR] raw/ 에서 '{pattern}' 파일을 찾을 수 없습니다.")
    if len(matches) > 1:
        print(f"[WARN] '{pattern}' 파일이 여러 개입니다. 최신 파일 사용: {matches[-1].name}")
    return matches[-1]


# ── 손생보합산 → D ────────────────────────────────────────────────────────────

def build_D(path: Path) -> dict:
    """
    예상 컬럼:
      FA명, 팀, 월, 총계약건, 정상, 실효, 해지, 총보험료, rate_25
    컬럼 이름이 다를 경우 아래 COLUMN_MAP 을 수정하세요.
    """
    COLUMN_MAP = {
        "FA명":    "name",
        "팀":      "team",
        "월":      "month",
        "총계약건": "total_contracts",
        "정상":    "normal",
        "실효":    "lapsed",
        "해지":    "cancelled",
        "총보험료": "total_prem",
        "25회차유지율": "rate_25",
    }

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip() if h else "" for h in rows[0]]

    result: dict = {}
    for row in rows[1:]:
        rec = dict(zip(headers, row))
        name  = str(rec.get("FA명", "") or "").strip()
        team  = str(rec.get("팀", "") or "").strip()
        month = str(rec.get("월", "") or "").strip().zfill(2)
        if not name or not month:
            continue

        if name not in result:
            result[name] = {"name": name, "team": team, "months": {}}

        result[name]["months"][month] = {
            "team":            team,
            "rate_25":         float(rec.get("25회차유지율") or 0),
            "total_contracts": int(rec.get("총계약건") or 0),
            "normal":          int(rec.get("정상") or 0),
            "lapsed":          int(rec.get("실효") or 0),
            "cancelled":       int(rec.get("해지") or 0),
            "total_prem":      float(rec.get("총보험료") or 0),
        }

    wb.close()
    return result


# ── 통산유지율 → LOST ─────────────────────────────────────────────────────────

def build_LOST(path: Path) -> dict:
    """
    예상 컬럼:
      FA명, 기간(01_02 등), 보험사, 상품명, 계약자, 초회보험료,
      현재상태, 해지일, 계약일, 납입회차
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip() if h else "" for h in rows[0]]

    result: dict = {}
    for row in rows[1:]:
        rec    = dict(zip(headers, row))
        name   = str(rec.get("FA명", "") or "").strip()
        period = str(rec.get("기간", "") or "").strip()
        if not name or not period:
            continue

        entry = {
            "insurer":      str(rec.get("보험사", "") or ""),
            "product":      str(rec.get("상품명", "") or ""),
            "holder":       str(rec.get("계약자", "") or ""),
            "first_perf":   int(rec.get("초회보험료") or 0),
            "curr_status":  str(rec.get("현재상태", "") or ""),
            "cancel_date":  str(rec.get("해지일", "") or ""),
            "start_date":   str(rec.get("계약일", "") or ""),
            "paid_round":   int(rec.get("납입회차") or 0),
        }

        result.setdefault(name, {}).setdefault(period, []).append(entry)

    wb.close()
    return result


# ── 건별실적 → PERF ───────────────────────────────────────────────────────────

def build_PERF(path: Path) -> dict:
    """
    예상 컬럼:
      FA명, 팀, 월, 건수, 보험료, 실적, 환산, 생보건수, 손보건수,
      정상건수, 실효건수, 해지건수, 무효건수
    상품별 집계(products) 및 보험사별 집계(insurers)는
    컬럼이 있을 경우 자동 수집합니다.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip() if h else "" for h in rows[0]]

    result: dict = {}
    for row in rows[1:]:
        rec   = dict(zip(headers, row))
        name  = str(rec.get("FA명", "") or "").strip()
        team  = str(rec.get("팀", "") or "").strip()
        month = str(rec.get("월", "") or "").strip().zfill(2)
        if not name or not month:
            continue

        if name not in result:
            result[name] = {"name": name, "months": {}}

        result[name]["months"][month] = {
            "cnt":     int(rec.get("건수") or 0),
            "prem":    int(rec.get("보험료") or 0),
            "perf":    int(rec.get("실적") or 0),
            "hwan":    int(rec.get("환산") or 0),
            "life":    int(rec.get("생보건수") or 0),
            "nonlife": int(rec.get("손보건수") or 0),
            "status": {
                "정상": int(rec.get("정상건수") or 0),
                "실효": int(rec.get("실효건수") or 0),
                "해지": int(rec.get("해지건수") or 0),
                "무효": int(rec.get("무효건수") or 0),
            },
            "products":  {},   # TODO: 상품별 컬럼이 있으면 여기에 집계
            "insurers":  {},   # TODO: 보험사별 컬럼이 있으면 여기에 집계
        }

    wb.close()
    return result


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    print("▶ 파일 탐색 중...")
    f_d    = find_file("손생보합산*.xlsx")
    f_lost = find_file("통산유지율*.xlsx")
    f_perf = find_file("건별실적*.xlsx")

    print(f"  D    ← {f_d.name}")
    print(f"  LOST ← {f_lost.name}")
    print(f"  PERF ← {f_perf.name}")

    print("▶ 데이터 변환 중...")
    D    = build_D(f_d)
    LOST = build_LOST(f_lost)
    PERF = build_PERF(f_perf)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with OUT.open("w", encoding="utf-8") as fh:
        json.dump({"D": D, "LOST": LOST, "PERF": PERF}, fh, ensure_ascii=False, indent=2)

    print(f"✅ {OUT} 갱신 완료")
    print(f"   D={len(D)}명  LOST={len(LOST)}명  PERF={len(PERF)}명")


if __name__ == "__main__":
    main()
