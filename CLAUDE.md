# FA-data — 프로젝트 가이드

## 1. 프로젝트 개요

FA(금융설계사) 성과·유지율 통합 관리 대시보드. 인카금융서비스 법인전략사업단 운영용.

- **구조**: 단일 페이지(`index.html`) + 외부 데이터(`data/data.json`) 분리
- **빌드 없음**: 정적 HTML/JS. fetch로 data.json 로드.
- **차트**: Chart.js (CDN)
- **데이터 출처**: `raw/YYYY-MM/` 엑셀 원본을 `scripts/build_data.py`로 머지하여 `data/data.json` 생성

## 2. 데이터 구조 (`data/data.json`)

최상위 키:

### `_meta`
```jsonc
{
  "lastUpdated": "ISO 8601",        // 헤더 "📅 최종 업데이트" 표시 원본
  "dataMonths": ["10","11","12","01","02","03","04", ...]
}
```

### `D` — 통산유지율 (손생보합산 원본)
FA별 월 단위 유지율·계약 통계. **모든 FA의 기준이 되는 메인 키.**
```jsonc
"FA명": {
  "name": "...",
  "team": "Team 0|1|2|3|강서지사|서울지사|기타",
  "status": "FA|MA|...",
  "months": {
    "MM": {
      "rate_25": 100.0,        // 25회차 유지율
      "team": "...",            // 월별 팀(이동 추적용)
      "total_contracts": 23,
      "normal": 23,
      "lapsed": 0,              // 실효
      "cancelled": 0,           // 해지
      "total_prem": 1322755.0   // 원 단위
    }
  }
}
```

### `LOST` — 이탈 건별 (유지율 계약리스트 원본)
FA별 월 구간(`mm1_mm2`)에 발생한 실효·해지 건별 raw.
```jsonc
"FA명": {
  "01_02": [
    {
      "insurer": "...",
      "product": "...",
      "holder": "...",          // 계약자
      "first_perf": 75902,      // 초회실적
      "curr_status": "실효|해지",
      "cancel_date": "YYYY-MM-DD",
      "start_date": "YYYY-MM-DD",
      "paid_round": 20          // 납입회차
    }
  ]
}
```

### `PERF` — 신계약 실적 (건별실적 원본을 월 집계)
FA별 월 단위 신계약 통계 + 누적(`totals`).
```jsonc
"FA명": {
  "name": "...",
  "team": "...",
  "status": "...",
  "months": {
    "MM": {
      "cnt": 7,                 // 신계약 건수
      "prem": 516832,           // 월납보험료 (원)
      "perf": 488923,           // 인정실적 (원)
      "hwan": 803424,           // 환산월초 (원)
      "life": 5,                // 생보 건
      "nonlife": 2,             // 손보 건
      "status": { "정상":7, "유예":0, "해지":0, "실효":0 },
      "products": { "보장성":2, "종신/CI":1, ... }
    }
  },
  "totals": {
    "cnt", "prem", "perf", "hwan",
    "life", "nonlife", "lost", "delay",
    "avg_perf",                 // 건당 평균 인정실적
    "life_ratio", "lost_rate",
    "growth": null|숫자,        // 직전 3개월 vs 최근 3개월 증감률 %
    "growth_label": "25.11~26.4",
    "top_products": { ... }
  },
  "insurers": { "메리츠화재": 7, ... }
}
```

### `TARGET` — 활동·목표 (manual.json 기반)
FA·월 단위 DB배정/외활/마감목표.
```jsonc
"FA명": {
  "MM": {
    "db": 15,        // DB 배정 건
    "act": 6,        // 외활 건수
    "goal": 1000000  // 월 마감목표 (원 단위, 월납보험료 기준)
  }
}
```
입력 소스: `raw/YYYY-MM/manual.json` 의 `target` 객체 + `goal` 별도 입력.

### `FEEDBACK` — 피드백 활동 (manual.json 기반)
```jsonc
"FA명": {
  "MM": { "done": 2, "hold": 0 }
}
```

### `ACT` (현재 미사용)
`index.html` 에 `let ACT={}` 로 선언되어 있으나 data.json 에 해당 키 없음. 외활 데이터는 `TARGET.MM.act` 에 통합되어 처리됨. 향후 별도 활동량 데이터 도입 시 사용 예정.

## 3. 활동데이터 제외 정책

DB배정/외활/피드백/마감목표 4영역의 산정·표시에서 제외되는 대상:

- **강서지사 팀 전원** (D 또는 PERF 의 team === '강서지사')
- **`ACTIVITY_EXCLUDED_NAMES` Set 에 명시된 FA** (현재: 민선경)

판정 함수: `isActivityExcluded(name, team)` (`index.html:483`)
```js
function isActivityExcluded(name, team) {
  if (team === KANGSEO) return true;
  if (ACTIVITY_EXCLUDED_NAMES.has(name)) return true;
  return false;
}
function isKangseo(name) { return isActivityExcluded(name, getTeam(name)); }
```

기존 호출처는 `isKangseo(name)` 로 유지되어 있고, 내부적으로 `isActivityExcluded` 로 위임. 신규 제외 대상 추가 시 `ACTIVITY_EXCLUDED_NAMES` Set 한 곳만 수정하면 전체 일관 적용됨.

**제외 적용 범위**:
- 카드 KPI 4셀 (DB/외활/피드백/달성률) → '—'
- 차트 pc4/pc5/pc7/pc6 → '🚫 표시 제외' 노티스
- 통합뷰/랭킹 테이블 셀 → '—'
- 영예 칭호 후보(DB왕/외활왕/피드백왕) 제외
- 자동 랭킹 배지 풀 제외
- 부정 배지(`db_inactive`, `fb_inactive`) 제외
- 정렬 시 `kangseoLastSort` 로 최하단 배치

**제외되지 않는 영역** (정상 처리):
- 실적/유지율/신계약/성장률/생손비중
- 실적왕/신계약왕/보험료왕/유지왕/성장왕 후보

## 4. 서범석(사번 2335304) D 제외 정책

서범석은 단장으로 FA 풀에서 제외. `data/data.json` 의 `D` 키에 포함되지 않음 (raw 추출 단계에서 제외).

- raw 파일명에 등장하는 `2335304` 는 데이터 추출 계정 사번(서범석 본인)이며, FA 본인 데이터로 사용하지 않음.
- TARGET/FEEDBACK 에도 추가하지 않음.

## 5. 일시납 보험료 정제 정책

`scripts/build_data.py` 의 PERF 빌드 단계에서 일시납 건을 행 단위로 제외 (`_check_lump_sum`):

- **룰 A** (연금일시납): 상품명/보험종목에 '연금' 포함 AND 영수보험료 ≥ 1천만원 AND 납입회차 = 1 AND 영수유형 = '초회'
- **룰 B** (고액안전망): 영수보험료 ≥ 5천만원 AND 납입회차 = 1

제외된 행은 `logs/excluded_lump_sum_YYYYMMDD.log` 에 기록 (FA, 계약자, 보험료, 상품명, 사유 코드).

## 6. 호스팅

- **URL**: https://returnofrd5461-pixel.github.io/FA-data/
- **방식**: GitHub Pages (main 브랜치 루트)
- **배포**: `git push` 즉시 반영 (수동 빌드 없음)
- `data.json` / `updates.json` fetch 시 `?v=Date.now()` 쿼리스트링으로 캐시 무효화

## 7. 주요 함수 위치 (`index.html`)

| 함수 | 라인 근처 | 역할 |
|---|---|---|
| `getTeam(name)` | 475 | D → PERF 순으로 팀 조회 |
| `isActivityExcluded(name, team)` | 483 | 활동데이터 제외 판정 (강서지사 + EXCLUDED_NAMES) |
| `isKangseo(name)` | 488 | `isActivityExcluded` 래퍼 (구 호출처 호환) |
| `kangseoLastSort(a, b)` | 489 | 정렬 시 제외 대상 최하단 |
| `allFAs()` | 597 | D∪PERF 의 FA 풀 |
| `perfOfMths(name)` | 821 | 선택월 PERF 집계 (cnt/prem/perf/...) |
| `metaBadges(name, pd)` | 845 | 메타분석 배지 산출 (긍정/부정/교차) |
| `targetOfMths(name)` | 948 | 선택월 TARGET 집계 + 달성률(ach) 계산 |
| `feedbackOfMths(name)` | 1133 | 선택월 FEEDBACK 집계 |
| `computeKings()` | 969 | 영예 칭호(왕) 1위 계산 |
| `newFASectionHTML(name)` | 1164 | 신규 합류 FA(`PERF.totals.cnt===0`) 카드 |
| `perfSectionHTML(name)` | 1206 | 일반 FA 카드 (실적+활동 통합) |
| `renderPerfView(agents)` | 1526 | 통합뷰(랭킹 테이블 + 상세카드 그리드) |
| `render()` | 2169 | 진입점, 필터·정렬 후 view 분기 |
| `renderUpdateNote(entries)` | 2308 | 우측 사이드 업데이트 노트 렌더 |

핵심 상수:
- `KANGSEO = '강서지사'` (474)
- `ACTIVITY_EXCLUDED_NAMES = new Set(['민선경'])` (475)
- `MONTH_ORDER` (466): 회계년도 10월 시작 정렬
- `ML_MAP` (470): 월 코드 → '26.4월' 라벨

## 8. 작업 폴더 구조

```
FA-data/
├── index.html              # 대시보드 단일 페이지 (~2,360 라인)
├── data/
│   ├── data.json           # 빌드 결과 (D/LOST/PERF/TARGET/FEEDBACK/_meta)
│   └── updates.json        # 우측 사이드 업데이트 노트 entries
├── raw/
│   └── YYYY-MM/            # 월별 원본 폴더
│       ├── 손생보합산_*.xlsx           → D
│       ├── 통산유지율_계약리스트_*.xlsx → LOST
│       ├── 건별실적_*.xlsx              → PERF
│       └── manual.json                  → TARGET(db/act) + FEEDBACK + fa_team
├── scripts/
│   └── build_data.py       # raw/ 머지 → data.json
├── logs/
│   └── excluded_lump_sum_*.log
├── README.md
└── CLAUDE.md               # 이 문서
```

**머지 원칙**: `build_data.py` 는 (FA, 월) 단위 덮어쓰기. 기존 월/기간 보존. `TARGET.goal` 과 `FEEDBACK` 은 manual.json 기반이며, `TARGET.goal` 만 수동으로 별도 입력(현재 `data/data.json` 직접 편집).

## 9. 단위 컨벤션

| 항목 | 단위 | 비고 |
|---|---|---|
| `PERF.prem` (월납보험료) | **원** | 표시 시 `fmtW()` 로 `÷10000` 후 '만' 표기 |
| `PERF.perf` (인정실적) | **원** | 동일 |
| `PERF.hwan` (환산월초) | **원** | 동일 |
| `D.total_prem` | **원** | |
| `TARGET.goal` (마감목표) | **원** | 월납보험료 기준. 표시 시 `÷10000` 후 '만원' |
| `TARGET.db` / `TARGET.act` | **건** | |
| `FEEDBACK.done` / `hold` | **건** | |
| `LOST.first_perf` | **원** | 초회실적 |
| 유지율 (`rate_25`, `lost_rate`, `life_ratio`) | **%** | 0~100 |
| 성장률 (`growth`) | **%** | 직전 3개월 vs 최근 3개월 |

**달성률(ach) 계산**: `targetOfMths` 에서 선택월 누적 `PERF.prem ÷ TARGET.goal × 100`. 인정실적이 아닌 **월납보험료** 기준 (실무 관행).

회계년도: **10월 시작** (10/11/12/01/02/03/04/...). 25.4Q = 10·11·12월, 26.1Q = 01·02·03월.
