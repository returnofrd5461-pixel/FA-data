# FA-data — 프로젝트 가이드

> 클로드 코드 ↔ Claude 웹 양쪽이 공유하는 **단일 소스 컨텍스트**.
> 새로운 세션 시작 시 이 파일을 먼저 참조하면 프로젝트 전체 상태를 빠르게 동기화할 수 있다.

---

## 프로젝트 개요

- **목적**: FA(Financial Advisor) 성과·유지율·신계약·활동 통합 대시보드. 인카금융서비스 법인전략사업단 내부 운영용.
- **구조**: 단일 페이지 HTML(`index.html`) + 외부 데이터 파일(`data/data.json`). 빌드 없음, 정적 호스팅.
- **호스팅**: GitHub Pages — https://returnofrd5461-pixel.github.io/FA-data/
- **대상 인원**: 약 30~40명. 팀 구성:
  - `Team 0` / `Team 1` / `Team 2` / `Team 3` — 4개 본부 팀
  - `강서지사` — 별도 정책(활동데이터 제외) 적용
  - `서울지사` — 코드 라벨. 실무에서는 **프라임지사**로 호칭하기도 함
  - `기타` / `미배정` — 폴백
- **차트 라이브러리**: Chart.js (CDN)

---

## 파일 구조

```
FA-data/
├── index.html              # 메인 대시보드 (단일 페이지, ~2,360 라인)
├── CLAUDE.md               # 본 파일 (프로젝트 컨텍스트)
├── README.md               # 외부용 짧은 안내
├── data/
│   ├── data.json           # D / LOST / PERF / TARGET / FEEDBACK / _meta
│   └── updates.json        # 업데이트 노트 (우측 사이드 패널 entries)
├── raw/
│   └── YYYY-MM/            # 월별 원본 (xls/xlsx + manual.json)
├── scripts/
│   └── build_data.py       # raw/ → data/data.json 머지 스크립트
└── logs/
    └── excluded_lump_sum_*.log
```

`raw/` 폴더는 git 추적 대상이지만 대용량 xls는 .gitignore로 제외될 수 있음(`.gitignore` 참조).

---

## 데이터 구조 (`data/data.json`)

최상위 키: `_meta`, `D`, `LOST`, `PERF`, `TARGET`, `FEEDBACK`.
(`ACT` 키는 현재 존재하지 않음 — 아래 별도 항목 참조)

### `_meta`
```jsonc
{
  "lastUpdated": "2026-05-12T19:28:33+09:00",  // 헤더 "📅 최종 업데이트" 표시
  "dataMonths": ["10","11","12","01","02","03","04"]
}
```

### `D[name]` — 통산유지율 (손생보합산 원본)
모든 FA의 베이스 키. 팀 정보는 여기서 우선 조회됨(`getTeam()` 참조).
```jsonc
"FA명": {
  "name": "...",
  "team": "Team 0|1|2|3|강서지사|서울지사|기타",
  "status": "FA|MA|...",
  "months": {
    "MM": {
      "rate_25": 100.0,        // 25회차 유지율 (%)
      "team": "...",            // 월별 팀(이동 추적)
      "total_contracts": 23,
      "normal": 23,
      "lapsed": 0,              // 실효
      "cancelled": 0,           // 해지
      "total_prem": 1322755.0   // 원 단위
    }
  }
}
```

### `LOST[name][period]` — 이탈 계약 상세
`period` 키 형식: `'MM_MM'` (예: `"01_02"` = 1→2월 사이 신규 이탈).
```jsonc
"FA명": {
  "01_02": [
    {
      "insurer": "DB생명",
      "product": "...",
      "holder": "전수민",         // 계약자
      "first_perf": 75902,        // 초회실적 (원)
      "curr_status": "실효|해지",
      "cancel_date": "YYYY-MM-DD",
      "start_date": "YYYY-MM-DD",
      "paid_round": 20            // 납입회차
    }
  ]
}
```

### `PERF[name]` — 신계약 성과 (건별실적 원본 → 월 집계)
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
      "life": 5, "nonlife": 2,
      "status": { "정상":7, "유예":0, "해지":0, "실효":0 },
      "products": { "보장성":2, "종신/CI":1, ... }
    }
  },
  "totals": {
    "cnt", "prem", "perf", "hwan",
    "life", "nonlife", "lost", "delay",
    "avg_perf",                 // 건당 평균 인정실적
    "life_ratio", "lost_rate",
    "growth": null|숫자,        // 직전 3개월 vs 최근 3개월 증감률 (%)
    "growth_label": "25.11~26.4",
    "top_products": { ... }
  },
  "insurers": { "메리츠화재": 7, ... }
}
```

### `TARGET[name][month]` — 활동·목표
```jsonc
"FA명": {
  "MM": {
    "db": 15,        // DB 배정 건수
    "act": 6,        // 외활 건수
    "goal": 1000000  // 월 마감목표 (원, 월납보험료 기준)
  }
}
```
- `db`/`act`는 `raw/YYYY-MM/manual.json` 의 `target` 객체에서 빌드.
- `goal`은 별도 수동 입력 (현재 `data/data.json` 직접 편집).

### `FEEDBACK[name][month]`
```jsonc
"FA명": {
  "MM": { "done": 2, "hold": 0 }
}
```
`raw/YYYY-MM/manual.json` 의 `feedback` 객체에서 빌드.

### `ACT` (현재 미사용)
- `index.html:539` 에 `let ACT={}` 로만 선언.
- 활동량(고객 미팅 등) 데이터를 **Notion DB → MCP/Anthropic API**로 동적 로드하는 자리표시자.
- 현재 외활/피드백 등 모든 활동성 지표는 `TARGET.MM.act` 와 `FEEDBACK` 에 통합되어 있어 ACT는 사실상 비어 있음.
- 향후 Notion 연계 활성화 시 `loadActivityData()` (`index.html:2200`) 가 ACT를 채움.

---

## 활동데이터 제외 정책

### 대상
- **강서지사 소속 FA 전원** (`getTeam(name) === '강서지사'`)
- **`ACTIVITY_EXCLUDED_NAMES` Set에 등록된 이름**
  - 현재: `new Set(['민선경'])` (`index.html:475`)

### 판정 함수 (`index.html:483-488`)
```js
function isActivityExcluded(name, team) {
  if (team === KANGSEO) return true;
  if (ACTIVITY_EXCLUDED_NAMES.has(name)) return true;
  return false;
}
function isKangseo(name) {
  return isActivityExcluded(name, getTeam(name));
}
```
기존 호출처는 `isKangseo(name)` 로 유지되어 있고 내부에서 `isActivityExcluded` 로 위임. **신규 제외 대상 추가 시 `ACTIVITY_EXCLUDED_NAMES` Set 한 곳만 수정**하면 전체 일관 적용됨.

### 제외 적용 범위 (4개 영역)
DB배정 / 외활 / 피드백 / 마감목표 — 다음 모든 위치에서:
- 카드 KPI 4셀 → `'—'`
- 차트 pc4/pc5/pc7/pc6 → `🚫 표시 제외` 노티스
- 통합뷰 / 랭킹 테이블 셀 → `'—'`
- 영예 칭호 후보: 🎯 DB왕 / 🚶 외활왕 / 💬 피드백왕 후보 풀에서 제외
- 자동 랭킹 배지 풀 제외
- 부정 배지 (`db_inactive`, `fb_inactive`) 제외
- 정렬 시 `kangseoLastSort` 로 최하단 배치

### 제외되지 않는 영역 (정상 처리)
- 실적 / 유지율 / 신계약 / 성장률 / 생손비중
- 실적왕 / 신계약왕 / 보험료왕 / 유지왕 / 성장왕 후보

---

## 데이터 정제 정책

### 서범석 (사번 2335304, 단장) 제외
- 사용자(서범석) 본인은 단장으로 FA 풀에서 제외.
- `D` / `PERF` / `TARGET` / `FEEDBACK` 어디에도 추가하지 않음.
- raw 파일명에 등장하는 `2335304` 는 **데이터 추출 계정 사번**(본인)이며 FA 데이터로 사용하지 않음.
- 원본 xls에 행이 있어도 build_data.py 또는 후처리 단계에서 제외.

### 일시납 보험료 제거
`scripts/build_data.py` 의 `_check_lump_sum()` (~L202) 에서 PERF 행 단위로 검사 후 제외:
- **룰 A** (연금일시납): 상품명/보험종목에 `'연금'` 포함 AND 영수보험료 ≥ 1천만원 AND 납입회차 = 1 AND 영수유형 = `'초회'`
- **룰 B** (고액안전망): 영수보험료 ≥ 5천만원 AND 납입회차 = 1

제외 행은 `logs/excluded_lump_sum_YYYYMMDD.log` 에 (FA, 계약자, 보험료, 상품명, 사유 코드) 형식으로 기록.

이유: 일시납은 월납 KPI/실적/단가를 왜곡함. 대시보드는 **월납 보장성 신계약 기준**이라 일시납은 제거해야 정합성이 맞음.

---

## 단위 컨벤션

코드 직접 확인 (`index.html`, `data/data.json`):

| 항목 | 원본 단위 | 표시 단위 | 변환 |
|---|---|---|---|
| `PERF.prem` (월납보험료) | **원** | **만원** | `fmtW()` → `Math.round(v/10000) + '만'` |
| `PERF.perf` (인정실적) | **원** | **만원** | 동일 |
| `PERF.hwan` (환산월초) | **원** | **만원** | 동일 |
| `D.total_prem` | **원** | **만원** | 동일 |
| `TARGET.goal` (마감목표) | **원** | **만원** | 동일. pc6 차트는 `Math.round(goal/10000)` |
| `TARGET.db` / `TARGET.act` | **건** | **건** | — |
| `FEEDBACK.done` / `hold` | **건** | **건** | — |
| `LOST.first_perf` (초회실적) | **원** | **만원** | 동일 |
| 유지율 / 생보비중 / 손보비중 / 이탈률 | **%** | **%** | 0~100 |
| 성장률 (`growth`) | **%** | **%** | 직전 3개월 vs 최근 3개월 |

**달성률(ach) 계산**: `targetOfMths(name)` 에서 선택월 누적 `PERF.prem ÷ TARGET.goal × 100`. **인정실적이 아닌 월납보험료 기준**(실무 관행).

**회계년도**: 10월 시작. `MONTH_ORDER` (`index.html:466`):
- 25.4Q = `10` / `11` / `12`
- 26.1Q = `01` / `02` / `03`
- 이후 `04` / `05` / ...

---

## 주요 함수·로직 위치 (`index.html`)

| 함수 | 라인 | 역할 |
|---|---|---|
| `getTeam(name)` | 475 | D → PERF 순으로 팀 조회 |
| `isActivityExcluded(name, team)` | 483 | 활동데이터 제외 판정 |
| `isKangseo(name)` | 488 | `isActivityExcluded` 래퍼 (구 호출처 호환) |
| `kangseoLastSort(a, b)` | 489 | 정렬 시 제외 대상 최하단 |
| `actMths()` | 543 | 선택월 배열 반환 |
| `allFAs()` | 597 | D∪PERF 의 FA 풀 |
| `getGrade(d)` | 664 | 유지율 등급 (s/a/b/c/d/n) |
| `perfOfMths(name)` | 821 | 선택월 PERF 집계 (cnt/prem/perf/...) |
| `metaBadges(name, pd)` | 845 | 자동분석 배지 산출 (긍정/부정/교차) |
| `targetOfMths(name)` | 948 | 선택월 TARGET 집계 + ach 계산 (prem ÷ goal × 100) |
| `computeKings()` | 969 | 영예 칭호(왕) 1위 계산 |
| `feedbackOfMths(name)` | 1133 | 선택월 FEEDBACK 집계 |
| `newFASectionHTML(name)` | 1164 | 신규 합류 FA(`PERF.totals.cnt===0`) 카드 |
| `perfSectionHTML(name)` | 1206 | 일반 FA 카드뷰 성과 섹션 (실적+활동 통합) |
| `renderPerfView(agents)` | 1526 | 통합뷰(랭킹 테이블 + FA별 상세카드 그리드) |
| `rankSort(agents)` | 2140 | 유지율 랭킹 정렬 |
| `perfSort(col)` | 2163 | 통합뷰 표 정렬 진입점 (헤더 클릭 핸들러) |
| `render()` | 2169 | 메인 진입점 (필터·정렬 → view 분기) |
| `actOfMths(name)` | 2193 | ACT 집계 (현재 미사용) |
| `loadActivityData()` | 2200 | Notion MCP + Anthropic API로 ACT 로드 (현재 비활성) |
| `downloadImg()` | 2286 | 현재 화면 PNG 이미지 저장 (HTML 다운로드 아님) |
| `renderUpdateNote(entries)` | 2308 | 우측 사이드 업데이트 노트 렌더 |

### 활동 관련 차트
| Canvas ID | 내용 | 제외 정책 |
|---|---|---|
| `pc1` | 신계약 건수 | 적용 안 함 |
| `pc2` | 월납보험료 (만원) | 적용 안 함 |
| `pc3` | 인정실적 (만원) | 적용 안 함 |
| **`pc4`** | DB 배정 (건) | ✓ 제외 대상 |
| **`pc5`** | 외활 (건) | ✓ 제외 대상 |
| **`pc6`** | 마감목표·달성률 (만원) | ✓ 제외 대상 |
| **`pc7`** | 피드백 (건) | ✓ 제외 대상 |

### 핵심 상수
- `KANGSEO = '강서지사'` (474)
- `ACTIVITY_EXCLUDED_NAMES = new Set(['민선경'])` (475)
- `MONTH_ORDER` (466), `ML_MAP` (470)
- `TC` / `TT` (471, 472): 팀 배경/텍스트 색상
- `ALL_TEAMS` (581), `TEAM_RANK` (715)

---

## UI/스타일 컨벤션

### 카드 헤더 라벨 줄바꿈 (한글 글자 단위 줄바꿈 방지)
"라벨 + 칩 병렬" 구조의 섹션 헤더(`.lost-hd`·`.perf-divider` 등, `justify-content:space-between` flex)는 다음 규칙을 지킨다:
- **라벨**(`.lost-title`·`.perf-divider-title`): `white-space:nowrap` + `flex-shrink:0` — 라벨은 절대 접히지 않음. 추가로 `word-break:keep-all`.
- **칩 컨테이너**(`.lost-chips`·`.perf-chips`): `flex-wrap:wrap` + `justify-content:flex-end` + `min-width:0` — 줄바꿈은 칩만 담당.
- **전역 예방**: 카드 내 한글 라벨 텍스트 블록(`.pmc-section-label` 등)에 `word-break:keep-all; overflow-wrap:break-word`.
- 이유: 월 칩이 늘어나면(5개월분 등) 라벨이 squeeze 되어 "이탈 계약 상/세"처럼 글자 단위로 줄바꿈되는 문제 방지.

### 기간 칩 활성 색상 (`MONTH_PALETTE`, `monthColor()`)
- 월별 칩 배경색은 CSS 하드코딩이 아니라 `MONTH_PALETTE`(12색) 배열을 `MONTH_ORDER` 인덱스로 순환 할당(`monthColor(m)`).
- 칩 생성 시 인라인 `--mc` CSS 변수 지정, CSS는 `.mbtn.on{background:var(--mc)}` 한 줄. **월이 늘어도 코드 수정 불필요** (신규 월 색상 누락 재발 방지).
- 팔레트 앞 7색은 기존 25.10~26.4 색상 순서 그대로(회귀 방지).

### 배지 카테고리 가드 (`metaBadges` `cat` 태그)
- 각 배지는 소스 카테고리로 태깅: `cat:'perf'`(실적·성장·단가·생손·활동균등) / `'retain'`(이탈·유예) / `'cross'`(교차) / `'goal'`(마감목표) / `'activity'`(DB·외활·피드백).
- 제외 대상자(`isKangseo`)는 `metaBadges` 말미에서 `cat ∈ {goal, activity}` 배지를 **일괄 스킵** — 개별 배지에 `!isKangseo` 가드 넣지 말 것. 신규 배지는 `cat`만 부여하면 자동 적용.
- 주의: "매월 꾸준한 활동"(`active_consistent`)은 라벨과 달리 소스가 **신계약 건수(PERF.months.cnt)** 라 `cat:'perf'` (제외 대상 아님).

### TARGET 부재 시 카드 골격 통일 (`actHidden`)
- `perfSectionHTML` 은 `actHidden = isKangseo(name) || !tHas` 로 활동/목표 섹션을 판정.
- **TARGET 데이터가 없어도(`tHas=false`) 누적 KPI 5셀 행·활동 차트·마감목표 섹션 골격은 항상 렌더**하고 값만 처리 → 카드 레이아웃 통일(김성한처럼 TARGET 없는 강서지사도 이연식과 동일 구조).
- 노티스: 제외 대상(`kang`) → `🚫 표시 제외`(`kangNotice`), 비제외·데이터부재 → `— 활동 데이터 없음`(`noDataNotice`).

---

## 워크플로우

### 데이터 갱신 주기
**월 단위**. 매월 초 전월 데이터가 확정되면 갱신.

### 작업 PC
- **회사·집 PC 모두**: `C:\Users\서범석\Documents\GitHub\FA-data\` (26.7월 회사 PC 계정 `SEO`→`서범석` 변경으로 경로 통일)
- 두 PC 경로가 동일해 **경로만으로는 구분 불가**. 양쪽 작업으로 인해 **pull 누락 시 커밋 갭 누적** 위험. 항상 작업 시작 시 `git pull --ff-only` 먼저.
- **환경 특이사항 (회사 PC)**: git/python 이 PATH 에 없음 → **GitHub Desktop 번들 git** + **winget user-scope Python 3.12(openpyxl 설치)** 사용.

### 클로드 코드 실행
```bash
cd C:\Users\서범석\Documents\GitHub\FA-data
claude --dangerously-skip-permissions
```

### Git 컨벤션
- **작업 시작**: 무조건 `git pull --ff-only` 먼저. 충돌 시 즉시 보고·대기.
- **작업 완료**: `git add → git commit -m "type(scope): 요약" → git push` 로 마무리.
- **커밋 메시지**: `feat | fix | chore | docs | refactor (scope): 한 줄 요약`

### 월별 데이터 추가 방식 (targeted merge 권장)
- `build_data.py` **전체 재실행은 지양**. `merge_PERF` 가 insurers 를 누적 add 하므로 전체 재스캔 시 기존 월 insurers 가 **이중 누적**됨.
- 대신 **해당 월·해당 영역만 대상으로 targeted merge**: 신월 데이터만 계산해 `data.json` 에 병합하고, 그 외(다른 월/PERF/D/LOST/TARGET/FEEDBACK)는 불변 유지. 병합 후 빌드 전/후 diff 로 "타 영역 불변 + 신월만 추가" 검증.
- 커밋도 영역별 분리(PERF / 유지 D·LOST / 활동 TARGET·FEEDBACK). PERF 만 재빌드할 땐 신월 `건별실적` 만 스캔 범위에 두어 insurers 재누적 방지.

---

## 유지율 데이터 갱신 시 필수 파일 3종

`raw/YYYY-MM/` 에 다음 파일 배치:

1. **이번 달 통산유지율 현황** (`.xls`)
   - 손생보합산 사원_통산유지율현황_YYYYMMDD.xls
   - → `D` 빌드
2. **이번 달 계약 리스트** (`.xlsx`)
   - 통산유지율_계약리스트_YYYYMMDDhhmmss_*.xlsx
   - → `LOST` 빌드 (이번 달 시점 기준 실효/해지 raw)
3. **전월 계약 리스트** (`.xlsx`)
   - 직전 월의 `통산유지율_계약리스트_*.xlsx`
   - → `LOST` 신규 이탈 추출에 **전월과의 diff** 필요
4. (보조) **건별실적** (`.xlsx`)
   - 건별실적_YYYYMMDDhhmmss_2335304.xlsx
   - → `PERF` 빌드
5. (보조) **manual.json**
   - DB배정/외활/피드백/팀 매핑 수동 입력

저장 후 `python scripts/build_data.py` 실행 → `data/data.json` 갱신.

### 데이터 갱신 후 필수 단계 (누락 주의)
`data/data.json` 을 갱신했으면 **반드시 `data/updates.json` 에 항목을 추가**한다. 우측 사이드 업데이트 노트(`renderUpdateNote`)에 표시되는 변경 이력으로, 빠지면 사용자가 무엇이 바뀌었는지 알 수 없다.
- `entries` 배열 **맨 앞**에 `{ "date": "YYYY-MM-DD", "items": [...] }` 추가 (최신이 위).
- `items` 는 간결한 한 줄 bullet. 무엇을(월/대상), 규모(인원/건수/합계), 부수효과(예: 성장률 윈도우 이동)를 기재.
- data.json 과 **같은 커밋**으로 묶는 것을 권장 (별도 커밋으로 뒤늦게 기록하면 누락되기 쉬움).

### PERF 원본 파일 주의
`PERF` 빌드는 **계약 행 단위 `건별실적_*.xlsx`** 만 사용한다. 시스템에서 비슷하게 보이는 **`인정 실적조회(사원)_*.xls`** 는 FA별 **집계 요약표**(HTML 위장 .xls)라 월보험료·환산월초·계약상태·보험회사·계약별 일시납 판정이 없어 사용 불가. `build_data.py` 도 `건별실적*.xlsx` 패턴으로만 검색한다.

---

## Claude 웹 ↔ Claude Code 분업

| 역할 | Claude 웹 | Claude Code |
|---|---|---|
| 이미지/PDF 변환·파싱 | ✓ | △ |
| 새 기능 기획·명세화 | ✓ | △ |
| 데이터 정합성 분석 | ✓ | △ |
| 의사결정 (정책/단위/제외 대상) | ✓ | — |
| 명세 기반 코드 구현 | — | ✓ |
| 데이터 처리 스크립트 작성 | — | ✓ |
| Git 작업 (pull/commit/push) | — | ✓ |
| 파일 검증·diff 확인 | — | ✓ |

**Claude Design**: 디자인 시안. 사용자가 별도 도구로 만들어 보내면 Claude Code가 구현에 반영.

기본 원칙: **Claude 웹이 "무엇을"을 정하고, Claude Code가 "어떻게"를 실행**. 본 CLAUDE.md는 둘 사이 컨텍스트 동기화 매개체.
