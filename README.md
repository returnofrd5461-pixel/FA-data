# FA-data

FA(금융설계사) 유지율 · 실적 현황 대시보드

## 구조

```
FA-data/
├── index.html          # 대시보드 뷰어 (data.json을 fetch로 로드)
├── data/
│   └── data.json       # 대시보드 데이터 (D / LOST / PERF)
├── scripts/
│   └── build_data.py   # raw/ Excel → data.json 변환 스크립트
└── raw/                # 원본 Excel 파일 (gitignore됨)
```

## 매월 업데이트 절차

1. `raw/` 폴더에 최신 Excel 파일 3종을 넣는다
   - `손생보합산_YYYYMM.xlsx` (또는 손생보합산 형식)
   - `통산유지율_YYYYMM.xlsx`
   - `건별실적_YYYYMM.xlsx`

2. 스크립트를 실행한다
   ```bash
   python scripts/build_data.py
   ```
   → `data/data.json` 이 갱신된다

3. 커밋 & 푸시
   ```bash
   git add data/data.json
   git commit -m "data: YYYY-MM 실적 반영"
   git push
   ```

4. GitHub Pages(또는 배포 환경)에서 `index.html`을 열면 최신 데이터가 로드된다
