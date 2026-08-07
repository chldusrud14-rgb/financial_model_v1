# 작업 지침

이 저장소는 태양광 PF 재무모델 생성기입니다.
먼저 `docs/SPEC.md`와 `docs/STATUS.md`를 읽으세요.

## 절대 규칙

1. `reference/dangjin_reference.json`이 정답입니다. 여기 값과 다르면 버그입니다.
2. 원본 모델의 기간 단위(분기)를 임의로 바꾸지 마세요.
3. 편의를 위한 단순화 금지. v1이 트랜치를 합치고 연 단위로 접었다가
   Equity IRR이 6%p 틀렸습니다.
4. 엑셀은 생성만으로 검증되지 않습니다. 반드시 재계산해서 화면 값과
   일치하는지 확인하세요: `python scripts/recalc.py <file>.xlsx 240`
5. 화면(app.js)과 엑셀(xlsxbuild.js)은 같은 엔진을 씁니다. 엔진을 고치면
   `npm run test:e2e`로 양쪽이 같이 움직이는지 확인하세요.

## 검증 순서

```bash
npm run test:ref    # 엔진 ↔ 원본 대조
npm run build       # 단일 HTML 번들
npm run test:e2e    # 화면 key-in → 다운로드 → 워크북 캡처 (방식 1/2/3)
python scripts/recalc.py dist/<생성된>.xlsx 240   # 재계산 오류 0 확인
```

## 검증용 테스트

```bash
node test/test_ops.js   # 운영기간(runOps) ↔ reference.results 대조, reference.json에서 직접 읽어옴
```

## 다음 작업

`docs/STATUS.md`의 "미완 — v2에서 해야 할 일" 1번(운영기간 로직)에 이어서:
1. `.xlsm`의 `Opex & Capex!C16:C27`(항목별 기준값), `Assum!F340~`(에스컬레이션),
   `A16:A27`(선순위=1/후순위=0 플래그)에서 정확한 값을 뽑아 하드코딩된
   `opexEok=49.8`/`opexEscal=0.7`/`opexSubShare`(현재 32.9% 균등근사)를 교체
2. 보증발전량/추정발전량 이원화 (4년차부터 일조시간 3.77→3.5) 반영 →
   총영업수익 오차(406,418 vs 405,781) 해결
3. 1·2번 끝나면 최소누적DSCR(1.2428 vs 1.4489), 연차배당(82,991 vs 85,239)
   재검증 — 원본 첫 운영분기(2025-Q4) 값과 직접 대조하면 빠름
   (`Report!C227`=2707.60, `Report!C231`=1868.75)

## `.xlsm` 원본에서 값을 뽑을 때

`python3 -c "import openpyxl; wb=openpyxl.load_workbook('reference/당진_..._final.xlsm', data_only=True); ..."`
로 셀 값을, `data_only=False`로 수식을 읽을 수 있습니다. 시트 이름의 한글이
터미널에서 깨질 수 있으니 `PYTHONIOENCODING=utf-8` 붙이고 필요하면 결과를
파일로 리다이렉트해서 Read 도구로 읽으세요.
