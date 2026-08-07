# 태양광 PF 재무모델 생성기

당진 태양광발전 FS 재무모델(.xlsm)을 분석해, 웹 화면에서 가정을 key-in 하면
수익률 지표가 계산되고 **수식이 살아있는 엑셀**로 추출되는 툴.

---

## 지금 상태 (한 줄 요약)

- **v1 (연 단위 · 단일 트랜치)** — 완성, 검증 완료, 배포 가능한 단일 HTML 파일
- **v2 (분기 · 5트랜치)** — 건설기간 로직까지 원본과 소수점 일치 확인. 운영기간부터 미완성

**v2를 완성하는 것이 이 작업의 목표입니다.** v1은 참고용/폴백으로 남겨두세요.

---

## 빠른 시작

```bash
npm install                     # exceljs, jsdom
node scripts/build.js           # src/ → dist/태양광_재무모델_생성기.html
node scripts/sample.js          # 당진 가정으로 샘플 xlsx 생성
node test/test_dangjin.js       # v2 엔진 ↔ 원본 대조 (건설기간)
node test/e2e.js 1              # HTML에 key-in → 다운로드 → 워크북 캡처 (상환방식 1)
```

엑셀 재계산 검증은 LibreOffice가 필요합니다:

```bash
python recalc.py <파일.xlsx> 240   # status/total_errors/total_formulas 출력
```

> `recalc.py`는 Anthropic 컨테이너의 `/mnt/skills/public/xlsx/scripts/recalc.py`에
> 있던 스크립트입니다. 로컬에 없으면 LibreOffice UNO로 `calculateAll()` 후
> 에러 셀을 세는 동등한 스크립트를 만들어 쓰면 됩니다.

---

## 디렉터리

```
src/          v1 — 연 단위 · 단일 트랜치 (완성)
  engine.js       계산엔진 (window.SolarModel)
  xlsxbuild.js    ExcelJS 워크북 빌더 (13시트)
  app.js          UI: 폼/KPI/SVG차트/검증/민감도/표/다운로드
  index.html      마크업 + CSS 템플릿 (__ENGINE__ / __XLSX__ / __APP__ 치환)

src2/         v2 — 분기 · 5트랜치 (작업 중)
  engine2.js      기간축 + 건설기간 완성. 운영기간 미검증

scripts/      build.js (단일 HTML 번들), sample.js
test/         test_dangjin.js (원본 대조), e2e.js (UI→엑셀 E2E), uitest.js
reference/    dangjin_reference.json  ← 원본에서 추출한 검증 기준값 전부
docs/         SPEC.md (원본 모델 사양), STATUS.md (완료/미완/함정)
```

---

## 핵심 원칙

1. **원본과 다르면 그건 버그입니다.** `reference/dangjin_reference.json`의
   값이 정답입니다. 편의를 위해 단순화하지 마세요 — v1이 그렇게 하다
   Equity IRR이 6%p 틀렸습니다.
2. **기간 단위를 바꾸지 마세요.** 원본이 분기면 분기로 갑니다.
3. **화면과 엑셀은 같은 코드를 씁니다.** 엔진을 고치면 양쪽이 같이 움직여야
   합니다. E2E 테스트가 이걸 지킵니다.
4. **검증 없이 "됐다"고 하지 마세요.** 엑셀은 생성만으로 부족하고, 재계산해서
   화면 값과 일치하는지까지 확인해야 합니다.
