# 태양광 PF 재무모델 생성기

당진 태양광발전 FS 재무모델(.xlsm)을 분석해, 웹 화면에서 가정을 key-in 하면
수익률 지표가 계산되고 **수식이 살아있는 엑셀**로 추출되는 툴.

---

## 지금 상태

- **v2 (분기 · 5트랜치)** — **완성.** 건설기간·운영기간 전 구간이 원본과 대조 검증됨
- **v1 (연 단위 · 단일 트랜치)** — 완성이지만 **구버전**. 참고/폴백용으로만 남겨둠

**앞으로의 작업은 v2(`src2/`) 기준입니다.** v1은 손대지 마세요.

### v2가 실제로 검증된 범위

| 항목 | 상태 |
|---|---|
| 건설기간 — 인출 순서, 건설이자(IDC) | 원본과 소수점 일치 |
| 운영기간 — 매출·운영비·감가상각·이자·법인세 | 원본과 오차 0.01 KRWm 미만 |
| 상환 스케줄 — 5트랜치 × 방식 1/2/3 | 원본과 일치 |
| DSCR — 단순/누적/연합산 | 원본과 일치 |
| 배당 — 연차배당 + 청산배당, 3중 캡 | 원본과 일치 |
| IRR — Project(세전/세후)·Equity(FCFE)·배당·Investor | 오차 0.1pp 미만 |
| 엑셀 — **입력값 시트 → 전 시트 라이브 수식 연동** | 순환참조 0, 재계산 불일치 0 |
| 발전원 — 태양광 / 풍력 전환 | 지원 |

`node test/test_ops.js` 가 13개 지표를 `reference/dangjin_reference.json` 과
직접 대조하며, 전부 통과합니다(exit 0).

> **엑셀은 값이 아니라 수식입니다.** "입력값" 시트(노란 탭)의 노란 셀만
> 키인 값이고, 나머지 9개 시트는 전부 그 시트를 참조하는 수식입니다.
> 엑셀에서 금리 하나를 바꾸면 IDC → 상환 → 법인세 → 배당 → IRR 까지
> 다시 계산됩니다.

---

## 바로 쓰려면 (개발 환경 없이)

**`dist/태양광_재무모델_생성기_v2.html`** 을 내려받아 브라우저로 열면 끝입니다.
설치도 인터넷 연결도 필요 없습니다(ExcelJS 까지 파일 안에 들어 있음).

화면에서 **"예시 불러오기 (당진1, 100MW급 PJT)" → "재무모델 생성"** 을 누르면
원본 FS 와 같은 숫자가 나오고, **"Excel 다운로드"** 로 수식이 살아있는
워크북을 받을 수 있습니다.

## 다른 시스템에 붙이려면

**`contract/`** 폴더가 인수인계 패키지입니다. **`contract/README.md` 부터** 읽으세요.

| 목적 | 볼 곳 |
|---|---|
| 입력/출력 데이터 스펙 | `contract/README.md` |
| Django 연동 | `contract/DJANGO.md` |
| Django 앱 (복사만 하면 됨) | `contract/django-app/` + `INSTALL.md` |
| 공통 레이아웃에 삽입 (좌측 메뉴 유지) | `contract/scoped/` + `README-scoped.md` |
| 색상·디자인 변경 | `contract/THEMING.md` |
| 연동 후 검증용 기준값 | `contract/sample-input.json` / `sample-output.json` |

계산 엔진(`contract/engine2.js`)은 **의존성이 전혀 없는 순수 JS** 라
Node 서버에서 그대로 `require()` 해서 쓸 수 있습니다.

---

## 빠른 시작

```bash
npm install          # exceljs, jsdom
npm run build2       # src2/ → dist/태양광_재무모델_생성기_v2.html (단일 파일)
```

빌드된 HTML을 브라우저로 열면 끝입니다. **인터넷 연결이 필요 없습니다** —
ExcelJS까지 파일 안에 들어 있어 오프라인으로 동작합니다.

화면에서 **"예시 불러오기 (당진1, 100MW급 PJT)"** → **"재무모델 생성"** 을
누르면 원본 FS와 같은 숫자가 나옵니다. 예시를 안 불러오고 빈 폼으로
생성하면 범용 근사치가 나오는 게 정상입니다.

### 검증

```bash
npm run test:ref     # 엔진 ↔ 원본 대조 (건설기간: 인출·IDC)
npm run test:ops     # 엔진 ↔ 원본 대조 (운영기간: 13개 지표)
npm run test:ui2     # 화면 key-in → 생성 → KPI → 엑셀 버퍼
npm run test:e2e     # 상환방식 1/2/3 각각 UI→엑셀 E2E
```

엑셀 워크북의 수식을 재계산해서 화면 값과 맞는지까지 보려면:

```bash
python scripts/recalc.py <파일.xlsx> 240
```

> `recalc.py` 는 LibreOffice UNO 로 `calculateAll()` 을 돌립니다.
> **Windows 환경에서는 동작하지 않습니다**(`socket.AF_UNIX` 미지원).
> 그 경우 엑셀 수식을 직접 파싱해 재귀·메모이제이션으로 재평가하는
> Python 스크립트로 대체해서 검증했습니다 — 순환참조·평가오류·값대조를
> 모두 0으로 확인. 방식은 `docs/STATUS.md` 참조.

---

## 디렉터리

```
src2/         v2 — 분기 · 5트랜치  ★ 현재 작업 대상
  engine2.js      계산엔진 (window.SolarModel2) — 브라우저 의존성 없음
  xlsxbuild2.js   ExcelJS 워크북 빌더 (9~10시트, 라이브 수식)
  app2.js         UI: 폼/트랜치표/지출스케줄/민감도/발전원 토글/다운로드
  index2.html     마크업 + CSS (색상 토큰 36개는 :root 한 곳)

src/          v1 — 연 단위 · 단일 트랜치 (구버전, 손대지 말 것)

scripts/      build2.js (단일 HTML 번들), sample2.js, recalc.py
test/         test_dangjin.js (건설기간 대조), test_ops.js (운영기간 대조),
              e2e.js (UI→엑셀 E2E), uitest2*.js (UI 회귀 7종)
reference/    dangjin_reference.json  ← 원본에서 추출한 검증 기준값 전부
docs/         SPEC.md (원본 모델 사양), STATUS.md (경과/함정)
contract/     타 시스템 인수인계 패키지 (git 추적 안 함 — .gitignore)
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
