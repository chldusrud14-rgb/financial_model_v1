# 화면 인수인계 — 구조와 리테마 가이드

구조·화면은 그대로 쓰되 색상만 다른 서비스와 맞추실 경우를 위한 문서입니다.

## 1. 프론트엔드 파일

| 파일 | 역할 |
|---|---|
| `index2.html` | 레이아웃 + **CSS 전체** (색상 토큰은 `:root` 한 곳) |
| `app2.js` | UI 로직 — 필드 정의, 그리드 생성, 발전원 토글, 입력 수집 |
| `standalone.html` | 위 둘 + 엔진 + ExcelJS를 합친 **완성 단일 파일** (참고용) |

빌드는 `index2.html`의 플레이스홀더를 치환하는 방식입니다.

```
__EXCELJS__      → exceljs.min.js
__DANGJIN_REF__  → 예시 데이터 JSON
__ENGINE2__      → engine2.js
__XLSX2__        → xlsxbuild2.js
__APP2__         → app2.js
```

프레임워크 의존성이 없습니다. jQuery·React 등을 쓰지 않는 순수 DOM 코드라
기존 페이지에 그대로 삽입하거나, 컴포넌트로 감싸도 충돌하지 않습니다.

## 2. 색상 변경 — `:root` 만 고치면 됩니다

CSS 규칙에는 색상 하드코딩이 **0곳**입니다. 모든 색은 `index2.html`
상단 `:root` 블록의 **36개 토큰**을 거쳐서 나갑니다.

### 기본 팔레트 (이것만 바꿔도 톤이 바뀝니다)

| 토큰 | 기본값 | 용도 |
|---|---|---|
| `--brand` | `#2E7D62` | 주 색상 — 선택 상태, 강조 헤더 |
| `--brand-d` | `#14483A` | 진한 주 색상 — 버튼, 제목 |
| `--brand-l` | `#E7F1ED` | 옅은 배경 — 배너, pill |
| `--deep` / `--deep2` | `#0F2E24` / `#173F32` | 최상단 헤더, 토스트 |
| `--accent` | `#4E9E80` | 강조 버튼 |
| `--bg` / `--card` | `#F4F7F5` / `#FFFFFF` | 페이지·카드 바탕 |
| `--ink` / `--ink2` / `--muted` | `#16261F` / `#3D524A` / `#84968E` | 본문·보조·흐린 글자 |
| `--line` | `#E4EBE8` | 테두리 |
| `--warn` / `--bad` / `--good` | `#C2703B` / `#B4483E` / `#2E7D62` | 상태색 |

### 파생 토큰

위 기본색의 명도·채도 변형들입니다. 기본 팔레트만 바꾸고 파생은 그대로 두면
색이 겉돌 수 있으니, 톤을 크게 바꾸실 때는 같이 조정하세요.

* 바탕 계열 — `--field` `--th` `--zebra` `--good-bg`
* 테두리 계열 — `--line2` `--brand-l2` `--brand-b`
* 브랜드 변형 — `--brand-l3` `--brand-h`(버튼 호버) `--accent-h`
* 흐린 글자 — `--dim` `--dim2`
* 주의(앰버) — `--warn-bg` `--warn-bg2` `--warn-line` `--warn-line2` `--warn-ink`
* 위험(레드) — `--bad-bg` `--bad-line` `--bad-dim` `--bad-hero`

### 예시 — 파랑 계열로 전환

```css
:root{
  --deep:#101A33; --deep2:#1B2A52;
  --brand:#2563EB; --brand-d:#1E3A8A; --brand-l:#E8EEFC;
  --brand-l2:#C7D7F8; --brand-l3:#DCE6FB; --brand-b:#AFC6F4;
  --brand-h:#16255C; --accent:#3B82F6; --accent-h:#2F6FE0;
  --good:#2563EB; --good-bg:#F2F6FE;
  --field:#FBFCFE; --th:#F5F8FE; --zebra:#FAFBFE;
  --line:#E3E8F2; --line2:#CBD6EA; --bg:#F4F6FB;
}
```

실제로 적용해서 화면 전체가 파랑으로 바뀌는 것을 확인했습니다.

## 3. 색 외에 바꾸실 만한 것

| 대상 | 위치 |
|---|---|
| 서체 | `index2.html`의 `body{font-family:...}` — 기본 Pretendard |
| 모서리 둥글기 | `--r:14px` |
| 그림자 | `--sh` |
| 2단 레이아웃 비율 | `.cols{grid-template-columns:...}` (기본 1.4 : 0.6) |
| 반응형 분기점 | `@media (max-width:980px)` — 이하에서 1단으로 |

## 4. 주의사항

1. **엑셀 색상은 별개입니다.** 입력 셀 노란색은 `xlsxbuild2.js`의
   `INPUT_FILL = 'FFFFF200'` 상수입니다. CSS 토큰과 연동되지 않으니
   엑셀 톤도 맞추시려면 그쪽을 따로 고치세요.
2. **`.plantBtn`에 `transition`이 걸려 있습니다.** 런타임에 JS로
   CSS 변수를 바꾸면 이 버튼만 다시 칠해지지 않고 이전 색이 남습니다
   (스타일시트를 고치는 정상적인 리테마에서는 문제없음). 다크모드 토글처럼
   런타임 전환을 넣으실 거면 이 점을 감안하세요.
3. `data-plant` 속성이 붙은 요소는 발전원(태양광/풍력) 전환 시
   `display`가 토글됩니다. **이 속성이나 `display` 처리를 CSS로 덮어쓰지 마세요.**
4. `id`는 JS가 직접 참조합니다(`#run` `#xls` `#core` `#trbox` `#spendbox`
   `#shbox` `#capexItemBox` `#opexItemBox` `#sensBox` `#plantHint` 등).
   **id는 바꾸지 마시고**, 스타일은 class로 거세요.
