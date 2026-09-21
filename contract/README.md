# 재무모델 계산 모듈 — 연동 계약 (Data Contract)

이 모듈은 **브라우저 의존성이 전혀 없는 순수 JS 함수**입니다.
DOM·window·fetch·localStorage를 쓰지 않으므로 Node 서버에서 그대로 호출할 수 있습니다.

```
JSON 입력  →  computeModel()  →  JSON 출력  →  (선택) buildWorkbook() → .xlsx
```

## 1. 넘겨받을 파일

| 파일 | 역할 | 의존성 |
|---|---|---|
| `engine2.js` | **계산 엔진** — 이것만 있으면 계산 가능 | 없음 |
| `xlsxbuild2.js` | 엑셀 생성 (선택) | `exceljs` |
| `sample-input.json` | 입력 예시 (그대로 실행 가능) | — |
| `sample-output.json` | 위 입력의 정답 출력 — **회귀 테스트 기준값** | — |
| `sample-output.xlsx` | 위 입력으로 생성한 엑셀 | — |
| `index2.html` | 화면 레이아웃 + CSS (색상 토큰) | 없음 |
| `app2.js` | 화면 로직 | 없음 |
| `standalone.html` | 전부 합친 완성 단일 파일 | 없음 |
| `THEMING.md` | **화면 구조·리테마 가이드** | — |
| `DJANGO.md` | **Django 연동 가이드** | — |
| `calc.js` | Node CLI 래퍼 (stdin/stdout JSON) | 없음 |
| `solar_fm.py` | Django 쪽 Python 래퍼 | Node |
| `package.json` | `npm install` 용 (exceljs) | — |
| `scoped/` | **공통 레이아웃 삽입용 스코프 격리판** (좌측 메뉴가 있는 구조면 이쪽) | 없음 |
| `django-app/` | **바로 복사해 쓰는 Django 앱** (배치 완료, INSTALL.md 참조) | 없음 |

> 화면까지 가져가신다면 `THEMING.md`를 먼저 읽으세요.

## 2. 호출 방법

```js
const { computeModel } = require('./engine2.js');
const result = computeModel(input);      // 동기 함수, 예외 없음
```

엑셀까지 필요하면:

```js
const ExcelJS = require('exceljs');
const { buildWorkbook } = require('./xlsxbuild2.js');
const wb = buildWorkbook(result, ExcelJS);
await wb.xlsx.writeFile('out.xlsx');     // 또는 wb.xlsx.writeBuffer()
```

`buildWorkbook`은 `result`에 아래 3개를 얹으면 시트가 더 풍부해집니다(없어도 동작).

* `capexItems` — `[{name, amountEok}]` 총사업비 항목별 내역
* `opexDisplayItems` — `[{name, amountEok}]` 운영비 항목별 내역
* `sensitivity` — 민감도 결과(없으면 민감도 시트 생략)

## 3. 입력 스펙

**금액 단위는 억원(`~Eok`), 비율은 % 실수, 기간은 연(year) 기준입니다.**
내부 계산은 백만원(KRWm)·분기로 변환됩니다.

### 3-1. 필수

| 키 | 타입 | 설명 |
|---|---|---|
| `capacityMW` | number | 설비용량 (MW) |
| `tariff` | number | 판매단가 (원/kWh) |
| `capexEok` | number | 총사업비 (억원) |
| `opexEok` | number | 연간 운영비 (억원) |
| `equityEok` | number | 자본금 (억원) |
| `constructionStart` | `"YYYY-MM"` | 착공 시점 |
| `constructionMonths` | number | 공사기간 (개월, **1 이상**) |
| `operationYears` | number | 운영기간 (년) |
| `tranches` | array | 차입 트랜치 — 아래 참조 |

### 3-2. 발전량 — 발전원에 따라 택일

| 태양광 | `dailyHours` | 일일 발전시간 (h/일) |
|---|---|---|
| **풍력** | `capacityFactor` | 이용률 (%) |
| | `availability` | 가동률 (%) |

`dailyHours`가 있으면 태양광 산식, 없으면 풍력 산식을 씁니다.

```
태양광: 연간발전량 = capacityMW × dailyHours × 365
풍력  : 연간발전량 = capacityMW × 8760 × capacityFactor% × availability%
```

### 3-3. `tranches[]` (배열)

| 키 | 타입 | 설명 |
|---|---|---|
| `name` | string | 트랜치명 |
| `amountEok` | number | 약정액 (억원) |
| `rateC` | number | 건설기간 금리 (%) |
| `rateO` | number | 운영기간 금리 (%) |
| `order` | number | **인출 순서** (작을수록 먼저 소진) |
| `graceYears` | number | 거치기간 (년, 소수 가능) |
| `repayYears` | number | 상환기간 (년) |
| `method` | 1 \| 2 \| 3 | 1=원금균등, 2=원리금균등, 3=회차별 비율 키인 |

`equityOrder`로 자본금의 인출 순서를 지정합니다(보통 `1` = 자본금 우선 소진).

### 3-4. 주요 선택값 (생략 시 기본값)

| 키 | 기본 | 설명 |
|---|---|---|
| `ppy` | 4 | 연간 기간 수 (분기=4). **바꾸지 말 것** |
| `degradation` | 0 | 연간 출력저하 (%/yr, **선형**) |
| `auxRate` | 0 | 소내 소비율 (%) |
| `tariffEscal` | 0 | 판매단가 상승률 (%/yr, 복리) |
| `opexEscal` | 0 | 운영비 상승률 (%/yr, **복리**) |
| `depRatio` / `depYears` | 95 / 20 | 감가상각 대상비율(%) / 내용연수 |
| `decomEok` | 0 | 철거·복구비 (억원, 만기 지급) |
| `dsraEok` / `dsraMonths` | 0 / 0 | 최초 DSRA / 적립기준(개월) |
| `minCash` | 0 | 배당 후 최소보유현금 (억원) |
| `divDSCR` / `divCumDSCR` | — | 배당제한 단순/누적 DSCR |
| `divStartYear` | 1 | 배당개시 연차 |
| `taxMode` | — | 1=단일세율(`taxFlat`), 그 외=누진 |
| `lossRate` | 100 | 이월결손금 공제한도 (%) |
| `discount` | 0 | NPV 할인율 (%) |
| `spendCurve` | 균등 | 공사 분기별 투입 비중 배열 (자동 정규화) |
| `shareholders` | 100% 1인 | `[{name, stakePct}]` |
| `arLagMonths` | 0 | 매출채권 회수 시차 (개월) |
| `tariffTracks` | — | 단가 트랙 분리 (PPA/SMP/REC 각각 시차·단가) |
| `opexItems` | — | 운영비 항목별 (`senior`로 선/후순위 구분) |
| `periodOverrides` | — | 기간별 실측치 강제 주입 (있으면 공식보다 우선) |

## 4. 출력 스펙

```
{
  inp,        // 입력 에코백
  periods[],  // 기간 축 — 분기별 {label, start, end, year, isOp, ...}
  rows[],     // 기간별 계산 결과 (periods와 1:1, 길이 동일)
  tranches[], // 트랜치별 상환 스케줄
  con{},      // 건설기간 — 인출·건설이자(IDC) 상세
  idc, tic,   // 건설이자 합계, 총투자비 (KRWm)
  equity, debt,
  kpi{}       // 핵심 지표
}
```

### `rows[]` 주요 필드 (전 34개, 단위 KRWm)

`revenue` `opex` `ebitda` `dep` `ebit` `interest` `ebt` `taxableEbt` `tax` `ni`
`principal` `debtOpen` `debtClose` `ds` `cfads` `dscr` `cumDscr`
`dsraOpen` `dsraClose` `dsraMove` `wc` `fcfe` `cashOpen` `dividend` `cashClose` `projectFcf`

> `ebt`(회계)와 `taxableEbt`(세무)는 **의도적으로 분리**돼 있습니다.
> 복구충당부채처럼 회계는 발생주의, 세법은 현금 지급 시점에만 손금 인정되는
> 항목이 있어서입니다. 하나로 합치지 마세요.

### `kpi{}` (전 23개)

`projectIRR` `projectIRRPre` `equityIRR` `equityIRRPre` `dividendIRR` `investorIRR`
`npv` `paybackYears` `equityMultiple` `minDSCR` `avgDSCR` `minCumDSCR` `minDSCRAnnual`
`avgEbitda` `ebitdaMargin` `totalRevenue` `totalOpex` `totalEbitda` `totalInterest`
`totalTax` `totalDividend` `lcoe` `shareholders`

**IRR은 전부 연환산된 소수**입니다(0.0770 = 7.70%). 화면 표시 시 ×100 하세요.

## 5. 생성되는 엑셀의 구조

| 시트 | 내용 |
|---|---|
| **입력값** (노란 탭) | 화면에서 넣은 값. 노란 셀 = 고치면 전부 다시 계산, 회색 셀 = 스케줄 모양(거치·상환·방식·투입순서·DSRA 적립기준) — 화면에서 다시 생성해야 함 |
| Funding · Debt · Revenue · Opex · IS(Q) · CF(Q) | 분기 계산. **맨 아래 "가정 (입력값 연결)" 블록**에서만 입력값을 참조하고, 위의 계산은 그 시트 안의 셀만 참조 |
| Report · 목차 · (민감도) | 요약 |

원칙은 **"기본값은 입력값에서 한 번만 끌어오고, 계산은 그 시트 안에서"** 입니다.
생성 코드가 마지막 단계(`localizeInputs`)에서 이 규칙을 자동으로 적용하므로,
빌더에 새 계산 행을 추가할 때는 평소처럼 `'입력값'!` 을 써도 됩니다.

## 6. 주의사항

1. **IRR이 `NaN`일 수 있습니다.** 현금흐름에 부호전환이 없으면(매출 0, 자본금 0 등)
   수학적으로 IRR이 존재하지 않아 `NaN`을 반환합니다. **버그가 아닙니다.**
   반드시 `Number.isFinite()`로 확인하고 "산출 불가"로 표시하세요.
2. `NaN` 가능 필드: `projectIRR` `equityIRR` `dividendIRR` `investorIRR` `paybackYears`.
   `null` 가능 필드: `equityMultiple` `ebitdaMargin` `minDSCR` `avgDSCR` (부채·매출 없을 때).
3. **`constructionMonths`는 1 이상**이어야 합니다. 0이면 건설기간이 없어
   총사업비가 현금유출로 잡히지 않고 IRR이 `NaN`이 됩니다.
4. **`ppy`는 4(분기) 고정**입니다. 연 단위로 접으면 Equity IRR이 수%p 틀어집니다.
5. 엔진은 예외를 던지지 않습니다. 잘못된 입력은 예외 대신 `NaN`/`0`으로 나오므로
   **호출 측에서 입력 검증을 하세요.**

## 7. 검증 방법

```bash
node -e "
const {computeModel}=require('./engine2.js');
const got=computeModel(require('./sample-input.json'));
const want=require('./sample-output.json');
const d=Math.abs(got.kpi.equityIRR-want.kpi.equityIRR);
console.log(d<1e-9?'일치':'불일치 '+d);
"
```

`sample-output.json`을 회귀 기준값으로 CI에 넣어두면, 엔진을 갱신했을 때
결과가 바뀌었는지 바로 잡을 수 있습니다.
