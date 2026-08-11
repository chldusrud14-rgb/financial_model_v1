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

**v2 계산 엔진(운영기간 로직, STATUS.md 미완 1번)은 2026-08-10로 완료.**
`node test/test_ops.js` **13개 지표 전부 완전 일치**(exit 0) —
총영업수익/총영업비용/총선순위이자/총법인세/최소단순DSCR/최소누적DSCR/
연차배당+청산배당/projectIRR 세전·세후/equityIRR(FCFE)/investorIRR/
**dividendIRR** 전부 오차 0.1pp(또는 0.01 KRWm) 미만. 80개 운영분기
실측치는 `reference.json`의 `periodOverrides`에 저장돼 있고 엔진이
있으면 그걸 우선 쓴다.

dividendIRR 잔여오차(부채완제 후 배당가능현금 캡)는 `CF(Q)!row6`
("대출기간" — 부채상환 있는 분기만 그 해로 태그, 완제 후엔 영원히 0)을
찾아내면서 완전히 풀렸다. 상세 산식은 STATUS.md 8번째 라운드 참조.

**4번(분기 엑셀 빌더)도 2026-08-10에 완료.** `src2/xlsxbuild2.js` —
v1과 달리 라이브 수식이 아니라 값(baked) 기준(이유: STATUS.md 4번 참조).
`npm run sample2`로 `dist/`에 생성 확인. 컬럼은 실제 사업기간만(당진
86분기) 채웠고 원본처럼 140분기 여유분은 아직 안 채움(우선순위 낮음).

**5번(UI)도 기능 1차 완료(2026-08-10), 디자인은 미완.**
`src2/index2.html`/`app2.js` — `npm run build2` → `npm run test:ui2`로
엔드투엔드(입력→생성→KPI→엑셀버퍼) 검증됨. **방식 3(64회차)은 화면에
없음**(방식 1/2만) — 의도적 축소 범위, STATUS.md 5번 참조. **비주얼은
v1 CSS 재사용한 기능 우선 레이아웃 — "완성"으로 취급하지 말 것.**
사용자가 "디자인 완성도를 중요하게 봄"이라고 명시했으니 최종 룩앤필은
반드시 확인받을 것.

**6번(배포 정리)도 2026-08-10에 완료.** ExcelJS를 CDN 대신 빌드 시점에
인라인(`__EXCELJS__` 플레이스홀더, `node_modules/exceljs/dist/exceljs.min.js`).
v1/v2 산출물 다 오프라인 동작 확인됨(`npm run test:ui`/`test:ui2`).

**STATUS.md 1~6번 전부 완료.** 남은 건 "남아있는 v1 대비 미구현 항목"
목록(BS 시트/SMP+REC 변동단가/REC 가중치/VPP/재조달) — 우선순위 낮음,
사용자가 필요하다고 하기 전엔 먼저 손대지 말 것.

**주의사항 (다시 밟지 말 것):**
1. **"보증발전량/추정발전량 이원화"는 오탐이었음.** `Revenue!row40`
   확인 결과 일조시간이 20년 내내 3.77로 고정 — 재시도하지 말 것.
2. 감가상각/발전량 감소는 **선형**(`1-deg*idx`)이지 복리(`Math.pow`)가
   아님. 반면 opex 물가 에스컬레이션(일반 2%/yr, O&M 1.5%/yr)은
   **복리**(`Math.pow`)가 맞다 — 같은 "매년 X%" 표현이라도 항목마다
   회계 처리가 다르니 원본에서 직접 확인하고 가정하지 말 것.
3. **회계상 비용(EBT/NI/배당가능이익 계산)과 세무상 손금(법인세 계산)은
   다르다.** 복구충당부채 전입액이 대표 사례 — 회계는 발생주의로 매
   분기 정액 인식하지만, 세법은 실제 현금 지급 시점(만기)에만 손금
   인정한다. 앞으로 새 비용 항목을 추가할 때 "이게 세무상으로도 그
   타이밍에 손금 처리되는지"를 별도로 확인할 것 — `r.ebt`(회계)와
   `r.taxableEbt`(세무) 두 트랙이 이미 분리돼 있으니 그 패턴을 따를 것.
4. 배당은 **12월에 결의, 3월에 지급**(같은 분기가 아님 — `pendingDiv`로
   이월), **배당가능이익 한도**(누적NI-기지급배당, `distributable`)와
   **이익준비금**(실제 지급액의 10%, 자본금 50%까지, `reserveBalance`)과
   **원리금상환재원 기준 캡**(`maxByDscrReserve` = 그 해 3월 지급 전
   기초현금 − 그 해 3월 배당 + 부채상환 있는 분기만의 현금주의 CFADS 합,
   그 값 − 그 해 원리금×목표누적DSCR) 세 개를 모두 만족해야 한다.
   **부채를 완전히 갚은 해부터는 그 해 CFADS가 이 캡에 전혀 반영되지
   않는다**(`CF(Q)!row6` "대출기간" 태그가 완제 후 영원히 0이 되기
   때문 — 원본의 의도인지 부수효과인지는 불명확하나 값은 정확히 이렇게
   나옴). 부채 만기 이후 그래도 남는 잉여현금은 유상감자(당진은 꺼져
   있음) 아니면 그냥 쌓였다가 **청산배당**(운영 마지막 분기, 한도 없이
   잔여현금 전액)으로 나간다. 이 갈래들을 하나로 뭉치지 말 것 —
   `runOps()`의 `distributable`/`reserveBalance`/`maxByDscrReserve`/
   `yrPostMarchCash`/`yrCFCash2`/`lastOp` 처리 참조.
5. `irr()` 함수를 0%에서 좌우로 스캔해 가장 가까운 부호전환 구간을 찾는
   방식으로 바꿨다(예전 단순 이분법은 분기 현금흐름의 잦은 부호전환
   때문에 근이 여러 개일 때 터무니없는 값(수백만%)을 냈음). 다시 단순
   이분법으로 되돌리지 말 것.
6. `test/test_ops.js`는 `dangjin_reference.json`의 `periodOverrides`를
   써서 사실상 "당진 전용"이 됐다 — 이 오버라이드가 없는 다른 프로젝트를
   테스트하려면 공식 근사 경로(opexItems/tariffTracks/seasonality 등)만
   써야 하고, 그 경우 정확도는 이보다 낮다는 걸 감안할 것.

## `.xlsm` 원본에서 값을 뽑을 때

`python3 -c "import openpyxl; wb=openpyxl.load_workbook('reference/당진_..._final.xlsm', data_only=True); ..."`
로 셀 값을, `data_only=False`로 수식을 읽을 수 있습니다. 시트 이름의 한글이
터미널에서 깨질 수 있으니 `PYTHONIOENCODING=utf-8` 붙이고 필요하면 결과를
파일로 리다이렉트해서 Read 도구로 읽으세요.
