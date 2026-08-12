/* ============================================================
   재무모델 Excel 빌더 v2 (ExcelJS) — 분기 · 5트랜치
   - engine2.js의 계산 결과(값)를 그대로 굽는다. v1과 달리 라이브 수식이
     아니다 — 통합투자세액공제/최저한세/이익준비금/배당가능현금 캡처럼
     원본이 반복계산·연도 태그 SUMIF 등으로 얽혀 있어 엑셀 수식만으로
     재현하면 오차가 생긴다(CLAUDE.md "스컬프팅은 수식 자동화 불가"와
     같은 이유 — 그 원칙을 엔진 전체로 확장).
   - 시트: 목차 / Report / Funding / Debt_A..후순위 / Debt_합계 /
           Revenue / Opex / IS(Q) / CF(Q)
   - 열 규칙: B=구분, C=unit, D=합계, E~=분기 1..N
   ============================================================ */
(function (global) {
  'use strict';

  var FONT = '맑은 고딕';
  var FMT_M = '#,##0;(#,##0);"-"';
  var FMT_P = '0.00%';
  var FMT_X = '0.0000"x"';
  var BLACK = 'FF000000', WHITE = 'FFFFFFFF';
  var HDR_FILL = 'FF14483A', SUB_FILL = 'FFE8F1ED';

  function colLetter(n) {
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }

  function buildWorkbook(model, ExcelJSLib) {
    var ExcelJS = ExcelJSLib || global.ExcelJS;
    var wb = new ExcelJS.Workbook();
    wb.creator = '재생E AI Agent — 재무모델 생성기 v2';
    wb.created = new Date();

    var inp = model.inp, periods = model.periods, rows = model.rows, N = periods.length;
    var C0 = 5; // E열 = 첫 분기
    var pc = function (n) { return colLetter(C0 + n); }; // 0-based 분기 인덱스 -> 열문자
    var firstC = pc(0), lastC = pc(N - 1);

    function sheet(name, tab) {
      var ws = wb.addWorksheet(name, {
        views: [{ state: 'frozen', xSplit: 4, ySplit: 7 }],
        properties: { tabColor: { argb: tab || 'FF2E7D62' }, defaultRowHeight: 16 }
      });
      ws.getColumn(1).width = 2.5;
      ws.getColumn(2).width = 26;
      ws.getColumn(3).width = 9;
      ws.getColumn(4).width = 13;
      for (var c = 5; c <= C0 + N; c++) ws.getColumn(c).width = 11;
      return ws;
    }
    function title(ws, name) {
      ws.getCell('A1').value = inp.projectName || '태양광 재무모델';
      ws.getCell('A1').font = { name: FONT, bold: true, size: 13, color: { argb: 'FF14483A' } };
      ws.getCell('A2').value = name;
      ws.getCell('A2').font = { name: FONT, bold: true, size: 10, color: { argb: 'FF6B7B76' } };
    }
    function section(ws, r, text) {
      var c = ws.getCell('B' + r);
      c.value = text;
      c.font = { name: FONT, bold: true, size: 11, color: { argb: WHITE } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HDR_FILL } };
      for (var k = 3; k <= C0 + N; k++) {
        ws.getCell(colLetter(k) + r).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HDR_FILL } };
      }
    }
    function periodHeader(ws, r) {
      ['구  분', 'unit', '합 계'].forEach(function (t, idx) {
        var c = ws.getCell(colLetter(2 + idx) + r);
        c.value = t;
        c.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
        c.alignment = { horizontal: 'center' };
      });
      for (var n = 0; n < N; n++) {
        var c1 = ws.getCell(pc(n) + r);
        c1.value = periods[n].endStr;
        c1.numFmt = '@';
        c1.font = { name: FONT, bold: true, size: 8, color: { argb: WHITE } };
        c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
        c1.alignment = { horizontal: 'center', textRotation: 90 };
      }
    }
    function label(ws, r, text, unit, opt) {
      opt = opt || {};
      var c = ws.getCell('B' + r);
      c.value = (opt.indent ? '  ' : '') + text;
      c.font = { name: FONT, size: 10, bold: !!opt.bold, color: { argb: BLACK } };
      if (opt.fill) {
        for (var k = 2; k <= C0 + N; k++) {
          ws.getCell(colLetter(k) + r).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opt.fill } };
        }
      }
      if (unit) {
        var u = ws.getCell('C' + r);
        u.value = unit;
        u.font = { name: FONT, size: 9, color: { argb: 'FF9AA6A1' } };
        u.alignment = { horizontal: 'center' };
      }
      return r;
    }
    function put(ws, addr, value, fmt, opt) {
      opt = opt || {};
      var c = ws.getCell(addr);
      c.value = value;
      c.numFmt = fmt || FMT_M;
      c.font = { name: FONT, size: 10, bold: !!opt.bold, color: { argb: opt.color || BLACK } };
      if (opt.fill) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opt.fill } };
      return c;
    }
    // opt.noSum: 잔액류(스톡)는 분기 합계가 의미 없으므로 D열 합계를 안 채운다.
    function fillPeriods(ws, r, fn, fmt, opt) {
      opt = opt || {};
      var sum = 0, hasNull = false;
      for (var n = 0; n < N; n++) {
        var v = fn(n);
        if (v === null || v === undefined) { hasNull = true; put(ws, pc(n) + r, null, fmt, opt); continue; }
        sum += v;
        put(ws, pc(n) + r, v, fmt, opt);
      }
      if (!hasNull && !opt.noSum) put(ws, 'D' + r, sum, fmt, Object.assign({ bold: true }, opt));
    }

    /* =========================================================
       1. Funding — 트랜치·자금조달
       ========================================================= */
    (function () {
      var ws = wb.addWorksheet('Funding', { properties: { tabColor: { argb: 'FF14483A' } } });
      ws.getColumn(1).width = 2.5; ws.getColumn(2).width = 14;
      ['금액[KRWm]', '투입순서', '건설금리', '운영금리', '거치(yr)', '상환(yr)', '방식', '상환개시'].forEach(function (h, idx) {
        ws.getColumn(3 + idx).width = 12;
      });
      title(ws, '자금조달 — 자본금 + 5개 트랜치');
      var r = 4;
      section(ws, r, '트랜치 조건'); r += 2;
      ['트랜치', '금액[KRWm]', '투입순서', '건설금리', '운영금리', '거치(yr)', '상환(yr)', '방식', '상환개시'].forEach(function (h, idx) {
        var c = ws.getCell(colLetter(2 + idx) + r);
        c.value = h;
        c.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
        c.alignment = { horizontal: 'center' };
      });
      r++;
      put(ws, 'B' + r, '자본금'); put(ws, 'C' + r, model.equity, FMT_M); r++;
      model.tranches.forEach(function (t) {
        put(ws, 'B' + r, t.name);
        put(ws, 'C' + r, t.amount, FMT_M);
        put(ws, 'D' + r, t.order === undefined ? '' : t.order, '0');
        put(ws, 'E' + r, t.rateO, FMT_P);
        put(ws, 'F' + r, t.rateO, FMT_P);
        put(ws, 'G' + r, t.graceYears, '0.00');
        put(ws, 'H' + r, t.repayYears, '0.00');
        put(ws, 'I' + r, t.method, '0');
        put(ws, 'J' + r, t.repayStartIdx >= 0 ? periods[t.repayStartIdx].endStr : '-', '@');
        r++;
      });
      r += 1;
      section(ws, r, '건설이자(IDC)'); r += 2;
      put(ws, 'B' + r, 'IDC 합계[KRWm]'); put(ws, 'C' + r, model.idc, FMT_M); r++;
      put(ws, 'B' + r, '총투자비(TIC)[KRWm]'); put(ws, 'C' + r, model.tic, FMT_M); r++;
      put(ws, 'B' + r, '총사업비(건설이자 제외)[KRWm]'); put(ws, 'C' + r, inp.capexEok * 100, FMT_M); r++;

      var sh = model.kpi && model.kpi.shareholders;
      if (sh && sh.length > 1) {
        r += 1;
        section(ws, r, '사업자 구성 (지분)'); r += 2;
        ['사업자', '지분율', '출자금액[KRWm]', '누적배당[KRWm]'].forEach(function (h, idx) {
          var c = ws.getCell(colLetter(2 + idx) + r);
          c.value = h;
          c.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
          c.alignment = { horizontal: 'center' };
        });
        r++;
        sh.forEach(function (s) {
          put(ws, 'B' + r, s.name);
          put(ws, 'C' + r, s.stakePct / 100, FMT_P);
          put(ws, 'D' + r, s.equityKRWm, FMT_M);
          put(ws, 'E' + r, s.dividendKRWm, FMT_M);
          r++;
        });
      }
    })();

    /* =========================================================
       2. Debt — 트랜치별 5개 시트 + 합계
       ========================================================= */
    model.tranches.forEach(function (t, ti) {
      var ws = sheet('Debt_' + t.name.replace(/\s/g, ''));
      title(ws, '차입금 상환 스케줄 — ' + t.name);
      section(ws, 4, '인출/상환');
      periodHeader(ws, 6);
      label(ws, 9, '기초잔액', '[KRWm]');
      var bal = 0;
      var opens = [], closes = [], ints = [], prins = [];
      for (var n = 0; n < N; n++) {
        opens.push(bal);
        var draw = t.draws[n] || 0;
        bal += draw;
        var open2 = bal - draw;
        var interest = periods[n].isOp ? open2 * t.rateO / (inp.ppy || 4) : 0;
        var prin = (rows[n].principalBy && rows[n].principalBy[ti]) || 0;
        bal -= prin;
        ints.push(interest); prins.push(prin); closes.push(bal);
      }
      fillPeriods(ws, 9, function (n) { return opens[n]; }, FMT_M, { noSum: true });
      label(ws, 10, '인출', '[KRWm]');
      fillPeriods(ws, 10, function (n) { return t.draws[n] || 0; }, FMT_M);
      label(ws, 11, '건설이자(IDC)', '[KRWm]');
      fillPeriods(ws, 11, function (n) { return t.idcSeries[n] || 0; }, FMT_M);
      label(ws, 12, '이자(운영)', '[KRWm]');
      fillPeriods(ws, 12, function (n) { return ints[n]; }, FMT_M);
      label(ws, 13, '원금상환', '[KRWm]');
      fillPeriods(ws, 13, function (n) { return prins[n]; }, FMT_M);
      label(ws, 14, '기말잔액', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillPeriods(ws, 14, function (n) { return closes[n]; }, FMT_M, { bold: true, noSum: true });
      label(ws, 16, '미상환 잔액(검증용)', '[KRWm]', { bold: true });
      var finalBal = closes[N - 1];
      put(ws, 'D16', finalBal, FMT_M, { bold: true });
      put(ws, 'F16', Math.abs(finalBal) < 1 ? '완전상환 확인 (OK)' : '경고: 미상환 잔액', null, { bold: true });
    });
    (function () {
      var ws = sheet('Debt_합계', 'FF14483A');
      title(ws, '차입금 상환 스케줄 — 전체 합계');
      section(ws, 4, '전체 트랜치 합계');
      periodHeader(ws, 6);
      label(ws, 9, '기초잔액', '[KRWm]');
      fillPeriods(ws, 9, function (n) { return rows[n].debtOpen; }, FMT_M, { noSum: true });
      label(ws, 10, '이자', '[KRWm]');
      fillPeriods(ws, 10, function (n) { return rows[n].interest; }, FMT_M);
      label(ws, 11, '원금상환', '[KRWm]');
      fillPeriods(ws, 11, function (n) { return rows[n].principal; }, FMT_M);
      label(ws, 12, '원리금(DS)', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillPeriods(ws, 12, function (n) { return rows[n].ds; }, FMT_M, { bold: true });
      label(ws, 13, '기말잔액', '[KRWm]');
      fillPeriods(ws, 13, function (n) { return rows[n].debtClose; }, FMT_M, { noSum: true });
      label(ws, 15, 'DSRA 기말잔액', '[KRWm]');
      fillPeriods(ws, 15, function (n) { return rows[n].dsraClose; }, FMT_M, { noSum: true });
    })();

    /* =========================================================
       3. Revenue
       ========================================================= */
    (function () {
      var ws = sheet('Revenue');
      title(ws, '영업수익 추정');
      section(ws, 4, '발전량 및 매출');
      periodHeader(ws, 6);
      label(ws, 9, '발전량', '[MWh]');
      fillPeriods(ws, 9, function (n) { return periods[n].gen; }, '#,##0');
      label(ws, 10, '판매단가', '[원/kWh]');
      fillPeriods(ws, 10, function (n) { return periods[n].isOp ? periods[n].price : null; }, '#,##0.0', { noSum: true });
      label(ws, 12, '영업수익', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillPeriods(ws, 12, function (n) { return rows[n].revenue; }, FMT_M, { bold: true });
    })();

    /* =========================================================
       4. Opex
       ========================================================= */
    (function () {
      var ws = sheet('Opex');
      title(ws, '영업비용 추정');
      section(ws, 4, '운영비용 (선순위/후순위 지급순위 반영)');
      periodHeader(ws, 6);
      label(ws, 9, '선순위운영비', '[KRWm]');
      fillPeriods(ws, 9, function (n) { return periods[n].opexSenior || 0; }, FMT_M);
      label(ws, 10, '후순위운영비', '[KRWm]');
      fillPeriods(ws, 10, function (n) { return periods[n].opexSub || 0; }, FMT_M);
      label(ws, 11, '영업비용 합계', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillPeriods(ws, 11, function (n) { return rows[n].opex; }, FMT_M, { bold: true });
    })();

    /* =========================================================
       5. IS(Q) — 손익계산서
       ========================================================= */
    (function () {
      var ws = sheet('IS(Q)');
      title(ws, '추정 손익계산서 (분기)');
      section(ws, 4, '손익계산서');
      periodHeader(ws, 6);
      label(ws, 9, '영업수익', '[KRWm]');
      fillPeriods(ws, 9, function (n) { return rows[n].revenue; }, FMT_M);
      label(ws, 10, '영업비용', '[KRWm]');
      fillPeriods(ws, 10, function (n) { return -rows[n].opex; }, FMT_M);
      label(ws, 11, 'EBITDA', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillPeriods(ws, 11, function (n) { return rows[n].ebitda; }, FMT_M, { bold: true });
      label(ws, 12, '감가상각비', '[KRWm]');
      fillPeriods(ws, 12, function (n) { return -rows[n].dep; }, FMT_M);
      label(ws, 13, '복구충당부채 전입액', '[KRWm]');
      fillPeriods(ws, 13, function (n) { return -(rows[n].decomAccrual || 0); }, FMT_M);
      label(ws, 14, '영업이익(EBIT)', '[KRWm]', { bold: true });
      fillPeriods(ws, 14, function (n) { return rows[n].ebit; }, FMT_M, { bold: true });
      label(ws, 15, '이자비용', '[KRWm]');
      fillPeriods(ws, 15, function (n) { return -rows[n].interest; }, FMT_M);
      label(ws, 16, '대리은행수수료', '[KRWm]');
      fillPeriods(ws, 16, function (n) { return -(rows[n].agentFee || 0); }, FMT_M);
      label(ws, 17, '법인세차감전순이익', '[KRWm]', { bold: true });
      fillPeriods(ws, 17, function (n) { return rows[n].ebt; }, FMT_M, { bold: true });
      label(ws, 18, '법인세비용', '[KRWm]');
      fillPeriods(ws, 18, function (n) { return -rows[n].tax; }, FMT_M);
      label(ws, 19, '당기순이익', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillPeriods(ws, 19, function (n) { return rows[n].ni; }, FMT_M, { bold: true });
    })();

    /* =========================================================
       6. CF(Q) — 현금흐름 / DSCR / 배당
       ========================================================= */
    (function () {
      var ws = sheet('CF(Q)');
      title(ws, '추정 현금흐름표 (분기) · DSCR · 배당');
      section(ws, 4, '영업/재무 현금흐름');
      periodHeader(ws, 6);
      label(ws, 9, 'CFADS (원리금상환재원)', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillPeriods(ws, 9, function (n) { return rows[n].cfads; }, FMT_M, { bold: true });
      label(ws, 10, '원리금(DS)', '[KRWm]');
      fillPeriods(ws, 10, function (n) { return rows[n].ds; }, FMT_M);
      label(ws, 11, '단순 DSCR', '[x]', { bold: true });
      fillPeriods(ws, 11, function (n) { return rows[n].dscr; }, FMT_X, { bold: true, noSum: true });
      label(ws, 13, 'FCFE', '[KRWm]', { bold: true });
      fillPeriods(ws, 13, function (n) { return rows[n].fcfe; }, FMT_M, { bold: true });
      label(ws, 14, 'DSRA 증감', '[KRWm]');
      fillPeriods(ws, 14, function (n) { return rows[n].dsraMove; }, FMT_M);
      label(ws, 15, '기초현금', '[KRWm]');
      fillPeriods(ws, 15, function (n) { return rows[n].cashOpen; }, FMT_M, { noSum: true });
      label(ws, 16, '배당(연차+청산)', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillPeriods(ws, 16, function (n) { return rows[n].dividend; }, FMT_M, { bold: true });
      label(ws, 17, '기말현금', '[KRWm]');
      fillPeriods(ws, 17, function (n) { return rows[n].cashClose; }, FMT_M, { noSum: true });

      var r = 20;
      section(ws, r, '검증'); r += 2;
      label(ws, r, '최소 단순DSCR(연 합산)', '[x]', { bold: true });
      put(ws, 'D' + r, model.kpi.minDSCRAnnual, FMT_X, { bold: true }); r++;
      label(ws, r, '최소 누적DSCR', '[x]', { bold: true });
      put(ws, 'D' + r, model.kpi.minCumDSCR, FMT_X, { bold: true }); r++;
      label(ws, r, '최종 기말현금(음수면 오류)', '[KRWm]', { bold: true });
      put(ws, 'D' + r, rows[N - 1].cashClose, FMT_M, { bold: true });
    })();

    /* =========================================================
       7. Report — 요약
       ========================================================= */
    (function () {
      var ws = wb.addWorksheet('Report', { properties: { tabColor: { argb: 'FF14483A' } } });
      ws.getColumn(1).width = 2.5; ws.getColumn(2).width = 30; ws.getColumn(3).width = 3; ws.getColumn(4).width = 16;
      title(ws, 'Executive Summary (v2 · 분기 · 5트랜치)');
      var r = 4;
      function kv(name, val, fmt) {
        ws.getCell('B' + r).value = name;
        ws.getCell('B' + r).font = { name: FONT, size: 10 };
        put(ws, 'D' + r, val, fmt, { bold: true });
        r++;
      }
      section(ws, r, '사업 개요'); r += 2;
      kv('사업명', inp.projectName, '@');
      kv('설비용량 [MW]', inp.capacityMW, '#,##0.000');
      kv('총 기간 수(분기)', N, '0');
      kv('건설 개시', inp.constructionStart, '@');
      r++;
      section(ws, r, '재원조달'); r += 2;
      kv('총투자비(TIC) [KRWm]', model.tic, FMT_M);
      kv('  건설이자(IDC) [KRWm]', model.idc, FMT_M);
      kv('자기자본 [KRWm]', model.equity, FMT_M);
      kv('차입금 합계 [KRWm]', model.debt, FMT_M);
      r++;
      section(ws, r, '수익성 지표'); r += 2;
      var k = model.kpi;
      kv('Project IRR 세전 [%]', k.projectIRRPre, FMT_P);
      kv('Project IRR 세후 [%]', k.projectIRR, FMT_P);
      kv('Equity IRR (FCFE) [%]', k.equityIRR, FMT_P);
      kv('Equity IRR (배당) [%]', k.dividendIRR, FMT_P);
      kv('Investor IRR [%]', k.investorIRR, FMT_P);
      kv('최소 단순DSCR(연 합산) [x]', k.minDSCRAnnual, FMT_X);
      kv('최소 누적DSCR [x]', k.minCumDSCR, FMT_X);
      kv('총영업수익(전체기간) [KRWm]', k.totalRevenue, FMT_M);
      kv('총영업비용(전체기간) [KRWm]', k.totalOpex, FMT_M);
      kv('총선순위이자 [KRWm]', k.totalInterest, FMT_M);
      kv('총법인세 [KRWm]', k.totalTax, FMT_M);
      kv('총배당(연차+청산) [KRWm]', k.totalDividend, FMT_M);
    })();

    /* =========================================================
       7-1. 민감도 — 화면에서 실행한 시나리오 비교 결과(있을 때만)
       ========================================================= */
    if (model.sensitivity && model.sensitivity.length) {
      (function () {
        var ws = wb.addWorksheet('민감도', { properties: { tabColor: { argb: 'FF14483A' } } });
        ws.getColumn(1).width = 2.5; ws.getColumn(2).width = 16;
        for (var ci = 0; ci < 9; ci++) ws.getColumn(3 + ci).width = 14;
        title(ws, '민감도 분석 — 시나리오별 핵심 지표 비교');
        var r = 4;
        ws.getCell('B' + r).value = '판매단가/총투자비/운영비는 %조정, 이자율은 bp(1/100%) 조정. 화면(생성기)에서 지정한 시나리오를 각각 독립적으로 재계산한 값 — 라이브 수식이 아니라 스냅샷임.';
        ws.getCell('B' + r).font = { name: FONT, size: 9, italic: true, color: { argb: 'FF6B7B76' } };
        r += 2;
        var heads = ['시나리오', '판매단가Δ[%]', '총투자비Δ[%]', '운영비Δ[%]', '이자율Δ[bp]',
          'Equity IRR(배당)', 'Equity IRR(FCFE)', 'Project IRR', '최소DSCR', 'NPV[억원]', '투자배수[x]'];
        heads.forEach(function (h, idx) {
          var c = ws.getCell(colLetter(2 + idx) + r);
          c.value = h;
          c.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
          c.alignment = { horizontal: 'center' };
        });
        r++;
        model.sensitivity.forEach(function (row) {
          put(ws, 'B' + r, row.name, '@');
          put(ws, 'C' + r, (row.sc && row.sc.tariffPct) || 0, '0.0');
          put(ws, 'D' + r, (row.sc && row.sc.capexPct) || 0, '0.0');
          put(ws, 'E' + r, (row.sc && row.sc.opexPct) || 0, '0.0');
          put(ws, 'F' + r, (row.sc && row.sc.ratebp) || 0, '0');
          if (row.error) {
            ws.getCell('G' + r).value = '계산 실패: ' + row.error;
            ws.getCell('G' + r).font = { name: FONT, size: 9, color: { argb: 'FFB4483E' } };
          } else {
            var k = row.kpi;
            put(ws, 'G' + r, k.dividendIRR, FMT_P);
            put(ws, 'H' + r, k.equityIRR, FMT_P);
            put(ws, 'I' + r, k.projectIRR, FMT_P);
            put(ws, 'J' + r, k.minDSCRAnnual, FMT_X);
            put(ws, 'K' + r, k.npv / 100, '#,##0');
            put(ws, 'L' + r, k.equityMultiple, '0.00');
          }
          r++;
        });
      })();
    }

    /* =========================================================
       8. 목차
       ========================================================= */
    (function () {
      var ws = wb.addWorksheet('목차', { properties: { tabColor: { argb: 'FF0B2F24' } } });
      ws.getColumn(1).width = 2.5; ws.getColumn(2).width = 14; ws.getColumn(3).width = 60;
      ws.getCell('B2').value = inp.projectName || '태양광 재무모델';
      ws.getCell('B2').font = { name: FONT, bold: true, size: 16, color: { argb: 'FF14483A' } };
      ws.getCell('B3').value = '재무모델 v2 (분기·5트랜치) — 재생E AI Agent 자동생성 · ' + new Date().toISOString().slice(0, 10);
      ws.getCell('B3').font = { name: FONT, size: 10, color: { argb: 'FF6B7B76' } };
      var list = [
        ['Report', 'Executive Summary — 핵심 KPI'],
        ['Funding', '자금조달 — 자본금 + 5트랜치 조건'],
        ['Debt_*', '트랜치별 상환 스케줄 (A/B/C/D/후순위) + 합계'],
        ['Revenue', '발전량 · 매출'],
        ['Opex', '운영비 (선순위/후순위)'],
        ['IS(Q)', '분기 손익계산서'],
        ['CF(Q)', '분기 현금흐름 · DSCR · 배당']
      ].concat(model.sensitivity && model.sensitivity.length ? [['민감도', '시나리오별 핵심 지표 비교']] : []);
      var r = 6;
      ['시트', '내용'].forEach(function (h, idx) {
        var c = ws.getCell(colLetter(2 + idx) + r);
        c.value = h;
        c.font = { name: FONT, bold: true, size: 10, color: { argb: WHITE } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HDR_FILL } };
      });
      r++;
      list.forEach(function (row) {
        ws.getCell('B' + r).value = row[0];
        ws.getCell('B' + r).font = { name: FONT, size: 10, bold: true, color: { argb: 'FF14483A' } };
        ws.getCell('C' + r).value = row[1];
        ws.getCell('C' + r).font = { name: FONT, size: 10 };
        r++;
      });
      r += 1;
      ws.getCell('B' + r).value = '※ 이 워크북은 값 기준(baked)입니다 — 가정을 바꾸려면 생성기에서 다시 뽑아야 합니다.';
      ws.getCell('B' + r).font = { name: FONT, size: 9, italic: true, color: { argb: 'FFB4573C' } };
    })();

    var order = ['목차', 'Report', 'Funding'].concat(
      model.tranches.map(function (t) { return 'Debt_' + t.name.replace(/\s/g, ''); }),
      ['Debt_합계', 'Revenue', 'Opex', 'IS(Q)', 'CF(Q)'],
      model.sensitivity && model.sensitivity.length ? ['민감도'] : []
    );
    order.forEach(function (n, idx) {
      var w = wb.getWorksheet(n);
      if (w) w.orderNo = idx + 1;
    });
    return wb;
  }

  var API = { buildWorkbook: buildWorkbook, colLetter: colLetter };
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  global.SolarXlsx2 = API;
})(typeof window !== 'undefined' ? window : globalThis);
