/* ============================================================
   재무모델 Excel 빌더 (ExcelJS)
   - 시트: 목차 / Assum / TIC_Funding / Debt / Revenue / Opex /
           IS(Y) / CF(Y) / DSCR / IRR / Report / Sensitivity
   - 열 규칙: B=구분, C=unit, D=합계, E=t0(건설), F~=운영 1~N년차
   - 가정(Assum)을 바꾸면 전 시트가 재계산되는 live 모델
   ============================================================ */
(function (global) {
  'use strict';

  var FONT = '맑은 고딕';
  var FMT_M = '#,##0;(#,##0);"-"';
  var FMT_M1 = '#,##0.0;(#,##0.0);"-"';
  var FMT_P = '0.0%';
  var FMT_P2 = '0.00%';
  var FMT_X = '0.00"x"';
  var FMT_N = '#,##0';
  var BLUE = 'FF0000FF', BLACK = 'FF000000', GREEN = 'FF008000', WHITE = 'FFFFFFFF';
  var HDR_FILL = 'FF14483A', SUB_FILL = 'FFE8F1ED', KEY_FILL = 'FFFFF2CC';

  function colLetter(n) { // 1 -> A
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }

  function buildWorkbook(model, ExcelJSLib) {
    var ExcelJS = ExcelJSLib || global.ExcelJS;
    var wb = new ExcelJS.Workbook();
    wb.creator = '재생E AI Agent — 재무모델 생성기';
    wb.created = new Date();

    var i = model.input, rows = model.rows, N = rows.length;
    var C0 = 5;                     // E열 = t0(건설기간)
    var yc = function (y) { return colLetter(C0 + y); };   // 운영 y년차 열문자 (F=1)
    var firstY = yc(1), lastY = yc(N);

    /* ---------- 공통 헬퍼 ---------- */
    function sheet(name, tabColor) {
      var ws = wb.addWorksheet(name, {
        views: [{ state: 'frozen', xSplit: 4, ySplit: 7 }],
        properties: { tabColor: { argb: tabColor || 'FF2E7D62' }, defaultRowHeight: 16 }
      });
      ws.getColumn(1).width = 2.5;
      ws.getColumn(2).width = 30;
      ws.getColumn(3).width = 11;
      ws.getColumn(4).width = 14;
      for (var c = 5; c <= C0 + N + 1; c++) ws.getColumn(c).width = 12;
      return ws;
    }

    function title(ws, name) {
      ws.getCell('A1').value = i.projectName;
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
      var labels = ['구  분', 'unit', '합 계', '건설기간'];
      ['B', 'C', 'D', 'E'].forEach(function (col, idx) {
        var c = ws.getCell(col + r);
        c.value = labels[idx];
        c.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
        c.alignment = { horizontal: 'center' };
        c.border = { bottom: { style: 'thin' } };
      });
      for (var y = 1; y <= N; y++) {
        var c1 = ws.getCell(yc(y) + r);
        c1.value = y + '년차';
        c1.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
        c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
        c1.alignment = { horizontal: 'center' };
        var c2 = ws.getCell(yc(y) + (r + 1));
        c2.value = rows[y - 1].year;
        c2.numFmt = '0';
        c2.font = { name: FONT, size: 9, color: { argb: 'FF6B7B76' } };
        c2.alignment = { horizontal: 'center' };
      }
      ws.getCell('E' + (r + 1)).value = model.cod.y - 1;
      ws.getCell('E' + (r + 1)).font = { name: FONT, size: 9, color: { argb: 'FF6B7B76' } };
      ws.getCell('E' + (r + 1)).alignment = { horizontal: 'center' };
    }

    // 라벨 + 단위
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

    // 값 셀
    function put(ws, addr, value, fmt, opt) {
      opt = opt || {};
      var c = ws.getCell(addr);
      if (value && typeof value === 'object' && value.formula !== undefined) c.value = value;
      else c.value = value;
      c.numFmt = fmt || FMT_M;
      c.font = { name: FONT, size: 10, bold: !!opt.bold, color: { argb: opt.color || BLACK } };
      if (opt.fill) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opt.fill } };
      if (opt.border) c.border = { top: { style: 'thin' }, bottom: { style: opt.border } };
      return c;
    }

    // 연도별 행 채우기: fn(y) -> {formula, result} 또는 숫자
    function fillYears(ws, r, fn, fmt, opt) {
      for (var y = 1; y <= N; y++) put(ws, yc(y) + r, fn(y), fmt, opt);
    }
    // 합계(D열)
    function sumRow(ws, r, fmt) {
      put(ws, 'D' + r, { formula: 'SUM(' + firstY + r + ':' + lastY + r + ')' }, fmt || FMT_M, { bold: true });
    }

    /* =========================================================
       1. Assum — 모든 입력 가정 (파란색 = 직접 입력)
       ========================================================= */
    var A = {};
    var CONS_IDC = "'Construction'!$D$13";
    (function () {
      var ws = wb.addWorksheet('Assum', { properties: { tabColor: { argb: 'FF14483A' } } });
      ws.getColumn(1).width = 2.5; ws.getColumn(2).width = 34; ws.getColumn(3).width = 4;
      ws.getColumn(4).width = 13; ws.getColumn(5).width = 3; ws.getColumn(6).width = 16;
      ws.getColumn(7).width = 42;
      title(ws, 'Assumption — 파란색 셀만 수정하면 전 시트가 재계산됩니다');

      var r = 4;
      function head(t) { section(ws, r, t); r += 2; }
      function inp(name, unit, value, fmt, key, note, isFormula) {
        ws.getCell('B' + r).value = name;
        ws.getCell('B' + r).font = { name: FONT, size: 10 };
        ws.getCell('D' + r).value = unit;
        ws.getCell('D' + r).font = { name: FONT, size: 9, color: { argb: 'FF9AA6A1' } };
        var c = ws.getCell('F' + r);
        c.value = value;
        c.numFmt = fmt;
        c.font = { name: FONT, size: 10, bold: true, color: { argb: isFormula ? BLACK : BLUE } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isFormula ? 'FFF2F5F4' : KEY_FILL } };
        c.border = { top: { style: 'hair' }, left: { style: 'hair' }, bottom: { style: 'hair' }, right: { style: 'hair' } };
        if (note) {
          ws.getCell('G' + r).value = note;
          ws.getCell('G' + r).font = { name: FONT, size: 9, color: { argb: 'FF9AA6A1' } };
        }
        if (key) A[key] = 'Assum!$F$' + r;
        r++;
      }
      function gap() { r++; }
      function L(key) { return A[key].replace('Assum!', ''); }

      // 섹션 설명
      ws.getCell('B3').value = '노랑 배경 + 파란 글씨 = 직접 입력값 / 회색 = 수식';
      ws.getCell('B3').font = { name: FONT, size: 9, italic: true, color: { argb: 'FF9AA6A1' } };

      head('General');
      inp('사업명', '[text]', i.projectName, '@', 'name');
      inp('설비용량', '[MW]', i.capacityMW, '#,##0.000', 'mw');
      inp('이용률', '[%]', i.capacityFactor, FMT_P2, 'cf', '일일발전시간 ÷ 24');
      inp('일일 발전시간(환산)', '[h/day]', { formula: L('cf') + '*24', result: i.capacityFactor * 24 }, '0.00', 'hday', null, true);
      inp('Degradation', '[%/yr]', i.degradation, FMT_P2, 'deg');
      inp('소내 소비율', '[%]', i.auxRate, FMT_P2, 'aux');
      gap();

      head('Timing');
      inp('착공시점', '[Date]', i.constructionStart, '@', 'cstart');
      inp('공사기간', '[Month]', i.constructionMonths, '0', 'cmonths');
      inp('준공(COD)', '[Date]', model.cod.label, '@', 'cod', null, true);
      inp('운영기간', '[Year]', i.operationYears, '0', 'oplife');
      gap();

      head('TIC & Funding');
      inp('총사업비(건설이자 제외)', '[KRWm]', model.capex0, FMT_M, 'capex0');
      inp('최초 DSRA', '[KRWm]', model.dsra0, FMT_M, 'dsra0');
      inp('차입비율(Gearing)', '[%]', i.gearing, FMT_P, 'gear', '기준: 총사업비 + 최초DSRA (건설이자는 전액 차입 조달)');
      inp('자기자본', '[KRWm]', { formula: '(' + L('capex0') + '+' + L('dsra0') + ')*(1-' + L('gear') + ')', result: model.equity }, FMT_M, 'equity', null, true);
      inp('건설이자(IDC)', '[KRWm]', { formula: CONS_IDC, result: model.idc }, FMT_M, 'idc', 'Construction 시트에서 월단위 산출', true);
      inp('총투자비(TIC)', '[KRWm]', { formula: L('capex0') + '+' + L('idc'), result: model.tic }, FMT_M, 'tic', null, true);
      inp('소요자금 합계', '[KRWm]', { formula: L('tic') + '+' + L('dsra0'), result: model.totalFunding }, FMT_M, 'need', null, true);
      inp('차입금', '[KRWm]', { formula: L('need') + '-' + L('equity'), result: model.debt }, FMT_M, 'debt', null, true);
      inp('실질 차입비율(TIC 기준)', '[%]', { formula: L('debt') + '/' + L('need'), result: model.actualGearing }, FMT_P, 'agear', null, true);
      gap();

      head('Debt');
      inp('이자율(건설기간)', '[%]', i.rateC, FMT_P2, 'rateC');
      inp('이자율(운영기간)', '[%]', i.rateO, FMT_P2, 'rateO');
      inp('거치기간', '[Year]', i.graceYears, '0', 'grace');
      inp('상환기간', '[Year]', i.repayYears, '0', 'repay');
      inp('상환방식', '[1/2/3]', i.repayType, '0', 'rtype', '1: 원금균등  2: 원리금균등  3: DSCR 스컬프팅(수동 스케줄)');
      inp('연간 상환횟수', '[회/년]', i.payPerYear, '0', 'ppy', '이자는 연중 평균잔액 기준으로 계산');
      inp('스컬프팅 균등 DSCR(산출)', '[x]', model.kpi.sculptDSCR, FMT_X, 'sculpt', '상환방식 3 선택 시 완전상환되는 균등 DSCR', true);
      gap();

      head('Revenue');
      inp('판매단가', '[KRW/kWh]', i.tariff, '#,##0.0', 'tariff', 'PPA 또는 SMP+REC 가중평균');
      inp('단가 상승률', '[%/yr]', i.tariffEscal, FMT_P2, 'tesc');
      gap();

      head('Cost');
      inp('운영비(연간, 1년차)', '[KRWm/yr]', i.opexEok * 100, FMT_M, 'opex');
      inp('운영비 상승률', '[%/yr]', i.opexEscal, FMT_P2, 'oesc');
      inp('철거·복구비(만기)', '[KRWm]', model.decom, FMT_M, 'decom');
      gap();

      head('Accounting & Tax');
      inp('감가상각 대상비율', '[%]', i.depRatio, FMT_P, 'depratio', 'TIC 중 감가상각 대상 (토지선납·금융비용 등 제외)');
      inp('감가상각 대상자산', '[KRWm]', { formula: L('tic') + '*' + L('depratio'), result: model.depBase }, FMT_M, 'depbase', null, true);
      inp('내용연수', '[Year]', i.depYears, '0', 'deplife');
      inp('법인세 방식', '[1/2]', i.taxMode, '0', 'taxmode', '1: 누진구간(9/19/21%)  2: 단일세율');
      inp('단일 법인세율', '[%]', i.taxFlat, FMT_P, 'taxflat');
      inp('이월결손금 공제비율', '[%]', i.lossRate, FMT_P, 'lossrate');
      gap();

      head('Equity & Reserve');
      inp('DSRA 적립기준', '[Month]', i.dsraMonths, '0', 'dsramon', '차기 원리금의 X개월분 적립');
      inp('배당 후 최소보유현금', '[KRWm]', i.minCash, FMT_M, 'mincash');
      inp('배당제한 DSCR', '[x]', i.divDSCR, FMT_X, 'divdscr');
      inp('배당개시 연차(Lock-up)', '[Year]', i.divStartYear, '0', 'divstart');
      gap();

      head('Evaluation');
      inp('할인율(NPV·LCOE)', '[%]', i.discount, FMT_P2, 'disc');
    })();

    var Aq = {}; // 수식용 (Assum!$F$n)
    for (var k in A) Aq[k] = A[k];

    /* =========================================================
       1-b. Construction — 건설기간 월별 인출 및 건설이자 (live)
       ========================================================= */
    (function () {
      var MM = Math.max(model.cSched.length + 12, 24);   // 여유 열 확보
      var ws = wb.addWorksheet('Construction', {
        views: [{ state: 'frozen', xSplit: 4, ySplit: 7 }],
        properties: { tabColor: { argb: 'FF2E7D62' } }
      });
      ws.getColumn(1).width = 2.5; ws.getColumn(2).width = 30; ws.getColumn(3).width = 11; ws.getColumn(4).width = 14;
      for (var c = 5; c <= 4 + MM; c++) ws.getColumn(c).width = 11;
      title(ws, '건설기간 자금인출 및 건설이자(IDC)');
      section(ws, 4, '월별 인출 스케줄');

      var mc = function (m) { return colLetter(4 + m); };   // 1개월차 = E
      for (var m = 1; m <= MM; m++) {
        var c1 = ws.getCell(mc(m) + '6');
        c1.value = m; c1.numFmt = '0"M"';
        c1.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
        c1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
        c1.alignment = { horizontal: 'center' };
      }
      ['구  분', 'unit', '합 계'].forEach(function (t, idx) {
        var c = ws.getCell(colLetter(2 + idx) + '6');
        c.value = t;
        c.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
        c.alignment = { horizontal: 'center' };
      });

      function mFill(r, fn, fmt, opt) {
        for (var m2 = 1; m2 <= MM; m2++) put(ws, mc(m2) + r, fn(m2), fmt, opt);
        put(ws, 'D' + r, { formula: 'SUM(' + mc(1) + r + ':' + mc(MM) + r + ')' }, fmt, { bold: true });
      }
      var cs = model.cSched, L = cs.length;
      var g = function (m, k) { return m <= L ? cs[m - 1][k] : 0; };

      label(ws, 8, '사업비 인출', '[KRWm]');
      mFill(8, function (m) {
        return { formula: 'IF(' + m + '<=' + Aq.cmonths + ',' + Aq.capex0 + '/' + Aq.cmonths + ',0)', result: g(m, 'cost') };
      }, FMT_M);

      label(ws, 9, '자기자본 인출', '[KRWm]');
      mFill(9, function (m) {
        var prev = m === 1 ? '0' : 'SUM($' + mc(1) + '$9:' + mc(m - 1) + '9)';
        return { formula: 'MIN(MAX(0,' + Aq.equity + '-' + prev + '),' + mc(m) + '8)', result: g(m, 'eq') };
      }, FMT_M);

      label(ws, 10, '차입금 인출', '[KRWm]');
      mFill(10, function (m) {
        return { formula: mc(m) + '8-' + mc(m) + '9', result: g(m, 'dt') };
      }, FMT_M);

      label(ws, 11, '기초 차입잔액', '[KRWm]');
      for (var m3 = 1; m3 <= MM; m3++) {
        put(ws, mc(m3) + '11', m3 === 1 ? 0 : { formula: mc(m3 - 1) + '13', result: g(m3, 'open') }, FMT_M);
      }

      label(ws, 12, '건설이자(월)', '[KRWm]');
      mFill(12, function (m) {
        return { formula: 'IF(' + m + '<=' + Aq.cmonths + ',(' + mc(m) + '11+' + mc(m) + '10)*' + Aq.rateC + '/12,0)', result: g(m, 'interest') };
      }, FMT_M);

      label(ws, 13, '기말 차입잔액', '[KRWm]', { bold: true, fill: SUB_FILL });
      for (var m4 = 1; m4 <= MM; m4++) {
        put(ws, mc(m4) + '13', { formula: mc(m4) + '11+' + mc(m4) + '10+' + mc(m4) + '12', result: g(m4, 'close') }, FMT_M, { bold: true });
      }
      // D13 = 건설이자 합계(IDC)  ← Assum이 참조하는 셀
      put(ws, 'D13', { formula: 'SUM(' + mc(1) + '12:' + mc(MM) + '12)', result: model.idc }, FMT_M, { bold: true });
      ws.getCell('C13').value = '[IDC 합계]';
      ws.getCell('C13').font = { name: FONT, size: 9, bold: true, color: { argb: 'FF2E7D62' } };

      section(ws, 15, '준공시점 정산');
      label(ws, 17, '기말 차입잔액(사업비)', '[KRWm]');
      put(ws, 'D17', { formula: mc(MM) + '13', result: model.debt - model.dsraDt }, FMT_M);
      label(ws, 18, 'DSRA — 자기자본 조달', '[KRWm]');
      put(ws, 'D18', { formula: 'MIN(MAX(0,' + Aq.equity + '-SUM(' + mc(1) + '9:' + mc(MM) + '9)),' + Aq.dsra0 + ')', result: model.dsraEq }, FMT_M);
      label(ws, 19, 'DSRA — 차입 조달', '[KRWm]');
      put(ws, 'D19', { formula: Aq.dsra0 + '-D18', result: model.dsraDt }, FMT_M);
      label(ws, 20, '차입금 합계(운영개시)', '[KRWm]', { bold: true, fill: SUB_FILL });
      put(ws, 'D20', { formula: 'D17+D19', result: model.debt }, FMT_M, { bold: true });
      label(ws, 21, 'Check (Assum 차입금과 일치)', '[KRWm]');
      put(ws, 'D21', { formula: 'D20-' + Aq.debt, result: 0 }, FMT_M);

      ws.getCell('B23').value = '※ 자기자본을 먼저 인출하고 잔여 소요액을 차입하는 구조이며, 건설이자는 매월 자본화(원금가산)됩니다.';
      ws.getCell('B23').font = { name: FONT, size: 9, italic: true, color: { argb: 'FF9AA6A1' } };
      ws.getCell('B24').value = '※ 공사기간을 ' + MM + '개월 이상으로 늘리려면 열을 추가하거나 생성기에서 다시 만들어 주세요.';
      ws.getCell('B24').font = { name: FONT, size: 9, italic: true, color: { argb: 'FF9AA6A1' } };
    })();

    /* =========================================================
       2. TIC_Funding
       ========================================================= */
    (function () {
      var ws = sheet('TIC_Funding');
      title(ws, '총투자비 및 재원조달');
      section(ws, 4, '총투자비');
      var r = 6;
      ws.getCell('B' + r).value = '구  분'; ws.getCell('D' + r).value = '금액'; ws.getCell('E' + r).value = '비중';
      ['B', 'D', 'E'].forEach(function (c) {
        var cc = ws.getCell(c + r);
        cc.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
        cc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
        cc.alignment = { horizontal: 'center' };
      });
      r = 7;
      label(ws, r, '총사업비(건설이자 제외)', '[KRWm]');
      put(ws, 'D' + r, { formula: Aq.capex0, result: model.capex0 }, FMT_M, { color: GREEN });
      put(ws, 'E' + r, { formula: 'D7/$D$10', result: model.capex0 / model.totalFunding }, FMT_P);
      r = 8; label(ws, r, '건설이자(IDC)', '[KRWm]');
      put(ws, 'D' + r, { formula: Aq.idc, result: model.idc }, FMT_M, { color: GREEN });
      put(ws, 'E' + r, { formula: 'D8/$D$10', result: model.idc / model.totalFunding }, FMT_P);
      r = 9; label(ws, r, '최초 DSRA', '[KRWm]');
      put(ws, 'D' + r, { formula: Aq.dsra0, result: model.dsra0 }, FMT_M, { color: GREEN });
      put(ws, 'E' + r, { formula: 'D9/$D$10', result: model.dsra0 / model.totalFunding }, FMT_P);
      r = 10; label(ws, r, '소요자금 합계', '[KRWm]', { bold: true, fill: SUB_FILL });
      put(ws, 'D' + r, { formula: 'SUM(D7:D9)', result: model.totalFunding }, FMT_M, { bold: true });
      put(ws, 'E' + r, { formula: 'D10/$D$10', result: 1 }, FMT_P, { bold: true });

      section(ws, 13, '재원조달');
      r = 15;
      label(ws, r, '자기자본', '[KRWm]');
      put(ws, 'D' + r, { formula: Aq.equity, result: model.equity }, FMT_M, { color: GREEN });
      put(ws, 'E' + r, { formula: 'D15/$D$17', result: model.equity / model.totalFunding }, FMT_P);
      r = 16; label(ws, r, '선순위 차입금', '[KRWm]');
      put(ws, 'D' + r, { formula: Aq.debt, result: model.debt }, FMT_M, { color: GREEN });
      put(ws, 'E' + r, { formula: 'D16/$D$17', result: model.debt / model.totalFunding }, FMT_P);
      r = 17; label(ws, r, '합계', '[KRWm]', { bold: true, fill: SUB_FILL });
      put(ws, 'D' + r, { formula: 'SUM(D15:D16)', result: model.totalFunding }, FMT_M, { bold: true });
      put(ws, 'E' + r, { formula: 'D17/$D$17', result: 1 }, FMT_P, { bold: true });
      r = 18; label(ws, r, 'Check (소요 - 조달)', '[KRWm]');
      put(ws, 'D' + r, { formula: 'D10-D17', result: 0 }, FMT_M);

      section(ws, 21, '감가상각');
      r = 23; label(ws, r, '감가상각 대상자산', '[KRWm]');
      put(ws, 'D' + r, { formula: Aq.depbase, result: model.depBase }, FMT_M, { color: GREEN });
      r = 24; label(ws, r, '내용연수', '[Year]');
      put(ws, 'D' + r, { formula: Aq.deplife, result: i.depYears }, '0', { color: GREEN });
      r = 25; label(ws, r, '연간 상각비', '[KRWm/yr]', { bold: true });
      put(ws, 'D' + r, { formula: 'D23/D24', result: model.depAnnual }, FMT_M, { bold: true });
      ws.getCell('B27').value = '※ 건설이자는 자기자본 우선인출 가정 하에 월단위로 자본화하여 산출한 값입니다.';
      ws.getCell('B27').font = { name: FONT, size: 9, italic: true, color: { argb: 'FF9AA6A1' } };
    })();
    var TIC = { dep: "'TIC_Funding'!$D$25", debt: "'TIC_Funding'!$D$16", equity: "'TIC_Funding'!$D$15", tic: "'TIC_Funding'!$D$10" };

    /* =========================================================
       3. Revenue
       ========================================================= */
    var REV = {};
    (function () {
      var ws = sheet('Revenue');
      title(ws, '영업수익 추정');
      section(ws, 4, '발전량 및 매출');
      periodHeader(ws, 6);
      var r = 9;
      label(ws, r, '설비용량', '[MW]');
      fillYears(ws, r, function () { return { formula: Aq.mw, result: i.capacityMW }; }, '#,##0.000');
      r = 10; label(ws, r, '이용률', '[%]');
      fillYears(ws, r, function () { return { formula: Aq.cf, result: i.capacityFactor }; }, FMT_P2);
      r = 11; label(ws, r, 'Degradation 계수', '[%]');
      fillYears(ws, r, function (y) {
        return { formula: '(1-' + Aq.deg + ')^(' + y + '-1)', result: Math.pow(1 - i.degradation, y - 1) };
      }, '0.000');
      r = 12; label(ws, r, '발전량(송전단)', '[MWh]', { bold: true });
      fillYears(ws, r, function (y) {
        return {
          formula: yc(y) + '9*8760*' + yc(y) + '10*(1-' + Aq.aux + ')*' + yc(y) + '11',
          result: rows[y - 1].gen
        };
      }, FMT_N, { bold: true });
      sumRow(ws, r, FMT_N);
      r = 14; label(ws, r, '판매단가', '[KRW/kWh]');
      fillYears(ws, r, function (y) {
        return { formula: Aq.tariff + '*(1+' + Aq.tesc + ')^(' + y + '-1)', result: rows[y - 1].price };
      }, '#,##0.0');
      r = 16; label(ws, r, '영업수익', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillYears(ws, r, function (y) {
        return { formula: yc(y) + '12*' + yc(y) + '14/1000', result: rows[y - 1].revenue };
      }, FMT_M, { bold: true });
      sumRow(ws, r);
      REV.gen = 12; REV.rev = 16;
    })();

    /* =========================================================
       4. Opex
       ========================================================= */
    (function () {
      var ws = sheet('Opex');
      title(ws, '영업비용 추정');
      section(ws, 4, '운영비용');
      periodHeader(ws, 6);
      var r = 9;
      label(ws, r, '운영비(1년차 기준)', '[KRWm]');
      fillYears(ws, r, function () { return { formula: Aq.opex, result: i.opexEok * 100 }; }, FMT_M);
      r = 10; label(ws, r, '물가상승 계수', '[%]');
      fillYears(ws, r, function (y) {
        return { formula: '(1+' + Aq.oesc + ')^(' + y + '-1)', result: Math.pow(1 + i.opexEscal, y - 1) };
      }, '0.000');
      r = 12; label(ws, r, '영업비용 합계', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillYears(ws, r, function (y) {
        return { formula: yc(y) + '9*' + yc(y) + '10', result: rows[y - 1].opex };
      }, FMT_M, { bold: true });
      sumRow(ws, r);
    })();

    /* =========================================================
       5. Debt — 상환스케줄
       ========================================================= */
    var DBT = { open: 9, prin: 10, close: 11, int: 12, ds: 13, dsra: 16, dsraMove: 17 };
    (function () {
      var ws = sheet('Debt');
      title(ws, '차입금 인출·상환 스케줄');
      section(ws, 4, '선순위 차입금');
      periodHeader(ws, 6);
      var sculpt = (i.repayType === 3);

      label(ws, 9, '기초잔액', '[KRWm]');
      fillYears(ws, 9, function (y) {
        return y === 1 ? { formula: TIC.debt, result: rows[0].debtOpen }
          : { formula: yc(y - 1) + '11', result: rows[y - 1].debtOpen };
      }, FMT_M);

      label(ws, 10, '원금상환', '[KRWm]', { bold: sculpt });
      fillYears(ws, 10, function (y) {
        var v = rows[y - 1].principal;
        if (sculpt) return v;   // 스컬프팅 = 산출된 스케줄을 key-in (원본 모델의 '원금불균등' 방식)
        var inWin = 'AND(' + y + '>' + Aq.grace + ',' + y + '<=' + Aq.grace + '+' + Aq.repay + ')';
        var eq = TIC.debt + '/' + Aq.repay;
        var an = '-PPMT(' + Aq.rateO + ',' + y + '-' + Aq.grace + ',' + Aq.repay + ',' + TIC.debt + ')';
        return { formula: 'IF(' + inWin + ',MIN(' + yc(y) + '9,IF(' + Aq.rtype + '=2,' + an + ',' + eq + ')),0)', result: v };
      }, FMT_M, sculpt ? { color: BLUE, fill: KEY_FILL } : {});
      sumRow(ws, 10);

      label(ws, 11, '기말잔액', '[KRWm]');
      fillYears(ws, 11, function (y) {
        return { formula: yc(y) + '9-' + yc(y) + '10', result: rows[y - 1].debtClose };
      }, FMT_M);

      label(ws, 12, '이자비용', '[KRWm]');
      fillYears(ws, 12, function (y) {
        return {
          formula: '(' + yc(y) + '9-' + yc(y) + '10*(' + Aq.ppy + '-1)/(2*' + Aq.ppy + '))*' + Aq.rateO,
          result: rows[y - 1].interest
        };
      }, FMT_M);
      sumRow(ws, 12);

      label(ws, 13, '원리금 합계(DS)', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillYears(ws, 13, function (y) {
        return { formula: yc(y) + '10+' + yc(y) + '12', result: rows[y - 1].ds };
      }, FMT_M, { bold: true });
      sumRow(ws, 13);

      section(ws, 15, 'DSRA (원리금상환적립금)');
      label(ws, 16, 'DSRA 잔액(기말)', '[KRWm]');
      fillYears(ws, 16, function (y) {
        var nxt = (y < N) ? yc(y + 1) + '13' : '0';
        return { formula: 'IF(' + y + '>=' + Aq.grace + '+' + Aq.repay + ',0,' + nxt + '*' + Aq.dsramon + '/12)', result: rows[y - 1].dsraClose };
      }, FMT_M);
      label(ws, 17, 'DSRA 증감((+)적립/(-)환입)', '[KRWm]');
      fillYears(ws, 17, function (y) {
        var prev = (y === 1) ? Aq.dsra0 : yc(y - 1) + '16';
        return { formula: yc(y) + '16-' + prev, result: rows[y - 1].dsraMove };
      }, FMT_M);
      sumRow(ws, 17);

      label(ws, 19, '상환 종료 후 미상환 잔액', '[KRWm]', { bold: true });
      put(ws, 'D19', { formula: yc(N) + '11', result: rows[N - 1].debtClose }, FMT_M, { bold: true });
      put(ws, 'F19', {
        formula: 'IF(ABS(D19)<1,"완전상환 확인 (OK)","경고: 미상환 잔액 발생 — 원금상환 행 재산출 필요")',
        result: Math.abs(rows[N - 1].debtClose) < 1 ? '완전상환 확인 (OK)' : '경고: 미상환 잔액 발생 — 원금상환 행 재산출 필요'
      }, null, { bold: true });

      if (sculpt) {
        ws.getCell('B21').value = '※ 상환방식 3(DSCR 스컬프팅): 위 원금상환 행(파란 셀)은 균등 DSCR ' +
          model.kpi.sculptDSCR.toFixed(3) + 'x 로 완전상환되도록 산출해 넣은 스케줄입니다.';
        ws.getCell('B22').value = '   스컬프팅은 반복계산이 필요해 수식으로 자동 재산출되지 않습니다. ' +
          'Assum의 이용률·금리·단가·상환기간 등을 엑셀에서 바꾸면 이 행은 그대로 남아 D19에 미상환 잔액이 뜹니다.';
        ws.getCell('B23').value = '   → 가정을 바꾼 뒤에는 생성기에서 다시 추출하거나, 상환방식 1(원금균등)·2(원리금균등)를 쓰면 엑셀에서 완전히 자동 재계산됩니다.';
        [21, 22, 23].forEach(function (r) {
          ws.getCell('B' + r).font = { name: FONT, size: 9, italic: true, color: { argb: 'FFB4573C' } };
        });
      }
    })();

    /* =========================================================
       6. IS(Y) — 손익계산서
       ========================================================= */
    var ISR = { rev: 9, opex: 10, ebitda: 11, dep: 12, ebit: 13, int: 14, ebt: 15, lossOpen: 16, deduct: 17, base: 18, tax: 19, lossEnd: 20, ni: 21 };
    (function () {
      var ws = sheet('IS(Y)');
      title(ws, '추정 손익계산서 (연간)');
      section(ws, 4, '손익계산서');
      periodHeader(ws, 6);

      label(ws, 9, '영업수익', '[KRWm]');
      fillYears(ws, 9, function (y) { return { formula: "'Revenue'!" + yc(y) + REV.rev, result: rows[y - 1].revenue }; }, FMT_M, { color: GREEN });
      sumRow(ws, 9);
      label(ws, 10, '영업비용', '[KRWm]');
      fillYears(ws, 10, function (y) { return { formula: "-'Opex'!" + yc(y) + '12', result: -rows[y - 1].opex }; }, FMT_M, { color: GREEN });
      sumRow(ws, 10);
      label(ws, 11, 'EBITDA', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillYears(ws, 11, function (y) { return { formula: yc(y) + '9+' + yc(y) + '10', result: rows[y - 1].ebitda }; }, FMT_M, { bold: true });
      sumRow(ws, 11);
      label(ws, 12, '감가상각비', '[KRWm]');
      fillYears(ws, 12, function (y) {
        return { formula: 'IF(' + y + '<=' + Aq.deplife + ',-' + TIC.dep + ',0)', result: -rows[y - 1].dep };
      }, FMT_M);
      sumRow(ws, 12);
      label(ws, 13, '영업이익(EBIT)', '[KRWm]', { bold: true });
      fillYears(ws, 13, function (y) { return { formula: yc(y) + '11+' + yc(y) + '12', result: rows[y - 1].ebit }; }, FMT_M, { bold: true });
      sumRow(ws, 13);
      label(ws, 14, '이자비용', '[KRWm]');
      fillYears(ws, 14, function (y) { return { formula: "-'Debt'!" + yc(y) + '12', result: -rows[y - 1].interest }; }, FMT_M, { color: GREEN });
      sumRow(ws, 14);
      label(ws, 15, '법인세차감전순이익', '[KRWm]', { bold: true });
      fillYears(ws, 15, function (y) { return { formula: yc(y) + '13+' + yc(y) + '14', result: rows[y - 1].ebt }; }, FMT_M, { bold: true });
      sumRow(ws, 15);

      label(ws, 16, '이월결손금(기초)', '[KRWm]');
      fillYears(ws, 16, function (y) {
        return y === 1 ? 0 : { formula: yc(y - 1) + '20', result: rows[y - 2].lossCF };
      }, FMT_M);
      label(ws, 17, '결손금 공제액', '[KRWm]');
      fillYears(ws, 17, function (y) {
        return { formula: 'MIN(MAX(0,' + yc(y) + '15)*' + Aq.lossrate + ',' + yc(y) + '16)', result: rows[y - 1].deduct };
      }, FMT_M);
      label(ws, 18, '과세표준', '[KRWm]');
      fillYears(ws, 18, function (y) {
        return { formula: 'MAX(0,' + yc(y) + '15-' + yc(y) + '17)', result: rows[y - 1].taxBase };
      }, FMT_M);
      label(ws, 19, '법인세비용', '[KRWm]');
      fillYears(ws, 19, function (y) {
        var b = yc(y) + '18';
        var bracket = 'IF(' + b + '<=200,' + b + '*0.09,IF(' + b + '<=20000,18+(' + b + '-200)*0.19,3780+(' + b + '-20000)*0.21))';
        return { formula: 'IF(' + Aq.taxmode + '=1,' + bracket + ',' + b + '*' + Aq.taxflat + ')', result: rows[y - 1].tax };
      }, FMT_M);
      sumRow(ws, 19);
      label(ws, 20, '이월결손금(기말)', '[KRWm]');
      fillYears(ws, 20, function (y) {
        return { formula: yc(y) + '16-' + yc(y) + '17+MAX(0,-' + yc(y) + '15)', result: rows[y - 1].lossCF };
      }, FMT_M);
      label(ws, 21, '당기순이익', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillYears(ws, 21, function (y) { return { formula: yc(y) + '15-' + yc(y) + '19', result: rows[y - 1].ni }; }, FMT_M, { bold: true });
      sumRow(ws, 21);
      label(ws, 23, 'EBITDA Margin', '[%]');
      fillYears(ws, 23, function (y) {
        return { formula: 'IFERROR(' + yc(y) + '11/' + yc(y) + '9,0)', result: rows[y - 1].ebitdaMargin };
      }, FMT_P);
    })();

    /* =========================================================
       7. CF(Y) — 현금흐름
       ========================================================= */
    var CFR = { cfads: 11, ds: 13, fcfe: 17, div: 19 };
    (function () {
      var ws = sheet('CF(Y)');
      title(ws, '추정 현금흐름표 (연간)');
      section(ws, 4, '영업활동 현금흐름');
      periodHeader(ws, 6);

      label(ws, 9, 'EBITDA', '[KRWm]');
      fillYears(ws, 9, function (y) { return { formula: "'IS(Y)'!" + yc(y) + '11', result: rows[y - 1].ebitda }; }, FMT_M, { color: GREEN });
      sumRow(ws, 9);
      label(ws, 10, '법인세 지급', '[KRWm]');
      fillYears(ws, 10, function (y) { return { formula: "-'IS(Y)'!" + yc(y) + '19', result: -rows[y - 1].tax }; }, FMT_M, { color: GREEN });
      sumRow(ws, 10);
      label(ws, 11, 'CFADS (상환가능현금)', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillYears(ws, 11, function (y) { return { formula: yc(y) + '9+' + yc(y) + '10', result: rows[y - 1].cfads }; }, FMT_M, { bold: true });
      sumRow(ws, 11);

      section(ws, 12, '재무활동 현금흐름');
      label(ws, 13, '원리금 상환', '[KRWm]');
      fillYears(ws, 13, function (y) { return { formula: "-'Debt'!" + yc(y) + '13', result: -rows[y - 1].ds }; }, FMT_M, { color: GREEN });
      sumRow(ws, 13);
      label(ws, 14, 'DSRA 증감', '[KRWm]');
      fillYears(ws, 14, function (y) { return { formula: "-'Debt'!" + yc(y) + '17', result: -rows[y - 1].dsraMove }; }, FMT_M, { color: GREEN });
      sumRow(ws, 14);
      label(ws, 15, '철거·복구비', '[KRWm]');
      fillYears(ws, 15, function (y) {
        return { formula: 'IF(' + y + '=' + Aq.oplife + ',-' + Aq.decom + ',0)', result: -rows[y - 1].decom };
      }, FMT_M);
      sumRow(ws, 15);
      label(ws, 17, '주주귀속 잉여현금(FCFE)', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillYears(ws, 17, function (y) {
        return { formula: yc(y) + '11+' + yc(y) + '13+' + yc(y) + '14+' + yc(y) + '15', result: rows[y - 1].fcfe };
      }, FMT_M, { bold: true });
      sumRow(ws, 17);

      section(ws, 18, '배당 및 현금계정');
      label(ws, 19, '배당금 지급', '[KRWm]', { bold: true });
      fillYears(ws, 19, function (y) {
        var cashBefore = yc(y) + '21+' + yc(y) + '17';
        var reserve = 'IF(' + y + '=' + Aq.oplife + ',0,' + Aq.mincash + ')';
        var cond = 'AND(' + y + '>=' + Aq.divstart + ",IF(ISNUMBER('DSCR'!" + yc(y) + "11),'DSCR'!" + yc(y) + '11,99)>=' + Aq.divdscr + ')';
        return { formula: 'IF(' + cond + ',MAX(0,' + cashBefore + '-' + reserve + '),0)', result: rows[y - 1].dividend };
      }, FMT_M, { bold: true });
      sumRow(ws, 19);
      label(ws, 21, '기초현금', '[KRWm]');
      fillYears(ws, 21, function (y) {
        return y === 1 ? 0 : { formula: yc(y - 1) + '22', result: rows[y - 1].cashOpen };
      }, FMT_M);
      label(ws, 22, '기말현금', '[KRWm]');
      fillYears(ws, 22, function (y) {
        return { formula: yc(y) + '21+' + yc(y) + '17-' + yc(y) + '19', result: rows[y - 1].cashClose };
      }, FMT_M);
      label(ws, 23, 'DSRA 잔액', '[KRWm]');
      fillYears(ws, 23, function (y) { return { formula: "'Debt'!" + yc(y) + '16', result: rows[y - 1].dsraClose }; }, FMT_M, { color: GREEN });
    })();

    /* =========================================================
       8. DSCR
       ========================================================= */
    (function () {
      var ws = sheet('DSCR');
      title(ws, '원리금 상환능력 분석');
      section(ws, 4, 'DSCR');
      periodHeader(ws, 6);
      label(ws, 9, 'CFADS', '[KRWm]');
      fillYears(ws, 9, function (y) { return { formula: "'CF(Y)'!" + yc(y) + '11', result: rows[y - 1].cfads }; }, FMT_M, { color: GREEN });
      label(ws, 10, '원리금(DS)', '[KRWm]');
      fillYears(ws, 10, function (y) { return { formula: "'Debt'!" + yc(y) + '13', result: rows[y - 1].ds }; }, FMT_M, { color: GREEN });
      label(ws, 11, '단순 DSCR', '[x]', { bold: true, fill: SUB_FILL });
      fillYears(ws, 11, function (y) {
        return { formula: 'IF(' + yc(y) + '10<=0.001,"-",' + yc(y) + '9/' + yc(y) + '10)', result: rows[y - 1].dscr === null ? '-' : rows[y - 1].dscr };
      }, FMT_X, { bold: true });
      label(ws, 13, '최소 DSCR', '[x]', { bold: true });
      put(ws, 'D13', { formula: 'MIN(' + firstY + '11:' + lastY + '11)', result: model.kpi.minDSCR }, FMT_X, { bold: true });
      label(ws, 14, '평균 DSCR', '[x]', { bold: true });
      put(ws, 'D14', { formula: 'IFERROR(AVERAGE(' + firstY + '11:' + lastY + '11),0)', result: model.kpi.avgDSCR }, FMT_X, { bold: true });
      ws.getCell('B16').value = '※ 상환이 없는 연도는 "-"로 표시되며 최소/평균 산출에서 제외됩니다.';
      ws.getCell('B16').font = { name: FONT, size: 9, italic: true, color: { argb: 'FF9AA6A1' } };
    })();

    /* =========================================================
       9. IRR
       ========================================================= */
    (function () {
      var ws = sheet('IRR');
      title(ws, '수익성 분석 (IRR / NPV)');
      section(ws, 4, 'Project IRR');
      periodHeader(ws, 6);

      label(ws, 9, '총투자비', '[KRWm]');
      put(ws, 'E9', { formula: '-' + TIC.tic + '+' + Aq.dsra0, result: -model.tic }, FMT_M);
      fillYears(ws, 9, function () { return 0; }, FMT_M);
      label(ws, 10, 'EBITDA', '[KRWm]');
      put(ws, 'E10', 0, FMT_M);
      fillYears(ws, 10, function (y) { return { formula: "'IS(Y)'!" + yc(y) + '11', result: rows[y - 1].ebitda }; }, FMT_M, { color: GREEN });
      label(ws, 11, '법인세', '[KRWm]');
      put(ws, 'E11', 0, FMT_M);
      fillYears(ws, 11, function (y) { return { formula: "-'IS(Y)'!" + yc(y) + '19', result: -rows[y - 1].tax }; }, FMT_M, { color: GREEN });
      label(ws, 12, '철거·복구비', '[KRWm]');
      put(ws, 'E12', 0, FMT_M);
      fillYears(ws, 12, function (y) { return { formula: "'CF(Y)'!" + yc(y) + '15', result: -rows[y - 1].decom }; }, FMT_M, { color: GREEN });
      label(ws, 13, 'Net Cashflow', '[KRWm]', { bold: true, fill: SUB_FILL });
      put(ws, 'E13', { formula: 'SUM(E9:E12)', result: -model.tic }, FMT_M, { bold: true });
      fillYears(ws, 13, function (y) {
        return { formula: 'SUM(' + yc(y) + '9:' + yc(y) + '12)', result: rows[y - 1].projectFcf };
      }, FMT_M, { bold: true });

      label(ws, 15, 'Project IRR (세후)', '[%]', { bold: true });
      put(ws, 'D15', { formula: 'IRR(E13:' + lastY + '13)', result: model.kpi.projectIRR }, FMT_P2, { bold: true });
      label(ws, 16, 'NPV @ 할인율', '[KRWm]', { bold: true });
      put(ws, 'D16', { formula: 'NPV(' + Aq.disc + ',' + firstY + '13:' + lastY + '13)+E13', result: model.kpi.npv }, FMT_M, { bold: true });

      section(ws, 18, 'Equity IRR');
      label(ws, 20, '자본금 납입', '[KRWm]');
      put(ws, 'E20', { formula: '-' + TIC.equity, result: -model.equity }, FMT_M);
      fillYears(ws, 20, function () { return 0; }, FMT_M);
      label(ws, 21, 'FCFE', '[KRWm]');
      put(ws, 'E21', 0, FMT_M);
      fillYears(ws, 21, function (y) { return { formula: "'CF(Y)'!" + yc(y) + '17', result: rows[y - 1].fcfe }; }, FMT_M, { color: GREEN });
      label(ws, 22, 'Net Cashflow (FCFE base)', '[KRWm]', { bold: true, fill: SUB_FILL });
      put(ws, 'E22', { formula: 'E20+E21', result: -model.equity }, FMT_M, { bold: true });
      fillYears(ws, 22, function (y) {
        return { formula: yc(y) + '20+' + yc(y) + '21', result: rows[y - 1].fcfe };
      }, FMT_M, { bold: true });
      label(ws, 24, '배당금', '[KRWm]');
      put(ws, 'E24', { formula: '-' + TIC.equity, result: -model.equity }, FMT_M);
      fillYears(ws, 24, function (y) { return { formula: "'CF(Y)'!" + yc(y) + '19', result: rows[y - 1].dividend }; }, FMT_M, { color: GREEN });

      label(ws, 26, 'Equity IRR (FCFE 기준)', '[%]', { bold: true });
      put(ws, 'D26', { formula: 'IRR(E22:' + lastY + '22)', result: model.kpi.equityIRR }, FMT_P2, { bold: true });
      label(ws, 27, 'Equity IRR (배당 기준)', '[%]', { bold: true });
      put(ws, 'D27', { formula: 'IRR(E24:' + lastY + '24)', result: model.kpi.dividendIRR }, FMT_P2, { bold: true });
    })();

    /* =========================================================
       10. Report — 요약
       ========================================================= */
    (function () {
      var ws = sheet('Report', 'FF14483A');
      title(ws, 'Executive Summary');
      section(ws, 4, '사업 개요');
      var r = 6;
      function kv(name, val, fmt, formula) {
        ws.getCell('B' + r).value = name;
        ws.getCell('B' + r).font = { name: FONT, size: 10 };
        var c = ws.getCell('D' + r);
        c.value = formula ? { formula: formula, result: val } : val;
        c.numFmt = fmt;
        c.font = { name: FONT, size: 10, bold: true, color: { argb: formula ? GREEN : BLACK } };
        r++;
      }
      kv('사업명', i.projectName, '@');
      kv('설비용량 [MW]', i.capacityMW, '#,##0.000', Aq.mw);
      kv('이용률 [%]', i.capacityFactor, FMT_P2, Aq.cf);
      kv('준공(COD)', model.cod.label, '@');
      kv('운영기간 [Year]', i.operationYears, '0', Aq.oplife);
      r++;
      section(ws, r, '재원조달'); r += 2;
      kv('총투자비(TIC) [KRWm]', model.tic, FMT_M, Aq.tic);
      kv('  건설이자(IDC) [KRWm]', model.idc, FMT_M, Aq.idc);
      kv('자기자본 [KRWm]', model.equity, FMT_M, TIC.equity);
      kv('선순위 차입금 [KRWm]', model.debt, FMT_M, TIC.debt);
      kv('차입비율 [%]', i.gearing, FMT_P, Aq.gear);
      r++;
      section(ws, r, '수익성 지표'); r += 2;
      kv('Project IRR (세후) [%]', model.kpi.projectIRR, FMT_P2, "'IRR'!$D$15");
      kv('Equity IRR (FCFE) [%]', model.kpi.equityIRR, FMT_P2, "'IRR'!$D$26");
      kv('Equity IRR (배당) [%]', model.kpi.dividendIRR, FMT_P2, "'IRR'!$D$27");
      kv('NPV [KRWm]', model.kpi.npv, FMT_M, "'IRR'!$D$16");
      kv('최소 DSCR [x]', model.kpi.minDSCR, FMT_X, "'DSCR'!$D$13");
      kv('평균 DSCR [x]', model.kpi.avgDSCR, FMT_X, "'DSCR'!$D$14");
      kv('LCOE [KRW/kWh]', model.kpi.lcoe, '#,##0.0');
      kv('투자비 회수기간 [Year]', model.kpi.payback, '0.0');
      kv('연평균 EBITDA [KRWm]', model.kpi.avgEbitda, FMT_M);
      kv('EBITDA Margin [%]', model.kpi.ebitdaMargin, FMT_P);

      var t = r + 1;
      section(ws, t, '연도별 요약');
      periodHeader(ws, t + 2);
      var base = t + 5;
      var defs = [
        ['영업수익', 'rev', "'Revenue'!" + '@16', FMT_M],
        ['영업비용', 'opex', "'Opex'!" + '@12', FMT_M],
        ['EBITDA', 'ebitda', "'IS(Y)'!" + '@11', FMT_M],
        ['당기순이익', 'ni', "'IS(Y)'!" + '@21', FMT_M],
        ['CFADS', 'cfads', "'CF(Y)'!" + '@11', FMT_M],
        ['원리금(DS)', 'ds', "'Debt'!" + '@13', FMT_M],
        ['DSCR', 'dscr', "'DSCR'!" + '@11', FMT_X],
        ['배당금', 'dividend', "'CF(Y)'!" + '@19', FMT_M],
        ['Project FCF', 'projectFcf', "'IRR'!" + '@13', FMT_M]
      ];
      defs.forEach(function (d, idx) {
        var rr = base + idx;
        label(ws, rr, d[0], null, { bold: idx === 2 });
        fillYears(ws, rr, function (y) {
          var v = rows[y - 1][d[1]];
          return { formula: d[2].replace('@', yc(y)), result: (v === null ? '-' : v) };
        }, d[3], { color: GREEN, bold: idx === 2 });
      });
    })();

    /* =========================================================
       11. Sensitivity — 민감도 (산출값)
       ========================================================= */
    (function (sens) {
      if (!sens) return;
      var ws = wb.addWorksheet('Sensitivity', { properties: { tabColor: { argb: 'FF2E7D62' } } });
      ws.getColumn(1).width = 2.5; ws.getColumn(2).width = 26;
      for (var c = 3; c <= 10; c++) ws.getColumn(c).width = 13;
      title(ws, '민감도 분석 (재무모델 생성기 산출값)');
      var r = 4;
      sens.forEach(function (blk) {
        section(ws, r, blk.title); r += 2;
        ws.getCell('B' + r).value = '변동값';
        ['Equity IRR', 'Project IRR', '최소 DSCR', 'NPV [KRWm]'].forEach(function (h, idx) {
          ws.getCell(colLetter(3 + idx) + r).value = h;
        });
        for (var c2 = 2; c2 <= 6; c2++) {
          var cc = ws.getCell(colLetter(c2) + r);
          cc.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
          cc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
          cc.alignment = { horizontal: 'center' };
        }
        r++;
        blk.data.forEach(function (d) {
          ws.getCell('B' + r).value = blk.fmtLabel(d.value);
          ws.getCell('B' + r).font = { name: FONT, size: 10 };
          put(ws, 'C' + r, d.equityIRR, FMT_P2);
          put(ws, 'D' + r, d.projectIRR, FMT_P2);
          put(ws, 'E' + r, d.minDSCR, FMT_X);
          put(ws, 'F' + r, d.npv, FMT_M);
          r++;
        });
        r += 2;
      });
      ws.getCell('B' + r).value = '※ 민감도 표는 각 케이스를 별도 계산한 결과값이며 수식으로 연결되어 있지 않습니다.';
      ws.getCell('B' + r).font = { name: FONT, size: 9, italic: true, color: { argb: 'FF9AA6A1' } };
    })(model.sensBlocks);

    /* =========================================================
       12. 목차 (맨 앞으로)
       ========================================================= */
    (function () {
      var ws = wb.addWorksheet('목차', { properties: { tabColor: { argb: 'FF0B2F24' } } });
      ws.getColumn(1).width = 2.5; ws.getColumn(2).width = 8; ws.getColumn(3).width = 20; ws.getColumn(4).width = 60;
      ws.getCell('B2').value = i.projectName;
      ws.getCell('B2').font = { name: FONT, bold: true, size: 16, color: { argb: 'FF14483A' } };
      ws.getCell('B3').value = '재무모델 (Financial Model) — 재생E AI Agent 자동생성  ·  ' +
        new Date().toISOString().slice(0, 10);
      ws.getCell('B3').font = { name: FONT, size: 10, color: { argb: 'FF6B7B76' } };
      var list = [
        ['1', 'Assum', '모든 입력 가정 — 노란 셀(파란 글씨)만 수정하면 전 시트 재계산'],
        ['2', 'TIC_Funding', '총투자비 · 건설이자 · 재원조달 · 감가상각 기준'],
        ['3', 'Revenue', '발전량 · 판매단가 · 영업수익 추정'],
        ['4', 'Opex', '운영비용 추정 (물가상승 반영)'],
        ['5', 'Debt', '차입금 인출/상환 스케줄 · 이자 · DSRA'],
        ['6', 'IS(Y)', '추정 손익계산서 (이월결손금·법인세 포함)'],
        ['7', 'CF(Y)', '추정 현금흐름표 · 배당 · 현금계정'],
        ['8', 'DSCR', '원리금 상환능력 (단순 DSCR, 최소/평균)'],
        ['9', 'IRR', 'Project IRR / Equity IRR / NPV'],
        ['10', 'Report', 'Executive Summary + 연도별 요약'],
        ['11', 'Sensitivity', '민감도 분석 (산출값)']
      ];
      var r = 6;
      ['No', '시트', '내용'].forEach(function (h, idx) {
        var c = ws.getCell(colLetter(2 + idx) + r);
        c.value = h;
        c.font = { name: FONT, bold: true, size: 10, color: { argb: WHITE } };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HDR_FILL } };
        c.alignment = { horizontal: 'center' };
      });
      r++;
      list.forEach(function (row) {
        ws.getCell('B' + r).value = row[0];
        ws.getCell('B' + r).alignment = { horizontal: 'center' };
        ws.getCell('C' + r).value = row[1];
        ws.getCell('C' + r).font = { name: FONT, size: 10, bold: true, color: { argb: 'FF14483A' } };
        ws.getCell('D' + r).value = row[2];
        ws.getCell('B' + r).font = ws.getCell('B' + r).font || { name: FONT, size: 10 };
        ws.getCell('D' + r).font = { name: FONT, size: 10 };
        r++;
      });
      r += 1;
      ws.getCell('B' + r).value = '색상 규칙';
      ws.getCell('B' + r).font = { name: FONT, bold: true, size: 10 };
      r++;
      [['파란 글씨 + 노란 배경', '직접 입력하는 가정값 (Assum 시트)'],
      ['검정 글씨', '해당 시트 내 수식'],
      ['초록 글씨', '다른 시트를 참조하는 수식'],
      ['단위', '별도 표기가 없는 금액은 모두 백만원(KRWm)']].forEach(function (x) {
        ws.getCell('C' + r).value = x[0];
        ws.getCell('C' + r).font = { name: FONT, size: 9, color: { argb: 'FF6B7B76' } };
        ws.getCell('D' + r).value = x[1];
        ws.getCell('D' + r).font = { name: FONT, size: 9, color: { argb: 'FF6B7B76' } };
        r++;
      });
    })();

    // 목차를 맨 앞으로 이동
    var order = ['목차', 'Report', 'Assum', 'Construction', 'TIC_Funding', 'Revenue', 'Opex', 'Debt', 'IS(Y)', 'CF(Y)', 'DSCR', 'IRR', 'Sensitivity'];
    order.forEach(function (n, idx) {
      var w = wb.getWorksheet(n);
      if (w) w.orderNo = idx + 1;
    });
    return wb;
  }

  global.SolarXlsx = { buildWorkbook: buildWorkbook, colLetter: colLetter };
})(typeof window !== 'undefined' ? window : globalThis);
