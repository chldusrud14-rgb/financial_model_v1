/* ============================================================
   재무모델 Excel 빌더 v2 (ExcelJS) — 분기 · 5트랜치
   - engine2.js의 계산 결과(값)를 그대로 굽는다. v1과 달리 라이브 수식이
     아니다 — 통합투자세액공제/최저한세/이익준비금/배당가능현금 캡처럼
     원본이 반복계산·연도 태그 SUMIF 등으로 얽혀 있어 엑셀 수식만으로
     재현하면 오차가 생긴다(CLAUDE.md "스컬프팅은 수식 자동화 불가"와
     같은 이유 — 그 원칙을 엔진 전체로 확장).
   - 시트: 목차 / Report / Funding / Debt(트랜치별+합계 한 시트) /
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
  var INPUT_FILL = 'FFFFF200'; // 재무모델링 관례 — 사용자가 직접 key-in한 값은 노란색으로 표시

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
    var IN = "'입력값'!"; // 다른 시트에서 입력값 시트 셀을 참조할 때 붙이는 접두사
    // 분기별 실측 오버라이드(engine2.js와 동일한 판정) — 있는 분기는 공식이
    // 아니라 실측 사실이라 수식화 대상에서 제외한다.
    var ovrByEnd = {};
    (inp.periodOverrides || []).forEach(function (o) { ovrByEnd[o.end] = o; });

    // 트랜치별 이자(운영) 시리즈 — Debt 시트와 IS(Q) 시트(이자비용 세부내역)가
    // 같은 계산을 각자 다시 하지 않도록 한 번만 구해서 공유한다.
    var trancheInterest = model.tranches.map(function (t, ti) {
      var bal = 0, ints = [];
      for (var n = 0; n < N; n++) {
        var draw = t.draws[n] || 0;
        bal += draw;
        var open2 = bal - draw;
        ints.push(periods[n].isOp ? open2 * t.rateO / (inp.ppy || 4) : 0);
        bal -= (rows[n].principalBy && rows[n].principalBy[ti]) || 0;
      }
      return { name: t.name, ints: ints };
    });

    // 항목별 운영비 추정치 — 실측 오버라이드(periodOverrides)가 있는
    // 분기는 항목별 실제값이 아니라 "선순위/후순위 합계"만 저장돼 있어서,
    // 공식(opexItems escalation)으로 각 항목의 상대 비중을 구한 뒤 실제
    // 합계에 맞춰 비례 배분한다 — 항목 합은 항상 검증된 실제 합계와
    // 정확히 일치하고, 항목 간 배분만 근사치다(오버라이드 없는 일반
    // 입력에서는 이 계산 자체가 곧 실제값이라 근사가 아니다).
    function itemizedOpex(n) {
      if (!inp.opexItems || !inp.opexItems.length) return null;
      var p = periods[n];
      var actual = (p.opexSenior || 0) + (p.opexSub || 0);
      if (!p.isOp || actual === 0) return inp.opexItems.map(function () { return 0; });
      var frac = p.opMonths / 12;
      var raws = inp.opexItems.map(function (it) {
        var esc = it.escal ? Math.pow(1 + it.escal / 100, p.opYearIdx) : 1;
        return it.annualKRWm * frac * esc;
      });
      var rawSum = raws.reduce(function (a, b) { return a + b; }, 0);
      var scale = rawSum > 0 ? actual / rawSum : 0;
      return raws.map(function (v) { return v * scale; });
    }
    // 오버라이드가 없는 분기는 itemizedOpex()의 비례배분(scale=1)이 곧
    // 순수 공식과 같으므로, 그 분기에 한해 "입력값" 시트를 참조하는
    // 수식 문자열을 만들어 준다(양수 기준 — 부호는 호출부에서 처리).
    function opexItemFormulaStr(idx, n) {
      var it = inp.opexItems[idx], ia = IN_ADDR.opexItem[idx];
      var p = periods[n];
      var fracF = '(' + p.opMonths + '/12)';
      var body = IN + ia.amount + '*100*' + fracF;
      if (it.escal) body += '*(1+' + IN + ia.escal + '/100)^' + p.opYearIdx;
      return body;
    }
    function opexPeriodIsFormulaable(n) {
      var p = periods[n];
      return !ovrByEnd[p.endStr] && p.isOp && ((p.opexSenior || 0) + (p.opexSub || 0)) !== 0;
    }
    // 오버라이드 분기의 항목별 금액 = 그 분기 실제 합계(입력값 시트 참조)를
    // 공식 기준 항목별 비중(raw)으로 비례배분한 수식 — itemizedOpex()가
    // JS로 하던 계산을 그대로 Excel 수식으로 옮긴 것. 항목 합은 실제
    // 합계와 정확히 일치하고, 항목 간 배분만 공식 비중 기준 근사치다.
    function opexItemOverrideFormula(idx, n) {
      var rawSum = inp.opexItems.map(function (it, j) { return '(' + opexItemFormulaStr(j, n) + ')'; }).join('+');
      var actual = IN + pc(n) + IN_ADDR.ovr.opexSenior + '+' + IN + pc(n) + IN_ADDR.ovr.opexSub;
      return 'IF(' + rawSum + '=0,0,(' + actual + ')*(' + opexItemFormulaStr(idx, n) + ')/(' + rawSum + '))';
    }
    // opex 실제값 유무와 무관하게, 오버라이드가 없는 운영 분기인지만 본다
    // (감가상각비·복구충당부채 전입액처럼 opex와 별개로 항상 발생하는
    // 항목에 쓴다).
    function isOpNonOverride(n) {
      var p = periods[n];
      return !ovrByEnd[p.endStr] && p.isOp;
    }
    // 오버라이드 분기의 선순위/후순위운영비도 다른 key-in 값들과 마찬가지로
    // Opex 시트에 직접 박아넣지 않고 "입력값" 시트의 "실측 오버라이드" 표를
    // 참조하게 한다. 오버라이드가 없는 분기(단순 0)는 그대로 값으로 둔다.
    function putOpexSeniorOrSub(ws, addr, n, key, fmt) {
      if (ovrByEnd[periods[n].endStr]) {
        putF(ws, addr, IN + pc(n) + IN_ADDR.ovr[key], fmt);
      } else {
        put(ws, addr, periods[n][key === 'opexSenior' ? 'opexSenior' : 'opexSub'] || 0, fmt);
      }
    }

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
        c1.alignment = { horizontal: 'center', textRotation: 0 };
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
    // put()과 같지만 값 대신 다른 시트를 참조하는 수식을 넣는다(입력값
    // 시트를 고치면 이 셀도 같이 바뀌는 라이브 연결 셀에 사용).
    function putF(ws, addr, formulaStr, fmt, opt) {
      opt = opt || {};
      var c = ws.getCell(addr);
      c.value = { formula: formulaStr };
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
      if (!hasNull && !opt.noSum) putF(ws, 'D' + r, sumFormula(r), fmt, Object.assign({ bold: true }, opt));
    }
    // D열 합계 — 값이 아니라 그 행의 분기 셀들(E~마지막 열)을 SUM하는
    // 수식으로 채운다. 각 분기 셀이 baked든 수식이든 상관없이 항상
    // 정확하고, 분기 셀을 나중에 고쳐도 합계가 같이 따라간다.
    function sumFormula(r) { return 'SUM(' + firstC + r + ':' + lastC + r + ')'; }

    /* =========================================================
       0. 입력값 — 화면(html)에서 key-in한 원본 입력을 한 시트에 모아둔다.
          Funding 등 다른 시트의 "단순 복사" 성격 셀(트랜치 조건/총사업비
          항목/사업자 지분 등)은 이 시트를 참조하는 수식으로 연결해서,
          입력값을 이 시트에서 고치면 해당 셀이 같이 바뀌는 걸 볼 수
          있게 한다. 다만 세금/최저한세/배당가능이익 캡처럼 반복계산·
          연도 태그 SUMIF로 얽힌 IS(Q)/CF(Q)의 계산 결과는 여기 대상이
          아니다 — 그대로 값(baked) 기준을 유지한다(이유는 파일 맨 위
          주석 참조). 실측 오버라이드(periodOverrides, "예시 불러오기")가
          적용된 분기는 애초에 공식이 아니라 실측 사실이라 수식화할 수
          없다 — 그런 분기는 이 시트의 입력값이 아니라 실측치가 곧
          정답이라는 뜻.
       ========================================================= */
    var IN_ADDR = { tranche: [], capexItem: [], opexItem: [], sh: [] };
    // Debt 시트가 채운 트랜치별 행 주소(openRow/intRow/prinRow/closeRow) —
    // IS(Q)의 이자비용 세부내역이 재계산하지 않고 그대로 참조하는 데 쓴다.
    var DEBT_TR_BLOCKS = [];
    var FUNDING_TIC_ADDR = null; // Funding 시트의 TIC(총투자비) 셀 주소 — 감가상각비 수식이 참조
    var FUNDING_IDC_ADDR = null, FUNDING_EQUITY_ADDR = null, FUNDING_DEBT_ADDR = null;
    var DEBT_DS_ROW = null; // Debt 시트 "전체 합계 원리금(DS)" 행 — CF(Q)가 참조
    var DEBT_INT_TOTAL_ROW = null, DEBT_PRIN_TOTAL_ROW = null; // Debt 전체합계 이자/원금 행 — Report의 IRR 현금흐름이 참조
    var DEBT_IDC_TOTAL_ROW = null, DEBT_CUMDRAW_ROW = {}, DEBT_IDC_SRC_ROW = {}; // 건설기간 자금조달 섹션 행 주소
    // 아직 안 만들어진 시트를 참조해야 하는 셀들 — 전부 만든 뒤 마지막에 채운다.
    var DEFERRED = [];
    var CFQ_MINDSCR_ROW = null, CFQ_MINCUMDSCR_ROW = null;
    var CFQ_PROJFLOW_ROW = null, CFQ_EQFLOW_ROW = null, CFQ_DIVFLOW_ROW = null, CFQ_INVFLOW_ROW = null;
    var CFQ_PROJIRR_ROW = null, CFQ_EQIRR_ROW = null, CFQ_DIVIRR_ROW = null, CFQ_INVIRR_ROW = null;
    var CFQ_PROJFLOWPRE_ROW = null, CFQ_PROJIRRPRE_ROW = null;
    var AR_WC_ROW = null; // Revenue 시트의 '운전자본 증감(A/R)' 행 — CF(Q)가 참조
    // Opex 시트가 채운 항목별 행 번호와 "영업비용 합계" 행 번호 — IS(Q)가
    // 재계산하지 않고 그대로 참조하는 데 쓴다.
    var OPEX_ITEM_ROWS = [];
    var OPEX_TOTAL_ROW = null;
    var OPEX_SENIOR_ROW = null, OPEX_SUB_ROW = null; // CF(Q)의 CFADS/FCFE가 참조
    var ISQ_TAX_ROW = null, ISQ_AGENTFEE_ROW = null; // CF(Q)의 CFADS가 참조
    var ISQ_NI_ROW = null; // CF(Q)의 배당가능이익 누적이 참조
    (function () {
      var ws = wb.addWorksheet('입력값', { properties: { tabColor: { argb: INPUT_FILL } } });
      ws.getColumn(1).width = 2.5; ws.getColumn(2).width = 22; ws.getColumn(3).width = 14;
      for (var c = 4; c <= 10; c++) ws.getColumn(c).width = 12;
      title(ws, '입력값 — 화면에서 입력한 값 (다른 시트가 이 시트를 참조)');
      var r = 4;
      section(ws, r, '사업 기본 가정'); r += 2;
      function kv(name, val, fmt) {
        var addr = 'C' + r;
        ws.getCell('B' + r).value = name;
        ws.getCell('B' + r).font = { name: FONT, size: 10 };
        put(ws, addr, val, fmt, { fill: INPUT_FILL });
        r++;
        return addr;
      }
      IN_ADDR.projectName = kv('사업명', inp.projectName, '@');
      IN_ADDR.capacityMW = kv('설비용량[MW]', inp.capacityMW, '#,##0.000');
      IN_ADDR.capexEok = kv('총사업비[억원]', inp.capexEok, FMT_M);
      IN_ADDR.equityEok = kv('자기자본[억원]', inp.equityEok, FMT_M);
      IN_ADDR.tariff = kv('판매단가[원/kWh]', inp.tariff, '#,##0.0');
      IN_ADDR.tariffEscal = kv('판매단가 에스컬레이션[%/yr]', inp.tariffEscal, '0.00');
      IN_ADDR.degradation = kv('발전량 열화율[%/yr]', inp.degradation, '0.00');
      IN_ADDR.auxRate = kv('소내소비율[%]', inp.auxRate, '0.00');
      if (inp.dailyHours != null) IN_ADDR.dailyHours = kv('일조시간[h/day]', inp.dailyHours, '0.000');
      else IN_ADDR.capacityFactor = kv('이용률(CF)[%]', inp.capacityFactor, '0.00');
      if (inp.depBaseOverride != null) IN_ADDR.depBaseOverride = kv('상각대상액(직접입력)[KRWm]', inp.depBaseOverride, FMT_M);
      else IN_ADDR.depRatio = kv('상각대상 비율(TIC 대비)[%]', inp.depRatio, '0.00');
      IN_ADDR.depYears = kv('감가상각 내용연수[yr]', inp.depYears, '0.0');
      IN_ADDR.decomEok = kv('철거비(복구충당부채)[억원]', inp.decomEok, FMT_M);
      if (!(inp.opexItems && inp.opexItems.length)) {
        IN_ADDR.opexEok = kv('운영비 총액[억원/yr]', inp.opexEok, FMT_M);
        IN_ADDR.opexEscal = kv('운영비 에스컬레이션[%/yr]', inp.opexEscal, '0.00');
        IN_ADDR.opexSubShare = kv('운영비 중 후순위 비중[%]', inp.opexSubShare || 0, '0.00');
      }
      IN_ADDR.agentFeeKRWm = kv('대리은행수수료(총액)[KRWm]', inp.agentFeeKRWm || 0, FMT_M);
      IN_ADDR.extraTaxDeductionKRWm = kv('기타 비현금 세무손금(총액)[KRWm]', inp.extraTaxDeductionKRWm || 0, FMT_M);
      IN_ADDR.dsraEok = kv('DSRA 최초 적립액[억원]', inp.dsraEok || 0, FMT_M);
      IN_ADDR.dsraMonths = kv('DSRA 적립기준(차기 X개월분)', inp.dsraMonths || 0, '0');
      IN_ADDR.minCash = kv('최소 보유현금[억원]', inp.minCash || 0, FMT_M);
      IN_ADDR.divDSCR = kv('배당 게이트 — 단순DSCR 최소치[x]', inp.divDSCR || 0, '0.0000');
      IN_ADDR.divCumDSCRVal = kv('배당 게이트 — 누적DSCR 최소치[x]', inp.divCumDSCR != null ? inp.divCumDSCR : 0, '0.0000');
      if (inp.dividendMonth != null) IN_ADDR.dividendMonth = kv('배당 지급월', inp.dividendMonth, '0');
      if (inp.firstDividendYear != null) IN_ADDR.firstDividendYear = kv('최초 배당 가능 연도', inp.firstDividendYear, '0');
      if (inp.dividendMonth == null) IN_ADDR.divStartYear = kv('배당 개시 운영연차', inp.divStartYear || 1, '0');
      r += 1;

      section(ws, r, '법인세 가정'); r += 2;
      IN_ADDR.taxFlat = inp.taxMode === 1 ? null : kv('법인세율(단일)[%]', inp.taxFlat || 0, '0.00');
      IN_ADDR.lossRate = kv('이월결손금 공제한도[%]', inp.lossRate != null ? inp.lossRate : 100, '0.00');
      IN_ADDR.amtRate = kv('최저한세율[%]', inp.amtRate || 0, '0.00');
      IN_ADDR.investmentCreditRate = kv('통합투자세액공제율[%]', inp.investmentCreditRate || 0, '0.00');
      IN_ADDR.localSurtaxRate = kv('지방소득세율(법인세 대비)[%]', inp.localSurtaxRate || 0, '0.00');
      IN_ADDR.creditSurtaxRate = kv('세액공제 농특세율[%]', inp.creditSurtaxRate || 0, '0.00');
      r += 1;

      // 건설기간 손실(연도별)·통합투자세액공제 기준액(연도별) — 연도를 키로 하는
      // 작은 표라서 항목명에 연도를 그대로 쓴다.
      function yearTable(title2, obj, addrKey) {
        var years = Object.keys(obj || {});
        if (!years.length) return;
        section(ws, r, title2); r += 2;
        ['연도', '금액[KRWm]'].forEach(function (h, idx) {
          var cc = ws.getCell(colLetter(2 + idx) + r);
          cc.value = h; cc.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
          cc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
          cc.alignment = { horizontal: 'center' };
        });
        r++;
        IN_ADDR[addrKey] = {};
        years.forEach(function (y) {
          put(ws, 'B' + r, Number(y), '0', { fill: INPUT_FILL });
          put(ws, 'C' + r, obj[y], FMT_M, { fill: INPUT_FILL });
          IN_ADDR[addrKey][y] = 'C' + r;
          r++;
        });
        r += 1;
      }
      yearTable('건설기간 손실(연도별)', inp.preOpLossByYear, 'preOpLossByYear');
      yearTable('통합투자세액공제 기준액(연도별)', inp.investmentCreditBaseByYear, 'investmentCreditBaseByYear');

      if (inp.tariffTracks && inp.tariffTracks.length) {
        section(ws, r, '판매단가 트랙 (PPA/SMP+REC 등)'); r += 2;
        ['트랙', '비중[%]', '단가[원/kWh]', '에스컬레이션[%/yr]', '대금회수시차[개월]'].forEach(function (h, idx) {
          var cc = ws.getCell(colLetter(2 + idx) + r);
          cc.value = h; cc.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
          cc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
          cc.alignment = { horizontal: 'center' };
        });
        r++;
        IN_ADDR.tariffTrack = [];
        inp.tariffTracks.forEach(function (tr, idx) {
          put(ws, 'B' + r, tr.name || ('트랙' + (idx + 1)), '@', { fill: INPUT_FILL });
          put(ws, 'C' + r, tr.share * 100, '0.00', { fill: INPUT_FILL });
          put(ws, 'D' + r, tr.price, '#,##0.0', { fill: INPUT_FILL });
          put(ws, 'E' + r, tr.escal || 0, '0.00', { fill: INPUT_FILL });
          put(ws, 'F' + r, tr.arLagMonths != null ? tr.arLagMonths : (inp.arLagMonths || 0), '0.0', { fill: INPUT_FILL });
          IN_ADDR.tariffTrack.push({ share: 'C' + r, price: 'D' + r, escal: 'E' + r, arLag: 'F' + r });
          r++;
        });
        r += 1;
      }

      section(ws, r, '트랜치 조건'); r += 2;
      ['트랜치', '금액[억원]', '건설금리[%]', '운영금리[%]', '거치(yr)', '상환(yr)', '방식', '투입순서'].forEach(function (h, idx) {
        var cc = ws.getCell(colLetter(2 + idx) + r);
        cc.value = h; cc.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
        cc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
        cc.alignment = { horizontal: 'center' };
      });
      r++;
      // model.tranches에는 rateC(건설금리)가 안 실려 있어서 con.srcs에서 꺼낸다
      // (srcs[0]=자본금, srcs[ti+1]=ti번째 트랜치). 예전엔 여기에 운영금리를
      // 그대로 넣어놔서 건설금리 칸이 사실상 잘못된 값이었음 — 이번에 IDC를
      // 수식화하면서 실제로 참조하게 되므로 바로잡는다.
      model.tranches.forEach(function (t, ti) {
        var src = model.con.srcs[ti + 1] || {};
        put(ws, 'B' + r, t.name, '@', { fill: INPUT_FILL });
        var a = { amount: 'C' + r, rateC: 'D' + r, rateO: 'E' + r, grace: 'F' + r, repay: 'G' + r, method: 'H' + r, order: 'I' + r };
        put(ws, a.amount, t.amount / 100, FMT_M, { fill: INPUT_FILL });
        put(ws, a.rateC, src.rateC != null ? src.rateC : t.rateO, FMT_P, { fill: INPUT_FILL });
        put(ws, a.rateO, t.rateO, FMT_P, { fill: INPUT_FILL });
        put(ws, a.grace, t.graceYears, '0.00', { fill: INPUT_FILL });
        put(ws, a.repay, t.repayYears, '0.00', { fill: INPUT_FILL });
        put(ws, a.method, t.method, '0', { fill: INPUT_FILL });
        put(ws, a.order, src.order != null ? src.order : '', '0', { fill: INPUT_FILL });
        IN_ADDR.tranche.push(a);
        r++;
      });
      // 자본금도 인출 순서상 트랜치들과 같은 워터폴에 참여한다(투입순서 기준).
      put(ws, 'B' + r, '자본금', '@');
      putF(ws, 'C' + r, IN_ADDR.equityEok, FMT_M);
      put(ws, 'I' + r, model.con.srcs[0] ? model.con.srcs[0].order : 1, '0', { fill: INPUT_FILL });
      IN_ADDR.equityOrder = 'I' + r;
      IN_ADDR.equityAmountRef = 'C' + r;
      r += 2;

      // 총사업비 지출 스케줄 — 건설 분기별 지출 비중(합 1). IDC 계산의
      // 출발점이라 수식화하려면 이 표가 필요하다.
      if (model.con.conPs && model.con.conPs.length) {
        for (var c = C0; c <= C0 + N - 1; c++) ws.getColumn(c).width = 11;
        section(ws, r, '총사업비 지출 스케줄 (건설 분기별 비중)'); r += 2;
        periodHeader(ws, r); r += 2;
        label(ws, r, '지출 비중', '[비율]');
        model.con.conPs.forEach(function (cp, ci) {
          put(ws, pc(cp.n) + r, model.con.curve[ci], '0.000000', { fill: INPUT_FILL });
        });
        putF(ws, 'D' + r, sumFormula(r), '0.000000');
        IN_ADDR.spendCurveRow = r;
        r += 2;
      }

      // 방식3(64회차 직접 키인, "예시 불러오기" 전용) 트랜치의 상환비율도
      // 결국 원본에서 그대로 가져온 key-in 데이터다 — Debt 시트에 직접
      // 박아넣지 않고 여기 표로 두고 참조하게 한다. 표 자체(비율 배열)를
      // "계산"할 수는 없지만(그 자체가 원본 값), Debt는 이제 이 표를
      // 참조하는 수식만 갖는다.
      var scheduleTranches = model.tranches.map(function (t, ti) { return t.method === 3 && t.schedule ? ti : -1; }).filter(function (ti) { return ti >= 0; });
      if (scheduleTranches.length) {
        for (var c = C0; c <= C0 + N - 1; c++) ws.getColumn(c).width = 11;
        section(ws, r, '방식3 상환비율 ("예시 불러오기" 전용 — 원본 상환 스케줄)'); r += 2;
        periodHeader(ws, r); r += 2;
        IN_ADDR.schedule = {};
        scheduleTranches.forEach(function (ti) {
          var t = model.tranches[ti];
          label(ws, r, t.name + ' 상환비율', '[비율]');
          for (var k = 0; k < t.schedule.length; k++) {
            var n = t.repayStartIdx + k;
            if (n >= 0 && n < N) put(ws, pc(n) + r, t.schedule[k], '0.000000', { fill: INPUT_FILL });
          }
          IN_ADDR.schedule[ti] = r;
          r++;
        });
        r += 1;
      }

      if (model.capexItems && model.capexItems.length) {
        section(ws, r, '총사업비 항목'); r += 2;
        ['항목', '금액[억원]'].forEach(function (h, idx) {
          var cc = ws.getCell(colLetter(2 + idx) + r);
          cc.value = h; cc.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
          cc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
          cc.alignment = { horizontal: 'center' };
        });
        r++;
        model.capexItems.forEach(function (it) {
          put(ws, 'B' + r, it.name, '@', { fill: INPUT_FILL });
          put(ws, 'C' + r, it.amountEok != null ? it.amountEok : null, FMT_M, { fill: INPUT_FILL });
          IN_ADDR.capexItem.push({ name: 'B' + r, amount: 'C' + r, hasAmount: it.amountEok != null });
          r++;
        });
        r += 1;
      }

      if (inp.opexItems && inp.opexItems.length) {
        section(ws, r, '운영비 항목'); r += 2;
        ['항목', '연간금액[억원/yr]', '에스컬레이션[%/yr]'].forEach(function (h, idx) {
          var cc = ws.getCell(colLetter(2 + idx) + r);
          cc.value = h; cc.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
          cc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
          cc.alignment = { horizontal: 'center' };
        });
        r++;
        inp.opexItems.forEach(function (it) {
          put(ws, 'B' + r, it.name, '@', { fill: INPUT_FILL });
          put(ws, 'C' + r, it.annualKRWm / 100, FMT_M, { fill: INPUT_FILL });
          put(ws, 'D' + r, it.escal || 0, '0.00', { fill: INPUT_FILL });
          IN_ADDR.opexItem.push({ name: 'B' + r, amount: 'C' + r, escal: 'D' + r });
          r++;
        });
        r += 1;
      }

      var sh0 = model.kpi && model.kpi.shareholders;
      if (sh0 && sh0.length > 1) {
        section(ws, r, '사업자 구성'); r += 2;
        ['사업자', '지분율[%]'].forEach(function (h, idx) {
          var cc = ws.getCell(colLetter(2 + idx) + r);
          cc.value = h; cc.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
          cc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
          cc.alignment = { horizontal: 'center' };
        });
        r++;
        sh0.forEach(function (s) {
          put(ws, 'B' + r, s.name, '@', { fill: INPUT_FILL });
          put(ws, 'C' + r, s.stakePct, '0.00', { fill: INPUT_FILL });
          IN_ADDR.sh.push({ name: 'B' + r, stake: 'C' + r });
          r++;
        });
      }

      var hasOverride = Object.keys(ovrByEnd).length > 0;
      if (hasOverride) {
        // "예시 불러오기"의 분기별 실측치도 결국은 key-in 데이터다(원본
        // FS에서 뽑아온 값이라는 것만 다르다) — 그래서 Revenue/Opex 시트에
        // 직접 박아넣지 않고, 다른 입력값처럼 여기 모아두고 그 시트들이
        // 참조하게 한다. 공식이 없는 분기는 셀을 비워둔다(그 분기만
        // "그 자체가 정답"이라는 뜻).
        for (var c = C0; c <= C0 + N - 1; c++) ws.getColumn(c).width = 11;
        r += 1;
        section(ws, r, '실측 오버라이드 ("예시 불러오기" 전용 — 공식이 아니라 원본 FS의 분기별 실제값)'); r += 2;
        periodHeader(ws, r); r += 2;
        IN_ADDR.ovr = {};
        [
          ['price', '판매단가(실측)', '[원/kWh]', '#,##0.0', function (n) {
            var ovr = ovrByEnd[periods[n].endStr];
            if (!ovr || !(periods[n].gen > 0)) return null;
            return ovr.revenue / periods[n].gen * 1000;
          }],
          ['opexSenior', '선순위운영비(실측)', '[KRWm]', FMT_M, function (n) {
            var ovr = ovrByEnd[periods[n].endStr];
            return ovr ? ovr.opexSenior : null;
          }],
          ['opexSub', '후순위운영비(실측)', '[KRWm]', FMT_M, function (n) {
            var ovr = ovrByEnd[periods[n].endStr];
            return ovr ? ovr.opexSub : null;
          }],
          ['decomAccrual', '복구충당부채 전입액(실측)', '[KRWm]', FMT_M, function (n) {
            var ovr = ovrByEnd[periods[n].endStr];
            return ovr ? ovr.decomAccrual : null;
          }],
          ['agentFee', '대리은행수수료(실측)', '[KRWm]', FMT_M, function (n) {
            var ovr = ovrByEnd[periods[n].endStr];
            return ovr ? ovr.agentFee : null;
          }],
          ['extraTaxDed', '기타 비현금 세무손금(실측)', '[KRWm]', FMT_M, function (n) {
            var ovr = ovrByEnd[periods[n].endStr];
            return ovr ? ovr.extraTaxDed : null;
          }],
          ['taxCash', '현금기준 법인세(실측)', '[KRWm]', FMT_M, function (n) {
            var ovr = ovrByEnd[periods[n].endStr];
            return ovr && ovr.taxCash != null ? ovr.taxCash : null;
          }],
          ['dsraFcfe', 'DSRA FCFE 조정(실측)', '[KRWm]', FMT_M, function (n) {
            var ovr = ovrByEnd[periods[n].endStr];
            return ovr && ovr.dsraFcfe != null ? ovr.dsraFcfe : null;
          }],
          ['wc', '운전자본 증감(실측)', '[KRWm]', FMT_M, function (n) {
            var ovr = ovrByEnd[periods[n].endStr];
            return ovr && ovr.wc != null ? ovr.wc : null;
          }]
        ].forEach(function (spec) {
          var key = spec[0], label2 = spec[1], unit = spec[2], fmt = spec[3], getter = spec[4];
          label(ws, r, label2, unit);
          for (var n = 0; n < N; n++) {
            var v = getter(n);
            put(ws, pc(n) + r, v, fmt, v != null ? { fill: INPUT_FILL } : {});
          }
          IN_ADDR.ovr[key] = r;
          r++;
        });
      }

      r += 1;
      ws.getCell('B' + r).value = '※ 노란색 셀 = 화면에서 직접 key-in했거나("예시 불러오기"의 실측치 포함) 그에 준하는 입력값(재무모델링 관례). 세금·최저한세·배당가능이익 캡 등 IS(Q)/CF(Q)의 계산 결과는 반복계산·조건부 누적 로직이 얽혀 있어 여기 대상이 아니며 값(baked) 기준입니다.';
      ws.getCell('B' + r).font = { name: FONT, size: 8, italic: true, color: { argb: 'FF9AA6A1' } };
    })();

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
      put(ws, 'B' + r, '자본금'); putF(ws, 'C' + r, IN + IN_ADDR.equityEok + '*100', FMT_M);
      FUNDING_EQUITY_ADDR = 'C' + r; r++;
      var trAmtRows = [];
      model.tranches.forEach(function (t, ti) {
        var ia = IN_ADDR.tranche[ti];
        put(ws, 'B' + r, t.name);
        putF(ws, 'C' + r, IN + ia.amount + '*100', FMT_M);
        trAmtRows.push(r);
        put(ws, 'D' + r, t.order === undefined ? '' : t.order, '0');
        putF(ws, 'E' + r, IN + ia.rateC, FMT_P);
        putF(ws, 'F' + r, IN + ia.rateO, FMT_P);
        putF(ws, 'G' + r, IN + ia.grace, '0.00');
        putF(ws, 'H' + r, IN + ia.repay, '0.00');
        putF(ws, 'I' + r, IN + ia.method, '0');
        put(ws, 'J' + r, t.repayStartIdx >= 0 ? periods[t.repayStartIdx].endStr : '-', '@');
        r++;
      });
      put(ws, 'B' + r, '차입금 합계', '@', { bold: true, fill: SUB_FILL });
      putF(ws, 'C' + r, trAmtRows.map(function (rr) { return 'C' + rr; }).join('+'), FMT_M, { bold: true, fill: SUB_FILL });
      FUNDING_DEBT_ADDR = 'C' + r; r++;
      r += 1;
      section(ws, r, '건설이자(IDC)'); r += 2;
      // IDC 합계/TIC는 Debt 시트의 "건설기간 자금조달" 섹션을 참조해야 하는데
      // 그 시트가 아직 안 만들어졌으므로, 주소만 잡아두고 실제 수식은 마지막에
      // 채운다(DEFERRED).
      put(ws, 'B' + r, 'IDC 합계[KRWm]'); FUNDING_IDC_ADDR = 'C' + r;
      DEFERRED.push(function () { putF(ws, FUNDING_IDC_ADDR, "'Debt'!D" + DEBT_IDC_TOTAL_ROW, FMT_M); });
      r++;
      put(ws, 'B' + r, '총투자비(TIC)[KRWm]'); FUNDING_TIC_ADDR = 'C' + r;
      DEFERRED.push(function () {
        putF(ws, FUNDING_TIC_ADDR, IN + IN_ADDR.capexEok + "*100+'Debt'!D" + DEBT_IDC_TOTAL_ROW, FMT_M);
      });
      r++;
      put(ws, 'B' + r, '총사업비(건설이자 제외)[KRWm]'); putF(ws, 'C' + r, IN + IN_ADDR.capexEok + '*100', FMT_M); r++;

      // 총사업비 세부내역 — 화면에서 항목별로 입력했으면 실제 금액, 합계만
      // 입력했으면 항목명만(금액은 빈칸) 표시한다. 어느 쪽이든 항목
      // 구성 자체는 항상 보이게 해달라는 요청 반영.
      if (model.capexItems && model.capexItems.length) {
        r += 1;
        section(ws, r, '총사업비 세부내역'); r += 2;
        ['항목', '금액[KRWm]'].forEach(function (h, idx) {
          var c = ws.getCell(colLetter(2 + idx) + r);
          c.value = h;
          c.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
          c.alignment = { horizontal: 'center' };
        });
        r++;
        model.capexItems.forEach(function (it, idx) {
          var ia = IN_ADDR.capexItem[idx];
          putF(ws, 'B' + r, IN + ia.name, '@');
          if (ia.hasAmount) putF(ws, 'C' + r, IN + ia.amount + '*100', FMT_M);
          r++;
        });
        // 항목별 금액을 안 넣었어도(전부 빈칸이어도) 합계는 항상 위
        // "총사업비" 입력값 그대로 나온다 — 항목 합계가 아니라 실제
        // 입력된 총사업비를 그대로 쓰므로 항목을 일부만 채워도 항상
        // 정확하다.
        put(ws, 'B' + r, '합계', '@', { bold: true, fill: SUB_FILL });
        putF(ws, 'C' + r, IN + IN_ADDR.capexEok + '*100', FMT_M, { bold: true, fill: SUB_FILL });
        r++;
      }

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
        sh.forEach(function (s, idx) {
          var ia = IN_ADDR.sh[idx];
          putF(ws, 'B' + r, IN + ia.name, '@');
          putF(ws, 'C' + r, IN + ia.stake + '/100', FMT_P);
          put(ws, 'D' + r, s.equityKRWm, FMT_M);
          put(ws, 'E' + r, s.dividendKRWm, FMT_M);
          r++;
        });
      }
    })();

    /* =========================================================
       2. Debt — 트랜치별 + 합계를 한 시트에 위아래로 이어붙임
          (예전엔 트랜치마다 별도 시트 5개 + 합계 시트였는데, 시트가
          너무 쪼개져 있어 하나로 합쳐달라는 요청 반영)
       ========================================================= */
    (function () {
      var ws = sheet('Debt');
      title(ws, '차입금 상환 스케줄 — 건설기간 인출 + 트랜치별 상환');
      var r = 4;

      /* =========================================================
         건설기간 자금조달 — 인출(draw)과 건설이자(IDC)를 수식으로 계산한다.

         원본 JS는 "분기마다 필요한 돈을 투입순서대로 나눠준다"를 반복문으로
         돌리는데, 그걸 그대로 옮기려면 복잡하다. 대신 수학적으로 동일한
         **누적 기준**으로 재구성했다:

           누적소요액(n) = 누적소요액(n-1) + 공사비(n) + DSRA(n) + 건설이자(n)
           누적인출(트랜치) = MIN(MAX(누적소요액 - 선순위 약정합, 0), 자기 그룹
                              약정합) × 자기 약정액 ÷ 그룹 약정합
           이번 분기 인출 = 누적인출(n) - 누적인출(n-1)
           건설이자(n)    = 누적인출(n-1) × 건설금리 ÷ 4

         "같은 투입순서 그룹 안에서는 약정금액 비율대로 나눠 가진다"는 원본
         로직이 매 분기 잔여여력 비율을 유지하므로(귀납적으로 항상 약정액
         비례), 누적 기준으로 봐도 정확히 같은 값이 나온다.

         핵심: 건설이자(n)은 "n-1까지의 누적인출"만 참조하므로 **순환참조가
         아니다** — Debt 시트의 "기초잔액=직전 기말잔액"과 같은 체인이다.
         (엔진의 60회 반복문도 실제로는 항상 2회 만에 같은 값으로 끝난다 —
         되먹임이 없는 구조라는 뜻.)
         ========================================================= */
      var srcs = model.con.srcs || [];
      var codIdxC = model.con.codIdx;
      // 소스별 약정액 셀 참조(자본금은 0번, 트랜치는 1번부터)
      function srcAmountF(si) {
        return si === 0
          ? (IN + IN_ADDR.equityEok + '*100')
          : (IN + IN_ADDR.tranche[si - 1].amount + '*100');
      }
      function sumAmountsF(list) {
        return list.length ? list.map(srcAmountF).join('+') : '0';
      }
      var srcPrior = [], srcGroup = [];
      srcs.forEach(function (s, si) {
        var prior = [], group = [];
        srcs.forEach(function (s2, si2) {
          if (s2.order < s.order) prior.push(si2);
          else if (s2.order === s.order) group.push(si2);
        });
        srcPrior.push(prior); srcGroup.push(group);
      });

      section(ws, r, '건설기간 자금조달 (인출·건설이자)'); r += 2;
      periodHeader(ws, r); r += 2;

      var capexSpendRow = r;
      label(ws, r, '공사비 지출', '[KRWm]');
      for (var n = 0; n < N; n++) {
        if (IN_ADDR.spendCurveRow != null) {
          putF(ws, pc(n) + r, IN + pc(n) + IN_ADDR.spendCurveRow + '*' + IN + IN_ADDR.capexEok + '*100', FMT_M);
        } else {
          put(ws, pc(n) + r, 0, FMT_M);
        }
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M); r++;

      var dsraFundRow = r;
      label(ws, r, 'DSRA 적립', '[KRWm]');
      for (var n = 0; n < N; n++) {
        if (n === codIdxC) putF(ws, pc(n) + r, IN + IN_ADDR.dsraEok + '*100', FMT_M);
        else put(ws, pc(n) + r, 0, FMT_M);
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M); r++;

      // 소스별 건설이자 — 자본금은 이자가 없으므로 트랜치만.
      var idcSrcRow = {};
      srcs.forEach(function (s, si) {
        if (si === 0) return;
        idcSrcRow[si] = r;
        label(ws, r, s.name + ' 건설이자', '[KRWm]', { indent: true });
        r++;
      });
      var idcTotalRow = r;
      label(ws, r, '건설이자 합계(당기)', '[KRWm]', { bold: true });
      r++;
      var cumNeedRow = r;
      label(ws, r, '누적 소요액', '[KRWm]');
      r++;
      var cumDrawRow = {};
      srcs.forEach(function (s, si) {
        cumDrawRow[si] = r;
        label(ws, r, s.name + ' 누적인출', '[KRWm]', { indent: true });
        r++;
      });

      // 열 방향으로 채운다 — 같은 열 안에서 건설이자 → 누적소요액 → 누적인출
      // 순서로 참조가 흐르고, 건설이자만 직전 열을 참조한다.
      for (var n = 0; n < N; n++) {
        var prevC2 = n > 0 ? pc(n - 1) : null;
        srcs.forEach(function (s, si) {
          if (si === 0) return;
          var addr = pc(n) + idcSrcRow[si];
          if (n <= codIdxC && prevC2) {
            putF(ws, addr, prevC2 + cumDrawRow[si] + '*' + IN + IN_ADDR.tranche[si - 1].rateC + '/' + (inp.ppy || 4), FMT_M);
          } else {
            put(ws, addr, 0, FMT_M);
          }
        });
        putF(ws, pc(n) + idcTotalRow,
          Object.keys(idcSrcRow).map(function (si) { return pc(n) + idcSrcRow[si]; }).join('+') || '0',
          FMT_M, { bold: true });
        putF(ws, pc(n) + cumNeedRow,
          (prevC2 ? prevC2 + cumNeedRow + '+' : '') +
          pc(n) + capexSpendRow + '+' + pc(n) + dsraFundRow + '+' + pc(n) + idcTotalRow,
          FMT_M, { noSum: true });
        srcs.forEach(function (s, si) {
          var groupF = sumAmountsF(srcGroup[si]);
          var priorF = sumAmountsF(srcPrior[si]);
          // 그룹 약정합이 0이면 분자(MIN(...,0))도 0이라 결과가 0이 되지만,
          // 분모를 IF로 감싸 0으로 나누는 상황 자체를 막는다(#DIV/0! 방지).
          putF(ws, pc(n) + cumDrawRow[si],
            'MIN(MAX(' + pc(n) + cumNeedRow + '-(' + priorF + '),0),' + groupF + ')*(' + srcAmountF(si) + ')/IF((' + groupF + ')=0,1,' + groupF + ')',
            FMT_M, { noSum: true });
        });
      }
      srcs.forEach(function (s, si) {
        if (si === 0) return;
        putF(ws, 'D' + idcSrcRow[si], sumFormula(idcSrcRow[si]), FMT_M);
      });
      putF(ws, 'D' + idcTotalRow, sumFormula(idcTotalRow), FMT_M, { bold: true });
      DEBT_IDC_TOTAL_ROW = idcTotalRow;
      DEBT_CUMDRAW_ROW = cumDrawRow;
      DEBT_IDC_SRC_ROW = idcSrcRow;
      r += 2;

      section(ws, r, '트랜치별 인출/상환 + 전체 합계'); r += 2;
      periodHeader(ws, r); r += 2;

      // 방식 1(원금균등)/2(원리금균등)은 화면에서 사용자가 직접 고를 수 있는
      // 방식이라 순수 재무공식(PMT 등)으로 완전히 재현 가능 — 이자/원금/잔액을
      // "입력값" 시트를 참조하는 수식으로 연결한다. 방식 3(64회차 직접 키인,
      // "예시 불러오기" 프리셋 전용, 화면에는 없음)은 스케줄 자체가 원본 실측
      // 데이터라 수식화 대상이 아니라 그대로 값(baked)을 쓴다.
      var trBlocks = DEBT_TR_BLOCKS;
      model.tranches.forEach(function (t, ti) {
        var ia = IN_ADDR.tranche[ti];
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
        var canFormula = (t.method === 1 || t.method === 2) && t.nRepay > 0;
        var canScheduleLink = t.method === 3 && t.schedule && IN_ADDR.schedule && IN_ADDR.schedule[ti] != null;
        label(ws, r, t.name, null, { bold: true, fill: SUB_FILL }); r++;
        var openRow = r, drawRow = r + 1, idcRow = r + 2, intRow = r + 3, prinRow = r + 4, closeRow = r + 5;
        trBlocks.push({ openRow: openRow, intRow: intRow, prinRow: prinRow, closeRow: closeRow });

        label(ws, openRow, '기초잔액', '[KRWm]');
        for (var n = 0; n < N; n++) {
          if (n === 0) put(ws, pc(0) + openRow, 0, FMT_M, { noSum: true });
          else putF(ws, pc(n) + openRow, pc(n - 1) + closeRow, FMT_M, { noSum: true });
        }
        // 인출 = 이번 분기 누적인출 - 직전 분기 누적인출, 건설이자는 위
        // "건설기간 자금조달" 섹션에서 이미 계산한 행을 그대로 참조한다.
        var cdRow = DEBT_CUMDRAW_ROW[ti + 1];
        label(ws, drawRow, '인출', '[KRWm]');
        for (var n = 0; n < N; n++) {
          putF(ws, pc(n) + drawRow, pc(n) + cdRow + (n > 0 ? ('-' + pc(n - 1) + cdRow) : ''), FMT_M);
        }
        putF(ws, 'D' + drawRow, sumFormula(drawRow), FMT_M);
        label(ws, idcRow, '건설이자(IDC)', '[KRWm]');
        for (var n = 0; n < N; n++) putF(ws, pc(n) + idcRow, pc(n) + DEBT_IDC_SRC_ROW[ti + 1], FMT_M);
        putF(ws, 'D' + idcRow, sumFormula(idcRow), FMT_M);

        label(ws, intRow, '이자(운영)', '[KRWm]');
        for (var n = 0; n < N; n++) {
          if (periods[n].isOp) putF(ws, pc(n) + intRow, pc(n) + openRow + '*' + IN + ia.rateO + '/4', FMT_M);
          else put(ws, pc(n) + intRow, 0, FMT_M);
        }
        putF(ws, 'D' + intRow, sumFormula(intRow), FMT_M, { bold: true });

        label(ws, prinRow, '원금상환', '[KRWm]');
        for (var n = 0; n < N; n++) {
          if (canFormula && n >= t.repayStartIdx && n <= t.repayEndIdx) {
            if (t.method === 1) {
              putF(ws, pc(n) + prinRow, IN + ia.amount + '*100/' + t.nRepay, FMT_M);
            } else {
              putF(ws, pc(n) + prinRow,
                'PMT(' + IN + ia.rateO + '/4,' + t.nRepay + ',-' + IN + ia.amount + '*100)-' +
                pc(n) + openRow + '*' + IN + ia.rateO + '/4', FMT_M);
            }
          } else if (canScheduleLink && n >= t.repayStartIdx && n < t.repayStartIdx + t.schedule.length) {
            putF(ws, pc(n) + prinRow, IN + pc(n) + IN_ADDR.schedule[ti] + '*' + IN + ia.amount + '*100', FMT_M);
          } else {
            put(ws, pc(n) + prinRow, prins[n], FMT_M);
          }
        }
        putF(ws, 'D' + prinRow, sumFormula(prinRow), FMT_M, { bold: true });

        label(ws, closeRow, '기말잔액', '[KRWm]', { bold: true });
        for (var n = 0; n < N; n++) {
          putF(ws, pc(n) + closeRow,
            pc(n) + openRow + '+' + pc(n) + drawRow + '-' + pc(n) + prinRow, FMT_M, { bold: true });
        }
        r = closeRow + 1;
        label(ws, r, '미상환 잔액(검증용)', '[KRWm]');
        putF(ws, 'D' + r, lastC + closeRow, FMT_M);
        putF(ws, 'F' + r, 'IF(ABS(D' + r + ')<1,"완전상환 확인 (OK)","경고: 미상환 잔액")', '@');
        if (canScheduleLink) {
          put(ws, 'H' + r, '※ 방식 3(직접 키인) — 원금상환은 "입력값" 시트의 상환비율 표를 참조하는 수식', null);
          ws.getCell('H' + r).font = { name: FONT, size: 8, italic: true, color: { argb: 'FF9AA6A1' } };
        } else if (!canFormula) {
          put(ws, 'H' + r, '※ 미사용 트랜치 — 값(baked) 기준', null);
          ws.getCell('H' + r).font = { name: FONT, size: 8, italic: true, color: { argb: 'FF9AA6A1' } };
        }
        r += 2;
      });

      section(ws, r, '전체 합계'); r++;
      label(ws, r, '기초잔액', '[KRWm]');
      for (var n = 0; n < N; n++) {
        putF(ws, pc(n) + r, trBlocks.map(function (b) { return pc(n) + b.openRow; }).join('+'), FMT_M, { noSum: true });
      }
      r++;
      var intTotalRow = r;
      DEBT_INT_TOTAL_ROW = r;
      label(ws, r, '이자', '[KRWm]');
      for (var n = 0; n < N; n++) {
        putF(ws, pc(n) + r, trBlocks.map(function (b) { return pc(n) + b.intRow; }).join('+'), FMT_M);
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M, { bold: true });
      r++;
      var prinTotalRow = r;
      DEBT_PRIN_TOTAL_ROW = r;
      label(ws, r, '원금상환', '[KRWm]');
      for (var n = 0; n < N; n++) {
        putF(ws, pc(n) + r, trBlocks.map(function (b) { return pc(n) + b.prinRow; }).join('+'), FMT_M);
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M, { bold: true });
      r++;
      label(ws, r, '원리금(DS)', '[KRWm]', { bold: true, fill: SUB_FILL });
      for (var n = 0; n < N; n++) {
        putF(ws, pc(n) + r, pc(n) + intTotalRow + '+' + pc(n) + prinTotalRow, FMT_M, { bold: true });
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M, { bold: true });
      DEBT_DS_ROW = r;
      r++;
      label(ws, r, '기말잔액', '[KRWm]');
      for (var n = 0; n < N; n++) {
        putF(ws, pc(n) + r, trBlocks.map(function (b) { return pc(n) + b.closeRow; }).join('+'), FMT_M, { noSum: true });
      }
      r++;
      r++;
      // DSRA 기초/기말잔액 — CF(Q)에서 이미 검증한 것과 같은 로직(차기 X개월분
      // 원리금을 미리 내다보는 SUM, COD 시점에 최초 적립액으로 리셋)을 이 시트
      // 자체의 DS 행(DEBT_DS_ROW)만으로 독립적으로 재현한다.
      var dsraOpenRowD = r, dsraCloseRowD = r + 1;
      var nqDsraD = Math.round((inp.dsraMonths || 0) / (12 / (inp.ppy || 4)));
      var codIdxD = model.con.codIdx;
      label(ws, dsraOpenRowD, 'DSRA 기초잔액', '[KRWm]');
      label(ws, dsraCloseRowD, 'DSRA 기말잔액', '[KRWm]');
      for (var n = 0; n < N; n++) {
        var prevColD = n > 0 ? pc(n - 1) : null;
        if (n === codIdxD) {
          putF(ws, pc(n) + dsraOpenRowD, IN + IN_ADDR.dsraEok + '*100', FMT_M, { noSum: true });
        } else {
          putF(ws, pc(n) + dsraOpenRowD, prevColD ? (prevColD + dsraCloseRowD) : '0', FMT_M, { noSum: true });
        }
        if (periods[n].isOp) {
          var needTermsD = [];
          for (var k = 1; k <= nqDsraD && n + k < N; k++) needTermsD.push(pc(n + k) + DEBT_DS_ROW);
          putF(ws, pc(n) + dsraCloseRowD, needTermsD.length ? needTermsD.join('+') : '0', FMT_M, { noSum: true });
        } else {
          putF(ws, pc(n) + dsraCloseRowD, pc(n) + dsraOpenRowD, FMT_M, { noSum: true });
        }
      }
    })();

    /* =========================================================
       3. Revenue — 발전량은 항상 수식화(오버라이드도 발전량 자체엔 영향
          없음). 판매단가/영업수익은 트랙별 입력(tariffTracks)이 없고
          해당 분기가 실측 오버라이드 대상이 아닐 때만 수식으로 연결
          (오버라이드 분기는 그 자체가 실측 사실이라 baked 유지).
       ========================================================= */
    (function () {
      var ws = sheet('Revenue');
      title(ws, '영업수익 추정');
      section(ws, 4, '발전량 및 매출');
      periodHeader(ws, 6);
      var ppy = inp.ppy || 4;
      label(ws, 9, '발전량', '[MWh]');
      for (var n = 0; n < N; n++) {
        var p = periods[n];
        var degF = 'MAX(0,1-' + IN + IN_ADDR.degradation + '/100*' + p.opYearIdx + ')';
        var annualGenF = inp.dailyHours != null
          ? IN + IN_ADDR.capacityMW + '*' + IN + IN_ADDR.dailyHours + '*365'
          : IN + IN_ADDR.capacityMW + '*8760*(' + IN + IN_ADDR.capacityFactor + '/100)';
        var frac = p.opMonths / 12;
        var full = ppy === 4 && p.opMonths === (12 / ppy);
        var genFrac = (inp.seasonality && full) ? inp.seasonality[p.end.getUTCMonth() + 1] : frac;
        putF(ws, pc(n) + 9, '(' + annualGenF + ')*' + degF + '*(1-' + IN + IN_ADDR.auxRate + '/100)*' + genFrac, '#,##0');
      }
      putF(ws, 'D9', sumFormula(9), '#,##0', { bold: true });

      label(ws, 10, '판매단가', '[원/kWh]');
      label(ws, 12, '영업수익', '[KRWm]', { bold: true, fill: SUB_FILL });
      var hasTracks = inp.tariffTracks && inp.tariffTracks.length;
      var anyOverride = false;
      for (var n = 0; n < N; n++) {
        var p = periods[n];
        var ovr = !!(ovrByEnd[p.endStr]);
        if (!p.isOp) {
          put(ws, pc(n) + 10, null, '#,##0.0');
          put(ws, pc(n) + 12, 0, FMT_M, { bold: true });
          continue;
        }
        if (ovr) {
          // 오버라이드 분기(실측치) — "매출 숫자를 통째로 박아넣지 말고
          // 발전량×단가로 계산되게 해달라"는 요청 반영. 실측 사실 자체는
          // "그 분기의 실제 정산단가"이고, 그 값 자체도 다른 key-in
          // 입력들과 마찬가지로 "입력값" 시트("실측 오버라이드" 표)에
          // 두고 여기서는 그 셀을 참조만 한다. 매출은 다른 분기와 완전히
          // 동일하게 발전량×단가/1000 수식으로 계산한다 — 값은 원본과
          // 그대로 일치(단가 자체가 실측 매출/발전량이므로).
          anyOverride = true;
          putF(ws, pc(n) + 10, IN + pc(n) + IN_ADDR.ovr.price, '#,##0.0');
          putF(ws, pc(n) + 12, pc(n) + '9*' + pc(n) + '10/1000', FMT_M, { bold: true });
        } else if (hasTracks) {
          // 트랙별(SMP+REC/PPA 등) 입력 — 매출 = 발전량×Σ(트랙비중×트랙단가)/1000,
          // 트랙단가는 각각 에스컬레이션 적용. 단가행은 그 매출의 역산치.
          var sumTerms = IN_ADDR.tariffTrack.map(function (ta) {
            return IN + ta.share + '/100*' + IN + ta.price + '*(1+' + IN + ta.escal + '/100)^' + p.opYearIdx;
          }).join('+');
          putF(ws, pc(n) + 12, pc(n) + '9*(' + sumTerms + ')/1000', FMT_M, { bold: true });
          putF(ws, pc(n) + 10, 'IF(' + pc(n) + '9=0,0,' + pc(n) + '12*1000/' + pc(n) + '9)', '#,##0.0');
        } else {
          putF(ws, pc(n) + 10,
            IN + IN_ADDR.tariff + '*(1+' + IN + IN_ADDR.tariffEscal + '/100)^' + p.opYearIdx, '#,##0.0');
          putF(ws, pc(n) + 12, pc(n) + '9*' + pc(n) + '10/1000', FMT_M, { bold: true });
        }
      }
      putF(ws, 'D12', sumFormula(12), FMT_M, { bold: true });

      if (anyOverride) {
        var note = 14;
        ws.getCell('B' + note).value = '※ 실측 오버라이드가 적용된 분기는 "판매단가"가 그 분기의 실제 정산단가(실측 사실, baked)이고, "영업수익"은 다른 분기와 동일하게 발전량×판매단가/1000 수식으로 계산됩니다(값은 원본과 그대로 일치).';
        ws.getCell('B' + note).font = { name: FONT, size: 8, italic: true, color: { argb: 'FF9AA6A1' } };
      }

      /* =========================================================
         매출채권(A/R) 회수 시차 → 운전자본 증감
         전기를 판 분기와 대금이 들어오는 분기가 다르다(PPA·SMP 통상 1개월,
         REC는 발급·거래 절차가 붙어 2개월). 분기 안에서 매출이 균등하다고
         보면 기말 미수금 = "마지막 lag개월치 매출"이고, 운전자본 증감은
         그 잔액의 감소분이다: wc[n] = AR[n-1] - AR[n].
         운영 마지막 분기엔 남은 미수금을 전액 회수하므로 AR = 0.
         ========================================================= */
      var arLagAny = (inp.tariffTracks || []).some(function (t) { return (t.arLagMonths || 0) > 0; }) ||
        (inp.arLagMonths || 0) > 0;
      if (arLagAny) {
        var perM = 12 / (inp.ppy || 4);
        var lastOpR = -1;
        periods.forEach(function (p, i) { if (p.isOp) lastOpR = i; });
        var arStart = 16;
        section(ws, arStart, '매출채권(A/R) 회수 시차'); arStart += 2;
        var trackRevRow = {}, arBalRow = {};
        var rr = arStart;
        if (hasTracks) {
          IN_ADDR.tariffTrack.forEach(function (ta, ti) {
            var nm = (inp.tariffTracks[ti].name) || ('트랙' + (ti + 1));
            trackRevRow[ti] = rr;
            label(ws, rr, nm + ' 매출', '[KRWm]', { indent: true });
            for (var n = 0; n < N; n++) {
              if (!periods[n].isOp) { put(ws, pc(n) + rr, 0, FMT_M); continue; }
              putF(ws, pc(n) + rr,
                pc(n) + '9*' + IN + ta.share + '/100*' + IN + ta.price +
                '*(1+' + IN + ta.escal + '/100)^' + periods[n].opYearIdx + '/1000', FMT_M);
            }
            putF(ws, 'D' + rr, sumFormula(rr), FMT_M);
            rr++;
          });
        } else {
          trackRevRow[0] = 12;   // 트랙이 없으면 총매출 행을 그대로 쓴다
        }
        var trackIdx = hasTracks ? IN_ADDR.tariffTrack.map(function (_, i) { return i; }) : [0];
        trackIdx.forEach(function (ti) {
          var nm = hasTracks ? ((inp.tariffTracks[ti].name) || ('트랙' + (ti + 1))) : '전체';
          arBalRow[ti] = rr;
          label(ws, rr, nm + ' 기말 미수금', '[KRWm]', { indent: true });
          var lagRef = hasTracks ? (IN + IN_ADDR.tariffTrack[ti].arLag) : String(inp.arLagMonths || 0);
          for (var n = 0; n < N; n++) {
            if (n >= lastOpR) { put(ws, pc(n) + rr, 0, FMT_M, { noSum: true }); continue; }
            // lag가 분기 길이를 넘을 수 있으므로 직전 분기 매출까지 함께 본다.
            var cur = pc(n) + trackRevRow[ti];
            var prv = n > 0 ? (pc(n - 1) + trackRevRow[ti]) : '0';
            putF(ws, pc(n) + rr,
              'MIN(' + lagRef + ',' + perM + ')/' + perM + '*' + cur +
              '+MAX(' + lagRef + '-' + perM + ',0)/' + perM + '*' + prv, FMT_M, { noSum: true });
          }
          rr++;
        });
        AR_WC_ROW = rr;
        label(ws, rr, '운전자본 증감(A/R)', '[KRWm]', { bold: true, fill: SUB_FILL });
        for (var n = 0; n < N; n++) {
          var terms = trackIdx.map(function (ti) {
            var prev = n > 0 ? (pc(n - 1) + arBalRow[ti]) : '0';
            return prev + '-' + pc(n) + arBalRow[ti];
          }).join('+');
          putF(ws, pc(n) + rr, terms, FMT_M, { bold: true });
        }
        putF(ws, 'D' + rr, sumFormula(rr), FMT_M, { bold: true });
        rr++;
        ws.getCell('B' + rr).value = '※ 20년 전체 합계는 0입니다(초기에 묶인 미수금을 마지막에 전액 회수). 다만 분기별로는 크게 움직여서 준공 직후 DSCR과 Equity IRR에 영향을 줍니다.';
        ws.getCell('B' + rr).font = { name: FONT, size: 8, italic: true, color: { argb: 'FF9AA6A1' } };
      }
    })();

    /* =========================================================
       4. Opex
       ========================================================= */
    (function () {
      var ws = sheet('Opex');
      title(ws, '영업비용 추정');
      var r = 4;
      section(ws, r, '운영비용 (선순위/후순위 지급순위 반영)'); r += 2;
      periodHeader(ws, r); r += 2;

      // IS(Q)와 같은 항목별 세부내역 — 화면에서 항목별로 입력했으면
      // 실제 금액, 합계만 입력했으면 항목명만(금액 빈칸) 표시.
      var opexRows = inp.opexItems && inp.opexItems.length;
      var opexLabelsOnly = !opexRows && model.opexDisplayItems && model.opexDisplayItems.length;
      var itemRows = OPEX_ITEM_ROWS;
      if (opexRows) {
        label(ws, r, '항목별 세부내역', null, { bold: true }); r++;
        inp.opexItems.forEach(function (it, idx) {
          label(ws, r, it.name || ('항목' + (idx + 1)), '[KRWm]', { indent: true });
          itemRows.push(r);
          for (var n = 0; n < N; n++) {
            if (opexPeriodIsFormulaable(n)) {
              putF(ws, pc(n) + r, '-(' + opexItemFormulaStr(idx, n) + ')', FMT_M);
            } else if (ovrByEnd[periods[n].endStr]) {
              putF(ws, pc(n) + r, '-(' + opexItemOverrideFormula(idx, n) + ')', FMT_M);
            } else {
              put(ws, pc(n) + r, 0, FMT_M);
            }
          }
          putF(ws, 'D' + r, sumFormula(r), FMT_M);
          r++;
        });
        r++;
      } else if (opexLabelsOnly) {
        label(ws, r, '항목별 세부내역', null, { bold: true }); r++;
        model.opexDisplayItems.forEach(function (it) {
          label(ws, r, it.name, '[KRWm]', { indent: true }); r++;
        });
        r++;
      }

      // "합계만 입력"(단순모드, 항목별 opexItems 미사용) — opexEok*에스컬레이션의
      // 단순 지수식이라 오버라이드 없는 분기는 이것도 수식화한다.
      var flatMode = !opexRows && IN_ADDR.opexEok != null;
      var totalRowAddr = r + 2; // 합계 행 주소(선순위/후순위가 그 행을 참조)
      var seniorRow = r;
      OPEX_SENIOR_ROW = r;
      OPEX_SUB_ROW = r + 1; // 후순위운영비 행 — 바로 다음 줄(subRow와 항상 같은 값)
      label(ws, r, '선순위운영비', '[KRWm]');
      if (opexRows && inp.opexItems.some(function (it) { return it.senior !== false; })) {
        var seniorIdx = inp.opexItems.map(function (it, i) { return it.senior !== false ? i : -1; }).filter(function (i) { return i >= 0; });
        for (var n = 0; n < N; n++) {
          if (opexPeriodIsFormulaable(n)) putF(ws, pc(n) + r, seniorIdx.map(function (i) { return '-' + pc(n) + itemRows[i]; }).join('+'), FMT_M);
          else putOpexSeniorOrSub(ws, pc(n) + r, n, 'opexSenior', FMT_M);
        }
        putF(ws, 'D' + r, sumFormula(r), FMT_M);
      } else if (flatMode) {
        for (var n = 0; n < N; n++) {
          if (opexPeriodIsFormulaable(n)) putF(ws, pc(n) + r, pc(n) + totalRowAddr + '-' + pc(n) + (totalRowAddr - 1), FMT_M);
          else putOpexSeniorOrSub(ws, pc(n) + r, n, 'opexSenior', FMT_M);
        }
        putF(ws, 'D' + r, sumFormula(r), FMT_M);
      } else {
        for (var n = 0; n < N; n++) putOpexSeniorOrSub(ws, pc(n) + r, n, 'opexSenior', FMT_M);
        putF(ws, 'D' + r, sumFormula(r), FMT_M);
      }
      r++;
      var subRow = r;
      label(ws, r, '후순위운영비', '[KRWm]');
      if (opexRows && inp.opexItems.some(function (it) { return it.senior === false; })) {
        var subIdx = inp.opexItems.map(function (it, i) { return it.senior === false ? i : -1; }).filter(function (i) { return i >= 0; });
        for (var n = 0; n < N; n++) {
          if (opexPeriodIsFormulaable(n)) putF(ws, pc(n) + r, subIdx.map(function (i) { return '-' + pc(n) + itemRows[i]; }).join('+'), FMT_M);
          else putOpexSeniorOrSub(ws, pc(n) + r, n, 'opexSub', FMT_M);
        }
        putF(ws, 'D' + r, sumFormula(r), FMT_M);
      } else if (flatMode) {
        for (var n = 0; n < N; n++) {
          if (opexPeriodIsFormulaable(n)) {
            putF(ws, pc(n) + r, pc(n) + totalRowAddr + '*' + IN + IN_ADDR.opexSubShare + '/100', FMT_M);
          } else putOpexSeniorOrSub(ws, pc(n) + r, n, 'opexSub', FMT_M);
        }
        putF(ws, 'D' + r, sumFormula(r), FMT_M);
      } else {
        for (var n = 0; n < N; n++) putOpexSeniorOrSub(ws, pc(n) + r, n, 'opexSub', FMT_M);
        putF(ws, 'D' + r, sumFormula(r), FMT_M);
      }
      r++;
      // 합계 행은 **모드에 따라 계산 방향이 반대**라 한쪽으로 통일하면 안 된다:
      //  - 항목별 입력(opexRows): 항목들 → 선순위/후순위 → 합계 (아래에서 위로)
      //  - 단순모드(flatMode)   : 합계(총액×에스컬) → 선순위/후순위 (위에서 아래로)
      // 예전에 합계를 무조건 "선순위+후순위"로 통일했다가 단순모드에서
      // 선순위=합계-후순위, 후순위=합계×비중, 합계=선순위+후순위가 서로를
      // 물고 도는 순환참조가 생겼다 — 모드별로 나눠서 채운다.
      OPEX_TOTAL_ROW = r;
      label(ws, r, '영업비용 합계', '[KRWm]', { bold: true, fill: SUB_FILL });
      for (var n = 0; n < N; n++) {
        var p = periods[n];
        if (flatMode && opexPeriodIsFormulaable(n)) {
          putF(ws, pc(n) + r,
            IN + IN_ADDR.opexEok + '*100*(1+' + IN + IN_ADDR.opexEscal + '/100)^' + p.opYearIdx + '*(' + p.opMonths + '/12)',
            FMT_M, { bold: true });
        } else {
          putF(ws, pc(n) + r, pc(n) + seniorRow + '+' + pc(n) + subRow, FMT_M, { bold: true });
        }
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M, { bold: true }); r++;

      if (opexLabelsOnly) {
        r++;
        ws.getCell('B' + r).value = '※ 화면에서 운영비를 합계로만 입력해서 항목별 금액은 비어 있습니다 — 항목별로 입력하면 이 표에 실제 금액이 채워집니다.';
        ws.getCell('B' + r).font = { name: FONT, size: 8, italic: true, color: { argb: 'FF9AA6A1' } };
      } else if (opexRows) {
        r++;
        ws.getCell('B' + r).value = '※ 항목별 실제값이 없는 분기(오버라이드 포함)는 실제 합계("입력값" 시트 참조)를 공식 기준 비중으로 비례 배분한 수식입니다 — 합계 자체는 항상 실제값과 정확히 일치하고, 항목 간 배분만 근사치입니다.';
        ws.getCell('B' + r).font = { name: FONT, size: 8, italic: true, color: { argb: 'FF9AA6A1' } };
      } else if (flatMode) {
        r++;
        ws.getCell('B' + r).value = '※ 실측 오버라이드가 적용된 분기(있다면)는 값(baked)이고, 그 외 분기는 "입력값" 시트의 운영비 총액/에스컬레이션/후순위 비중을 참조하는 수식입니다.';
        ws.getCell('B' + r).font = { name: FONT, size: 8, italic: true, color: { argb: 'FF9AA6A1' } };
      }
    })();

    /* =========================================================
       5. IS(Q) — 손익계산서
       ========================================================= */
    (function () {
      var ws = sheet('IS(Q)');
      title(ws, '추정 손익계산서 (분기)');
      var r = 4;
      section(ws, r, '손익계산서'); r += 2;
      periodHeader(ws, r); r += 2;

      // 새로 계산하지 않고 Revenue 시트를 그대로 참조한다(Revenue 시트의
      // 12행 "영업수익"이 이미 값/수식을 다 가지고 있음 — 로직 중복 방지).
      var revRowIS = r;
      label(ws, r, '영업수익', '[KRWm]', { bold: true });
      for (var n = 0; n < N; n++) putF(ws, pc(n) + r, "'Revenue'!" + pc(n) + '12', FMT_M, { bold: true });
      putF(ws, 'D' + r, sumFormula(r), FMT_M, { bold: true });
      r += 2;

      // 새로 계산하지 않고 Opex 시트를 그대로 참조한다(항목별 로직이 이미
      // Opex 시트에 있음 — 두 곳에서 같은 계산을 유지하는 위험을 없앤다).
      var opexRows = inp.opexItems && inp.opexItems.length;
      var opexLabelsOnly = !opexRows && model.opexDisplayItems && model.opexDisplayItems.length;
      if (opexRows) {
        label(ws, r, '영업비용 세부내역', null, { bold: true }); r++;
        inp.opexItems.forEach(function (it, idx) {
          label(ws, r, it.name || ('항목' + (idx + 1)), '[KRWm]', { indent: true });
          var opexRowAddr = OPEX_ITEM_ROWS[idx];
          for (var n = 0; n < N; n++) putF(ws, pc(n) + r, "'Opex'!" + pc(n) + opexRowAddr, FMT_M);
          putF(ws, 'D' + r, sumFormula(r), FMT_M);
          r++;
        });
      } else if (opexLabelsOnly) {
        // 합계만 입력한 경우 — 항목 이름만 보여주고 금액칸은 비워둔다
        // (화면에서 항목별 금액을 안 넣었으니 추정치를 임의로 채우지 않음).
        label(ws, r, '영업비용 세부내역', null, { bold: true }); r++;
        model.opexDisplayItems.forEach(function (it) {
          label(ws, r, it.name, '[KRWm]', { indent: true }); r++;
        });
      }
      var opexTotalRowIS = r;
      label(ws, r, '영업비용 합계', '[KRWm]', { bold: true });
      for (var n = 0; n < N; n++) putF(ws, pc(n) + r, "-'Opex'!" + pc(n) + OPEX_TOTAL_ROW, FMT_M, { bold: true });
      putF(ws, 'D' + r, sumFormula(r), FMT_M, { bold: true }); r++;
      r++;
      var ebitdaRow = r;
      label(ws, r, 'EBITDA', '[KRWm]', { bold: true, fill: SUB_FILL });
      for (var n = 0; n < N; n++) putF(ws, pc(n) + r, pc(n) + revRowIS + '+' + pc(n) + opexTotalRowIS, FMT_M, { bold: true });
      putF(ws, 'D' + r, sumFormula(r), FMT_M, { bold: true }); r++;
      label(ws, r, 'EBITDA 마진', '[%]');
      for (var n = 0; n < N; n++) {
        putF(ws, pc(n) + r, 'IF(' + pc(n) + revRowIS + '<=0,"",' + pc(n) + ebitdaRow + '/' + pc(n) + revRowIS + ')', FMT_P, { noSum: true });
      }
      r++;
      r++;
      var depRow = r;
      label(ws, r, '감가상각비', '[KRWm]');
      // 상각대상액/내용연수 기준 정액법 — 감가상각 판정(어느 분기가 상각
      // 대상인지)은 이 시점의 내용연수를 기준으로 정해지는 구조적 사실이라
      // 값을 바꿔도 어느 열이 대상인지는 안 바뀐다(그 안에서는 정확히 수식).
      var depBaseF = inp.depBaseOverride != null
        ? IN + IN_ADDR.depBaseOverride
        : "'Funding'!" + FUNDING_TIC_ADDR + '*' + IN + IN_ADDR.depRatio + '/100';
      for (var n = 0; n < N; n++) {
        var p = periods[n];
        if (p.isOp && p.opYearIdx < inp.depYears) {
          putF(ws, pc(n) + r, '-((' + depBaseF + ')/' + IN + IN_ADDR.depYears + '*(' + p.opMonths + '/12))', FMT_M);
        } else put(ws, pc(n) + r, 0, FMT_M);
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M); r++;

      var decomRow = r;
      label(ws, r, '복구충당부채 전입액', '[KRWm]');
      var opPeriodCountIS = periods.filter(function (p) { return p.isOp; }).length;
      for (var n = 0; n < N; n++) {
        if (ovrByEnd[periods[n].endStr] && IN_ADDR.ovr && IN_ADDR.ovr.decomAccrual != null) {
          // 오버라이드 분기도 실측치를 "입력값" 시트에 두고 참조한다(다른
          // 오버라이드 항목들과 동일한 취급 — 예전엔 여기만 값이 박혀 있었음).
          putF(ws, pc(n) + r, '-' + IN + pc(n) + IN_ADDR.ovr.decomAccrual, FMT_M);
        } else if (isOpNonOverride(n)) {
          putF(ws, pc(n) + r, '-(' + IN + IN_ADDR.decomEok + '*100/' + opPeriodCountIS + ')', FMT_M);
        } else put(ws, pc(n) + r, -(rows[n].decomAccrual || 0), FMT_M);
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M); r++;

      // "기타 비현금 세무손금"(선납임대료 상각액 등)과 "대리은행수수료"는
      // 원본 순서대로 EBIT 위(EBITDA와 EBIT 사이)에서 빠진다 — 예전엔
      // 대리은행수수료를 이자비용 밑(EBT 바로 위)에 표시해서 실제 계산
      // 순서(EBIT 자체에 이미 반영됨)와 화면 배치가 어긋나 있었다.
      var extraTaxDedRow = r;
      label(ws, r, '기타 비현금 세무손금', '[KRWm]');
      for (var n = 0; n < N; n++) {
        if (isOpNonOverride(n)) {
          putF(ws, pc(n) + r, '-(' + IN + IN_ADDR.extraTaxDeductionKRWm + '/' + opPeriodCountIS + ')', FMT_M);
        } else if (ovrByEnd[periods[n].endStr]) {
          putF(ws, pc(n) + r, '-' + IN + pc(n) + IN_ADDR.ovr.extraTaxDed, FMT_M);
        } else {
          put(ws, pc(n) + r, 0, FMT_M);
        }
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M); r++;

      var agentFeeRow = r;
      ISQ_AGENTFEE_ROW = r;
      label(ws, r, '대리은행수수료', '[KRWm]');
      for (var n = 0; n < N; n++) {
        if (isOpNonOverride(n)) {
          putF(ws, pc(n) + r, '-(' + IN + IN_ADDR.agentFeeKRWm + '/' + opPeriodCountIS + ')', FMT_M);
        } else if (ovrByEnd[periods[n].endStr]) {
          putF(ws, pc(n) + r, '-' + IN + pc(n) + IN_ADDR.ovr.agentFee, FMT_M);
        } else {
          put(ws, pc(n) + r, 0, FMT_M);
        }
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M); r++;

      var ebitRow = r;
      label(ws, r, '영업이익(EBIT)', '[KRWm]', { bold: true });
      for (var n = 0; n < N; n++) {
        putF(ws, pc(n) + r,
          [ebitdaRow, depRow, decomRow, extraTaxDedRow, agentFeeRow].map(function (rr) { return pc(n) + rr; }).join('+'),
          FMT_M, { bold: true });
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M, { bold: true }); r++;
      label(ws, r, '영업이익률', '[%]');
      for (var n = 0; n < N; n++) {
        putF(ws, pc(n) + r, 'IF(' + pc(n) + revRowIS + '<=0,"",' + pc(n) + ebitRow + '/' + pc(n) + revRowIS + ')', FMT_P, { noSum: true });
      }
      r++;
      r++;

      var intTotalRowIS;
      if (trancheInterest.length) {
        // 새로 계산하지 않고 Debt 시트의 트랜치별 이자 행을 그대로 참조한다
        // (Debt 시트 자체가 이미 "입력값"을 참조하는 수식이므로, 여기서
        // 또 계산하면 같은 로직이 두 곳에 흩어져 유지보수 위험만 커진다).
        label(ws, r, '이자비용 세부내역', null, { bold: true }); r++;
        var intDetailRows = [];
        trancheInterest.forEach(function (ti, tii) {
          label(ws, r, ti.name + ' 이자', '[KRWm]', { indent: true });
          var blk = DEBT_TR_BLOCKS[tii];
          if (blk) {
            for (var n = 0; n < N; n++) putF(ws, pc(n) + r, "-'Debt'!" + pc(n) + blk.intRow, FMT_M);
            putF(ws, 'D' + r, sumFormula(r), FMT_M);
            intDetailRows.push(r);
          } else {
            fillPeriods(ws, r, function (n) { return -(ti.ints[n] || 0); }, FMT_M);
          }
          r++;
        });
        intTotalRowIS = r;
        label(ws, r, '이자비용 합계', '[KRWm]', { bold: true });
        // 위 세부내역 행들의 단순 합 — Debt 전체합계를 또 참조하지 않고
        // 바로 위에 펼쳐놓은 값들을 더해서 표 안에서 검산이 되게 한다.
        for (var n = 0; n < N; n++) {
          putF(ws, pc(n) + r,
            intDetailRows.length ? intDetailRows.map(function (rr) { return pc(n) + rr; }).join('+') : ("-'Debt'!" + pc(n) + DEBT_INT_TOTAL_ROW),
            FMT_M, { bold: true });
        }
        putF(ws, 'D' + r, sumFormula(r), FMT_M, { bold: true }); r++;
      } else {
        intTotalRowIS = r;
        label(ws, r, '이자비용', '[KRWm]');
        for (var n = 0; n < N; n++) putF(ws, pc(n) + r, "-'Debt'!" + pc(n) + DEBT_INT_TOTAL_ROW, FMT_M);
        putF(ws, 'D' + r, sumFormula(r), FMT_M); r++;
      }
      r++;
      var ebtRow = r;
      label(ws, r, '법인세차감전순이익', '[KRWm]', { bold: true });
      for (var n = 0; n < N; n++) putF(ws, pc(n) + r, pc(n) + ebitRow + '+' + pc(n) + intTotalRowIS, FMT_M, { bold: true });
      putF(ws, 'D' + r, sumFormula(r), FMT_M, { bold: true }); r++;
      var taxRow = r;
      ISQ_TAX_ROW = r;
      label(ws, r, '법인세비용', '[KRWm]'); r++; // 데이터는 연도별 산출 표를 만든 뒤 아래서 채운다
      ISQ_NI_ROW = r;
      label(ws, r, '당기순이익', '[KRWm]', { bold: true, fill: SUB_FILL });
      for (var n = 0; n < N; n++) putF(ws, pc(n) + r, pc(n) + ebtRow + '+' + pc(n) + taxRow, FMT_M, { bold: true });
      putF(ws, 'D' + r, sumFormula(r), FMT_M, { bold: true }); r++;

      /* =========================================================
         연도별 법인세 산출 — 이월결손금·통합투자세액공제(10년 이월)·
         최저한세(AMT)를 연도 단위로 체인 수식화한다(Debt 시트의
         "기초잔액=직전 기말잔액"과 같은 방식 — 반복계산이 아니라 그냥
         직전 열을 참조하는 체인이라 수식화 가능). 분기별 법인세는 그
         연도 법인세를 "그 분기 양수EBT / 그 해 양수EBT합" 비중으로
         비례배분한다(오버라이드 없는 세계에서는 이 자체가 정확한 값).
         ========================================================= */
      r += 2;
      var taxSectionTitleRow = r;
      section(ws, r, '연도별 법인세 산출'); r += 2;
      // 분기 -> 연도 매핑, 연도 목록(구간 전체 + preOpLossByYear/투자세액공제
      // 연도까지 포함해 JS 엔진의 Object.keys(yrEBT)와 동일하게 구성)
      var yearSet = {};
      periods.forEach(function (p) { yearSet[p.year] = true; });
      Object.keys(inp.preOpLossByYear || {}).forEach(function (y) { yearSet[y] = true; });
      Object.keys(inp.investmentCreditBaseByYear || {}).forEach(function (y) { yearSet[y] = true; });
      var years = Object.keys(yearSet).map(Number).sort(function (a, b) { return a - b; });
      var lastOpIdx = -1;
      periods.forEach(function (p, i) { if (p.isOp) lastOpIdx = i; });

      // 분기별 "세무상 과세표준" 행 — 회계상 EBT에 복구충당부채(비현금,
      // 세무상 손금 아님)를 되돌리고, 실제 철거비 현금지출(마지막
      // 운영분기에만)만 손금 반영한다.
      var taxableRow = r;
      label(ws, r, '세무상 과세표준(분기, 연간집계용)', '[KRWm]');
      for (var n = 0; n < N; n++) {
        var cashDecomF = (n === lastOpIdx) ? IN + IN_ADDR.decomEok + '*100' : '0';
        putF(ws, pc(n) + r, pc(n) + ebtRow + '-' + pc(n) + decomRow + '-(' + cashDecomF + ')', FMT_M);
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M); r += 2;

      var yc = C0; // 연도 열은 분기 열과 같은 시작열(E)을 재사용 — 아래쪽 별도 구간
      var yearCol = {}; // year -> column letter
      years.forEach(function (y, yi) { yearCol[y] = colLetter(yc + yi); });

      function yearLabelRow(text) { label(ws, r, text, null); }
      // 연도 헤더
      label(ws, r, '연도', null);
      years.forEach(function (y) { put(ws, yearCol[y] + r, y, '0', { bold: true }); });
      r++;
      var taxableYrRow = r;
      label(ws, r, '과세표준(연간, 손실이월 전)', '[KRWm]');
      years.forEach(function (y) {
        var qCols = [];
        periods.forEach(function (p, n) { if (p.year === y) qCols.push(pc(n) + taxableRow); });
        var lossAddr = (IN_ADDR.preOpLossByYear && IN_ADDR.preOpLossByYear[y]) ? ('-' + IN + IN_ADDR.preOpLossByYear[y]) : '';
        putF(ws, yearCol[y] + r, qCols.join('+') + lossAddr, FMT_M);
      });
      r++;
      var carryPrevRow = r; label(ws, r, '직전 이월결손금(기초)', '[KRWm]'); r++;
      var dedRow = r; label(ws, r, '당해 결손금 공제액', '[KRWm]'); r++;
      var carryNewRow = r; label(ws, r, '이월결손금(기말)', '[KRWm]'); r++;
      var baseRow = r; label(ws, r, '과세표준(공제후)', '[KRWm]'); r++;
      var grossTaxRow = r; label(ws, r, '산출세액', '[KRWm]'); r++;
      var amtRow = r; label(ws, r, '최저한세', '[KRWm]'); r++;
      var creditPoolPrevRow = r; label(ws, r, '직전 세액공제 잔액(기초)', '[KRWm]'); r++;
      var newCreditRow = r; label(ws, r, '당해 신규 세액공제', '[KRWm]'); r++;
      var afterCreditRow = r; label(ws, r, '공제후 세액', '[KRWm]'); r++;
      var taxFinalRow = r; label(ws, r, '최종 산출세액(MAX)', '[KRWm]'); r++;
      var creditUsedRow = r; label(ws, r, '세액공제 사용액', '[KRWm]'); r++;
      var creditPoolNewRow = r; label(ws, r, '세액공제 잔액(기말)', '[KRWm]'); r++;
      var taxByYearRow = r; label(ws, r, '법인세비용(연간)', '[KRWm]', { bold: true }); r++;
      var yrPosRow = r; label(ws, r, '양수 EBT 합(연간, 분기배분용)', '[KRWm]'); r++;

      var grossTaxF = inp.taxMode === 1
        ? function (baseExpr) {
          return 'MIN(MAX(' + baseExpr + ',0),200)*0.09' +
            '+MIN(MAX(' + baseExpr + '-200,0),19800)*0.19' +
            '+MIN(MAX(' + baseExpr + '-20000,0),280000)*0.21' +
            '+MAX(' + baseExpr + '-300000,0)*0.24';
        }
        : function (baseExpr) { return '(' + baseExpr + ')*' + IN + IN_ADDR.taxFlat + '/100'; };

      years.forEach(function (y, yi) {
        var c = yearCol[y];
        var prevC = yi > 0 ? yearCol[years[yi - 1]] : null;
        putF(ws, c + carryPrevRow, prevC ? (prevC + carryNewRow) : '0', FMT_M);
        putF(ws, c + dedRow, 'IF(' + c + taxableYrRow + '<=0,0,MIN(' + c + carryPrevRow + ',' + c + taxableYrRow + '*' + IN + IN_ADDR.lossRate + '/100))', FMT_M);
        putF(ws, c + carryNewRow, 'IF(' + c + taxableYrRow + '<=0,' + c + carryPrevRow + '-' + c + taxableYrRow + ',' + c + carryPrevRow + '-' + c + dedRow + ')', FMT_M);
        putF(ws, c + baseRow, 'MAX(0,' + c + taxableYrRow + '-' + c + dedRow + ')', FMT_M);
        putF(ws, c + grossTaxRow, 'IF(' + c + taxableYrRow + '<=0,0,' + grossTaxF(c + baseRow) + ')', FMT_M);
        putF(ws, c + amtRow, c + baseRow + '*' + IN + IN_ADDR.amtRate + '/100', FMT_M);
        putF(ws, c + creditPoolPrevRow, prevC ? (prevC + creditPoolNewRow) : '0', FMT_M);
        var creditBaseAddr = (IN_ADDR.investmentCreditBaseByYear && IN_ADDR.investmentCreditBaseByYear[y]) ? (IN + IN_ADDR.investmentCreditBaseByYear[y]) : '0';
        putF(ws, c + newCreditRow, '(' + creditBaseAddr + ')*' + IN + IN_ADDR.investmentCreditRate + '/100', FMT_M);
        putF(ws, c + afterCreditRow, 'MAX(0,' + c + grossTaxRow + '-(' + c + creditPoolPrevRow + '+' + c + newCreditRow + '))', FMT_M);
        putF(ws, c + taxFinalRow, 'IF(' + c + taxableYrRow + '<=0,0,MAX(' + c + amtRow + ',' + c + afterCreditRow + '))', FMT_M);
        putF(ws, c + creditUsedRow, c + grossTaxRow + '-' + c + taxFinalRow, FMT_M);
        putF(ws, c + creditPoolNewRow, c + creditPoolPrevRow + '+' + c + newCreditRow + '-' + c + creditUsedRow, FMT_M);
        putF(ws, c + taxByYearRow, c + taxFinalRow + '*(1+' + IN + IN_ADDR.localSurtaxRate + '/100)+' + c + creditUsedRow + '*' + IN + IN_ADDR.creditSurtaxRate + '/100', FMT_M, { bold: true });
        var posTerms = [];
        periods.forEach(function (p, n) { if (p.year === y) posTerms.push('MAX(0,' + pc(n) + ebtRow + ')'); });
        putF(ws, c + yrPosRow, posTerms.join('+') || '0', FMT_M);
      });
      r += 1;
      ws.getCell('B' + r).value = '※ 이월결손금(공제한도 이내)·세액공제(10년 이월, 통합투자세액공제)·최저한세(MAX)를 원본과 같은 순서로 연도 단위 수식 체인으로 계산합니다(반복계산 아님 — Debt 시트의 "기초잔액=직전 기말잔액"과 같은 방식). 분기별 법인세비용은 이 표의 연간 법인세를 그 분기 양수EBT 비중으로 비례배분합니다.';
      ws.getCell('B' + r).font = { name: FONT, size: 8, italic: true, color: { argb: 'FF9AA6A1' } };

      // 이제 분기별 "법인세비용" 행을 채운다(연도 표를 참조하는 수식).
      for (var n = 0; n < N; n++) {
        var p = periods[n];
        var yc2 = yearCol[p.year];
        putF(ws, pc(n) + taxRow,
          'IF(AND(' + pc(n) + ebtRow + '>0,' + yc2 + yrPosRow + '>0),-' + yc2 + taxByYearRow + '*' + pc(n) + ebtRow + '/' + yc2 + yrPosRow + ',0)',
          FMT_M);
      }
      putF(ws, 'D' + taxRow, sumFormula(taxRow), FMT_M);

      if (opexRows) {
        r++;
        ws.getCell('B' + r).value = '※ 영업비용 세부내역은 Opex 시트를 그대로 참조합니다 — 항목별 실제값이 없는 분기(오버라이드 포함)는 실제 합계를 공식 기준 비중으로 비례 배분한 수식이고, 합계 자체는 검증된 실제값과 정확히 일치합니다.';
        ws.getCell('B' + r).font = { name: FONT, size: 8, italic: true, color: { argb: 'FF9AA6A1' } };
      } else if (opexLabelsOnly) {
        r++;
        ws.getCell('B' + r).value = '※ 화면에서 운영비를 합계로만 입력해서 항목별 금액은 비어 있습니다 — 항목별로 입력하면 이 표에 실제 금액이 채워집니다.';
        ws.getCell('B' + r).font = { name: FONT, size: 8, italic: true, color: { argb: 'FF9AA6A1' } };
      }
    })();

    /* =========================================================
       6. CF(Q) — 현금흐름 / DSCR / 배당
       ========================================================= */
    (function () {
      var ws = sheet('CF(Q)');
      title(ws, '추정 현금흐름표 (분기) · DSCR · 배당');
      section(ws, 4, '영업/재무 현금흐름');
      periodHeader(ws, 6);
      // CFADS = 영업수익 - 선순위운영비 + 법인세비용(이미 음수) +
      // 대리은행수수료(이미 음수) + 운전자본증감(오버라이드 분기만 실측,
      // 그 외엔 항상 0 — 엔진 자체가 일반 입력 경로에서 wc를 안 씀).
      // 운전자본 증감 — 실측 오버라이드가 있으면 그 값을, 없으면 Revenue
      // 시트에서 A/R 회수 시차로 계산한 행을 참조한다(둘 다 없으면 0).
      var wcTerm = function (n) {
        var ovr = ovrByEnd[periods[n].endStr];
        if (ovr && ovr.wc != null) return '+' + IN + pc(n) + IN_ADDR.ovr.wc;
        if (AR_WC_ROW) return "+'Revenue'!" + pc(n) + AR_WC_ROW;
        return '';
      };
      label(ws, 9, 'CFADS (원리금상환재원)', '[KRWm]', { bold: true, fill: SUB_FILL });
      for (var n = 0; n < N; n++) {
        putF(ws, pc(n) + 9,
          "'Revenue'!" + pc(n) + '12-\'Opex\'!' + pc(n) + OPEX_SENIOR_ROW +
          "+'IS(Q)'!" + pc(n) + ISQ_TAX_ROW + "+'IS(Q)'!" + pc(n) + ISQ_AGENTFEE_ROW + wcTerm(n),
          FMT_M, { bold: true });
      }
      putF(ws, 'D9', sumFormula(9), FMT_M, { bold: true });
      label(ws, 10, '원리금(DS)', '[KRWm]');
      for (var n = 0; n < N; n++) putF(ws, pc(n) + 10, "'Debt'!" + pc(n) + DEBT_DS_ROW, FMT_M);
      putF(ws, 'D10', sumFormula(10), FMT_M);
      label(ws, 11, '단순 DSCR', '[x]', { bold: true });
      // CFADS(9행)÷원리금(10행) — 둘 다 이미 이 시트 안에 있으니 순수 비율
      // 수식으로 연결(원리금이 0에 가까우면 원본과 동일하게 공란 처리).
      for (var n = 0; n < N; n++) {
        if (rows[n].ds > 1e-9) putF(ws, pc(n) + 11, pc(n) + '9/' + pc(n) + '10', FMT_X, { bold: true });
        else put(ws, pc(n) + 11, null, FMT_X, { bold: true });
      }
      // "누적DSCR"은 실제로는 분기 누적이 아니라 연 단위 지표다(Report!row233
      // "(A+B)/C" — A=그 해 첫 분기 기초현금, B=그 해 첫 분기 DSRA 기초잔액,
      // C=그 해 CFADS합, D=그 해 원리금합 → (A+B+C)/D). "단순DSCR(연 합산)"도
      // 마찬가지로 그 해 CFADS합/원리금합이지 분기별 비율이 아니다. 아래
      // "검증" 섹션에 연도별 요약 표를 별도로 만들어서 계산한다.

      label(ws, 13, 'FCFE', '[KRWm]', { bold: true });
      // FCFE = CFADS - DS - DSRA조정(원본: 부호 반대로 뺀다, 오버라이드
      // 있으면 실측 dsraFcfe, 없으면 DSRA증감을 그대로) - 후순위운영비 -
      // 철거비현금유출(마지막 운영분기만) + 세금 현금조정(taxCash 오버라이드
      // 있을 때만, 발생주의 법인세를 실제 납부액으로 되돌림).
      var lastOpIdx2 = -1;
      periods.forEach(function (p, i) { if (p.isOp) lastOpIdx2 = i; });
      var dsraMoveRow = 14, dsraOpenRow = 20, dsraCloseRow = 21;
      var nqDsra = Math.round((inp.dsraMonths || 0) / (12 / (inp.ppy || 4)));
      var codIdx = model.con.codIdx;
      label(ws, dsraOpenRow, '(DSRA 기초잔액, 내부용)', '[KRWm]', { noHide: true });
      label(ws, dsraCloseRow, '(DSRA 기말잔액, 내부용)', '[KRWm]', { noHide: true });
      for (var n = 0; n < N; n++) {
        var prevCol = n > 0 ? pc(n - 1) : null;
        if (n === codIdx) {
          putF(ws, pc(n) + dsraOpenRow, IN + IN_ADDR.dsraEok + '*100', FMT_M, { noSum: true });
        } else {
          putF(ws, pc(n) + dsraOpenRow, prevCol ? (prevCol + dsraCloseRow) : '0', FMT_M, { noSum: true });
        }
        if (periods[n].isOp) {
          var needTerms = [];
          for (var k = 1; k <= nqDsra && n + k < N; k++) needTerms.push(pc(n + k) + '10');
          putF(ws, pc(n) + dsraCloseRow, needTerms.length ? needTerms.join('+') : '0', FMT_M, { noSum: true });
        } else {
          putF(ws, pc(n) + dsraCloseRow, pc(n) + dsraOpenRow, FMT_M, { noSum: true });
        }
      }
      label(ws, dsraMoveRow, 'DSRA 증감', '[KRWm]');
      for (var n = 0; n < N; n++) putF(ws, pc(n) + dsraMoveRow, pc(n) + dsraCloseRow + '-' + pc(n) + dsraOpenRow, FMT_M);
      putF(ws, 'D' + dsraMoveRow, sumFormula(dsraMoveRow), FMT_M);

      // 세금 현금조정(taxAdj) — 오버라이드 분기 중 실제 납부액(taxCash)이 있는
      // 분기만 발생주의 법인세를 실제 현금 납부액으로 되돌린다. FCFE와
      // "그 해 원리금상환재원(현금기준)" 둘 다 이 값을 참조하므로 행 하나로
      // 공유한다.
      var taxAdjRow = 19;
      label(ws, taxAdjRow, '(세금 현금조정, 내부용)', '[KRWm]');
      for (var n = 0; n < N; n++) {
        var ovr0 = ovrByEnd[periods[n].endStr];
        if (ovr0 && ovr0.taxCash != null) {
          putF(ws, pc(n) + taxAdjRow, "-'IS(Q)'!" + pc(n) + ISQ_TAX_ROW + '-' + IN + pc(n) + IN_ADDR.ovr.taxCash, FMT_M);
        } else {
          put(ws, pc(n) + taxAdjRow, 0, FMT_M);
        }
      }

      for (var n = 0; n < N; n++) {
        var ovr = ovrByEnd[periods[n].endStr];
        // fcfe -= dsraForFcfe. dsraForFcfe(오버라이드) = -dsraFcfe_실측값이므로
        // "- dsraForFcfe" = "+dsraFcfe_실측값"(부호 반전에 주의).
        var dsraForFcfeF = (ovr && ovr.dsraFcfe != null) ? ('+' + IN + pc(n) + IN_ADDR.ovr.dsraFcfe) : ('-' + pc(n) + dsraMoveRow);
        var decomF = (n === lastOpIdx2) ? ('-' + IN + IN_ADDR.decomEok + '*100') : '';
        putF(ws, pc(n) + 13,
          pc(n) + '9-' + pc(n) + '10' + dsraForFcfeF + "-'Opex'!" + pc(n) + OPEX_SUB_ROW + decomF + '+' + pc(n) + taxAdjRow,
          FMT_M, { bold: true });
      }
      putF(ws, 'D13', sumFormula(13), FMT_M, { bold: true });

      // ---- 가용현금(avail) = 기초현금 + FCFE ----
      var availRow = 18;
      label(ws, availRow, '(가용현금=기초현금+FCFE, 내부용)', '[KRWm]');
      for (var n = 0; n < N; n++) putF(ws, pc(n) + availRow, pc(n) + '15+' + pc(n) + '13', FMT_M);

      // ==========================================================
      // 배당 게이트 — distributable(배당가능이익 누적)/reserveBalance(이익
      // 준비금 누적)/pendingDiv(결의~지급 이연)을 분기 단위 체인 수식으로
      // 연결한다. 실제 "얼마를 배당할지" 결정(decided)과 이익준비금 적립액
      // (reserveNeed)은 연 1회(12월, decideMonth) 계산되는데, 그 계산엔 그
      // 해 전체 분기 데이터가 필요해서 아래 "연도별 배당 게이트" 표에서
      // 연도 단위로 구하고, 분기 체인은 그 표의 결과를 결의월에만 반영한다.
      // ==========================================================
      var hasDivMonth = inp.dividendMonth != null;
      var decideMonth = hasDivMonth ? (inp.dividendMonth === 3 ? 12 : inp.dividendMonth) : null;
      var isPayQ = periods.map(function (p) { return hasDivMonth && (p.end.getUTCMonth() + 1) === inp.dividendMonth; });
      var isDecideQ = periods.map(function (p) { return hasDivMonth && p.isOp && (p.end.getUTCMonth() + 1) === decideMonth; });
      var yrFirstOpIdx = {};
      periods.forEach(function (p, n) { if (p.isOp && yrFirstOpIdx[p.year] === undefined) yrFirstOpIdx[p.year] = n; });
      var payQtrOpIdx = {};
      periods.forEach(function (p, n) { if (isPayQ[n] && p.isOp) payQtrOpIdx[p.year] = n; });

      // 연도별 배당 게이트 표 — 행 번호를 먼저 정해서(아래서 실제로 채움)
      // 분기 체인이 그 주소를 미리 참조할 수 있게 한다.
      var g_yrDS = 27, g_yrCF = 28, g_annualDscr = 29, g_annualCumDscr = 30,
        g_yrCFCash = 31, g_yrReserve = 32, g_yrCashStart = 33, g_yrPostMarch = 34,
        g_pSimpleOK = 35, g_pCumOK = 36, g_canDecide = 37, g_maxByProfit = 38,
        g_maxByCash = 39, g_yearlyAvail = 40, g_maxByDscrReserve = 41,
        g_decided = 42, g_reserveNeed = 43;
      var divYearSet = {};
      periods.forEach(function (p, n) { if (isDecideQ[n]) divYearSet[p.year] = true; });
      var dsYearSet2 = {};
      periods.forEach(function (p, n) { if (rows[n].ds > 1e-9) dsYearSet2[p.year] = true; });
      var allYearSet = {};
      Object.keys(divYearSet).forEach(function (y) { allYearSet[y] = true; });
      Object.keys(dsYearSet2).forEach(function (y) { allYearSet[y] = true; });
      var allYears = Object.keys(allYearSet).map(Number).sort(function (a, b) { return a - b; });
      var gYearCol = {};
      allYears.forEach(function (y, yi) { gYearCol[y] = colLetter(C0 + yi); });

      var distributableRow = 22, reserveBalanceRow = 23, pendingDivRow = 24;
      label(ws, distributableRow, '(배당가능이익 누적, 내부용)', '[KRWm]');
      label(ws, reserveBalanceRow, '(이익준비금 누적, 내부용)', '[KRWm]');
      label(ws, pendingDivRow, '(결의 대기 배당금, 내부용)', '[KRWm]');
      for (var n = 0; n < N; n++) {
        var p = periods[n];
        var prevCol = n > 0 ? pc(n - 1) : null;
        var gc = hasDivMonth && isDecideQ[n] ? gYearCol[p.year] : null;
        // distributable += (그 분기 isOp면 NI) - (결의월이면 decided+reserveNeed)
        var distTerms = [];
        if (prevCol) distTerms.push(prevCol + distributableRow);
        if (p.isOp) distTerms.push("'IS(Q)'!" + pc(n) + ISQ_NI_ROW);
        var distFormula = distTerms.length ? distTerms.join('+') : '0';
        if (gc) distFormula += '-(' + gc + g_decided + '+' + gc + g_reserveNeed + ')';
        putF(ws, pc(n) + distributableRow, distFormula, FMT_M, { noSum: true });
        // reserveBalance += (결의월이면 reserveNeed)
        var resFormula = prevCol ? (prevCol + reserveBalanceRow) : '0';
        if (gc) resFormula += '+' + gc + g_reserveNeed;
        putF(ws, pc(n) + reserveBalanceRow, resFormula, FMT_M, { noSum: true });
        // pendingDiv: 결의월이면 decided로 교체, 지급월이면 0으로 리셋,
        // 그 외에는 직전 값을 그대로 이어간다.
        var pendF;
        if (gc) pendF = gc + g_decided;
        else if (hasDivMonth && isPayQ[n]) pendF = '0';
        else pendF = prevCol ? (prevCol + pendingDivRow) : '0';
        putF(ws, pc(n) + pendingDivRow, pendF, FMT_M, { noSum: true });
      }

      // 일반 입력 경로(오버라이드/dividendMonth 없음) 전용 — "그 즉시 스윕"
      // 근사에 쓰는 분기 누적DSCR 체인(운영개시부터, Debt식 직전열 참조).
      // dividendMonth가 있는 정밀 경로(당진 프리셋)에서는 안 쓴다.
      var cumCfadsRow = 46, cumDsRow = 47, cumDscrRow2 = 48;
      if (!hasDivMonth) {
        label(ws, cumCfadsRow, '(누적 CFADS, 내부용)', '[KRWm]');
        label(ws, cumDsRow, '(누적 DS, 내부용)', '[KRWm]');
        label(ws, cumDscrRow2, '(누적 DSCR, 내부용)', '[x]');
        for (var n = 0; n < N; n++) {
          var prevCol = n > 0 ? pc(n - 1) : null;
          if (periods[n].isOp) {
            putF(ws, pc(n) + cumCfadsRow, (prevCol ? prevCol + cumCfadsRow + '+' : '') + pc(n) + '9', FMT_M, { noSum: true });
            putF(ws, pc(n) + cumDsRow, (prevCol ? prevCol + cumDsRow + '+' : '') + pc(n) + '10', FMT_M, { noSum: true });
            putF(ws, pc(n) + cumDscrRow2, 'IF(' + pc(n) + cumDsRow + '>0,' + pc(n) + cumCfadsRow + '/' + pc(n) + cumDsRow + ',"")', FMT_X, { noSum: true });
          } else {
            putF(ws, pc(n) + cumCfadsRow, prevCol ? (prevCol + cumCfadsRow) : '0', FMT_M, { noSum: true });
            putF(ws, pc(n) + cumDsRow, prevCol ? (prevCol + cumDsRow) : '0', FMT_M, { noSum: true });
            put(ws, pc(n) + cumDscrRow2, null, FMT_X, { noSum: true });
          }
        }
      }

      label(ws, 15, '기초현금', '[KRWm]');
      for (var n = 0; n < N; n++) putF(ws, pc(n) + 15, n > 0 ? (pc(n - 1) + '17') : '0', FMT_M, { noSum: true });
      label(ws, 16, '배당(연차+청산)', '[KRWm]', { bold: true, fill: SUB_FILL });
      for (var n = 0; n < N; n++) {
        var divF;
        if (n === lastOpIdx2) {
          divF = pc(n) + availRow; // 청산배당: 한도 없이 잔여현금 전액
        } else if (hasDivMonth) {
          // 결의월≠지급월 전제(표준 12월결의/3월지급 구조) — 지급월엔 직전
          // 분기까지 쌓인 결의액을 그대로 지급, 그 외엔 0.
          divF = (isPayQ[n] && n > 0) ? (pc(n - 1) + pendingDivRow) : '0';
        } else {
          // 오버라이드 없는 일반 입력 경로 — 즉시 스윕형 근사(12월결의/3월지급
          // 타이밍 없이, DSCR 게이트만 보고 그 즉시 최소보유현금 초과분을 배당).
          if (periods[n].isOp) {
            var dscrOK = 'OR(ISBLANK(' + pc(n) + '11),' + pc(n) + '11>=' + IN + IN_ADDR.divDSCR + ')';
            var cumOK = 'OR(ISBLANK(' + pc(n) + cumDscrRow2 + '),' + pc(n) + cumDscrRow2 + '>=' + IN + IN_ADDR.divCumDSCRVal + ')';
            var startOK = periods[n].opYearIdx + '>=(' + IN + IN_ADDR.divStartYear + '-1)';
            divF = 'IF(AND(' + startOK + ',' + dscrOK + ',' + cumOK + '),MAX(0,' + pc(n) + availRow + '-' + IN + IN_ADDR.minCash + '*100),0)';
          } else {
            divF = '0';
          }
        }
        putF(ws, pc(n) + 16, divF, FMT_M, { bold: true });
      }
      putF(ws, 'D16', sumFormula(16), FMT_M, { bold: true });
      label(ws, 17, '기말현금', '[KRWm]');
      for (var n = 0; n < N; n++) putF(ws, pc(n) + 17, pc(n) + availRow + '-' + pc(n) + '16', FMT_M, { noSum: true });

      // 연도별 DSCR·배당 게이트 요약 — "단순DSCR(연 합산)"/"누적DSCR"은 분기
      // 비율이 아니라 연 단위 지표(Report!row233 방식)이고, 배당 결의(decided)/
      // 이익준비금 적립(reserveNeed)도 그 해 전체 분기 데이터가 있어야 계산
      //되는 연 단위 값이라 같은 표에 함께 둔다. 위쪽 분기 체인(distributable/
      // reserveBalance/pendingDiv)은 이 표의 decided/reserveNeed를 결의월에만
      // 참조해서 반영한다.
      var yrFirstIdx = {};
      periods.forEach(function (p, n) { if (yrFirstIdx[p.year] === undefined) yrFirstIdx[p.year] = n; });
      var decQtrIdx = {};
      periods.forEach(function (p, n) { if (isDecideQ[n]) decQtrIdx[p.year] = n; });
      var reserveCapF = IN + IN_ADDR.equityEok + '*100*0.5';
      var r = 24;
      section(ws, r, '연도별 DSCR · 배당 게이트'); r += 2;
      label(ws, r, '연도', null);
      allYears.forEach(function (y) { put(ws, gYearCol[y] + r, y, '0', { bold: true }); });
      r++;
      label(ws, g_yrDS, '원리금 합(연간)', '[KRWm]');
      label(ws, g_yrCF, 'CFADS 합(연간)', '[KRWm]');
      label(ws, g_annualDscr, '단순 DSCR(연 합산)', '[x]');
      label(ws, g_annualCumDscr, '누적 DSCR(연간)', '[x]');
      label(ws, g_yrCFCash, 'CFADS 합(현금기준, 연간)', '[KRWm]');
      label(ws, g_yrReserve, '그 해 첫 운영분기 기초현금+DSRA', '[KRWm]');
      label(ws, g_yrCashStart, '그 해 첫 운영분기 기초현금', '[KRWm]');
      label(ws, g_yrPostMarch, '그 해 지급월 기초현금-배당', '[KRWm]');
      label(ws, g_pSimpleOK, '단순DSCR 게이트 통과', '[bool]');
      label(ws, g_pCumOK, '누적DSCR 게이트 통과', '[bool]');
      label(ws, g_canDecide, '배당 결의 가능', '[bool]');
      label(ws, g_maxByProfit, '배당가능이익 한도', '[KRWm]');
      label(ws, g_maxByCash, '현금 한도', '[KRWm]');
      label(ws, g_yearlyAvail, '원리금상환재원 기준선', '[KRWm]');
      label(ws, g_maxByDscrReserve, '원리금상환재원 한도', '[KRWm]');
      label(ws, g_decided, '결정 배당액', '[KRWm]');
      label(ws, g_reserveNeed, '이익준비금 적립액', '[KRWm]');
      r = 44;
      allYears.forEach(function (y) {
        var c = gYearCol[y];
        var dsTerms = [], cfTerms = [], cfCashTerms = [];
        periods.forEach(function (p, n) {
          if (p.year === y && rows[n].ds > 1e-9) {
            dsTerms.push(pc(n) + '10'); cfTerms.push(pc(n) + '9'); cfCashTerms.push(pc(n) + '9+' + pc(n) + taxAdjRow);
          }
        });
        putF(ws, c + g_yrDS, dsTerms.length ? dsTerms.join('+') : '0', FMT_M);
        putF(ws, c + g_yrCF, cfTerms.length ? cfTerms.join('+') : '0', FMT_M);
        putF(ws, c + g_annualDscr, 'IF(' + c + g_yrDS + '=0,"",' + c + g_yrCF + '/' + c + g_yrDS + ')', FMT_X);
        var r0 = yrFirstIdx[y];
        putF(ws, c + g_annualCumDscr, 'IF(' + c + g_yrDS + '=0,"",(' + pc(r0) + '15+' + pc(r0) + dsraOpenRow + '+' + c + g_yrCF + ')/' + c + g_yrDS + ')', FMT_X);
        if (!hasDivMonth || !divYearSet[y]) return; // 배당 게이트는 정밀(dividendMonth) 경로·결의연도만
        putF(ws, c + g_yrCFCash, cfCashTerms.length ? cfCashTerms.join('+') : '0', FMT_M);
        var r0op = yrFirstOpIdx[y];
        putF(ws, c + g_yrReserve, pc(r0op) + '15+' + pc(r0op) + dsraOpenRow, FMT_M);
        putF(ws, c + g_yrCashStart, pc(r0op) + '15', FMT_M);
        var pq = payQtrOpIdx[y];
        if (pq !== undefined) putF(ws, c + g_yrPostMarch, pc(pq) + '15-' + pc(pq) + '16', FMT_M);
        putF(ws, c + g_pSimpleOK, 'IF(' + c + g_yrDS + '=0,TRUE,' + c + g_annualDscr + '>=' + IN + IN_ADDR.divDSCR + ')', '@');
        putF(ws, c + g_pCumOK, 'IF(' + c + g_yrDS + '=0,TRUE,' + c + g_annualCumDscr + '>=' + IN + IN_ADDR.divCumDSCRVal + ')', '@');
        var firstDivOK = (inp.firstDividendYear == null || (y + 1) >= inp.firstDividendYear) ? 'TRUE' : 'FALSE';
        putF(ws, c + g_canDecide, 'AND(' + firstDivOK + ',' + c + g_pSimpleOK + ',' + c + g_pCumOK + ')', '@');
        var dq = decQtrIdx[y];
        // 주의: distributableRow[dq] 셀 자체는 이미 이번 분기의 decided/
        // reserveNeed 차감이 반영된 "사후" 값이라(그 값이 바로 이 g_decided를
        // 참조해서 계산되므로) 여기서 그 셀을 그대로 쓰면 순환참조가 된다.
        // JS 원본도 "차감 전" distributable(직전 잔액+이번 분기 NI)을 쓰므로
        // 그 값을 직접 다시 조립한다.
        var distBeforeThisQ = (dq > 0 ? pc(dq - 1) + distributableRow : '0') + "+'IS(Q)'!" + pc(dq) + ISQ_NI_ROW;
        putF(ws, c + g_maxByProfit, '(' + distBeforeThisQ + ')/1.1', FMT_M);
        putF(ws, c + g_maxByCash, pc(dq) + availRow + '-' + IN + IN_ADDR.minCash + '*100', FMT_M);
        // yearlyAvail = (그 해 3월 지급분기가 있으면 그 시점 값, 없으면
        // 그 해 첫 운영분기 기초현금) + 그 해 현금기준 CFADS합 — 어느 쪽이든
        // yrCFCash는 항상 더해진다(원본 수식 그대로).
        putF(ws, c + g_yearlyAvail, (pq !== undefined ? (c + g_yrPostMarch) : (c + g_yrCashStart)) + '+' + c + g_yrCFCash, FMT_M);
        putF(ws, c + g_maxByDscrReserve, 'MAX(0,' + c + g_yearlyAvail + '-' + c + g_yrDS + '*' + IN + IN_ADDR.divCumDSCRVal + ')', FMT_M);
        putF(ws, c + g_decided, 'IF(' + c + g_canDecide + ',MAX(0,MIN(MIN(' + c + g_maxByCash + ',' + c + g_maxByProfit + '),' + c + g_maxByDscrReserve + ')),0)', FMT_M);
        // reserveBalanceRow[dq]도 마찬가지로 이미 이번 분기 reserveNeed가
        //반영된 사후값이라 순환참조가 된다 — 차감(적립) 전 값(직전 열)을 쓴다.
        var reserveBefore = dq > 0 ? pc(dq - 1) + reserveBalanceRow : '0';
        putF(ws, c + g_reserveNeed, 'IF(' + c + g_canDecide + ',MAX(0,MIN(' + c + g_decided + '*0.10,' + reserveCapF + '-' + reserveBefore + ')),0)', FMT_M);
      });
      r += 1;
      section(ws, r, '검증'); r += 2;
      var dsYears2 = allYears.filter(function (y) { return dsYearSet2[y]; });
      CFQ_MINDSCR_ROW = r;
      label(ws, r, '최소 단순DSCR(연 합산)', '[x]', { bold: true });
      putF(ws, 'D' + r, 'MIN(' + gYearCol[dsYears2[0]] + g_annualDscr + ':' + gYearCol[dsYears2[dsYears2.length - 1]] + g_annualDscr + ')', FMT_X, { bold: true }); r++;
      CFQ_MINCUMDSCR_ROW = r;
      label(ws, r, '최소 누적DSCR', '[x]', { bold: true });
      putF(ws, 'D' + r, 'MIN(' + gYearCol[dsYears2[0]] + g_annualCumDscr + ':' + gYearCol[dsYears2[dsYears2.length - 1]] + g_annualCumDscr + ')', FMT_X, { bold: true }); r++;
      label(ws, r, '최종 기말현금(음수면 오류)', '[KRWm]', { bold: true });
      putF(ws, 'D' + r, lastC + '17', FMT_M, { bold: true });

      /* =========================================================
         IRR용 분기별 현금흐름 계열 — Report의 IRR들이 참조한다. 부호가
         자주 바뀌는 분기 현금흐름이라 Excel IRR()의 기본 탐색(10% 근방
         뉴턴법)이 엉뚱한 근(수백만% 등)에 수렴할 위험이 있다(JS의 irr()가
         "0% 좌우로 가장 가까운 부호전환 구간"을 스캔하도록 만든 이유와
         동일) — 그래서 IRR() 두 번째 인자(guess)에 JS가 이미 찾아둔 정답을
         그대로 넣어 그 근처에서 시작하게 한다. 입력값을 크게 바꾸면 guess가
         낡아서 다른 근으로 수렴할 수 있으니, 결과가 비정상적으로 크면
         (수백% 이상) 화면(HTML) 값과 대조할 것.
         ========================================================= */
      r += 2;
      section(ws, r, 'IRR용 현금흐름 (참고)'); r += 2;
      periodHeader(ws, r); r += 2;
      // 건설기간 공사비 유출·자본금 인출 모두 이제 라이브 — Debt 시트의
      // "건설기간 자금조달" 섹션(누적인출)과 입력값 시트의 지출 스케줄을
      // 그대로 참조한다.
      var capOutF = function (n) {
        return IN_ADDR.spendCurveRow != null
          ? (IN + pc(n) + IN_ADDR.spendCurveRow + '*' + IN + IN_ADDR.capexEok + '*100')
          : '0';
      };
      var equityDrawRow = r;
      label(ws, r, '자본금 인출', '[KRWm]');
      for (var n = 0; n < N; n++) {
        var cd0 = DEBT_CUMDRAW_ROW[0];
        putF(ws, pc(n) + r, "'Debt'!" + pc(n) + cd0 + (n > 0 ? ("-'Debt'!" + pc(n - 1) + cd0) : ''), FMT_M);
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M); r++;
      CFQ_PROJFLOW_ROW = r;
      label(ws, r, 'Project 현금흐름', '[KRWm]');
      for (var n = 0; n < N; n++) {
        var decomCashF = (n === lastOpIdx2) ? ('-' + IN + IN_ADDR.decomEok + '*100') : '';
        // projectFcf = EBITDA(=매출-영업비용합계) - 법인세 - 철거비현금(마지막
        // 운영분기만) - 대리은행수수료 + 운전자본 - 건설기간 유출.
        putF(ws, pc(n) + r,
          "'Revenue'!" + pc(n) + "12-'Opex'!" + pc(n) + OPEX_TOTAL_ROW +
          "+'IS(Q)'!" + pc(n) + ISQ_TAX_ROW + decomCashF +
          "+'IS(Q)'!" + pc(n) + ISQ_AGENTFEE_ROW + wcTerm(n) + '-(' + capOutF(n) + ')',
          FMT_M);
      }
      r++;
      CFQ_PROJFLOWPRE_ROW = r;
      label(ws, r, 'Project 현금흐름(세전)', '[KRWm]');
      for (var n = 0; n < N; n++) {
        var decomCashF2 = (n === lastOpIdx2) ? ('-' + IN + IN_ADDR.decomEok + '*100') : '';
        // preFlows = EBITDA - 철거비현금 - 대리은행수수료 + 운전자본 -
        // 건설기간 유출 (법인세만 세후 버전과 다름 — 안 뺀다).
        putF(ws, pc(n) + r,
          "'Revenue'!" + pc(n) + "12-'Opex'!" + pc(n) + OPEX_TOTAL_ROW +
          decomCashF2 + "+'IS(Q)'!" + pc(n) + ISQ_AGENTFEE_ROW + wcTerm(n) + '-(' + capOutF(n) + ')',
          FMT_M);
      }
      r++;
      CFQ_EQFLOW_ROW = r;
      label(ws, r, 'Equity 현금흐름(FCFE 기준)', '[KRWm]');
      for (var n = 0; n < N; n++) putF(ws, pc(n) + r, pc(n) + '13-' + pc(n) + equityDrawRow, FMT_M);
      r++;
      CFQ_DIVFLOW_ROW = r;
      label(ws, r, 'Equity 현금흐름(배당 기준)', '[KRWm]');
      for (var n = 0; n < N; n++) putF(ws, pc(n) + r, pc(n) + '16-' + pc(n) + equityDrawRow, FMT_M);
      r++;
      CFQ_INVFLOW_ROW = r;
      label(ws, r, 'Investor 현금흐름', '[KRWm]');
      for (var n = 0; n < N; n++) {
        // 부채 인출 합계 = 트랜치별 (이번 누적인출 - 직전 누적인출)의 합.
        var debtDrawF = model.tranches.map(function (t, ti) {
          var cdr = DEBT_CUMDRAW_ROW[ti + 1];
          return "'Debt'!" + pc(n) + cdr + (n > 0 ? ("-'Debt'!" + pc(n - 1) + cdr) : '');
        }).join('+') || '0';
        var idcTermF = (n === model.con.codIdx) ? ("+'Debt'!D" + DEBT_IDC_TOTAL_ROW) : '';
        putF(ws, pc(n) + r,
          '-' + pc(n) + equityDrawRow + '-(' + debtDrawF + ")+'Debt'!" + pc(n) + DEBT_INT_TOTAL_ROW + "+'Debt'!" + pc(n) + DEBT_PRIN_TOTAL_ROW + '+' + pc(n) + '16' + idcTermF,
          FMT_M);
      }
      r += 2;
      var lastColRef = lastC;
      var ppyIrr = inp.ppy || 4;
      // IRR()은 분기 현금흐름 그대로 넣으면 "분기 기준" 내부수익률을 돌려준다
      // (연 환산 아님) — JS의 annualize()와 동일하게 (1+분기IRR)^ppy-1로
      // 연 환산해야 model.kpi 값과 맞는다. guess도 연 단위가 아니라 분기
      // 단위로 환산해서 넣어야 한다 — 분기 현금흐름은 배당처럼 부호가 자주
      // 바뀌어 IRR()의 뉴턴법이 guess 스케일이 안 맞으면(연율을 그대로
      // 넣으면) 엉뚱한 근으로 튀는 걸 실제로 확인했다(Equity/배당 IRR에서
      // 재현). guess를 분기 스케일로 정확히 맞추면 JS의 스캔 방식과 동일한
      // 근으로 수렴한다.
      function irrRow(label2, flowRow, annualGuess) {
        var qGuess = Math.pow(1 + (annualGuess || 0.08), 1 / ppyIrr) - 1;
        label(ws, r, label2, '[%]', { bold: true });
        putF(ws, 'D' + r, '(1+IRR(' + firstC + flowRow + ':' + lastColRef + flowRow + ',' + qGuess + '))^' + ppyIrr + '-1', FMT_P, { bold: true });
        var rr = r; r++;
        return rr;
      }
      CFQ_PROJIRRPRE_ROW = irrRow('Project IRR(세전)', CFQ_PROJFLOWPRE_ROW, model.kpi.projectIRRPre || 0.09);
      CFQ_PROJIRR_ROW = irrRow('Project IRR(세후)', CFQ_PROJFLOW_ROW, model.kpi.projectIRR || 0.08);
      CFQ_EQIRR_ROW = irrRow('Equity IRR(FCFE)', CFQ_EQFLOW_ROW, model.kpi.equityIRR || 0.1);
      CFQ_DIVIRR_ROW = irrRow('Equity IRR(배당)', CFQ_DIVFLOW_ROW, model.kpi.dividendIRR || 0.1);
      CFQ_INVIRR_ROW = irrRow('Investor IRR', CFQ_INVFLOW_ROW, model.kpi.investorIRR || 0.06);
      r++;
      ws.getCell('B' + r).value = '※ IRR()의 두 번째 인자(초기 추정치)에 현재 계산된 값을 넣어 엉뚱한 근으로 수렴하는 걸 방지했습니다. 입력값을 크게 바꿔서 결과가 비정상적으로 크게(수백% 이상) 나오면, 그 근처의 추정치로 다시 만들어야 정확합니다 — 화면(HTML) 값과 대조해서 확인하세요.';
      ws.getCell('B' + r).font = { name: FONT, size: 8, italic: true, color: { argb: 'FF9AA6A1' } };
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
      function kvF(name, formulaStr, fmt) {
        ws.getCell('B' + r).value = name;
        ws.getCell('B' + r).font = { name: FONT, size: 10 };
        putF(ws, 'D' + r, formulaStr, fmt, { bold: true });
        r++;
      }
      section(ws, r, '사업 개요'); r += 2;
      kvF('사업명', IN + IN_ADDR.projectName, '@');
      kvF('설비용량 [MW]', IN + IN_ADDR.capacityMW, '#,##0.000');
      kv('총 기간 수(분기)', N, '0');
      kv('건설 개시', inp.constructionStart, '@');
      r++;
      section(ws, r, '재원조달'); r += 2;
      kvF('총투자비(TIC) [KRWm]', "'Funding'!" + FUNDING_TIC_ADDR, FMT_M);
      kvF('  건설이자(IDC) [KRWm]', "'Funding'!" + FUNDING_IDC_ADDR, FMT_M);
      kvF('자기자본 [KRWm]', "'Funding'!" + FUNDING_EQUITY_ADDR, FMT_M);
      kvF('차입금 합계 [KRWm]', "'Funding'!" + FUNDING_DEBT_ADDR, FMT_M);
      r++;
      section(ws, r, '수익성 지표'); r += 2;
      kvF('Project IRR 세전 [%]', "'CF(Q)'!D" + CFQ_PROJIRRPRE_ROW, FMT_P);
      kvF('Project IRR 세후 [%]', "'CF(Q)'!D" + CFQ_PROJIRR_ROW, FMT_P);
      kvF('Equity IRR (FCFE) [%]', "'CF(Q)'!D" + CFQ_EQIRR_ROW, FMT_P);
      kvF('Equity IRR (배당) [%]', "'CF(Q)'!D" + CFQ_DIVIRR_ROW, FMT_P);
      kvF('Investor IRR [%]', "'CF(Q)'!D" + CFQ_INVIRR_ROW, FMT_P);
      kvF('최소 단순DSCR(연 합산) [x]', "'CF(Q)'!D" + CFQ_MINDSCR_ROW, FMT_X);
      kvF('최소 누적DSCR [x]', "'CF(Q)'!D" + CFQ_MINCUMDSCR_ROW, FMT_X);
      kvF('총영업수익(전체기간) [KRWm]', "'Revenue'!D12", FMT_M);
      kvF('총영업비용(전체기간) [KRWm]', "'Opex'!D" + OPEX_TOTAL_ROW, FMT_M);
      kvF('총선순위이자 [KRWm]', "'Debt'!D" + DEBT_INT_TOTAL_ROW, FMT_M);
      kvF('총법인세 [KRWm]', "-'IS(Q)'!D" + ISQ_TAX_ROW, FMT_M);
      kvF('총배당(연차+청산) [KRWm]', "'CF(Q)'!D16", FMT_M);
    })();

    /* =========================================================
       7-1. 민감도 — 화면에서 실행한 시나리오 비교 결과(있을 때만)
       ========================================================= */
    if (model.sensitivity && model.sensitivity.length) {
      (function () {
        var ws = wb.addWorksheet('민감도', { properties: { tabColor: { argb: 'FF14483A' } } });
        ws.getColumn(1).width = 2.5; ws.getColumn(2).width = 18;
        for (var ci = 0; ci < 10; ci++) ws.getColumn(3 + ci).width = 14;
        title(ws, '민감도 분석 — 시나리오별 핵심 지표 비교');
        var r = 4;
        ws.getCell('B' + r).value = '판매단가/총투자비/운영비/금리는 델타가 아니라 절대값 — 빈 칸이면 그 시나리오는 "사업 기본 가정"에 입력한 값을 그대로 쓴다. Base(현재 입력값)는 시나리오를 하나도 안 바꾼 기준선. 화면(생성기)에서 지정한 시나리오를 각각 독립적으로 재계산한 값 — 라이브 수식이 아니라 스냅샷임.';
        ws.getCell('B' + r).font = { name: FONT, size: 9, italic: true, color: { argb: 'FF6B7B76' } };
        r += 2;
        var heads = ['시나리오', '판매단가[원/kWh]', '총사업비[억원]', '운영비[억원]', '금리[%]',
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
          var sc = row.sc || {};
          put(ws, 'B' + r, row.name, '@');
          put(ws, 'C' + r, sc.tariffAbs != null ? sc.tariffAbs : '(기본값)', sc.tariffAbs != null ? '#,##0.0' : '@');
          put(ws, 'D' + r, sc.capexAbs != null ? sc.capexAbs : '(기본값)', sc.capexAbs != null ? '#,##0.0' : '@');
          put(ws, 'E' + r, sc.opexAbs != null ? sc.opexAbs : '(기본값)', sc.opexAbs != null ? '#,##0.0' : '@');
          put(ws, 'F' + r, sc.rateAbs != null ? sc.rateAbs : '(기본값)', sc.rateAbs != null ? '0.00' : '@');
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
        ['입력값', '화면에서 입력한 원본 값 — Funding 등이 이 시트를 참조'],
        ['Funding', '자금조달 — 자본금 + 5트랜치 조건'],
        ['Debt', '트랜치별 상환 스케줄 (A/B/C/D/후순위) + 합계'],
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
      ws.getCell('B' + r).value = '※ Funding 시트의 트랜치 조건·총사업비 항목·사업자 지분은 "입력값" 시트를 참조하는 수식입니다(입력값 시트를 고치면 같이 바뀝니다). 그 외 계산 결과(세금·DSCR·배당 등)는 값(baked) 기준이라 가정을 바꾸려면 생성기에서 다시 뽑아야 합니다.';
      ws.getCell('B' + r).font = { name: FONT, size: 9, italic: true, color: { argb: 'FFB4573C' } };
    })();

    // 아직 안 만들어진 시트를 참조해야 했던 셀들을 이제 채운다.
    DEFERRED.forEach(function (fn) { fn(); });

    var order = ['목차', 'Report', '입력값', 'Funding', 'Debt', 'Revenue', 'Opex', 'IS(Q)', 'CF(Q)'].concat(
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
