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
    // Opex 시트가 채운 항목별 행 번호와 "영업비용 합계" 행 번호 — IS(Q)가
    // 재계산하지 않고 그대로 참조하는 데 쓴다.
    var OPEX_ITEM_ROWS = [];
    var OPEX_TOTAL_ROW = null;
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
      r += 1;

      if (inp.tariffTracks && inp.tariffTracks.length) {
        section(ws, r, '판매단가 트랙 (PPA/SMP+REC 등)'); r += 2;
        ['트랙', '비중[%]', '단가[원/kWh]', '에스컬레이션[%/yr]'].forEach(function (h, idx) {
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
          IN_ADDR.tariffTrack.push({ share: 'C' + r, price: 'D' + r, escal: 'E' + r });
          r++;
        });
        r += 1;
      }

      section(ws, r, '트랜치 조건'); r += 2;
      ['트랜치', '금액[억원]', '건설금리[%]', '운영금리[%]', '거치(yr)', '상환(yr)', '방식'].forEach(function (h, idx) {
        var cc = ws.getCell(colLetter(2 + idx) + r);
        cc.value = h; cc.font = { name: FONT, bold: true, size: 9, color: { argb: WHITE } };
        cc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D62' } };
        cc.alignment = { horizontal: 'center' };
      });
      r++;
      model.tranches.forEach(function (t, ti) {
        put(ws, 'B' + r, t.name, '@', { fill: INPUT_FILL });
        var a = { amount: 'C' + r, rateC: 'D' + r, rateO: 'E' + r, grace: 'F' + r, repay: 'G' + r, method: 'H' + r };
        put(ws, a.amount, t.amount / 100, FMT_M, { fill: INPUT_FILL });
        put(ws, a.rateC, t.rateO, FMT_P, { fill: INPUT_FILL });
        put(ws, a.rateO, t.rateO, FMT_P, { fill: INPUT_FILL });
        put(ws, a.grace, t.graceYears, '0.00', { fill: INPUT_FILL });
        put(ws, a.repay, t.repayYears, '0.00', { fill: INPUT_FILL });
        put(ws, a.method, t.method, '0', { fill: INPUT_FILL });
        IN_ADDR.tranche.push(a);
        r++;
      });
      r += 1;

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
      put(ws, 'B' + r, 'IDC 합계[KRWm]'); put(ws, 'C' + r, model.idc, FMT_M); FUNDING_IDC_ADDR = 'C' + r; r++;
      put(ws, 'B' + r, '총투자비(TIC)[KRWm]'); put(ws, 'C' + r, model.tic, FMT_M); FUNDING_TIC_ADDR = 'C' + r; r++;
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
      title(ws, '차입금 상환 스케줄 — 트랜치별 + 합계');
      var r = 4;
      section(ws, r, '트랜치별 인출/상환 + 전체 합계'); r += 2;
      periodHeader(ws, r); r += 2;

      // 방식 1(원금균등)/2(원리금균등)은 화면에서 사용자가 직접 고를 수 있는
      // 방식이라 순수 재무공식(PMT 등)으로 완전히 재현 가능 — 이자/원금/잔액을
      // "입력값" 시트를 참조하는 수식으로 연결한다. 방식 3(64회차 직접 키인,
      // "예시 불러오기" 프리셋 전용, 화면에는 없음)은 스케줄 자체가 원본 실측
      // 데이터라 수식화 대상이 아니라 그대로 값(baked)을 쓴다. 인출/IDC는
      // 지출곡선·건설기간 배분 로직이 얽혀 있어 이번 단계에서는 baked 유지.
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
        label(ws, drawRow, '인출', '[KRWm]');
        fillPeriods(ws, drawRow, function (n) { return t.draws[n] || 0; }, FMT_M);
        label(ws, idcRow, '건설이자(IDC)', '[KRWm]');
        fillPeriods(ws, idcRow, function (n) { return t.idcSeries[n] || 0; }, FMT_M);

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
        var finalBal = closes[N - 1];
        label(ws, r, '미상환 잔액(검증용)', '[KRWm]');
        putF(ws, 'D' + r, lastC + closeRow, FMT_M);
        put(ws, 'F' + r, Math.abs(finalBal) < 1 ? '완전상환 확인 (OK)' : '경고: 미상환 잔액');
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
      fillPeriods(ws, r, function (n) { return rows[n].debtOpen; }, FMT_M, { noSum: true }); r++;
      var intTotalRow = r;
      label(ws, r, '이자', '[KRWm]');
      for (var n = 0; n < N; n++) {
        putF(ws, pc(n) + r, trBlocks.map(function (b) { return pc(n) + b.intRow; }).join('+'), FMT_M);
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M, { bold: true });
      r++;
      var prinTotalRow = r;
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
      label(ws, r, 'DSRA 기말잔액', '[KRWm]');
      fillPeriods(ws, r, function (n) { return rows[n].dsraClose; }, FMT_M, { noSum: true });
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
      // 선순위/후순위 행이 이제 항상 수식이나 "입력값" 참조로 채워져
      // 있으므로(진짜 0인 비운영분기만 예외), 합계는 그냥 그 둘을 더하면
      // 된다 — 여기서 다시 계산할 필요가 없다.
      OPEX_TOTAL_ROW = r;
      label(ws, r, '영업비용 합계', '[KRWm]', { bold: true, fill: SUB_FILL });
      for (var n = 0; n < N; n++) {
        putF(ws, pc(n) + r, pc(n) + seniorRow + '+' + pc(n) + subRow, FMT_M, { bold: true });
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
      label(ws, r, '영업비용 합계', '[KRWm]', { bold: true });
      for (var n = 0; n < N; n++) putF(ws, pc(n) + r, "-'Opex'!" + pc(n) + OPEX_TOTAL_ROW, FMT_M, { bold: true });
      putF(ws, 'D' + r, sumFormula(r), FMT_M, { bold: true }); r++;
      r++;
      label(ws, r, 'EBITDA', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillPeriods(ws, r, function (n) { return rows[n].ebitda; }, FMT_M, { bold: true }); r++;
      label(ws, r, 'EBITDA 마진', '[%]');
      fillPeriods(ws, r, function (n) { return rows[n].revenue > 0 ? rows[n].ebitda / rows[n].revenue : null; }, FMT_P, { noSum: true }); r++;
      r++;
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

      label(ws, r, '복구충당부채 전입액', '[KRWm]');
      var opPeriodCountIS = periods.filter(function (p) { return p.isOp; }).length;
      for (var n = 0; n < N; n++) {
        if (isOpNonOverride(n)) {
          putF(ws, pc(n) + r, '-(' + IN + IN_ADDR.decomEok + '*100/' + opPeriodCountIS + ')', FMT_M);
        } else put(ws, pc(n) + r, -(rows[n].decomAccrual || 0), FMT_M);
      }
      putF(ws, 'D' + r, sumFormula(r), FMT_M); r++;
      label(ws, r, '영업이익(EBIT)', '[KRWm]', { bold: true });
      fillPeriods(ws, r, function (n) { return rows[n].ebit; }, FMT_M, { bold: true }); r++;
      label(ws, r, '영업이익률', '[%]');
      fillPeriods(ws, r, function (n) { return rows[n].revenue > 0 ? rows[n].ebit / rows[n].revenue : null; }, FMT_P, { noSum: true }); r++;
      r++;

      if (trancheInterest.length) {
        // 새로 계산하지 않고 Debt 시트의 트랜치별 이자 행을 그대로 참조한다
        // (Debt 시트 자체가 이미 "입력값"을 참조하는 수식이므로, 여기서
        // 또 계산하면 같은 로직이 두 곳에 흩어져 유지보수 위험만 커진다).
        label(ws, r, '이자비용 세부내역', null, { bold: true }); r++;
        trancheInterest.forEach(function (ti, tii) {
          label(ws, r, ti.name + ' 이자', '[KRWm]', { indent: true });
          var blk = DEBT_TR_BLOCKS[tii];
          if (blk) {
            for (var n = 0; n < N; n++) putF(ws, pc(n) + r, "-'Debt'!" + pc(n) + blk.intRow, FMT_M);
            putF(ws, 'D' + r, sumFormula(r), FMT_M);
          } else {
            fillPeriods(ws, r, function (n) { return -(ti.ints[n] || 0); }, FMT_M);
          }
          r++;
        });
        label(ws, r, '이자비용 합계', '[KRWm]', { bold: true });
        fillPeriods(ws, r, function (n) { return -rows[n].interest; }, FMT_M, { bold: true }); r++;
      } else {
        label(ws, r, '이자비용', '[KRWm]');
        fillPeriods(ws, r, function (n) { return -rows[n].interest; }, FMT_M); r++;
      }
      label(ws, r, '대리은행수수료', '[KRWm]');
      fillPeriods(ws, r, function (n) { return -(rows[n].agentFee || 0); }, FMT_M); r++;
      r++;
      label(ws, r, '법인세차감전순이익', '[KRWm]', { bold: true });
      fillPeriods(ws, r, function (n) { return rows[n].ebt; }, FMT_M, { bold: true }); r++;
      label(ws, r, '법인세비용', '[KRWm]');
      fillPeriods(ws, r, function (n) { return -rows[n].tax; }, FMT_M); r++;
      label(ws, r, '당기순이익', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillPeriods(ws, r, function (n) { return rows[n].ni; }, FMT_M, { bold: true }); r++;

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
      label(ws, 9, 'CFADS (원리금상환재원)', '[KRWm]', { bold: true, fill: SUB_FILL });
      fillPeriods(ws, 9, function (n) { return rows[n].cfads; }, FMT_M, { bold: true });
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
