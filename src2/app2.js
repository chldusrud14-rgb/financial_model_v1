(function () {
  'use strict';
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var el = function (t, c, h) { var e = document.createElement(t); if (c) e.className = c; if (h !== undefined) e.innerHTML = h; return e; };
  var M = window.SolarModel2, X = window.SolarXlsx2;

  /* ---------- 기본 입력 필드 ---------- */
  var CORE = [
    ['projectName', '사업명', 'text', '태양광 발전사업'],
    ['capacityMW', '설비용량', 'number', 99.998, 'MW'],
    ['dailyHours', '일일 발전시간(환산)', 'number', 3.77, 'h/day'],
    ['degradation', 'Degradation', 'number', 0.5, '%/yr'],
    ['auxRate', '소내 소비율', 'number', 0, '%'],
    ['constructionStart', '착공시점(YYYY-MM)', 'text', '2024-06'],
    ['constructionMonths', '공사기간', 'number', 16, 'Month'],
    ['operationYears', '운영기간', 'number', 20, 'Year'],
    ['capexEok', '총사업비(건설이자 제외)', 'number', 1410.69, '억원'],
    ['dsraEok', '최초 DSRA', 'number', 50, '억원'],
    ['equityEok', '자본금', 'number', 150, '억원'],
    ['tariff', '판매단가(가중평균)', 'number', 154.8, '원/kWh'],
    ['tariffEscal', '단가 상승률', 'number', 0, '%/yr'],
    ['opexEok', '운영비(1년차)', 'number', 49.8, '억원/yr'],
    ['opexEscal', '운영비 상승률', 'number', 0.7, '%/yr'],
    ['decomEok', '철거·복구비(만기)', 'number', 20, '억원'],
    ['depRatio', '감가상각 대상비율', 'number', 95, '%'],
    ['depYears', '내용연수', 'number', 20, 'Year'],
    ['lossRate', '이월결손금 공제한도', 'number', 80, '%'],
    ['taxFlat', '단일 법인세율(taxMode=2용)', 'number', 21, '%'],
    ['dsraMonths', 'DSRA 적립기준', 'number', 6, 'Month'],
    ['minCash', '배당 후 최소보유현금', 'number', 10, '억원'],
    ['divDSCR', '배당제한 단순DSCR', 'number', 1.1, 'x'],
    ['divCumDSCR', '배당제한 누적DSCR', 'number', 1.15, 'x'],
    ['divStartYear', '배당개시 연차(폴백용)', 'number', 2, 'Year'],
    ['discount', '할인율(NPV)', 'number', 5.5, '%']
  ];

  var TRANCHES = [
    { key: 'A', name: '선순위A', amountEok: 500, order: 2, rateO: 5.6, graceYears: 2, repayYears: 16 },
    { key: 'B', name: '선순위B', amountEok: 500, order: 2, rateO: 5.5, graceYears: 2, repayYears: 16 },
    { key: 'C', name: '선순위C', amountEok: 350, order: 4, rateO: 5.5, graceYears: 1.25, repayYears: 16 },
    { key: 'D', name: '선순위D', amountEok: 0, order: 3, rateO: 2.25, graceYears: 5, repayYears: 10 },
    { key: 'sub', name: '후순위', amountEok: 0, order: 1, rateO: 5.65, graceYears: 15, repayYears: 2.75 }
  ];

  function buildCore() {
    var wrap = $('#core');
    CORE.forEach(function (d) {
      var wrapf = el('div', 'f' + (d[2] === 'text' ? ' full' : ''));
      wrapf.appendChild(el('label', null, d[1]));
      var inw = el('div', 'in2');
      var input = document.createElement('input');
      input.type = d[2] === 'text' ? 'text' : 'number';
      input.step = 'any';
      input.value = d[3];
      input.dataset.k = d[0];
      inw.appendChild(input);
      if (d[4]) { var u = el('span', 'unit', d[4]); inw.appendChild(u); }
      wrapf.appendChild(inw);
      wrap.appendChild(wrapf);
    });
  }

  function buildTrancheGrid() {
    var box = $('#trbox');
    var t = el('table', 'tr');
    t.innerHTML = '<thead><tr>' +
      '<th>트랜치</th><th>금액(억원)</th><th>투입순서</th><th>건설금리(%)</th><th>운영금리(%)</th>' +
      '<th>거치(yr)</th><th>상환(yr)</th><th>방식</th></tr></thead>';
    var tb = document.createElement('tbody');
    TRANCHES.forEach(function (tr) {
      var row = document.createElement('tr');
      var cells = [
        ['name', tr.name, 'text'],
        ['amountEok', tr.amountEok, 'number'],
        ['order', tr.order, 'number'],
        ['rateC', tr.rateO, 'number'],
        ['rateO', tr.rateO, 'number'],
        ['graceYears', tr.graceYears, 'number'],
        ['repayYears', tr.repayYears, 'number']
      ];
      cells.forEach(function (c, idx) {
        var td = document.createElement('td');
        if (idx === 0) { td.textContent = c[1]; row.appendChild(td); return; }
        var input = document.createElement('input');
        input.type = c[2] === 'text' ? 'text' : 'number';
        input.step = 'any';
        input.value = c[1];
        input.dataset.tr = tr.key; input.dataset.f = c[0];
        td.appendChild(input);
        row.appendChild(td);
      });
  var mtd = document.createElement('td');
      var sel = document.createElement('select');
      sel.innerHTML = '<option value="1">원금균등</option><option value="2">원리금균등</option>';
      sel.dataset.tr = tr.key; sel.dataset.f = 'method';
      mtd.appendChild(sel);
      row.appendChild(mtd);
      tb.appendChild(row);
    });
    t.appendChild(tb);
    box.appendChild(t);
  }

  function buildSpendCurve() {
    var box = $('#spendbox');
    box.innerHTML = '';
    var cm = Number($('[data-k="constructionMonths"]').value) || 16;
    var nq = Math.max(1, Math.ceil(cm / 3));
    var capex = Number($('[data-k="capexEok"]').value) || 0;
    for (var i = 0; i < nq; i++) {
      var f = el('div', 'f');
      f.appendChild(el('label', null, (i + 1) + '분기'));
      var inw = el('div', 'in2');
      var input = document.createElement('input');
      input.type = 'number'; input.step = 'any';
      input.value = (capex / nq).toFixed(2);
      input.dataset.spend = i;
      inw.appendChild(input);
      inw.appendChild(el('span', 'unit', '억원'));
      f.appendChild(inw);
      box.appendChild(f);
    }
  }

  function readCore() {
    var inp = {};
    CORE.forEach(function (d) {
      var v = $('[data-k="' + d[0] + '"]').value;
      inp[d[0]] = d[2] === 'text' ? v : Number(v);
    });
    return inp;
  }
  function readTranches() {
    return TRANCHES.map(function (tr) {
      var o = { name: tr.name, method: Number($('select[data-tr="' + tr.key + '"]').value) };
      ['amountEok', 'order', 'rateC', 'rateO', 'graceYears', 'repayYears'].forEach(function (f) {
        o[f] = Number($('input[data-tr="' + tr.key + '"][data-f="' + f + '"]').value);
      });
      return o;
    });
  }
  function readSpendCurve() {
    return Array.prototype.slice.call(document.querySelectorAll('[data-spend]')).map(function (el2) {
      return Number(el2.value);
    });
  }

  /* ---------- 당진 FS 프리셋 ----------
     test/test_ops.js·scripts/sample2.js와 완전히 동일한 방식으로 reference.json
     (80개 분기 실측 오버라이드 + 세액공제/AMT + 배당 결의 타이밍)을 그대로
     엔진 입력으로 구성한다 — 이게 없으면 화면은 검증된 숫자를 재현하지 못한다. */
  var usingPreset = false, presetInp = null, suppressDirty = false;

  function buildDangjinInp(ref) {
    function findTranche(name) { return ref.tranches.find(function (t) { return t.name === name; }); }
    function scheduleFor(letter) {
      var s = ref.repaySchedule_case3[letter];
      return s ? s.map(function (p) { return p[1]; }) : null;
    }
    var trancheDefs = [['선순위A', 'A'], ['선순위B', 'B'], ['선순위C', 'C'], ['선순위D', null], ['후순위', null]];
    return {
      projectName: '당진 태양광발전', ppy: 4,
      capacityMW: ref.project.capacityMW,
      capacityFactor: ref.project.capacityFactor * 100,
      dailyHours: ref.project.dailyHours,
      degradation: ref.project.degradation * 100,
      auxRate: 0,
      constructionStart: ref.project.constructionStart,
      constructionMonths: ref.project.constructionMonths,
      operationYears: ref.project.operationYears,
      capexEok: ref.funding.TIC_exIDC / 100,
      dsraEok: ref.funding.DSRA / 100,
      opexItems: ref.opexItems.map(function (it) { return { annualKRWm: it.annualKRWm, escal: it.escalRate * 100, senior: it.senior }; }),
      spendCurve: ref.spendCurve_KRWm,
      tariffTracks: ref.tariffTracks.map(function (t) { return { share: t.share, price: t.price }; }),
      seasonality: (function () { var o = {}; Object.keys(ref.seasonality).forEach(function (m) { o[Number(m)] = ref.seasonality[m]; }); return o; })(),
      equityEok: ref.funding.equity / 100, equityOrder: ref.funding.equityOrder,
      tranches: trancheDefs.map(function (d) {
        var name = d[0], letter = d[1];
        var t = findTranche(name);
        return {
          name: t.name, amountEok: t.amount / 100, order: t.order,
          rateC: t.rateCon * 100, rateO: t.rateOp * 100,
          graceYears: t.graceYears, repayYears: t.repayYears, method: t.method,
          repayStart: letter ? t.repayStart : null,
          schedule: letter ? scheduleFor(letter) : null
        };
      }),
      depRatio: 95, depYears: 20, depBaseOverride: ref.depreciableBaseKRWm, taxMode: 1, taxFlat: 21, lossRate: 80,
      investmentCreditRate: ref.taxCredit.investmentCreditRate, amtRate: ref.taxCredit.amtRate,
      localSurtaxRate: ref.taxCredit.localSurtaxRate, creditSurtaxRate: ref.taxCredit.creditSurtaxRate,
      investmentCreditBaseByYear: ref.taxCredit.investmentCreditBaseByYear, preOpLossKRWm: ref.taxCredit.preOpLossKRWm,
      preOpLossByYear: ref.taxCredit.preOpLossByYear,
      extraTaxDeductionKRWm: ref.extraTaxDeductionKRWm, agentFeeKRWm: ref.agentFeeKRWm,
      periodOverrides: ref.periodOverrides,
      decomEok: ref.results.철거비 / 100,
      dsraMonths: 6, minCash: 10, divDSCR: 1.1, divCumDSCR: 1.15, divStartYear: 2, discount: 5.5,
      dividendMonth: ref.dividendMonth, firstDividendYear: ref.firstDividendYear
    };
  }

  function setVal(sel, v) { var e = $(sel); if (e) e.value = v; }

  function loadDangjin() {
    var ref = window.__DANGJIN_REFERENCE__;
    if (!ref) { toast('당진 실측 데이터가 이 빌드에 포함되어 있지 않습니다'); return; }
    suppressDirty = true;
    setVal('[data-k="projectName"]', '당진 태양광발전');
    setVal('[data-k="capacityMW"]', ref.project.capacityMW);
    setVal('[data-k="dailyHours"]', ref.project.dailyHours);
    setVal('[data-k="degradation"]', ref.project.degradation * 100);
    setVal('[data-k="auxRate"]', 0);
    setVal('[data-k="constructionStart"]', ref.project.constructionStart);
    setVal('[data-k="constructionMonths"]', ref.project.constructionMonths);
    setVal('[data-k="operationYears"]', ref.project.operationYears);
    setVal('[data-k="capexEok"]', (ref.funding.TIC_exIDC / 100).toFixed(2));
    setVal('[data-k="dsraEok"]', ref.funding.DSRA / 100);
    setVal('[data-k="equityEok"]', ref.funding.equity / 100);
    setVal('[data-k="tariff"]', ref.project.tariff_wavg);
    setVal('[data-k="tariffEscal"]', 0);
    var totalOpex = ref.opexItems.reduce(function (a, it) { return a + it.annualKRWm; }, 0);
    setVal('[data-k="opexEok"]', (totalOpex / 100).toFixed(2));
    setVal('[data-k="opexEscal"]', 1.2); // 표시용 근사(항목별 실제 값은 프리셋 계산에 그대로 반영됨)
    setVal('[data-k="decomEok"]', ref.results.철거비 / 100);
    setVal('[data-k="depRatio"]', 95); setVal('[data-k="depYears"]', 20);
    setVal('[data-k="lossRate"]', 80); setVal('[data-k="taxFlat"]', 21);
    setVal('[data-k="dsraMonths"]', 6); setVal('[data-k="minCash"]', 10);
    setVal('[data-k="divDSCR"]', 1.1); setVal('[data-k="divCumDSCR"]', 1.15);
    setVal('[data-k="divStartYear"]', 2); setVal('[data-k="discount"]', 5.5);

    buildSpendCurve();
    ref.spendCurve_KRWm.forEach(function (v, i) {
      var e = $('[data-spend="' + i + '"]'); if (e) e.value = (v / 100).toFixed(2);
    });

    var byName = {};
    ref.tranches.forEach(function (t) { byName[t.name] = t; });
    TRANCHES.forEach(function (tr) {
      var t = byName[tr.name];
      if (!t) return;
      setVal('input[data-tr="' + tr.key + '"][data-f="amountEok"]', t.amount / 100);
      setVal('input[data-tr="' + tr.key + '"][data-f="order"]', t.order);
      setVal('input[data-tr="' + tr.key + '"][data-f="rateC"]', (t.rateCon * 100).toFixed(3));
      setVal('input[data-tr="' + tr.key + '"][data-f="rateO"]', (t.rateOp * 100).toFixed(3));
      setVal('input[data-tr="' + tr.key + '"][data-f="graceYears"]', t.graceYears);
      setVal('input[data-tr="' + tr.key + '"][data-f="repayYears"]', t.repayYears);
      var sel = $('select[data-tr="' + tr.key + '"]');
      if (sel && t.method === 3 && !sel.querySelector('option[value="3"]')) {
        sel.appendChild(new Option('64회차(당진 실측)', '3'));
      }
      if (sel) sel.value = String(t.method);
    });

    presetInp = buildDangjinInp(ref);
    usingPreset = true;
    suppressDirty = false;
    toast('당진 FS 실측치를 불러왔습니다 — 이제 "생성"을 누르면 검증된 원본과 동일한 숫자가 나옵니다');
  }

  var model = null;
  var f0 = function (n) { return Math.round(n).toLocaleString('ko-KR'); };
  var pct = function (n) { return n === null || n === undefined || isNaN(n) ? '—' : (n * 100).toFixed(2); };

  function renderKPIs() {
    var k = model.kpi;
    var items = [
      ['Project IRR (세후)', pct(k.projectIRR), '%'],
      ['Equity IRR (FCFE)', pct(k.equityIRR), '%'],
      ['Equity IRR (배당)', pct(k.dividendIRR), '%'],
      ['Investor IRR', pct(k.investorIRR), '%'],
      ['최소 단순DSCR(연 합산)', k.minDSCRAnnual === null ? '—' : k.minDSCRAnnual.toFixed(3), 'x'],
      ['최소 누적DSCR', k.minCumDSCR === null ? '—' : k.minCumDSCR.toFixed(3), 'x'],
      ['총 배당(연차+청산)', f0(k.totalDividend), '억원 기준 아님(KRWm)'],
      ['총투자비(TIC)', f0(model.tic), 'KRWm']
    ];
    var box = $('#kpis'); box.innerHTML = '';
    items.forEach(function (it) {
      var d = el('div', 'kpi');
      d.appendChild(el('div', 'k', it[0]));
      d.appendChild(el('div', 'v', it[1] + '<em>' + it[2] + '</em>'));
      box.appendChild(d);
    });
    $('#metaNote').textContent = '기간 수(분기): ' + model.periods.length +
      ' · 마지막 분기: ' + model.periods[model.periods.length - 1].endStr +
      ' · 총영업수익 ' + f0(k.totalRevenue) + 'KRWm · 총선순위이자 ' + f0(k.totalInterest) + 'KRWm';
  }

  function toast(msg) {
    var t = $('#toast'); t.textContent = msg; t.classList.add('on');
    setTimeout(function () { t.classList.remove('on'); }, 2600);
  }

  function run() {
    var inp;
    if (usingPreset && presetInp) {
      inp = presetInp;
    } else {
      var core = readCore();
      var tranches = readTranches();
      var spendCurve = readSpendCurve();
      inp = Object.assign({}, core, {
        ppy: 4,
        capacityFactor: undefined, // dailyHours 우선 사용
        spendCurve: spendCurve,
        equityOrder: 1,
        tranches: tranches.map(function (t) {
          return {
            name: t.name, amountEok: t.amountEok, order: t.order,
            rateC: t.rateC, rateO: t.rateO, graceYears: t.graceYears, repayYears: t.repayYears,
            method: t.method
          };
        }),
        taxMode: 1,
        localSurtaxRate: 10   // 한국 지방소득세(법인세의 10%)는 기본 적용
      });
    }
    try {
      model = M.computeModel(inp);
      renderKPIs();
      $('#xls').disabled = false;
      toast(usingPreset ? '당진 FS 실측치 기준으로 생성 완료 (원본과 검증된 값)' : '재무모델 생성 완료');
    } catch (e) {
      toast('생성 실패: ' + e.message);
      console.error(e);
    }
  }

  function download() {
    if (!model) return;
    var btn = $('#xls'); btn.disabled = true; btn.textContent = '생성 중…';
    X.buildWorkbook(model, window.ExcelJS).xlsx.writeBuffer().then(function (buf) {
      var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = (model.inp.projectName || '재무모델') + '_v2.xlsx';
      document.body.appendChild(a); a.click(); a.remove();
      btn.disabled = false; btn.textContent = 'Excel 다운로드';
      toast('Excel 다운로드 완료');
    }).catch(function (e) {
      btn.disabled = false; btn.textContent = 'Excel 다운로드';
      toast('다운로드 실패: ' + e.message);
      console.error(e);
    });
  }

  buildCore();
  buildTrancheGrid();
  buildSpendCurve();
  $('[data-k="constructionMonths"]').addEventListener('change', buildSpendCurve);
  $('[data-k="capexEok"]').addEventListener('change', buildSpendCurve);
  $('#run').addEventListener('click', run);
  $('#xls').addEventListener('click', download);
  $('#loadDangjin').addEventListener('click', loadDangjin);

  // 프리셋 로드 후 사용자가 아무 값이나 직접 고치면 더는 "검증된 원본값"이
  // 아니므로 usingPreset을 해제한다(다시 폼을 읽어서 계산하는 일반 경로로).
  // 이때 "64회차(당진 실측)" 방식으로 남아있던 select는 일반 경로에서
  // schedule 데이터가 없어 원금이 영원히 0으로 계산돼 결과가 깨지므로
  // 방식 1(원금균등)로 되돌려준다.
  function dropUnsupportedMethod3() {
    document.querySelectorAll('select[data-tr]').forEach(function (sel) {
      if (sel.value === '3') sel.value = '1';
    });
  }
  document.addEventListener('input', function (e) {
    if (suppressDirty || !usingPreset) return;
    if (e.target.matches('[data-k],[data-tr],[data-spend]')) { usingPreset = false; dropUnsupportedMethod3(); }
  });
  document.addEventListener('change', function (e) {
    if (suppressDirty || !usingPreset) return;
    if (e.target.matches('[data-k],[data-tr],[data-spend]')) { usingPreset = false; dropUnsupportedMethod3(); }
  });
})();
