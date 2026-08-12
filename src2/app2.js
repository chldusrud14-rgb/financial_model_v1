(function () {
  'use strict';
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var el = function (t, c, h) { var e = document.createElement(t); if (c) e.className = c; if (h !== undefined) e.innerHTML = h; return e; };
  var M = window.SolarModel2, X = window.SolarXlsx2;

  /* ---------- 기본 입력 필드 ----------
     essential: 항상 보이는 핵심 가정. 나머지는 그룹별로 묶어서
     "상세 가정" 접힘 영역에 넣는다 — 26개를 한 화면에 쭉 나열하지 않는다. */
  var CORE = [
    { k: 'projectName', label: '사업명', type: 'text', def: '태양광 발전사업', essential: true },
    { k: 'capacityMW', label: '설비용량', type: 'number', def: 99.998, unit: 'MW', essential: true },
    { k: 'dailyHours', label: '일일 발전시간(환산)', type: 'number', def: 3.77, unit: 'h/day', essential: true,
      hint: '하루 24시간 중 정격출력으로 발전한 것으로 환산한 시간. 이용률(%) = 이 값÷24. 보통 3.3~3.9' },
    { k: 'tariff', label: 'PPA 또는 SMP+REC 판매단가', type: 'number', def: 154.8, unit: '원/kWh', essential: true,
      hint: 'PPA·SMP+REC 등 여러 단가를 물량가중평균한 값' },
    { k: 'capexEok', label: '총사업비(건설이자 제외)', type: 'number', def: 1410.69, unit: '억원', essential: true,
      hint: 'EPC·인허가·개발비 등 순수 공사비 합계 — 건설이자(IDC)는 여기 포함 안 함, 자동 계산됨' },
    { k: 'opexEok', label: '운영비(1년차 기준)', type: 'number', def: 49.8, unit: '억원/yr', essential: true },
    { k: 'equityEok', label: '자본금(Equity)', type: 'number', def: 150, unit: '억원', essential: true,
      hint: '전체 출자자의 출자금 합계(Equity 총액)입니다 — 개별 출자자 지분율은 아래 "사업자 구성"에서 나눕니다.' },
    { k: 'equityRatioPct', label: '자기자본비율(선택)', type: 'number', def: '', unit: '%', essential: true, uiOnly: true,
      hint: '입력하면 총사업비 대비 비율로 자본금을 자동 계산합니다(비워두면 위 자본금 절대값을 그대로 씀). 자본금을 직접 고치면 이 값도 같이 갱신됩니다.' },

    { k: 'degradation', label: '연간 출력저하(Degradation)', type: 'number', def: 0.5, unit: '%/yr', group: '발전소 특성',
      hint: '매년 정액으로 발전량이 이만큼씩 줄어든다고 가정 (복리 아님)' },
    { k: 'auxRate', label: '소내 소비율', type: 'number', def: 0, unit: '%', group: '발전소 특성',
      hint: '발전한 전력 중 설비 자체가 쓰는 비율' },
    { k: 'constructionStart', label: '착공시점', type: 'text', def: '2024-06', unit: 'YYYY-MM', group: '발전소 특성' },
    { k: 'constructionMonths', label: '공사기간', type: 'number', def: 16, unit: 'Month', group: '발전소 특성' },
    { k: 'operationYears', label: '운영기간', type: 'number', def: 20, unit: 'Year', group: '발전소 특성' },

    { k: 'dsraEok', label: '최초 DSRA', type: 'number', def: 50, unit: '억원', group: '재원조달·감가상각',
      hint: '준공 시점에 별도로 적립해두는 원리금상환 예비재원' },
    { k: 'depRatio', label: '감가상각 대상비율', type: 'number', def: 95, unit: '%', group: '재원조달·감가상각',
      hint: '총투자비 중 감가상각 대상 자산의 비율(토지 등 제외분 빼고)' },
    { k: 'depYears', label: '감가상각 내용연수', type: 'number', def: 20, unit: 'Year', group: '재원조달·감가상각' },

    { k: 'tariffEscal', label: '판매단가 상승률', type: 'number', def: 0, unit: '%/yr', group: '매출' },
    { k: 'rpsShare', label: 'RPS(SMP+REC) 비중', type: 'number', def: 0, unit: '%', group: '매출',
      hint: '설비용량 중 SMP+REC(RPS)로 정산받는 비중. 나머지는 PPA(위 "판매단가")로 계산됩니다. 0이면 기존과 동일하게 전량 PPA로 계산됩니다.' },
    { k: 'smpPrice', label: 'SMP 단가', type: 'number', def: 135, unit: '원/kWh', group: '매출',
      hint: 'RPS 비중이 0보다 클 때만 사용됩니다.' },
    { k: 'recPrice', label: 'REC 단가', type: 'number', def: 70, unit: '원/kWh', group: '매출',
      hint: 'RPS 비중이 0보다 클 때만 사용됩니다.' },
    { k: 'recWeight', label: 'REC 가중치', type: 'number', def: 1, unit: '배', group: '매출',
      hint: '설비용량 구간별 REC 가중치. RPS 비중이 0(전량 PPA)이면 의미가 없어 자동으로 1로 고정되고 입력이 잠깁니다.' },

    { k: 'opexEscal', label: '운영비 상승률', type: 'number', def: 0.7, unit: '%/yr', group: '운영비' },
    { k: 'decomEok', label: '철거·복구비(만기 시점)', type: 'number', def: 20, unit: '억원', group: '운영비' },

    { k: 'lossRate', label: '이월결손금 공제한도', type: 'number', def: 80, unit: '%', group: '세무',
      hint: '그 해 과세소득 중 이월결손금으로 상계 가능한 비율' },
    { k: 'taxFlat', label: '단일 법인세율', type: 'number', def: 21, unit: '%', group: '세무',
      hint: '누진세율 대신 단일세율을 쓰고 싶을 때만 참고(기본은 누진 브래킷 적용)' },

    { k: 'dsraMonths', label: 'DSRA 적립기준', type: 'number', def: 6, unit: 'Month', group: '현금관리·배당',
      hint: '차기 몇 개월분 원리금을 항상 예비로 쌓아둘지' },
    { k: 'minCash', label: '배당 후 최소보유현금', type: 'number', def: 10, unit: '억원', group: '현금관리·배당' },
    { k: 'divDSCR', label: '배당제한 — 단순DSCR', type: 'number', def: 1.1, unit: 'x', group: '현금관리·배당',
      hint: '이 값 미만이면 그 분기 원리금 상환여력이 부족하다고 보고 배당을 막음' },
    { k: 'divCumDSCR', label: '배당제한 — 누적DSCR', type: 'number', def: 1.15, unit: 'x', group: '현금관리·배당' },
    { k: 'divStartYear', label: '배당개시 연차', type: 'number', def: 2, unit: 'Year', group: '현금관리·배당',
      hint: '"당진 FS 불러오기"를 안 썼을 때만 적용되는 간이 기준(실측 배당 로직 미사용 시 폴백)' },

    { k: 'discount', label: '할인율(NPV 계산용)', type: 'number', def: 5.5, unit: '%', group: '평가' }
  ];

  var SHAREHOLDERS = [{ name: '출자자1', stakePct: 100 }];

  // 민감도 분석 시나리오 — 판매단가/총투자비/운영비는 %조정, 이자율은 bp
  // 조정(모든 트랜치 건설·운영금리에 동일하게 가산), 총사업비 변동은
  // 억원 단위 절대 가산(%조정과 별개로 같이 적용됨). 기본 3행은 시작
  // 템플릿일 뿐, 값은 전부 직접 편집 가능.
  var SENS_ROWS = [
    { name: 'Base', tariffPct: 0, capexPct: 0, opexPct: 0, ratebp: 0, capexAbsEok: 0 },
    { name: 'Upside', tariffPct: 10, capexPct: -5, opexPct: -5, ratebp: -50, capexAbsEok: 0 },
    { name: 'Downside', tariffPct: -10, capexPct: 10, opexPct: 10, ratebp: 50, capexAbsEok: 0 }
  ];
  var lastSensResults = null;

  var TRANCHES = [
    { key: 'A', name: '선순위A', amountEok: 500, order: 2, rateO: 5.6, graceYears: 2, repayYears: 16 },
    { key: 'B', name: '선순위B', amountEok: 500, order: 2, rateO: 5.5, graceYears: 2, repayYears: 16 },
    { key: 'C', name: '선순위C', amountEok: 350, order: 4, rateO: 5.5, graceYears: 1.25, repayYears: 16 },
    { key: 'D', name: '선순위D', amountEok: 0, order: 3, rateO: 2.25, graceYears: 5, repayYears: 10 },
    { key: 'sub', name: '후순위', amountEok: 0, order: 1, rateO: 5.65, graceYears: 15, repayYears: 2.75 }
  ];

  function fieldEl(d) {
    var wrapf = el('div', 'f' + (d.type === 'text' ? ' full' : ''));
    wrapf.appendChild(el('label', null, d.label));
    var inw = el('div', 'in2');
    var input = document.createElement('input');
    input.type = d.type === 'text' ? 'text' : 'number';
    input.step = 'any';
    input.value = d.def;
    input.dataset.k = d.k;
    inw.appendChild(input);
    if (d.unit) inw.appendChild(el('span', 'unit', d.unit));
    wrapf.appendChild(inw);
    if (d.hint) wrapf.appendChild(el('div', 'fhint', d.hint));
    return wrapf;
  }

  // 필수 입력 배치 — 자동 2열 순서채움 대신 명시적으로 행을 지정해서
  // "총사업비+운영비", "자본금+자기자본비율"이 나란히 붙게 한다.
  var ESSENTIAL_ROWS = [
    ['projectName'],
    ['capacityMW', 'dailyHours'],
    ['tariff'],
    ['capexEok', 'opexEok'],
    ['equityEok', 'equityRatioPct']
  ];
  function fieldByKey(k) { return CORE.filter(function (d) { return d.k === k; })[0]; }

  function buildCore() {
    var wrap = $('#core');
    var essential = el('div', 'grid');
    ESSENTIAL_ROWS.forEach(function (row) {
      row.forEach(function (k) {
        var d = fieldByKey(k);
        if (d) essential.appendChild(fieldEl(d));
      });
    });
    wrap.appendChild(essential);

    // 그룹별로 묶어서 "상세 가정" 접힘 영역에 — 26개를 한 화면에 쭉 나열하지 않는다.
    var groups = [];
    CORE.filter(function (d) { return !d.essential; }).forEach(function (d) {
      var g = groups.filter(function (x) { return x.name === d.group; })[0];
      if (!g) { g = { name: d.group, items: [] }; groups.push(g); }
      g.items.push(d);
    });
    var det = document.createElement('details');
    det.className = 'more';
    var sum = document.createElement('summary');
    sum.textContent = '상세 가정 (발전소 특성 · 재원조달 · 세무 · 배당 등)';
    det.appendChild(sum);
    var body2 = el('div', 'body2');
    groups.forEach(function (g) {
      var grp = el('div', 'grp');
      grp.appendChild(el('b', null, g.name));
      var gg = el('div', 'grid');
      g.items.forEach(function (d) { gg.appendChild(fieldEl(d)); });
      grp.appendChild(gg);
      body2.appendChild(grp);
    });
    det.appendChild(body2);
    wrap.appendChild(det);
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

  // 트랜치 조건을 몰라도(또는 대략적인 민감도 확인만 하고 싶을 때) 쓰는
  // 간편설정 — 부채 전액(총사업비-자본금 추정)을 선순위A 하나로 몰아넣고
  // 나머지는 0으로 비워서 표준적인 조건(5.5%/5.5%, 거치2년, 상환15년,
  // 원금균등)으로 즉시 계산 가능하게 만든다. 스프레드곡선의
  // "균등분배로 재설정"과 같은 성격의 단순화 버튼.
  function quickFillTranches() {
    var capex = Number($('[data-k="capexEok"]').value) || 0;
    var equity = Number($('[data-k="equityEok"]').value) || 0;
    var debt = Math.max(0, capex - equity);
    var std = { rateC: 5.5, rateO: 5.5, graceYears: 2, repayYears: 15, method: 1 };
    TRANCHES.forEach(function (tr, idx) {
      var amt = idx === 0 ? debt : 0;
      setVal('input[data-tr="' + tr.key + '"][data-f="amountEok"]', amt);
      setVal('input[data-tr="' + tr.key + '"][data-f="order"]', 1);
      setVal('input[data-tr="' + tr.key + '"][data-f="rateC"]', std.rateC);
      setVal('input[data-tr="' + tr.key + '"][data-f="rateO"]', std.rateO);
      setVal('input[data-tr="' + tr.key + '"][data-f="graceYears"]', std.graceYears);
      setVal('input[data-tr="' + tr.key + '"][data-f="repayYears"]', std.repayYears);
      setVal('select[data-tr="' + tr.key + '"]', std.method);
    });
    if (usingPreset) { usingPreset = false; }
    toast('부채 전액(' + f0(debt) + '억원 추정)을 선순위A 하나로 단순화했습니다 — 필요하면 표에서 직접 조정하세요');
  }

  /* ---------- 사업자(출자자) 구성 ----------
     자본금을 나눠 낼 여러 출자자를 이름+지분율로 입력받는다. 모든 출자자는
     지분율만큼 자본금 납입·배당 수령에 비례 참여한다고 가정(pro-rata). */
  function shRow(sh, idx) {
    var row = el('div', 'shrow');
    var name = document.createElement('input');
    name.type = 'text'; name.value = sh.name; name.dataset.shName = idx;
    var stake = document.createElement('input');
    stake.type = 'number'; stake.step = 'any'; stake.value = sh.stakePct; stake.dataset.shStake = idx;
    var unit = el('span', 'unit', '%');
    row.appendChild(name); row.appendChild(stake); row.appendChild(unit);
    var rm = document.createElement('button');
    rm.type = 'button'; rm.className = 'rm'; rm.textContent = '×';
    rm.addEventListener('click', function () {
      if (SHAREHOLDERS.length <= 1) return;
      SHAREHOLDERS.splice(idx, 1);
      buildShareholderGrid();
    });
    row.appendChild(rm);
    return row;
  }

  function buildShareholderGrid() {
    var box = $('#shbox');
    box.innerHTML = '';
    SHAREHOLDERS.forEach(function (sh, idx) { box.appendChild(shRow(sh, idx)); });
    updateShareholderSum();
  }

  function readShareholders() {
    var names = Array.prototype.slice.call(document.querySelectorAll('[data-sh-name]'));
    return names.map(function (e) {
      var idx = e.dataset.shName;
      var stakeEl = document.querySelector('[data-sh-stake="' + idx + '"]');
      return { name: e.value || ('출자자' + (Number(idx) + 1)), stakePct: Number(stakeEl.value) || 0 };
    });
  }

  function updateShareholderSum() {
    var box = $('#shsum');
    if (!box) return;
    var list = readShareholders();
    // 화면에 그려진 값을 상태(SHAREHOLDERS)에도 반영해둔다(재렌더 시 유지).
    SHAREHOLDERS = list;
    var sum = list.reduce(function (a, s) { return a + (s.stakePct || 0); }, 0);
    var ok = Math.abs(sum - 100) < 0.01;
    box.className = 'spendsum ' + (ok ? 'ok' : 'bad');
    box.innerHTML = '<span>지분율 합계: ' + sum.toFixed(2) + '%</span><span>' + (ok ? '100% — 일치' : '100%와 차이 ' + (sum - 100).toFixed(2) + '%p') + '</span>';
  }

  /* ---------- 민감도 분석 ----------
     화면 입력(일반 경로) 기준으로 판매단가/총투자비/운영비/이자율을
     시나리오별로 조정해 여러 번 재계산한다. 프리셋(당진 FS)은 실측
     분기 데이터(periodOverrides)로 결과가 고정돼 있어 이 조정들이
     반영되지 않으므로 의도적으로 막는다. */
  function sensRow(sc, idx) {
    var tr = document.createElement('tr');
    function cell(val, f, isText) {
      var td = document.createElement('td');
      var input = document.createElement('input');
      input.type = isText ? 'text' : 'number'; input.step = 'any';
      input.value = val; input.dataset.sensIdx = idx; input.dataset.sensF = f;
      td.appendChild(input);
      return td;
    }
    tr.appendChild(cell(sc.name, 'name', true));
    tr.appendChild(cell(sc.tariffPct, 'tariffPct'));
    tr.appendChild(cell(sc.capexPct, 'capexPct'));
    tr.appendChild(cell(sc.opexPct, 'opexPct'));
    tr.appendChild(cell(sc.ratebp, 'ratebp'));
    tr.appendChild(cell(sc.capexAbsEok, 'capexAbsEok'));
    var rmTd = document.createElement('td');
    var rm = document.createElement('button');
    rm.type = 'button'; rm.className = 'rm'; rm.textContent = '×';
    rm.addEventListener('click', function () {
      if (SENS_ROWS.length <= 1) return;
      SENS_ROWS.splice(idx, 1);
      buildSensGrid();
    });
    rmTd.appendChild(rm);
    tr.appendChild(rmTd);
    return tr;
  }

  function buildSensGrid() {
    var box = $('#sensBox');
    box.innerHTML = '';
    var t = el('table', 'tr');
    t.innerHTML = '<thead><tr><th>시나리오</th><th>판매단가(%)</th><th>총투자비(%)</th><th>운영비(%)</th><th>이자율(bp)</th><th>총사업비 변동(억원)</th><th></th></tr></thead>';
    var tb = document.createElement('tbody');
    SENS_ROWS.forEach(function (sc, idx) { tb.appendChild(sensRow(sc, idx)); });
    t.appendChild(tb);
    box.appendChild(t);
  }

  function readSensRows() {
    var trs = Array.prototype.slice.call(document.querySelectorAll('#sensBox tbody tr'));
    return trs.map(function (tr) {
      var o = {};
      Array.prototype.slice.call(tr.querySelectorAll('input')).forEach(function (input) {
        o[input.dataset.sensF] = input.dataset.sensF === 'name' ? input.value : Number(input.value);
      });
      return o;
    });
  }

  // 판매단가/총투자비/운영비는 배율(%), 이자율은 %p(bp/100)를 기본 입력에
  // 더해 새 입력 객체를 만든다 — 트랜치 개별 rateC/rateO를 모두 같은
  // bp만큼 평행이동, tariffTracks/opexItems가 있으면 그 항목들도 같이 조정.
  // 총사업비 변동(억원)은 %조정과 별개로 절대금액을 더/뺀다(둘 다 채우면
  // 둘 다 같이 적용됨 — 순서는 %조정 먼저, 그다음 절대금액 가산).
  function applyScenario(baseInp, sc) {
    var c = JSON.parse(JSON.stringify(baseInp));
    var tf = 1 + (sc.tariffPct || 0) / 100;
    var cf = 1 + (sc.capexPct || 0) / 100;
    var of = 1 + (sc.opexPct || 0) / 100;
    var rb = (sc.ratebp || 0) / 100;
    if (c.tariff != null) c.tariff = c.tariff * tf;
    if (c.tariffTracks) c.tariffTracks.forEach(function (t) { t.price = t.price * tf; });
    if (c.capexEok != null) c.capexEok = c.capexEok * cf + (sc.capexAbsEok || 0);
    if (c.opexEok != null) c.opexEok = c.opexEok * of;
    if (c.opexItems) c.opexItems.forEach(function (it) { it.annualKRWm = it.annualKRWm * of; });
    if (c.tranches) c.tranches.forEach(function (t) {
      t.rateC = (t.rateC || 0) + rb; t.rateO = (t.rateO || 0) + rb;
    });
    return c;
  }

  function runSensitivity() {
    if (usingPreset) {
      toast('민감도 분석은 "당진 FS 불러오기" 프리셋 상태에서는 의미가 없습니다(실측치로 고정됨) — 아무 값이나 수정해 프리셋을 해제한 뒤 사용하세요');
      return;
    }
    var core = readCore();
    var tranches = readTranches();
    var spendCurve = readSpendCurve();
    var baseInp = Object.assign({}, core, {
      ppy: 4, capacityFactor: undefined, spendCurve: spendCurve, equityOrder: 1,
      tranches: tranches.map(function (t) {
        return { name: t.name, amountEok: t.amountEok, order: t.order, rateC: t.rateC, rateO: t.rateO, graceYears: t.graceYears, repayYears: t.repayYears, method: t.method };
      }),
      taxMode: 1, localSurtaxRate: 10
    });
    if (baseInp.rpsShare > 0) {
      baseInp.tariffTracks = [
        { share: baseInp.rpsShare / 100, price: baseInp.smpPrice + baseInp.recWeight * baseInp.recPrice, escal: baseInp.tariffEscal },
        { share: 1 - baseInp.rpsShare / 100, price: baseInp.tariff, escal: baseInp.tariffEscal }
      ];
    }
    baseInp.shareholders = readShareholders();

    SENS_ROWS = readSensRows();
    var results = SENS_ROWS.map(function (sc) {
      try {
        var m = M.computeModel(applyScenario(baseInp, sc));
        return { name: sc.name, sc: sc, kpi: m.kpi };
      } catch (e) {
        return { name: sc.name, sc: sc, error: e.message };
      }
    });
    lastSensResults = results;
    renderSensResults(results);
    toast('민감도 분석 완료 — ' + results.length + '개 시나리오');
  }

  function renderSensResults(results) {
    var box = $('#sensResults');
    box.innerHTML = '';
    var t = el('table', 'tr');
    t.innerHTML = '<thead><tr><th>시나리오</th><th>Equity IRR(배당)</th><th>Equity IRR(FCFE)</th>' +
      '<th>Project IRR</th><th>최소DSCR</th><th>NPV(억원)</th><th>투자배수</th></tr></thead>';
    var tb = document.createElement('tbody');
    results.forEach(function (r) {
      var tr = document.createElement('tr');
      if (r.error) {
        tr.innerHTML = '<td>' + r.name + '</td><td colspan="6" style="color:var(--bad)">계산 실패: ' + r.error + '</td>';
      } else {
        var k = r.kpi;
        tr.innerHTML = '<td>' + r.name + '</td>' +
          '<td style="text-align:right">' + pct(k.dividendIRR) + '%</td>' +
          '<td style="text-align:right">' + pct(k.equityIRR) + '%</td>' +
          '<td style="text-align:right">' + pct(k.projectIRR) + '%</td>' +
          '<td style="text-align:right">' + (k.minDSCRAnnual === null ? '—' : k.minDSCRAnnual.toFixed(3)) + 'x</td>' +
          '<td style="text-align:right">' + feok(k.npv) + '</td>' +
          '<td style="text-align:right">' + fx(k.equityMultiple) + '배</td>';
      }
      tb.appendChild(tr);
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
    updateSpendSum();
  }

  function updateSpendSum() {
    var box = $('#spendsum');
    if (!box) return;
    var capex = Number($('[data-k="capexEok"]').value) || 0;
    var sum = readSpendCurve().reduce(function (a, b) { return a + (b || 0); }, 0);
    var diff = sum - capex;
    var ok = Math.abs(diff) < 0.5;
    box.className = 'spendsum ' + (ok ? 'ok' : 'bad');
    box.innerHTML = '<span>지출 합계: ' + sum.toFixed(2) + '억원</span><span>총사업비: ' + capex.toFixed(2) + '억원' +
      (ok ? ' — 일치' : ' — 차이 ' + diff.toFixed(2) + '억원') + '</span>';
  }

  // RPS 비중이 0(전량 PPA)이면 REC 가중치는 의미가 없으므로 1로 고정하고
  // 입력을 잠근다 — RPS 비중이 0보다 커지면 다시 편집 가능하게 푼다.
  function updateRecWeightState() {
    var rpsEl = $('[data-k="rpsShare"]');
    var rwEl = $('[data-k="recWeight"]');
    if (!rpsEl || !rwEl) return;
    var rps = Number(rpsEl.value) || 0;
    if (rps <= 0) { rwEl.value = 1; rwEl.disabled = true; }
    else { rwEl.disabled = false; }
  }

  function readCore() {
    var inp = {};
    CORE.forEach(function (d) {
      var v = $('[data-k="' + d.k + '"]').value;
      inp[d.k] = d.type === 'text' ? v : Number(v);
    });
    return inp;
  }

  // 자본금(절대값) ↔ 자기자본비율(%, 총사업비 대비) 양방향 동기화.
  // fromRatio=true면 비율→자본금, false면 자본금→비율로 갱신한다.
  function syncEquityRatio(fromRatio) {
    var capexEl = $('[data-k="capexEok"]'), eqEl = $('[data-k="equityEok"]'), ratioEl = $('[data-k="equityRatioPct"]');
    if (!capexEl || !eqEl || !ratioEl) return;
    var capex = Number(capexEl.value) || 0;
    if (fromRatio) {
      if (ratioEl.value !== '' && capex > 0) eqEl.value = (capex * Number(ratioEl.value) / 100).toFixed(2);
    } else {
      if (capex > 0) ratioEl.value = (Number(eqEl.value) / capex * 100).toFixed(2);
    }
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
    syncEquityRatio(false); // 참고용 — 실제 계산은 presetInp를 그대로 씀, 이 필드는 표시만

    buildSpendCurve();
    ref.spendCurve_KRWm.forEach(function (v, i) {
      var e = $('[data-spend="' + i + '"]'); if (e) e.value = (v / 100).toFixed(2);
    });
    updateSpendSum();

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
  var fx = function (n) { return n === null || n === undefined || isNaN(n) ? '—' : n.toFixed(2); };
  var fyr = function (n) { return n === null || n === undefined || isNaN(n) ? '회수 안 됨' : n.toFixed(1); };
  var feok = function (n) { return Math.round(n / 100).toLocaleString('ko-KR'); }; // KRWm → 억원

  // 값 부호/구간에 따라 강조색을 입힌다 — 음수(적자)는 빨강, DSCR은
  // 임계값(1.0/1.2)에 따라 빨강/주황/초록, 회수 안 됨도 빨강.
  function kpiTone(raw, kind) {
    if (raw === null || raw === undefined || isNaN(raw)) return kind === 'payback' ? 'bad' : '';
    if (kind === 'dscr') return raw < 1.0 ? 'bad' : raw < 1.2 ? 'warn' : 'good';
    if (kind === 'neg') return raw < 0 ? 'bad' : '';
    return '';
  }

  function kpiGroup(title, items) {
    var wrap = el('div', 'grp');
    wrap.appendChild(el('b', null, title));
    var grid = el('div', 'kpis');
    items.forEach(function (it) {
      var d = el('div', 'kpi' + (it[4] ? ' ' + it[4] : ''));
      d.appendChild(el('div', 'k', it[0]));
      d.appendChild(el('div', 'v', it[1] + '<em>' + it[2] + '</em>'));
      if (it[3]) d.appendChild(el('div', 'n2', it[3]));
      grid.appendChild(d);
    });
    wrap.appendChild(grid);
    return wrap;
  }

  function renderKPIs() {
    var k = model.kpi;
    var capMW = model.inp.capacityMW;
    var capexPerMW = capMW > 0 ? model.tic / 100 / capMW : null;         // 억원/MW
    var opexPerMWyr = capMW > 0 ? (k.totalOpex / model.inp.operationYears) / capMW : null; // KRWm/MW/yr

    var box = $('#kpis'); box.innerHTML = '';

    // 히어로 카드 — 사업자가 가장 먼저 보고 싶어할 "실수령 기준" 수익률
    var hero = el('div', 'kpiHero');
    hero.innerHTML = '<div class="hk">Equity IRR (배당)</div>' +
      '<div class="hv' + (k.dividendIRR !== null && k.dividendIRR < 0 ? ' bad' : '') + '">' + pct(k.dividendIRR) + '<em>%</em></div>' +
      '<div class="hsub">실제 배당 수령 기준 — 가장 보수적/현실적인 수익률 · 투자배수 ' + fx(k.equityMultiple) + '배 · 회수기간 ' + fyr(k.paybackYears) + (k.paybackYears === null ? '' : '년') + '</div>';
    box.appendChild(hero);

    // 세전은 세후보다 덜 중요한 참고값이라 왼쪽에 두고 톤을 흐리게(dim),
    // 세후는 오른쪽에 두고 원래 톤 그대로 — 같은 지표를 나란히 비교하기 쉽게.
    box.appendChild(kpiGroup('수익성 상세', [
      ['Equity IRR (FCFE) 세전', pct(k.equityIRRPre), '%', '', (kpiTone(k.equityIRRPre, 'neg') + ' dim').trim()],
      ['Equity IRR (FCFE) 세후', pct(k.equityIRR), '%', '', kpiTone(k.equityIRR, 'neg')],
      ['Project IRR 세전', pct(k.projectIRRPre), '%', '', (kpiTone(k.projectIRRPre, 'neg') + ' dim').trim()],
      ['Project IRR 세후', pct(k.projectIRR), '%', '', kpiTone(k.projectIRR, 'neg')]
    ]));
    box.appendChild(kpiGroup('사업 규모·수익구조', [
      ['연평균 EBITDA', f0(k.avgEbitda), 'KRWm/yr', '', kpiTone(k.avgEbitda, 'neg')],
      ['EBITDA 마진', k.ebitdaMargin === null ? '—' : (k.ebitdaMargin * 100).toFixed(1), '%'],
      ['NPV(프로젝트)', feok(k.npv), '억원', '', kpiTone(k.npv, 'neg')],
      ['LCOE(균등화발전단가)', k.lcoe ? k.lcoe.toFixed(1) : '—', '원/kWh', '판매단가와 같은 단위 — LCOE보다 판매단가가 높아야 수익이 남'],
      ['MW당 총투자비(TIC)', capexPerMW === null ? '—' : capexPerMW.toFixed(2), '억원/MW', '건설이자(IDC) 포함 총투자비 기준 — 위에 입력한 "총사업비"(건설이자 제외)와는 다른 값입니다'],
      ['MW당 연평균 운영비', opexPerMWyr === null ? '—' : f0(opexPerMWyr), 'KRWm/MW/yr']
    ]));
    box.appendChild(kpiGroup('리스크', [
      ['최소 DSCR(연 합산)', k.minDSCRAnnual === null ? '—' : k.minDSCRAnnual.toFixed(3), 'x',
        '연도별 CFADS합/원리금합 중 최솟값 — 1.0 미만이면 그 해 상환재원이 부족했다는 뜻', kpiTone(k.minDSCRAnnual, 'dscr')]
    ]));

    $('#metaNote').textContent = '기간 수(분기): ' + model.periods.length +
      ' · 마지막 분기: ' + model.periods[model.periods.length - 1].endStr +
      ' · 총영업수익 ' + f0(k.totalRevenue) + 'KRWm · 총선순위이자 ' + f0(k.totalInterest) + 'KRWm' +
      ' · 총투자비(TIC) ' + f0(model.tic) + 'KRWm';

    var shBox = $('#shResults');
    shBox.innerHTML = '';
    if (k.shareholders && k.shareholders.length > 1) {
      var wrap = el('div', 'grp');
      wrap.appendChild(el('b', null, '사업자별 배분'));
      var t = el('table', 'tr');
      t.innerHTML = '<thead><tr><th>사업자</th><th>지분율</th><th>출자금액(KRWm)</th><th>누적배당(KRWm)</th></tr></thead>';
      var tb = document.createElement('tbody');
      k.shareholders.forEach(function (sh) {
        var row = document.createElement('tr');
        row.innerHTML = '<td>' + sh.name + '</td><td style="text-align:right">' + sh.stakePct.toFixed(2) +
          '%</td><td style="text-align:right">' + f0(sh.equityKRWm) + '</td><td style="text-align:right">' + f0(sh.dividendKRWm) + '</td>';
        tb.appendChild(row);
      });
      t.appendChild(tb);
      wrap.appendChild(t);
      shBox.appendChild(wrap);
    }
  }

  function toast(msg) {
    var t = $('#toast'); t.textContent = msg; t.classList.add('on');
    setTimeout(function () { t.classList.remove('on'); }, 2600);
  }

  function run() {
    // 새로 생성하면 이전 민감도 결과는 다른 기준값에서 나온 것이라 무효 —
    // 엑셀에 잘못 딸려가지 않게 비운다(다시 보려면 민감도를 재실행).
    lastSensResults = null;
    var sr = $('#sensResults'); if (sr) sr.innerHTML = '';
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
      // RPS(SMP+REC)/PPA 매출 이원화 — RPS 비중이 0보다 크면 두 트랙으로
      // 나눠 엔진에 전달한다(엔진은 이미 tariffTracks를 지원). RPS 단가는
      // 원본 Revenue!row58/63 구조와 같이 "SMP 단가 + REC가중치×REC단가"로
      // 계산한다. 0(기존과 동일, 전량 PPA)이면 기존 단일 tariff 경로 그대로.
      if (core.rpsShare > 0) {
        inp.tariffTracks = [
          { share: core.rpsShare / 100, price: core.smpPrice + core.recWeight * core.recPrice, escal: core.tariffEscal },
          { share: 1 - core.rpsShare / 100, price: core.tariff, escal: core.tariffEscal }
        ];
      }
    }
    // 사업자 구성은 프리셋 여부와 무관하게 화면 입력을 그대로 쓴다 —
    // 자본금 총액을 나눠 낸 여러 출자자에게 지분율만큼 배당을 배분하는
    // 표시용 계산이라 원본 검증치(periodOverrides 등)와는 독립적이다.
    inp = Object.assign({}, inp, { shareholders: readShareholders() });
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
    if (lastSensResults) model.sensitivity = lastSensResults;
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
  buildShareholderGrid();
  buildSpendCurve();
  updateRecWeightState();
  $('#trQuick').addEventListener('click', quickFillTranches);
  $('[data-k="rpsShare"]').addEventListener('input', updateRecWeightState);
  $('[data-k="constructionMonths"]').addEventListener('change', buildSpendCurve);
  $('[data-k="capexEok"]').addEventListener('change', buildSpendCurve);
  $('[data-k="capexEok"]').addEventListener('input', function () { if ($('[data-k="equityRatioPct"]').value !== '') syncEquityRatio(true); });
  $('[data-k="equityEok"]').addEventListener('input', function () { syncEquityRatio(false); });
  $('[data-k="equityRatioPct"]').addEventListener('input', function () { syncEquityRatio(true); });
  $('#spendReset').addEventListener('click', buildSpendCurve);
  $('#spendbox').addEventListener('input', updateSpendSum);
  $('#shAdd').addEventListener('click', function () {
    SHAREHOLDERS.push({ name: '출자자' + (SHAREHOLDERS.length + 1), stakePct: 0 });
    buildShareholderGrid();
  });
  $('#shbox').addEventListener('input', updateShareholderSum);
  buildSensGrid();
  $('#sensAdd').addEventListener('click', function () {
    SENS_ROWS.push({ name: '시나리오' + (SENS_ROWS.length + 1), tariffPct: 0, capexPct: 0, opexPct: 0, ratebp: 0, capexAbsEok: 0 });
    buildSensGrid();
  });
  $('#sensRun').addEventListener('click', runSensitivity);
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
