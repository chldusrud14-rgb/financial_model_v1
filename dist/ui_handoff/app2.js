(function () {
  'use strict';
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var el = function (t, c, h) { var e = document.createElement(t); if (c) e.className = c; if (h !== undefined) e.innerHTML = h; return e; };
  var M = window.SolarModel2, X = window.SolarXlsx2;

  // 착공시점+공사기간 -> 준공시점(COD) 표시용 계산 — engine2.js의
  // ym()/addM()과 같은 방식(월 단위, "착공월+공사기간=준공 다음달"이라
  // 마지막 달은 -1). 순수 화면 표시용이라 엔진 계산에는 관여하지 않는다.
  function updateCodDisplay() {
    var disp = $('[data-display="codDisplay"]');
    if (!disp) return;
    var startEl = $('[data-k="constructionStart"]'), monthsEl = $('[data-k="constructionMonths"]');
    if (!startEl || !monthsEl) return;
    var m = /^(\d{4})-(\d{1,2})$/.exec(startEl.value || '');
    var months = Number(monthsEl.value);
    if (!m || !months) { disp.textContent = '—'; return; }
    var y = Number(m[1]), mo = Number(m[2]);
    var t = y * 12 + (mo - 1) + months - 1; // 준공월(착공 포함 공사기간 마지막 달)
    var cy = Math.floor(t / 12), cm = (t % 12) + 1;
    disp.textContent = cy + '-' + (cm < 10 ? '0' + cm : cm);
  }

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
      toggle: { id: 'capexDetailToggle', label: '항목별 입력' },
      hint: 'EPC·인허가·개발비 등 순수 공사비 합계 — 건설이자(IDC)는 여기 포함 안 함, 자동 계산됨. 항목별 금액을 모르면 체크하지 말고 합계만 입력(빠른 사업성 검토용) — 체크하면 아래 항목 표가 나타나고 이 필드는 항목 합계로 자동 계산됩니다.' },
    { k: 'opexEok', label: '운영비(1년차 기준)', type: 'number', def: 49.8, unit: '억원/yr', essential: true,
      toggle: { id: 'opexDetailToggle', label: '항목별 입력' },
      hint: '항목별 금액을 모르면 체크하지 말고 합계만 입력하세요. 체크하면 항목별 상승률·선순위/후순위까지 반영돼서 계산 정확도가 올라갑니다(총액 근사 대신 항목별 계산 사용).' },
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
    { k: 'codDisplay', label: '준공시점(COD, 자동계산)', type: 'display', group: '발전소 특성',
      hint: '착공시점+공사기간으로 자동 계산됩니다 — 직접 입력하는 칸이 아닙니다.' },
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

  // 총사업비/운영비 항목별 상세 입력 — 신속한 사업성 검토(합계만 입력)와
  // 정밀 검토(항목별 입력) 둘 다 지원. 기본은 합계만 쓰는 간단 모드고,
  // 체크박스로 항목별 입력 모드를 켤 수 있다. 항목명은 원본 당진 FS의
  // 총사업비/운영비 세부내역 구성을 그대로 템플릿으로 씀(건설이자·DSRA는
  // 별도 필드로 이미 있어서 제외).
  var DEFAULT_CAPEX_ITEMS = ['EPC', '감리비', '공사보험료', '토지임대료(선납)', '토지임대료(분납)',
    '사업개발비', '민원처리비', '사업성자문비', '법인운영비', '기타예비비', '신주발행비용', '금융부대비용'];
  // 지급순위 기본값은 전부 선순위(원본 당진 실측치는 항목별로 다르지만,
  // 화면 템플릿 기본값은 단순하게 전부 선순위로 시작 — 필요하면 사용자가
  // 개별 항목만 후순위로 바꾸면 됨).
  var DEFAULT_OPEX_ITEMS = [
    { name: '부지임대료', escal: 0, senior: true },
    { name: 'O&M', escal: 1.5, senior: true },
    { name: '보험료', escal: 0, senior: true },
    { name: '환경모니터링비용', escal: 0, senior: true },
    { name: '소내전력비', escal: 2, senior: true },
    { name: '인건비', escal: 2, senior: true },
    { name: '법인운영비', escal: 2, senior: true },
    { name: '주민보상비', escal: 0, senior: true },
    { name: '도로점용료', escal: 0, senior: true },
    { name: '예비비', escal: 0, senior: true },
    { name: '전력거래수수료', escal: 0, senior: true },
    { name: '관리운영비 성과보수', escal: 0, senior: true }
  ];
  var capexDetailOn = false, opexDetailOn = false;
  var CAPEX_ITEMS = DEFAULT_CAPEX_ITEMS.map(function (n) { return { name: n, amountEok: null }; });
  var OPEX_ITEMS = DEFAULT_OPEX_ITEMS.map(function (d) { return { name: d.name, amountEok: null, escal: d.escal, senior: d.senior }; });

  var SHAREHOLDERS = [{ name: '출자자1', stakePct: 100 }];

  // 민감도 분석 시나리오 — 판매단가(원/kWh)/총사업비(억원)/운영비(억원)/
  // 금리(%) 전부 델타가 아니라 절대값. 빈 칸이면 "사업 기본 가정"에 입력한
  // 값을 그대로 쓴다. Base는 위 입력값을 그대로 참조하는 기준행이라
  // "↑" 버튼으로 채울 수 있고, Case들은 사용자가 원하는 값을 직접
  // 입력하는 행이라 버튼이 없다 — 더 필요하면 "+ 시나리오 추가"로 늘린다.
  var SENS_ROWS = [
    { name: 'Base', tariffAbs: null, capexAbs: null, opexAbs: null, rateAbs: null },
    { name: 'Case1', tariffAbs: null, capexAbs: null, opexAbs: null, rateAbs: null },
    { name: 'Case2', tariffAbs: null, capexAbs: null, opexAbs: null, rateAbs: null }
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
    if (d.toggle) {
      var labelRow = el('div', 'flabel');
      labelRow.appendChild(el('label', null, d.label));
      var tg = document.createElement('label');
      tg.className = 'ftoggle';
      var cb = document.createElement('input');
      cb.type = 'checkbox'; cb.id = d.toggle.id;
      tg.appendChild(cb);
      tg.appendChild(document.createTextNode(' ' + d.toggle.label));
      labelRow.appendChild(tg);
      wrapf.appendChild(labelRow);
    } else {
      wrapf.appendChild(el('label', null, d.label));
    }
    var inw = el('div', 'in2');
    if (d.type === 'display') {
      // 입력칸이 아니라 다른 필드에서 자동 계산되는 값을 보여주기만
      // 하는 필드(예: 준공시점) — mwref와 같은 성격으로 input 대신
      // 회색 이탤릭 텍스트로 렌더링.
      var disp = document.createElement('span');
      disp.className = 'mwref'; disp.style.textAlign = 'left';
      disp.dataset.display = d.k;
      disp.textContent = '—';
      inw.appendChild(disp);
    } else {
      var input = document.createElement('input');
      input.type = d.type === 'text' ? 'text' : 'number';
      input.step = 'any';
      input.value = d.def;
      input.dataset.k = d.k;
      inw.appendChild(input);
      if (d.unit) inw.appendChild(el('span', 'unit', d.unit));
    }
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
      var colsUsed = 0;
      row.forEach(function (k) {
        var d = fieldByKey(k);
        if (!d) return;
        essential.appendChild(fieldEl(d));
        colsUsed += d.type === 'text' ? 2 : 1; // text 필드는 .full로 2칸 다 씀
      });
      // 2칸 그리드는 자식을 그냥 순서대로 채우기 때문에, 한 칸만 쓴 행은
      // 빈 칸을 하나 더 넣어줘야 다음 행이 새 줄에서 시작한다(안 그러면
      // 다음 행 첫 필드가 이 행의 남은 칸에 끼어들어가 버림).
      if (colsUsed === 1) essential.appendChild(el('div'));
      // 총사업비/운영비 항목별 입력 표는 해당 필드 바로 아래 줄에 — 위
      // 필드들처럼 반반(1칸씩) 나눠서 나란히 배치한다. display:none으로
      // 숨기면 그리드 흐름에서 아예 빠져버려 다음 필드(자본금)가 옆으로
      // 끌려오는 버그가 있어서, 항상 그리드에 남겨두고 내용만 비워둔다
      // (비어있으면 CSS로 높이 0 처리 — .itemBox:empty 참고).
      if (row.indexOf('capexEok') >= 0) {
        var capexBox = el('div', 'itemBox'); capexBox.id = 'capexItemBox';
        essential.appendChild(capexBox);
      }
      if (row.indexOf('opexEok') >= 0) {
        var opexBox = el('div', 'itemBox'); opexBox.id = 'opexItemBox';
        essential.appendChild(opexBox);
      }
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

  function capexItemRow(it, idx) {
    var tr = document.createElement('tr');
    var nameTd = document.createElement('td');
    var nameInput = document.createElement('input');
    nameInput.type = 'text'; nameInput.value = it.name; nameInput.dataset.capexIdx = idx; nameInput.dataset.capexF = 'name';
    nameTd.appendChild(nameInput); tr.appendChild(nameTd);
    var amtTd = document.createElement('td');
    var amtInput = document.createElement('input');
    amtInput.type = 'number'; amtInput.step = 'any';
    amtInput.value = it.amountEok === null || it.amountEok === undefined ? '' : it.amountEok;
    amtInput.dataset.capexIdx = idx; amtInput.dataset.capexF = 'amountEok';
    amtTd.appendChild(amtInput); tr.appendChild(amtTd);
    var rmTd = document.createElement('td');
    var rm = document.createElement('button');
    rm.type = 'button'; rm.className = 'rm'; rm.textContent = '×';
    rm.addEventListener('click', function () {
      if (CAPEX_ITEMS.length <= 1) return;
      CAPEX_ITEMS.splice(idx, 1);
      buildCapexItemGrid();
      updateCapexItemSum();
    });
    rmTd.appendChild(rm); tr.appendChild(rmTd);
    return tr;
  }

  function buildCapexItemGrid() {
    var box = $('#capexItemBox');
    if (!box) return;
    box.innerHTML = '';
    var t = el('table', 'tr');
    t.innerHTML = '<thead><tr><th>항목</th><th>금액<br>(억원)</th><th></th></tr></thead>';
    var tb = document.createElement('tbody');
    CAPEX_ITEMS.forEach(function (it, idx) { tb.appendChild(capexItemRow(it, idx)); });
    t.appendChild(tb);
    box.appendChild(t);
    var addBtn = document.createElement('button');
    addBtn.type = 'button'; addBtn.className = 'btn ghost'; addBtn.style.marginTop = '8px';
    addBtn.textContent = '+ 항목 추가';
    addBtn.addEventListener('click', function () {
      CAPEX_ITEMS.push({ name: '항목' + (CAPEX_ITEMS.length + 1), amountEok: null });
      buildCapexItemGrid();
    });
    box.appendChild(addBtn);
    var sum = el('div', 'spendsum'); sum.id = 'capexItemSum';
    box.appendChild(sum);
    updateCapexItemSum();
  }

  function readCapexItems() {
    var trs = Array.prototype.slice.call(document.querySelectorAll('#capexItemBox tbody tr'));
    return trs.map(function (tr) {
      var nameEl = tr.querySelector('[data-capex-f="name"]');
      var amtEl = tr.querySelector('[data-capex-f="amountEok"]');
      return { name: nameEl.value, amountEok: amtEl.value === '' ? null : Number(amtEl.value) };
    });
  }

  function updateCapexItemSum() {
    CAPEX_ITEMS = readCapexItems();
    var sum = CAPEX_ITEMS.reduce(function (a, it) { return a + (it.amountEok || 0); }, 0);
    var box = $('#capexItemSum');
    if (box) box.innerHTML = '<span>항목 합계</span><span>' + sum.toFixed(2) + '억원</span>';
    if (capexDetailOn) {
      var capexEl = $('[data-k="capexEok"]');
      capexEl.value = sum.toFixed(2);
      capexEl.dispatchEvent(new Event('change', { bubbles: true }));
    }
  }

  // box를 display:none으로 숨기면 CSS 그리드 배치에서 그 자리가 통째로
  // 빠져버려서(그리드는 display:none 요소를 아예 없는 것처럼 취급),
  // 바로 다음 필드(자본금)가 옆 칸(운영비 표 자리)으로 끌려 올라오는
  // 버그가 있었다 — 항상 그리드 흐름에는 남겨두고, 안 켰을 때는
  // 내용만 비워서 높이가 0에 가깝게 만드는 방식으로 고친다.
  function toggleCapexDetail(on) {
    capexDetailOn = on;
    var capexEl = $('[data-k="capexEok"]');
    if (capexEl) capexEl.readOnly = on;
    if (on) { buildCapexItemGrid(); updateCapexItemSum(); }
    else { var box = $('#capexItemBox'); if (box) box.innerHTML = ''; }
  }

  function opexItemRow(it, idx) {
    var tr = document.createElement('tr');
    function cell(type, val, f) {
      var td = document.createElement('td');
      var input = document.createElement('input');
      input.type = type; if (type === 'number') input.step = 'any';
      input.value = val === null || val === undefined ? '' : val;
      input.dataset.opexIdx = idx; input.dataset.opexF = f;
      td.appendChild(input);
      return td;
    }
    tr.appendChild(cell('text', it.name, 'name'));
    tr.appendChild(cell('number', it.amountEok, 'amountEok'));
    // MW당 금액은 입력칸이 아니라 위 금액을 설비용량으로 나눈 참고치일
    // 뿐이다 — 입력 가능한 것처럼 보이면 헷갈리니 입력칸(input) 대신
    // 회색 텍스트로만 보여준다(원본 Assum!E341 "per MW" 보조열과 같은
    // 성격, 다만 원본과 반대로 여긴 입력이 아니라 계산 결과 표시).
    var mwTd = document.createElement('td');
    var mwSpan = document.createElement('span');
    mwSpan.className = 'mwref';
    mwSpan.dataset.opexMwref = idx;
    mwSpan.textContent = '—';
    mwTd.appendChild(mwSpan); tr.appendChild(mwTd);
    tr.appendChild(cell('number', it.escal, 'escal'));
    var selTd = document.createElement('td');
    var sel = document.createElement('select');
    sel.innerHTML = '<option value="1">선순위</option><option value="0">후순위</option>';
    sel.value = it.senior === false ? '0' : '1';
    sel.dataset.opexIdx = idx; sel.dataset.opexF = 'senior';
    selTd.appendChild(sel); tr.appendChild(selTd);
    var rmTd = document.createElement('td');
    var rm = document.createElement('button');
    rm.type = 'button'; rm.className = 'rm'; rm.textContent = '×';
    rm.addEventListener('click', function () {
      if (OPEX_ITEMS.length <= 1) return;
      OPEX_ITEMS.splice(idx, 1);
      buildOpexItemGrid();
      updateOpexItemSum();
    });
    rmTd.appendChild(rm); tr.appendChild(rmTd);
    return tr;
  }

  function buildOpexItemGrid() {
    var box = $('#opexItemBox');
    if (!box) return;
    box.innerHTML = '';
    var t = el('table', 'tr');
    t.innerHTML = '<thead><tr><th>항목</th><th>연간금액<br>(억원/yr)</th><th>MW당<br>(참고)</th><th>상승률<br>(%/yr)</th><th>지급<br>순위</th><th></th></tr></thead>';
    var tb = document.createElement('tbody');
    OPEX_ITEMS.forEach(function (it, idx) { tb.appendChild(opexItemRow(it, idx)); });
    t.appendChild(tb);
    box.appendChild(t);
    box.appendChild(el('div', 'trlegend',
      '<b>MW당</b>: 입력칸이 아니라 왼쪽 금액을 위 "설비용량"으로 나눈 참고치입니다(자동 계산, 직접 입력 불가). ' +
      '<b>선순위</b>: 원리금 상환 전에 먼저 빠지는 비용 — 이 비용이 많을수록 원리금 상환여력(DSCR)이 낮게 계산됩니다. ' +
      '<b>후순위</b>: 원리금 상환 후 배당 전에 빠지는 비용 — DSCR 계산엔 영향 없음. 기본값은 전부 선순위입니다.'));
    var addBtn = document.createElement('button');
    addBtn.type = 'button'; addBtn.className = 'btn ghost'; addBtn.style.marginTop = '8px';
    addBtn.textContent = '+ 항목 추가';
    addBtn.addEventListener('click', function () {
      OPEX_ITEMS.push({ name: '항목' + (OPEX_ITEMS.length + 1), amountEok: null, escal: 0, senior: true });
      buildOpexItemGrid();
    });
    box.appendChild(addBtn);
    var sum = el('div', 'spendsum'); sum.id = 'opexItemSum';
    box.appendChild(sum);
    updateOpexItemSum();
  }

  // 항목별 금액을 위 "설비용량"으로 나눈 MW당 참고치를 새로고침한다 —
  // 항목 금액이 바뀌거나 설비용량이 바뀔 때마다 호출.
  function updateOpexMWRefs() {
    var capMW = Number(($('[data-k="capacityMW"]') || {}).value) || 0;
    Array.prototype.slice.call(document.querySelectorAll('#opexItemBox [data-opex-mwref]')).forEach(function (span) {
      var idx = span.dataset.opexMwref;
      var amtEl = document.querySelector('#opexItemBox [data-opex-f="amountEok"][data-opex-idx="' + idx + '"]');
      var amt = amtEl && amtEl.value !== '' ? Number(amtEl.value) : null;
      span.textContent = (amt != null && capMW > 0) ? (amt / capMW).toFixed(3) + '억원/MW' : '—';
    });
  }

  function readOpexItemsDetailed() {
    var trs = Array.prototype.slice.call(document.querySelectorAll('#opexItemBox tbody tr'));
    return trs.map(function (tr) {
      var nameEl = tr.querySelector('[data-opex-f="name"]');
      var amtEl = tr.querySelector('[data-opex-f="amountEok"]');
      var escalEl = tr.querySelector('[data-opex-f="escal"]');
      var seniorEl = tr.querySelector('[data-opex-f="senior"]');
      return {
        name: nameEl.value,
        amountEok: amtEl.value === '' ? null : Number(amtEl.value),
        escal: escalEl.value === '' ? 0 : Number(escalEl.value),
        senior: seniorEl.value === '1'
      };
    });
  }

  function updateOpexItemSum() {
    OPEX_ITEMS = readOpexItemsDetailed();
    var sum = OPEX_ITEMS.reduce(function (a, it) { return a + (it.amountEok || 0); }, 0);
    var box = $('#opexItemSum');
    if (box) box.innerHTML = '<span>항목 합계(1년차 기준)</span><span>' + sum.toFixed(2) + '억원/yr</span>';
    if (opexDetailOn) {
      var opexEl = $('[data-k="opexEok"]');
      opexEl.value = sum.toFixed(2);
      opexEl.dispatchEvent(new Event('change', { bubbles: true }));
    }
    updateOpexMWRefs();
  }

  function toggleOpexDetail(on) {
    opexDetailOn = on;
    var opexEl = $('[data-k="opexEok"]');
    if (opexEl) opexEl.readOnly = on;
    if (on) { buildOpexItemGrid(); updateOpexItemSum(); }
    else { var box = $('#opexItemBox'); if (box) box.innerHTML = ''; }
  }

  function buildTrancheGrid() {
    var box = $('#trbox');
    box.innerHTML = ''; // × 삭제로 재호출될 때 이전 표가 안 지워지고 쌓이던 버그 방지
    var t = el('table', 'tr');
    // 열 폭을 고정 배분 — "트랜치"(선순위A 등) 라벨이 줄바꿈되던 문제와
    // "방식" 드롭다운(원리금균등 등) 글자가 잘리던 문제를 같이 해결.
    t.innerHTML = '<colgroup>' +
      '<col style="width:10%"><col style="width:12%"><col style="width:8%">' +
      '<col style="width:10%"><col style="width:10%"><col style="width:8%">' +
      '<col style="width:8%"><col style="width:28%"><col style="width:6%">' +
      '</colgroup>' +
      '<thead><tr>' +
      '<th>트랜치</th><th>금액(억원)</th><th>투입순서</th><th>건설금리(%)</th><th>운영금리(%)</th>' +
      '<th>거치(yr)</th><th>상환(yr)</th><th>방식</th><th></th></tr></thead>';
    var tb = document.createElement('tbody');
    TRANCHES.forEach(function (tr, ti) {
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
      // 해당 없는 트랜치는 지울 수 있게 — 사업자구성/민감도 분석과
      // 같은 × 버튼 패턴. 최소 1개는 남겨야 함(전부 지우면 부채가
      // 아예 없는 상태가 돼서 계산 자체는 되지만 UI가 텅 비어버림).
      var rmTd = document.createElement('td');
      var rm = document.createElement('button');
      rm.type = 'button'; rm.className = 'rm'; rm.textContent = '×';
      rm.addEventListener('click', function () {
        if (TRANCHES.length <= 1) return;
        TRANCHES.splice(ti, 1);
        buildTrancheGrid();
      });
      rmTd.appendChild(rm); row.appendChild(rmTd);
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
    // 입력칸 위에 "출자자명"/"지분율" 열 제목을 달아준다 — 아래 행과
    // 같은 flex 비율(1.4/0.7)로 맞추고, unit·삭제버튼 자리만큼 빈 칸을
    // 채워서 좌우 정렬이 어긋나지 않게 한다.
    var head = el('div', 'shrow shrow-head');
    head.appendChild(el('span', 'shrow-label', '출자자명'));
    head.appendChild(el('span', 'shrow-label', '지분율'));
    head.appendChild(el('span', 'unit', ''));
    var spacer = el('span'); spacer.style.width = '30px'; spacer.style.flex = '0 0 auto';
    head.appendChild(spacer);
    box.appendChild(head);
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
      input.value = val === null || val === undefined ? '' : val;
      input.dataset.sensIdx = idx; input.dataset.sensF = f;
      td.appendChild(input);
      return td;
    }
    tr.appendChild(cell(sc.name, 'name', true));
    tr.appendChild(cell(sc.tariffAbs, 'tariffAbs'));
    tr.appendChild(cell(sc.capexAbs, 'capexAbs'));
    tr.appendChild(cell(sc.opexAbs, 'opexAbs'));
    tr.appendChild(cell(sc.rateAbs, 'rateAbs'));
    // "불러오기"는 Base행에서만 의미가 있다 — Case는 사용자가 직접
    // 값을 채우는 행이라 버튼을 안 둔다.
    var pullTd = document.createElement('td');
    if (sc.name === 'Base') {
      var pull = document.createElement('button');
      pull.type = 'button'; pull.className = 'pull'; pull.title = '위 "사업 기본 가정" 입력값을 그대로 불러와 채우기';
      pull.textContent = '불러오기';
      pull.addEventListener('click', function () { pullFromForm(idx); });
      pullTd.appendChild(pull);
    }
    tr.appendChild(pullTd);
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

  // 그 순간의 위 "사업 기본 가정" 값(판매단가/총사업비/운영비, 금리는
  // 금액이 가장 큰 트랜치의 운영금리를 대표값으로)을 해당 시나리오 행에
  // 그대로 채워넣는다 — Base 행을 실제 숫자로 눈에 보이게 채우고 싶을 때,
  // 또는 다른 행을 "위 값에서 살짝만 바꿔보고 싶을 때" 시작점으로 쓴다.
  function pullFromForm(idx) {
    var row = document.querySelectorAll('#sensBox tbody tr')[idx];
    if (!row) return;
    setVal2(row, 'tariffAbs', $('[data-k="tariff"]').value);
    setVal2(row, 'capexAbs', $('[data-k="capexEok"]').value);
    setVal2(row, 'opexAbs', $('[data-k="opexEok"]').value);
    var tranches = readTranches();
    var maxT = null;
    tranches.forEach(function (t) { if (!maxT || t.amountEok > maxT.amountEok) maxT = t; });
    setVal2(row, 'rateAbs', maxT && maxT.amountEok > 0 ? maxT.rateO : '');
  }
  function setVal2(row, f, v) { var e = row.querySelector('[data-sens-f="' + f + '"]'); if (e) e.value = v; }

  function buildSensGrid() {
    var box = $('#sensBox');
    box.innerHTML = '';
    var t = el('table', 'tr');
    t.innerHTML = '<thead><tr><th>시나리오</th><th>판매단가(원/kWh)</th><th>총사업비(억원)</th><th>운영비(억원)</th><th>금리(%)</th><th></th><th></th></tr></thead>';
    var tb = document.createElement('tbody');
    SENS_ROWS.forEach(function (sc, idx) { tb.appendChild(sensRow(sc, idx)); });
    t.appendChild(tb);
    box.appendChild(t);
  }

  // 빈 칸은 "위 사업 기본가정 값을 그대로 쓴다"는 뜻이라 null로 남긴다
  // (Number('')=0으로 바뀌면 "0으로 강제 지정"과 구분이 안 되므로).
  function readSensRows() {
    var trs = Array.prototype.slice.call(document.querySelectorAll('#sensBox tbody tr'));
    return trs.map(function (tr) {
      var o = {};
      Array.prototype.slice.call(tr.querySelectorAll('input')).forEach(function (input) {
        var f = input.dataset.sensF;
        if (f === 'name') { o[f] = input.value; return; }
        o[f] = input.value === '' ? null : Number(input.value);
      });
      return o;
    });
  }

  // 시나리오 값은 델타(증감)가 아니라 절대값이다 — 빈 칸이면 위 "사업 기본
  // 가정"에 입력한 값을 그대로 쓰고, 채워져 있으면 그 값으로 완전히
  // 대체한다. 판매단가·운영비는 tariffTracks/opexItems가 있을 때 그
  // 항목들도 같은 비율로 같이 조정, 금리는 모든 트랜치의 건설·운영금리를
  // 동일하게 덮어쓴다.
  function applyScenario(baseInp, sc) {
    var c = JSON.parse(JSON.stringify(baseInp));
    if (sc.tariffAbs != null) {
      var tRatio = c.tariff > 0 ? sc.tariffAbs / c.tariff : 1;
      if (c.tariffTracks) c.tariffTracks.forEach(function (t) { t.price = t.price * tRatio; });
      c.tariff = sc.tariffAbs;
    }
    if (sc.capexAbs != null) c.capexEok = sc.capexAbs;
    if (sc.opexAbs != null) {
      var oRatio = c.opexEok > 0 ? sc.opexAbs / c.opexEok : 1;
      if (c.opexItems) c.opexItems.forEach(function (it) { it.annualKRWm = it.annualKRWm * oRatio; });
      c.opexEok = sc.opexAbs;
    }
    if (sc.rateAbs != null && c.tranches) {
      c.tranches.forEach(function (t) { t.rateC = sc.rateAbs; t.rateO = sc.rateAbs; });
    }
    return c;
  }

  function buildBaseInp() {
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
    if (opexDetailOn) {
      baseInp.opexItems = readOpexItemsDetailed().map(function (it) {
        return { name: it.name, annualKRWm: (it.amountEok || 0) * 100, escal: it.escal, senior: it.senior };
      });
    }
    baseInp.shareholders = readShareholders();
    return baseInp;
  }

  function runSensitivity() {
    var wasPreset = usingPreset;
    if (usingPreset) {
      // 실측치(periodOverrides) 기준으로는 판매단가/총사업비/운영비/금리를
      // 바꿔도 결과가 그대로라 민감도 분석 자체가 의미 없다. 막는 대신
      // 프리셋을 해제하고 일반 계산식 기준으로 자동 전환해서 실행한다.
      usingPreset = false;
    }
    var baseInp = buildBaseInp();

    // Base도 다른 행과 똑같은 일반 시나리오 행이다 — 빈 칸이면 applyScenario가
    // 알아서 위 폼 값을 그대로 쓰므로 특별 취급이 필요 없다.
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
    toast(
      (wasPreset ? '프리셋을 해제하고 일반 계산식 기준으로 전환했습니다 — ' : '') +
      '민감도 분석 완료 — ' + results.length + '개 시나리오'
    );
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
      if (d.type === 'display') return; // 계산 결과 표시용일 뿐 입력값이 아님
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
      opexItems: ref.opexItems.map(function (it) { return { name: it.name, annualKRWm: it.annualKRWm, escal: it.escalRate * 100, senior: it.senior }; }),
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

    // 운영비는 실제 항목별 실측치(reference.json opexItems)가 있으니
    // 그대로 채워서 항목별 입력 모드를 켜준다 — 위 opexEok 합계 필드도
    // 이 항목들의 합과 정확히 같은 값이라 서로 어긋나지 않는다.
    OPEX_ITEMS = ref.opexItems.map(function (it) {
      return { name: it.name, amountEok: it.annualKRWm / 100, escal: it.escalRate * 100, senior: it.senior !== false };
    });
    var opexToggleEl = $('#opexDetailToggle');
    if (opexToggleEl) opexToggleEl.checked = true;
    toggleOpexDetail(true);
    // 총사업비는 원본에도 항목별 실측 내역(EPC/감리비 등 세부금액)이
    // 없어서(총액만 존재) 항목별 모드로 켜지 않는다 — 켜면 빈 항목
    // 합계(0)로 총사업비가 덮어써져서 오히려 틀린 값이 됨.
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
    updateCodDisplay();
    toast('예시값(당진1, 100MW급 PJT)을 불러왔습니다 — 바로 "생성"을 눌러보시거나, 숫자를 고쳐가며 시나리오를 만들어보세요');
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
    // 예전엔 "생성"을 다시 누르면 민감도 결과를 지웠는데(기준값이
    // 달라질 수 있다는 이유였음), 순서를 안 지키면 다운로드에서
    // 민감도가 빠져버려 헷갈린다는 피드백으로 제거 — 이제 민감도는
    // "민감도 분석 실행"을 다시 누르기 전까지 그대로 유지되고, 생성을
    // 몇 번을 다시 눌러도 Excel 다운로드에 그대로 딸려간다.
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
      // 운영비 항목별 입력을 켰으면 엔진이 이미 지원하는 opexItems 경로로
      // 넘겨서 계산 자체가 항목별 상승률·선순위/후순위를 반영하게 한다
      // (안 켰으면 기존처럼 총액 근사).
      if (opexDetailOn) {
        inp.opexItems = readOpexItemsDetailed().map(function (it) {
          return { name: it.name, annualKRWm: (it.amountEok || 0) * 100, escal: it.escal, senior: it.senior };
        });
      }
    }
    // 사업자 구성은 프리셋 여부와 무관하게 화면 입력을 그대로 쓴다 —
    // 자본금 총액을 나눠 낸 여러 출자자에게 지분율만큼 배당을 배분하는
    // 표시용 계산이라 원본 검증치(periodOverrides 등)와는 독립적이다.
    inp = Object.assign({}, inp, { shareholders: readShareholders() });
    try {
      model = M.computeModel(inp);
      // 총사업비/운영비 세부 항목은 계산 결과가 아니라 엑셀 표시용
      // 부가정보 — 합계만 입력한 경우에도 엑셀에는 항목 이름이 나오고
      // 금액만 비워두고 싶다는 요청 반영. capexItems는 항상 붙이고
      // (엔진이 쓰는 값이 아니라 순수 표시용), opexDisplayItems는
      // inp.opexItems(실제 계산에 쓰인 항목별 값)가 없을 때만 붙인다 —
      // 있으면 xlsxbuild2.js가 그 실제값으로 이미 항목별 표시를 한다.
      model.capexItems = capexDetailOn ? readCapexItems() : DEFAULT_CAPEX_ITEMS.map(function (n) { return { name: n, amountEok: null }; });
      if (!inp.opexItems) {
        model.opexDisplayItems = DEFAULT_OPEX_ITEMS.map(function (d) { return { name: d.name, amountEok: null }; });
      }
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

  // 예시 불러오기 버튼을 다시 전체 공개로 되돌리면서(빠른 시작용으로
  // 재해석), 민감도 카드의 "프리셋 상태에선 의미 없음" 안내도 같이
  // 보이게 한다 — 이제 누구나 누를 수 있는 버튼이라 그 주의사항도
  // 다시 필요함.
  var sensHint = $('#adminSensHint'); if (sensHint) sensHint.style.display = 'inline';

  buildCore();
  buildTrancheGrid();
  buildShareholderGrid();
  buildSpendCurve();
  updateRecWeightState();
  updateCodDisplay();
  $('#trQuick').addEventListener('click', quickFillTranches);
  $('[data-k="rpsShare"]').addEventListener('input', updateRecWeightState);
  $('[data-k="constructionStart"]').addEventListener('input', updateCodDisplay);
  $('[data-k="constructionMonths"]').addEventListener('input', updateCodDisplay);
  $('[data-k="constructionMonths"]').addEventListener('change', buildSpendCurve);
  $('[data-k="capexEok"]').addEventListener('change', buildSpendCurve);
  $('[data-k="capexEok"]').addEventListener('input', function () { if ($('[data-k="equityRatioPct"]').value !== '') syncEquityRatio(true); });
  $('[data-k="equityEok"]').addEventListener('input', function () { syncEquityRatio(false); });
  $('[data-k="equityRatioPct"]').addEventListener('input', function () { syncEquityRatio(true); });
  $('#spendReset').addEventListener('click', buildSpendCurve);
  $('#spendbox').addEventListener('input', updateSpendSum);
  $('#capexDetailToggle').addEventListener('change', function (e) { toggleCapexDetail(e.target.checked); });
  $('#capexItemBox').addEventListener('input', updateCapexItemSum);
  $('#opexDetailToggle').addEventListener('change', function (e) { toggleOpexDetail(e.target.checked); });
  $('#opexItemBox').addEventListener('input', updateOpexItemSum);
  $('[data-k="capacityMW"]').addEventListener('input', updateOpexMWRefs);
  $('#shAdd').addEventListener('click', function () {
    SHAREHOLDERS.push({ name: '출자자' + (SHAREHOLDERS.length + 1), stakePct: 0 });
    buildShareholderGrid();
  });
  $('#shbox').addEventListener('input', updateShareholderSum);
  buildSensGrid();
  $('#sensAdd').addEventListener('click', function () {
    SENS_ROWS.push({ name: 'Case' + SENS_ROWS.length, tariffAbs: null, capexAbs: null, opexAbs: null, rateAbs: null });
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
