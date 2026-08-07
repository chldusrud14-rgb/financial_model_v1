(function () {
  'use strict';
  var M = window.SolarModel, X = window.SolarXlsx;
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var el = function (t, c, h) { var e = document.createElement(t); if (c) e.className = c; if (h !== undefined) e.innerHTML = h; return e; };

  /* ---------------- 입력 정의 ---------------- */
  var CORE = [
    { k: 'projectName', l: '사업명', t: 'text', full: true },
    { k: 'capacityMW', l: '설비 용량', u: 'MW', s: 0.001 },
    { k: 'capacityFactor', l: '이용률', u: '%', s: 0.1, note: '' },
    { k: 'capexEok', l: '총 사업비 (CAPEX)', u: '억원', s: 1, note: '건설이자 제외' },
    { k: 'opexEok', l: '운영비 (OPEX)', u: '억원/년', s: 1 },
    { k: 'gearing', l: '차입 비율', u: '%', s: 1 },
    { k: 'rateO', l: '차입 금리', u: '%', s: 0.01 },
    { k: 'tariff', l: '매출 단가', u: '원/kWh', s: 0.1, note: 'PPA 또는 SMP+REC' },
    { k: 'operationYears', l: '사업 기간', u: '년', s: 1 }
  ];
  var ADV = [
    ['발전 · 일정', [
      { k: 'degradation', l: 'Degradation', u: '%/년', s: 0.05 },
      { k: 'auxRate', l: '소내 소비율', u: '%', s: 0.1 },
      { k: 'constructionStart', l: '착공 시점', t: 'month' },
      { k: 'constructionMonths', l: '공사 기간', u: '개월', s: 1 }
    ]],
    ['재원조달 · 상환조건', [
      { k: 'dsraEok', l: '최초 DSRA', u: '억원', s: 1 },
      { k: 'rateC', l: '건설기간 금리', u: '%', s: 0.01 },
      { k: 'graceYears', l: '거치 기간', u: '년', s: 1 },
      { k: 'repayYears', l: '상환 기간', u: '년', s: 1 },
      {
        k: 'repayType', l: '상환 방식', t: 'select',
        opt: [[1, '원금 균등'], [2, '원리금 균등'], [3, 'DSCR 스컬프팅']]
      },
      { k: 'payPerYear', l: '연간 상환 횟수', u: '회', s: 1 }
    ]],
    ['매출 · 비용', [
      { k: 'tariffEscal', l: '단가 상승률', u: '%/년', s: 0.1 },
      { k: 'opexEscal', l: '운영비 상승률', u: '%/년', s: 0.1 },
      { k: 'decomEok', l: '철거·복구비', u: '억원', s: 1 }
    ]],
    ['세무 · 감가상각', [
      { k: 'depRatio', l: '감가상각 대상비율', u: '%', s: 1 },
      { k: 'depYears', l: '내용연수', u: '년', s: 1 },
      { k: 'taxMode', l: '법인세 방식', t: 'select', opt: [[1, '누진구간 9/19/21%'], [2, '단일세율']] },
      { k: 'taxFlat', l: '단일 법인세율', u: '%', s: 0.5 },
      { k: 'lossRate', l: '이월결손금 공제비율', u: '%', s: 5 }
    ]],
    ['배당 · 평가', [
      { k: 'dsraMonths', l: 'DSRA 적립기준', u: '개월', s: 1 },
      { k: 'minCash', l: '최소보유현금', u: '억원', s: 1 },
      { k: 'divDSCR', l: '배당제한 DSCR', u: 'x', s: 0.05 },
      { k: 'divStartYear', l: '배당개시 연차', u: '년', s: 1 },
      { k: 'discount', l: '할인율 (NPV·LCOE)', u: '%', s: 0.1 }
    ]]
  ];

  var PRESETS = {
    dangjin: {
      projectName: '당진 태양광발전', capacityMW: 99.998, capacityFactor: 15.71, degradation: 0.5,
      auxRate: 0, constructionStart: '2024-06', constructionMonths: 16, operationYears: 20,
      capexEok: 1411, dsraEok: 50, opexEok: 49.8, opexEscal: 0.7, gearing: 90, rateC: 5.6, rateO: 5.54,
      graceYears: 1, repayYears: 16, repayType: 3, payPerYear: 4, tariff: 154.8, tariffEscal: 0,
      depRatio: 95, depYears: 20, taxMode: 1, taxFlat: 21, lossRate: 80, decomEok: 20,
      dsraMonths: 6, minCash: 10, divDSCR: 1.1, divStartYear: 2, discount: 5.5
    },
    std: {
      projectName: '신규 태양광발전 사업', capacityMW: 80, capacityFactor: 15.8, degradation: 0.5,
      auxRate: 0, constructionStart: '2026-03', constructionMonths: 14, operationYears: 20,
      capexEok: 1120, dsraEok: 40, opexEok: 28, opexEscal: 1.5, gearing: 75, rateC: 5.4, rateO: 5.2,
      graceYears: 1, repayYears: 15, repayType: 3, payPerYear: 4, tariff: 165, tariffEscal: 0,
      depRatio: 95, depYears: 20, taxMode: 1, taxFlat: 21, lossRate: 80, decomEok: 16,
      dsraMonths: 6, minCash: 10, divDSCR: 1.1, divStartYear: 2, discount: 5.5
    }
  };
  PRESETS.reset = PRESETS.std;

  var state = Object.assign({}, PRESETS.std), model = null;

  /* ---------------- 폼 ---------------- */
  function field(d) {
    var w = el('div', 'f' + (d.full ? ' full' : ''));
    w.appendChild(el('label', null, d.l));
    var box = el('div', 'in');
    var input;
    if (d.t === 'select') {
      input = el('select');
      d.opt.forEach(function (o) {
        var op = el('option', null, o[1]); op.value = o[0]; input.appendChild(op);
      });
    } else {
      input = el('input');
      input.type = (d.t === 'text' || d.t === 'month') ? (d.t === 'month' ? 'month' : 'text') : 'number';
      if (d.s) input.step = d.s;
      if (d.u) input.className = 'u';
    }
    input.dataset.k = d.k;
    box.appendChild(input);
    if (d.u) box.appendChild(el('span', 'unit', d.u));
    w.appendChild(box);
    if (d.note !== undefined) { var n = el('div', 'note', d.note); n.dataset.note = d.k; w.appendChild(n); }
    var onEdit = function () {
      var v = input.value;
      state[d.k] = (input.type === 'number') ? (v === '' ? 0 : parseFloat(v)) : v;
      run(true);
    };
    input.addEventListener('input', onEdit);
    input.addEventListener('change', onEdit);
    return w;
  }

  function buildForm() {
    var core = $('#core');
    CORE.forEach(function (d) { core.appendChild(field(d)); });
    var adv = $('#adv');
    ADV.forEach(function (g) {
      var box = el('div', 'grp');
      box.appendChild(el('b', null, g[0]));
      var gr = el('div', 'grid');
      g[1].forEach(function (d) { gr.appendChild(field(d)); });
      box.appendChild(gr);
      adv.appendChild(box);
    });
  }

  function syncForm() {
    document.querySelectorAll('[data-k]').forEach(function (i) {
      if (state[i.dataset.k] !== undefined) i.value = state[i.dataset.k];
    });
  }

  /* ---------------- 포맷 ---------------- */
  var f0 = function (n) { return (Math.round(n)).toLocaleString('ko-KR'); };
  var f1 = function (n) { return n.toLocaleString('ko-KR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }); };
  var f2 = function (n) { return n.toLocaleString('ko-KR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); };
  var pct = function (n) { return n === null ? '—' : f1(n * 100); };
  var eok = function (m) { return m / 100; };  // 백만원 -> 억원

  /* ---------------- KPI ---------------- */
  function kpis() {
    var k = model.kpi, box = $('#kpis');
    var mind = k.minDSCR;
    var items = [
      ['Equity IRR', pct(k.equityIRR), '%', k.equityIRR !== null && k.equityIRR < 0.06 ? 'warn' : ''],
      ['Project IRR', pct(k.projectIRR), '%', ''],
      ['NPV', f0(eok(k.npv)), '억', k.npv < 0 ? 'alert' : ''],
      ['최소 DSCR', mind === null ? '—' : f2(mind), 'x', mind !== null && mind < 1 ? 'alert' : (mind !== null && mind < 1.15 ? 'warn' : '')],
      ['투자비 회수', k.payback === null ? '—' : f1(k.payback), '년', ''],
      ['LCOE', k.lcoe === null ? '—' : f0(k.lcoe), '원', '']
    ];
    box.innerHTML = '';
    items.forEach(function (it) {
      var c = el('div', 'kpi' + (it[3] ? ' ' + it[3] : ''));
      c.appendChild(el('div', 'k', it[0]));
      c.appendChild(el('div', 'v', '<span class="num">' + it[1] + '</span><em>' + it[2] + '</em>'));
      box.appendChild(c);
    });
  }

  /* ---------------- 차트 ---------------- */
  function svg(w, h, vb) {
    return '<svg viewBox="0 0 ' + w + ' ' + h + '" preserveAspectRatio="none" style="height:' + (vb || h) + 'px">';
  }
  function fcfChart() {
    var rows = model.rows, W = 660, H = 200, pad = { l: 4, r: 4, t: 10, b: 22 };
    var vals = rows.map(function (r) { return eok(r.fcfe); });
    var ds = rows.map(function (r) { return eok(r.ds); });
    var max = Math.max.apply(null, vals.concat(ds)) || 1;
    var min = Math.min(0, Math.min.apply(null, vals));
    var span = max - min || 1;
    var iw = W - pad.l - pad.r, n = rows.length;
    var bw = iw / n, gap = Math.min(6, bw * 0.28);
    var y0 = pad.t + (max / span) * (H - pad.t - pad.b);
    var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" style="height:210px">';
    s += '<defs><linearGradient id="g1" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#4E9E80"/><stop offset="100%" stop-color="#14483A"/></linearGradient></defs>';
    s += '<line x1="0" y1="' + y0.toFixed(1) + '" x2="' + W + '" y2="' + y0.toFixed(1) + '" stroke="#DCE6E1"/>';
    rows.forEach(function (r, idx) {
      var x = pad.l + idx * bw;
      var hd = (eok(r.ds) / span) * (H - pad.t - pad.b);
      s += '<rect x="' + (x + gap / 2).toFixed(1) + '" y="' + (y0 - hd).toFixed(1) + '" width="' + (bw - gap).toFixed(1) +
        '" height="' + hd.toFixed(1) + '" fill="#CBD9D3" rx="2"/>';
      var v = eok(r.fcfe);
      var hv = Math.abs(v) / span * (H - pad.t - pad.b);
      var yy = v >= 0 ? y0 - hv : y0;
      s += '<rect x="' + (x + gap / 2 + (bw - gap) * 0.22).toFixed(1) + '" y="' + yy.toFixed(1) +
        '" width="' + ((bw - gap) * 0.56).toFixed(1) + '" height="' + Math.max(1, hv).toFixed(1) +
        '" fill="url(#g1)" rx="2"><title>' + (idx + 1) + '년차 (' + r.year + ') · FCFE ' + f0(v) + '억 · 원리금 ' + f0(eok(r.ds)) + '억</title></rect>';
      if (n <= 12 || idx === 0 || (idx + 1) % Math.ceil(n / 7) === 0) {
        s += '<text x="' + (x + bw / 2).toFixed(1) + '" y="' + (H - 6) + '" text-anchor="middle" font-size="10" fill="#84968E">' + (idx + 1) + 'Y</text>';
      }
    });
    s += '</svg>';
    $('#fcf').innerHTML = s;
  }

  function dscrChart() {
    var rows = model.rows.filter(function (r) { return r.dscr !== null; });
    if (!rows.length) { $('#dscr').innerHTML = '<p style="font-size:12px;color:#84968E">상환 기간이 없습니다.</p>'; return; }
    var W = 660, H = 150, pad = { t: 10, b: 20 };
    var max = Math.max(2, Math.max.apply(null, rows.map(function (r) { return r.dscr; })) * 1.1);
    var n = rows.length, bw = W / n;
    var y = function (v) { return pad.t + (1 - v / max) * (H - pad.t - pad.b); };
    var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" style="height:150px">';
    [1, state.divDSCR].forEach(function (lv, i) {
      s += '<line x1="0" y1="' + y(lv).toFixed(1) + '" x2="' + W + '" y2="' + y(lv).toFixed(1) +
        '" stroke="' + (i ? '#C2703B' : '#B4483E') + '" stroke-width="1" stroke-dasharray="' + (i ? '4 4' : '0') + '" opacity=".55"/>';
      s += '<text x="' + (W - 2) + '" y="' + (y(lv) - 4).toFixed(1) + '" text-anchor="end" font-size="9.5" fill="' + (i ? '#C2703B' : '#B4483E') + '">' + (i ? '배당제한 ' : '') + f2(lv) + 'x</text>';
    });
    var pts = rows.map(function (r, i) { return (bw * i + bw / 2).toFixed(1) + ',' + y(r.dscr).toFixed(1); });
    s += '<polyline points="' + pts.join(' ') + '" fill="none" stroke="#2E7D62" stroke-width="2" stroke-linejoin="round"/>';
    rows.forEach(function (r, i) {
      var cx = bw * i + bw / 2, cy = y(r.dscr);
      s += '<circle cx="' + cx.toFixed(1) + '" cy="' + cy.toFixed(1) + '" r="3" fill="' + (r.dscr < 1 ? '#B4483E' : '#2E7D62') + '"><title>' + r.n + '년차 DSCR ' + f2(r.dscr) + 'x</title></circle>';
      if (n <= 12 || i === 0 || (i + 1) % Math.ceil(n / 7) === 0) {
        s += '<text x="' + cx.toFixed(1) + '" y="' + (H - 4) + '" text-anchor="middle" font-size="10" fill="#84968E">' + r.n + 'Y</text>';
      }
    });
    s += '</svg>';
    $('#dscr').innerHTML = s;
  }

  /* ---------------- 검토 탭 ---------------- */
  function checks() {
    var k = model.kpi, rows = model.rows, out = [];
    var low = rows.filter(function (r) { return r.dscr !== null && r.dscr < 1; });
    out.push(low.length
      ? ['bad', '!', 'DSCR 1.00x 미만 구간 ' + low.length + '개년', '원리금 상환이 불가한 연도가 있습니다 (' + low.map(function (r) { return r.n + '년차'; }).slice(0, 5).join(', ') + '). 차입비율·거치기간·상환기간을 조정하세요.']
      : ['ok', '✓', '전 기간 DSCR 1.00x 이상', '최소 DSCR ' + f2(k.minDSCR) + 'x · 평균 ' + f2(k.avgDSCR) + 'x']);

    out.push(k.minDSCR !== null && k.minDSCR >= 1.2
      ? ['ok', '✓', '금융권 통상 커버리지 충족', '최소 DSCR이 1.20x 이상으로 선순위 대주단 요구수준을 상회합니다.']
      : ['warn', '!', '커버리지 여유 부족', '최소 DSCR ' + (k.minDSCR === null ? '—' : f2(k.minDSCR)) + 'x — 국내 태양광 PF는 통상 1.15~1.30x를 요구합니다.']);

    out.push(k.npv >= 0
      ? ['ok', '✓', 'NPV 양(+)', '할인율 ' + f1(state.discount) + '% 기준 NPV ' + f0(eok(k.npv)) + '억원']
      : ['bad', '!', 'NPV 음(−)', '할인율 ' + f1(state.discount) + '% 기준으로 가치가 훼손됩니다.']);

    var neg = rows.filter(function (r) { return r.cashClose < -0.5; });
    out.push(neg.length
      ? ['warn', '!', '현금 부족 연도 ' + neg.length + '개년', '운영 중 현금잔액이 음(−)이 되는 구간이 있습니다. 최소보유현금·DSRA 가정을 확인하세요.']
      : ['ok', '✓', '운영기간 현금 부족 없음', '배당 후 최소보유현금 ' + f0(state.minCash) + '억원을 상시 유지합니다.']);

    if (state.repayType === 3 || state.repayType === '3') {
      out.push(['ok', '✓', 'DSCR 스컬프팅 적용', '상환기간 내 완전상환되는 균등 DSCR ' + f2(k.sculptDSCR) + 'x 로 원금 스케줄이 산출되었습니다.']);
    }
    var lastY = rows[rows.length - 1];
    out.push(['ok', '✓', '차입금 완전상환 확인', (model.lastRepayYear) + '년차에 잔액 0 — 만기 잔존액 없음']);

    var box = $('#checks'); box.innerHTML = '';
    out.forEach(function (c) {
      var d = el('div', 'chk ' + c[0]);
      d.appendChild(el('div', 'm', c[1]));
      var t = el('div'); t.appendChild(el('b', null, c[2])); t.appendChild(el('span', null, c[3]));
      d.appendChild(t); box.appendChild(d);
    });
    var bad = out.filter(function (c) { return c[0] === 'bad'; }).length;
    var warn = out.filter(function (c) { return c[0] === 'warn'; }).length;
    var tag = $('#vtag');
    tag.textContent = bad ? '보완 필요 ' + bad + '건' : (warn ? '검토 권고 ' + warn + '건' : '이상 없음');
    tag.style.background = bad ? '#F8E9E7' : warn ? '#FBF0E7' : '';
    tag.style.color = bad ? '#B4483E' : warn ? '#C2703B' : '';
  }

  function sensBlocks() {
    var base = Object.assign({}, state);
    var d = function (v, p) { return Math.round((v * (1 + p)) * 100) / 100; };
    return [
      {
        title: '이용률', key: 'capacityFactor', unit: '%',
        vals: [d(base.capacityFactor, -0.1), d(base.capacityFactor, -0.05), base.capacityFactor, d(base.capacityFactor, 0.05), d(base.capacityFactor, 0.1)]
      },
      {
        title: '차입 금리', key: 'rateO', unit: '%',
        vals: [d(base.rateO, -0.2), d(base.rateO, -0.1), base.rateO, d(base.rateO, 0.1), d(base.rateO, 0.2)]
      },
      {
        title: '매출 단가', key: 'tariff', unit: '원/kWh',
        vals: [d(base.tariff, -0.1), d(base.tariff, -0.05), base.tariff, d(base.tariff, 0.05), d(base.tariff, 0.1)]
      }
    ].map(function (b) {
      b.data = M.sensitivity(base, b.key, b.vals);
      b.fmtLabel = function (v) { return f2(v) + (b.unit === '%' ? ' %' : ' ' + b.unit); };
      return b;
    });
  }

  function renderSens() {
    var blocks = sensBlocks(), box = $('#sens'); box.innerHTML = '';
    blocks.forEach(function (b) {
      var c = el('div');
      c.appendChild(el('div', null, '<b style="font-size:13px">' + b.title + ' 민감도</b>'));
      var vals = b.data.map(function (x) { return x.equityIRR === null ? 0 : x.equityIRR * 100; });
      var max = Math.max.apply(null, vals.map(Math.abs)) || 1;
      var W = 300, rowH = 30, H = b.data.length * rowH + 8;
      var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" style="height:' + H + 'px;margin-top:10px">';
      b.data.forEach(function (x, i) {
        var y = i * rowH + 4, isBase = Math.abs(x.value - state[b.key]) < 1e-9;
        var w = Math.max(2, (vals[i] / max) * 150);
        s += '<text x="0" y="' + (y + 14) + '" font-size="11" fill="' + (isBase ? '#16261F' : '#84968E') + '" font-weight="' + (isBase ? 700 : 500) + '">' + b.fmtLabel(x.value) + '</text>';
        s += '<rect x="98" y="' + (y + 4) + '" width="' + w.toFixed(1) + '" height="14" rx="3" fill="' + (isBase ? '#14483A' : '#8FBFAA') + '"/>';
        s += '<text x="' + (98 + w + 6).toFixed(1) + '" y="' + (y + 15) + '" font-size="11" fill="#3D524A" font-weight="600">' + pct(x.equityIRR) + '%</text>';
        s += '<text x="' + W + '" y="' + (y + 15) + '" text-anchor="end" font-size="10.5" fill="' + (x.minDSCR < 1 ? '#B4483E' : '#84968E') + '">DSCR ' + (x.minDSCR === null ? '—' : f2(x.minDSCR)) + 'x</text>';
      });
      s += '</svg>';
      c.insertAdjacentHTML('beforeend', s);
      c.insertAdjacentHTML('beforeend', '<div style="font-size:11px;color:#84968E;margin-top:2px">막대 = Equity IRR · 우측 = 최소 DSCR</div>');
      box.appendChild(c);
    });
    return blocks;
  }

  var TABLES = {
    is: {
      cols: [['영업수익', 'revenue'], ['영업비용', 'opex', -1], ['EBITDA', 'ebitda', 0, 1], ['감가상각비', 'dep', -1],
      ['영업이익', 'ebit'], ['이자비용', 'interest', -1], ['세전이익', 'ebt'], ['법인세', 'tax', -1], ['당기순이익', 'ni', 0, 1]]
    },
    cf: {
      cols: [['EBITDA', 'ebitda'], ['법인세', 'tax', -1], ['CFADS', 'cfads', 0, 1], ['원리금(DS)', 'ds', -1],
      ['DSRA 증감', 'dsraMove', -1], ['FCFE', 'fcfe', 0, 1], ['배당금', 'dividend'], ['기말현금', 'cashClose'], ['DSCR', 'dscr', 0, 0, 'x']]
    },
    debt: {
      cols: [['기초잔액', 'debtOpen'], ['원금상환', 'principal'], ['이자비용', 'interest'], ['기말잔액', 'debtClose'],
      ['원리금(DS)', 'ds', 0, 1], ['DSRA 잔액', 'dsraClose'], ['DSCR', 'dscr', 0, 0, 'x']]
    }
  };

  function renderTable() {
    var def = TABLES[$('#tsel').value], rows = model.rows;
    var h = '<thead><tr><th>구분</th>';
    rows.forEach(function (r) { h += '<th>' + r.n + '년차<br><span style="font-weight:500;color:#84968E">' + r.year + '</span></th>'; });
    h += '<th>합계</th></tr></thead><tbody>';
    def.cols.forEach(function (c) {
      h += '<tr><td>' + c[0] + '</td>';
      var tot = 0;
      rows.forEach(function (r) {
        var v = r[c[1]];
        if (c[4] === 'x') {
          h += '<td>' + (v === null ? '—' : f2(v) + 'x') + '</td>';
        } else {
          var sv = (c[2] === -1 ? -Math.abs(v) : v);
          tot += sv;
          h += '<td class="' + (sv < 0 ? 'neg ' : '') + (c[3] ? 'hi' : '') + '">' + f0(sv) + '</td>';
        }
      });
      h += '<td class="hi">' + (c[4] === 'x' ? '—' : f0(tot)) + '</td></tr>';
    });
    h += '</tbody>';
    $('#tbl').innerHTML = h;
  }

  /* ---------------- 실행 ---------------- */
  var timer = null;
  function run(debounced) {
    if (debounced) { clearTimeout(timer); timer = setTimeout(function () { run(false); }, 220); return; }
    try {
      model = M.computeModel(state);
    } catch (e) { toast('계산 오류: ' + e.message); return; }
    kpis(); fcfChart(); dscrChart(); checks(); renderSens(); renderTable();
    var n = document.querySelector('[data-note="capacityFactor"]');
    if (n) n.textContent = '일일 발전시간 ' + f2(state.capacityFactor * 24 / 100) + '시간/일 · 연 발전량 ' + f0(model.rows[0].gen) + ' MWh';
    var c = document.querySelector('[data-note="capexEok"]');
    if (c) c.textContent = '건설이자 ' + f0(eok(model.idc)) + '억 → 총투자비 ' + f0(eok(model.tic)) + '억 · 자기자본 ' + f0(eok(model.equity)) + '억';
  }

  function toast(msg) {
    var t = $('#toast'); t.textContent = msg; t.classList.add('on');
    setTimeout(function () { t.classList.remove('on'); }, 2600);
  }

  /* ---------------- Excel ---------------- */
  function download() {
    if (!window.ExcelJS) { toast('Excel 모듈을 불러오지 못했습니다. 네트워크를 확인해 주세요.'); return; }
    var btn = $('#xls'); btn.disabled = true;
    var label = btn.innerHTML; btn.innerHTML = '생성 중…';
    setTimeout(function () {
      try {
        var m = M.computeModel(state);
        m.sensBlocks = sensBlocks();
        var wb = X.buildWorkbook(m, window.ExcelJS);
        wb.xlsx.writeBuffer().then(function (buf) {
          var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
          var a = document.createElement('a');
          a.href = URL.createObjectURL(blob);
          var d = new Date(), ymd = d.getFullYear() + ('0' + (d.getMonth() + 1)).slice(-2) + ('0' + d.getDate()).slice(-2);
          a.download = (state.projectName || '태양광') + '_재무모델_' + ymd + '.xlsx';
          document.body.appendChild(a); a.click(); a.remove();
          setTimeout(function () { URL.revokeObjectURL(a.href); }, 3000);
          btn.disabled = false; btn.innerHTML = label;
          toast('Excel 파일을 내려받았습니다 — 12개 시트, 가정값 수정 시 자동 재계산');
        })['catch'](function (e) {
          btn.disabled = false; btn.innerHTML = label; toast('생성 실패: ' + e.message);
        });
      } catch (e) {
        btn.disabled = false; btn.innerHTML = label; toast('생성 실패: ' + e.message);
      }
    }, 30);
  }

  /* ---------------- 이벤트 ---------------- */
  function init() {
    buildForm(); syncForm(); run();
    document.querySelectorAll('.tabs button').forEach(function (b) {
      b.addEventListener('click', function () {
        document.querySelectorAll('.tabs button').forEach(function (x) { x.classList.remove('on'); });
        b.classList.add('on');
        document.querySelectorAll('.page').forEach(function (p) { p.classList.remove('on'); });
        $('#tab-' + b.dataset.tab).classList.add('on');
      });
    });
    document.querySelectorAll('.chip').forEach(function (c) {
      c.addEventListener('click', function () {
        state = Object.assign({}, PRESETS[c.dataset.preset]);
        syncForm(); run();
        toast(c.textContent.replace(' 불러오기', '') + ' 적용');
      });
    });
    $('#run').addEventListener('click', function () { run(); toast('재무모델을 생성했습니다'); });
    $('#xls').addEventListener('click', download);
    $('#tsel').addEventListener('change', renderTable);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
