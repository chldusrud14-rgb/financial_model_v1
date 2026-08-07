/* ============================================================
   태양광 PF 재무모델 엔진
   - 단위: 내부 계산은 백만원(KRWm), 입력 일부는 억원
   - 기간: 건설기간(월 단위) + 운영 N개년(연 단위)
   ============================================================ */
(function (global) {
  'use strict';

  var EOK = 100; // 1억원 = 100 백만원

  function num(v, d) { v = parseFloat(v); return isFinite(v) ? v : (d || 0); }

  /* ---------- 재무 함수 ---------- */
  function npv(rate, flows) { // flows[0] = t0
    var s = 0;
    for (var t = 0; t < flows.length; t++) s += flows[t] / Math.pow(1 + rate, t);
    return s;
  }

  function irr(flows) {
    var pos = false, neg = false, i;
    for (i = 0; i < flows.length; i++) { if (flows[i] > 0) pos = true; if (flows[i] < 0) neg = true; }
    if (!pos || !neg) return null;
    var lo = -0.9499, hi = 2.0, mid, f;
    var flo = npv(lo, flows), fhi = npv(hi, flows);
    if (flo * fhi > 0) return null;
    for (i = 0; i < 200; i++) {
      mid = (lo + hi) / 2;
      f = npv(mid, flows);
      if (flo * f <= 0) { hi = mid; fhi = f; } else { lo = mid; flo = f; }
    }
    return (lo + hi) / 2;
  }

  function pmt(rate, nper, pv) {
    if (nper <= 0) return 0;
    if (rate === 0) return pv / nper;
    return pv * rate / (1 - Math.pow(1 + rate, -nper));
  }

  /* 법인세: 누진 구간 (단위 백만원) */
  function corpTax(base, mode, flatRate) {
    if (base <= 0) return 0;
    if (mode !== 1) return base * flatRate;
    if (base <= 200) return base * 0.09;
    if (base <= 20000) return 18 + (base - 200) * 0.19;
    return 3780 + (base - 20000) * 0.21;
  }

  function addMonths(ym, m) { // 'YYYY-MM' + months
    var p = String(ym).split('-');
    var y = parseInt(p[0], 10), mo = parseInt(p[1], 10) - 1 + m;
    y += Math.floor(mo / 12); mo = ((mo % 12) + 12) % 12;
    return { y: y, m: mo + 1, label: y + '-' + ('0' + (mo + 1)).slice(-2) };
  }

  /* ============================================================
     메인 계산
     ============================================================ */
  function computeModel(inp) {
    var i = {
      projectName: inp.projectName || '태양광 발전사업',
      capacityMW: num(inp.capacityMW, 100),
      capacityFactor: num(inp.capacityFactor, 15.7) / 100,
      degradation: num(inp.degradation, 0.5) / 100,
      auxRate: num(inp.auxRate, 0) / 100,
      constructionStart: inp.constructionStart || '2026-01',
      constructionMonths: Math.max(1, num(inp.constructionMonths, 16)),
      operationYears: Math.max(1, Math.min(40, Math.round(num(inp.operationYears, 20)))),
      capexEok: num(inp.capexEok, 1410),
      dsraEok: num(inp.dsraEok, 50),
      opexEok: num(inp.opexEok, 50),
      opexEscal: num(inp.opexEscal, 1.5) / 100,
      gearing: num(inp.gearing, 90) / 100,
      rateC: num(inp.rateC, 5.6) / 100,
      rateO: num(inp.rateO, 5.5) / 100,
      graceYears: num(inp.graceYears, 1),
      repayYears: Math.max(1, num(inp.repayYears, 16)),
      repayType: num(inp.repayType, 1),      // 1 원금균등 / 2 원리금균등 / 3 DSCR 스컬프팅
      targetDSCR: Math.max(1.0, num(inp.targetDSCR, 1.2)),
      payPerYear: Math.max(1, num(inp.payPerYear, 4)),
      minCash: num(inp.minCash, 10) * EOK,   // 배당 후 최소보유현금
      divDSCR: num(inp.divDSCR, 1.1),        // 배당제한 DSCR
      divStartYear: Math.max(1, num(inp.divStartYear, 2)),  // 배당개시 연차(Lock-up)
      dsraMonths: num(inp.dsraMonths, 6),    // DSRA 적립기준 개월수
      tariff: num(inp.tariff, 154.8),
      tariffEscal: num(inp.tariffEscal, 0) / 100,
      depRatio: num(inp.depRatio, 95) / 100,
      depYears: Math.max(1, num(inp.depYears, 20)),
      taxMode: num(inp.taxMode, 1),          // 1 누진구간 / 2 단일세율
      taxFlat: num(inp.taxFlat, 21) / 100,
      lossRate: num(inp.lossRate, 80) / 100,
      decomEok: num(inp.decomEok, 20),
      discount: num(inp.discount, 5.5) / 100
    };

    var N = i.operationYears;
    var capex0 = i.capexEok * EOK;
    var dsra0 = i.dsraEok * EOK;
    var decom = i.decomEok * EOK;

    /* ---- 1. 건설기간 인출 & 건설이자(IDC) ----
       자기자본 = (총사업비 + 최초DSRA) x (1 - 차입비율)  → 우선 인출
       건설이자는 전액 차입으로 조달·자본화 (순환참조 없음 = Excel에서 재계산 가능) */
    var equity = (capex0 + dsra0) * (1 - i.gearing);
    var monthly = capex0 / i.constructionMonths;
    var eqLeft = equity, bal = 0, idc = 0, m;
    var cSched = [];
    for (m = 1; m <= i.constructionMonths; m++) {
      var draw = monthly;
      var eqDraw = Math.min(eqLeft, draw); eqLeft -= eqDraw;
      var dtDraw = draw - eqDraw;
      var open = bal;
      var it0 = (open + dtDraw) * i.rateC / 12;
      idc += it0; bal = open + dtDraw + it0;
      cSched.push({ m: m, cost: draw, eq: eqDraw, dt: dtDraw, open: open, interest: it0, close: bal });
    }
    var dsraEq = Math.min(eqLeft, dsra0); eqLeft -= dsraEq;
    var dsraDt = dsra0 - dsraEq;
    bal += dsraDt;
    var debt = bal;
    var tic = capex0 + idc;              // 총투자비(건설이자 포함)
    var totalFunding = tic + dsra0;      // 소요자금 합계 = 자기자본 + 차입금
    var actualGearing = totalFunding > 0 ? debt / totalFunding : 0;

    var cod = addMonths(i.constructionStart, i.constructionMonths);
    var startYear = parseInt(String(i.constructionStart).split('-')[0], 10);

    /* ---- 2~3. 차입금 상환 + 손익 + 현금흐름 ---- */
    var P = i.payPerYear;
    var annuityPay = pmt(i.rateO, i.repayYears, debt);
    var depBase = tic * i.depRatio;
    var depAnnual = depBase / i.depYears;
    var lastRepayYear = Math.min(N, i.graceYears + i.repayYears);

    function runOps(targetD) {
      var rows = [], lossCF = 0, balOpen = debt, cashOpen = 0;
      for (var y = 1; y <= N; y++) {
        var gen = i.capacityMW * 8760 * i.capacityFactor * (1 - i.auxRate) *
          Math.pow(1 - i.degradation, y - 1);                       // MWh
        var price = i.tariff * Math.pow(1 + i.tariffEscal, y - 1);   // 원/kWh
        var revenue = gen * price / 1000;                            // 백만원
        var opex = i.opexEok * EOK * Math.pow(1 + i.opexEscal, y - 1);
        var ebitda = revenue - opex;
        var dep = (y <= i.depYears) ? depAnnual : 0;
        var inRepay = (y > i.graceYears) && (y <= i.graceYears + i.repayYears);
        var isLastRepay = (y === lastRepayYear);

        /* 원금·이자·법인세 상호의존 → 반복 수렴 */
        var prin = 0, interest = 0, tax = 0, ebt = 0, deduct = 0, taxBase = 0, cfads = 0, lossEnd = lossCF;
        for (var k2 = 0; k2 < 20; k2++) {
          interest = Math.max(0, balOpen - prin * (P - 1) / (2 * P)) * i.rateO;
          ebt = ebitda - dep - interest;
          deduct = Math.min(Math.max(0, ebt) * i.lossRate, lossCF);
          taxBase = Math.max(0, ebt - deduct);
          tax = corpTax(taxBase, i.taxMode, i.taxFlat);
          lossEnd = lossCF - deduct + Math.max(0, -ebt);
          cfads = ebitda - tax;

          var np = 0;
          if (inRepay) {
            if (i.repayType === 3) np = Math.max(0, cfads / targetD - interest);
            else if (i.repayType === 2) np = Math.max(0, annuityPay - balOpen * i.rateO);
            else np = debt / i.repayYears;
            if (isLastRepay && i.repayType !== 3) np = balOpen;
            np = Math.min(np, balOpen);
          }
          if (Math.abs(np - prin) < 1e-8) { prin = np; break; }
          prin = np;
        }
        lossCF = lossEnd;

        var balClose = balOpen - prin;
        var ds = prin + interest;
        var dscr = ds > 0.0001 ? cfads / ds : null;
        var decomOut = (y === N) ? decom : 0;

        rows.push({
          n: y, year: cod.y + y - 1, gen: gen, price: price, revenue: revenue, opex: opex,
          ebitda: ebitda, ebitdaMargin: revenue ? ebitda / revenue : 0, dep: dep, ebit: ebitda - dep,
          interest: interest, principal: prin, debtOpen: balOpen, debtClose: balClose,
          ebt: ebt, deduct: deduct, taxBase: taxBase, tax: tax, ni: ebt - tax, lossCF: lossCF,
          cfads: cfads, ds: ds, dscr: dscr, decom: decomOut,
          projectFcf: ebitda - tax - decomOut
        });
        balOpen = balClose;
      }

      /* DSRA(차기 원리금 X개월분) → 현금계정 → 배당 */
      var dsraPrev = dsra0, cash = 0;
      for (var z = 0; z < rows.length; z++) {
        var r0 = rows[z];
        var nextDS = (z + 1 < rows.length) ? rows[z + 1].ds : 0;
        var dsraTarget = (z + 1 >= lastRepayYear) ? 0 : nextDS * i.dsraMonths / 12;
        var dsraMove = dsraTarget - dsraPrev;          // (+)적립 (-)환입
        r0.dsraOpen = dsraPrev; r0.dsraClose = dsraTarget; r0.dsraMove = dsraMove;
        r0.fcfe = r0.cfads - r0.ds - dsraMove - r0.decom;

        var cashBefore = cash + r0.fcfe;
        var lockup = (r0.n < i.divStartYear);
        var allowed = !lockup && (r0.dscr === null || r0.dscr >= i.divDSCR);
        var reserve = (r0.n === N) ? 0 : i.minCash;
        var dividend = allowed ? Math.max(0, cashBefore - reserve) : 0;
        r0.cashOpen = cash; r0.dividend = dividend; r0.cashClose = cashBefore - dividend;
        cash = r0.cashClose; dsraPrev = dsraTarget;
      }
      return { rows: rows, endBalance: rows[lastRepayYear - 1].debtClose };
    }

    /* 스컬프팅: 상환기간 내 완전상환이 되는 균등 DSCR을 역산 */
    var sculptDSCR = i.targetDSCR, res;
    if (i.repayType === 3) {
      var lo2 = 1.0, hi2 = 8.0, mid2;
      if (runOps(lo2).endBalance > 1e-6) { sculptDSCR = lo2; }
      else {
        for (var b = 0; b < 60; b++) {
          mid2 = (lo2 + hi2) / 2;
          if (runOps(mid2).endBalance > 1e-6) hi2 = mid2; else lo2 = mid2;
        }
        sculptDSCR = lo2;
      }
      res = runOps(sculptDSCR);
    } else {
      res = runOps(i.targetDSCR);
    }

    var rows = res.rows, cumProject = -tic, cumEquity = -equity;
    var pvOpex = 0, pvGen = 0;
    rows.forEach(function (r) {
      cumProject += r.projectFcf; cumEquity += r.fcfe;
      r.cumProject = cumProject; r.cumEquity = cumEquity;
      var df = Math.pow(1 + i.discount, r.n);
      pvOpex += r.opex / df; pvGen += r.gen / df;
    });
    var balloon = 0;

    /* ---- 4. 지표 ---- */
    var projFlows = [-tic], eqFlows = [-equity], preTaxFlows = [-tic], divFlows = [-equity];
    rows.forEach(function (r) {
      projFlows.push(r.projectFcf);
      preTaxFlows.push(r.ebitda - r.decom);
      eqFlows.push(r.fcfe);
      divFlows.push(r.dividend);
    });

    var dscrs = rows.filter(function (r) { return r.dscr !== null && r.ds > 0.0001; })
      .map(function (r) { return r.dscr; });
    var minDSCR = dscrs.length ? Math.min.apply(null, dscrs) : null;
    var avgDSCR = dscrs.length ? dscrs.reduce(function (a, b) { return a + b; }, 0) / dscrs.length : null;

    // 투자비 회수기간 (Project FCF 누적 기준, 선형보간)
    var payback = null, prev = -tic;
    for (var q = 0; q < rows.length; q++) {
      var cum = rows[q].cumProject;
      if (cum >= 0) { payback = q + (prev < 0 ? (-prev) / (cum - prev) : 0); break; }
      prev = cum;
    }

    var lcoe = pvGen > 0 ? (tic + pvOpex) * 1e6 / (pvGen * 1000) : null;
    var totalRev = rows.reduce(function (a, r) { return a + r.revenue; }, 0);
    var totalEbitda = rows.reduce(function (a, r) { return a + r.ebitda; }, 0);

    return {
      input: i, rows: rows,
      tic: tic, capex0: capex0, idc: idc, dsra0: dsra0, totalFunding: totalFunding,
      debt: debt, equity: equity, depBase: depBase, depAnnual: depAnnual,
      cod: cod, startYear: startYear, lastRepayYear: lastRepayYear, decom: decom,
      annuityPay: annuityPay, cSched: cSched, actualGearing: actualGearing,
      dsraEq: dsraEq, dsraDt: dsraDt, monthlyCost: monthly,
      kpi: {
        equityIRR: irr(eqFlows),
        dividendIRR: irr(divFlows),
        sculptDSCR: sculptDSCR,
        totalDividend: rows.reduce(function (a, r) { return a + r.dividend; }, 0),
        projectIRR: irr(projFlows),
        projectIRRPre: irr(preTaxFlows),
        npv: npv(i.discount, projFlows),
        npvEquity: npv(i.discount, eqFlows),
        minDSCR: minDSCR, avgDSCR: avgDSCR,
        payback: payback, lcoe: lcoe,
        avgEbitda: totalEbitda / N,
        ebitdaMargin: totalRev ? totalEbitda / totalRev : 0,
        totalRevenue: totalRev
      }
    };
  }

  /* 민감도: 하나의 변수만 변화 */
  function sensitivity(base, key, values) {
    return values.map(function (v) {
      var inp = {}; for (var k in base) inp[k] = base[k];
      inp[key] = v;
      var r = computeModel(inp);
      return {
        value: v, equityIRR: r.kpi.equityIRR, projectIRR: r.kpi.projectIRR,
        minDSCR: r.kpi.minDSCR, npv: r.kpi.npv, avgEbitda: r.kpi.avgEbitda
      };
    });
  }

  global.SolarModel = {
    computeModel: computeModel, sensitivity: sensitivity,
    irr: irr, npv: npv, pmt: pmt, corpTax: corpTax, EOK: EOK, addMonths: addMonths
  };
})(typeof window !== 'undefined' ? window : globalThis);
