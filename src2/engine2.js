/* =========================================================================
   태양광 PF 재무모델 엔진 v2
   - 기간 단위: 월(12) / 분기(4) / 연(1) 선택. 원본 모델과 동일하게 분기 기본.
   - 자금조달: 자본금 + 선순위A/B/C/D + 후순위 (5개 트랜치), 투입순서 기반 인출
   - 건설이자: 기말인출 / 기초잔액 기준 / 이자부잔액에 자본화하지 않음 (원본 규약)
   ========================================================================= */
(function (root) {
  'use strict';

  /* ---------- 날짜 유틸 ---------- */
  function ym(s) { var p = String(s).split('-'); return { y: +p[0], m: +p[1] }; }
  function addM(o, n) {
    var t = o.y * 12 + (o.m - 1) + n;
    return { y: Math.floor(t / 12), m: (t % 12) + 1 };
  }
  function eom(y, m) { return new Date(Date.UTC(y, m, 0)); }
  function key(o) { return o.y * 12 + o.m; }          // 월 일련번호
  function iso(d) { return d.toISOString().slice(0, 10); }

  /* ---------- 금융 유틸 ---------- */
  function npv(rate, flows) {
    var v = 0; for (var i = 0; i < flows.length; i++) v += flows[i] / Math.pow(1 + rate, i);
    return v;
  }
  function irr(flows, guess) {
    var lo = -0.9999, hi = 10, f = function (r) { return npv(r, flows); };
    if (f(lo) * f(hi) > 0) return NaN;
    for (var i = 0; i < 200; i++) {
      var mid = (lo + hi) / 2;
      if (f(lo) * f(mid) <= 0) hi = mid; else lo = mid;
    }
    return (lo + hi) / 2;
  }
  function annualize(rPeriod, ppy) { return Math.pow(1 + rPeriod, ppy) - 1; }

  /* ---------- 법인세 (누진) ---------- */
  function corpTax(base) {
    if (base <= 0) return 0;
    var b = [[200, 0.09], [20000, 0.19], [300000, 0.21], [Infinity, 0.24]];  // 단위 백만원
    var tax = 0, prev = 0;
    for (var i = 0; i < b.length; i++) {
      var cap = b[i][0], rate = b[i][1];
      if (base > prev) { tax += (Math.min(base, cap) - prev) * rate; prev = cap; }
      else break;
    }
    return tax;
  }

  /* =====================================================================
     기간 축 생성
     ===================================================================== */
  function buildPeriods(inp) {
    var ppy = inp.ppy, plen = 12 / ppy;
    var cs = ym(inp.constructionStart);                       // 착공 (해당 월 1일)
    var codM = addM(cs, inp.constructionMonths);              // COD 월
    var endM = addM(codM, Math.round(inp.operationYears * 12));

    // 시작 기간: 착공월이 속한 기간의 첫 월
    var startIdx = Math.floor((cs.m - 1) / plen) * plen;
    var cur = { y: cs.y, m: startIdx + 1 };

    var ps = [], n = 0;
    while (key(cur) < key(endM)) {
      var pEndM = addM(cur, plen - 1);
      var conM = 0, opM = 0;
      for (var k = 0; k < plen; k++) {
        var mm = addM(cur, k), kk = key(mm);
        if (kk >= key(cs) && kk < key(codM)) conM++;
        if (kk >= key(codM) && kk < key(endM)) opM++;
      }
      ps.push({
        n: n++, end: eom(pEndM.y, pEndM.m), endStr: iso(eom(pEndM.y, pEndM.m)),
        year: pEndM.y, conMonths: conM, opMonths: opM,
        isCon: conM > 0, isOp: opM > 0
      });
      cur = addM(cur, plen);
    }
    // 운영연차 (degradation·에스컬레이션용)
    var cum = 0;
    ps.forEach(function (p) { p.opYearIdx = Math.floor(cum / 12); cum += p.opMonths; });
    return ps;
  }

  /* =====================================================================
     건설기간: 인출 + 건설이자 (순환구조 → 반복수렴)
     ===================================================================== */
  function buildConstruction(inp, ps) {
    var ppy = inp.ppy;
    var conPs = ps.filter(function (p) { return p.isCon; });
    var totConM = conPs.reduce(function (a, p) { return a + p.conMonths; }, 0);
    var codIdx = -1;
    for (var i = 0; i < ps.length; i++) if (ps[i].isCon) codIdx = i;   // 마지막 건설기간

    // 공사비 지출 곡선 (기본: 월수 비례)
    var curve = conPs.map(function (p) { return p.conMonths / totConM; });
    if (inp.spendCurve && inp.spendCurve.length === conPs.length) {
      var s = inp.spendCurve.reduce(function (a, b) { return a + b; }, 0);
      curve = inp.spendCurve.map(function (x) { return x / s; });
    }

    var capex = inp.capexEok * 100, dsra = inp.dsraEok * 100;
    var srcs = [{ name: '자본금', amount: inp.equityEok * 100, order: inp.equityOrder || 1, equity: true }]
      .concat(inp.tranches.map(function (t) {
        return { name: t.name, amount: t.amountEok * 100, order: t.order, rateC: t.rateC / 100, ref: t };
      }));

    var idcTot = 0, res = null;
    for (var iter = 0; iter < 60; iter++) {
      // 기간별 소요액: 공사비(곡선) + 건설이자(발생분) + DSRA(준공기)
      var need = ps.map(function () { return 0; });
      conPs.forEach(function (p, i) { need[p.n] += capex * curve[i]; });
      need[codIdx] += dsra;

      // 인출 배분 (투입순서 → 동일순서는 금액 비례)
      var drawn = srcs.map(function () { return 0; });
      var draws = srcs.map(function () { return ps.map(function () { return 0; }); });
      var idc = srcs.map(function () { return ps.map(function () { return 0; }); });
      var openBal = srcs.map(function () { return 0; });

      var orders = srcs.map(function (s) { return s.order; })
        .filter(function (v, i, a) { return a.indexOf(v) === i; }).sort(function (a, b) { return a - b; });

      for (var pi = 0; pi <= codIdx; pi++) {
        // 1) 기초잔액 기준 건설이자 (자본화하지 않음)
        var idcThis = 0;
        srcs.forEach(function (s, si) {
          if (s.equity) return;
          var v = openBal[si] * s.rateC / ppy;
          idc[si][pi] = v; idcThis += v;
        });
        // 2) 소요액 = 공사비 + 당기 건설이자
        var remain = need[pi] + idcThis;
        // 3) 순서대로 인출
        for (var oi = 0; oi < orders.length && remain > 1e-9; oi++) {
          var grp = [];
          srcs.forEach(function (s, si) { if (s.order === orders[oi]) grp.push(si); });
          var room = grp.reduce(function (a, si) { return a + Math.max(0, srcs[si].amount - drawn[si]); }, 0);
          if (room <= 1e-9) continue;
          var take = Math.min(remain, room);
          grp.forEach(function (si) {
            var r = Math.max(0, srcs[si].amount - drawn[si]);
            var d = room > 0 ? take * (r / room) : 0;
            draws[si][pi] += d; drawn[si] += d;
          });
          remain -= take;
        }
        // 4) 기말인출 → 다음기 기초잔액
        srcs.forEach(function (s, si) { openBal[si] += draws[si][pi]; });
      }

      var newIdc = 0;
      idc.forEach(function (a) { a.forEach(function (v) { newIdc += v; }); });
      res = { srcs: srcs, draws: draws, idc: idc, curve: curve, conPs: conPs, codIdx: codIdx, drawn: drawn };
      if (Math.abs(newIdc - idcTot) < 1e-7) { idcTot = newIdc; break; }
      idcTot = newIdc;
      capex = inp.capexEok * 100;   // 공사비는 고정, 건설이자는 need에 별도 가산
    }
    res.idcTotal = idcTot;
    res.idcByTranche = res.srcs.map(function (s, si) {
      return res.idc[si].reduce(function (a, b) { return a + b; }, 0);
    });
    return res;
  }

  /* =====================================================================
     본 계산
     ===================================================================== */
  function computeModel(inp) {
    var ppy = inp.ppy || 4;
    inp = Object.assign({}, inp, { ppy: ppy });
    var ps = buildPeriods(inp);
    var con = buildConstruction(inp, ps);

    var capex = inp.capexEok * 100, dsra0 = inp.dsraEok * 100;
    var idc = con.idcTotal;
    var tic = capex + idc;                       // 총투자비(건설이자 포함)
    var equity = inp.equityEok * 100;

    /* ---- 매출 / 운영비 ---- */
    var ppyF = 12 / ppy;
    ps.forEach(function (p) {
      var frac = p.opMonths / 12;
      var deg = Math.pow(1 - inp.degradation / 100, p.opYearIdx);
      p.gen = inp.capacityMW * 8760 * (inp.capacityFactor / 100) * deg * (1 - inp.auxRate / 100) * frac;
      p.price = inp.tariff * Math.pow(1 + inp.tariffEscal / 100, p.opYearIdx);
      p.revenue = p.gen * p.price / 1000;
      p.opex = p.isOp ? inp.opexEok * 100 * Math.pow(1 + inp.opexEscal / 100, p.opYearIdx) * frac : 0;
      // 지급순위: 선순위운영비(대부분 항목)는 CFADS/DSCR 계산 전에, 후순위운영비(O&M
      // 등)는 원리금 상환 뒤 배당 전에 빠진다 (Opex&Capex!A16:A34 SUMPRODUCT 플래그 구조,
      // 당진 원본 기준 후순위 비중 = 34,685.50 / 105,377.38 ≈ 32.9%). 분리하지 않으면
      // CFADS가 과소평가되어 DSCR이 원본보다 낮게 나온다.
      var subShare = inp.opexSubShare != null ? inp.opexSubShare / 100 : 0;
      p.opexSub = p.opex * subShare;
      p.opexSenior = p.opex - p.opexSub;
      p.ebitda = p.revenue - p.opex;
    });

    /* ---- 감가상각 ---- */
    var depBase = tic * (inp.depRatio / 100), depAnnual = depBase / inp.depYears;
    ps.forEach(function (p) {
      p.dep = (p.isOp && p.opYearIdx < inp.depYears) ? depAnnual * (p.opMonths / 12) : 0;
    });

    /* ---- 트랜치별 상환 스케줄 ----
       상환개시는 트랜치별 명시 날짜(t.repayStart)를 우선 사용한다.
       거치기간 역산은 A(2026-06-30)/C(2026-09-30)처럼 트랜치마다 다른 규칙이
       섞여 원본과 안 맞아서, 명시 날짜가 없을 때만 폴백으로 쓴다. (SPEC.md 5장) */
    var trs = inp.tranches.map(function (t, ti) {
      var si = ti + 1;                                   // srcs[0] = 자본금
      var amt = con.drawn[si];
      var firstDrawIdx = -1;
      con.draws[si].forEach(function (v, i) { if (v > 1e-9 && firstDrawIdx < 0) firstDrawIdx = i; });
      var repayStartIdx = -1;
      if (t.repayStart) {
        for (var i = 0; i < ps.length; i++) { if (ps[i].endStr === t.repayStart) { repayStartIdx = i; break; } }
      }
      var schedule = t.schedule || null;
      var nRepay = schedule ? schedule.length : Math.round(t.repayYears * ppy);
      return {
        name: t.name, amount: amt, rateO: t.rateO / 100, method: t.method,
        graceYears: t.graceYears, repayYears: t.repayYears,
        firstDrawIdx: firstDrawIdx, schedule: schedule,
        repayStartIdx: repayStartIdx, nRepay: nRepay,
        draws: con.draws[si], idcSeries: con.idc[si]
      };
    });

    trs.forEach(function (t) {
      if (t.amount <= 1e-9 || t.firstDrawIdx < 0) { t.nRepay = 0; return; }
      if (t.repayStartIdx < 0) {
        var fd = ps[t.firstDrawIdx].end;
        var target = new Date(Date.UTC(fd.getUTCFullYear(), fd.getUTCMonth() + Math.round(t.graceYears * 12) + 1, 0));
        for (var i = 0; i < ps.length; i++) {
          if (ps[i].end.getTime() > target.getTime() ||
            (ps[i].end.getTime() === target.getTime() && t.graceYears % 1 !== 0)) { t.repayStartIdx = i; break; }
        }
        if (t.repayStartIdx < 0) t.repayStartIdx = ps.length - 1;
      }
      t.repayEndIdx = Math.min(ps.length - 1, t.repayStartIdx + t.nRepay - 1);
    });

    /* ---- 상환 + 손익/현금흐름 ----
       당진은 세 트랜치 모두 방식 3(64회차 상환비율 직접 키인)이라 원금이
       CFADS와 무관하게 정해진다. 즉 스컬프팅 순환참조가 없으므로 단일
       패스로 충분하다 (SPEC.md 10장 함정 5 참조 — 스컬프팅 자체는 여전히
       미지원이며, 스케줄 없는 방식3 입력 시 원금이 0이 된다). */
    function runOps() {
      var bal = trs.map(function () { return 0; });
      var rows = ps.map(function () { return {}; });
      var yrEBT = {};

      for (var i = 0; i < ps.length; i++) {
        var p = ps[i], r = rows[i];
        // 인출 반영
        trs.forEach(function (t, ti) { bal[ti] += t.draws[i]; });

        // 이자 (기초잔액)
        var openTot = 0, interest = 0;
        trs.forEach(function (t, ti) {
          var ob = bal[ti] - t.draws[i];
          openTot += ob;
          if (p.isOp) interest += ob * t.rateO / ppy;
        });
        r.debtOpen = openTot;
        r.interest = interest;
        r.idc = trs.reduce(function (a, t) { return a + t.idcSeries[i]; }, 0);

        // 원금상환
        var prin = 0, prinBy = [];
        trs.forEach(function (t, ti) {
          var v = 0;
          if (t.nRepay > 0 && i >= t.repayStartIdx && i <= t.repayEndIdx && bal[ti] > 1e-9) {
            var k = i - t.repayStartIdx;
            if (t.method === 1) v = t.amount / t.nRepay;
            else if (t.method === 2) {
              var rr = t.rateO / ppy;
              var pay = rr > 0 ? t.amount * rr / (1 - Math.pow(1 + rr, -t.nRepay)) : t.amount / t.nRepay;
              v = pay - bal[ti] * rr;
            } else if (t.method === 3 && t.schedule) {
              v = t.amount * (t.schedule[k] || 0);
            }
            v = Math.min(v, bal[ti]);
          }
          prinBy.push(v); prin += v;
        });
        r.principalBy = prinBy;

        // 손익
        r.revenue = p.revenue; r.opex = p.opex; r.ebitda = p.ebitda; r.dep = p.dep;
        r.ebit = p.ebitda - p.dep;
        r.ebt = r.ebit - interest;
        yrEBT[p.year] = (yrEBT[p.year] || 0) + r.ebt;

        r.principal = prin;
        trs.forEach(function (t, ti) { bal[ti] -= prinBy[ti]; });
        r.debtClose = bal.reduce(function (a, b) { return a + b; }, 0);
        r.ds = prin + interest;
      }

      // 연도별 법인세 산출 (이월결손금 반영)
      var taxByYear = {};
      var yrs = Object.keys(yrEBT).map(Number).sort(function (a, b) { return a - b; });
      var carry = 0;
      yrs.forEach(function (y) {
        var ebt = yrEBT[y];
        if (ebt <= 0) { carry += -ebt; taxByYear[y] = 0; return; }
        var ded = Math.min(carry, ebt * (inp.lossRate / 100));
        carry -= ded;
        var base = ebt - ded;
        taxByYear[y] = inp.taxMode === 1 ? corpTax(base) : base * inp.taxFlat / 100;
      });

      // 세금을 기간에 배분 (연도내 양의 EBT 비례, 음수 EBT 분기는 0)
      var yrPos = {};
      ps.forEach(function (p, i) { if (rows[i].ebt > 0) yrPos[p.year] = (yrPos[p.year] || 0) + rows[i].ebt; });
      ps.forEach(function (p, i) {
        var r = rows[i];
        r.tax = (yrPos[p.year] > 0 && r.ebt > 0) ? (taxByYear[p.year] || 0) * (r.ebt / yrPos[p.year]) : 0;
        r.ni = r.ebt - r.tax;
      });

      // 현금흐름 / DSRA / 배당 / 누적DSCR
      var decomTot = inp.decomEok * 100;
      var cashBal = 0, dsraBal = 0;
      var lastOp = -1;
      ps.forEach(function (p, i) { if (p.isOp) lastOp = i; });
      var divCumDSCR = inp.divCumDSCR != null ? inp.divCumDSCR : 0;

      var cumCfads = 0, cumDs = 0;
      ps.forEach(function (p, i) {
        var r = rows[i];
        r.decom = (i === lastOp) ? decomTot : 0;
        // CFADS(원리금 상환 재원)는 선순위운영비까지만 차감한다. 후순위운영비는
        // 원리금 상환 뒤에 빠지므로 DSCR 분자에서 제외 — Report!C210/C217 순서 그대로.
        r.cfads = p.revenue - p.opexSenior - r.tax;
        r.dscr = r.ds > 1e-9 ? r.cfads / r.ds : null;

        // 누적 DSCR: 운영개시부터 누적 CFADS / 누적 원리금
        if (p.isOp) { cumCfads += r.cfads; cumDs += r.ds; }
        r.cumDscr = (p.isOp && cumDs > 1e-9) ? cumCfads / cumDs : null;

        // DSRA: 차기 X개월분 원리금
        var need = 0;
        if (p.isOp && i + 1 < ps.length) {
          var nq = Math.round(inp.dsraMonths / (12 / ppy));
          for (var k = 1; k <= nq && i + k < ps.length; k++) need += rows[i + k].ds;
        }
        if (i === con.codIdx) dsraBal = dsra0;
        r.dsraOpen = dsraBal;
        r.dsraClose = p.isOp ? need : dsraBal;
        r.dsraMove = r.dsraClose - r.dsraOpen;
        dsraBal = r.dsraClose;

        // 후순위운영비 + 철거비(청산기)는 원리금 상환 뒤, 배당 전에 빠진다.
        r.fcfe = r.cfads - r.ds - r.dsraMove - p.opexSub - r.decom;
        r.cashOpen = cashBal;
        var avail = cashBal + r.fcfe;

        // 배당제한: 단순 DSCR + 누적 DSCR 동시 충족
        var canDiv = p.isOp && p.opYearIdx >= (inp.divStartYear - 1) &&
          (r.dscr === null || r.dscr >= inp.divDSCR) &&
          (r.cumDscr === null || r.cumDscr >= divCumDSCR);
        var div = 0;
        if (canDiv) div = Math.max(0, avail - inp.minCash * 100);
        r.dividend = div;
        cashBal = avail - div;
        r.cashClose = cashBal;
        r.projectFcf = p.ebitda - r.tax - r.decom;
      });

      return {
        rows: rows,
        endBalance: rows[rows.length - 1].debtClose
      };
    }

    var out = runOps();
    var rows = out.rows;

    /* ---- 지표 ---- */
    var projFlows = [], eqFlows = [], divFlows = [], preFlows = [];
    // 건설기간 유출
    ps.forEach(function (p, i) {
      var capOut = 0;
      con.conPs.forEach(function (cp, ci) { if (cp.n === i) capOut = capex * con.curve[ci]; });
      projFlows.push(rows[i].projectFcf - capOut);
      preFlows.push(ps[i].ebitda - rows[i].decom - capOut);
      eqFlows.push(rows[i].fcfe - con.draws[0][i]);
      divFlows.push(rows[i].dividend - con.draws[0][i]);
    });

    var pIRRp = irr(projFlows), eIRRp = irr(eqFlows), dIRRp = irr(divFlows), preIRRp = irr(preFlows);
    var dscrs = rows.filter(function (r) { return r.dscr !== null && r.ds > 1e-9; }).map(function (r) { return r.dscr; });
    var cumDscrs = rows.filter(function (r) { return r.cumDscr !== null; }).map(function (r) { return r.cumDscr; });
    // 연도 합산 DSCR (원본 Report 시트가 분기가 아닌 연 단위로 표시하는 것으로 보임 — 검증 중)
    var yrDS = {}, yrCF = {};
    rows.forEach(function (r, i) { if (r.ds > 1e-9) { var y = ps[i].year; yrDS[y] = (yrDS[y] || 0) + r.ds; yrCF[y] = (yrCF[y] || 0) + r.cfads; } });
    var annualDscrs = Object.keys(yrDS).map(function (y) { return yrCF[y] / yrDS[y]; });

    var pvOpex = 0, pvGen = 0, dr = Math.pow(1 + inp.discount / 100, 1 / ppy) - 1;
    ps.forEach(function (p, i) {
      var df = Math.pow(1 + dr, i);
      pvOpex += (p.opex + (i <= con.codIdx ? capex * (con.curve[con.conPs.map(function (c) { return c.n; }).indexOf(i)] || 0) : 0)) / df;
      pvGen += p.gen / df;
    });

    return {
      inp: inp, periods: ps, rows: rows, tranches: trs, con: con,
      idc: idc, tic: tic, equity: equity,
      debt: trs.reduce(function (a, t) { return a + t.amount; }, 0),
      kpi: {
        projectIRR: annualize(pIRRp, ppy),
        projectIRRPre: annualize(preIRRp, ppy),
        equityIRR: annualize(eIRRp, ppy),
        dividendIRR: annualize(dIRRp, ppy),
        npv: npv(dr, projFlows),
        minDSCR: dscrs.length ? Math.min.apply(null, dscrs) : null,
        avgDSCR: dscrs.length ? dscrs.reduce(function (a, b) { return a + b; }, 0) / dscrs.length : null,
        minCumDSCR: cumDscrs.length ? Math.min.apply(null, cumDscrs) : null,
        minDSCRAnnual: annualDscrs.length ? Math.min.apply(null, annualDscrs) : null,
        totalDividend: rows.reduce(function (a, r) { return a + r.dividend; }, 0),
        totalRevenue: rows.reduce(function (a, r) { return a + r.revenue; }, 0),
        totalOpex: rows.reduce(function (a, r) { return a + r.opex; }, 0),
        totalInterest: rows.reduce(function (a, r) { return a + r.interest; }, 0),
        totalTax: rows.reduce(function (a, r) { return a + r.tax; }, 0),
        lcoe: pvGen > 0 ? pvOpex / pvGen * 1000 : 0
      }
    };
  }

  var API = { computeModel: computeModel, irr: irr, npv: npv, corpTax: corpTax, buildPeriods: buildPeriods };
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  root.SolarModel2 = API;
})(typeof globalThis !== 'undefined' ? globalThis : this);
