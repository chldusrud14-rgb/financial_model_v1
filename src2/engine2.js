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
    var f = function (r) { return npv(r, flows); };
    // 분기 현금흐름은 부호가 자주 바뀌어(배당은 일부 분기에만 지급 등) NPV(r)에
    // 실근이 여러 개 있을 수 있다. 0%에서 좌우로 촘촘히 스캔해 "0%에 가장
    // 가까운"(경제적으로 타당한, 엑셀 IRR의 기본 guess=10%와 같은 관례) 부호
    // 전환 구간을 찾고 그 구간만 이분법으로 정밀화한다. 양쪽 다 못 찾으면 NaN.
    var step = 0.002, maxR = 10, f0 = f(0);
    if (f0 === 0) return 0;
    var bracket = null;
    for (var k = 1; k * step <= maxR; k++) {
      var rUp = k * step, rDn = -k * step;
      if (rDn > -0.9999) {
        var vDn = f(rDn), vPrevDn = f(rDn + step);
        if ((vPrevDn > 0) !== (vDn > 0)) { bracket = [rDn, rDn + step]; break; }
      }
      var vUp = f(rUp), vPrevUp = f(rUp - step);
      if ((vPrevUp > 0) !== (vUp > 0)) { bracket = [rUp - step, rUp]; break; }
    }
    if (!bracket) return NaN;
    var a = bracket[0], b = bracket[1];
    for (var i = 0; i < 100; i++) {
      var mid = (a + b) / 2;
      if ((f(a) > 0) !== (f(mid) > 0)) b = mid; else a = mid;
    }
    return (a + b) / 2;
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

    // 분기별 실측 오버라이드(inp.periodOverrides, endStr 키) — 있으면 공식
    // 근사 대신 원본에서 직접 뽑은 값을 그대로 쓴다. 계절성/에스컬레이션
    // 근사가 연간 합계는 맞아도 "균등 분배" 가정 때문에 분기별로는 원본과
    // 어긋나는 항목(부지임대료 연1회 납부 등)이 있어 도입 — 100% 일치 목적.
    var ovrByEnd = {};
    (inp.periodOverrides || []).forEach(function (o) { ovrByEnd[o.end] = o; });

    /* ---- 매출 / 운영비 ---- */
    var ppyF = 12 / ppy;
    ps.forEach(function (p) {
      var ovr = ovrByEnd[p.endStr];
      var frac = p.opMonths / 12;
      // 계절성: 원본은 연간 발전량을 균등 분기가 아니라 계절 비중(예: 당진 2Q
      // 30.57% / 1Q 23.97% / 4Q 23.47% / 3Q 21.99%, 합 100%)으로 배분한다.
      // COD/만기가 분기말과 정확히 맞아떨어져 부분분기가 없는 한(당진 기준)
      // 풀분기에만 적용 — 없으면 기존 균등 배분으로 폴백.
      var full = ppy === 4 && p.opMonths === ppyF;
      var genFrac = (inp.seasonality && full) ? inp.seasonality[p.end.getUTCMonth() + 1] : frac;
      // 효율감소는 복리(Math.pow)가 아니라 선형입니다. 원본: Revenue!row34
      // `=(1-$C34*(운영연차-1))*(운영중)` — 매년 정액으로 감소, 복리 아님.
      // 복리로 계산하면 후반 연차일수록(20년차 기준 0.4%p 이상) 발전량이
      // 원본보다 과대평가된다.
      var deg = Math.max(0, 1 - (inp.degradation / 100) * p.opYearIdx);
      // 연간 발전량 기준: dailyHours(일조시간)가 있으면 원본과 동일하게
      // capacityMW × dailyHours × 365로 계산(더 정확). 없으면 capacityFactor
      // 기반(capacityMW × 8760 × cf)으로 폴백 — 8760h 환산 반올림 때문에
      // capacityFactor를 dailyHours에서 역산해 넣으면 0.01%대 오차가 남는다.
      var annualGen = inp.dailyHours != null
        ? inp.capacityMW * inp.dailyHours * 365
        : inp.capacityMW * 8760 * (inp.capacityFactor / 100);
      p.gen = annualGen * deg * (1 - inp.auxRate / 100) * genFrac;
      // 매출단가: 트랙별 입력(inp.tariffTracks, 예 PPA1/PPA2/SMP+REC 비중·단가)이
      // 있으면 가중평균 대신 트랙별로 정확히 계산. 없으면 기존 단일 tariff 폴백.
      if (ovr) {
        p.revenue = ovr.revenue;
        p.price = p.gen > 0 ? p.revenue / p.gen * 1000 : 0;
      } else if (inp.tariffTracks && inp.tariffTracks.length) {
        var rev = 0;
        inp.tariffTracks.forEach(function (tr) {
          var price = tr.price * Math.pow(1 + (tr.escal || 0) / 100, p.opYearIdx);
          rev += p.gen * tr.share * price / 1000;
        });
        p.revenue = rev;
        p.price = p.gen > 0 ? rev / p.gen * 1000 : 0;
      } else {
        p.price = inp.tariff * Math.pow(1 + inp.tariffEscal / 100, p.opYearIdx);
        p.revenue = p.gen * p.price / 1000;
      }

      // 운영비: 분기 실측 오버라이드 > 항목별 입력(inp.opexItems) > 총액 근사
      // (opexEok/opexEscal/opexSubShare) 순으로 폴백한다.
      // 지급순위: 선순위운영비는 CFADS/DSCR 계산 전에, 후순위운영비(O&M 등)는
      // 원리금 상환 뒤 배당 전에 빠진다 (Opex&Capex!A16:A34 SUMPRODUCT 플래그 구조).
      // 분리하지 않으면 CFADS가 과소평가되어 DSCR이 원본보다 낮게 나온다.
      if (ovr) {
        p.opexSenior = ovr.opexSenior; p.opexSub = ovr.opexSub; p.opex = ovr.opexSenior + ovr.opexSub;
      } else if (p.isOp && inp.opexItems && inp.opexItems.length) {
        var senior = 0, sub = 0;
        inp.opexItems.forEach(function (it) {
          var esc = it.escal ? Math.pow(1 + it.escal / 100, p.opYearIdx) : 1;
          var v = it.annualKRWm * frac * esc;
          if (it.senior === false) sub += v; else senior += v;
        });
        p.opexSenior = senior; p.opexSub = sub; p.opex = senior + sub;
      } else if (p.isOp) {
        p.opex = inp.opexEok * 100 * Math.pow(1 + inp.opexEscal / 100, p.opYearIdx) * frac;
        var subShare = inp.opexSubShare != null ? inp.opexSubShare / 100 : 0;
        p.opexSub = p.opex * subShare;
        p.opexSenior = p.opex - p.opexSub;
      } else {
        p.opex = 0; p.opexSenior = 0; p.opexSub = 0;
      }
      p.ebitda = p.revenue - p.opex;
    });

    /* ---- 감가상각 ----
       상각대상은 TIC 전액이 아니라 "건설중인자산 계상액"(TIC!row34) — 토지
       선납임대료 등 상각 제외 항목을 뺀 값이다. inp.depBaseOverride가 있으면
       그 값을 그대로 쓰고, 없으면 tic*depRatio%로 근사한다. */
    var depBase = inp.depBaseOverride != null ? inp.depBaseOverride : tic * (inp.depRatio / 100);
    var depAnnual = depBase / inp.depYears;
    ps.forEach(function (p) {
      p.dep = (p.isOp && p.opYearIdx < inp.depYears) ? depAnnual * (p.opMonths / 12) : 0;
    });

    /* ---- 복구충당부채(철거비) 전입액 ----
       원본은 철거비를 만기 시점에 한 번에 비용처리하지 않고, 운영기간 전체에
       걸쳐 균등하게 충당부채로 전입한다(IS(Q)!row36, BS(Q)!row131 — 당진은
       분기 25 = 2,000/80분기 정액). 세무상으로만 손금(EBT 차감) 처리되는
       비현금 항목이고, 실제 현금 지출은 기존처럼 마지막 운영분기에 한 번에
       일어난다(아래 r.decom). 이걸 빼먹으면 EBT·법인세가 원본보다 과대평가된다. */
    var opPeriodCount = ps.filter(function (p) { return p.isOp; }).length;
    var decomAccrualQ = opPeriodCount > 0 ? (inp.decomEok * 100) / opPeriodCount : 0;

    /* ---- 기타 비현금 세무손금 (예: 선납임대료 상각액) ----
       Opex&Capex의 부지임대료(분납, 현금)와 별개로 IS(Y)의 부지임대료 라인은
       선납임대료 상각액까지 포함한다(당진: 25,853.65 vs 현금기준 19,110,
       차액 6,743.65). CFADS/배당엔 영향 없는 비현금 손금이라 opex가 아니라
       여기서 EBT만 줄인다. inp.extraTaxDeductionKRWm으로 총액을 받아 운영
       분기에 균등 배분. */
    var extraTaxDedQ = opPeriodCount > 0 ? (inp.extraTaxDeductionKRWm || 0) / opPeriodCount : 0;

    /* ---- 대리은행수수료 ----
       Report!row213 / IS(Q)!row53에서 확인 — 운영기간 내내 연 20(=분기 5) 발생하는
       실제 현금성 채무관리비용. EBT와 CFADS(원리금 상환 재원) 양쪽에서 차감된다. */
    var agentFeeQ = opPeriodCount > 0 ? (inp.agentFeeKRWm || 0) / opPeriodCount : 0;

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
      var decomTot = inp.decomEok * 100;
      var lastOp = -1;
      ps.forEach(function (p, i) { if (p.isOp) lastOp = i; });

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
        var ovrQ = ovrByEnd[p.endStr];
        r.decomAccrual = ovrQ ? ovrQ.decomAccrual : (p.isOp ? decomAccrualQ : 0);
        r.extraTaxDed = ovrQ ? ovrQ.extraTaxDed : (p.isOp ? extraTaxDedQ : 0);
        r.agentFee = ovrQ ? ovrQ.agentFee : (p.isOp ? agentFeeQ : 0);
        r.ebit = p.ebitda - p.dep - r.decomAccrual - r.extraTaxDed - r.agentFee;
        r.ebt = r.ebit - interest;
        // 세무상 손금 인식 시점은 회계상(발생주의)과 다르다 — 복구충당부채
        // 전입액(회계상 비용)은 세법상 손금불산입(더해서 되돌림)이고, 실제
        // 철거비를 현금 지급하는 시점(만기)에만 손금산입된다(IS(Y)!row64/65).
        // 회계상 EBT/NI(r.ebt)는 그대로 두고, 법인세 계산용 과세소득만 조정.
        var cashDecomThis = i === lastOp ? decomTot : 0;
        r.taxableEbt = r.ebt + r.decomAccrual - cashDecomThis;
        yrEBT[p.year] = (yrEBT[p.year] || 0) + r.taxableEbt;

        r.principal = prin;
        trs.forEach(function (t, ti) { bal[ti] -= prinBy[ti]; });
        r.debtClose = bal.reduce(function (a, b) { return a + b; }, 0);
        r.ds = prin + interest;
      }

      // 연도별 법인세 산출 (이월결손금 + 통합투자세액공제 + 최저한세 + 지방소득세/농특세 반영)
      // 원본(IS(Y)!row90~125) 구조: 산출세액(누진 브래킷) → 세액공제(당해+10년 이월,
      // 건설기간 EPC 투자액×공제율) 차감하되 최저한세(과세표준×AMT율) 밑으로는 못 내려감.
      // 법인세비용 = 법인세×(1+지방소득세율) + 실제공제액×농특세율.
      // 건설기간(비용화된 개발비 등) 손실은 하나의 시드로 뭉뚱그리지 않고
      // 실제 발생 연도(2024/2025)별로 정확히 반영한다 — preOpLossKRWm 전체를
      // 2026 진입 전 시드로 넣으면 2025년(첫 운영분기만 이익)이 실제로는
      // 그 해 전체로 보면 적자인데도 내 모델에서는 미세하게 흑자로 잡혀
      // 세금이 잘못 붙는다(2025 taxByYear가 0이어야 하는데 그렇지 않았음).
      var preOpLossByYear = inp.preOpLossByYear || {};
      Object.keys(preOpLossByYear).forEach(function (y) {
        yrEBT[y] = (yrEBT[y] || 0) - preOpLossByYear[y];
      });
      var taxByYear = {};
      var yrs = Object.keys(yrEBT).map(Number).sort(function (a, b) { return a - b; });
      var carry = 0;
      var creditPool = 0;
      var creditRate = (inp.investmentCreditRate || 0) / 100;
      var amtRate = (inp.amtRate || 0) / 100;
      var surtaxRate = (inp.localSurtaxRate || 0) / 100;
      var creditSurtaxRate = (inp.creditSurtaxRate || 0) / 100;
      var creditBase = inp.investmentCreditBaseByYear || {};
      yrs.forEach(function (y) {
        creditPool += (creditBase[y] || 0) * creditRate;
        var ebt = yrEBT[y];
        if (ebt <= 0) { carry += -ebt; taxByYear[y] = 0; return; }
        var ded = Math.min(carry, ebt * (inp.lossRate / 100));
        carry -= ded;
        var base = ebt - ded;
        var grossTax = inp.taxMode === 1 ? corpTax(base) : base * inp.taxFlat / 100;
        var amtFloor = base * amtRate;
        var afterCredit = Math.max(0, grossTax - creditPool);
        var taxFinal = Math.max(amtFloor, afterCredit);
        var creditUsed = grossTax - taxFinal;
        creditPool -= creditUsed;
        taxByYear[y] = taxFinal * (1 + surtaxRate) + creditUsed * creditSurtaxRate;
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
      var cashBal = 0, dsraBal = 0, pendingDiv = 0;
      var divCumDSCR = inp.divCumDSCR != null ? inp.divCumDSCR : 0;
      // 배당가능이익(상법상 배당 한도) = 누적 당기순이익 - 기지급 배당. 원본은
      // 배당가능이익을 넘는 잉여현금을 유상감자로 돌리는데(BS(Q)!row253
      // "유상감자 가능한도(배당가능이익 초과현금)") 당진은 유상감자 스위치가
      // 꺼져 있어(CF(Q)!row14 전 기간 0) 그 초과분이 그냥 쌓였다가 마지막에
      // 청산배당으로 나간다. 이 한도를 안 걸면 부채 상환 완료(2042) 이후
      // 현금이 넉넉해지면서 연차배당이 원본보다 훨씬 커진다.
      var distributable = 0;
      // 이익준비금(상법 제458조): 배당액의 10%를 자본금의 50%까지 강제 적립,
      // 배당가능이익에서 영구히 빠진다(BS(Q)!row242~244). 없으면 부채 만기
      // 이후 배당가능이익이 실제보다 커져서 배당이 과대평가된다.
      var reserveBalance = 0, reserveCap = equity * 0.5;

      // 배당 게이트용 연도별 누적치. "누적DSCR"이 (그해 기초현금+CFADS)/DS이므로
      // 배당 시점(연 1회, 통상 3월)엔 직전 연도가 이미 다 끝나 있어 이 값들이
      // 확정돼 있다 — 매 분기 갱신해두면 3월 분기 도달 시 바로 조회 가능.
      var yrDS2 = {}, yrCF2 = {}, yrCFCash2 = {}, yrReserve2 = {}, yrCashStart = {}, yrPostMarchCash = {};

      var cumCfads = 0, cumDs = 0;
      ps.forEach(function (p, i) {
        var r = rows[i];
        r.decom = (i === lastOp) ? decomTot : 0;
        // CFADS(원리금 상환 재원)는 선순위운영비까지만 차감한다. 후순위운영비는
        // 원리금 상환 뒤에 빠지므로 DSCR 분자에서 제외 — Report!C210/C217 순서 그대로.
        // 운전자본 증감(매출채권 회수 타이밍 등, CF(Q)!row29)도 반영 — 20년
        // 합계는 0이지만 분기별로는 크다(특히 준공 직후 첫 분기).
        var ovrC = ovrByEnd[p.endStr];
        // wc는 CF(Q)!row29 값을 그대로 쓴다 — 이미 부호가 들어있다(현금 유출이면
        // 음수). 따라서 빼는 게 아니라 더한다.
        r.wc = ovrC ? (ovrC.wc || 0) : 0;
        r.cfads = p.revenue - p.opexSenior - r.tax - r.agentFee + r.wc;
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
        // FCFE의 DSRA 항목은 IRR!row83과 대조해보니 "차기 6개월분" 선행 계산
        // 룰이 아니라 원본만의 다른(미확인) 스케줄로 움직인다 — DSRA 잔액
        // 자체(dsraOpen/dsraClose, 배당가능 현금·누적DSCR 계산에 씀)는 총액이
        // 이미 원본과 맞아떨어져서 그대로 두고, FCFE 계산에서만 실측값
        // (periodOverrides의 dsraFcfe, 부호: 양수=환입/유입)으로 대체한다.
        var ovrD = ovrByEnd[p.endStr];
        var dsraForFcfe = ovrD && ovrD.dsraFcfe != null ? -ovrD.dsraFcfe : r.dsraMove;
        // 법인세도 CFADS/DSCR 계산에선 발생주의(r.tax)를 쓰지만, 실제 현금
        // 유출(FCFE·배당 재원)은 원본이 다음 해 3월에 한 번에 몰아 낸다
        // (CF(Q)!row31 — 배당과 같은 타이밍). taxCash가 있으면 발생주의 세금을
        // 되돌리고 실제 납부액으로 바꿔치기.
        var taxAdj = ovrD && ovrD.taxCash != null ? (r.tax - ovrD.taxCash) : 0;
        r.fcfe = r.cfads - r.ds - dsraForFcfe - p.opexSub - r.decom + taxAdj;
        r.cashOpen = cashBal;
        var avail = cashBal + r.fcfe;
        if (p.isOp) distributable += r.ni;

        if (p.isOp && yrReserve2[p.year] === undefined) yrReserve2[p.year] = r.cashOpen + r.dsraOpen;
        if (p.isOp && yrCashStart[p.year] === undefined) yrCashStart[p.year] = r.cashOpen;
        if (r.ds > 1e-9) {
          yrDS2[p.year] = (yrDS2[p.year] || 0) + r.ds; yrCF2[p.year] = (yrCF2[p.year] || 0) + r.cfads;
          // CF(Q)!row33("영업현금흐름") 기준 원리금상환재원은 Report와 달리
          // 세금을 발생주의가 아니라 현금주의(taxAdj로 이미 구한 실제
          // 납부액)로 뺀다 — DSCR용 r.cfads(발생주의)와는 다른 트랙.
          yrCFCash2[p.year] = (yrCFCash2[p.year] || 0) + (r.cfads + taxAdj);
        }

        // 배당: 원본은 **12월에 결의, 이듬해 3월에 지급**한다
        // (BS(Q)!row242~248 — 결의 시점 현금·배당가능이익으로 금액을 정하고,
        // row248이 그 결의액을 다음 분기로 그대로 넘겨 지급). 그래서 "직전
        // 연도의 DSCR로 3월에 게이트"가 아니라 "그 해 12월(그 해 DSCR이
        // 막 확정된 시점)에 그 해 자기 실적으로 게이트"가 맞다. 결의만
        // 하고 지급은 다음 분기(pendingDiv)로 넘긴다.
        // inp.dividendMonth/firstDividendYear 없으면 기존 opYearIdx 기반
        // 근사(즉시 스윕)로 폴백.
        var div = 0;
        if (inp.dividendMonth != null) {
          var decideMonth = inp.dividendMonth === 3 ? 12 : inp.dividendMonth;
          var isPayQ = (p.end.getUTCMonth() + 1) === inp.dividendMonth;
          if (isPayQ) { div = pendingDiv; pendingDiv = 0; }

          if (p.isOp && (p.end.getUTCMonth() + 1) === decideMonth) {
            var pSimple = yrDS2[p.year] ? yrCF2[p.year] / yrDS2[p.year] : null;
            var pCum = yrDS2[p.year] ? (yrReserve2[p.year] + yrCF2[p.year]) / yrDS2[p.year] : null;
            var canDecide = (inp.firstDividendYear == null || p.year + 1 >= inp.firstDividendYear) &&
              (pSimple === null || pSimple >= inp.divDSCR) &&
              (pCum === null || pCum >= divCumDSCR);
            if (canDecide) {
              // 이익준비금은 "배당가능이익(적립전)의 10%/1.1"이 아니라
              // "실제로 지급하기로 한 배당금의 10%"다(BS(Q)!row208
              // `=MIN(S247*10%,...)`, S247이 배당금 그 자체). row243
              // (10%/1.1)은 배당가능이익 한도를 역산하는 상한 계산용이라
              // 별개 — 현금이 부족해 배당이 이익한도보다 적게 나가면
              // 준비금도 그 실제 결의액의 10%만 쌓인다.
              var maxByProfit = distributable / 1.1;
              var maxByCash = avail - inp.minCash * 100;
              // BS(Q)!row239 "배당 후 누적DSCR 고려한 배당가능현금" — 배당
              // 가능 여부의 이분법적 게이트(pCum>=divCumDSCR)와 별개로, 매년
              // "그 해 원리금상환재원(기초현금+영업현금흐름) − 그 해
              // 원리금상환액×목표누적DSCR"만큼은 항상 현금으로 남겨야 한다.
              // 부채를 다 갚은 해도 그 해에 갚은 원금·이자가 조금이라도
              // 있으면 이 캡이 걸린다 — 완전상환 다음 해부터는 원리금상환액이
              // 0이라 자동으로 풀린다. 이게 없으면 부채 만기 직후 배당이
              // 급증한다.
              // "원리금상환재원"(CF(Y)!row101=S92+S14) = 그 해 CFADS(S92,
              // CF(Q)!row6="대출기간" 태그 기준 SUMIF — 그 분기에 원리금상환이
              // 있는 분기만 그 해로 잡힌다. 부채를 다 갚은 분기부터는 태그가
              // 0이 돼서 이후로는 영원히 0) + S14(그 해 3월 지급 직후 현금,
              // = 기초현금 − 그 해 자체 3월 배당). 부채를 다 갚은 해는 S92가
              // 0으로 얼어붙어서 "그 해 자체의 새 영업이익"은 이 캡에 전혀
              // 더해지지 않는다 — 원본 수식을 그대로 재현(의도된 기능인지
              // 부수효과인지는 불명확하지만 값은 정확히 이렇게 나온다).
              var yearlyAvail = (yrPostMarchCash[p.year] != null ? yrPostMarchCash[p.year] : (yrCashStart[p.year] || 0)) + (yrCFCash2[p.year] || 0);
              var yearlyDS = yrDS2[p.year] || 0;
              var maxByDscrReserve = Math.max(0, yearlyAvail - yearlyDS * divCumDSCR);
              var decided = Math.max(0, Math.min(Math.min(maxByCash, maxByProfit), maxByDscrReserve));
              var reserveNeed = Math.max(0, Math.min(decided * 0.10, reserveCap - reserveBalance));
              reserveBalance += reserveNeed;
              distributable -= (decided + reserveNeed);
              pendingDiv = decided;
            }
          }
        } else {
          var canDiv = p.isOp && p.opYearIdx >= (inp.divStartYear - 1) &&
            (r.dscr === null || r.dscr >= inp.divDSCR) &&
            (r.cumDscr === null || r.cumDscr >= divCumDSCR);
          if (canDiv) div = Math.max(0, avail - inp.minCash * 100);
        }
        // 청산배당: 운영 마지막 분기는 최소보유현금도, 배당가능이익 한도도
        // 안 보고 (결의 대기 중인 pendingDiv 포함) 잔여현금 전액을 배당한다
        // (CF(Q)!row77 = SUMIF(운영종료분기, 기말현금) — 연차배당과 별개의
        // 마지막 정산이라 상법상 이익배당 한도가 아니라 청산/자본반환 성격).
        if (i === lastOp) { div = Math.max(0, avail); pendingDiv = 0; }
        r.dividend = div;
        cashBal = avail - div;
        r.cashClose = cashBal;
        // CF(Y)!row14 "배당&유상감자 후 잔액" = 그 해 기초현금(3월 분기
        // 자체의 영업현금흐름은 안 더함) − 그 해 3월 배당. 부채 완제 후
        // "원리금상환재원" 캡의 기준선(아래 yearlyAvail)이 된다.
        if (inp.dividendMonth != null && p.isOp && (p.end.getUTCMonth() + 1) === inp.dividendMonth) {
          yrPostMarchCash[p.year] = r.cashOpen - div;
        }
        // Project FCF: 건설이자는 유출에서 빼고(아래 capOut이 TIC_exIDC 기준),
        // 유입에서는 운전자본 증감·철거비·대리은행수수료(금융수수료)를
        // 차감한다 — `IRR!row12~27`로 직접 대조해서 확인한 정의. 원본이 두
        // 오차(IDC 포함 + 수수료 누락)를 서로 반대 방향으로 갖고 있어서
        // 우연히 비슷해 보였던 v1 함정과 같은 문제.
        r.projectFcf = p.ebitda - r.tax - r.decom - r.agentFee + r.wc;
      });

      return {
        rows: rows,
        endBalance: rows[rows.length - 1].debtClose
      };
    }

    var out = runOps();
    var rows = out.rows;

    /* ---- 지표 ---- */
    var projFlows = [], eqFlows = [], divFlows = [], preFlows = [], investorFlows = [];
    // 건설기간 유출
    ps.forEach(function (p, i) {
      var capOut = 0;
      con.conPs.forEach(function (cp, ci) { if (cp.n === i) capOut = capex * con.curve[ci]; });
      projFlows.push(rows[i].projectFcf - capOut);
      preFlows.push(ps[i].ebitda - rows[i].decom - rows[i].agentFee + rows[i].wc - capOut);
      eqFlows.push(rows[i].fcfe - con.draws[0][i]);
      divFlows.push(rows[i].dividend - con.draws[0][i]);

      // Investor IRR: 자본+부채 조달 전체(유출) vs 건설이자·원리금·배당
      // 전체(유입) — `IRR!row33~70`으로 직접 대조해서 확인한 정의. 건설이자
      // (IDC)는 실제 현금흐름이 아니라 준공 시점에 한 번에 정산되는 것으로
      // 처리한다(원본도 분기별로 안 풀고 총액 한 방으로 처리 — Funding!E55
      // 건설이자(값) 하드코딩과 같은 방식).
      var debtDraw = 0;
      con.srcs.forEach(function (s, si) { if (!s.equity) debtDraw += con.draws[si][i]; });
      investorFlows.push(
        -con.draws[0][i] - debtDraw +
        rows[i].interest + rows[i].principal + rows[i].dividend +
        (i === con.codIdx ? con.idcTotal : 0)
      );
    });

    var pIRRp = irr(projFlows), eIRRp = irr(eqFlows), dIRRp = irr(divFlows), preIRRp = irr(preFlows);
    var invIRRp = irr(investorFlows);
    var dscrs = rows.filter(function (r) { return r.dscr !== null && r.ds > 1e-9; }).map(function (r) { return r.dscr; });
    // DSCR은 원본 Report 시트 기준 연 단위 지표다(분기 롤링이 아님).
    // 단순DSCR = 그 해 CFADS합/DS합.
    // 누적DSCR = Report!row233 "(A+B)/C" — A=그 해 기초현금(DSRA 포함), B=그 해
    // 영업현금흐름(CFADS), C=그 해 원리금상환. "COD부터 누적한 CFADS/DS"가
    // 아니라 "그동안 쌓인 현금 여유분 + 이번 해 벌이"가 "이번 해 원리금"을
    // 몇 배 덮는지를 본다 — 배당으로 현금을 다 빼가지 않는 한 매년 우상향하는
    // 이유가 이것. row224(기초현금)로 검증: 2025 (0+2707.60)/1868.75=1.4489,
    // 2026 (5463.85+17798.98)/14407.14=1.6146 — 실제로 정확히 일치했다.
    var yrDS = {}, yrCF = {};
    rows.forEach(function (r, i) { if (r.ds > 1e-9) { var y = ps[i].year; yrDS[y] = (yrDS[y] || 0) + r.ds; yrCF[y] = (yrCF[y] || 0) + r.cfads; } });
    var yrsSorted = Object.keys(yrDS).map(Number).sort(function (a, b) { return a - b; });
    var annualDscrs = yrsSorted.map(function (y) { return yrCF[y] / yrDS[y]; });
    var yrFirstIdx = {};
    ps.forEach(function (p, i) { if (yrFirstIdx[p.year] === undefined) yrFirstIdx[p.year] = i; });
    var annualCumDscrs = yrsSorted.map(function (y) {
      var r0 = rows[yrFirstIdx[y]];
      var reserve = r0.cashOpen + r0.dsraOpen;
      return (reserve + yrCF[y]) / yrDS[y];
    });

    var pvOpex = 0, pvGen = 0, dr = Math.pow(1 + inp.discount / 100, 1 / ppy) - 1;
    ps.forEach(function (p, i) {
      var df = Math.pow(1 + dr, i);
      pvOpex += (p.opex + (i <= con.codIdx ? capex * (con.curve[con.conPs.map(function (c) { return c.n; }).indexOf(i)] || 0) : 0)) / df;
      pvGen += p.gen / df;
    });

    var totalDividendKRWm = rows.reduce(function (a, r) { return a + r.dividend; }, 0);

    /* ---- 사업자(출자자) 구성 ----
       모든 출자자가 매 인출·배당에 지분율대로 비례 참여한다고 가정(pro-rata
       cap table) — 타이밍이 전부 동일하므로 IRR은 출자자별로 동일하고
       금액만 지분율만큼 scale된다. inp.shareholders가 없으면 단일 100%
       출자자로 취급한다. */
    var shList = (inp.shareholders && inp.shareholders.length) ? inp.shareholders : [{ name: '출자자', stakePct: 100 }];
    var shareholders = shList.map(function (s) {
      var frac = (s.stakePct || 0) / 100;
      return {
        name: s.name, stakePct: s.stakePct,
        equityKRWm: equity * frac,
        dividendKRWm: totalDividendKRWm * frac,
        equityIRR: annualize(eIRRp, ppy),
        dividendIRR: annualize(dIRRp, ppy)
      };
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
        investorIRR: annualize(invIRRp, ppy),
        npv: npv(dr, projFlows),
        minDSCR: dscrs.length ? Math.min.apply(null, dscrs) : null,
        avgDSCR: dscrs.length ? dscrs.reduce(function (a, b) { return a + b; }, 0) / dscrs.length : null,
        minCumDSCR: annualCumDscrs.length ? Math.min.apply(null, annualCumDscrs) : null,
        minDSCRAnnual: annualDscrs.length ? Math.min.apply(null, annualDscrs) : null,
        totalDividend: totalDividendKRWm,
        totalRevenue: rows.reduce(function (a, r) { return a + r.revenue; }, 0),
        totalOpex: rows.reduce(function (a, r) { return a + r.opex; }, 0),
        totalInterest: rows.reduce(function (a, r) { return a + r.interest; }, 0),
        totalTax: rows.reduce(function (a, r) { return a + r.tax; }, 0),
        lcoe: pvGen > 0 ? pvOpex / pvGen * 1000 : 0,
        shareholders: shareholders
      }
    };
  }

  var API = { computeModel: computeModel, irr: irr, npv: npv, corpTax: corpTax, buildPeriods: buildPeriods };
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
  root.SolarModel2 = API;
})(typeof globalThis !== 'undefined' ? globalThis : this);
