/* 운영기간(v2 engine2.js runOps) ↔ 원본 대조.
   tranches/상환스케줄/목표값을 reference json에서 그대로 읽어 손으로 옮겨적는
   과정에서 생기는 오타를 없앤다. */
const fs = require('fs');
const path = require('path');
const M = require('../src2/engine2.js');

const ref = JSON.parse(fs.readFileSync(path.join(__dirname, '../reference/dangjin_reference.json'), 'utf8'));

function findTranche(name) { return ref.tranches.find(t => t.name === name); }
function scheduleFor(letter) {
  const s = ref.repaySchedule_case3[letter];
  return s ? s.map(([, ratio]) => ratio) : null;
}

const trancheDefs = [
  ['선순위A', 'A'], ['선순위B', 'B'], ['선순위C', 'C'], ['선순위D', null], ['후순위', null]
];

const dangjin = {
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
  opexItems: ref.opexItems.map(it => ({ annualKRWm: it.annualKRWm, escal: it.escalRate * 100, senior: it.senior })),
  spendCurve: ref.spendCurve_KRWm,
  tariffTracks: ref.tariffTracks.map(t => ({ share: t.share, price: t.price })),
  seasonality: Object.fromEntries(Object.entries(ref.seasonality).map(([m, v]) => [Number(m), v])),
  equityEok: ref.funding.equity / 100, equityOrder: ref.funding.equityOrder,
  tranches: trancheDefs.map(([name, letter]) => {
    const t = findTranche(name);
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

const r = M.computeModel(dangjin);
const k = r.kpi;

function cmp(label, got, want, tol) {
  const diff = Math.abs(got - want);
  const bad = diff > tol;
  console.log((bad ? '  << ' : '  OK  ') + label + ': ' + got.toFixed(4) + ' (원본 ' + want.toFixed(4) + ', diff ' + diff.toFixed(4) + ')');
  return !bad;
}

console.log('=== 운영기간 결과 대조 (reference.results) ===');
let ok = true;
ok = cmp('총영업수익(20yr)', k.totalRevenue, ref.results.총영업수익_20yr, 0.01) && ok;
ok = cmp('총영업비용(20yr)', k.totalOpex, ref.results.총영업비용_20yr, 0.01) && ok;
ok = cmp('총선순위이자', k.totalInterest, ref.results.총선순위이자, 0.01) && ok;
ok = cmp('총법인세', k.totalTax, ref.results.총법인세, 0.01) && ok;
console.log('  (참고) 분기 최소DSCR:', k.minDSCR.toFixed(4), '/ 연 합산 최소DSCR:', k.minDSCRAnnual.toFixed(4));
ok = cmp('최소단순DSCR(연 합산 기준)', k.minDSCRAnnual, ref.results.최소단순DSCR, 0.01) && ok;
ok = cmp('최소누적DSCR', k.minCumDSCR, ref.results.최소누적DSCR, 0.01) && ok;
ok = cmp('연차배당', k.totalDividend, ref.results.연차배당 + ref.results.청산배당, 0.1) && ok;
ok = cmp('projectIRR(세전)', k.projectIRRPre, ref.results.projectIRR_preTax, 0.001) && ok;
ok = cmp('projectIRR(세후)', k.projectIRR, ref.results.projectIRR_postTax, 0.001) && ok;
ok = cmp('equityIRR(FCFE)', k.equityIRR, ref.results.equityIRR_FCFE, 0.001) && ok;
ok = cmp('investorIRR', k.investorIRR, ref.results.investorIRR, 0.001) && ok;
ok = cmp('dividendIRR', k.dividendIRR, ref.results.equityIRR_dividend, 0.001) && ok;

console.log('\n' + (ok ? '핵심 채무/세무/DSCR 지표 일치' : '불일치 항목 있음 — 위 << 표시 확인'));
process.exitCode = ok ? 0 : 1;
