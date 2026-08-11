/* v2(분기·5트랜치) 샘플 엑셀 생성 — reference.json의 당진 실측치를 그대로 사용 */
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const M = require('../src2/engine2.js');
const X = require('../src2/xlsxbuild2.js');

const ref = JSON.parse(fs.readFileSync(path.join(__dirname, '../reference/dangjin_reference.json'), 'utf8'));

function findTranche(name) { return ref.tranches.find(t => t.name === name); }
function scheduleFor(letter) {
  const s = ref.repaySchedule_case3[letter];
  return s ? s.map(([, ratio]) => ratio) : null;
}
const trancheDefs = [['선순위A', 'A'], ['선순위B', 'B'], ['선순위C', 'C'], ['선순위D', null], ['후순위', null]];

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

const model = M.computeModel(dangjin);
const wb = X.buildWorkbook(model, ExcelJS);
const outPath = path.join(__dirname, '../dist/당진_태양광발전_재무모델_v2_생성예시.xlsx');
fs.mkdirSync(path.dirname(outPath), { recursive: true });
wb.xlsx.writeFile(outPath).then(() => {
  console.log('v2 sample written:', outPath);
  console.log('projectIRR', (model.kpi.projectIRR * 100).toFixed(2) + '%',
    'equityIRR', (model.kpi.equityIRR * 100).toFixed(2) + '%',
    'minDSCR', model.kpi.minDSCRAnnual.toFixed(4));
});
