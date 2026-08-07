const M = require('../src2/engine2.js');

const dangjin = {
  projectName: '당진 태양광발전', ppy: 4,
  capacityMW: 99.998, capacityFactor: 15.71, degradation: 0.5, auxRate: 0,
  constructionStart: '2024-06', constructionMonths: 16, operationYears: 20,
  capexEok: 1410.68821, dsraEok: 50,
  opexEok: 49.8, opexEscal: 0.7,
  spendCurve: [34870.32, 11153.04, 22227.33, 22239.02, 23245.03, 27334.07],
  tariff: 154.8, tariffEscal: 0,
  equityEok: 150, equityOrder: 1,
  tranches: [
    { name: '선순위A', amountEok: 500, order: 2, rateC: 5.6, rateO: 5.6, graceYears: 2, repayYears: 16, method: 1 },
    { name: '선순위B', amountEok: 500, order: 2, rateC: 5.5, rateO: 5.5, graceYears: 2, repayYears: 16, method: 1 },
    { name: '선순위C', amountEok: 350, order: 4, rateC: 5.5, rateO: 5.5, graceYears: 1.25, repayYears: 16, method: 1 },
    { name: '선순위D', amountEok: 0, order: 3, rateC: 2.25, rateO: 2.25, graceYears: 5, repayYears: 10, method: 1 },
    { name: '후순위', amountEok: 0, order: 1, rateC: 5.65, rateO: 5.65, graceYears: 15, repayYears: 2.75, method: 1 }
  ],
  depRatio: 95, depYears: 20, taxMode: 1, taxFlat: 21, lossRate: 80, decomEok: 20,
  dsraMonths: 6, minCash: 10, divDSCR: 1.1, divStartYear: 2, discount: 5.5
};

const r = M.computeModel(dangjin);

console.log('=== 기간 축 ===');
console.log('총 기간 수:', r.periods.length, '| 첫 기말:', r.periods[0].endStr, '| 마지막:', r.periods[r.periods.length - 1].endStr);
const con = r.periods.filter(p => p.isCon);
console.log('건설기간 수:', con.length, '| 건설월수 합:', con.reduce((a, p) => a + p.conMonths, 0));
console.log('운영월수 합:', r.periods.reduce((a, p) => a + p.opMonths, 0));

console.log('\n=== 건설기간 인출 / 건설이자 (원본 대조) ===');
const TGT = {
  '2024-06-30': [15000, 9935.16, 9935.16, 0, 0, 0, 0],
  '2024-09-30': [0, 5714.37, 5714.37, 0, 139.09, 136.61, 0],
  '2024-12-31': [0, 11330.80, 11330.80, 0, 219.09, 215.18, 0],
  '2025-03-31': [0, 11493.86, 11493.86, 0, 377.72, 370.98, 0],
  '2025-06-30': [0, 11525.80, 11525.80, 1261.09, 538.64, 529.02, 0],
  '2025-09-30': [0, 0, 0, 33738.91, 700.00, 687.50, 17.34]
};
console.log('기말        | 자본금    A인출     B인출     C인출  |  IDC_A   IDC_B   IDC_C');
r.periods.forEach((p, i) => {
  if (i > r.con.codIdx) return;
  const d = k => r.con.draws[k][i], ic = k => r.con.idc[k][i];
  const row = [d(0), d(1), d(2), d(3), ic(1), ic(2), ic(3)];
  console.log(p.endStr, '|', row.map(v => v.toFixed(2).padStart(9)).join(' '));
  const t = TGT[p.endStr];
  if (t) {
    const diff = row.map((v, j) => Math.abs(v - t[j]));
    const bad = diff.map((v, j) => v > 0.5 ? j : -1).filter(j => j >= 0);
    console.log('  원본     |', t.map(v => v.toFixed(2).padStart(9)).join(' '), bad.length ? '  << 불일치 idx ' + bad : '  OK');
  }
});

console.log('\n=== 합계 대조 ===');
const idcT = r.con.idcByTranche;
console.log('IDC  A:', idcT[1].toFixed(3), '(원본 1974.549)');
console.log('IDC  B:', idcT[2].toFixed(3), '(원본 1939.290)');
console.log('IDC  C:', idcT[3].toFixed(3), '(원본 17.340)');
console.log('IDC 합:', r.idc.toFixed(3), '(원본 3931.179)');
console.log('총투자비:', r.tic.toFixed(2), '(원본 145000)');
console.log('인출합계 A/B/C:', r.con.drawn.slice(1, 4).map(v => v.toFixed(2)).join(' / '), '(원본 50000/50000/35000)');
