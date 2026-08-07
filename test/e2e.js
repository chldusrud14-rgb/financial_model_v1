const fs = require('fs');
const { JSDOM } = require('jsdom');
const html = fs.readFileSync('dist/태양광_재무모델_생성기.html', 'utf8');
const rt = process.argv[2] || '1';

const BASE = {
  projectName: 'E2E검증_방식' + rt,
  capacityMW: 45.5, capacityFactor: 14.8, capexEok: 620, opexEok: 22,
  gearing: 82, rateO: 5.85, tariff: 172, operationYears: 20,
  constructionStart: '2027-04', constructionMonths: 12,
  dsraEok: 25, rateC: 6.1, graceYears: 2, repayYears: 15,
  repayType: rt, payPerYear: 2, tariffEscal: 1.2, opexEscal: 1.8,
  degradation: 0.5, auxRate: 0, depRatio: 95, depYears: 18,
  taxMode: 1, taxFlat: 21, lossRate: 80, decomEok: 16,
  discount: 6.0, divStartYear: 3, dsraMonths: 6, minCash: 10, divDSCR: 1.1
};

const dom = new JSDOM(html, { runScripts: 'dangerously', pretendToBeVisual: true });
const w = dom.window, d = w.document;
w.ExcelJS = require('exceljs');
let errs = [];
w.addEventListener('error', e => errs.push(e.message));

setTimeout(() => {
  for (const k in BASE) {
    const i = d.querySelector('[data-k="' + k + '"]');
    if (!i) { console.log('!! 필드없음', k); continue; }
    i.value = String(BASE[k]);
    i.dispatchEvent(new w.Event('input', { bubbles: true }));
    i.dispatchEvent(new w.Event('change', { bubbles: true }));
  }
  setTimeout(() => {
    const kpi = {};
    d.querySelectorAll('#kpis .kpi').forEach(c => {
      kpi[c.querySelector('.k').textContent.trim()] = c.querySelector('.v').textContent.trim();
    });
    const chk = [];
    d.querySelectorAll('#checks .chk').forEach(c =>
      chk.push(c.className.replace('chk ', '') + ':' + c.querySelector('b').textContent.trim()));
    console.log('[방식' + rt + '] 화면 KPI:', JSON.stringify(kpi, null, 0));
    console.log('[방식' + rt + '] 상환방식 필드값 =', d.querySelector('[data-k="repayType"]').value);
    console.log('[방식' + rt + '] 검증:', chk.join(' | '));

    const OB = w.Blob; let cap = null;
    w.Blob = function (p, o) { cap = p[0]; return new OB(p, o); };
    w.URL.createObjectURL = () => 'blob:x'; w.URL.revokeObjectURL = () => {};
    d.querySelector('#xls').click();
    setTimeout(() => {
      if (!cap) { console.log('!! 다운로드 실패'); process.exit(1); }
      fs.writeFileSync('./e2e_t' + rt + '.xlsx', Buffer.from(cap));
      fs.writeFileSync('./e2e_kpi' + rt + '.json', JSON.stringify(kpi));
      console.log('[방식' + rt + '] 다운로드 OK, 런타임에러:', errs.length ? errs.join(';') : '없음');
      process.exit(0);
    }, 2500);
  }, 800);
}, 600);
