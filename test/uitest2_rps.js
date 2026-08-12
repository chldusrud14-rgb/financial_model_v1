const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');
const html = fs.readFileSync(path.join(__dirname, '../dist/태양광_재무모델_생성기_v2.html'), 'utf8');
const dom = new JSDOM(html, { runScripts: 'dangerously', pretendToBeVisual: true, resources: undefined, url: 'file://' + path.join(__dirname, '../dist/x.html') });
const w = dom.window;
w.ExcelJS = require('exceljs');
w.addEventListener('error', e => console.log('WINDOW ERROR:', e.message, e.error && e.error.stack));

function fireClick(d, sel) { d.querySelector(sel).dispatchEvent(new w.Event('click', { bubbles: true })); }
function fireInput(el) { el.dispatchEvent(new w.Event('input', { bubbles: true })); }

setTimeout(() => {
  const d = w.document;
  console.log('=== 기본 상태(RPS 비중 0) ===');
  const recW = d.querySelector('[data-k="recWeight"]');
  console.log('REC 가중치 값:', recW.value, '(기대 1) / disabled:', recW.disabled, '(기대 true)');

  fireClick(d, '#run');
  setTimeout(() => {
    const hero0 = d.querySelector('#kpis .kpiHero .hv').textContent;
    console.log('RPS 0%(전량 PPA) 생성 결과:', hero0);

    console.log('\n=== RPS 비중을 40%로 설정 ===');
    const rps = d.querySelector('[data-k="rpsShare"]');
    rps.value = '40'; fireInput(rps);
    console.log('REC 가중치 disabled:', recW.value === '1' ? recW.disabled : recW.disabled, '(기대 false)');
    const smp = d.querySelector('[data-k="smpPrice"]');
    const rec = d.querySelector('[data-k="recPrice"]');
    smp.value = '130'; fireInput(smp);
    rec.value = '80'; fireInput(rec);
    recW.value = '1.2'; fireInput(recW);

    fireClick(d, '#run');
    setTimeout(() => {
      const hero1 = d.querySelector('#kpis .kpiHero .hv').textContent;
      console.log('RPS 40% 적용 후 생성 결과(값이 달라지면 트랙 반영된 것):');
      console.log('  총영업수익 포함 메타노트:', d.querySelector('#metaNote').textContent);
      console.log('  Equity IRR(배당):', hero1, '(RPS 0% 케이스', hero0, '와 달라야 함)');

      console.log('\n=== RPS 비중 다시 0으로 ===');
      rps.value = '0'; fireInput(rps);
      console.log('REC 가중치 값:', recW.value, '(기대 1) / disabled:', recW.disabled, '(기대 true)');

      console.log('\n테스트 종료(WINDOW ERROR 없으면 정상)');
    }, 200);
  }, 200);
}, 200);
