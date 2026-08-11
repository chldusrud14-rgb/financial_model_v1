const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');
const html = fs.readFileSync(path.join(__dirname, '../dist/태양광_재무모델_생성기_v2.html'), 'utf8');
const dom = new JSDOM(html, { runScripts: 'dangerously', pretendToBeVisual: true, resources: undefined, url: 'file://' + path.join(__dirname, '../dist/x.html') });
const w = dom.window;
w.ExcelJS = require('exceljs');
w.addEventListener('error', e => console.log('WINDOW ERROR:', e.message, e.error && e.error.stack));

setTimeout(() => {
  const d = w.document;
  console.log('입력 필드(core):', d.querySelectorAll('[data-k]').length, '(기대 26)');
  console.log('트랜치 입력:', d.querySelectorAll('[data-tr]').length, '(기대 5행*6필드=30)');
  console.log('트랜치 방식 select:', d.querySelectorAll('select[data-tr]').length, '(기대 5)');
  console.log('공사비 지출곡선 입력:', d.querySelectorAll('[data-spend]').length);

  d.querySelector('#run').dispatchEvent(new w.Event('click'));
  setTimeout(() => {
    console.log('--- 생성 결과 ---');
    const kpis = d.querySelectorAll('#kpis .kpi');
    console.log('KPI 카드:', kpis.length);
    kpis.forEach(k => console.log('  ', k.querySelector('.k').textContent, '=', k.querySelector('.v').textContent));
    console.log('메타:', d.querySelector('#metaNote').textContent);
    console.log('Excel 버튼 활성화:', !d.querySelector('#xls').disabled);

    d.querySelector('#xls').dispatchEvent(new w.Event('click'));
    setTimeout(() => {
      console.log('다운로드 클릭 후 버튼 텍스트:', d.querySelector('#xls').textContent);
      console.log('테스트 종료(에러 없으면 정상)');
    }, 600);
  }, 200);
}, 200);
