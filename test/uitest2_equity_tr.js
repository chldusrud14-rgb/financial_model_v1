const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');
const html = fs.readFileSync(path.join(__dirname, '../dist/태양광_재무모델_생성기_v2.html'), 'utf8');
const dom = new JSDOM(html, { runScripts: 'dangerously', pretendToBeVisual: true, resources: undefined, url: 'file://' + path.join(__dirname, '../dist/x.html') });
const w = dom.window;
w.ExcelJS = require('exceljs');
w.addEventListener('error', e => console.log('WINDOW ERROR:', e.message, e.error && e.error.stack));

function fireInput(el) { el.dispatchEvent(new w.Event('input', { bubbles: true })); }
function fireClick(d, sel) { d.querySelector(sel).dispatchEvent(new w.Event('click', { bubbles: true })); }

setTimeout(() => {
  const d = w.document;
  console.log('=== 카드 순서 확인 ===');
  const titles = Array.from(d.querySelectorAll('.card .hd h4')).map(h => h.textContent);
  console.log(titles.join(' -> '));

  console.log('\n=== 운영비가 총사업비 옆(essential)에 있는지 ===');
  const opexEl = d.querySelector('[data-k="opexEok"]');
  console.log('운영비 필드가 상세가정 아코디언 밖(essential)에 있음:', !!opexEl && !opexEl.closest('details'));

  console.log('\n=== 자기자본비율 -> 자본금 동기화 ===');
  const capex = d.querySelector('[data-k="capexEok"]');
  const equity = d.querySelector('[data-k="equityEok"]');
  const ratio = d.querySelector('[data-k="equityRatioPct"]');
  capex.value = '1000'; fireInput(capex);
  ratio.value = '15'; fireInput(ratio);
  console.log('자본금(총사업비 1000억 * 15%):', equity.value, '(기대 150.00)');

  console.log('\n=== 자본금 -> 자기자본비율 동기화 ===');
  equity.value = '200'; fireInput(equity);
  console.log('자기자본비율(200/1000*100):', ratio.value, '(기대 20.00)');

  console.log('\n=== 트랜치 간편설정 ===');
  fireClick(d, '#trQuick');
  const amtA = d.querySelector('input[data-tr="A"][data-f="amountEok"]');
  const amtB = d.querySelector('input[data-tr="B"][data-f="amountEok"]');
  const methodA = d.querySelector('select[data-tr="A"]');
  console.log('선순위A 금액(총사업비1000-자본금200=800 기대):', amtA.value);
  console.log('선순위B 금액(기대 0):', amtB.value);
  console.log('선순위A 방식(기대 1):', methodA.value);
  console.log('토스트:', d.querySelector('#toast').textContent);

  console.log('\n테스트 종료(WINDOW ERROR 없으면 정상)');
}, 200);
