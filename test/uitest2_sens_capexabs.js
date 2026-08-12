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
  console.log('=== 헤더에 "총사업비 변동(억원)" 열 존재 확인 ===');
  const heads = Array.from(d.querySelectorAll('#sensBox thead th')).map(h => h.textContent);
  console.log(heads.join(' | '));

  // 값 수정(프리셋 해제 목적) 후 -5억/+10억 시나리오로 덮어쓰기
  const cap = d.querySelector('[data-k="capacityMW"]');
  fireInput(cap);

  const rows = d.querySelectorAll('#sensBox tbody tr');
  const absInputs = Array.from(rows).map(tr => tr.querySelector('input[data-sens-f="capexAbsEok"]'));
  absInputs[0].value = '-5'; fireInput(absInputs[0]);   // Base -> -5억
  absInputs[1].value = '10'; fireInput(absInputs[1]);   // Upside -> +10억
  // capexPct는 0으로 둬서(기본 시나리오값 그대로) 절대금액만의 순수 효과를 본다
  const capexPctInputs = Array.from(rows).map(tr => tr.querySelector('input[data-sens-f="capexPct"]'));
  capexPctInputs.forEach(inp => { inp.value = '0'; fireInput(inp); });

  fireClick(d, '#sensRun');
  setTimeout(() => {
    const resultRows = d.querySelectorAll('#sensResults table.tr tbody tr');
    console.log('\n결과:');
    resultRows.forEach(r => console.log('  ', r.textContent.replace(/\s+/g, ' ').trim()));
    console.log('\n(총사업비만 다르면 IRR도 서로 달라야 함 — Base(-5억)와 Upside(+10억)의 Project IRR이 다른지 확인)');
    console.log('\n테스트 종료(WINDOW ERROR 없으면 정상)');
  }, 200);
}, 200);
