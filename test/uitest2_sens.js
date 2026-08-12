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
  console.log('=== 기본 상태: Base + Case1~2, 총 3행 ===');
  const rows = d.querySelectorAll('#sensBox tbody tr');
  const names = Array.from(rows).map(tr => tr.querySelector('[data-sens-f="name"]').value);
  console.log('시나리오 이름:', names.join(', '), '(기대 Base, Case1, Case2)');

  console.log('\n=== "가져오기" 버튼은 Base행에만 있는지 확인 ===');
  rows.forEach((tr, i) => {
    const pullBtn = tr.querySelector('button.pull');
    console.log('  ', names[i], ':', pullBtn ? '"' + pullBtn.textContent + '" 버튼 있음' : '버튼 없음', names[i] === 'Base' ? '(기대: 있음)' : '(기대: 없음)');
  });

  console.log('\n=== placeholder 텍스트가 사라졌는지 확인 ===');
  const anyInput = rows[1].querySelector('[data-sens-f="tariffAbs"]');
  console.log('Case1 판매단가 placeholder:', JSON.stringify(anyInput.placeholder), '(기대 빈 문자열)');

  console.log('\n=== 시나리오 추가 시 이름이 Case3으로 붙는지 확인 ===');
  fireClick(d, '#sensAdd');
  const rows2 = d.querySelectorAll('#sensBox tbody tr');
  console.log('행 개수:', rows2.length, '(기대 4)');
  console.log('마지막 행 이름:', rows2[rows2.length - 1].querySelector('[data-sens-f="name"]').value, '(기대 Case3)');

  console.log('\n테스트 종료(WINDOW ERROR 없으면 정상)');
}, 200);
