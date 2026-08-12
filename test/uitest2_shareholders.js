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
  console.log('=== 기본 상태: 출자자 1명(100%) ===');
  console.log('shrow 개수:', d.querySelectorAll('.shrow').length, '(기대 1)');
  console.log('지분합 표시:', d.querySelector('#shsum').textContent);

  console.log('\n=== 출자자 추가 ===');
  fireClick(d, '#shAdd');
  fireClick(d, '#shAdd');
  console.log('shrow 개수:', d.querySelectorAll('.shrow').length, '(기대 3)');

  const stakes = d.querySelectorAll('[data-sh-stake]');
  stakes[0].value = '60'; fireInput(stakes[0]);
  stakes[1].value = '30'; fireInput(stakes[1]);
  stakes[2].value = '10'; fireInput(stakes[2]);
  const names = d.querySelectorAll('[data-sh-name]');
  names[0].value = 'A사'; fireInput(names[0]);
  names[1].value = 'B사'; fireInput(names[1]);
  names[2].value = 'C사'; fireInput(names[2]);
  console.log('지분합 표시(60+30+10=100):', d.querySelector('#shsum').textContent);

  console.log('\n=== 프리셋 로드해도 출자자 구성은 유지되는지 확인 ===');
  fireClick(d, '#loadDangjin');
  setTimeout(() => {
    console.log('shrow 개수(프리셋 로드 후):', d.querySelectorAll('.shrow').length, '(기대 3, 유지)');

    fireClick(d, '#run');
    setTimeout(() => {
      const rows = d.querySelectorAll('#shResults table.tr tbody tr');
      console.log('\n=== 생성 결과 — 사업자별 배분 테이블 ===');
      console.log('행 개수:', rows.length, '(기대 3)');
      rows.forEach(r => console.log('  ', r.textContent.replace(/\s+/g, ' ').trim()));

      console.log('\n=== 지분 불일치(합 100% 아님) 시 경고 ===');
      stakes[0].value = '50'; fireInput(stakes[0]);
      console.log('지분합 표시:', d.querySelector('#shsum').textContent, '(bad 클래스 기대)');
      console.log('shsum class:', d.querySelector('#shsum').className);

      console.log('\n테스트 종료(WINDOW ERROR 없으면 정상)');
    }, 200);
  }, 200);
}, 200);
