const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');
const html = fs.readFileSync(path.join(__dirname, '../dist/태양광_재무모델_생성기_v2.html'), 'utf8');
const dom = new JSDOM(html, { runScripts: 'dangerously', pretendToBeVisual: true, resources: undefined, url: 'file://' + path.join(__dirname, '../dist/x.html') });
const w = dom.window;
w.ExcelJS = require('exceljs');
w.addEventListener('error', e => console.log('WINDOW ERROR:', e.message, e.error && e.error.stack));

function fireClick(d, sel) { d.querySelector(sel).dispatchEvent(new w.Event('click', { bubbles: true })); }

setTimeout(() => {
  const d = w.document;
  console.log('=== 기본 상태: 시나리오 3행(Base/Upside/Downside) ===');
  console.log('시나리오 행 개수:', d.querySelectorAll('#sensBox tbody tr').length, '(기대 3)');

  console.log('\n=== 프리셋 상태에서 민감도 실행 시 차단되는지 확인 ===');
  fireClick(d, '#loadDangjin');
  setTimeout(() => {
    fireClick(d, '#sensRun');
    console.log('토스트:', d.querySelector('#toast').textContent);
    console.log('결과 테이블 존재:', !!d.querySelector('#sensResults table'), '(기대 false)');

    console.log('\n=== 프리셋 해제(일반 폼) 후 민감도 실행 ===');
    const cap = d.querySelector('[data-k="capacityMW"]');
    cap.dispatchEvent(new w.Event('input', { bubbles: true })); // 값은 그대로, usingPreset만 해제
    fireClick(d, '#sensRun');
    setTimeout(() => {
      const rows = d.querySelectorAll('#sensResults table.tr tbody tr');
      console.log('결과 행 개수:', rows.length, '(기대 3)');
      rows.forEach(r => console.log('  ', r.textContent.replace(/\s+/g, ' ').trim()));

      console.log('\n=== 시나리오 추가 후 재실행 ===');
      fireClick(d, '#sensAdd');
      console.log('시나리오 행 개수:', d.querySelectorAll('#sensBox tbody tr').length, '(기대 4)');
      fireClick(d, '#sensRun');
      setTimeout(() => {
        console.log('결과 행 개수:', d.querySelectorAll('#sensResults table.tr tbody tr').length, '(기대 4)');

        console.log('\n=== 생성 버튼 누르면 민감도 결과 초기화되는지 확인 ===');
        fireClick(d, '#run');
        setTimeout(() => {
          console.log('민감도 결과 잔존 여부:', !!d.querySelector('#sensResults table'), '(기대 false)');
          console.log('\n테스트 종료(WINDOW ERROR 없으면 정상)');
        }, 200);
      }, 200);
    }, 200);
  }, 200);
}, 200);
