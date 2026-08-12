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
  console.log('=== 기본 상태: 시나리오 2행(Upside/Downside, Base는 행 아님) ===');
  console.log('시나리오 행 개수:', d.querySelectorAll('#sensBox tbody tr').length, '(기대 2)');
  console.log('헤더:', Array.from(d.querySelectorAll('#sensBox thead th')).map(h => h.textContent).join(' | '));

  console.log('\n=== 프리셋 상태에서 민감도 실행 시 차단되는지 확인 ===');
  fireClick(d, '#loadDangjin');
  setTimeout(() => {
    fireClick(d, '#sensRun');
    console.log('토스트:', d.querySelector('#toast').textContent);
    console.log('결과 테이블 존재:', !!d.querySelector('#sensResults table'), '(기대 false)');

    console.log('\n=== 프리셋 해제(일반 폼) 후 아무 값도 안 채우고 민감도 실행 — Base만 자동 추가돼서 3행(Upside/Downside는 빈칸=기본값 사용) ===');
    const cap = d.querySelector('[data-k="capacityMW"]');
    cap.dispatchEvent(new w.Event('input', { bubbles: true }));
    fireClick(d, '#sensRun');
    setTimeout(() => {
      const rows = d.querySelectorAll('#sensResults table.tr tbody tr');
      console.log('결과 행 개수:', rows.length, '(기대 3: Base+Upside+Downside)');
      rows.forEach(r => console.log('  ', r.textContent.replace(/\s+/g, ' ').trim()));
      console.log('(빈칸 시나리오는 Base와 동일한 결과가 나와야 정상)');

      console.log('\n=== Downside에 절대값(판매단가 100원, 총투자비 2000억) 입력 후 재실행 ===');
      const dsRow = d.querySelectorAll('#sensBox tbody tr')[1];
      dsRow.querySelector('[data-sens-f="tariffAbs"]').value = '100';
      fireInput(dsRow.querySelector('[data-sens-f="tariffAbs"]'));
      dsRow.querySelector('[data-sens-f="capexAbs"]').value = '2000';
      fireInput(dsRow.querySelector('[data-sens-f="capexAbs"]'));
      fireClick(d, '#sensRun');
      setTimeout(() => {
        const rows2 = d.querySelectorAll('#sensResults table.tr tbody tr');
        console.log('결과:');
        rows2.forEach(r => console.log('  ', r.textContent.replace(/\s+/g, ' ').trim()));
        console.log('(Downside가 Base/Upside와 값이 달라야 정상 — 판매단가 낮추고 투자비 늘렸으니 IRR도 낮아야 함)');

        console.log('\n=== 시나리오 추가 ===');
        fireClick(d, '#sensAdd');
        console.log('시나리오 행 개수:', d.querySelectorAll('#sensBox tbody tr').length, '(기대 3)');

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
