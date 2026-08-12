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
  console.log('=== 기본 상태: Base + Case1~5, 총 6행 ===');
  const names = Array.from(d.querySelectorAll('#sensBox tbody tr [data-sens-f="name"]')).map(i => i.value);
  console.log('시나리오 이름:', names.join(', '), '(기대 Base, Case1, Case2, Case3, Case4, Case5)');
  console.log('헤더:', Array.from(d.querySelectorAll('#sensBox thead th')).map(h => h.textContent).filter(Boolean).join(' | '));

  console.log('\n=== 프리셋 상태에서 민감도 실행 시 차단 확인 ===');
  fireClick(d, '#loadDangjin');
  setTimeout(() => {
    fireClick(d, '#sensRun');
    console.log('토스트:', d.querySelector('#toast').textContent);

    console.log('\n=== 프리셋 해제 후 Base행 "↑" 버튼으로 위 입력값 가져오기 ===');
    const cap = d.querySelector('[data-k="capacityMW"]');
    cap.dispatchEvent(new w.Event('input', { bubbles: true }));
    const tariffVal = d.querySelector('[data-k="tariff"]').value;
    const capexVal = d.querySelector('[data-k="capexEok"]').value;
    const opexVal = d.querySelector('[data-k="opexEok"]').value;
    console.log('위 입력값 — 판매단가:', tariffVal, '총사업비:', capexVal, '운영비:', opexVal);

    const baseRow = d.querySelectorAll('#sensBox tbody tr')[0];
    fireClick(d, '#sensBox tbody tr .rm'); // 첫 번째 .rm은 pull(↑) 버튼(rm 클래스 재사용)
    console.log('Base행 판매단가:', baseRow.querySelector('[data-sens-f="tariffAbs"]').value, '(기대', tariffVal + ')');
    console.log('Base행 총사업비:', baseRow.querySelector('[data-sens-f="capexAbs"]').value, '(기대', capexVal + ')');
    console.log('Base행 운영비:', baseRow.querySelector('[data-sens-f="opexAbs"]').value, '(기대', opexVal + ')');
    console.log('Base행 금리(최대 트랜치 운영금리):', baseRow.querySelector('[data-sens-f="rateAbs"]').value);

    fireClick(d, '#sensRun');
    setTimeout(() => {
      const rows = d.querySelectorAll('#sensResults table.tr tbody tr');
      console.log('\n결과 행 개수:', rows.length, '(기대 6)');
      rows.forEach(r => console.log('  ', r.textContent.replace(/\s+/g, ' ').trim()));
      console.log('(Base 실제 숫자로 채워졌어도 나머지 Case가 다 빈칸이면 전부 Base와 동일해야 정상)');

      console.log('\n테스트 종료(WINDOW ERROR 없으면 정상)');
    }, 200);
  }, 200);
}, 200);
