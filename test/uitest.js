const fs=require('fs');
const {JSDOM}=require('jsdom');
const html=fs.readFileSync('../dist/태양광_재무모델_생성기.html','utf8');
const dom=new JSDOM(html,{runScripts:'dangerously',pretendToBeVisual:true,resources:undefined});
const w=dom.window;
w.ExcelJS=require('exceljs');
// 외부 스크립트(CDN)는 로드되지 않으므로 인라인만 실행됨
setTimeout(()=>{
  const d=w.document;
  console.log('KPI 카드:', d.querySelectorAll('#kpis .kpi').length);
  d.querySelectorAll('#kpis .kpi').forEach(k=>console.log('  ',k.querySelector('.k').textContent,'=',k.querySelector('.v').textContent));
  console.log('입력 필드:', d.querySelectorAll('[data-k]').length);
  console.log('FCF chart bars:', d.querySelectorAll('#fcf rect').length);
  console.log('DSCR points:', d.querySelectorAll('#dscr circle').length);
  console.log('검증 항목:', d.querySelectorAll('#checks .chk').length);
  d.querySelectorAll('#checks .chk').forEach(c=>console.log('  ['+c.className.replace('chk ','')+']',c.querySelector('b').textContent));
  console.log('민감도 블록:', d.querySelectorAll('#sens svg').length);
  console.log('표 행:', d.querySelectorAll('#tbl tbody tr').length, '/ 열', d.querySelectorAll('#tbl thead th').length);
  console.log('note capacityFactor:', d.querySelector('[data-note="capacityFactor"]').textContent);
  console.log('note capex:', d.querySelector('[data-note="capexEok"]').textContent);
  // 프리셋 전환 테스트
  d.querySelector('[data-preset="dangjin"]').dispatchEvent(new w.Event('click'));
  setTimeout(()=>{
    console.log('--- 당진 프리셋 적용 후');
    d.querySelectorAll('#kpis .kpi').forEach(k=>console.log('  ',k.querySelector('.k').textContent,'=',k.querySelector('.v').textContent));
    console.log('에러 없음');
  },400);
},400);
w.addEventListener('error',e=>console.log('WINDOW ERROR:',e.message));
