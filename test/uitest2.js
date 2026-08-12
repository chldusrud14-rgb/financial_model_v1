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
  console.log('=== 기본 폼(당진 프리셋 로드 전) ===');
  console.log('입력 필드(core):', d.querySelectorAll('[data-k]').length, '(기대 30)');
  console.log('공사비 지출곡선 입력:', d.querySelectorAll('[data-spend]').length);

  fireClick(d, '#run');
  setTimeout(() => {
    const genericKpis = {};
    d.querySelectorAll('#kpis .kpi').forEach(k => { genericKpis[k.querySelector('.k').textContent] = k.querySelector('.v').textContent; });
    const genericHero = d.querySelector('#kpis .kpiHero .hv').textContent;
    console.log('기본값 생성 결과(범용 근사, 원본과 다름):', genericHero, genericKpis['최소 DSCR(연 합산)']);

    console.log('\n=== 당진 FS 불러오기 클릭 ===');
    fireClick(d, '#loadDangjin');
    setTimeout(() => {
      console.log('사업명:', d.querySelector('[data-k="projectName"]').value);
      console.log('공사비 지출곡선(1분기):', d.querySelector('[data-spend="0"]').value, '(기대 348.70)');
      console.log('선순위A 방식 select:', d.querySelector('select[data-tr="A"]').value, '(기대 3)');

      fireClick(d, '#run');
      setTimeout(() => {
        console.log('\n=== 당진 프리셋 생성 결과 (test_ops.js 검증값과 대조) ===');
        const kpis = {};
        d.querySelectorAll('#kpis .kpi').forEach(k => { kpis[k.querySelector('.k').textContent] = k.querySelector('.v').textContent; });
        kpis['Equity IRR (배당)'] = d.querySelector('#kpis .kpiHero .hv').textContent;
        const expect = {
          'Project IRR 세후': '7.86', 'Equity IRR (FCFE) 세후': '13.33', 'Equity IRR (배당)': '12.12',
          '최소 DSCR(연 합산)': '1.218'
        };
        let allOk = true;
        Object.keys(expect).forEach(k => {
          const got = kpis[k];
          const ok = got && got.startsWith(expect[k].slice(0, 4));
          if (!ok) allOk = false;
          console.log((ok ? '  OK  ' : '  << ') + k + ': ' + got + ' (기대 ' + expect[k] + '대)');
        });
        console.log('\n' + (allOk ? '당진 프리셋 → 검증된 원본값과 일치' : '불일치 있음 — 위 << 확인'));

        console.log('\n=== 프리셋 로드 후 값 수정 시 usingPreset 해제 확인 ===');
        const cap = d.querySelector('[data-k="capacityMW"]');
        cap.value = '50';
        cap.dispatchEvent(new w.Event('input', { bubbles: true }));
        fireClick(d, '#run');
        setTimeout(() => {
          const kpis2 = {};
          d.querySelectorAll('#kpis .kpi').forEach(k => { kpis2[k.querySelector('.k').textContent] = k.querySelector('.v').textContent; });
          kpis2['Equity IRR (배당) [히어로]'] = d.querySelector('#kpis .kpiHero .hv').textContent;
          console.log('용량 50MW(부채는 원래 규모 그대로라 과중) 재생성 결과 전체:');
          Object.keys(kpis2).forEach(k => console.log('  ', k, '=', kpis2[k]));
          console.log('토스트 메시지:', d.querySelector('#toast').textContent);
          console.log('\n테스트 종료(WINDOW ERROR 없으면 정상)');
        }, 200);
      }, 200);
    }, 200);
  }, 200);
}, 200);
