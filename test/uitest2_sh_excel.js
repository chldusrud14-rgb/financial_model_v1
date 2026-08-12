// 사업자 구성(지분) 화면 입력이 실제로 다운로드되는 엑셀 파일에도
// 반영되는지, 브라우저(jsdom) UI 흐름을 그대로 태워서 끝까지 확인한다.
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

  // 출자자 2명 추가 + 지분 40/60 입력
  fireClick(d, '#shAdd');
  const stakes = d.querySelectorAll('[data-sh-stake]');
  const names = d.querySelectorAll('[data-sh-name]');
  names[0].value = '갑'; fireInput(names[0]);
  names[1].value = '을'; fireInput(names[1]);
  stakes[0].value = '40'; fireInput(stakes[0]);
  stakes[1].value = '60'; fireInput(stakes[1]);
  console.log('지분합 표시:', d.querySelector('#shsum').textContent);

  fireClick(d, '#run');
  setTimeout(() => {
    // download()가 쓰는 model은 클로저 안에 있어 직접 접근 불가하니,
    // window.URL.createObjectURL을 가로채서 buildWorkbook에 실제로
    // 넘어간 workbook을 파일로 저장해 검사한다.
    const origBuildWorkbook = w.SolarXlsx2.buildWorkbook;
    let captured = null;
    w.SolarXlsx2.buildWorkbook = function (model, ExcelJSLib) {
      captured = model;
      return origBuildWorkbook(model, ExcelJSLib);
    };

    // jsdom엔 URL.createObjectURL이 없어 다운로드 트리거용 a.click()까지는
    // 못 가더라도, buildWorkbook 호출 시점까지는 도달하는지 확인 가능.
    w.URL.createObjectURL = function () { return 'blob:mock'; };
    w.HTMLAnchorElement.prototype.click = function () {}; // 실제 다운로드 방지

    fireClick(d, '#xls');
    setTimeout(() => {
      if (!captured) { console.log('FAIL: buildWorkbook 호출 안 됨'); return; }
      console.log('\nmodel.kpi.shareholders:', JSON.stringify(captured.kpi.shareholders));

      origBuildWorkbook(captured, w.ExcelJS).xlsx.writeFile(
        path.join(__dirname, '../dist/_sh_excel_test.xlsx')
      ).then(() => {
        console.log('엑셀 파일 작성 완료 — openpyxl로 별도 검사 필요');
        console.log('\n테스트 종료(WINDOW ERROR 없으면 정상)');
      });
    }, 200);
  }, 200);
}, 200);
