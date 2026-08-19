// 총사업비/운영비 항목별 입력 토글 검증: (1) 기본(합계만) 상태에서도
// 엑셀에 항목명이 나오는지, (2) 토글 켜고 항목 입력하면 합계 필드가
// 자동 갱신되는지, (3) 실제 다운로드되는 엑셀에 항목별 금액이 반영되는지.
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
function fireChange(el) { el.dispatchEvent(new w.Event('change', { bubbles: true })); }

setTimeout(() => {
  const d = w.document;
  console.log('=== 기본 상태: 토글 꺼짐, 총사업비/운영비는 직접 입력 가능 ===');
  const capexToggle = d.querySelector('#capexDetailToggle');
  const opexToggle = d.querySelector('#opexDetailToggle');
  console.log('capex 토글:', capexToggle.checked, '(기대 false)');
  console.log('capexEok readonly:', d.querySelector('[data-k="capexEok"]').readOnly, '(기대 false)');

  const captured = {};
  const origBuildWorkbook = w.SolarXlsx2.buildWorkbook;
  w.SolarXlsx2.buildWorkbook = function (model, ExcelJSLib) {
    captured.model = model;
    return origBuildWorkbook(model, ExcelJSLib);
  };
  w.URL.createObjectURL = function () { return 'blob:mock'; };
  w.HTMLAnchorElement.prototype.click = function () {};

  fireClick(d, '#run');
  setTimeout(() => {
    fireClick(d, '#xls');
    setTimeout(() => {
    console.log('model.capexItems 개수:', captured.model.capexItems.length, '(기대 12)');
    console.log('model.capexItems[0]:', JSON.stringify(captured.model.capexItems[0]), '(금액 null 기대)');
    console.log('model.opexDisplayItems 존재:', !!captured.model.opexDisplayItems, '(기대 true, opexItems 안 켰으므로)');

      origBuildWorkbook(captured.model, w.ExcelJS).xlsx.writeFile(
        path.join(__dirname, '../dist/_items_test1.xlsx')
      ).then(() => {
        console.log('기본모드 엑셀 작성 완료(항목명만, 금액 없음 기대)');

        console.log('\n=== 총사업비 항목별 입력 켜기 ===');
        capexToggle.checked = true; fireChange(capexToggle);
        console.log('capexEok readonly:', d.querySelector('[data-k="capexEok"]').readOnly, '(기대 true)');
        const capexRows = d.querySelectorAll('#capexItemBox tbody tr');
        console.log('capex 항목 행 개수:', capexRows.length, '(기대 12)');
        const amt0 = capexRows[0].querySelector('[data-capex-f="amountEok"]');
        amt0.value = '500'; fireInput(amt0);
        const amt1 = capexRows[1].querySelector('[data-capex-f="amountEok"]');
        amt1.value = '100'; fireInput(amt1);
        console.log('총사업비 필드 값(500+100=600 기대):', d.querySelector('[data-k="capexEok"]').value);

        console.log('\n=== 운영비 항목별 입력 켜기 ===');
        opexToggle.checked = true; fireChange(opexToggle);
        const opexRows = d.querySelectorAll('#opexItemBox tbody tr');
        console.log('opex 항목 행 개수:', opexRows.length, '(기대 12)');
        const oamt0 = opexRows[0].querySelector('[data-opex-f="amountEok"]');
        oamt0.value = '10'; fireInput(oamt0);
        const oamt1 = opexRows[1].querySelector('[data-opex-f="amountEok"]');
        oamt1.value = '15'; fireInput(oamt1);
        console.log('운영비 필드 값(10+15=25 기대):', d.querySelector('[data-k="opexEok"]').value);

        fireClick(d, '#run');
        setTimeout(() => {
          fireClick(d, '#xls');
          setTimeout(() => {
            console.log('model.capexItems[0] 금액(500 기대):', captured.model.capexItems[0].amountEok);
            console.log('model.inp.opexItems 존재(엑셀 실제 계산 반영 기대):', !!captured.model.inp.opexItems);
            console.log('model.opexDisplayItems 존재(opexItems 켰으므로 기대 false/undefined):', captured.model.opexDisplayItems);

            origBuildWorkbook(captured.model, w.ExcelJS).xlsx.writeFile(
              path.join(__dirname, '../dist/_items_test2.xlsx')
            ).then(() => {
              console.log('상세모드 엑셀 작성 완료');
              console.log('\n테스트 종료(WINDOW ERROR 없으면 정상)');
            });
          }, 200);
        }, 200);
      });
    }, 200);
  }, 200);
}, 200);
