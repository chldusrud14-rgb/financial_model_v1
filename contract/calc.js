#!/usr/bin/env node
/* Django 등 외부 프로세스에서 호출하는 CLI 래퍼.
   stdin 으로 입력 JSON 을 받아 stdout 으로 결과 JSON 을 낸다.

     echo '{"capacityMW":100,...}' | node calc.js
     node calc.js --xlsx out.xlsx  < input.json     (엑셀도 같이 생성)

   실패해도 프로세스는 항상 JSON 을 뱉는다: {"ok":false,"error":"..."} */
'use strict';
const path = require('path');

function readStdin() {
  return new Promise((resolve, reject) => {
    let buf = '';
    process.stdin.setEncoding('utf8');
    process.stdin.on('data', d => { buf += d; });
    process.stdin.on('end', () => resolve(buf));
    process.stdin.on('error', reject);
  });
}

(async function main() {
  try {
    const raw = await readStdin();
    if (!raw.trim()) throw new Error('입력 JSON 이 비어 있습니다 (stdin)');

    let inp;
    try { inp = JSON.parse(raw); }
    catch (e) { throw new Error('입력 JSON 파싱 실패: ' + e.message); }

    const { computeModel } = require(path.join(__dirname, 'engine2.js'));
    const result = computeModel(inp);

    const xi = process.argv.indexOf('--xlsx');
    if (xi !== -1) {
      const out = process.argv[xi + 1];
      if (!out) throw new Error('--xlsx 뒤에 출력 파일 경로가 필요합니다');
      const ExcelJS = require('exceljs');
      const { buildWorkbook } = require(path.join(__dirname, 'xlsxbuild2.js'));
      await buildWorkbook(result, ExcelJS).xlsx.writeFile(out);
      result.xlsxPath = out;
    }

    // IRR 등은 NaN 일 수 있는데 JSON.stringify 가 null 로 바꾼다.
    // 호출 측에서 "값 없음"과 구분할 수 있도록 명시적으로 표시해 둔다.
    const nanKeys = Object.keys(result.kpi || {}).filter(
      k => typeof result.kpi[k] === 'number' && !isFinite(result.kpi[k]));

    /* 조달 충족 여부 — 약정(자본금+트랜치)이 소요액(총사업비+DSRA+건설이자)보다 작으면
       못 빌린 돈이 공짜로 조달된 셈이 되어 이자가 과소계상되고 수익률이 과대평가된다.
       화면(생성기)은 이 상태를 막지만, API 로 쓰는 쪽은 이 값을 직접 확인해야 한다. */
    const need = (result.con && result.con.needTotal) || 0;
    const unfunded = (result.con && result.con.unfunded) || 0;
    const funding = { needTotalKRWm: need, unfundedKRWm: unfunded, sufficient: unfunded <= 1 };
    process.stdout.write(JSON.stringify({ ok: true, undefinedKpis: nanKeys, funding: funding, result }));
  } catch (e) {
    process.stdout.write(JSON.stringify({ ok: false, error: String(e && e.message || e) }));
    process.exitCode = 1;
  }
})();
