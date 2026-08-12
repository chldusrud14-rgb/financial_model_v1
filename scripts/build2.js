const fs = require('fs');
const rd = p => fs.readFileSync('src2/' + p, 'utf8');
// ExcelJS를 CDN 대신 인라인 — 사내망(오프라인) 환경 대응
const exceljs = fs.readFileSync('node_modules/exceljs/dist/exceljs.min.js', 'utf8');
// 당진 실측치(reference.json)를 통째로 인라인 — "당진 FS 불러오기" 프리셋용.
// 이게 없으면 화면 기본값(범용 근사)으로만 계산돼서 검증된 숫자와 달라진다.
const refJson = fs.readFileSync('reference/dangjin_reference.json', 'utf8');
let html = rd('index2.html');
html = html.replace('__EXCELJS__', () => exceljs)
           .replace('__DANGJIN_REF__', () => 'window.__DANGJIN_REFERENCE__ = ' + refJson + ';')
           .replace('__ENGINE2__', () => rd('engine2.js'))
           .replace('__XLSX2__', () => rd('xlsxbuild2.js'))
           .replace('__APP2__', () => rd('app2.js'));
const out = 'dist/태양광_재무모델_생성기_v2.html';
fs.writeFileSync(out, html);
console.log('built', out, (html.length / 1024).toFixed(1) + 'KB');
