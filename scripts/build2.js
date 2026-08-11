const fs = require('fs');
const rd = p => fs.readFileSync('src2/' + p, 'utf8');
// ExcelJS를 CDN 대신 인라인 — 사내망(오프라인) 환경 대응
const exceljs = fs.readFileSync('node_modules/exceljs/dist/exceljs.min.js', 'utf8');
let html = rd('index2.html');
html = html.replace('__EXCELJS__', () => exceljs)
           .replace('__ENGINE2__', () => rd('engine2.js'))
           .replace('__XLSX2__', () => rd('xlsxbuild2.js'))
           .replace('__APP2__', () => rd('app2.js'));
const out = 'dist/태양광_재무모델_생성기_v2.html';
fs.writeFileSync(out, html);
console.log('built', out, (html.length / 1024).toFixed(1) + 'KB');
