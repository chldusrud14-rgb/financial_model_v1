const fs = require('fs');
const rd = p => fs.readFileSync('src/' + p, 'utf8');
// ExcelJS를 CDN 대신 인라인 — 사내망(오프라인) 환경 대응
const exceljs = fs.readFileSync('node_modules/exceljs/dist/exceljs.min.js', 'utf8');
let html = rd('index.html');
html = html.replace('__EXCELJS__', () => exceljs)
           .replace('__ENGINE__', () => rd('engine.js'))
           .replace('__XLSX__', () => rd('xlsxbuild.js'))
           .replace('__APP__', () => rd('app.js'));
const out = 'dist/태양광_재무모델_생성기.html';
fs.writeFileSync(out, html);
console.log('built', out, (html.length / 1024).toFixed(1) + 'KB');
