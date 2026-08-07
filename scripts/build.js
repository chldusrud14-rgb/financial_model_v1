const fs = require('fs');
const rd = p => fs.readFileSync('src/' + p, 'utf8');
let html = rd('index.html');
html = html.replace('__ENGINE__', () => rd('engine.js'))
           .replace('__XLSX__', () => rd('xlsxbuild.js'))
           .replace('__APP__', () => rd('app.js'));
const out = 'dist/태양광_재무모델_생성기.html';
fs.writeFileSync(out, html);
console.log('built', out, (html.length / 1024).toFixed(1) + 'KB');
