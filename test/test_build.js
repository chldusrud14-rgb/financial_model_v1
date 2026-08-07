const path='../src/';
require(path+'engine.js'); require(path+'xlsxbuild.js');
const ExcelJS=require('exceljs');
const M=globalThis.SolarModel, X=globalThis.SolarXlsx;
const base={projectName:'당진 태양광발전 (테스트)',capacityMW:99.998,capacityFactor:15.7083,degradation:0.5,auxRate:0,constructionStart:'2024-06',constructionMonths:16,operationYears:20,capexEok:1410.68821,dsraEok:50,opexEok:49.82,opexEscal:0.7,gearing:90,rateC:5.6,rateO:5.539,graceYears:1,repayYears:16,repayType:1,payPerYear:4,tariff:154.8,tariffEscal:0,depRatio:95,depYears:20,taxMode:1,lossRate:80,decomEok:20,discount:5.5,minCash:10,divDSCR:1.1,divStartYear:2,dsraMonths:6};
const mode = parseInt(process.argv[3]||'1',10);
const inp = Object.assign({},base,{repayType:mode});
const m=M.computeModel(inp);
m.sensBlocks=[
 {title:'이용률 민감도',fmtLabel:v=>v.toFixed(2)+' %',data:M.sensitivity(inp,'capacityFactor',[13.7,14.7,15.7083,16.7,17.7])},
 {title:'차입금리 민감도',fmtLabel:v=>v.toFixed(2)+' %',data:M.sensitivity(inp,'rateO',[4.5,5.0,5.539,6.0,6.5])},
 {title:'판매단가 민감도',fmtLabel:v=>v.toFixed(1)+' 원/kWh',data:M.sensitivity(inp,'tariff',[135,145,154.8,165,175])}
];
const wb=X.buildWorkbook(m,ExcelJS);
wb.xlsx.writeFile(process.argv[2]).then(()=>console.log('written',process.argv[2],'| eIRR',(m.kpi.equityIRR*100).toFixed(2),'pIRR',(m.kpi.projectIRR*100).toFixed(2),'minDSCR',m.kpi.minDSCR.toFixed(3)));
