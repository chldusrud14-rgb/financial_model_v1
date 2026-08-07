require('../src/engine.js'); require('../src/xlsxbuild.js');
const ExcelJS=require('exceljs');
const M=globalThis.SolarModel,X=globalThis.SolarXlsx;
const p={projectName:'당진 태양광발전',capacityMW:99.998,capacityFactor:15.71,degradation:0.5,auxRate:0,constructionStart:'2024-06',constructionMonths:16,operationYears:20,capexEok:1411,dsraEok:50,opexEok:49.8,opexEscal:0.7,gearing:90,rateC:5.6,rateO:5.54,graceYears:1,repayYears:16,repayType:3,payPerYear:4,tariff:154.8,tariffEscal:0,depRatio:95,depYears:20,taxMode:1,taxFlat:21,lossRate:80,decomEok:20,dsraMonths:6,minCash:10,divDSCR:1.1,divStartYear:2,discount:5.5};
const m=M.computeModel(p);
const d=(v,r)=>Math.round(v*(1+r)*100)/100;
m.sensBlocks=[
 {title:'이용률',fmtLabel:v=>v.toFixed(2)+' %',data:M.sensitivity(p,'capacityFactor',[d(p.capacityFactor,-0.1),d(p.capacityFactor,-0.05),p.capacityFactor,d(p.capacityFactor,0.05),d(p.capacityFactor,0.1)])},
 {title:'차입 금리',fmtLabel:v=>v.toFixed(2)+' %',data:M.sensitivity(p,'rateO',[d(p.rateO,-0.2),d(p.rateO,-0.1),p.rateO,d(p.rateO,0.1),d(p.rateO,0.2)])},
 {title:'매출 단가',fmtLabel:v=>v.toFixed(2)+' 원/kWh',data:M.sensitivity(p,'tariff',[d(p.tariff,-0.1),d(p.tariff,-0.05),p.tariff,d(p.tariff,0.05),d(p.tariff,0.1)])}
];
X.buildWorkbook(m,ExcelJS).xlsx.writeFile('../dist/당진_태양광발전_재무모델_생성예시.xlsx')
 .then(()=>console.log('sample written | eIRR',(m.kpi.equityIRR*100).toFixed(2),'pIRR',(m.kpi.projectIRR*100).toFixed(2),'minDSCR',m.kpi.minDSCR.toFixed(3),'sculpt',m.kpi.sculptDSCR.toFixed(3)));
