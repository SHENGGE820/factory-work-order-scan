import fs from 'fs';
const code=fs.readFileSync(new URL('../apps-script/Code.gs', import.meta.url),'utf8');
const rows=[["工令單明細","","","",""],["產品編號","產品名稱","數量","單位","工令單號"],
 ["4100000000011","測試品A","1,000","袋","V1150101001"],
 ["重工-\n4100000000042","測試品B","623","盒","重工換包裝\nX1150102004"],
 ["","","","",""]];
const props={PIN:"1234"};let cached=null;
const g={
 PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k]})},
 CacheService:{getScriptCache:()=>({get:()=>cached,put:(k,v)=>{cached=v},remove:()=>{cached=null}})},
 SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheets:()=>[{getDataRange:()=>({getDisplayValues:()=>rows})}],getSheetByName:()=>null})},
 Utilities:{sleep(){}},
 ContentService:{MimeType:{JSON:'json'},createTextOutput:t=>({setMimeType:()=>t})},
 Logger:{log:(...a)=>console.log('LOG',...a)},
};
const f=new Function(...Object.keys(g),code+';return {doGet,checkSetup};')(...Object.values(g));
for(const q of [{no:'V1150101001',pin:'1234'},{no:'x1150102004',pin:'1234'},{no:'V1150101009',pin:'1234'},{no:'V1150101001',pin:'0'},{no:'H1',pin:'1234'}]) console.log(JSON.stringify(q),f.doGet({parameter:q}));
f.checkSetup();
