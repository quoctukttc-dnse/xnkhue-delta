const ExcelJS=require('exceljs'); const C=require('../core.js');
(async()=>{
 const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(process.argv[2]);
 const d=C.parseScafSheet(wb.worksheets[0]);
 console.log(JSON.stringify(d.info), d.rows.length, d.rows[7]);
 const combos=C.listCombos(d.rows); console.log(combos);
 const mapping={}; combos.forEach(c=>mapping[c.key]={code:'4AC089-977',styleNo:'AOFYDDQ11',hs:'61082210',price:1.25});
 const groups=C.buildGroups(d.rows,mapping,process.argv[4]||'number',{});
 const H={invoiceNo:d.info.invoiceNo,invoiceDate:d.info.invoiceDate,sailingDate:new Date(Date.UTC(2026,8,21)),buyer1:d.info.buyer1,buyer2:d.info.buyer2,
  consignee1:'Pond Group Co., Ltd.',consignee2:'42F, Tower1 Parc1, 108, Yeoui-daero, Yeongdeungpo-gu',consignee3:'Seoul, Republic of Korea',consignee4:'REPUBLIC OF KOREA (07335)',
  terms:'TT, 30 days from B/L',po:'3035152',portLoading:'DA NANG PORT',destination:'INCHEON, KOREA',carrier:'SITC GUANGXI 2625N',shipperPer:'SEA',
  description:'BRIEF',season:'S24',brand:'ADIDAS',termDelivery:'FOB DANANG - VIETNAM',buyerTel:'TEL: 41 79 586 7403',cbm:1.31,nw:217.66,gw:251.46,container:'TIIU2686743',containerNote:'A PART OF 1*40HC',pcsPerSet:1};
 const out=C.buildWorkbook(ExcelJS,{header:H,groups});
 await out.xlsx.writeFile(process.argv[3]); console.log('ok',C.suggestFileName(H,groups));
})().catch(e=>{console.error(e);process.exit(1)});
