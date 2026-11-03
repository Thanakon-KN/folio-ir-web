'use strict';

// Deterministic, browser-local update functions. These rules deliberately stop
// on ambiguous mappings instead of inventing values.
const FolioEngine = (() => {
  const periodFrom = value => {
    const text = String(value ?? '').replace(/<[^>]*>/g, ' ').replace(/&nbsp;|&#160;/gi, ' ').replace(/\s+/g, ' ').trim();
    const q = text.match(/^(?:([1-4])\s*Q|Q\s*([1-4])|(?:Quarter|ไตรมาส)\s*([1-4]))\s*[/ .-]?\s*((?:19|20|25|26)\d{2}|\d{2})(?:\b|$)/i);
    const duration=text.match(/^(3|6|9)\s*M\s*[/ .-]?\s*((?:19|20|25|26)\d{2}|\d{2})(?:\b|$)/i);
    const half=text.match(/^1\s*H\s*[/ .-]?\s*((?:19|20|25|26)\d{2}|\d{2})(?:\b|$)/i);
    let year, quarter = 0, months=0, halfYear=false;
    if (q) {
      year = Number(q[4]); quarter = Number(q[1] || q[2] || q[3]);
      if (year < 100) year += year >= 50 ? 2500 : 2000;
    } else if(duration||half){
      year=Number(duration?.[2]||half?.[1]);months=duration?Number(duration[1]):6;halfYear=Boolean(half);
      if(year<100)year+=year>=50?2500:2000;
    } else {
      const annual = text.match(/^(?:(?:FY|ปี)\s*)?((?:19|20|25|26)\d{2})(?:\s*[AE])?$/i);
      if (!annual) return null;
      year = Number(annual[1]);
    }
    if (year >= 2400) year -= 543;
    if (year < 1900 || year > 2199) return null;
    const id=String(year)+(quarter?'-Q'+quarter:halfYear?'-1H':months?'-'+months+'M':'');
    return {id,year,quarter,months,halfYear,label:formatPeriodLabel({year,quarter,months,halfYear},'en')};
  };
  const formatPeriodLabel = (period,language='en',buddhist=false) => {const year=period.year+(buddhist?543:0);if(period.quarter)return (language==='th'?'ไตรมาส ':'Q')+period.quarter+'/'+year;if(period.halfYear)return '1H/'+year;if(period.months)return period.months+'M/'+year;return String(year);};
  const isSubAnnual = period => Boolean(period&&(period.quarter||period.months||period.halfYear));
  const periodOrder = period => period.quarter?period.quarter*3:period.months||(period.halfYear?6:13);
  const uniquePeriods = list => [...new Map(list.filter(Boolean).map(p => [p.id, p])).values()].sort((a,b)=>a.year-b.year||periodOrder(a)-periodOrder(b)||String(a.id).localeCompare(String(b.id)));
  const normalize = value => plainText(value).toLowerCase().normalize('NFKC').replace(/&amp;/g,'&').replace(/[^\p{L}\p{N}]+/gu,'');
  const plainText = value => {
    const box=document.createElement('div'); box.innerHTML=String(value??'');
    return (box.textContent||'').replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim();
  };
  const numberFrom = value => {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    let text=String(value??'').trim(); if(!text||/^[-–—]$/.test(text)) return null;
    const negative=/^\(.*\)$/.test(text); text=text.replace(/[(),%\s]/g,'').replace(/,/g,'');
    const n=Number(text); return Number.isFinite(n)?(negative?-Math.abs(n):n):null;
  };
  const formatLike = (oldValue, value) => {
    if (value === null || value === undefined) return String(oldValue ?? '');
    const old=String(oldValue??'').trim();
    if (/\(/.test(old)&&/\)/.test(old)&&value<0) return '('+Math.abs(value).toLocaleString('en-US',{minimumFractionDigits:decimalPlaces(old),maximumFractionDigits:decimalPlaces(old)})+')';
    const decimals=decimalPlaces(old), grouped=/\d,\d{3}/.test(old), pct=/%/.test(old);
    let result=Math.abs(value).toLocaleString('en-US',{useGrouping:grouped,minimumFractionDigits:decimals,maximumFractionDigits:decimals});
    if(!decimals && !old.includes('.')) result=String(Math.round(value));
    if(value<0) result='-'+result;
    return result+(pct?'%':'');
  };
  const decimalPlaces = value => { const match=String(value??'').replace(/,/g,'').match(/\.(\d+)/); return match?match[1].length:0; };
  const parseWorkbook = async files => {
    if (!window.XLSX) throw new Error('โหลดตัวอ่าน Excel ไม่สำเร็จ');
    if (!files.length) throw new Error('เลือกไฟล์ข้อมูลต้นทางก่อน');
    const books=[];
    for(const file of files){
      const buffer=await file.arrayBuffer();
      const wb=XLSX.read(buffer,{type:'array',cellDates:false,cellNF:true,cellText:true});
      const sheets=wb.SheetNames.map(name=>{
        const ws=wb.Sheets[name]; if(!ws['!ref']) return {name,rows:[]};
          const range=XLSX.utils.decode_range(ws['!ref']), rows=[];
        for(let r=range.s.r;r<=range.e.r;r++){
          const cells=[],grid=[];
          for(let c=range.s.c;c<=range.e.c;c++){
            const address=XLSX.utils.encode_cell({r,c}), cell=ws[address];
            const raw=cell?.v ?? '', text=cell?.w ?? (raw instanceof Date?raw.toISOString().slice(0,10):String(raw));
            const value={address,row:r+1,col:c,text:String(text).trim(),raw,format:cell?.z||'',present:Boolean(cell)};grid.push(value);
            if(String(text).trim()||String(raw).trim()) cells.push(value);
          }
          if(cells.length) rows.push({row:r+1,cells,grid});
        }
        return {name,rows};
      });
      books.push({name:file.name,sheets});
    }
    return {books,cellCount:books.reduce((n,b)=>n+b.sheets.reduce((m,s)=>m+s.rows.reduce((k,r)=>k+r.cells.length,0),0),0)};
  };
  const extractFinancial = workbook => {
    const metrics=[], sourcePeriods=[];
    for(const book of workbook.books) for(const sheet of book.sheets){
      const all=sheet.rows;
      for(let ri=0;ri<all.length;ri++){
        const row=all[ri];
        const quarterOnly=row.cells.map(cell=>({cell,match:cell.text.match(/^(?:Q\s*([1-4])|([1-4])\s*Q|Quarter\s*([1-4])|ไตรมาส\s*([1-4]))$/i)})).filter(x=>x.match);
        const durationOnly=row.cells.filter(cell=>/^(?:(?:3|6|9)\s*M|1\s*H)$/i.test(cell.text.trim()));
        const partialPeriods=[...quarterOnly.map(x=>x.cell),...durationOnly];
        const direct=row.cells.map(cell=>({cell,period:periodFrom(cell.text)})).filter(x=>x.period);
        const hasNumericBelow=col=>{for(let i=ri+1;i<Math.min(all.length,ri+31);i++){const next=all[i],nextPeriods=next.cells.filter(cell=>periodFrom(cell.text)),nextPartial=next.cells.filter(cell=>/^(?:Q\s*[1-4]|[1-4]\s*Q|Quarter\s*[1-4]|ไตรมาส\s*[1-4]|(?:3|6|9)\s*M|1\s*H)$/i.test(cell.text));if(nextPeriods.length>=2||nextPartial.length>=2)break;const grid=next.grid||next.cells,cell=grid.find(item=>item.col===col);if(cell&&(numberFrom(cell.raw)!==null||numberFrom(cell.text)!==null||/^[-–—]$/.test(cell.text)))return true;if(grid.some(item=>item.col!==col&&(numberFrom(item.raw)!==null||numberFrom(item.text)!==null)))return true;}return false;};
        const periodsByCol=[];
        if(partialPeriods.length){
          for(const item of partialPeriods){
            let yearCell=direct.filter(x=>!isSubAnnual(x.period)&&x.cell.col<=item.col).sort((a,b)=>b.cell.col-a.cell.col)[0];
            if(!yearCell){for(let previous=ri-1;previous>=Math.max(0,ri-3)&&!yearCell;previous--){const yearHeaders=all[previous].cells.map(cell=>({cell,period:periodFrom(cell.text)})).filter(x=>x.period&&!isSubAnnual(x.period)&&x.cell.col<=item.col);yearCell=yearHeaders.sort((a,b)=>b.cell.col-a.cell.col)[0];}}
            if(!yearCell)continue;const period=periodFrom(`${item.text}/${yearCell.period.year}`);if(period&&hasNumericBelow(item.col))periodsByCol.push({period,col:item.col});
          }
        }
        for(const item of direct){const groupingHeader=partialPeriods.some(partial=>partial.col===item.cell.col);if((!partialPeriods.length||!groupingHeader)&&hasNumericBelow(item.cell.col))periodsByCol.push({period:item.period,col:item.cell.col});}
        const candidates=[...new Map(periodsByCol.map(x=>[x.col+'|'+x.period.id,x])).values()];
        if(!candidates.length)continue;
        const uniqueHeaderPeriods=uniquePeriods(candidates.map(x=>x.period));
        if(uniqueHeaderPeriods.length<2&&!candidates.some(x=>isSubAnnual(x.period))&&!(row.cells.length<=2&&row.cells.some(cell=>/^(?:year|period|ปี|งวด|FY)$/i.test(cell.text.trim()))))continue;
        sourcePeriods.push(...candidates.map(x=>x.period));
        const firstPeriodCol=Math.min(...candidates.map(x=>x.col));
        const sectionContext=all.slice(Math.max(0,ri-4),ri+1).map(item=>item.cells.map(cell=>cell.text).join(' ')).join(' ');
        const sheetLabel=(book.name+' / '+sheet.name).toLowerCase();
        const explicitDetail=/consolidated financial statements?|งบการเงินรวม|detailed financial|รายละเอียดงบการเงิน/i.test(sectionContext);
        const summaryBlock=/financial highlights|highlights summary|graph|chart|summary|home|overview|สรุป|กราฟ/i.test(sheetLabel+' '+sectionContext);
        const kind=explicitDetail?'detail':summaryBlock?'summary':'detail';
        const periods=candidates.map(x=>({period:x.period,col:x.col}));
        const comparisonColumns=row.cells.map(cell=>({cell,type:comparisonType(cell.text)})).filter(item=>item.type&&!periodFrom(item.cell.text)).map(item=>{const owner=periods.filter(period=>period.col<=item.cell.col).sort((a,b)=>b.col-a.col)[0]||[...periods].sort((a,b)=>Math.abs(a.col-item.cell.col)-Math.abs(b.col-item.cell.col))[0];return owner?{type:item.type,col:item.cell.col,period:owner.period}:null;}).filter(Boolean);
        const headerLabels=[];
        if(firstPeriodCol===0){
          for(let previous=ri-1;previous>=0&&row.row-all[previous].row<=2;previous--){
            const prior=all[previous];if(prior.cells.some(cell=>periodFrom(cell.text)))break;
            const labels=prior.cells.filter(cell=>numberFrom(cell.raw)===null&&numberFrom(cell.text)===null&&!/^(?:ปี|งวด)$/i.test(cell.text.trim()));
            if(!labels.length)break;headerLabels.unshift(...labels);
          }
        }
        let looseHeaderLabelsAvailable=headerLabels.length>0;
        for(let rj=ri+1;rj<all.length;rj++){
          const data=all[rj];
          if(data.row>all[rj-1].row+20) break;
          const nextPeriods=data.cells.filter(cell=>periodFrom(cell.text)),nextPartial=data.cells.filter(cell=>/^(?:Q\s*[1-4]|[1-4]\s*Q|Quarter\s*[1-4]|ไตรมาส\s*[1-4]|(?:3|6|9)\s*M|1\s*H)$/i.test(cell.text));
          if(nextPeriods.length>=2||nextPartial.length>=2)break;
          const before=data.cells.filter(cell=>cell.col<firstPeriodCol && cell.text.trim());
          const labelCells=before.length?before:(looseHeaderLabelsAvailable?headerLabels:[]);
          if(!labelCells.length) continue;
          // Workbooks may put English and Thai labels in adjacent columns.
          // Keep every label for matching, and use the first label as the
          // display name so HTML can match either language.
          const labels=[...new Set(labelCells.map(cell=>cell.text.trim()).filter(Boolean))];
          const label=labels[0];
          if(periodFrom(label)) break;
          const values={},comparisons={}; let found=false;
          for(const item of periods){
            const cell=(data.grid||data.cells).find(x=>x.col===item.col);
            let val=numberFrom(cell?.raw);if(val===null)val=numberFrom(cell?.text);
            if(val!==null){values[item.period.id]={value:val,text:cell?.text||String(val),address:cell?.address||''};found=true;}
            else if(cell&&(cell.present||/^[-–—]$/.test(cell.text))){values[item.period.id]={value:null,text:cell.text,address:cell.address||'',empty:true};if(cell.present||cell.text)found=true;}
          }
          for(const comparison of comparisonColumns){const cell=(data.grid||data.cells).find(x=>x.col===comparison.col),key=comparison.period.id+':'+comparison.type;let val=numberFrom(cell?.raw);if(val===null)val=numberFrom(cell?.text);if(val!==null){comparisons[key]={value:val,text:cell?.text||String(val),address:cell?.address||''};found=true;}else if(cell&&(cell.present||/^[-–—]$/.test(cell.text))){comparisons[key]={value:null,text:cell.text,address:cell.address||'',empty:true};if(cell.present||cell.text)found=true;}}
          if(found)for(const item of periods)if(!values[item.period.id])values[item.period.id]={value:null,text:'',empty:true};
          if(found && normalize(label)) {
            metrics.push({label,labels,normalized:normalize(label),kind,bookName:book.name,sheetName:sheet.name,sheet:book.name+' / '+sheet.name,row:data.row,values,comparisons});
            if(!before.length)looseHeaderLabelsAvailable=false;
          }
        }
      }
    }
    // Several workbooks repeat the same period header for different blocks.
    // Merge those blocks by source row instead of letting duplicates make an
    // otherwise exact label look ambiguous.
    const bySourceRow=new Map();
    for(const metric of metrics){
      const key=[metric.normalized,metric.sheetName,metric.kind,metric.bookName,metric.row].join('|');
      const prior=bySourceRow.get(key);if(prior){Object.assign(prior.values,metric.values);Object.assign(prior.comparisons,metric.comparisons||{});}else bySourceRow.set(key,{...metric,values:{...metric.values},comparisons:{...(metric.comparisons||{})}});
    }
    const bySection=new Map();for(const metric of bySourceRow.values()){const key=[metric.normalized,metric.sheetName,metric.kind].join('|');if(!bySection.has(key))bySection.set(key,[]);bySection.get(key).push(metric);}
    const consolidated=[];
    for(const group of bySection.values()){
      let combined=null;
      for(const metric of group){
        const conflict=combined&&(metric.bookName===combined.bookName||Object.entries(metric.values).some(([id,item])=>combined.values[id]&&combined.values[id].value!==item.value)||Object.entries(metric.comparisons||{}).some(([id,item])=>combined.comparisons?.[id]&&combined.comparisons[id].value!==item.value));
        if(conflict){consolidated.push(...group);combined=null;break;}
        if(!combined)combined={...metric,values:{...metric.values},comparisons:{...(metric.comparisons||{})}};else{Object.assign(combined.values,metric.values);Object.assign(combined.comparisons,metric.comparisons||{});}
      }
      if(combined)consolidated.push(combined);
    }
    return {periods:uniquePeriods(sourcePeriods),metrics:consolidated};
  };
  const extractExistingPeriods = payload => {
    const found=[];
    for(const m of String(payload.html||'').replace(/<!--[\s\S]*?-->/g,'').matchAll(/<(?:th|td)\b[^>]*>([\s\S]*?)<\/(?:th|td)>/gi)) found.push(periodFrom(m[1]));
    const js=String(payload.javascript||'').replace(/\/\*[\s\S]*?\*\/|^\s*\/\/[^\r\n]*/gm,'');
    for(const m of js.matchAll(/(?:["']?period["']?)\s*:\s*\[([^\]]*)\]/g)) for(const p of m[1].matchAll(/(['"])([^'"]*)\1|\b(?:19|20|25|26)\d{2}\b/g)) found.push(periodFrom(p[2]??p[0]));
    return uniquePeriods(found);
  };
  const inspect = (workbook,payload,workflow) => {
    const files=workbook.books.map(book=>({name:book.name,sheets:book.sheets.map(sheet=>({name:sheet.name,rows:sheet.rows.length,cells:sheet.rows.reduce((n,r)=>n+r.cells.length,0),preview:sheet.rows.slice(0,7).map(row=>({row:row.row,cells:row.cells.slice(0,8).map(c=>({address:c.address,text:c.text.slice(0,140)}))}))}))}));
    if(workflow==='financial-highlights'){
      const parsed=extractFinancial(workbook), existing=extractExistingPeriods(payload), sourceKeys=new Set(parsed.periods.map(p=>p.id)), oldKeys=new Set(existing.map(p=>p.id));
      const displayPeriods=uniquePeriods([...parsed.periods,...existing]).map(p=>({...p,inSource:sourceKeys.has(p.id),inExisting:oldKeys.has(p.id)}));
      const warnings=[];if(!parsed.periods.length)warnings.push('ยังไม่พบหัวงวดที่ระบบรู้จักในแนวนอน กรุณาตรวจไฟล์หรือเลือก workbook ที่มีแถวหัวงวด');if(!existing.length)warnings.push('ยังไม่พบ period ในโค้ดเดิม');if(parsed.periods.length&&!parsed.metrics.length)warnings.push('พบงวด แต่ยังจับคู่แถวตัวเลขไม่ชัดเจน');
      return {files,cellCount:workbook.cellCount,sourcePeriods:parsed.periods,existingPeriods:existing,displayPeriods,metrics:parsed.metrics,warnings};
    }
    const blocks=[
      {id:'overview',label:'Overview · ภาพรวมผู้ถือหุ้น',re:/overview|rights?\s*type|total\s*(?:number\s*of\s*)?shareholders|scripless|ภาพรวม|ประเภทสิทธิ|ผู้ถือหุ้นทั้งหมด|ระบบไร้ใบหุ้น/i},
      {id:'free-float',label:'Free float · ผู้ถือหุ้นรายย่อย',re:/free\s*float|minor\s*shareholders|ผู้ถือหุ้นรายย่อย/i},
      {id:'graph',label:'Graph · ข้อมูลกราฟ',re:/\bgraph\b|\bchart\b|majorShareholders|กราฟ|แผนภูมิ/i},
      {id:'list',label:'Shareholder list · รายชื่อผู้ถือหุ้น',re:/shareholder\s*(?:list|name)|major\s*shareholders|รายชื่อผู้ถือหุ้น|ผู้ถือหุ้นรายใหญ่|ชื่อผู้ถือหุ้น/i}
    ].map(block=>{const refs=[];for(const book of workbook.books)for(const sheet of book.sheets){if(block.re.test(sheet.name))refs.push(book.name+' / '+sheet.name);for(const row of sheet.rows)for(const cell of row.cells)if(block.re.test(cell.text))refs.push(book.name+' / '+sheet.name+' / '+cell.address);}return {...block,references:[...new Set(refs)].slice(0,12)};});
    const warnings=blocks.filter(b=>!b.references.length).map(b=>'ยังไม่พบหัวข้อ '+b.label+' ใน workbook');
    const shareholderRows=extractShareholderRows(workbook);
    const shareholderFacts=extractShareholderFacts(workbook);
    return {files,cellCount:workbook.cellCount,displayPeriods:[],sourcePeriods:[],existingPeriods:[],blocks:blocks.map(({re,...b})=>b),shareholderRows,shareholderFacts,warnings};
  };
  function extractShareholderFacts(workbook){
    const rules=[
      ['rightsType',/^(rights?type|typeofrights|ประเภทสิทธิ|สิทธิประโยชน์)$/i,'overview'],
      ['overviewDate',/^(overviewdate|asofdate|วันที่ข้อมูล|ณวันที่)$/i,'overview'],
      ['totalShareholders',/^(totalshareholders|numberofshareholders|ผู้ถือหุ้นทั้งหมด|จำนวนผู้ถือหุ้น)$/i,'overview'],
      ['scripless',/^(scripless|scriplessholding|ระบบไร้ใบหุ้น|ถือครองผ่านระบบไร้ใบหุ้น)$/i,'overview'],
      ['freeFloatDate',/^(freefloatdate|minoritydate|วันที่ผู้ถือหุ้นรายย่อย)$/i,'free-float'],
      ['minorShareholders',/^(minorshareholders|numberofminorshareholders|ผู้ถือหุ้นรายย่อย|จำนวนผู้ถือหุ้นรายย่อย)$/i,'free-float'],
      ['freeFloatPercent',/^(freefloat|freefloatpercent|freefloatpercentage|สัดส่วนผู้ถือหุ้นรายย่อย|ร้อยละผู้ถือหุ้นรายย่อย)$/i,'free-float']
    ];
    const facts={};for(const book of workbook.books)for(const sheet of book.sheets)for(const row of sheet.rows)for(const cell of row.cells){
      const key=normalize(cell.text);for(const [id,re,section] of rules){if(!re.test(key))continue;const value=row.cells.find(c=>c.col===cell.col+1);if(!value||!value.text)continue;const candidate={value:value.text,raw:value.raw,source:`${book.name} / ${sheet.name} / ${cell.address}:${value.address}`,sheet:sheet.name};const priority=section==='free-float'?/free\s*float|minor|รายย่อย/i.test(sheet.name):/overview|ภาพรวม/i.test(sheet.name);if(!facts[id]||priority)facts[id]={...candidate,section};}
    }return facts;
  }
  function extractShareholderRows(workbook){
    const rows=[];
    for(const book of workbook.books)for(const sheet of book.sheets){
      for(let i=0;i<sheet.rows.length;i++){
        const header=sheet.rows[i], labels=header.cells.map(c=>normalize(c.text));
        const nameCol=labels.findIndex(x=>/^(shareholder|shareholders|shareholdername|nameofshareholder|name|majorholder|ชื่อผู้ถือหุ้น|ผู้ถือหุ้น)$/.test(x)||/^shareholdername(?:english|thai|en|th)$/.test(x));
        const shareCol=labels.findIndex(x=>/^(shares|numberofshares|noofshares|shareholding|จำนวนหุ้น|หุ้น)$/.test(x)||/^shares(?:held|owned)$/.test(x));
        const pctCol=labels.findIndex(x=>/^(percent|percentage|shareholdingpercent|ownership|สัดส่วน|ร้อยละ|%)$/.test(x)||/(?:percent|percentage)$/.test(x));
        const graphSheet=/graph|chart|กราฟ/i.test(sheet.name);
        if(nameCol<0||(shareCol<0&&!(graphSheet&&pctCol>=0)))continue;
        for(const row of sheet.rows.slice(i+1)){const name=row.cells.find(c=>c.col===header.cells[nameCol].col)?.text||'';const share=shareCol>=0?row.cells.find(c=>c.col===header.cells[shareCol].col):null;const pct=pctCol>=0?row.cells.find(c=>c.col===header.cells[pctCol].col):null;const shareRaw=numberFrom(share?.raw),percentRaw=numberFrom(pct?.raw);const shares=shareRaw!==null?shareRaw:numberFrom(share?.text),percent=percentRaw!==null?percentRaw:numberFrom(pct?.text);if(name&&(shares!==null||percent!==null))rows.push({name,shares,percent,sheet:sheet.name,source:`${book.name} / ${sheet.name} / row ${row.row}`});}
      }
    }
    return rows;
  }
  const validateSelection=(selection,summary)=>{
    if(!Array.isArray(selection)||!selection.length)throw new Error('เลือกงวดที่ต้องการแสดงอย่างน้อย 1 งวด');
    const byId=new Map(summary.displayPeriods.map(period=>[period.id,period]));if(selection.some(id=>!byId.has(id))||new Set(selection).size!==selection.length)throw new Error('ตัวเลือกงวดไม่ตรงกับผลตรวจ กรุณาตรวจไฟล์ใหม่');return selection.map(id=>byId.get(id));
  };
  const aliases={
    revenue:['sales','netsales','turnover','totalrevenue','totalrevenues','revenuefromoperations','revenuefromcontractswithcustomers','salesandserviceincome','รายได้','รายได้รวม','รายได้จากการขายและบริการ','รายได้จากการขายสินค้าและการให้บริการ'],
    costofsales:['costofgoodsold','costofsalesandservices','ต้นทุนขาย','ต้นทุนขายและบริการ','ต้นทุนสินค้าและบริการ'],
    netprofit:['netincome','profitforperiod','profitfortheyear','netprofitfortheperiod','profitattributabletoownersoftheparent','netprofitattributabletoownersoftheparent','กำไรสุทธิ','กำไรขาดทุนสุทธิ','กำไร(ขาดทุน)สุทธิ','กำไรสุทธิส่วนที่เป็นของผู้ถือหุ้นบริษัทใหญ่'],
    grossprofit:['กำไรขั้นต้น','กำไร(ขาดทุน)ขั้นต้น'], operatingprofit:['operatingincome','profitfromoperations','กำไรจากการดำเนินงาน','กำไรจากการปฏิบัติงาน','กำไร(ขาดทุน)จากการดำเนินงาน'],
    profitbeforetax:['ebt','profitbeforeincometax','profitbeforetaxation','กำไรก่อนภาษี','กำไรก่อนภาษีเงินได้','กำไร(ขาดทุน)ก่อนภาษี'],
    sellingandadministrativeexpenses:['sga','sellingandadministrativeexpenses','ค่าใช้จ่ายในการขายและบริหาร','ค่าใช้จ่ายในการขายและการบริหาร'],
    totalassets:['assets','สินทรัพย์รวม'], totalliabilities:['liabilities','หนี้สินรวม'],
    totalshareholdersequity:['shareholdersequity','equity','ส่วนของผู้ถือหุ้น','ส่วนของเจ้าของ'],
    earningspershare:['eps','กำไรต่อหุ้น','กำไรสุทธิต่อหุ้น'],
    ebitda:['กำไรก่อนดอกเบี้ยภาษีค่าเสื่อมราคาและค่าตัดจำหน่าย'],
    cashflowfromoperatingactivities:['operatingcashflow','กระแสเงินสดจากกิจกรรมดำเนินงาน']
  };
  const canonicalLabel=value=>{
    const key=normalize(value);
    for(const [canonical,variants] of Object.entries(aliases))if(key===canonical||variants.some(item=>normalize(item)===key))return canonical;
    return key;
  };
  const comparisonType=value=>{const text=plainText(value).replace(/[_-]+/g,' ');if(/\b(?:y\s*o\s*y|yoy|year\s*over\s*year)\b|ปีก่อน/i.test(text))return 'yoy';if(/\b(?:q\s*o\s*q|qoq|quarter\s*over\s*quarter)\b|ไตรมาสก่อน/i.test(text))return 'qoq';return '';};
  const comparisonBase=value=>plainText(value).replace(/\b(?:y\s*o\s*y|yoy|year\s*over\s*year|q\s*o\s*q|qoq|quarter\s*over\s*quarter|growth|change|comparison)\b|ปีก่อน|ไตรมาสก่อน/gi,' ').replace(/[_-]+/g,' ').trim();
  const labelSimilarity=(left,right)=>{
    const leftComparison=comparisonType(left),rightComparison=comparisonType(right);
    if(leftComparison||rightComparison){if(leftComparison!==rightComparison)return 0;const a=normalize(comparisonBase(left)),b=normalize(comparisonBase(right));if(a===b)return 1;if(!a||!b)return 0;const at=new Set(a.match(/[\p{L}\p{N}]+/gu)||[]),bt=new Set(b.match(/[\p{L}\p{N}]+/gu)||[]);let overlap=0;for(const token of at)if(bt.has(token))overlap++;return overlap/(at.size+bt.size-overlap);}
    const a=canonicalLabel(left),b=canonicalLabel(right);if(!a||!b)return 0;if(a===b)return 1;
    if(a.length>3&&b.length>3&&(a.includes(b)||b.includes(a)))return Math.min(a.length,b.length)/Math.max(a.length,b.length)*0.96;
    const at=new Set((plainText(left).toLowerCase().match(/[\p{L}\p{N}]+/gu)||[]).filter(x=>!['total','million','baht','บาท','summary','detailed','the','of'].includes(x)));
    const bt=new Set((plainText(right).toLowerCase().match(/[\p{L}\p{N}]+/gu)||[]).filter(x=>!['total','million','baht','บาท','summary','detailed','the','of'].includes(x)));
    if(!at.size||!bt.size)return 0;let overlap=0;for(const token of at)if(bt.has(token))overlap++;
    return overlap/(at.size+bt.size-overlap);
  };
  const findMetric=(metrics,label,kind='',periods=[])=>{
    const ranked=metrics.map(metric=>({metric,score:Math.max(...(metric.labels?.length?metric.labels:[metric.label]).map(candidate=>labelSimilarity(label,candidate))),coverage:periods.filter(period=>metric.values[period.id]).length})).filter(x=>x.score>=0.55);
    if(!ranked.length)return null;
    ranked.sort((a,b)=>b.score-a.score||b.coverage-a.coverage||(kind&&a.metric.kind===kind?-1:0)-(kind&&b.metric.kind===kind?-1:0));
    const top=ranked[0],sameScore=ranked.filter(x=>Math.abs(x.score-top.score)<0.035&&x.coverage===top.coverage);
    if(sameScore.length>1&&kind){const preferred=sameScore.filter(x=>x.metric.kind===kind);if(preferred.length===1)return preferred[0].metric;}
    if(sameScore.length===1)return top.metric;
    const signatures=new Set(sameScore.map(x=>periods.map(p=>x.metric.values[p.id]?.value??'').join('|')));
    return signatures.size===1?top.metric:null;
  };
  function scanRows(html){
    const out=[],occupied=new Map(),re=/<tr\b[^>]*>[\s\S]*?<\/tr>/gi;let m;
    while((m=re.exec(html))){const raw=m[0],cells=[],carried=new Set(occupied.keys()),cellRe=/<(td|th)\b([^>]*)>([\s\S]*?)<\/\1>/gi;let c,column=0;
      while((c=cellRe.exec(raw))){const span=Math.max(1,Number(c[2].match(/\bcolspan\s*=\s*["']?(\d+)/i)?.[1]||1));while(Array.from({length:span},(_,i)=>occupied.has(column+i)).some(Boolean))column++;
        const rowSpan=Math.max(1,Number(c[2].match(/\browspan\s*=\s*["']?(\d+)/i)?.[1]||1));cells.push({tag:c[1],attrs:c[2],inner:c[3],raw:c[0],start:c.index,end:cellRe.lastIndex,text:plainText(c[3]),column,span,endColumn:column+span,rowSpan});
        if(rowSpan>1)for(let col=column;col<column+span;col++)occupied.set(col,rowSpan-1);column+=span;}
      for(const col of carried){const remaining=occupied.get(col);if(remaining<=1)occupied.delete(col);else occupied.set(col,remaining-1);}
      if(cells.length)out.push({raw,start:m.index,end:re.lastIndex,cells});}
    return out;
  }
  const cellAtColumn=(row,column)=>row.cells.find(cell=>column>=cell.column&&column<cell.endColumn)||null;
  const replaceCellText=(raw,oldText,newText)=>{
    let next=raw;
    // Vue numeric components commonly store a value in :number or number attributes.
    if(/(?:\:?(?:number|value|count|percent|percentage)|\bnumber)\s*=\s*(["'])/i.test(next)){
      const re=/((?:\:?(?:number|value|count|percent|percentage)|\bnumber)\s*=\s*)(["'])([\s\S]*?)\2/i;return next.replace(re,(_,prefix,q,value)=>prefix+q+(/-?\d[\d,.]*/.test(value)?value.replace(/-?\d[\d,.]*/,newText):newText)+q);
    }
    const content=next.match(/^(<[^>]+>)([\s\S]*)(<\/[^>]+>)$/);if(!content)return raw;
    const inner=content[2];
    if(/<[^>]+>/.test(inner)){
      const textNode=inner.match(/([^<>]*-?\d[\d,.]*%?[^<>]*)/);
      if(!textNode)return raw;
      const updated=textNode[0].replace(/-?\d[\d,.]*%?/,newText);return content[1]+inner.replace(textNode[0],updated)+content[3];
    }
    return content[1]+newText+content[3];
  };
  function formatFinancialValue(previous,value,label){
    const prior=String(previous??'').trim();
    const precise=/(?:eps|per share|margin|ratio|percent|percentage|yield|roe|roa|งบต่อหุ้น|กำไรต่อหุ้น|อัตราส่วน|ร้อยละ|เปอร์เซ็นต์)/i.test(label);
    const decimals=decimalPlaces(prior);
    if(decimals>0)return formatLike(prior,value);
    const places=precise?2:0;
    const grouped=/\d,\d{3}/.test(prior);
    const sign=value<0?'-':'';
    const result=Math.abs(value).toLocaleString('en-US',{useGrouping:grouped,minimumFractionDigits:places,maximumFractionDigits:places});
    return sign+result+(/%/.test(prior)?'%':'');
  }
  function emptyFinancialValue(source,previous){
    const sourceText=String(source?.text??'').trim();if(/^[-–—]$/.test(sourceText))return sourceText;
    const original=String(previous??''),prior=plainText(original).trim();
    if(!prior||/^[-–—]$/.test(prior)||numberFrom(prior)===0)return original;
    return '';
  }
  function financialValueCell(oldCell,templateCell,source,label){
    const cell=oldCell||templateCell;if(!cell)return '<td></td>';
    const previous=oldCell?.text??templateCell?.text??'';
    const value=source?(source.value===null?emptyFinancialValue(source,previous):formatFinancialValue(previous,source.value,label)):oldCell?oldCell.text:emptyFinancialValue({text:''},previous);
    return replaceCellText(cell.raw,cell.text,value);
  }
  function emptyArrayValue(source,previous,fallback){
    const sourceText=String(source?.text??'').trim();if(/^[-–—]$/.test(sourceText))return JSON.stringify(sourceText);
    const preserve=value=>{if(value===null)return 'null';if(value===undefined)return null;if(typeof value==='number')return value===0?'0':null;const text=String(value).trim();if(!text)return '""';if(/^[-–—]$/.test(text)||numberFrom(text)===0)return JSON.stringify(value);return null;};
    return preserve(previous)??preserve(fallback)??'null';
  }
  function replaceRowCells(row,newCells){
    let raw=row.raw;for(let i=row.cells.length-1;i>=0;i--){const c=row.cells[i];raw=raw.slice(0,c.start)+raw.slice(c.end);}const close=raw.lastIndexOf('</tr>');return close<0?row.raw:raw.slice(0,close)+newCells.join('')+raw.slice(close);
  }
  function updateFinancialTitle(html,selected,report){
    const latest=[...selected].sort((a,b)=>a.year-b.year||periodOrder(a)-periodOrder(b)||String(a.id).localeCompare(String(b.id))).at(-1);if(!latest)return html;
    return html.replace(/<(h[1-6])\b([^>]*)>([\s\S]*?)<\/\1>/gi,(whole,tag,attrs,inner)=>{
      const title=plainText(inner);if(!/(?:financial\s+highlights?|ข้อมูลสำคัญทางการเงิน|สรุปข้อมูลทางการเงิน)/i.test(title))return whole;
      const thai=/[\u0E00-\u0E7F]/.test(title),language=thai?'th':'en';
      const periodPattern=/(?:Q\s*[1-4]|[1-4]\s*Q|Quarter\s*[1-4]|ไตรมาส\s*[1-4]|[369]\s*M|1\s*H)\s*[/. -]?\s*((?:19|20|25|26)\d{2})/gi;
      const existingPeriod=[...title.matchAll(periodPattern)].at(-1);
      const existingYear=[...title.matchAll(/\b((?:19|20|25|26)\d{2})\b/g)].at(-1);
      const oldToken=existingPeriod?.[0]||existingYear?.[1];if(!oldToken)return whole;
      const oldYear=Number(existingPeriod?.[1]||existingYear[1]),label=formatPeriodLabel(latest,language,oldYear>=2400);
      if(oldToken===label)return whole;
      const offset=inner.lastIndexOf(oldToken);if(offset<0)return whole;
      const nextInner=inner.slice(0,offset)+label+inner.slice(offset+oldToken.length);
      report.changes.push({item:title,oldValue:oldToken,newValue:label,source:'งวดล่าสุดที่เลือก · '+formatPeriodLabel(latest,language,oldYear>=2400)});
      return '<'+tag+attrs+'>'+nextInner+'</'+tag+'>';
    });
  }
  function updateFinancialHtml(html,summary,selected,mode,language,report){
    const tables=[...html.matchAll(/<table\b[\s\S]*?<\/table>/gi)];let changed=false,matched=false;
    let updated=html.replace(/<table\b[\s\S]*?<\/table>/gi,table=>{
      const rows=scanRows(table);let headIndex=-1,periodCols=[];
      rows.forEach((row,index)=>{const ps=row.cells.map(c=>({period:periodFrom(c.text),column:c.column,span:c.span})).filter(x=>x.period);if(ps.length>periodCols.length&&ps.length>=1){periodCols=ps;headIndex=index;}});
      if(headIndex<0||!periodCols.length)return table;
      if(periodCols.some(period=>period.span!==1)){report.unmatched.push('พบหัวงวดที่ใช้ colspan · คงตารางเดิมไว้เพื่อไม่ให้คอลัมน์เลื่อน');return table;}
      const comparisonCols=rows[headIndex].cells.map(cell=>({cell,type:comparisonType(cell.text)})).filter(item=>item.type&&!periodFrom(item.cell.text)).map(item=>{const owner=periodCols.filter(period=>period.column<=item.cell.column).sort((a,b)=>b.column-a.column)[0]||[...periodCols].sort((a,b)=>Math.abs(a.column-item.cell.column)-Math.abs(b.column-item.cell.column))[0];return owner?{type:item.type,column:item.cell.column,period:owner.period}:null;}).filter(Boolean);
      const comparisonTypes=[...new Set(comparisonCols.map(item=>item.type))],comparisonKnown=new Map(comparisonCols.map(item=>[item.period.id+':'+item.type,item.column]));
      const dataColumns=[...periodCols.map(item=>({column:item.column,span:item.span})),...comparisonCols.map(item=>({column:item.column,span:1}))];
      const tablePeriodStart=Math.min(...dataColumns.map(p=>p.column)),tablePeriodEnd=Math.max(...dataColumns.map(p=>p.column+p.span));
      const unsafeRow=rows.find((row,index)=>index!==headIndex&&row.cells.some(cell=>cell.span>1&&cell.column<tablePeriodEnd&&cell.endColumn>tablePeriodStart)&&targetLabelCells(row,periodCols,language).some(cell=>findMetric(summary.metrics,cell.text,'detail',selected)));
      if(unsafeRow){report.unmatched.push(`ตารางมี colspan ในแถว “${metricLabelCell(unsafeRow,periodCols)?.text||'ไม่ทราบชื่อ'}” · คงตารางเดิมไว้เพื่อไม่ให้ข้อมูลเหลื่อม`);return table;}
      const known=new Map(periodCols.map(x=>[x.period.id,x.column]));const wanted=selected.map(p=>p.id);const sourcePeriod=new Map(summary.displayPeriods.map(p=>[p.id,p]));
      let tableRows=rows.slice();
      for(let ri=0;ri<tableRows.length;ri++){
        const row=tableRows[ri];if(ri===headIndex)continue;
        const labelCandidates=targetLabelCells(row,periodCols,language);if(!labelCandidates.length)continue;
        let labelCell=labelCandidates[0],metric=null;for(const candidate of labelCandidates){metric=findMetric(summary.metrics,candidate.text,'detail',selected);if(metric){labelCell=candidate;break;}}
        if(!metric){const containsNumber=periodCols.some(period=>numberFrom(cellAtColumn(row,period.column)?.text)!==null);if(containsNumber)report.unmatched.push(`HTML “${labelCell.text}” มีค่าตัวเลขแต่จับคู่กับแถว workbook ไม่ได้ · คงแถวเดิมไว้`);continue;}
        matched=true;
        const startCol=tablePeriodStart,endCol=tablePeriodEnd;
        const baseCells=row.cells.filter(c=>c.endColumn<=startCol).map(c=>c.raw),trailingCells=row.cells.filter(c=>c.column>=endCol).map(c=>c.raw),valueCells=[],periodTemplate=cellAtColumn(row,periodCols[0].column),comparisonTemplate=comparisonCols.length?cellAtColumn(row,comparisonCols[0].column):null;
        for(const id of wanted){const oldIndex=known.get(id),old=oldIndex===undefined?null:cellAtColumn(row,oldIndex),source=metric.values[id];valueCells.push(financialValueCell(old,periodTemplate,source,metric.label));
          if(!source&&oldIndex===undefined)report.unmatched.push(`HTML “${labelCell.text}” ไม่มีค่า ${formatPeriodLabel(sourcePeriod.get(id)||{year:Number(id.slice(0,4)),quarter:Number(id.match(/-Q([1-4])/)?.[1]||0)},language)} จาก workbook`);
          for(const type of comparisonTypes){const compareColumn=comparisonKnown.get(id+':'+type),oldCompare=compareColumn===undefined?null:cellAtColumn(row,compareColumn),sourceCompare=metric.comparisons?.[id+':'+type];valueCells.push(financialValueCell(oldCompare,oldCompare||comparisonTemplate||periodTemplate,sourceCompare,metric.label+' '+type));if(sourceCompare&&sourceCompare.value!==null)matched=true;}
        }
        const newRaw=replaceRowCells(row,[...baseCells,...valueCells,...trailingCells]);
        if(newRaw!==row.raw){changed=true;tableRows[ri]={...row,raw:newRaw};report.changes.push({item:labelCell.text,oldValue:'ตารางเดิม',newValue:wanted.map(id=>formatPeriodLabel(sourcePeriod.get(id)||{year:Number(id.slice(0,4)),quarter:Number(id.match(/-Q([1-4])/)?.[1]||0)},language)).join(', '),source:metric.sheet+' row '+metric.row});}
      }
      // Restrict rewriting to tables containing recognizable financial metric rows.
      if(!tableRows.some((row,i)=>i!==headIndex&&targetLabelCells(row,periodCols,language).some(cell=>findMetric(summary.metrics,cell.text,'detail',selected))))return table;
      const header=tableRows[headIndex],firstPeriod=tablePeriodStart,lastPeriod=tablePeriodEnd;const base=header.cells.filter(c=>c.endColumn<=firstPeriod).map(c=>c.raw),trailing=header.cells.filter(c=>c.column>=lastPeriod).map(c=>c.raw);const template=cellAtColumn(header,periodCols[0].column),comparisonHeaderTemplate=comparisonCols.length?cellAtColumn(header,comparisonCols[0].column):null,headCells=[];for(const id of wanted){const found=known.get(id),cell=found===undefined?template:cellAtColumn(header,found),period=sourcePeriod.get(id)||{year:Number(id.slice(0,4)),quarter:Number(id.match(/-Q([1-4])/)?.[1]||0)},label=formatPeriodLabel(period,language);headCells.push(cell?replaceCellText(cell.raw,cell.text,label):`<th>${label}</th>`);for(const type of comparisonTypes){const compareColumn=comparisonKnown.get(id+':'+type),compareCell=compareColumn===undefined?comparisonHeaderTemplate:cellAtColumn(header,compareColumn);headCells.push(compareCell?replaceCellText(compareCell.raw,compareCell.text,type.toUpperCase()):`<th>${type.toUpperCase()}</th>`);}}
      tableRows[headIndex]={...header,raw:replaceRowCells(header,[...base,...headCells,...trailing])};
      // Apply row-level column filtering/insertion only if the table header exposes period columns.
      return tableRows.map((row,i)=>({row,i})).sort((a,b)=>b.row.start-a.row.start).reduce((result,item)=>result.slice(0,item.row.start)+item.row.raw+result.slice(item.row.end),table);
    });
    const beforeComponents=report.changes.length;updated=updateFinancialComponents(updated,summary,selected,language,report);const componentChanged=report.changes.length>beforeComponents;
    const beforeTitle=report.changes.length;updated=updateFinancialTitle(updated,selected,report);const titleChanged=report.changes.length>beforeTitle;
    if(!matched&&tables.length&&!titleChanged)report.unmatched.push('ไม่พบตาราง HTML ที่มีหัวงวดและชื่อรายการตรงกับแถวใน workbook');else if(matched&&!changed)report.notes.push('ตารางมีรายการตรงกับ workbook แต่ค่าที่เลือกไม่ทำให้โค้ดเปลี่ยนแปลง');
    if(tables.length===0&&!componentChanged&&!titleChanged)report.unmatched.push('ไม่พบ <table> หรือ counter-number ที่เชื่อมกับรายการใน workbook');
    return updated;
  }
  function updateFinancialComponents(html,summary,selected,language,report){
    const tags=[...html.matchAll(/<counter-number\b[^>]*>/gi)];if(!tags.length)return html;
    let output=html,changed=0;const latest=[...selected].sort((a,b)=>a.year-b.year||periodOrder(a)-periodOrder(b)||String(a.id).localeCompare(String(b.id))).at(-1);
    for(let i=tags.length-1;i>=0;i--){const tag=tags[i],start=tag.index;
      const before=output.slice(Math.max(0,start-500),start);
      const headings=[...before.matchAll(/<(?:h[1-6]|strong|label|p)\b[^>]*>([\s\S]*?)<\/(?:h[1-6]|strong|label|p)>/gi)];
      const context=headings.map(x=>plainText(x[1])).filter(Boolean).slice(-2).join(' ');
      if(!matchesTargetLanguage(context,language))continue;
      let metric=findMetric(summary.metrics,context,'summary',selected)||findMetric(summary.metrics,context,'detail',selected);
      if(!metric||!latest||!metric.values[latest.id])continue;
      const source=metric.values[latest.id],old=tag[0].match(/(?:\:number|\bnumber)\s*=\s*(["'])(.*?)\1/i);
      if(!old)continue;const value=source.value===null?emptyFinancialValue(source,old[2]):formatFinancialValue(old[2],source.value,metric.label);
      const replacement=tag[0].replace(old[0],old[0].replace(old[2],value));
      output=output.slice(0,start)+replacement+output.slice(start+tag[0].length);changed++;
      report.changes.push({item:context||metric.label,oldValue:old[2],newValue:value,source:source.address?metric.sheet+' row '+metric.row+' · '+source.address:metric.sheet+' row '+metric.row});
    }
    if(changed)report.notes.push('อัปเดตตัวเลข counter-number จากงวดล่าสุดที่เลือก ('+latest.label+')');
    return output;
  }
  function metricLabelCells(row,periodCols){const first=Math.min(...periodCols.map(p=>p.column));return row.cells.filter(cell=>cell.endColumn<=first&&cell.text.trim()).reverse();}
  function matchesTargetLanguage(text,language='en'){
    const value=plainText(text),hasThai=/[\u0E00-\u0E7F]/.test(value),hasLatin=/[A-Za-z]/.test(value);
    if(hasThai&&hasLatin)return true;
    if(!hasThai&&!hasLatin)return true;
    return language==='th'?hasThai:hasLatin;
  }
  function targetLabelCells(row,periodCols,language){return metricLabelCells(row,periodCols).filter(cell=>matchesTargetLanguage(cell.text,language));}
  function metricLabelCell(row,periodCols){return metricLabelCells(row,periodCols)[0];}
  function parseArray(text){
    const parts=[];let start=0,quote='',escaped=false;
    for(let i=0;i<text.length;i++){const ch=text[i];if(quote){if(escaped)escaped=false;else if(ch==='\\')escaped=true;else if(ch===quote)quote='';continue;}if(ch==='"'||ch==="'"){quote=ch;continue;}if(ch===','){parts.push(text.slice(start,i).trim());start=i+1;}}
    if(text.slice(start).trim())parts.push(text.slice(start).trim());
    return parts.map(x=>{try{return JSON.parse(x.replace(/^'/,'"').replace(/'$/,'"'));}catch{return x.replace(/^['"]|['"]$/g,'');}});
  }
  function scanObjectScopes(code){
    const stack=[],scopes=[];let quote='',escaped=false,lineComment=false,blockComment=false;
    for(let i=0;i<code.length;i++){const ch=code[i],next=code[i+1];
      if(lineComment){if(ch==='\n')lineComment=false;continue;}
      if(blockComment){if(ch==='*'&&next==='/'){blockComment=false;i++;}continue;}
      if(quote){if(escaped)escaped=false;else if(ch==='\\')escaped=true;else if(ch===quote)quote='';continue;}
      if(ch==='/'&&next==='/'){lineComment=true;i++;continue;}if(ch==='/'&&next==='*'){blockComment=true;i++;continue;}
      if(ch==='"'||ch==="'"||ch==='`'){quote=ch;continue;}
      if(ch==='{')stack.push(i);else if(ch==='}'&&stack.length){const start=stack.pop();scopes.push({start,end:i+1});}
    }
    return scopes.sort((a,b)=>(a.end-a.start)-(b.end-b.start));
  }
  function arrayProperty(code,scope,name,scopes){
    const body=code.slice(scope.start+1,scope.end-1),re=new RegExp('\\b'+name+'\\b\\s*:\\s*\\[','g');let match;
    while((match=re.exec(body))){const at=scope.start+1+match.index;if(scopes.some(child=>child!==scope&&child.start>scope.start&&child.end<scope.end&&at>child.start&&at<child.end))continue;
      const open=scope.start+1+match.index+match[0].lastIndexOf('[');let depth=1,quote='',escaped=false;
      for(let i=open+1;i<scope.end;i++){const ch=code[i];if(quote){if(escaped)escaped=false;else if(ch==='\\')escaped=true;else if(ch===quote)quote='';continue;}if(ch==='"'||ch==="'"){quote=ch;continue;}if(ch==='[')depth++;else if(ch===']'&&!--depth)return {open,close:i,inner:code.slice(open+1,i)};}
    }
    return null;
  }
  function directArrayProperties(code,scope,scopes){
    const body=code.slice(scope.start+1,scope.end-1),re=/([\w$]+)\s*:\s*\[/g,result=[];let match;
    while((match=re.exec(body))){const at=scope.start+1+match.index;if(scopes.some(child=>child!==scope&&child.start>scope.start&&child.end<scope.end&&at>child.start&&at<child.end))continue;
      const array=arrayProperty(code,scope,match[1],scopes);if(array)result.push({name:match[1],...array});
    }
    return result;
  }
  function directTextProperty(code,scope,name,scopes){
    const body=code.slice(scope.start+1,scope.end-1),re=new RegExp("\\b"+name+"\\b\\s*:\\s*([\"'])(.*?)\\1",'gi');let match;
    while((match=re.exec(body))){const at=scope.start+1+match.index;if(scopes.some(child=>child!==scope&&child.start>scope.start&&child.end<scope.end&&at>child.start&&at<child.end))continue;return match[2];}
    return '';
  }
  function directBooleanProperty(code,scope,name,scopes){
    const body=code.slice(scope.start+1,scope.end-1),re=new RegExp('\\b'+name+'\\b\\s*:\\s*(?:true|false)','gi');let match;
    while((match=re.exec(body))){const at=scope.start+1+match.index;if(scopes.some(child=>child!==scope&&child.start>scope.start&&child.end<scope.end&&at>child.start&&at<child.end))continue;return true;}
    return false;
  }
  function scopeKey(code,scope){const before=code.slice(Math.max(0,scope.start-100),scope.start);return before.match(/["']?([\w$]+)["']?\s*:\s*$/)?.[1]||'';}
  function arrayLiteral(item,template,kind,period,sourceThai,language='en'){
    if(kind==='number')return String(item);
    const text=String(template??'');
    if(isSubAnnual(period)){const buddhist=/(?:25\d{2}|26\d{2})/.test(text)||(!text&&sourceThai);const year=period.year+(buddhist?543:0);return JSON.stringify(formatPeriodLabel({...period,year},language));}
    const year=period.year+(/(?:25|26)\d{2}/.test(text)||(!text&&sourceThai)?543:0);return !text||/^\s*\d{4}\s*$/.test(text)?String(year):JSON.stringify(String(year));
  }
  function formatArrayItems(items,original){
    const multiline=/\r?\n/.test(original);if(!multiline)return items.join(', ');
    const newline=original.includes('\r\n')?'\r\n':'\n';const firstLine=original.match(/\r?\n([\t ]*)\S/)?.[1]||'    ';
    if(!items.length)return '';
    return newline+firstLine+items.join(','+newline+firstLine)+newline;
  }
  function updateFinancialJs(code,summary,selected,mode,report,language='en'){
    const scopes=scanObjectScopes(code),patches=[],handled=new Set(),qoqPatches=new Set();let changed=0;
    const sourceThai=/[\u0E00-\u0E7F]/.test(code);
    for(const scope of scopes){
      const periodArray=arrayProperty(code,scope,'period',scopes),dataArray=arrayProperty(code,scope,'dataset',scopes);
      if(!periodArray||!dataArray)continue;
      const oldPeriodTokens=parseArray(periodArray.inner),oldPeriods=oldPeriodTokens.map(periodFrom);
      const sectionKey=scopeKey(code,scope),quarterly=/(?:quarter|quarterly|qoq)/i.test(sectionKey),yearly=/(?:yearly|annual)/i.test(sectionKey),periodicName=/(?:monthly|month|semiannual|halfyear|year.?to.?date|ytd)/i.test(sectionKey);
      if(!oldPeriods.some(Boolean)&&!quarterly&&!yearly&&!periodicName)continue;
      const parentScopes=scopes.filter(x=>x.start<scope.start&&x.end>scope.end).sort((a,b)=>(a.end-a.start)-(b.end-b.start));
      const labels=[];
      for(const parent of [scope,...parentScopes]){
        for(const key of ['name','label','title','metric','series']){const value=directTextProperty(code,parent,key,scopes);if(value)labels.push(value);}
        const objectKey=scopeKey(code,parent);if(objectKey)labels.push(objectKey);
      }
      const comparisonKind=comparisonType(sectionKey)||labels.map(comparisonType).find(Boolean)||'';
      let metric=null;for(const label of labels){metric=findMetric(summary.metrics,label,'summary',selected);if(metric)break;}
      const oldDataTokens=parseArray(dataArray.inner),oldIndex=new Map();oldPeriods.forEach((period,index)=>{if(period)oldIndex.set(period.id,index);});
      const seriesProperties=directArrayProperties(code,scope,scopes).filter(series=>series.name!=='period'&&series.name!=='dataset'&&!/markPoint|axis|annotation|color|symbol|label/i.test(series.name));
      const seriesCandidates=seriesProperties.map(series=>{const type=comparisonType(series.name),comparisonMetric=type?(findMetric(summary.metrics,series.name,'summary',selected)||findMetric(summary.metrics,(metric?.label||labels[0]||'')+' '+series.name,'summary',selected)):null,seriesMetric=findMetric(summary.metrics,type?comparisonBase(series.name):series.name,'summary',selected)||(type?metric:null);return {series,type,values:parseArray(series.inner),metric:seriesMetric,comparisonMetric};});
      if(!metric)metric=seriesCandidates.find(candidate=>candidate.metric)?.metric||null;
      if(!metric){report.unmatched.push('ไม่พบแถวสรุปใน workbook ที่ตรงกับกราฟ '+(labels[0]||'ไม่มีชื่อ'));continue;}
      const scopeComparisonMetric=comparisonKind&&comparisonType(metric.label)!==comparisonKind?findMetric(summary.metrics,metric.label+' '+comparisonKind,'summary',selected):null;
      const matchedSeries=seriesCandidates.map(candidate=>{const type=candidate.type;return {...candidate,metric:candidate.metric||(type?metric:null),comparisonMetric:candidate.comparisonMetric||(type?findMetric(summary.metrics,metric.label+' '+candidate.series.name,'summary',selected):null)}}).filter(({values,metric:seriesMetric})=>seriesMetric&&values.length===oldPeriodTokens.length&&values.every(value=>value===null||numberFrom(value)!==null||typeof value==='string'&&(!value.trim()||/^[-–—]$/.test(value))));
      const sourceFor=(sourceMetric,comparisonMetric,period,type='')=>{if(!sourceMetric)return null;if(!type)return sourceMetric.values[period.id]||null;return sourceMetric.comparisons?.[period.id+':'+type]||(comparisonType(sourceMetric.label)===type?sourceMetric.values[period.id]:null)||(comparisonMetric&&comparisonType(comparisonMetric.label)===type?comparisonMetric.values[period.id]:null)||null;};
      const isPeriodic=quarterly||periodicName||oldPeriods.some(isSubAnnual);
      const targetPeriods=selected.filter(p=>isSubAnnual(p)===isPeriodic);
      const nextIds=[],nextPeriods=[],nextData=[];
      for(const period of targetPeriods){
        const at=oldIndex.get(period.id),source=sourceFor(metric,scopeComparisonMetric,period,comparisonKind);
        const keepOld=at!==undefined&&!source;
        if(keepOld){nextIds.push(period.id);nextPeriods.push(arrayLiteral(oldPeriodTokens[at],oldPeriodTokens[0], 'period',period,sourceThai,language));nextData.push(oldDataTokens[at]??'null');}
        else if(source){nextIds.push(period.id);nextPeriods.push(arrayLiteral(period.id,oldPeriodTokens[0],'period',period,sourceThai,language));nextData.push(source.value===null?emptyArrayValue(source,oldDataTokens[at],oldDataTokens[0]):String(Number(source.value.toPrecision?.(12)??source.value)));}
        else if(at!==undefined){nextIds.push(period.id);nextPeriods.push(arrayLiteral(oldPeriodTokens[at],oldPeriodTokens[0],'period',period,sourceThai,language));nextData.push(oldDataTokens[at]??'null');report.unmatched.push(`${metric.label} ไม่มีค่า ${formatPeriodLabel(period,language)} ใน workbook · คงค่าเดิมไว้`);}
        else report.unmatched.push(`${metric.label} ไม่มีค่า ${formatPeriodLabel(period,language)} ใน workbook และกราฟเดิม`);
      }
      const signature=scope.start+':'+dataArray.open;if(handled.has(signature))continue;handled.add(signature);
      patches.push({start:periodArray.open+1,end:periodArray.close,text:formatArrayItems(nextPeriods,periodArray.inner)});
      patches.push({start:dataArray.open+1,end:dataArray.close,text:formatArrayItems(nextData,dataArray.inner)});
      for(const {series,type,values,metric:seriesMetric,comparisonMetric} of matchedSeries){
        const aligned=[];for(const period of targetPeriods){
          if(!nextIds.includes(period.id))continue;
          const at=oldIndex.get(period.id),source=sourceFor(seriesMetric,comparisonMetric,period,type);
          if(at!==undefined&&!source)aligned.push(values[at]??'null');
          else if(source)aligned.push(source.value===null?emptyArrayValue(source,values[at],values[0]):String(Number(source.value.toPrecision?.(12)??source.value)));
          else aligned.push(at===undefined?'null':values[at]??'null');
        }
        patches.push({start:series.open+1,end:series.close,text:formatArrayItems(aligned,series.inner)});
        report.changes.push({item:series.name,oldValue:'chart series',newValue:aligned.length+' aligned points',source:seriesMetric.sheet+' row '+seriesMetric.row});
      }
      const chartScope=parentScopes.find(parent=>directBooleanProperty(code,parent,'QoQ',scopes)||directBooleanProperty(code,parent,'qoq',scopes));
      if(chartScope&&(!comparisonKind||comparisonKind==='qoq')){for(const key of ['QoQ','qoq']){const body=code.slice(chartScope.start+1,chartScope.end-1),re=new RegExp('('+key+'\\s*:\\s*)(true|false)','i'),match=re.exec(body);if(match){const at=chartScope.start+1+match.index+match[1].length;if(!scopes.some(child=>child!==chartScope&&child.start>chartScope.start&&child.end<chartScope.end&&at>child.start&&at<child.end)){const id=chartScope.start+':'+at;if(!qoqPatches.has(id)){patches.push({start:at,end:at+match[2].length,text:nextData.some(v=>numberFrom(v)!==null)?'true':'false'});qoqPatches.add(id);}break;}}}}
      changed++;report.changes.push({item:labels[0]||metric.label,oldValue:'chart '+(isPeriodic?'periodic':'yearly')+' dataset',newValue:nextData.length+' points',source:metric.sheet+' row '+metric.row});
    }
    if(!changed&&code.trim()&&!report.unmatched.length)report.unmatched.push('ไม่พบ JavaScript chart source ที่มี period/dataset และชื่อ series เชื่อมกับแถวสรุปใน workbook');
    let output=code;for(const patch of patches.sort((a,b)=>b.start-a.start))output=output.slice(0,patch.start)+patch.text+output.slice(patch.end);
    return output;
  }
  const updateFinancial=(payload,summary,selection,mode,language='en')=>{
    const periods=validateSelection(selection,summary);if(!['add','replace'].includes(mode))throw new Error('เลือกวิธีเพิ่มหรือแทนที่ข้อมูล');
    const report={changes:[],unmatched:[],notes:[]};
    const html=payload.html?updateFinancialHtml(payload.html,summary,periods,mode,language,report):'';
    const javascript=payload.javascript?updateFinancialJs(payload.javascript,summary,periods,mode,report,language):'';
    report.notes.push('ภาษา HTML ที่เลือก: '+({en:'English',th:'ภาษาไทย'}[language]||'English'));
    if(mode==='add'&&!summary.sourcePeriods.some(p=>!summary.existingPeriods.some(o=>o.id===p.id)))report.notes.push('ไม่พบงวดใหม่ในไฟล์; โหมดเพิ่มจึงอัปเดตตัวเลขของงวดที่เลือกจาก workbook');
    return {html,javascript,...report};
  };
  const updateShareholders=(payload,summary)=>{
    const report={changes:[],unmatched:[],notes:[]};const allRows=summary.shareholderRows||[];
    const uniqueByName=items=>[...new Map(items.map(r=>[normalize(r.name),r])).values()];
    const graphRows=uniqueByName(allRows.filter(r=>/graph|chart/i.test(r.sheet||'')));
    const listRows=uniqueByName(allRows.filter(r=>/shareholder|holder|list|รายชื่อ|ผู้ถือหุ้น/i.test(r.sheet||'')));
    const rows=listRows.length?listRows:uniqueByName(allRows);
    if(!rows.length)report.unmatched.push('ยังไม่พบตารางรายชื่อที่มีหัวคอลัมน์ Shareholder/Name และ Shares');
    const totalRows=graphRows.length?graphRows:rows.filter(r=>r.percent!==null);const total=totalRows.reduce((sum,row)=>sum+(row.percent??0),0);
    if(totalRows.length&&Math.abs(total-100)>0.01)report.notes.push(`เปอร์เซ็นต์จาก ${graphRows.length?'Graph':'ตารางรายชื่อ'} รวม ${total.toFixed(2)}% ไม่เท่ากับ 100.00% · ไม่ปรับยอดอัตโนมัติ`);
    if(totalRows.length&&Math.abs(total-100)<=0.01)report.notes.push(`ตรวจเปอร์เซ็นต์จาก ${graphRows.length?'Graph':'ตารางรายชื่อ'} รวม 100.00%`);
    let html=payload.html||'';
    const factRules={
      rightsType:{labels:['Rights Type','Type of rights','ประเภทสิทธิ','สิทธิประโยชน์'],format:'text'},
      overviewDate:{labels:['Overview Date','As of date','วันที่ข้อมูล','ณ วันที่'],format:'text'},
      totalShareholders:{labels:['Total Shareholders','Number of Shareholders','ผู้ถือหุ้นทั้งหมด','จำนวนผู้ถือหุ้น'],format:'number'},
      scripless:{labels:['Scripless','Scripless Holding','ระบบไร้ใบหุ้น','ถือครองผ่านระบบไร้ใบหุ้น'],format:'number'},
      freeFloatDate:{labels:['Free Float Date','Minority Date','วันที่ผู้ถือหุ้นรายย่อย'],format:'text'},
      minorShareholders:{labels:['Minor Shareholders','Number of Minor Shareholders','ผู้ถือหุ้นรายย่อย','จำนวนผู้ถือหุ้นรายย่อย'],format:'number'},
      freeFloatPercent:{labels:['Free Float','Free Float %','Free Float Percentage','สัดส่วนผู้ถือหุ้นรายย่อย','ร้อยละผู้ถือหุ้นรายย่อย'],format:'percent'}
    };
    for(const [id,fact] of Object.entries(summary.shareholderFacts||{})){
      const rule=factRules[id];if(!rule)continue;const rawNumber=numberFrom(fact.raw),numeric=rawNumber!==null?rawNumber:numberFrom(fact.value);const percentFact=rule.format==='percent'||(id==='scripless'&&/%/.test(fact.value));const replacement=rule.format==='number'&&numeric!==null?numeric.toLocaleString('en-US'):percentFact&&numeric!==null?numeric.toFixed(2)+'%':fact.value;
      const result=replaceSiblingValue(html,rule.labels,replacement);
      if(result.count===1){html=result.html;report.changes.push({item:id,oldValue:'HTML value',newValue:replacement,source:fact.source});}
      else report.unmatched.push(result.count?'พบ label “'+rule.labels[0]+'” หลายจุด จึงไม่แทนที่':'พบ '+id+' ใน workbook แต่ไม่พบ label/value ที่รองรับใน HTML');
    }
    const names=new Map(rows.map(r=>[normalize(r.name),r]));let found=0;
    html=html.replace(/<table\b[\s\S]*?<\/table>/gi,table=>{
      const trs=scanRows(table);if(!trs.length)return table;const header=trs.find(r=>/shareholder|ผู้ถือหุ้น/i.test(r.cells.map(c=>c.text).join(' ')));if(!header)return table;
      const h=header.cells.map(c=>normalize(c.text));const nameCol=h.findIndex(x=>/shareholder|name|ผู้ถือหุ้น|ชื่อ/.test(x));const sharesCol=h.findIndex(x=>/shares|จำนวนหุ้น|หุ้น/.test(x));const pctCol=h.findIndex(x=>/percent|percentage|ร้อยละ|สัดส่วน/.test(x));if(nameCol<0||sharesCol<0)return table;
      let changedTable=table;const dataRows=trs.slice(trs.indexOf(header)+1);for(const row of dataRows.slice().reverse()){const label=normalize(row.cells[nameCol]?.text);const item=names.get(label);if(!item)continue;let raw=row.raw;const replacements=[];if(row.cells[sharesCol]&&item.shares!==null)replacements.push([row.cells[sharesCol],formatLike(row.cells[sharesCol].text,item.shares)]);if(pctCol>=0&&row.cells[pctCol]&&item.percent!==null)replacements.push([row.cells[pctCol],formatLike(row.cells[pctCol].text,item.percent)]);for(const [cell,value] of replacements.reverse())raw=raw.slice(0,cell.start)+replaceCellText(cell.raw,cell.text,value)+raw.slice(cell.end);changedTable=changedTable.slice(0,row.start)+raw+changedTable.slice(row.end);found++;report.changes.push({item:item.name,oldValue:'ค่าจากโค้ดเดิม',newValue:`${item.shares===null?'':item.shares.toLocaleString('en-US')+' หุ้น'}${item.percent===null?'':` · ${item.percent.toFixed(2)}%`}`,source:item.source});}
      const template=dataRows[0];if(template){let additions='';for(const item of rows){if([...names.keys()].some(name=>name===normalize(item.name))&&!dataRows.some(row=>normalize(row.cells[nameCol]?.text)===normalize(item.name))){let raw=template.raw;const cells=scanRows(raw)[0]?.cells||[];const replacements=[];if(cells[nameCol])replacements.push([cells[nameCol],item.name]);if(cells[sharesCol]&&item.shares!==null)replacements.push([cells[sharesCol],formatLike(cells[sharesCol].text,item.shares)]);if(pctCol>=0&&cells[pctCol]&&item.percent!==null)replacements.push([cells[pctCol],formatLike(cells[pctCol].text,item.percent)]);for(const [cell,value] of replacements.reverse())raw=raw.slice(0,cell.start)+replaceCellText(cell.raw,cell.text,value)+raw.slice(cell.end);additions+=raw;found++;report.changes.push({item:item.name,oldValue:'ไม่มีแถวในโค้ดเดิม',newValue:`${item.shares===null?'':item.shares.toLocaleString('en-US')+' หุ้น'}${item.percent===null?'':` · ${item.percent.toFixed(2)}%`}`,source:item.source});}}
        if(additions){const close=/<\/tbody\s*>/i.exec(changedTable);const at=close?close.index:changedTable.toLowerCase().lastIndexOf('</table>');if(at>=0)changedTable=changedTable.slice(0,at)+additions+changedTable.slice(at);}}
      return changedTable;
    });
    if(!found&&rows.length)report.unmatched.push('ไม่พบแถวผู้ถือหุ้นใน HTML ที่ตรงกับรายชื่อและคอลัมน์จำนวนหุ้น');
    // The desktop skill accepts both window.chart.majorShareholders and the
    // common window.chart.shareholders shape. Restrict edits to an object whose
    // own name or enclosing chart key clearly identifies shareholder data.
    let javascript=payload.javascript||'';const sourceGraph=graphRows.filter(r=>r.percent!==null);const jsScopes=scanObjectScopes(javascript),shareholderArrays=[];
    for(const scope of jsScopes){
      const names=arrayProperty(javascript,scope,'labels',jsScopes)||arrayProperty(javascript,scope,'names',jsScopes)||arrayProperty(javascript,scope,'shareholderNames',jsScopes);
      const values=arrayProperty(javascript,scope,'dataset',jsScopes)||arrayProperty(javascript,scope,'data',jsScopes)||arrayProperty(javascript,scope,'percentages',jsScopes)||arrayProperty(javascript,scope,'values',jsScopes);
      if(!names||!values)continue;
      const parents=jsScopes.filter(x=>x.start<scope.start&&x.end>scope.end);
      const identity=[scopeKey(javascript,scope),...parents.map(x=>scopeKey(javascript,x))].join(' ');
      if(/major.?shareholders?|shareholders?|ownership/i.test(identity))shareholderArrays.push({scope,names,values});
    }
    if(shareholderArrays.length&&sourceGraph.length){const jsPatches=[];
      const arrayText=(items,original,strings=false)=>formatArrayItems(items.map((item,index)=>{
        if(!strings)return String(item);
        const old=parseArray(original)[index];const quote=typeof old==='string'&&original.trim().startsWith("'")?"'":'"';
        return quote+String(item).replaceAll('\\','\\\\').replaceAll(quote,'\\'+quote)+quote;
      }),original);
      for(const block of shareholderArrays){
        jsPatches.push({start:block.names.open+1,end:block.names.close,text:arrayText(sourceGraph.map(r=>r.name),block.names.inner,true)});
        jsPatches.push({start:block.values.open+1,end:block.values.close,text:arrayText(sourceGraph.map(r=>r.percent),block.values.inner)});
      }
      for(const patch of jsPatches.sort((a,b)=>b.start-a.start))javascript=javascript.slice(0,patch.start)+patch.text+javascript.slice(patch.end);
      report.changes.push({item:'Major shareholders graph',oldValue:'กราฟเดิม',newValue:`${sourceGraph.length} รายการ`,source:sourceGraph[0].source});
    }else if(payload.javascript.trim()&&shareholderArrays.length)report.unmatched.push('ไม่พบแถว Graph ที่มีเปอร์เซ็นต์ใน workbook · คง JavaScript เดิมไว้');
    else if(payload.javascript.trim()&&sourceGraph.length)report.unmatched.push('พบข้อมูล Graph ใน workbook แต่ไม่พบ labels/names และ dataset/data ใน object ผู้ถือหุ้น');
    else if(payload.javascript.trim())report.unmatched.push('ไม่พบ object กราฟผู้ถือหุ้นใน JavaScript');
    return {html,javascript,...report};
  };
  function replaceSiblingValue(html,labels,value){
    const nodes=[],stack=[],token=/<\!--[\s\S]*?-->|<\/?[A-Za-z][\w:-]*\b[^>]*>/g,voidTags=new Set(['area','base','br','col','embed','hr','img','input','link','meta','param','source','track','wbr']);let match;
    while((match=token.exec(html))){const raw=match[0];if(raw.startsWith('<!--'))continue;const closing=/^<\//.test(raw),tag=raw.match(/^<\/?([\w:-]+)/)?.[1]?.toLowerCase();if(!tag)continue;
      if(closing){let index=stack.length-1;while(index>=0&&stack[index].tag!==tag)index--;if(index>=0){for(let i=stack.length-1;i>=index;i--){stack[i].closeStart=match.index;stack[i].end=token.lastIndex;}stack.length=index;}continue;}
      const node={tag,start:match.index,contentStart:token.lastIndex,closeStart:token.lastIndex,end:token.lastIndex,parent:stack.at(-1)||null,children:[],rawOpen:raw};if(node.parent)node.parent.children.push(node);nodes.push(node);
      if(!voidTags.has(tag)&&! /\/\s*>$/.test(raw))stack.push(node);
      if(tag==='script'||tag==='style'){const close=html.toLowerCase().indexOf('</'+tag,token.lastIndex);if(close>=0)token.lastIndex=close;}
    }
    for(const node of stack){node.closeStart=html.length;node.end=html.length;}
    const inner=node=>html.slice(node.contentStart,node.closeStart);
    const hasValue=node=>/:(?:number|value|count|percent|percentage)\s*=|\b(?:number|value|count|percent|percentage)\s*=/i.test(node.rawOpen+inner(node))||Boolean(plainText(inner(node)));
    const matched=[];
    for(const labelNode of nodes){if(!['div','span','td','th','dd','dt','p','li','label','strong','small'].includes(labelNode.tag)&&!/^h[1-6]$/.test(labelNode.tag))continue;
      const label=plainText(inner(labelNode));if(!labels.some(item=>normalize(item)===normalize(label)))continue;
      let probe=labelNode,parent=labelNode.parent,target=null,levels=0;
      while(parent&&levels++<4&&!target){const at=parent.children.indexOf(probe);if(at>=0)target=parent.children.slice(at+1).find(hasValue)||null;if(!target){probe=parent;parent=parent.parent;}}
      if(target)matched.push({target,label});
    }
    const unique=[...new Map(matched.map(item=>[item.target.start,item])).values()];if(unique.length!==1)return {html,count:unique.length};
    const {target}=unique[0],raw=html.slice(target.start,target.end),old=plainText(inner(target)),replacement=replaceCellText(raw,old,value);
    return {html:html.slice(0,target.start)+replacement+html.slice(target.end),count:1};
  }
  return {parseWorkbook,inspect,periodFrom,extractExistingPeriods,extractFinancial,updateFinancial,updateShareholders,plainText,numberFrom,uniquePeriods};
})();
window.FolioEngine=FolioEngine;
