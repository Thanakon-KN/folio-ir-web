'use strict';

// Deterministic, browser-local update functions. These rules deliberately stop
// on ambiguous mappings instead of inventing values.
const FolioEngine = (() => {
  const periodFrom = value => {
    const text = String(value ?? '').replace(/<[^>]*>/g, ' ').replace(/&nbsp;|&#160;/gi, ' ').replace(/\s+/g, ' ').trim();
    const q = text.match(/^(?:([1-4])\s*Q|Q\s*([1-4])|(?:Quarter|ไตรมาส)\s*([1-4]))\s*[/ .-]?\s*((?:19|20|25|26)\d{2}|\d{2})(?:\b|$)/i);
    let year, quarter = 0;
    if (q) {
      year = Number(q[4]); quarter = Number(q[1] || q[2] || q[3]);
      if (year < 100) year += year >= 50 ? 2500 : 2000;
    } else {
      const annual = text.match(/^(?:(?:FY|ปี)\s*)?((?:19|20|25|26)\d{2})(?:\s*[AE])?$/i);
      if (!annual) return null;
      year = Number(annual[1]);
    }
    if (year >= 2400) year -= 543;
    if (year < 1900 || year > 2199) return null;
    return { id: String(year) + (quarter ? '-Q' + quarter : ''), year, quarter, label: quarter ? quarter + 'Q/' + year : String(year) };
  };
  const uniquePeriods = list => [...new Map(list.filter(Boolean).map(p => [p.id, p])).values()].sort((a,b)=>a.year-b.year||a.quarter-b.quarter);
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
    let result=Math.abs(value).toLocaleString(grouped?'en-US':'en-US',{useGrouping:grouped,minimumFractionDigits:decimals,maximumFractionDigits:Math.max(decimals, Math.min(8, decimals+2))});
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
          const cells=[];
          for(let c=range.s.c;c<=range.e.c;c++){
            const address=XLSX.utils.encode_cell({r,c}), cell=ws[address];
            const raw=cell?.v ?? '', text=cell?.w ?? (raw instanceof Date?raw.toISOString().slice(0,10):String(raw));
            if(String(text).trim()||String(raw).trim()) cells.push({address,row:r+1,col:c,text:String(text).trim(),raw,format:cell?.z||''});
          }
          if(cells.length) rows.push({row:r+1,cells});
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
        const row=all[ri], candidates=row.cells.map(cell=>({cell,period:periodFrom(cell.text)})).filter(x=>x.period);
        if(!candidates.length) continue;
        if(candidates.length<2 && !candidates.some(x=>x.period.quarter)) continue;
        sourcePeriods.push(...candidates.map(x=>x.period));
        const firstPeriodCol=Math.min(...candidates.map(x=>x.cell.col));
        const periods=candidates.map(x=>({period:x.period,col:x.cell.col}));
        for(let rj=ri+1;rj<all.length;rj++){
          const data=all[rj];
          if(data.row>all[rj-1].row+2) break;
          const before=data.cells.filter(cell=>cell.col<firstPeriodCol && cell.text.trim());
          if(!before.length) continue;
          const label=before[before.length-1].text;
          if(periodFrom(label)) break;
          const values={}; let found=false;
          for(const item of periods){
            const cell=data.cells.find(x=>x.col===item.col);
            const val=numberFrom(cell?.raw ?? cell?.text);
            if(val!==null){values[item.period.id]={value:val,text:cell?.text||String(val),address:cell?.address||''};found=true;}
          }
          if(found && normalize(label)) metrics.push({label,normalized:normalize(label),sheet:book.name+' / '+sheet.name,row:data.row,values});
        }
      }
    }
    return {periods:uniquePeriods(sourcePeriods),metrics};
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
        const nameCol=labels.findIndex(x=>/^(shareholder|shareholdername|nameofshareholder|name|majorholder|ชื่อผู้ถือหุ้น|ผู้ถือหุ้น)$/.test(x));
        const shareCol=labels.findIndex(x=>/^(shares|numberofshares|noofshares|shareholding|จำนวนหุ้น|หุ้น)$/.test(x));
        const pctCol=labels.findIndex(x=>/^(percent|percentage|shareholdingpercent|ownership|สัดส่วน|ร้อยละ)$/.test(x));
        const graphSheet=/graph|chart|กราฟ/i.test(sheet.name);
        if(nameCol<0||(shareCol<0&&!(graphSheet&&pctCol>=0)))continue;
        for(const row of sheet.rows.slice(i+1)){const name=row.cells.find(c=>c.col===header.cells[nameCol].col)?.text||'';const share=shareCol>=0?row.cells.find(c=>c.col===header.cells[shareCol].col):null;const pct=pctCol>=0?row.cells.find(c=>c.col===header.cells[pctCol].col):null;const shares=share?numberFrom(share.raw??share.text):null;const percent=pct?numberFrom(pct.raw??pct.text):null;if(name&&(shares!==null||percent!==null))rows.push({name,shares,percent,sheet:sheet.name,source:`${book.name} / ${sheet.name} / row ${row.row}`});}
      }
    }
    return rows;
  }
  const validateSelection=(selection,summary)=>{
    if(!Array.isArray(selection)||!selection.length)throw new Error('เลือกงวดที่ต้องการแสดงอย่างน้อย 1 งวด');
    const valid=new Set(summary.displayPeriods.map(p=>p.id));if(selection.some(id=>!valid.has(id))||new Set(selection).size!==selection.length)throw new Error('ตัวเลือกงวดไม่ตรงกับผลตรวจ กรุณาตรวจไฟล์ใหม่');return summary.displayPeriods.filter(p=>selection.includes(p.id));
  };
  const findMetric=(metrics,label)=>{
    const key=normalize(label);if(!key)return null;
    const exact=metrics.filter(m=>m.normalized===key);if(exact.length===1)return exact[0];
    const likely=metrics.filter(m=>m.normalized.length>3&&key.length>3&&(m.normalized.includes(key)||key.includes(m.normalized)));
    return likely.length===1?likely[0]:null;
  };
  function scanRows(html){
    const out=[];const re=/<tr\b[^>]*>[\s\S]*?<\/tr>/gi;let m;
    while((m=re.exec(html))){const raw=m[0], cells=[];const cellRe=/<(td|th)\b([^>]*)>([\s\S]*?)<\/\1>/gi;let c;
      while((c=cellRe.exec(raw)))cells.push({tag:c[1],attrs:c[2],inner:c[3],raw:c[0],start:c.index,end:cellRe.lastIndex,text:plainText(c[3])});
      if(cells.length)out.push({raw,start:m.index,end:re.lastIndex,cells});}
    return out;
  }
  const replaceCellText=(raw,oldText,newText)=>{
    let next=raw;
    // Vue numeric components commonly store a value in :number or number attributes.
    if(/(?:\:number|\bnumber)\s*=\s*(["'])/.test(next)){
      const re=/(\:?number\s*=\s*)(["'])([\s\S]*?)\2/i;return next.replace(re,(_,prefix,q,value)=>prefix+q+value.replace(/-?\d[\d,.]*/,newText)+q);
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
  function replaceRowCells(row,newCells){
    let raw=row.raw;for(let i=row.cells.length-1;i>=0;i--){const c=row.cells[i];raw=raw.slice(0,c.start)+raw.slice(c.end);}const close=raw.lastIndexOf('</tr>');return close<0?row.raw:raw.slice(0,close)+newCells.join('')+raw.slice(close);
  }
  function updateFinancialHtml(html,summary,selected,mode,report){
    const tables=[...html.matchAll(/<table\b[\s\S]*?<\/table>/gi)];let changed=false,matched=false;
    const updated=html.replace(/<table\b[\s\S]*?<\/table>/gi,table=>{
      const rows=scanRows(table);let headIndex=-1,periodCols=[];
      rows.forEach((row,index)=>{const ps=row.cells.map((c,i)=>({period:periodFrom(c.text),i})).filter(x=>x.period);if(ps.length>periodCols.length&&ps.length>=1){periodCols=ps;headIndex=index;}});
      if(headIndex<0||!periodCols.length)return table;
      const known=new Map(periodCols.map(x=>[x.period.id,x.i]));const wanted=selected.map(p=>p.id);const sourcePeriod=new Map(summary.displayPeriods.map(p=>[p.id,p]));
      let tableRows=rows.slice();
      for(let ri=0;ri<tableRows.length;ri++){
        const row=tableRows[ri];if(ri===headIndex)continue;
        const labelCell=metricLabelCell(row,periodCols);if(!labelCell)continue;
        const metric=findMetric(summary.metrics,labelCell.text);
        if(!metric)continue;
        matched=true;
        const startCol=Math.min(...periodCols.map(p=>p.i));const baseCells=row.cells.filter((_,i)=>i<startCol).map(c=>c.raw);const valueCells=[];
        for(const id of wanted){const oldIndex=known.get(id),old=oldIndex===undefined?null:row.cells[oldIndex];const source=metric.values[id];let value;
          if(mode==='add'&&old)value=old.text;
          else if(source)value=formatLike(old?.text??source.text,source.value);
          else if(old)value=old.text;
          else {value='';report.unmatched.push(`HTML “${labelCell.text}” ไม่มีค่า ${sourcePeriod.get(id)?.label||id} จาก workbook`);}
          if(old){let raw=replaceCellText(old.raw,old.text,value);valueCells.push(raw);}else{const template=row.cells[periodCols[0].i];if(template)valueCells.push(replaceCellText(template.raw,template.text,value));}
        }
        const newRaw=replaceRowCells(row,[...baseCells,...valueCells]);
        if(newRaw!==row.raw){changed=true;tableRows[ri]={...row,raw:newRaw};report.changes.push({item:labelCell.text,oldValue:'ตารางเดิม',newValue:wanted.map(id=>sourcePeriod.get(id)?.label||id).join(', '),source:metric.sheet+' row '+metric.row});}
      }
      // Restrict rewriting to tables containing recognizable financial metric rows.
      if(!tableRows.some((row,i)=>i!==headIndex&&findMetric(summary.metrics,metricLabelCell(row,periodCols)?.text)))return table;
      const header=tableRows[headIndex],firstPeriod=Math.min(...periodCols.map(p=>p.i));const base=header.cells.filter((_,i)=>i<firstPeriod).map(c=>c.raw);const template=header.cells[periodCols[0].i];const headCells=wanted.map(id=>{const found=known.get(id);const cell=found===undefined?template:header.cells[found];const label=sourcePeriod.get(id)?.label||id;return cell?replaceCellText(cell.raw,cell.text,label):`<th>${label}</th>`;});
      tableRows[headIndex]={...header,raw:replaceRowCells(header,[...base,...headCells])};
      // Apply row-level column filtering/insertion only if the table header exposes period columns.
      return tableRows.map((row,i)=>({row,i})).sort((a,b)=>b.row.start-a.row.start).reduce((result,item)=>result.slice(0,item.row.start)+item.row.raw+result.slice(item.row.end),table);
    });
    if(!matched)report.unmatched.push('ไม่พบตาราง HTML ที่มีหัวงวดและชื่อรายการตรงกับแถวใน workbook');else if(!changed)report.notes.push('ตารางมีรายการตรงกับ workbook แต่ค่าที่เลือกไม่ทำให้โค้ดเปลี่ยนแปลง');
    if(tables.length===0)report.unmatched.push('ไม่พบ <table> ใน HTML/Vue');
    return updated;
  }
  function metricLabelCell(row,periodCols){const first=Math.min(...periodCols.map(p=>p.i));return row.cells.filter((cell,i)=>i<first&&cell.text.trim()).at(-1);}
  function parseArray(text){return text.split(',').map(x=>x.trim()).filter(Boolean).map(x=>{try{return JSON.parse(x.replace(/^'/,'"').replace(/'$/,'"'));}catch{return x.replace(/^['"]|['"]$/g,'');}});}
  function updateFinancialJs(code,summary,selected,mode,report){
    // Update well-formed chart objects only when a nearby property name maps
    // uniquely to a workbook metric. Unknown chart structures remain untouched.
    let output=code;const periods=[...code.matchAll(/(?:["']?period["']?)\s*:\s*\[([^\]]*)\]/g)];let changed=0;
    for(let pi=periods.length-1;pi>=0;pi--){const match=periods[pi],inside=match[1],periodVals=parseArray(inside).map(periodFrom);if(!periodVals.some(Boolean))continue;
      const start=match.index,close=output.indexOf(']',start)+1;const nextChunk=output.slice(close,Math.min(output.length,close+800));const ds=nextChunk.match(/(?:["']?dataset["']?)\s*:\s*\[([^\]]*)\]/);if(!ds)continue;
      const before=output.slice(Math.max(0,start-700),start);const labels=[...before.matchAll(/(?:name|label|title|metric|series)\s*:\s*(["'])(.*?)\1/gi)].map(x=>x[2]);let metric=null;for(const label of labels.reverse()){metric=findMetric(summary.metrics,label);if(metric)break;}
      if(!metric)continue;const oldValues=parseArray(ds[1]);const selectedSource=selected.map(p=>({period:p,source:metric.values[p.id]}));
      const quarterlyGroup=periodVals.some(p=>p?.quarter),aligned=selectedSource.filter(x=>Boolean(x.period.quarter)===quarterlyGroup);
      const nextPeriods=[],nextData=[];for(const item of aligned){const oldIndex=periodVals.findIndex(p=>p?.id===item.period.id);const keep=mode==='add'&&oldIndex>=0;const value=keep?oldValues[oldIndex]:item.source?.value??(oldIndex>=0?oldValues[oldIndex]:null);if(value===null)continue;nextPeriods.push(item.period.quarter?`${item.period.quarter}Q/${item.period.year}`:item.period.year);nextData.push(value);}
      const newPeriodText=nextPeriods.map(x=>typeof x==='number'?String(x):JSON.stringify(x)).join(', '),newDataText=nextData.join(', ');
      const pOpen=start+match[0].indexOf('['),pClose=start+match[0].lastIndexOf(']');
      const dOpen=close+ds.index+ds[0].indexOf('['),dClose=dOpen+ds[0].lastIndexOf(']');
      const patches=[{start:pOpen+1,end:pClose,text:newPeriodText},{start:dOpen+1,end:dClose,text:newDataText}].sort((a,b)=>b.start-a.start);
      if(quarterlyGroup){const before=output.slice(Math.max(0,start-1200),start);const matches=[...before.matchAll(/QoQ\s*:\s*(?:true|false)/g)];if(matches.length){const last=matches[matches.length-1],qStart=Math.max(0,start-1200)+last.index+last[0].indexOf(':')+1;qStart+=before.slice(last.index+last[0].indexOf(':')+1).match(/^\s*/)?.[0].length||0;const qEnd=qStart+(before.slice(qStart-Math.max(0,start-1200)).match(/(?:true|false)/)?.[0].length||0);patches.push({start:qStart,end:qEnd,text:nextData.length?'true':'false'});}}
      for(const patch of patches)output=output.slice(0,patch.start)+patch.text+output.slice(patch.end);
      changed++;report.changes.push({item:labels[0]||metric.label,oldValue:'chart dataset',newValue:String(nextData.length)+' จุดข้อมูล'+(quarterlyGroup?` · QoQ ${nextData.length?'true':'false'}`:''),source:metric.sheet+' row '+metric.row});
    }
    if(!changed&&code.trim())report.unmatched.push('ไม่พบ JavaScript chart object ที่มี period/dataset และชื่อ series ตรงกับแถวใน workbook');
    return output;
  }
  const updateFinancial=(payload,summary,selection,mode)=>{
    const periods=validateSelection(selection,summary);if(!['add','replace'].includes(mode))throw new Error('เลือกวิธีเพิ่มหรือแทนที่ข้อมูล');
    const report={changes:[],unmatched:[],notes:[]};
    const html=payload.html?updateFinancialHtml(payload.html,summary,periods,mode,report):'';
    const javascript=payload.javascript?updateFinancialJs(payload.javascript,summary,periods,mode,report):'';
    if(mode==='add'&&!summary.sourcePeriods.some(p=>!summary.existingPeriods.some(o=>o.id===p.id)))report.notes.push('ไม่พบงวดใหม่ในไฟล์; โหมดเพิ่มจึงคงค่าของงวดซ้ำไว้');
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
      const rule=factRules[id];if(!rule)continue;const numeric=numberFrom(fact.raw??fact.value);const replacement=rule.format==='number'&&numeric!==null?numeric.toLocaleString('en-US'):rule.format==='percent'&&numeric!==null?numeric.toFixed(2)+'%':fact.value;
      const result=replaceSiblingValue(html,rule.labels,replacement);
      if(result.count===1){html=result.html;report.changes.push({item:id,oldValue:'HTML value',newValue:replacement,source:fact.source});}
      else report.unmatched.push(result.count?'พบ label “'+rule.labels[0]+'” หลายจุด จึงไม่แทนที่':'พบ '+id+' ใน workbook แต่ไม่พบ label/value ที่รองรับใน HTML');
    }
    const names=new Map(rows.map(r=>[normalize(r.name),r]));let found=0;
    html=html.replace(/<table\b[\s\S]*?<\/table>/gi,table=>{
      const trs=scanRows(table);if(!trs.length)return table;const header=trs.find(r=>/shareholder|ผู้ถือหุ้น/i.test(r.cells.map(c=>c.text).join(' ')));if(!header)return table;
      const h=header.cells.map(c=>normalize(c.text));const nameCol=h.findIndex(x=>/shareholder|name|ผู้ถือหุ้น|ชื่อ/.test(x));const sharesCol=h.findIndex(x=>/shares|จำนวนหุ้น|หุ้น/.test(x));const pctCol=h.findIndex(x=>/percent|percentage|ร้อยละ|สัดส่วน/.test(x));if(nameCol<0||sharesCol<0)return table;
      let changedTable=table;for(const row of trs.slice(trs.indexOf(header)+1).reverse()){const label=normalize(row.cells[nameCol]?.text);const item=names.get(label);if(!item)continue;let raw=row.raw;const replacements=[];if(row.cells[sharesCol]&&item.shares!==null)replacements.push([row.cells[sharesCol],formatLike(row.cells[sharesCol].text,item.shares)]);if(pctCol>=0&&row.cells[pctCol]&&item.percent!==null)replacements.push([row.cells[pctCol],formatLike(row.cells[pctCol].text,item.percent)]);for(const [cell,value] of replacements.reverse())raw=raw.slice(0,cell.start)+replaceCellText(cell.raw,cell.text,value)+raw.slice(cell.end);changedTable=changedTable.slice(0,row.start)+raw+changedTable.slice(row.end);found++;report.changes.push({item:item.name,oldValue:'ค่าจากโค้ดเดิม',newValue:`${item.shares===null?'':item.shares.toLocaleString('en-US')+' หุ้น'}${item.percent===null?'':` · ${item.percent.toFixed(2)}%`}`,source:item.source});}
      return changedTable;
    });
    if(!found&&rows.length)report.unmatched.push('ไม่พบแถวผู้ถือหุ้นใน HTML ที่ตรงกับรายชื่อและคอลัมน์จำนวนหุ้น');
    // For JS, only replace common majorShareholders arrays when labels match.
    let javascript=payload.javascript||'';const sourceGraph=graphRows.filter(r=>r.percent!==null);const objectStart=javascript.indexOf('majorShareholders');
    if(objectStart>=0&&sourceGraph.length){const labelsRe=/((?:labels|names|shareholderNames)\s*:\s*)\[([^\]]*)\]/i;const dataRe=/((?:dataset|data|percentages|values)\s*:\s*)\[([^\]]*)\]/i;const tail=javascript.slice(objectStart);const lm=tail.match(labelsRe),dm=tail.match(dataRe);if(lm&&dm){const labels=sourceGraph.map(r=>JSON.stringify(r.name)).join(', '),values=sourceGraph.map(r=>r.percent).join(', ');const changed=tail.replace(labelsRe,`$1[${labels}]`).replace(dataRe,`$1[${values}]`);javascript=javascript.slice(0,objectStart)+changed;report.changes.push({item:'Major shareholders graph',oldValue:'กราฟเดิม',newValue:`${sourceGraph.length} รายการ`,source:sourceGraph[0].source});}else report.unmatched.push('พบ majorShareholders ใน JavaScript แต่ไม่พบ labels/names และ dataset/data ในรูปแบบที่รองรับ');}
    else if(payload.javascript.trim()&&objectStart>=0)report.unmatched.push('ไม่พบแถว Graph ที่มีเปอร์เซ็นต์ใน workbook · คง JavaScript เดิมไว้');else if(payload.javascript.trim())report.unmatched.push('ไม่พบ object ชื่อ majorShareholders ใน JavaScript');
    report.notes.push('กฎเว็บตรวจ Overview และ Free float จากตำแหน่งที่ระบุได้ แต่ยังไม่เปลี่ยนฟิลด์อิสระที่ไม่ตรงกับตารางหรือ object ที่รองรับ');
    return {html,javascript,...report};
  };
  function replaceSiblingValue(html,labels,value){
    let count=0,output=html;const tag='(?:div|span|td|th|dd|dt|p|li)';
    const re=new RegExp(`(<(${tag})\\b[^>]*>)([\\s\\S]*?)(<\\/\\2>)(\\s*)(<(${tag})\\b[^>]*>)([\\s\\S]*?)(<\\/\\7>)`,'gi');
    output=output.replace(re,(whole,open,labelTag,labelInner,labelClose,space,valueOpen,valueTag,valueInner,valueClose)=>{
      const actual=normalize(labelInner);if(!labels.some(x=>normalize(x)===actual))return whole;
      count++;if(count>1)return whole;const full=valueOpen+valueInner+valueClose;const next=replaceCellText(full,plainText(valueInner),value);return open+labelInner+labelClose+space+next;
    });
    return {html:output,count};
  }
  return {parseWorkbook,inspect,periodFrom,extractExistingPeriods,extractFinancial,updateFinancial,updateShareholders,plainText,numberFrom,uniquePeriods};
})();
window.FolioEngine=FolioEngine;
