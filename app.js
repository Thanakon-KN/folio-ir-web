'use strict';
(() => {
  const $=selector=>document.querySelector(selector), $$=selector=>[...document.querySelectorAll(selector)];
  const ui={files:[],workbook:null,inspection:null,inspectBusy:false,busy:false,result:false,notice:'',drafts:{},selectedPeriods:[],mode:''};
  const maxFiles=5,maxBytes=20*1024*1024;
  $('#active-skill').textContent='ฟังก์ชันในเบราว์เซอร์';
  $$('[data-guide-check="connection"]').forEach(el=>el.hidden=true);
  $('#guide-title').textContent='เริ่มจากเลือกไฟล์และโค้ดต้นฉบับ';$('#guide-detail').textContent='ฟังก์ชันจะตรวจข้อมูลในเครื่องและรายงานเมื่อจับคู่รายการไม่ได้';$('#guide-phase').textContent='เตรียมข้อมูล';
  const key=()=>`folio-ir-web-${activeWorkflow}`;
  const hasCode=()=>Boolean($('#html-code').value.trim()||$('#javascript-code').value.trim());
  const mode=()=>isShareholder()?'shareholder':$('input[name="update-mode"]:checked')?.value||'';
  const targetLanguage=()=>$('input[name="target-language"]:checked')?.value||'en';
  const periodOrder=()=>$('#period-order')?.value||'asc';
  const orderedPeriods=periods=>periodOrder()==='desc'?[...periods].reverse():periods;
  const selected=()=>$$('input[name="display-period"]:checked').map(input=>input.value);
  const msg=(text,kind='')=>{ui.notice=text;$('#status').textContent=text;$('#status').dataset.kind=kind;renderGuide();};
  function fileProblem(){if(ui.files.length>maxFiles)return`เลือกได้ไม่เกิน ${maxFiles} ไฟล์`;const size=ui.files.reduce((n,f)=>n+f.size,0);if(size>maxBytes)return`ขนาดไฟล์รวมเกิน ${maxBytes/1024/1024} MB`;const allowed=['xlsx','xls','xlsm','csv','tsv'];const bad=ui.files.find(f=>!allowed.includes(f.name.split('.').pop().toLowerCase()));return bad?`ไม่รองรับไฟล์ ${bad.name}`:'';}
  function readyToInspect(){return !ui.inspectBusy&&!ui.busy&&!fileProblem()&&ui.files.length>0&&hasCode();}
  function readyToUpdate(){return Boolean(ui.inspection&&!ui.busy&&!ui.inspectBusy&&(isShareholder()||(selected().length&&mode())));}
  function updateButtons(){
    $('#inspect-files').disabled=!readyToInspect();$('#run-update').disabled=!readyToUpdate();$('#cancel-update').hidden=true;
    $('#display-selection').hidden=!ui.inspection||isShareholder();$('#display-selection').disabled=!ui.inspection;
    $('#update-language').hidden=isShareholder();$('#update-language').disabled=isShareholder();
    $('#update-mode').hidden=!ui.inspection||isShareholder();$('#update-mode').disabled=!ui.inspection;
    $('#inspect-help').textContent=ui.inspectBusy?'กำลังอ่านไฟล์ในเบราว์เซอร์…':!hasCode()?'วาง HTML/Vue หรือ JavaScript ก่อนตรวจไฟล์':!ui.files.length?'เลือกไฟล์ข้อมูลต้นทางก่อนตรวจไฟล์':fileProblem()||'พร้อมตรวจข้อมูลในเครื่อง · หากแก้ไฟล์หรือโค้ด ต้องตรวจใหม่';
    $('#run-help').textContent=ui.busy?'กำลังสร้างโค้ดด้วยกฎที่กำหนดไว้…':!ui.inspection?'ตรวจไฟล์ก่อนดำเนินการ':isShareholder()?'พร้อมอัปเดตตามส่วนของผู้ถือหุ้นที่จับคู่ได้':'เลือกงวดและวิธีอัปเดตก่อนสร้างผลลัพธ์';
    renderGuide();
  }
  function renderGuide(){
    if(!$('#workflow-guide'))return;
    const phase=ui.result?'ตรวจทานและส่งออก':ui.inspection?(isShareholder()?'ตรวจไฟล์ต้นทาง':'ตรวจและเลือกงวด'):ui.files.length&&hasCode()?'เตรียมตรวจไฟล์':'เตรียมข้อมูล';
    const step=ui.result?4:ui.inspection?2:1;
    $('#guide-phase').textContent=step+'/4 · '+phase;$('#dock-title').textContent=ui.result?'ผลลัพธ์พร้อมตรวจ':ui.inspection?'เลือกข้อมูลและสร้างผลลัพธ์':'เตรียมไฟล์และโค้ด';
    const checks={code:hasCode(),files:ui.files.length>0,inspection:Boolean(ui.inspection),periods:isShareholder()||selected().length>0,mode:isShareholder()||Boolean(mode())};
    const count=Object.entries(checks).filter(([name])=>!isShareholder()||!['periods','mode'].includes(name)).filter(([,v])=>v).length;
    $('#guide-ready-count').textContent=`พร้อม ${count}/${isShareholder()?3:5} รายการ`;
    for(const [name,done] of Object.entries(checks)){const button=$(`[data-guide-check="${name}"]`);if(button){button.dataset.ready=String(done);button.hidden=isShareholder()&&['periods','mode'].includes(name);}}
    const state=ui.result?['ผลลัพธ์พร้อมตรวจทาน','อ่านรายงาน ตรวจ HTML/JavaScript แล้วคัดลอกหรือดาวน์โหลดได้']:ui.inspection&&readyToUpdate()?['พร้อมสร้างผลลัพธ์','ข้อมูลและตัวเลือกครบแล้ว ฟังก์ชันจะเปลี่ยนเฉพาะจุดที่จับคู่ได้ชัดเจน']:ui.inspection?['เลือกข้อมูลที่จะอัปเดต','เลือกอย่างน้อยหนึ่งงวดและวิธีเพิ่มหรือแทนที่']:['เตรียมไฟล์และโค้ด','ตรวจงวดและส่วนข้อมูลที่พบก่อนสร้างผลลัพธ์'];
    $('#guide-title').textContent=state[0];$('#guide-detail').textContent=state[1];
    const action=$('#guide-action'), label=action.querySelector('[data-action-text]');
    action.disabled=ui.busy||ui.inspectBusy;label.textContent=ui.result?'ดูผลลัพธ์':ui.inspection?'สร้างผลลัพธ์':'ตรวจไฟล์';
    document.querySelectorAll('[data-step]').forEach(li=>{const n=Number(li.dataset.step);li.classList.toggle('active',n===step);li.classList.toggle('done',n<step);if(n===step)li.setAttribute('aria-current','step');else li.removeAttribute('aria-current');});
    $('#workflow-dock').hidden=true;
  }
  function renderFiles(){const list=$('#file-list');list.replaceChildren(...ui.files.map((file,index)=>{const row=document.createElement('div');row.className='file';const name=document.createElement('span');name.textContent=`${file.name} · ${(file.size/1024).toFixed(0)} KB`;const remove=document.createElement('button');remove.type='button';remove.className='file-remove';remove.textContent='นำออก';remove.addEventListener('click',()=>{ui.files.splice(index,1);invalidate('ไฟล์เปลี่ยนแล้ว กรุณาตรวจใหม่');renderFiles();});row.append(name,remove);return row;}));$('#files-feedback').textContent=fileProblem()||(ui.files.length?`เลือกแล้ว ${ui.files.length} ไฟล์ · รวม ${(ui.files.reduce((n,f)=>n+f.size,0)/1024/1024).toFixed(2)} MB`:'ยังไม่ได้เลือกไฟล์');$('#files-feedback').dataset.invalid=String(Boolean(fileProblem()));}
  function invalidate(reason){ui.inspection=null;ui.workbook=null;ui.selectedPeriods=[];ui.mode='';ui.result=false;$('#inspection-report').hidden=true;$('#inspection-report').replaceChildren();$('#period-options').replaceChildren();$('#period-selection-summary').textContent='';$('#mode-add').checked=false;$('#mode-replace').checked=false;clearOutput();if(reason)msg(reason);updateButtons();}
  function clearOutput(){for(const id of ['html-result','javascript-result','report-result','notes-result'])$('#'+id).value='';$('#result-badge').textContent='รออัปเดต';$('#result-badge').dataset.state='empty';$('.result-card').dataset.state='empty';$$('.copy-result,.download-result').forEach(b=>b.disabled=true);}
  function renderInspection(summary){
    const container=$('#inspection-report');container.replaceChildren();container.hidden=false;
    const heading=document.createElement('h3');heading.textContent='ผลการตรวจในเบราว์เซอร์';container.append(heading);
    const detail=document.createElement('p');detail.textContent=`อ่าน ${summary.files.length} ไฟล์ · ${summary.cellCount.toLocaleString()} เซลล์`+(summary.metrics?` · พบ ${summary.metrics.length} แถวตัวเลข`:summary.shareholderRows?` · พบ ${summary.shareholderRows.length} รายชื่อ`:'' );container.append(detail);
    for(const file of summary.files){const title=document.createElement('strong');title.textContent=file.name;container.append(title);const ul=document.createElement('ul');for(const sheet of file.sheets){const li=document.createElement('li');li.textContent=`${sheet.name}: ${sheet.rows} แถว, ${sheet.cells} เซลล์`;ul.append(li);}container.append(ul);}
    if(summary.sourcePeriods?.length){const p=document.createElement('p');p.textContent='งวดในไฟล์: '+summary.sourcePeriods.map(x=>x.label).join(', ');container.append(p);}
    if(summary.existingPeriods?.length){const p=document.createElement('p');p.textContent='งวดในโค้ด: '+summary.existingPeriods.map(x=>x.label).join(', ');container.append(p);}
    if(summary.blocks){for(const block of summary.blocks){const p=document.createElement('p');p.textContent=`${block.label}: ${block.references.length?block.references.join(' · '):'ไม่พบ'}`;container.append(p);}}
    if(summary.shareholderFacts&&Object.keys(summary.shareholderFacts).length){const facts=document.createElement('ul');for(const [key,item] of Object.entries(summary.shareholderFacts)){const li=document.createElement('li');li.textContent=`${key}: ${item.value} · ${item.source}`;facts.append(li);}container.append(facts);}
    for(const warning of summary.warnings||[]){const p=document.createElement('p');p.className='inspection-warning';p.textContent=warning;container.append(p);}
    if(!isShareholder())renderPeriods(summary.displayPeriods||[]);
    $('#result-badge').textContent='ตรวจไฟล์แล้ว';$('#result-badge').dataset.state='ready';
  }
  function renderPeriods(periods){const host=$('#period-options'),checked=new Set(selected());host.replaceChildren();for(const period of orderedPeriods(periods)){const label=document.createElement('label');label.className='period-choice';const input=document.createElement('input');input.type='checkbox';input.name='display-period';input.value=period.id;input.checked=checked.has(period.id);input.addEventListener('change',()=>{updatePeriodSummary();updateButtons();clearOutput();});const strong=document.createElement('strong');strong.textContent=period.label;const badges=[];if(period.inSource)badges.push('ไฟล์');if(period.inExisting)badges.push('โค้ดเดิม');const small=document.createElement('small');small.textContent=badges.join(' · ');label.append(input,strong,small);host.append(label);}updatePeriodSummary();}
  function updatePeriodSummary(){const chosen=orderedPeriods(ui.inspection?.displayPeriods?.filter(p=>selected().includes(p.id))||[]);$('#period-selection-summary').textContent=chosen.length?'จะแสดง '+chosen.length+' งวด: '+chosen.map(p=>p.label).join(', '):'ยังไม่ได้เลือกงวด · เลือกอย่างน้อย 1 งวดก่อนอัปเดต';}
  async function inspect(){if(!readyToInspect())return;ui.inspectBusy=true;ui.result=false;clearOutput();updateButtons();try{ui.workbook=await FolioEngine.parseWorkbook(ui.files);ui.inspection=FolioEngine.inspect(ui.workbook,{html:$('#html-code').value,javascript:$('#javascript-code').value},activeWorkflow);renderInspection(ui.inspection);msg('ตรวจไฟล์เสร็จแล้ว · เลือกงวดและวิธีอัปเดต','success');}catch(error){ui.inspection=null;msg(error.message||'อ่านไฟล์ไม่สำเร็จ','error');}finally{ui.inspectBusy=false;updateButtons();}}
  function renderResult(result){
    $('#html-result').value=result.html||'';$('#javascript-result').value=result.javascript||'';
    const lines=[...result.changes.map(c=>`เปลี่ยน: ${c.item} · ${c.oldValue} → ${c.newValue} · ที่มา ${c.source}`),...result.unmatched.map(x=>'จับคู่ไม่ได้: '+x)];
    $('#report-result').value=lines.length?lines.join('\n'):'ไม่มีการเปลี่ยนแปลงที่จับคู่ได้';$('#notes-result').value=(result.notes||[]).join('\n')||'ไม่มีหมายเหตุเพิ่มเติม';
    $('#result-badge').textContent=result.changes.length?'สร้างผลลัพธ์แล้ว · '+result.changes.length+' รายการ':'ไม่มีจุดที่จับคู่ได้';$('#result-badge').dataset.state=result.changes.length?'success':'warning';$('.result-card').dataset.state=result.changes.length?'complete':'incomplete';
    $$('button.copy-result,button.download-result').forEach(b=>b.disabled=!$('#'+b.dataset.target).value.trim());ui.result=true;
    if(typeof queueHtmlPreview==='function')queueHtmlPreview(true);msg(result.changes.length?'ฟังก์ชันสร้างผลลัพธ์แล้ว กรุณาตรวจรายงานก่อนนำไปใช้':'ยังไม่มีข้อมูลที่จับคู่ได้ชัดเจน · ดูหมายเหตุและตรวจรูปแบบโค้ด','success');renderGuide();$('#result-title').scrollIntoView({behavior:'smooth',block:'center'});
  }
  function run(){if(!readyToUpdate())return;ui.busy=true;updateButtons();try{const payload={html:$('#html-code').value,javascript:$('#javascript-code').value};const result=isShareholder()?FolioEngine.updateShareholders(payload,ui.inspection):FolioEngine.updateFinancial(payload,ui.inspection,selected(),mode(),targetLanguage(),periodOrder());renderResult(result);}catch(error){msg(error.message||'สร้างผลลัพธ์ไม่สำเร็จ','error');}finally{ui.busy=false;updateButtons();}}
  function saveDraft(){try{localStorage.setItem(key(),JSON.stringify({html:$('#html-code').value,javascript:$('#javascript-code').value,request:$('#request').value}));}catch{}}
  function loadDraft(){try{const draft=JSON.parse(localStorage.getItem(key())||'{}');$('#html-code').value=draft.html||'';$('#javascript-code').value=draft.javascript||'';$('#request').value=draft.request||'';}catch{}}
  function language(){const text=$('#html-code').value+' '+$('#javascript-code').value;const thai=(text.match(/[\u0E00-\u0E7F]/g)||[]).length;$('#language-detected').textContent=!text.trim()?'ภาษา: รออ่านจากโค้ด':thai?'มีข้อความภาษาไทย · คงข้อความแต่ละภาษาตามต้นฉบับ':'ภาษา: English';$('#html-char-count').textContent=$('#html-code').value.length.toLocaleString();$('#javascript-char-count').textContent=$('#javascript-code').value.length.toLocaleString();}
  function switchWorkspace(workflow){if(workflow===activeWorkflow||ui.busy||ui.inspectBusy)return;saveDraft();activeWorkflow=workflow;ui.files=[];$('#files').value='';renderFiles();invalidate();applyWorkspaceCopy();loadDraft();language();msg('สลับพื้นที่ทำงานแล้ว · เลือกไฟล์และตรวจใหม่');}
  window.addEventListener('folio:workspace-request',event=>switchWorkspace(event.detail.workflow));
  window.addEventListener('workspacechange',()=>{if(typeof htmlPreviewState!=='undefined'){htmlPreviewState.source='original';}if(typeof queueHtmlPreview==='function')queueHtmlPreview(true);});
  for(const id of ['html-code','javascript-code','request'])$('#'+id).addEventListener('input',()=>{language();saveDraft();if(id!=='request')invalidate('โค้ดเปลี่ยนแล้ว กรุณาตรวจไฟล์ใหม่ก่อนเลือกวิธีอัปเดต');if(typeof queueHtmlPreview==='function')queueHtmlPreview();});
  $('#files').addEventListener('change',()=>{ui.files=[...$('#files').files];invalidate('ไฟล์เปลี่ยนแล้ว กรุณาตรวจใหม่');renderFiles();updateButtons();});
  $('#inspect-files').addEventListener('click',inspect);$('#run-update').addEventListener('click',run);
  $('#clear').addEventListener('click',()=>{ui.files=[];$('#files').value='';$('#html-code').value='';$('#javascript-code').value='';$('#request').value='';saveDraft();renderFiles();invalidate('ล้างข้อมูลในหน้านี้แล้ว');language();});
  $$('[data-select-periods]').forEach(button=>button.addEventListener('click',()=>{const kind=button.dataset.selectPeriods;$$('input[name="display-period"]').forEach(input=>{const p=ui.inspection.displayPeriods.find(x=>x.id===input.value);input.checked=kind==='all'||(kind==='source'&&p.inSource)||(kind==='existing'&&p.inExisting);});updatePeriodSummary();clearOutput();updateButtons();}));
  $('#period-order').addEventListener('change',()=>{if(ui.inspection)renderPeriods(ui.inspection.displayPeriods||[]);clearOutput();updateButtons();});
  $$('input[name="update-mode"]').forEach(input=>input.addEventListener('change',()=>{clearOutput();updateButtons();}));
  $$('input[name="target-language"]').forEach(input=>input.addEventListener('change',()=>{ui.result=false;clearOutput();updateButtons();}));
  $('#guide-action').addEventListener('click',()=>{if(ui.result){$('#result-title').scrollIntoView({behavior:'smooth'});return;}if(ui.inspection){$('#display-selection').scrollIntoView({behavior:'smooth'});return;}if(readyToInspect())inspect();else $('#html-code').focus();});
  document.querySelectorAll('[data-guide-check]').forEach(button=>button.addEventListener('click',()=>{const map={code:'#html-code',files:'#files',inspection:'#inspect-section',periods:'#display-selection',mode:'#update-mode',connection:'#status'};$(map[button.dataset.guideCheck]||'#workspace').scrollIntoView({behavior:'smooth',block:'center'});}));
  document.querySelectorAll('[data-step-jump]').forEach(button=>button.addEventListener('click',()=>{const map={1:'#input-title',2:'#inspect-section',3:'#update-mode',4:'#result-title'};$(map[button.dataset.stepJump]).scrollIntoView({behavior:'smooth',block:'center'});}));
  $$('button.copy-result').forEach(button=>button.addEventListener('click',async()=>{try{await navigator.clipboard.writeText($('#'+button.dataset.target).value);msg('คัดลอกโค้ดแล้ว','success');}catch{msg('คัดลอกไม่สำเร็จ · เลือกข้อความในช่องผลลัพธ์แล้วคัดลอก','error');}}));
  $$('button.download-result').forEach(button=>button.addEventListener('click',()=>{const content=$('#'+button.dataset.target).value;if(!content)return;const blob=new Blob([content],{type:'text/plain;charset=utf-8'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=button.dataset.name||'folio-ir-output.txt';a.click();URL.revokeObjectURL(url);}));
  document.querySelectorAll('[data-guide-check="connection"]').forEach(b=>b.hidden=true);
  const sampleButton=document.createElement('button');sampleButton.type='button';sampleButton.className='btn secondary';sampleButton.textContent='โหลดตัวอย่าง';sampleButton.title='เติมโค้ดและ workbook ตัวอย่างเพื่อทดลอง';$('#source-files-field').querySelector('.field-heading').append(sampleButton);
  sampleButton.addEventListener('click',()=>{
    const financial=activeWorkflow==='financial-highlights';
    if(financial){
      $('#html-code').value='<section>\n  <h2>Financial Highlights</h2>\n  <table><thead><tr><th>รายการ</th><th>2025</th><th>1Q/2026</th><th>2Q/2026</th></tr></thead><tbody>\n    <tr><td>Revenue</td><td>4,500.00</td><td>1,100.00</td><td>1,250.00</td></tr>\n    <tr><td>Net profit</td><td>370.00</td><td>95.00</td><td>102.00</td></tr>\n  </tbody></table>\n</section>';
      $('#javascript-code').value='window.chart = {\n  revenue: { label: "Revenue", source: { quarterly: { period: ["1Q/2026", "2Q/2026"], dataset: [1100, 1250] } } },\n  netProfit: { label: "Net profit", source: { quarterly: { period: ["1Q/2026", "2Q/2026"], dataset: [95, 102] } } }\n};';
      const wb=XLSX.utils.book_new();const ws=XLSX.utils.aoa_to_sheet([['Financial Highlights','2025','1Q/2026','2Q/2026'],['Revenue',4500,1100,1250],['Net profit',370,95,102]]);XLSX.utils.book_append_sheet(wb,ws,'Financial Highlights');const bytes=XLSX.write(wb,{bookType:'xlsx',type:'array'});ui.files=[new File([bytes],'Folio-IR-Financial-Sample.xlsx',{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})];
    }else{
      $('#html-code').value='<section><h2>Major Shareholders</h2><table><thead><tr><th>Shareholder</th><th>Shares</th><th>Percent</th></tr></thead><tbody><tr><td>Siam Strategic Fund</td><td>400,000,000</td><td>40.00%</td></tr><tr><td>Northstar Asset Management</td><td>200,000,000</td><td>20.00%</td></tr></tbody></table></section>';
      $('#javascript-code').value='window.chart = { majorShareholders: { labels: ["Siam Strategic Fund", "Northstar Asset Management", "Others"], dataset: [40, 20, 40] } };';
      const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Shareholder','Shares','Percent'],['Siam Strategic Fund',400000000,40],['Northstar Asset Management',200000000,20],['Others',400000000,40]]),'Graph');const bytes=XLSX.write(wb,{bookType:'xlsx',type:'array'});ui.files=[new File([bytes],'Folio-IR-Shareholders-Sample.xlsx',{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})];
    }
    $('#files').value='';renderFiles();language();saveDraft();invalidate('ตัวอย่างพร้อมแล้ว · กดตรวจไฟล์ก่อนดำเนินการ');if(typeof queueHtmlPreview==='function')queueHtmlPreview(true);
  });
  loadDraft();applyWorkspaceCopy();language();renderFiles();clearOutput();updateButtons();msg('พร้อมทำงาน · เลือกไฟล์และวางโค้ด หรือกด “โหลดตัวอย่าง”','success');
})();
