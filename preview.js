'use strict';
const htmlPreviewState = { source: 'original', width: 'desktop', lastHtml: null, timer: null };
const jsPreviewState = { source: 'original', timer: null };
const previewBaseStyle = `
  *{box-sizing:border-box}html{color-scheme:light}body{margin:0;padding:28px;background:#fff;color:#24364b;font:14px/1.65 "Segoe UI",Tahoma,sans-serif;overflow-wrap:anywhere}
  h1,h2,h3,h4{color:#14283f;line-height:1.4;margin:0 0 16px}h1{font-size:28px}h2{font-size:24px}h3{font-size:19px}p{margin:0 0 16px}
  .container,.container-fluid{width:100%;max-width:1160px;margin:0 auto}.row{display:flex;flex-wrap:wrap;margin:0 -10px}.row>*{padding:0 10px;min-width:0;flex:1 1 100%}
  table{border-collapse:collapse;width:100%;margin:16px 0 24px;font-size:13px}th,td{padding:12px 14px;border-bottom:1px solid #e4e9ef;text-align:left}th{background:#f3f6f9;font-weight:600}thead th{color:#fff;background:#243d57}tbody tr:nth-child(even){background:#f8fafb}
  .table-responsive{overflow-x:auto}.text-center{text-align:center}.text-right,.text-end{text-align:right}.text-left,.text-start{text-align:left}.d-flex{display:flex}.justify-content-center{justify-content:center}.align-items-center{align-items:center}
  img,svg{max-width:100%;height:auto}.card{border:1px solid #e0e7ed;border-radius:12px;margin-bottom:16px;padding:20px}.section{padding:12px 0}.mb-0{margin-bottom:0}.mb-3{margin-bottom:16px}.mb-4{margin-bottom:24px}.fw-bold,.font-weight-bold{font-weight:700}
  financial-chart,total-assets{display:flex;min-height:90px;align-items:center;justify-content:center;background:#f6f8fa;border:1px dashed #d7e0e7;border-radius:8px;margin-top:12px}financial-chart:empty::before,total-assets:empty::before{content:"กราฟนี้ต้องใช้ JavaScript ของเว็บไซต์";font-size:11px;color:#8293a3}
  [class*="col-"]{flex:1 1 0}.col-12{flex:0 0 100%}.col-6{flex:0 0 50%}.col-4{flex:0 0 33.333%}.col-3{flex:0 0 25%}
  @media(min-width:768px){.col-md-6{flex:0 0 50%}.col-md-4{flex:0 0 33.333%}.col-md-3{flex:0 0 25%}.col-md-12{flex:0 0 100%}}
  @media(min-width:992px){.col-lg-6{flex:0 0 50%}.col-lg-4{flex:0 0 33.333%}.col-lg-3{flex:0 0 25%}.col-lg-12{flex:0 0 100%}}
  @media(max-width:480px){body{padding:18px}th,td{padding:9px;font-size:12px}h1{font-size:24px}h2{font-size:21px}}
`;
function queueHtmlPreview(force = false) {
  clearTimeout(htmlPreviewState.timer);
  htmlPreviewState.timer = setTimeout(() => renderHtmlPreview(force), force ? 0 : 220);
}
function prepareHtmlPreview(html) {
  const template = document.createElement('template');
  template.innerHTML = html;
  function unwrapTemplates(fragment) {
    fragment.querySelectorAll('template').forEach(element => {
      unwrapTemplates(element.content);
      element.replaceWith(element.content);
    });
  }
  // Browsers keep <template> content inert. Unwrap only this disposable copy,
  // including nested slots, before sanitizing its complete element tree.
  unwrapTemplates(template.content);
  template.content.querySelectorAll('script,iframe,frame,frameset,object,embed,base,link,meta,portal').forEach(element => element.remove());
  template.content.querySelectorAll('*').forEach(element => {
    for (const attribute of [...element.attributes]) {
      const name = attribute.name.toLowerCase();
      // Vue normally removes v-cloak after mounting. Static previews do not mount Vue.
      if (/^on/i.test(name) || ['v-cloak','href','xlink:href','action','formaction','srcdoc','autofocus','target','ping'].includes(name)) element.removeAttribute(attribute.name);
    }
  });
  const visibleContent = template.content.cloneNode(true);
  visibleContent.querySelectorAll('style,title').forEach(element => element.remove());
  const hasContent = Boolean(visibleContent.textContent.trim() || visibleContent.querySelector('img,svg,canvas,table,hr,input,textarea,select,button,[style]') || [...visibleContent.querySelectorAll('*')].some(element => element.localName.includes('-')));
  return { markup:template.innerHTML, hasContent };
}
function showPreviewEmpty(title, detail) {
  const empty = document.querySelector('#preview-empty');
  empty.querySelector('strong').textContent = title;
  empty.querySelector('p').textContent = detail;
  empty.hidden = false;
  document.querySelector('#html-preview').hidden = true;
}
function renderHtmlPreview(force = false) {
  const get = selector => document.querySelector(selector);
  const resultAvailable = Boolean(get('#html-result').value.trim());
  if (!resultAvailable && htmlPreviewState.source === 'result') htmlPreviewState.source = 'original';
  document.querySelectorAll('[data-preview-source]').forEach(button => {
    button.setAttribute('aria-pressed', String(button.dataset.previewSource === htmlPreviewState.source));
    if (button.dataset.previewSource === 'result') button.disabled = !resultAvailable;
  });
  if (get('#output-preview-panel').hidden) return;
  const html = get(htmlPreviewState.source === 'result' ? '#html-result' : '#html-code').value;
  const iframe = get('#html-preview');
  get('#preview-document-name').textContent = 'ตัวอย่าง HTML · ' + (htmlPreviewState.source === 'result' ? 'หลังอัปเดต' : 'ต้นฉบับ');
  if (html === htmlPreviewState.lastHtml && !force) return;
  if (!html.trim()) {
    showPreviewEmpty('วาง HTML เพื่อดูตัวอย่าง', 'ใส่โค้ดในช่อง HTML / Vue ต้นฉบับ ตัวอย่างจะแสดงอัตโนมัติ · JavaScript เพียงอย่างเดียวยังไม่มีหน้า HTML ให้แสดง');
    iframe.removeAttribute('srcdoc');
    htmlPreviewState.lastHtml = html;
    return;
  }
  // Work on an inert, disposable copy. Never write serialized preview markup
  // back to a source/result field, the draft, or a download.
  try {
    const prepared = prepareHtmlPreview(html);
    if (!prepared.hasContent) {
      showPreviewEmpty('ไม่พบเนื้อหา HTML สำหรับแสดงผล', 'โค้ดนี้มีเฉพาะสไตล์หรือสคริปต์ ให้ใส่ส่วนตารางหรือเนื้อหาของหน้าในช่อง HTML / Vue ด้วย');
      iframe.removeAttribute('srcdoc');
    } else {
      const policy = "default-src 'none'; script-src 'none'; style-src 'unsafe-inline'; img-src data:; font-src 'none'; connect-src 'none'; frame-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'; media-src 'none'";
      iframe.srcdoc = '<!doctype html><html lang="th"><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="' + policy + '"><meta name="viewport" content="width=device-width,initial-scale=1"><style>' + previewBaseStyle + '</style></head><body>' + prepared.markup + '</body></html>';
      get('#preview-empty').hidden = true;
      iframe.hidden = false;
    }
    htmlPreviewState.lastHtml = html;
  } catch {
    showPreviewEmpty('ยังสร้างตัวอย่างจากโค้ดนี้ไม่ได้', 'ตรวจว่าใส่ HTML / Vue ครบถ้วน แล้วกดรีเฟรชตัวอย่างอีกครั้ง โค้ดต้นฉบับยังคงเดิม');
  }
}
document.querySelectorAll('[data-preview-source]').forEach(button => button.addEventListener('click', () => {
  htmlPreviewState.source = button.dataset.previewSource; renderHtmlPreview();
}));
document.querySelectorAll('[data-preview-width]').forEach(button => button.addEventListener('click', () => {
  htmlPreviewState.width = button.dataset.previewWidth;
  document.querySelector('#preview-stage').dataset.width = htmlPreviewState.width;
  document.querySelectorAll('[data-preview-width]').forEach(item => item.setAttribute('aria-pressed', String(item === button)));
}));
document.querySelector('#refresh-preview').addEventListener('click', () => renderHtmlPreview(true));
document.querySelector('.preview-card').addEventListener('toggle', event => {
  if (event.currentTarget.open) queueHtmlPreview(true);
});
function queueJsDataPreview(){clearTimeout(jsPreviewState.timer);jsPreviewState.timer=setTimeout(renderJsDataPreview,120);}
function renderJsDataPreview(){
  const original=document.querySelector('#javascript-code').value,result=document.querySelector('#javascript-result').value,available=Boolean(result.trim());
  if(!available&&jsPreviewState.source==='result')jsPreviewState.source='original';
  document.querySelectorAll('[data-js-preview-source]').forEach(button=>{button.setAttribute('aria-pressed',String(button.dataset.jsPreviewSource===jsPreviewState.source));if(button.dataset.jsPreviewSource==='result')button.disabled=!available;});
  const code=jsPreviewState.source==='result'?result:original,content=document.querySelector('#js-preview-content'),count=document.querySelector('#js-preview-count');
  content.replaceChildren();
  const charts=window.FolioEngine?.parseChartPreview(code)||[];
  count.textContent=charts.length?`${charts.length} ชุดข้อมูล`:'ไม่มีชุดข้อมูลที่รู้จัก';
  if(!charts.length){const empty=document.createElement('div');empty.className='preview-empty js-preview-empty';const title=document.createElement('strong'),detail=document.createElement('p');title.textContent=code.trim()?'ไม่พบ array period/dataset ที่รองรับ':'วาง JavaScript เพื่อดูข้อมูลกราฟ';detail.textContent=code.trim()?'ต้องมี array period และ dataset ใน object เดียวกัน':'ตัวอย่างจะแสดงค่า period และ dataset จากโค้ดที่พบ';empty.append(title,detail);content.append(empty);return;}
  charts.forEach(chart=>{
    const section=document.createElement('section');section.className='js-chart-preview';const heading=document.createElement('div');heading.className='js-chart-preview-heading';const title=document.createElement('strong'),meta=document.createElement('span');title.textContent=chart.title;meta.textContent=`${chart.key} · ${chart.section}`;heading.append(title,meta);
    const table=document.createElement('table');table.className='js-data-preview-table';const thead=document.createElement('thead'),headRow=document.createElement('tr');for(const label of ['Period','Dataset']){const th=document.createElement('th');th.textContent=label;headRow.append(th);}thead.append(headRow);const tbody=document.createElement('tbody');
    chart.rows.forEach(row=>{const tr=document.createElement('tr'),period=document.createElement('td'),value=document.createElement('td');period.textContent=row.period;value.textContent=row.value;value.className='js-data-value';tr.append(period,value);tbody.append(tr);});table.append(thead,tbody);section.append(heading,table);content.append(section);
  });
}
document.querySelectorAll('[data-js-preview-source]').forEach(button=>button.addEventListener('click',()=>{jsPreviewState.source=button.dataset.jsPreviewSource;renderJsDataPreview();}));
document.querySelector('.js-preview-card').addEventListener('toggle',event=>{if(event.currentTarget.open)queueJsDataPreview();});
document.querySelector('#javascript-code').addEventListener('input',queueJsDataPreview);
document.querySelector('#refresh-preview').addEventListener('click',queueJsDataPreview);
queueJsDataPreview();
