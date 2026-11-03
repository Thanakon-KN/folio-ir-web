'use strict';
let activeWorkflow='financial-highlights';
const isShareholder=()=>activeWorkflow==='major-shareholders';
const workspaceCopy={
  'financial-highlights':{name:'Financial Highlights',skill:'ฟังก์ชัน Financial Highlights',title:'อัปเดตข้อมูลการเงิน',description:'อ่าน workbook จับคู่รายการและงวด แล้วสร้างโค้ดด้วยกฎที่ตรวจสอบได้',html:'วาง HTML / Vue ของตาราง Financial Highlights ที่นี่…',javascript:'วาง window.chart หรือ JavaScript ของกราฟที่นี่…',upload:'เพิ่มไฟล์ข้อมูลการเงิน',step:'ตรวจและเลือกงวด',stepDetail:'กำหนดข้อมูลที่จะแสดง',inspect:'ตรวจไฟล์และเปรียบเทียบงวด',request:'หมายเหตุสำหรับรายงานเท่านั้น · ฟังก์ชันไม่ตีความคำสั่งอิสระ'},
  'major-shareholders':{name:'Major Shareholders',skill:'ฟังก์ชัน Major Shareholders',title:'อัปเดตข้อมูลผู้ถือหุ้น',description:'อ่าน Overview, Free float, Graph และรายชื่อจาก workbook แล้วอัปเดตส่วนที่จับคู่ได้',html:'วาง HTML / Vue ของหน้าข้อมูลผู้ถือหุ้นที่นี่…',javascript:'วาง window.chart.majorShareholders หรือ JavaScript ของกราฟที่นี่…',upload:'เพิ่มไฟล์ข้อมูลผู้ถือหุ้น',step:'ตรวจไฟล์ต้นทาง',stepDetail:'ภาพรวม รายชื่อ และกราฟ',inspect:'ตรวจไฟล์ข้อมูลผู้ถือหุ้น',request:'หมายเหตุสำหรับรายงานเท่านั้น · ฟังก์ชันไม่ตีความคำสั่งอิสระ'}
};
function applyWorkspaceCopy(){
  const copy=workspaceCopy[activeWorkflow];document.documentElement.dataset.workflow=activeWorkflow;
  document.querySelector('#workspace-panel').setAttribute('aria-labelledby',isShareholder()?'tab-shareholders':'tab-financial');
  document.querySelectorAll('[data-workflow-tab]').forEach(button=>{const active=button.dataset.workflowTab===activeWorkflow;button.setAttribute('aria-selected',String(active));button.tabIndex=active?0:-1;});
  document.querySelector('#workspace-name').textContent=copy.name;document.querySelector('#hero-title').textContent=copy.title;document.querySelector('#hero-description').textContent=copy.description;document.querySelector('#hero-eyebrow').textContent=copy.name.toUpperCase();
  document.querySelector('#active-skill').textContent=copy.skill;document.querySelector('#html-code').placeholder=copy.html;document.querySelector('#javascript-code').placeholder=copy.javascript;document.querySelector('#upload-title').textContent=copy.upload;
  document.querySelector('#step-inspect-title').textContent=copy.step;document.querySelector('#step-inspect-detail').textContent=copy.stepDetail;document.querySelector('#step-update-detail').textContent=isShareholder()?'อัปเดตตามข้อมูลต้นทาง':'เลือกเพิ่มหรือแทนที่';document.querySelector('#inspect-button-label').textContent=copy.inspect;
  document.querySelector('#shareholder-scope').hidden=!isShareholder();document.querySelector('#request').placeholder=copy.request;
  document.querySelector('.download-result[data-target="html-result"]').dataset.name=isShareholder()?'updated-shareholders.html':'updated-ir-page.html';document.querySelector('.download-result[data-target="javascript-result"]').dataset.name=isShareholder()?'updated-shareholders-chart.js':'updated-ir-chart.js';document.title='Folio IR · '+copy.name;
  window.dispatchEvent(new CustomEvent('workspacechange',{detail:{workflow:activeWorkflow}}));
}
document.querySelectorAll('[data-workflow-tab]').forEach(button=>{
  button.addEventListener('click',()=>window.dispatchEvent(new CustomEvent('folio:workspace-request',{detail:{workflow:button.dataset.workflowTab}})));
  button.addEventListener('keydown',event=>{if(!['ArrowLeft','ArrowRight','Home','End'].includes(event.key))return;event.preventDefault();const tabs=[...document.querySelectorAll('[data-workflow-tab]')],i=event.key==='Home'?0:event.key==='End'?tabs.length-1:(tabs.indexOf(button)+(event.key==='ArrowRight'?1:-1)+tabs.length)%tabs.length;tabs[i].focus();tabs[i].click();});
});
