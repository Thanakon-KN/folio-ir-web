'use strict';
(() => {
  const palettes = {peach:'#efd5c6','rose-pastel':'#efd3df',mint:'#c6e5dc',sky:'#cde3f4',lavender:'#dcd7f2',sage:'#dce5cb',butter:'#f4e7b9',apricot:'#f5d8b9',lilac:'#e6d3eb',aqua:'#c7e6e7',sand:'#e7ddcb',periwinkle:'#d2dbf5'};
  const names = {peach:'พีช','rose-pastel':'โรส',mint:'มิ้นต์',sky:'สกาย',lavender:'ลาเวนเดอร์',sage:'เซจ',butter:'บัตเตอร์',apricot:'แอปริคอต',lilac:'ไลแลค',aqua:'อควา',sand:'แซนด์',periwinkle:'เพอริวิงเคิล'};
  const storageKey='folio-ir-appearance-v2';
  const defaults={color:palettes.mint,financial:palettes.mint,shareholders:palettes.lavender,separate:false,mode:'light',intensity:50,editor:'ink',followColor:true,uiSize:16,codeSize:14,section:'screen'};
  let workspace=document.documentElement.dataset.workflow||'financial-highlights',settings={...defaults};
  const system=matchMedia('(prefers-color-scheme: dark)'),$=selector=>document.querySelector(selector);
  const hex=value=>typeof value==='string'&&/^#[0-9a-f]{6}$/i.test(value);
  function restore(saved){
    if(!saved||typeof saved!=='object')return;
    for(const key of ['color','financial','shareholders'])if(hex(saved[key]))settings[key]=saved[key];
    for(const key of ['separate','followColor'])if(typeof saved[key]==='boolean')settings[key]=saved[key];
    for(const [key,values]of [['mode',['light','dark','system']],['editor',['ink','paper']],['section',['screen','code','type']]])if(values.includes(saved[key]))settings[key]=saved[key];
    for(const [key,min,max]of [['intensity',0,100],['uiSize',14,20],['codeSize',12,20]])if(typeof saved[key]==='number'&&Number.isFinite(saved[key]))settings[key]=Math.min(max,Math.max(min,saved[key]));
  }
  try{
    const saved=JSON.parse(localStorage.getItem(storageKey)||'null');
    if(saved)restore(saved);
    else{
      // Migrate appearance only. Drafts and code content use separate storage.
      const financial=JSON.parse(localStorage.getItem('folio-ir-appearance-financial-highlights')||'null');
      const shareholders=JSON.parse(localStorage.getItem('folio-ir-appearance-major-shareholders')||'null');
      if(financial||shareholders){const previous=financial||shareholders;restore(previous);settings.financial=hex(financial?.color)?financial.color:defaults.financial;settings.shareholders=hex(shareholders?.color)?shareholders.color:defaults.shareholders;settings.color=hex(previous.color)?previous.color:defaults.color;settings.separate=true;settings.uiSize=previous.textSize==='large'?18:16;settings.codeSize=previous.textSize==='large'?16:14;}
    }
  }catch{}
  const rgb=color=>[1,3,5].map(index=>parseInt(color.slice(index,index+2),16));
  function mix(first,second,amount){const a=rgb(first),b=rgb(second);return '#'+a.map((value,index)=>Math.round(value*(1-amount)+b[index]*amount).toString(16).padStart(2,'0')).join('');}
  function hueOf(color){const [r,g,b]=rgb(color).map(value=>value/255),max=Math.max(r,g,b),min=Math.min(r,g,b),delta=max-min;let hue=0;if(delta)hue=max===r?60*((g-b)/delta%6):max===g?60*((b-r)/delta+2):60*((r-g)/delta+4);return Math.round((hue+360)%360);}
  function luminance(color){const c=rgb(color).map(value=>{value/=255;return value<=.04045?value/12.92:((value+.055)/1.055)**2.4;});return c[0]*.2126+c[1]*.7152+c[2]*.0722;}
  function readable(color,background,minimum=4.5){const destination=luminance(background)>.4?'#000000':'#ffffff';for(let index=0;index<24;index++){const a=luminance(color),b=luminance(background);if((Math.max(a,b)+.05)/(Math.min(a,b)+.05)>=minimum)break;color=mix(color,destination,.13);}return color;}
  const workspaceKey=()=>workspace==='major-shareholders'?'shareholders':'financial';
  const colorKey=()=>settings.separate?workspaceKey():'color';
  function apply(save=true){
    const root=document.documentElement,color=settings[colorKey()],hue=hueOf(color),seed=mix(color,'#ffffff',.35-settings.intensity*.0035);
    const dark=settings.mode==='dark'||settings.mode==='system'&&system.matches,surface=dark?'#232c31':'#ffffff',ink=dark?'#eef4f2':'#273b37';
    const soft=dark?mix(color,'#20272b',.86):mix(seed,'#ffffff',.6-settings.intensity*.003);
    const hsl=(h,s,l)=>`hsl(${h} ${s}% ${l}%)`,codeHue=settings.followColor?hue:210,codeDark=settings.editor==='ink';
    const hover=readable(mix(color,'#16262b',.82),'#ffffff',7),accent=readable(color,surface),brand=dark?seed:readable(mix(color,'#18312d',.78),'#ffffff',7);
    const vars={
      '--paper':dark?mix(color,'#171d21',.97):mix(seed,'#ffffff',.94),'--surface':surface,'--surface-soft':dark?mix(color,'#243137',.94):mix(seed,'#ffffff',.89),'--surface-hover':soft,
      '--ink':ink,'--muted':dark?'#bdcbc6':'#5c7069','--faint':dark?'#b1c3bb':'#63766e','--line':dark?mix(color,'#313e43',.87):mix(seed,'#dce5e2',.6),'--sidebar':soft,
      '--brand':brand,'--brand-ink':dark?'#263b37':'#ffffff','--brand-hover':hover,'--brand-hover-ink':'#ffffff','--accent':accent,'--accent-ink':dark?'#192a24':'#ffffff','--accent-bright':accent,'--accent-soft':soft,
      '--accent-line':dark?mix(color,'#344449',.8):mix(seed,'#dce5e2',.4),'--hover-bg':hover,'--hover-ink':'#ffffff','--hover-line':hover,'--folio-note':soft,'--folio-note-ink':ink,
      '--studio-seed':seed,'--studio-ink':readable('#273b37',seed),'--read-scale':String(settings.uiSize/16),'--editor-font-size':settings.codeSize+'px',
      '--editor-bg':hsl(codeHue,28,codeDark?13:96),'--editor-ink':hsl(codeHue,22,codeDark?96:15),'--editor-placeholder':hsl(codeHue,18,codeDark?79:35),'--editor-border':hsl(codeHue,18,codeDark?40:65),
      '--editor-head':hsl(codeHue,23,codeDark?20:90),'--editor-head-ink':hsl(codeHue,22,codeDark?96:15),'--editor-thumb':hsl(codeHue,20,codeDark?65:40),'--editor-selection':hsl(codeHue,35,80),'--editor-selection-ink':hsl(codeHue,30,15),
      '--editor-tag':hsl((codeHue+12)%360,codeDark?60:47,codeDark?83:29),'--editor-string':hsl(codeHue,codeDark?52:45,codeDark?81:25)
    };
    root.dataset.appearance=dark?'dark':'light';root.dataset.editorAppearance=settings.editor;root.dataset.palette=Object.keys(palettes).find(key=>palettes[key]===color)||'custom';
    Object.entries(vars).forEach(([key,value])=>root.style.setProperty(key,value));
    document.querySelectorAll('[data-theme-preset]').forEach(button=>{button.style.setProperty('--swatch',palettes[button.dataset.themePreset]);button.setAttribute('aria-pressed',String(palettes[button.dataset.themePreset]===color));});
    for(const [attribute,value]of [['appearance',settings.mode],['editor-appearance',settings.editor],['settings-tab',settings.section]])document.querySelectorAll('[data-'+attribute+']').forEach(button=>button.setAttribute(attribute==='settings-tab'?'aria-selected':'aria-pressed',String(button.getAttribute('data-'+attribute)===value)));
    document.querySelectorAll('[data-settings-panel]').forEach(panel=>panel.hidden=panel.dataset.settingsPanel!==settings.section);
    $('#theme-color').value=color;$('#theme-color-label').textContent=color.toUpperCase();$('#theme-current-color').textContent=names[root.dataset.palette]||'กำหนดเอง';
    $('#theme-intensity').value=settings.intensity;$('#theme-intensity-value').textContent=settings.intensity+'%';$('#theme-separate').checked=settings.separate;$('#theme-follow-color').checked=settings.followColor;
    $('#theme-ui-size').value=settings.uiSize;$('#theme-ui-size-value').textContent=settings.uiSize+' px';$('#theme-code-size').value=settings.codeSize;$('#theme-code-size-value').textContent=settings.codeSize+' px';
    $('meta[name="theme-color"]').content=vars['--paper'];if(save)try{localStorage.setItem(storageKey,JSON.stringify(settings));}catch{}
  }
  document.querySelectorAll('[data-theme-preset]').forEach(button=>button.addEventListener('click',()=>{settings[colorKey()]=palettes[button.dataset.themePreset];apply();}));
  document.querySelectorAll('[data-appearance]').forEach(button=>button.addEventListener('click',()=>{settings.mode=button.dataset.appearance;apply();}));
  document.querySelectorAll('[data-editor-appearance]').forEach(button=>button.addEventListener('click',()=>{settings.editor=button.dataset.editorAppearance;apply();}));
  const themeTabs=[...document.querySelectorAll('[data-settings-tab]')];
  themeTabs.forEach((button,index)=>{button.addEventListener('click',()=>{settings.section=button.dataset.settingsTab;apply();});button.addEventListener('keydown',event=>{let next;if(event.key==='ArrowRight')next=(index+1)%themeTabs.length;if(event.key==='ArrowLeft')next=(index+themeTabs.length-1)%themeTabs.length;if(event.key==='Home')next=0;if(event.key==='End')next=themeTabs.length-1;if(next===undefined)return;event.preventDefault();themeTabs[next].focus();themeTabs[next].click();});});
  for(const [id,key]of [['theme-intensity','intensity'],['theme-ui-size','uiSize'],['theme-code-size','codeSize']])$('#'+id).addEventListener('input',event=>{settings[key]=Number(event.target.value);apply();});
  $('#theme-separate').addEventListener('change',event=>{if(event.target.checked)settings[workspaceKey()]=settings.color;else settings.color=settings[workspaceKey()];settings.separate=event.target.checked;apply();});
  $('#theme-follow-color').addEventListener('change',event=>{settings.followColor=event.target.checked;apply();});
  $('#theme-color').addEventListener('input',event=>{if(hex(event.target.value)){settings[colorKey()]=event.target.value;apply();}});
  $('#theme-reset').addEventListener('click',()=>{settings={...defaults,section:settings.section};apply();});
  function closeMenu(){const menu=$('#theme-menu');menu.open=false;menu.querySelector('summary').focus();}
  $('#theme-close').addEventListener('click',closeMenu);
  document.addEventListener('pointerdown',event=>{const menu=$('#theme-menu');if(!menu.contains(event.target))menu.open=false;});
  document.addEventListener('keydown',event=>{if(event.key==='Escape'&&$('#theme-menu').open)closeMenu();});
  system.addEventListener('change',()=>{if(settings.mode==='system')apply(false);});
  window.addEventListener('workspacechange',event=>{workspace=event.detail.workflow;apply(false);});
  window.addEventListener('storage',event=>{if(event.key===storageKey){try{settings={...defaults};restore(JSON.parse(event.newValue));apply(false);}catch{}}});
  apply(false);
})();
