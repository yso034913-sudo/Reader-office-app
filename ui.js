// Reader All — UI glue. Loaded after app.js (uses its globals: S, sync, mkHl, hs, tx, dbPut, home, openDoc).
(()=>{
'use strict';
const $=s=>document.querySelector(s),E=(t,c,h)=>{const e=document.createElement(t);if(c)e.className=c;if(h!=null)e.innerHTML=h;return e};

// ---------- hooks that app.js expects ----------
window.LS={r(){},d(){},m(){},u(){}};        // lasso is not part of this build
window.SN=p=>p;                              // no grid snapping
window.raBusy=(on,t)=>{let b=$('#busy');if(!on){b&&b.remove();return}
 if(!b){b=E('div','','<div class="bx"><i class="spin"></i><span></span></div>');b.id='busy';document.body.append(b)}
 b.querySelector('span').textContent=t||''};
const toast=t=>{const d=E('div','oftoast');d.textContent=t;document.body.append(d);setTimeout(()=>d.remove(),2200)};

// ---------- home ----------
const tabOrder=['ppt','doc','xls','settings'];
const tabPaint=()=>{const t=raGetTab();document.querySelectorAll('#hnav button').forEach(b=>{const on=b.dataset.tab===t;b.classList.toggle('on',on);b.setAttribute('aria-selected',on)});document.querySelectorAll('.tp').forEach(p=>p.hidden=p.dataset.tab!==t);const isSettings=t==='settings';$('#list').hidden=isSettings;['#ls','#so','#tgs'].forEach(q=>{const e=$(q);if(e)e.hidden=isSettings});};
const switchHomeTab=t=>{if(!tabOrder.includes(t)||t===raGetTab())return;raSetTab(t);S.tg='';tabPaint();home();tabPaint();$('#home').scrollTop=0};
document.querySelectorAll('#hnav button').forEach(b=>b.onclick=()=>switchHomeTab(b.dataset.tab));
// Swipe horizontally across the home screen to move between PowerPoint, Word and Excel.
// Vertical scrolling and taps on controls remain native and unaffected.
const homeEl=$('#home');let homeTouch=null;
homeEl.addEventListener('touchstart',e=>{
 if(e.touches.length!==1)return;
 const target=e.target;
 if(target.closest('button,label,input,select,textarea,a,[contenteditable="true"]')){homeTouch=null;return;}
 const t=e.touches[0];homeTouch={x:t.clientX,y:t.clientY};
},{passive:true});
homeEl.addEventListener('touchend',e=>{
 if(!homeTouch||e.changedTouches.length!==1)return;
 const t=e.changedTouches[0],dx=t.clientX-homeTouch.x,dy=t.clientY-homeTouch.y;homeTouch=null;
 if(Math.abs(dx)<55||Math.abs(dx)<Math.abs(dy)*1.25)return;
 const current=tabOrder.indexOf(raGetTab());if(current<0)return;
 const next=dx<0?Math.min(tabOrder.length-1,current+1):Math.max(0,current-1);
 if(next!==current)switchHomeTab(tabOrder[next]);
},{passive:true});
homeEl.addEventListener('touchcancel',()=>{homeTouch=null},{passive:true});
tabPaint();

// ---------- app-wide settings ----------
const PREF='readerall.settings.v1';
const defaults={theme:'dark',eye:false,eyeLevel:20,motion:false,font:16,deleteConfirm:true};
const readPrefs=()=>{try{return {...defaults,...JSON.parse(localStorage.getItem(PREF)||'{}')}}catch(e){return {...defaults}}};
let prefs=readPrefs();
const savePrefs=()=>{try{localStorage.setItem(PREF,JSON.stringify(prefs))}catch(e){toast('تعذر حفظ الإعدادات على هذا الجهاز')}};
const applyPrefs=()=>{
 const systemDark=()=>!!(window.matchMedia&&matchMedia('(prefers-color-scheme: dark)').matches);
 const theme=prefs.theme==='system'?(systemDark()?'dark':'light'):prefs.theme;
 document.documentElement.dataset.theme=theme;
 document.documentElement.style.setProperty('--eye-alpha',prefs.eye?String(prefs.eyeLevel/100):'0');
 document.body.classList.toggle('eye-protect',!!prefs.eye);
 document.body.classList.toggle('reduce-motion',!!prefs.motion);
 document.documentElement.style.setProperty('--app-font',prefs.font+'px');
 const meta=document.querySelector('meta[name="theme-color"]');if(meta)meta.content=theme==='dark'?'#0f1523':'#f4f6fb';
};
const bindSetting=(id,key,convert=v=>v)=>{const el=$('#'+id);if(!el)return;el.value=String(prefs[key]);el.checked=!!prefs[key];el.addEventListener('change',()=>{prefs[key]=convert(el.type==='checkbox'?el.checked:el.value);savePrefs();applyPrefs();if(id==='set-eye')$('#eye-level-row').hidden=!prefs.eye})};
bindSetting('set-theme','theme',v=>v);bindSetting('set-eye','eye',Boolean);bindSetting('set-eye-level','eyeLevel',Number);bindSetting('set-motion','motion',Boolean);bindSetting('set-font','font',Number);bindSetting('set-delete-confirm','deleteConfirm',Boolean);
$('#eye-level-row').hidden=!prefs.eye;applyPrefs();
if(window.matchMedia){const mq=matchMedia('(prefers-color-scheme: dark)');mq.addEventListener&&mq.addEventListener('change',()=>{if(prefs.theme==='system')applyPrefs()})}
$('#set-reset').onclick=async()=>{if(!await raConfirm('إعادة جميع إعدادات التطبيق إلى الوضع الافتراضي؟'))return;prefs={...defaults};savePrefs();['set-theme','set-eye','set-eye-level','set-motion','set-font','set-delete-confirm'].forEach(id=>{const el=$('#'+id);if(el){if(el.type==='checkbox')el.checked=!!prefs[({ 'set-eye':'eye','set-motion':'motion','set-delete-confirm':'deleteConfirm'})[id]];else el.value=String(prefs[({'set-theme':'theme','set-eye-level':'eyeLevel','set-font':'font'})[id]])}});$('#eye-level-row').hidden=true;applyPrefs();toast('تمت إعادة ضبط الإعدادات')};
// Unified file manager inside Settings.
const fileList=$('#set-file-list');
const renderSettingsFiles=async()=>{
 if(!fileList)return;fileList.hidden=false;fileList.textContent='جارٍ تحميل الملفات…';
 try{const files=await window.raListAllFiles();fileList.innerHTML='';
  if(!files.length){fileList.textContent='لا توجد ملفات محفوظة حاليًا.';return}
  files.forEach(f=>{const row=document.createElement('div');row.className='settings-file-row';
   const info=document.createElement('div'),name=document.createElement('b'),meta=document.createElement('small');name.textContent=f.n;meta.textContent=({doc:'Word',xls:'Excel',ppt:'PowerPoint',note:'ملاحظة',pdf:'PDF'}[f.k]||'ملف')+(f.lk?' · 🔒':'');info.append(name,meta);
   const open=document.createElement('button');open.className='btn';open.textContent='فتح';open.onclick=()=>window.raOpenStoredFile(f.id);
   const rename=document.createElement('button');rename.className='btn';rename.textContent='✏️';rename.setAttribute('aria-label','إعادة تسمية');rename.onclick=async()=>{const n=prompt('الاسم الجديد',f.n);if(n&&n.trim()){await window.raRenameStoredFile(f.id,n.trim().slice(0,120));renderSettingsFiles();refreshStorage()}};
   const del=document.createElement('button');del.className='btn danger';del.textContent='حذف';del.onclick=async()=>{if(f.lk&&window.raVerify&&!await window.raVerify({title:'الملف مقفول 🔒',sub:f.n}))return;if(!await raConfirm('حذف «'+f.n+'» نهائيًا؟'))return;await window.raDeleteStoredFile(f.id);renderSettingsFiles();refreshStorage()};
   row.append(info,open,rename,del);fileList.append(row);
  });
 }catch(e){fileList.textContent='تعذر تحميل قائمة الملفات.'}
};
$('#set-file-manager').onclick=()=>{if(fileList.hidden||!fileList.childElementCount)renderSettingsFiles();else fileList.hidden=true};
const humanBytes=n=>n<1024?Math.round(n)+' B':n<1048576?(n/1024).toFixed(1)+' KB':(n/1048576).toFixed(2)+' MB';
const refreshStorage=async()=>{try{const x=await window.raStorageStats();$('#storage-count').textContent=x.count+' ملف';$('#storage-size').textContent=humanBytes(x.bytes);const quota=x.quota||0,used=x.usage||x.bytes;const pct=quota?Math.min(100,used/quota*100):0;$('#storage-meter-fill').style.width=pct+'%';$('#storage-detail').textContent=quota?'المساحة المقدرة المستخدمة من مساحة الموقع: '+humanBytes(used)+' من '+humanBytes(quota)+'.':'الحجم المحسوب لملفات التطبيق فقط؛ تقدير الحصة غير متاح على هذا الجهاز.'}catch(e){$('#storage-detail').textContent='تعذر قراءة معلومات التخزين.'}};
$('#set-storage-refresh').onclick=refreshStorage;refreshStorage();
const goTab=t=>document.querySelector('#hnav [data-tab="'+t+'"]').click();
$('#set-convert-word').onclick=()=>goTab('doc');$('#set-convert-excel').onclick=()=>goTab('xls');$('#set-convert-ppt').onclick=()=>goTab('ppt');

$('#set-backup').onclick=()=>{$('#hmore').hidden=false;$('#bkx').click()};
window.raAppSettings=()=>({...prefs});
$('#hmn').onclick=()=>{const m=$('#hmore');m.hidden=!m.hidden};

// ---------- editor: tools, popups ----------
S.t='hand';sync();                            // start in reading mode; pick a tool to write
const dot=()=>{$('#dot').style.background=S.c};dot();
$('#tcb').onclick=()=>{const p=$('#pc2'),h=p.hidden;$('#mm').hidden=true;p.hidden=!h};
$('#pc2').addEventListener('click',dot);
$('#pc2').addEventListener('input',dot);
$('#mn').addEventListener('click',()=>{$('#pc2').hidden=true});
document.addEventListener('pointerdown',e=>{if(!e.target.closest('#pc2,#mm,#tcb,#mn')){$('#pc2').hidden=true;$('#mm').hidden=true}},true);
$('#mm').addEventListener('click',e=>{const b=e.target.closest('button');if(b&&!['m-po','m-nt','tp','zo','zi'].includes(b.id))$('#mm').hidden=true});
document.querySelectorAll('[data-sh]').forEach(b=>b.onclick=()=>{S.sh=b.dataset.sh;S.t='sh';sync()});
$('#m-po').addEventListener('click',()=>$('#ed').classList.toggle('po',!!S.po));
// opening a blank note starts with the pen
const _o=window.openDoc;
window.openDoc=async function(id){
 let d;try{d=await tx('d','readonly',s=>s.get(id))}catch(e){}
 if(d&&d.lk&&d.k==='note'&&!await raVerify({title:'الملف مقفول 🔒',sub:d.n}))return;
 const r=await _o.apply(this,arguments);
 if(S.d&&S.d.k==='note'&&S.t==='hand'){S.t='pen';sync()}
 return r};

// ---------- text selection bar: highlight / copy ----------
{const c=$('#sa .sac');['#facc15','#86efac','#93c5fd','#f9a8d4'].forEach(k=>{const b=E('button','sw');b.style.background=k;b.setAttribute('aria-label','هايلايت');b.onclick=()=>mkHl(k);c.append(b)})}
window.updSA=()=>{$('#sa').hidden=!hs()};
$('#sa').addEventListener('pointerdown',e=>e.preventDefault());
$('#cp').onclick=async()=>{const t=getSelection().toString();try{await navigator.clipboard.writeText(t);toast('اتنسخ')}catch(e){toast('تعذر النسخ')}getSelection().removeAllRanges()};

// ---------- translation of the selected text ----------
{
 const AR=/[\u0600-\u06FF]/;
 const split=t=>{const out=[];let c='';
  for(const p of t.split(/\s+/).reduce((a,w)=>{const l=a.length-1;if(l>=0&&/[.!?؟]$/.test(a[l])===false&&(a[l]+' '+w).length<=1200)a[l]+=' '+w;else a.push(w);return a},[])){
   if((c+' '+p).length>1200&&c){out.push(c);c=''}c=c?c+' '+p:p}
  if(c)out.push(c);return out};
 const google=async(q,tl)=>{const r=await fetch('https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl='+tl+'&dt=t&q='+encodeURIComponent(q));
  if(!r.ok)throw new Error('g'+r.status);const j=await r.json();return j[0].map(x=>x[0]).join('')};
 const memory=async(q,tl)=>{const sl=tl==='ar'?'en':'ar';const r=await fetch('https://api.mymemory.translated.net/get?q='+encodeURIComponent(q.slice(0,450))+'&langpair='+sl+'|'+tl);
  if(!r.ok)throw new Error('m'+r.status);const j=await r.json();const t=j&&j.responseData&&j.responseData.translatedText;if(!t)throw new Error('empty');return t};
 const translate=async(text,tl)=>{const parts=split(text),res=[];
  for(const p of parts){try{res.push(await google(p,tl))}catch(e){res.push(await memory(p,tl))}}return res.join(' ')};
 window.raTranslate=translate;
 let tok=0;
 const show=(text)=>{
  document.querySelector('#trc')&&document.querySelector('#trc').remove();
  let tl=AR.test(text)?'en':'ar';
  const w=E('div','','<div class="trh"><b>🌐 ترجمة</b><span class="trs"></span><button class="trf" aria-label="تغيير اللغة"></button><button class="trx" aria-label="إغلاق">✕</button></div><div class="tro" dir="auto"></div><div class="trt" dir="auto"></div><div class="trbt"><button class="trcp">📋 نسخ الترجمة</button></div>');
  w.id='trc';w.querySelector('.tro').textContent=text.length>300?text.slice(0,300)+'…':text;
  const o=w.querySelector('.trt'),f=w.querySelector('.trf');
  const go=async()=>{const my=++tok;f.textContent=tl==='ar'?'إلى: العربية ⇄':'إلى: English ⇄';o.dir=tl==='ar'?'rtl':'ltr';o.textContent='جاري الترجمة…';o.className='trt ld';
   try{const r=await translate(text,tl);if(my!==tok)return;o.textContent=r;o.className='trt'}
   catch(e){if(my!==tok)return;o.textContent='تعذرت الترجمة — اتأكد إن النت شغال وجرّب تاني.';o.className='trt er'}};
  f.onclick=()=>{tl=tl==='ar'?'en':'ar';go()};
  w.querySelector('.trx').onclick=()=>{tok++;w.remove()};
  w.querySelector('.trcp').onclick=async()=>{try{await navigator.clipboard.writeText(o.textContent);toast('اتنسخ')}catch(e){toast('تعذر النسخ')}};
  w.addEventListener('pointerdown',e=>e.stopPropagation());
  document.body.append(w);go()};
 $('#trb').onclick=()=>{const t=getSelection().toString().replace(/\s+/g,' ').trim();if(!t)return;show(t.slice(0,4000));getSelection().removeAllRanges()};
}

// ---------- PIN lock (per file) ----------
const KEY='inkpin',enc=new TextEncoder();
const hex=b=>[...new Uint8Array(b)].map(x=>x.toString(16).padStart(2,'0')).join('');
const getPin=()=>{try{return JSON.parse(localStorage.getItem(KEY)||'null')}catch(e){return null}};
async function hash(pin,salt){const k=await crypto.subtle.importKey('raw',enc.encode(pin),'PBKDF2',false,['deriveBits']);
 return hex(await crypto.subtle.deriveBits({name:'PBKDF2',salt,iterations:150000,hash:'SHA-256'},k,256))}
const unhex=h=>Uint8Array.from(h.match(/../g),x=>parseInt(x,16));
function askPin(title,sub){return new Promise(ok=>{
 const w=E('div','dlg','<div class="dbox" role="dialog" aria-modal="true"><p></p><input type="password" inputmode="numeric" autocomplete="off" maxlength="12" dir="ltr"><div class="dbt"><button class="dc">إلغاء</button><button class="do">تأكيد</button></div></div>');
 w.querySelector('p').textContent=title+(sub?'\n'+sub:'');const i=w.querySelector('input');
 const done=v=>{w.remove();ok(v)};
 w.querySelector('.dc').onclick=()=>done(null);w.querySelector('.do').onclick=()=>done(i.value);
 i.onkeydown=e=>{if(e.key==='Enter')done(i.value)};document.body.append(w);setTimeout(()=>i.focus(),60)})}
window.raVerify=async o=>{const p=getPin();if(!p)return true;
 for(;;){const v=await askPin(o.title||'الرقم السري',o.sub);if(v==null)return false;
  if(await hash(v,unhex(p.s))===p.h)return true;
  if(!await raYN('الرقم غلط. تحاول تاني؟','حاول تاني','إلغاء'))return false}};
window.raLockToggle=async d=>{
 if(d.lk){if(!await raVerify({title:'فك القفل',sub:d.n}))return;d.lk=0}
 else{
  if(!getPin()){
   const a=await askPin('اختار رقم سري (4 أرقام على الأقل)');if(a==null)return;
   if(a.length<4)return alert('الرقم لازم يكون 4 أرقام على الأقل');
   const b=await askPin('اكتب الرقم تاني للتأكيد');if(b==null)return;
   if(a!==b)return alert('الرقمين مش متطابقين');
   const s=crypto.getRandomValues(new Uint8Array(16));
   try{localStorage.setItem(KEY,JSON.stringify({s:hex(s),h:await hash(a,s)}))}catch(e){return alert('تعذر حفظ الرقم السري')}}
  d.lk=1}
 await dbPut(d);home()};
})();

// ---------- Font picker (Arabic / English) — shared by the Word and PowerPoint editors ----------
// window.raFontPicker(onPick): step 1 asks "عربي or English", step 2 lists that language's fonts, each drawn in its own face.
(()=>{
 'use strict';
 const E=(t,c,h)=>{const e=document.createElement(t);if(c)e.className=c;if(h!=null)e.innerHTML=h;return e};
 const FONTS={
  ar:['Traditional Arabic','Simplified Arabic','Sakkal Majalla','Arabic Typesetting','Cairo','Tajawal','Almarai','Amiri','Noto Naskh Arabic','Noto Kufi Arabic','Noto Sans Arabic','IBM Plex Sans Arabic','Mada','Reem Kufi','El Messiri','Lateef','Scheherazade New','Markazi Text'],
  en:['Calibri','Arial','Times New Roman','Cambria','Georgia','Verdana','Tahoma','Segoe UI','Trebuchet MS','Courier New']
 };
 const SAMPLE={ar:'أبجد هوز حطي كلمن 123',en:'The quick brown fox 123'};
 const cv=document.createElement('canvas').getContext('2d');
 const wd=(f,t)=>{cv.font='32px '+f;return cv.measureText(t).width};
 // a font counts as available when drawing with it differs from every generic fallback
 const has=async(n,smp)=>{try{await document.fonts.load('32px "'+n+'"',smp)}catch(e){}
  return ['monospace','serif','sans-serif'].some(b=>wd('"'+n+'",'+b,smp)!==wd(b,smp))};
 if(!document.getElementById('rafp-css')){const c=E('style');c.id='rafp-css';c.textContent=`
.rafp-ov{position:fixed;inset:0;z-index:2147483000;background:#0008;display:flex;align-items:flex-end;justify-content:center}
.rafp-sh{width:100%;max-width:560px;max-height:82vh;display:flex;flex-direction:column;background:var(--card,#fff);color:var(--fg,#111);border-radius:18px 18px 0 0;padding:14px 14px calc(14px + env(safe-area-inset-bottom));box-sizing:border-box;direction:rtl;font-family:inherit}
.rafp-hd{display:flex;align-items:center;gap:8px;margin-bottom:10px}
.rafp-hd b{flex:1;font-size:17px}
.rafp-hd button{flex:none;min-width:40px;height:40px;border-radius:12px;border:1px solid var(--line,#ccd);background:var(--bg,#f4f6fb);color:inherit;font-size:18px;font-family:inherit}
.rafp-ch{display:flex;gap:10px}
.rafp-ch button{flex:1;display:flex;flex-direction:column;align-items:center;gap:4px;padding:18px 8px;border-radius:16px;border:2px solid var(--line,#ccd);background:var(--bg,#f4f6fb);color:inherit;font-family:inherit}
.rafp-ch button i{font-style:normal;font-size:34px;line-height:1.2}
.rafp-ch button span{font-size:15px;font-weight:700}
.rafp-ls{overflow-y:auto;display:flex;flex-direction:column;gap:6px;-webkit-overflow-scrolling:touch}
.rafp-row{flex:none;display:block;width:100%;text-align:start;padding:10px 12px;border-radius:12px;border:1px solid var(--line,#ccd);background:var(--bg,#f4f6fb);color:inherit;font-family:inherit}
.rafp-row small{display:block;font-size:12px;opacity:.65;direction:ltr;text-align:start}
.rafp-row .s{display:block;font-size:22px;line-height:1.5;margin-top:2px}
.rafp-row .s.en{direction:ltr;text-align:left}
.rafp-row.cur{border-color:var(--acc,#2f5fe0);border-width:2px}
.rafp-row.cur small::before{content:'✓ ';font-weight:700}
.rafp-row.off{opacity:.5}
.rafp-row.off small::after{content:'  ⚠ مش متاح على الجهاز ده'}
`;document.head.append(c)}
 let cur=null;
 const close=()=>{if(cur){cur.remove();cur=null}};
 window.raFontPicker=(onPick,current)=>{
  close();
  const ov=E('div','rafp-ov'),sh=E('div','rafp-sh');ov.append(sh);cur=ov;
  ov.addEventListener('mousedown',e=>e.preventDefault()); // keep the editor's text selection
  ov.addEventListener('click',e=>{if(e.target===ov)close()});
  const head=(title,back)=>{const h=E('div','rafp-hd');if(back){const b=E('button','', '→');b.setAttribute('aria-label','رجوع');b.onclick=step1;h.append(b)}
   h.append(E('b','',title));const x=E('button','','✕');x.setAttribute('aria-label','إغلاق');x.onclick=close;h.append(x);return h};
  const step1=()=>{sh.innerHTML='';sh.append(head('اختار نوع الخط'));
   const ch=E('div','rafp-ch');
   for(const [k,ic,lb] of [['ar','ع','خطوط عربي'],['en','Aa','English fonts']]){
    const b=E('button','','<i></i><span></span>');b.querySelector('i').textContent=ic;b.querySelector('span').textContent=lb;b.onclick=()=>step2(k);ch.append(b)}
   sh.append(ch)};
  const step2=k=>{sh.innerHTML='';sh.append(head(k==='ar'?'الخطوط العربي':'English fonts',1));
   const ls=E('div','rafp-ls');sh.append(ls);
   for(const n of FONTS[k]){
    const r=E('button','rafp-row','<small></small><span class="s"></span>');r.querySelector('small').textContent=n;
    const s=r.querySelector('.s');s.textContent=SAMPLE[k];s.style.fontFamily='"'+n+'",'+(k==='ar'?'"Noto Naskh Arabic",':'')+'sans-serif';if(k==='en')s.classList.add('en');
    if(current&&String(current).toLowerCase()===n.toLowerCase())r.classList.add('cur');
    r.onclick=()=>{close();try{onPick(n)}catch(e){console.warn('font',e)}};ls.append(r);
    has(n,SAMPLE[k]).then(ok=>{if(!ok)r.classList.add('off')})}};
  document.body.append(ov);
  const inK=current&&['ar','en'].find(k=>FONTS[k].some(n=>n.toLowerCase()===String(current).toLowerCase()));
  inK?step2(inK):step1();
 };
})();

// Applies a font to the current selection of `root` and removes inline fonts inside it that would override the choice.
// Also returns the font name at the caret/selection via window.raFontAt.
window.raFontAt=()=>{const g=getSelection();if(!g.rangeCount)return '';let n=g.anchorNode;if(n&&n.nodeType===3)n=n.parentElement;if(!n)return '';
 return (getComputedStyle(n).fontFamily||'').split(',')[0].replace(/["']/g,'').trim()};
window.raApplyFont=(root,name)=>{
 const q='"'+name+'"';
 document.execCommand('styleWithCSS',false,true);document.execCommand('fontName',false,name);
 const g=getSelection();if(!g.rangeCount||g.isCollapsed||!root)return;
 const r=g.getRangeAt(0),nr=document.createRange();
 root.querySelectorAll('*').forEach(n=>{if(!r.intersectsNode(n))return;nr.selectNodeContents(n);
  if(r.compareBoundaryPoints(Range.START_TO_START,nr)>0||r.compareBoundaryPoints(Range.END_TO_END,nr)<0)return; // only elements fully inside the selection
  if(n.tagName==='FONT'&&n.hasAttribute('face'))n.setAttribute('face',name);
  if(n.style&&n.style.fontFamily)n.style.fontFamily=q})};
