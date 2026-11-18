// Reader — UI glue. Loaded after app.js (uses its globals: S, sync, mkHl, hs, tx, dbPut, home, openDoc).
(()=>{
'use strict';
const $=s=>document.querySelector(s),E=(t,c,h)=>{const e=document.createElement(t);if(c)e.className=c;if(h!=null)e.innerHTML=h;return e};

// ---------- hooks that app.js expects ----------
window.LS={r(){},d(){},m(){},u(){}};        // lasso is not part of this build
window.SN=p=>p;                              // no grid snapping
window.inkBusy=(on,t)=>{let b=$('#busy');if(!on){b&&b.remove();return}
 if(!b){b=E('div','','<div class="bx"><i class="spin"></i><span></span></div>');b.id='busy';document.body.append(b)}
 b.querySelector('span').textContent=t||''};
const toast=t=>{const d=E('div','oftoast');d.textContent=t;document.body.append(d);setTimeout(()=>d.remove(),2200)};

// ---------- home ----------
const tabPaint=()=>{const t=inkGetTab();document.querySelectorAll('#hnav button').forEach(b=>{const on=b.dataset.tab===t;b.classList.toggle('on',on);b.setAttribute('aria-selected',on)});document.querySelectorAll('.tp').forEach(p=>p.hidden=p.dataset.tab!==t)};
document.querySelectorAll('#hnav button').forEach(b=>b.onclick=()=>{inkSetTab(b.dataset.tab);S.tg='';tabPaint();home();$('#home').scrollTop=0});
tabPaint();
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
 if(d&&d.lk&&(d.k==='pdf'||d.k==='note')&&!await inkVerify({title:'الملف مقفول 🔒',sub:d.n}))return;
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
 window.inkTranslate=translate;
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
window.inkVerify=async o=>{const p=getPin();if(!p)return true;
 for(;;){const v=await askPin(o.title||'الرقم السري',o.sub);if(v==null)return false;
  if(await hash(v,unhex(p.s))===p.h)return true;
  if(!await inkYN('الرقم غلط. تحاول تاني؟','حاول تاني','إلغاء'))return false}};
window.inkLockToggle=async d=>{
 if(d.lk){if(!await inkVerify({title:'فك القفل',sub:d.n}))return;d.lk=0}
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
