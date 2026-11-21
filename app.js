(()=>{const sh=m=>{let d=document.getElementById('err');if(!d){d=document.createElement('div');d.id='err';d.style.cssText='position:fixed;left:0;right:0;top:0;z-index:99;background:#b91c1c;color:#fff;font:12px monospace;padding:6px 8px;direction:ltr;white-space:pre-wrap';d.onclick=()=>d.remove();document.body.append(d)}d.textContent='خطأ (اضغط للإخفاء): '+m};
addEventListener('error',e=>sh(e.message+' @'+(e.lineno||'')));addEventListener('unhandledrejection',e=>sh(String(e.reason&&e.reason.message||e.reason)))})();
const $=s=>document.querySelector(s),E=(t,c,h)=>{const e=document.createElement(t);if(c)e.className=c;if(h!=null)e.innerHTML=h;return e};
// one-time migration from the old internal names (localStorage keys + IndexedDB 'inkpad') so existing files and settings survive the rename
try{for(const k of['so','tab','pos','col','sz']){const v=localStorage.getItem('ink'+k);if(v!=null&&localStorage.getItem('ra'+k)==null)localStorage.setItem('ra'+k,v)}}catch(e){}
const migrateOldDB=db=>new Promise(ok=>{try{if(localStorage.getItem('ramigrated')){return ok()}}catch(e){}
 const q=indexedDB.open('inkpad');let fresh=0;q.onupgradeneeded=()=>{fresh=1;q.transaction.abort()};q.onerror=()=>{if(fresh)try{indexedDB.deleteDatabase('inkpad')}catch(e){}ok()};
 q.onsuccess=()=>{const o=q.result;if(!o.objectStoreNames.contains('d')||!o.objectStoreNames.contains('f')){o.close();return ok()}
  const t=o.transaction(['d','f'],'readonly'),d=t.objectStore('d').getAll(),fk=t.objectStore('f').getAllKeys(),fv=t.objectStore('f').getAll();
  t.oncomplete=()=>{o.close();const w=db.transaction(['d','f'],'readwrite');
   const cd=w.objectStore('d'),cf=w.objectStore('f');cd.count().onsuccess=e=>{if(!e.target.result){(d.result||[]).forEach(x=>cd.put(x));(fk.result||[]).forEach((k,i)=>cf.put(fv.result[i],k))}};
   w.oncomplete=()=>{try{localStorage.setItem('ramigrated','1')}catch(e){}ok()};w.onerror=w.onabort=()=>ok()};t.onerror=t.onabort=()=>{o.close();ok()}}});
const DB=new Promise((r,no)=>{const q=indexedDB.open('readerall',1);q.onupgradeneeded=()=>{q.result.createObjectStore('d',{keyPath:'id'});q.result.createObjectStore('f')};q.onsuccess=()=>{const db=q.result;migrateOldDB(db).then(()=>r(db))};q.onerror=()=>{alert('تعذر فتح قاعدة البيانات. لو بتستخدم تصفح خفي جرّب وضع عادي.');no(q.error)};q.onblocked=()=>alert('اقفل أي نافذة تانية مفتوحة من InkPad وحاول تاني')});
try{navigator.storage&&navigator.storage.persist&&navigator.storage.persist()}catch(e){}
const tx=async(st,m,f)=>{const db=await DB;return new Promise((ok,no)=>{const t=db.transaction(st,m),r=f(t.objectStore(st));t.oncomplete=()=>ok(r&&r.result);t.onerror=()=>no(t.error);t.onabort=()=>no(t.error||new Error('abort'))})};
const dbPut=d=>tx('d','readwrite',s=>s.put(d));
const S={t:'pen',lp:'pen',c:'#1e40af',w:3,z:1,fit:1,sh:'line',und:[],red:[],pg:[]},W=595,H=842;
// NV = وضع التحكم في الشاشة بس (من غير أداة) أو وضع تحديد النص
const NV=()=>S.t==='hand'||S.t==='th',PENT=['pen','hl','lz','er'];
// ---------- home ----------
const SZ=()=>{try{return JSON.parse(localStorage.getItem('rasz')||'{}')}catch(e){return{}}},
SO=()=>{try{return localStorage.getItem('raso')||'u'}catch(e){return'u'}},
fmt=n=>n<1048576?Math.max(1,Math.round(n/1024))+' KB':(n/1048576).toFixed(1)+' MB',
raSort=(x,y)=>{const o=SO(),z=SZ();return o==='n'?String(x.n).localeCompare(String(y.n),'ar'):o==='sz'?(z[y.id]||0)-(z[x.id]||0):y.u-x.u};
const TABK={ppt:['ppt'],doc:['doc'],xls:['xls'],settings:[]},TABM={ppt:'لسه مفيش عروض PowerPoint. افتح ملف .pptx.',doc:'لسه مفيش مستندات Word. افتح ملف .docx أو اعمل مستند جديد.',xls:'لسه مفيش جداول Excel. افتح ملف أو اعمل جدول جديد.',settings:''};
let TAB=(()=>{try{const t=localStorage.getItem('ratab');return TABK[t]?t:'ppt'}catch(e){return'ppt'}})();
window.raGetTab=()=>TAB;window.raSetTab=t=>{if(!TABK[t])return;TAB=t;try{localStorage.setItem('ratab',t)}catch(e){}};
const inTab=d=>TABK[TAB].includes(d.k);
async function home(){$('#ed').hidden=true;$('#home').hidden=false;const l=$('#list');l.innerHTML='';
 const a=(await tx('d','readonly',s=>s.getAll())).sort(raSort);
 const ta=a.filter(inTab);lib(l,ta);{const hs=$('#hsub'),visibleCount=a.filter(inTab).length;hs&&(hs.textContent=visibleCount?visibleCount+' ملف':'')}const q=$('#ls').value.trim().toLowerCase(),fl=ta.filter(d=>(!q||(d.n+' '+(d.tag||'')).toLowerCase().includes(q))&&(!S.tg||d.tag===S.tg));if(!ta.length)l.innerHTML='<p class="em">'+TABM[TAB]+'</p>';
 fl.forEach(d=>{const r=E('div','it','<i class="ft"></i><div><b></b><small></small></div><button class="rn" aria-label="إعادة تسمية">✏️</button><button aria-label="حذف"><svg class="i" viewBox="0 0 24 24" aria-hidden="true"><path d="M3 6h18"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/></svg></button><button class="tg" aria-label="وسم"><svg class="i" viewBox="0 0 24 24" aria-hidden="true"><path d="M20.6 13.4 13.4 20.6a2 2 0 0 1-2.8 0L2 12V2h10l8.6 8.6a2 2 0 0 1 0 2.8z"/><circle cx="7" cy="7" r="1.2"/></svg></button><button class="lkb" aria-label="قفل الملف"></button>');{const FT={ppt:'PPT',doc:'DOC',xls:'XLS'},f=r.querySelector('.ft');f.textContent=FT[d.k]||'✎';f.classList.add(FT[d.k]?d.k:'nt')}
  r.querySelector('b').textContent=d.n;r.querySelector('small').textContent=(window.raKind?raKind(d):'ملاحظة')+' · '+new Date(d.u).toLocaleDateString('ar-EG')+(d.tag?' · 📁'+d.tag:'');
  r.querySelector('.ft').onclick=r.querySelector('div').onclick=()=>openDoc(d.id);r.querySelector('.rn').onclick=async()=>{const n=await raAsk('اسم الملف الجديد:',d.n);if(n&&n.trim()){d.n=n.trim().slice(0,120);d.u=Date.now();await dbPut(d);home()}};r.querySelector('.tg').onclick=async()=>{const t=await raAsk('اسم المجلد/الوسم (فاضي = بدون):',d.tag||'');if(t!==null){d.tag=t.trim();await dbPut(d);home()}};
  {const lb=r.querySelector('.lkb');lb.textContent=d.lk?'🔒':'🔓';lb.classList.toggle('on',!!d.lk);lb.onclick=()=>window.raLockToggle&&raLockToggle(d)}
  r.querySelector('button[aria-label="حذف"]').onclick=async()=>{if(d.lk&&window.raVerify&&!await raVerify({title:'الملف مقفول 🔒',sub:d.n}))return;if((()=>{try{return JSON.parse(localStorage.getItem('readerall.settings.v1')||'{}').deleteConfirm!==false}catch(e){return true}})()&&!await raConfirm('حذف الملف؟'))return;{try{for(const a of d.au||[])await tx('f','readwrite',s=>s.delete(a.k));await tx('d','readwrite',s=>s.delete(d.id));await tx('f','readwrite',s=>s.delete(d.id))}catch(e){alert('فشل الحذف: '+(e&&e.message||e))}home()}};l.append(r)})}
// ---------- settings: unified file manager + storage information ----------
window.raListAllFiles=async()=>{
 const rows=await tx('d','readonly',s=>s.getAll());
 return rows.sort((a,b)=>(b.u||0)-(a.u||0)).map(d=>({id:d.id,n:d.n||'بدون اسم',k:d.k||'other',u:d.u||0,tag:d.tag||'',lk:!!d.lk}));
};
window.raStorageStats=async()=>{
 const rows=await tx('d','readonly',s=>s.getAll());let bytes=0;
 for(const d of rows){try{bytes+=new Blob([JSON.stringify({...d,au:undefined})]).size}catch(e){};for(const a of d.au||[]){try{const v=await tx('f','readonly',s=>s.get(a.k));if(v)bytes+=v.byteLength||v.size||0}catch(e){}};try{const v=await tx('f','readonly',s=>s.get(d.id));if(v)bytes+=v.byteLength||v.size||0}catch(e){}}
 let estimate=null;try{if(navigator.storage&&navigator.storage.estimate)estimate=await navigator.storage.estimate()}catch(e){}
 return {count:rows.length,bytes,quota:estimate&&estimate.quota,usage:estimate&&estimate.usage};
};
window.raOpenStoredFile=async id=>openDoc(id);
window.raRenameStoredFile=async(id,name)=>{const d=await tx('d','readonly',s=>s.get(id));if(!d)return false;d.n=name;d.u=Date.now();await dbPut(d);return true};
window.raDeleteStoredFile=async id=>{const d=await tx('d','readonly',s=>s.get(id));if(!d)return false;for(const a of d.au||[])await tx('f','readwrite',s=>s.delete(a.k));await tx('d','readwrite',s=>s.delete(id));await tx('f','readwrite',s=>s.delete(id));return true};

// ---------- editor ----------
async function openDoc(id){const my=S.ot=(S.ot||0)+1;const d=await tx('d','readonly',s=>s.get(id));if(my!==S.ot)return;if(!d)return home();if(d.k==='pdf'){alert('قارئ PDF غير متاح في هذا التطبيق. يمكنك استخدام أدوات تحويل المستندات إلى PDF.');return home()}S.d=d;S.pdf=null;S.und=[];S.red=[];S.tc=[];S.sr=[];S.si=-1;S.lq='';clearHl();$('#ed').classList.toggle('nt',!!d.nt);$('#sb').hidden=true;$('#cnt').textContent='';$('#si').value='';$('#ed').classList.remove('fs');
 $('#home').hidden=true;$('#ed').hidden=false;$('#ttl').textContent=d.n;$('#ap').hidden=d.k!=='note';if($('#sq'))$('#sq').hidden=true;
 let n=d.np;
 const p1={width:W,height:H};if(my!==S.ot)return;
 S.fit=Math.min(2,($('#sc').clientWidth-8)/p1.width);S.z=Math.max(S.fit*.5,Math.min(S.fit*3,S.fit*(d.zr||1)));
 const g=$('#pgs');g.innerHTML='';S.pg=[];S.io&&S.io.disconnect();
 S.io=new IntersectionObserver(es=>es.forEach(e=>{const el=e.target;el.vis=e.isIntersecting;if(e.isIntersecting)rp(el.i);else gcSoon()}),{root:$('#sc'),rootMargin:'300px 0px'});
 for(let i=0;i<n;i++)addPg(i,[p1.width,p1.height]);S.rs=1;rst(d);setTimeout(()=>S.rs=0,700)}
function rst(d){const el=S.pg[d.pi||0],sc=$('#sc');if(el&&(d.pi||d.po)){const x=el.getBoundingClientRect(),y=sc.getBoundingClientRect();sc.scrollTop+=x.top-y.top+(d.po||0)*x.height}up()}
function addPg(i,dim){const el=E('div','pg');el.i=i;el.dim=dim;el.bg=E('canvas');el.hc=E('canvas','hlc');el.ink=E('canvas','ink');el.append(el.bg,el.hc,el.ink);
 el.style.width=dim[0]*S.z+'px';el.style.height=dim[1]*S.z+'px';bind(el);S.pg[i]=el;$('#pgs').append(el);S.io.observe(el)}
async function bgDraw(i,c,k,own){const el=S.pg[i];c.fillStyle='#fff';c.fillRect(0,0,el.dim[0]*k,el.dim[1]*k);
 const t=S.d.tp,s=28*k;c.strokeStyle=c.fillStyle='#cbd5e1';c.lineWidth=Math.max(1,k*.6);
 if(t==='lines')for(let y=s*1.5;y<H*k;y+=s){c.beginPath();c.moveTo(0,y);c.lineTo(W*k,y);c.stroke()}
 else if(t==='grid'){for(let y=s;y<H*k;y+=s){c.beginPath();c.moveTo(0,y);c.lineTo(W*k,y);c.stroke()}for(let x=s;x<W*k;x+=s){c.beginPath();c.moveTo(x,0);c.lineTo(x,H*k);c.stroke()}}
 else if(t==='dots')for(let y=s;y<H*k;y+=s)for(let x=s;x<W*k;x+=s)c.fillRect(x-k*.8,y-k*.8,k*1.6,k*1.6)}
async function rp0(i){const el=S.pg[i];if(!el||el.r===S.z)return;el.r=S.z;const t=el.tk=(el.tk||0)+1;
 const w=el.dim[0]*S.z,h=el.dim[1]*S.z;let d=Math.min(devicePixelRatio||1,2)*(S.df||1);d=Math.max(.8,Math.min(d,Math.sqrt(3e6/(w*h))));const oc=E('canvas');oc.width=w*d;oc.height=h*d;
 const t0=performance.now();try{await bgDraw(i,oc.getContext('2d'),S.z*d,1)}catch(e){if(t!==el.tk||(e&&e.name==='RenderingCancelledException')){oc.width=oc.height=0;return}throw e}
 if(t!==el.tk){oc.width=oc.height=0;return}if(performance.now()-t0>700)S.df=Math.max(.6,(S.df||1)*.85);
 el.style.width=w+'px';el.style.height=h+'px';
 for(const c of[el.bg,el.hc,el.ink]){c.width=w*d;c.height=h*d}
 el.bg.getContext('2d').drawImage(oc,0,0);oc.width=oc.height=0;ink(i);buildTl(i)}
// رسم الصفحات: صفحة واحدة في المرة، الأقرب لنص الشاشة الأول، وبعد ما السحب يقف. والصفحات القريبة بتفضل محفوظة
let rq=new Set(),rb=0,rt=0,lsc=0,gt=0;
const nearEl=(el,m)=>{const a=el.getBoundingClientRect(),b=$('#sc').getBoundingClientRect();return a.bottom>b.top-m&&a.top<b.bottom+m};
function rp(i){rq.add(i);rpump()}
function rpump(){if(rb||!rq.size)return;const w=130-(Date.now()-lsc);if(w>0){clearTimeout(rt);rt=setTimeout(rpump,w);return}
 const r=$('#sc').getBoundingClientRect(),cy=r.top+r.height/2;let best=-1,bd=1e9;
 for(const i of[...rq]){const el=S.pg&&S.pg[i];if(!el||el.r===S.z||!nearEl(el,700)){rq.delete(i);continue}const q=el.getBoundingClientRect(),dd=Math.abs(q.top+q.height/2-cy);if(dd<bd){bd=dd;best=i}}
 if(best<0)return;rq.delete(best);rb=1;rp0(best).catch(()=>{}).finally(()=>{rb=0;rpump();gcSoon()})}
function gcSoon(){clearTimeout(gt);gt=setTimeout(()=>{const m=$('#sc').clientHeight*2;(S.pg||[]).forEach(el=>{if(el&&el.r&&!nearEl(el,m)){el.bg.width=el.hc.width=el.ink.width=1;el.r=null;el.tk=(el.tk||0)+1;el.tl&&(el.tl.remove(),el.tl=null)}})},400)}
$('#sc').addEventListener('scroll',()=>{if(Date.now()-(S.zt||0)>400)lsc=Date.now();clearTimeout(rt);rt=setTimeout(rpump,140)},{passive:true});
function ds(c,o,k){c.save();c.lineCap=c.lineJoin='round';c.strokeStyle=c.fillStyle=o.c;c.lineWidth=o.w*k*(o.t==='hl'?4:1);if(o.t==='hl'){c.globalAlpha=.55;c.globalCompositeOperation='multiply'}
 if(o.t==='im'){const g=im(o.src);if(g.complete&&g.naturalWidth)c.drawImage(g,o.p[0][0]*k,o.p[0][1]*k,o.w*k,o.h*k);c.restore();return}
 if(o.t==='hr'){c.globalAlpha=.55;c.globalCompositeOperation='multiply';c.fillStyle=o.c;o.rs.forEach(r=>c.fillRect(r[0]*k,r[1]*k,r[2]*k,r[3]*k));c.restore();return}
 const p=o.p,X=i=>p[i][0]*k,Y=i=>p[i][1]*k;c.beginPath();
 if(o.t==='pen'&&p.some(q=>q[2]>0&&q[2]!==.5)){for(let i=1;i<p.length;i++){c.beginPath();c.lineWidth=o.w*k*(.4+1.2*(p[i][2]+p[i-1][2])/2);c.moveTo(X(i-1),Y(i-1));c.lineTo(X(i),Y(i));c.stroke()}}else if(o.t==='pen'||o.t==='hl'){c.moveTo(X(0),Y(0));for(let i=1;i<p.length-1;i++)c.quadraticCurveTo(X(i),Y(i),(X(i)+X(i+1))/2,(Y(i)+Y(i+1))/2);c.lineTo(X(p.length-1),Y(p.length-1));c.stroke()}
 else if(o.t==='line'){c.moveTo(X(0),Y(0));c.lineTo(X(1),Y(1));c.stroke()}
 else if(o.t==='rect')c.strokeRect(X(0),Y(0),X(1)-X(0),Y(1)-Y(0));
 else if(o.t==='ell'){c.ellipse((X(0)+X(1))/2,(Y(0)+Y(1))/2,Math.abs(X(1)-X(0))/2||1,Math.abs(Y(1)-Y(0))/2||1,0,0,7);c.stroke()}
 else{c.font=o.w*k+'px sans-serif';c.textBaseline='top';o.s.split('\n').forEach((t,j)=>c.fillText(t,X(0),Y(0)+j*o.w*k*1.25))}
 c.restore()}
const HLT=o=>o.t==='hl'||o.t==='hr';
function ink(i){const el=S.pg[i],c=el.ink.getContext('2d'),h=el.hc.getContext('2d'),k=el.ink.width/el.dim[0];c.clearRect(0,0,el.ink.width,el.ink.height);h.clearRect(0,0,el.hc.width,el.hc.height);
 (S.d.s[i]||[]).forEach(o=>ds(HLT(o)?h:c,o,k));if(el.cur)ds(HLT(el.cur)?h:c,el.cur,k);el.lz&&el.lz.length&&lzd(c,el,k);LS.r(c,el,k);window.HSR&&HSR(c,el,k)}
// ---------- drawing ----------
const pt=(e,el)=>{const r=el.getBoundingClientRect();return[(e.clientX-r.left)/S.z,(e.clientY-r.top)/S.z,e.pointerType==='pen'&&e.pressure?+e.pressure.toFixed(2):.5]};
function sp(o){if(o.t==='im')return[0,1,2,3,4].flatMap(u=>[0,1,2,3,4].map(v=>[o.p[0][0]+o.w*u/4,o.p[0][1]+o.h*v/4]));if(o.t==='hr')return o.rs.flatMap(r=>[0,1,2,3,4].map(k=>[r[0]+r[2]*k/4,r[1]+r[3]/2]));const p=o.p;if(o.t==='pen'||o.t==='hl'||o.t==='tx')return p;const[a,b]=p,r=[],L=(u,v)=>{for(let k=0;k<=10;k++)r.push([u[0]+(v[0]-u[0])*k/10,u[1]+(v[1]-u[1])*k/10])};
 if(o.t==='line')L(a,b);else if(o.t==='rect'){L(a,[b[0],a[1]]);L([b[0],a[1]],b);L(b,[a[0],b[1]]);L([a[0],b[1]],a)}
 else{const cx=(a[0]+b[0])/2,cy=(a[1]+b[1])/2,rx=Math.abs(b[0]-a[0])/2,ry=Math.abs(b[1]-a[1])/2;for(let k=0;k<32;k++){const t=k/32*6.283;r.push([cx+rx*Math.cos(t),cy+ry*Math.sin(t)])}}return r}
const segDist=(p,a,b)=>{const dx=b[0]-a[0],dy=b[1]-a[1],L=dx*dx+dy*dy;let t=L?((p[0]-a[0])*dx+(p[1]-a[1])*dy)/L:0;t=Math.max(0,Math.min(1,t));return Math.hypot(p[0]-a[0]-t*dx,p[1]-a[1]-t*dy)};
function hitO(o,p,r){if((o.t==='pen'||o.t==='hl')&&o.p.length>1){for(let i=1;i<o.p.length;i++)if(segDist(p,o.p[i-1],o.p[i])<r)return true;return false}return sp(o).some(q=>Math.hypot(q[0]-p[0],q[1]-p[1])<r)}
function era(el,p){const l=S.d.s[el.i]||[],r=8+S.w*2;for(let j=l.length-1;j>=0;j--)if(hitO(l[j],p,r))el.er.push(...l.splice(j,1));ink(el.i)}
function bind(el){const c=el.ink;
 c.onpointerdown=e=>{if(NV()||!e.isPrimary||S.po&&e.pointerType==='touch')return;e.preventDefault();const p=SN(pt(e,el)),i=el.i;if(S.t==='ls'){c.setPointerCapture(e.pointerId);el.dn=1;LS.d(el,p);return}
  if(S.t==='mv'){const l=S.d.s[i]||[];let h=null;for(let j=l.length-1;j>=0&&!h;j--)if(hitO(l[j],p,14))h=l[j];if(h){el.mo={o:h,p0:p,sv:JSON.stringify(h)};c.setPointerCapture(e.pointerId);el.dn=1}return}
  if(S.t==='lz'){const o={c:S.c,w:S.w,p:[p],t:[Date.now()],end:0,d:(S.lzT||7)*1000};(el.lz=el.lz||[]).push(o);el.lo=o;c.setPointerCapture(e.pointerId);el.dn=1;lzLoop();return}
  if(S.t==='tx'){raAsk('اكتب النص:').then(s=>{if(s)add(i,{t:'tx',c:S.c,w:S.w*4+8,p:[p],s})});return}
  if(S.t==='er'){el.er=[];era(el,p)}
  else el.cur={t:S.t==='sh'?S.sh:S.t,c:S.c,w:S.w,p:S.t==='sh'?[p,p]:[p]};
  c.setPointerCapture(e.pointerId);el.dn=1};
 c.onpointermove=e=>{if(!el.dn)return;const p=SN(pt(e,el));if(S.t==='ls'){LS.m(el,p);return}
  if(S.t==='mv'){const m=el.mo;if(m){const dx=p[0]-m.p0[0],dy=p[1]-m.p0[1],s=JSON.parse(m.sv),o=m.o;if(o.rs)o.rs=s.rs.map(r=>[r[0]+dx,r[1]+dy,r[2],r[3]]);else o.p=s.p.map(q=>q.map((v,n)=>n==0?v+dx:n==1?v+dy:v));cancelAnimationFrame(el.f);el.f=requestAnimationFrame(()=>ink(el.i))}return}
  el.lt=Date.now();
  if(S.t==='lz'){if(el.lo){el.lo.p.push(p);el.lo.t.push(Date.now())}return}
  if(S.t==='er')era(el,p);else if(S.t==='sh')el.cur.p[1]=p;else if(S.rl&&S.t==='pen')el.cur.p=[el.cur.p[0],p];else el.cur.p.push(p);
  cancelAnimationFrame(el.f);el.f=requestAnimationFrame(()=>ink(el.i))};
 c.onpointerup=c.onpointercancel=()=>{if(!el.dn)return;el.dn=0;if(S.t==='ls'){LS.u(el);return}
  if(S.t==='mv'){el.mo=null;save();return}
  if(S.t==='lz'){if(el.lo)el.lo.end=Date.now();el.lo=null;return}
  if(S.t==='er'){if(el.er.length)push({a:'del',i:el.i,o:el.er});el.er=null}
  else if(el.cur){const o=el.cur;el.cur=null;if(o.t==='pen'&&o.p.length>8&&Date.now()-(el.lt||0)>550)snap(o);add(el.i,o)}}}
function add(i,o){(S.d.s[i]=S.d.s[i]||[]).push(o);push({a:'add',i,o:[o]})}
function push(a){S.und.push(a);S.red=[];ink(a.i);save()}
function apply(a,inv){if(a.a==='col'){a.o.forEach(o=>o.c=inv?a.f:a.t);ink(a.i);save();return}const l=S.d.s[a.i]=S.d.s[a.i]||[];
 if((a.a==='add')===inv)a.o.forEach(o=>{const x=l.indexOf(o);if(x>=0)l.splice(x,1)});else l.push(...a.o);ink(a.i);save()}
$('#un').onclick=()=>{const a=S.und.pop();if(a){apply(a,true);S.red.push(a)}};
$('#re').onclick=()=>{const a=S.red.pop();if(a){apply(a,false);S.und.push(a)}};
let st;function save(){clearTimeout(st);st=setTimeout(flush,600)}
let fq=0;const fe=e=>{console.warn(e);if(Date.now()-fq>20000){fq=Date.now();alert('⚠ تعذر الحفظ: ممكن مساحة الجهاز خلصت. فضّي مساحة أو اعمل نسخة احتياطية.')}};
function flush(){clearTimeout(st);if(!S.d||$('#ed').hidden)return;S.d.u=Date.now();return dbPut(S.d).catch(fe)}
document.addEventListener('visibilitychange',()=>document.hidden&&flush());addEventListener('pagehide',flush);
// ---------- tools / zoom ----------
function sync(){document.querySelectorAll('[data-t]').forEach(b=>b.classList.toggle('on',b.dataset.t===S.t));$('#ed').classList.toggle('dr',!NV())}
// الأدوات اللي في الشنطة: الضغط تاني على الأداة الشغالة بيرجّع التحكم في الشاشة بس. أدوات القلم (الشريط الفرعي) بتبدّل بس.
document.querySelectorAll('[data-t]').forEach(b=>b.onclick=()=>{const t=b.dataset.t;if(b.closest('.sbar')){S.t=t;PENT.includes(t)&&(S.lp=t)}else S.t=S.t===t?'hand':t;sync()});
$('#shs').onchange=e=>{S.sh=e.target.value;S.t='sh';sync()};
$('#wd').oninput=e=>S.w=+e.target.value;
const cl=$('#cl');
function pc(){cl.querySelectorAll('.sw').forEach(b=>b.classList.toggle('on',b.dataset.c===S.c))}
['#1e40af','#dc2626','#16a34a','#111827','#facc15','#9333ea'].forEach(c=>{const b=E('button','sw');b.style.background=c;b.dataset.c=c;b.onclick=()=>{S.c=c;$('#co').value=c;pc()};cl.append(b)});
$('#co').oninput=e=>{S.c=e.target.value;pc()};pc();sync();
const vis=el=>{const a=el.getBoundingClientRect(),b=$('#sc').getBoundingClientRect();return a.bottom>b.top-600&&a.top<b.bottom+600};
// نقطة ارتكاز ثابتة (رقم الصفحة + نسبة جواها) عشان التكبير/التصغير ما يزحّفش الصفحة
function pgAnchor(fx,fy){const sc=$('#sc'),sr=sc.getBoundingClientRect();let best=0,bd=1e9;
 S.pg.forEach((el,i)=>{const r=el.getBoundingClientRect(),t=r.top-sr.top+sc.scrollTop,b=t+r.height,d=fy<t?t-fy:fy>b?fy-b:0;if(d<bd){bd=d;best=i}});
 const r=S.pg[best].getBoundingClientRect(),t=r.top-sr.top+sc.scrollTop,l=r.left-sr.left+sc.scrollLeft;return{i:best,u:(fy-t)/r.height,v:(fx-l)/r.width}}
function zoom(z,A){z=Math.max(S.fit*.5,Math.min(S.fit*3,z));const sc=$('#sc');if(!S.pg.length)return;
 const cx=A?A.cx:sc.clientWidth/2,cy=A?A.cy:sc.clientHeight/2,fx=A?A.fx:sc.scrollLeft+cx,fy=A?A.fy:sc.scrollTop+cy,an=pgAnchor(fx,fy);
 S.z=z;S.zt=Date.now();lsc=0;S.pg.forEach(el=>{try{el.rt&&el.rt.cancel()}catch(e){}el.tk=(el.tk||0)+1;el.style.width=el.dim[0]*z+'px';el.style.height=el.dim[1]*z+'px';el.r=null;el.tl&&(el.tl.style.transform='scale('+z+')')});S.d.zr=z/S.fit;save();
 const sr=sc.getBoundingClientRect(),r=S.pg[an.i].getBoundingClientRect(),t=r.top-sr.top+sc.scrollTop,l=r.left-sr.left+sc.scrollLeft;
 sc.scrollTop=t+an.u*r.height-cy;sc.scrollLeft=l+an.v*r.width-cx;
 S.pg.forEach(el=>vis(el)&&rp(el.i))}
$('#zi').onclick=()=>zoom(S.z*1.25);$('#zo').onclick=()=>zoom(S.z/1.25);
const dist=e=>Math.hypot(e.touches[0].clientX-e.touches[1].clientX,e.touches[0].clientY-e.touches[1].clientY);
// تكبير بالصوابع: معاينة فورية بـ transform (GPU) وبعدين تثبيت الزوم الحقيقي لما ترفع صوابعك
let PN=null,pnr=0;
const mid2=e=>{const sr=$('#sc').getBoundingClientRect();return[(e.touches[0].clientX+e.touches[1].clientX)/2-sr.left,(e.touches[0].clientY+e.touches[1].clientY)/2-sr.top]};
const pnApply=()=>{pnr=0;if(!PN)return;const g=$('#pgs');g.style.transform='translate('+(PN.cx-PN.cx0)+'px,'+(PN.cy-PN.cy0)+'px) scale('+PN.r+')'};
function pnEnd(){if(!PN)return;const P=PN,g=$('#pgs');PN=null;S.pd=0;cancelAnimationFrame(pnr);pnr=0;g.style.transform='';g.style.transformOrigin='';g.style.willChange='';
 if(Math.abs(P.r-1)>.02)zoom(P.z0*P.r,{fx:P.fx,fy:P.fy,cx:P.cx,cy:P.cy})}
$('#sc').addEventListener('touchstart',e=>{if(e.touches.length===2&&!PN){const sc=$('#sc'),g=$('#pgs'),[cx,cy]=mid2(e),pr=g.getBoundingClientRect(),sr=sc.getBoundingClientRect();
  PN={d0:dist(e),z0:S.z,r:1,cx0:cx,cy0:cy,cx:cx,cy:cy,sl:sc.scrollLeft,st:sc.scrollTop,fx:sc.scrollLeft+cx,fy:sc.scrollTop+cy};S.pd=PN.d0;
  g.style.transformOrigin=(cx+sr.left-pr.left)+'px '+(cy+sr.top-pr.top)+'px';g.style.willChange='transform';
  S.pg.forEach(el=>{el.dn=0;el.cur=null;if(el.ink.width>1)ink(el.i)});e.cancelable&&e.preventDefault()}},{passive:false});
$('#sc').addEventListener('touchmove',e=>{if(PN&&e.touches.length===2){const sc=$('#sc'),[cx,cy]=mid2(e);
  PN.r=Math.max(S.fit*.5/PN.z0,Math.min(S.fit*3/PN.z0,dist(e)/PN.d0));PN.cx=cx;PN.cy=cy;sc.scrollLeft=PN.sl;sc.scrollTop=PN.st;
  pnr||(pnr=requestAnimationFrame(pnApply));e.cancelable&&e.preventDefault()}},{passive:false});
$('#sc').addEventListener('touchend',e=>{if(PN&&e.touches.length<2)pnEnd()},{passive:true});
$('#sc').addEventListener('touchcancel',()=>pnEnd(),{passive:true});
function up(){const sc=$('#sc'),m=sc.getBoundingClientRect().top+sc.clientHeight/3;let c=0;for(let i=0;i<S.pg.length;i++){if(S.pg[i].getBoundingClientRect().top<=m)c=i;else break}$('#pn').textContent=(c+1)+'/'+S.pg.length;S.cp=c;updBm();if(!S.rs&&S.d&&S.pg[c]){const x=S.pg[c].getBoundingClientRect(),y=sc.getBoundingClientRect();S.d.pi=c;S.d.po=(y.top-x.top)/x.height;save()}}
let ur=0;$('#sc').onscroll=()=>{ur||(ur=requestAnimationFrame(()=>{ur=0;up()}))};
$('#pn').onclick=()=>{raAsk('رقم الصفحة:').then(v=>{const n=+v;if(n>0&&n<=S.pg.length)S.pg[n-1].scrollIntoView()})};
$('#ttl').onclick=()=>{raAsk('الاسم:',S.d.n).then(n=>{if(n){S.d.n=n;$('#ttl').textContent=n;save()}})};
$('#ap').onclick=()=>{const i=S.pg.length;S.d.np++;addPg(i,[W,H]);rp(i);S.pg[i].scrollIntoView();save()};
$('#bk').onclick=async()=>{S.ot=(S.ot||0)+1;$('#ed').classList.remove('fs');document.fullscreenElement&&document.exitFullscreen();clearHl();await flush();S.pdf=null;S.d=null;home()};
// ---------- export ----------
async function exImg(b){if(!window.PDFLib&&window.raNeed)await window.raNeed('pl');const D=await PDFLib.PDFDocument.create(),n=S.pg.length;
 for(let i=0;i<n;i++){const el=S.pg[i];
  const k=2,c=E('canvas');c.width=el.dim[0]*k;c.height=el.dim[1]*k;const x=c.getContext('2d');await bgDraw(i,x,k);(S.d.s[i]||[]).forEach(o=>ds(x,o,k));
  const im=await D.embedJpg(await(await fetch(c.toDataURL('image/jpeg',.88))).arrayBuffer());
  D.addPage([el.dim[0],el.dim[1]]).drawImage(im,{x:0,y:0,width:el.dim[0],height:el.dim[1]});b.textContent=(i+1)+'/'+n}
 return D}
$('#ex').onclick=async()=>{const b=$('#ex');b.disabled=true;
 try{const D=await exImg(b)
  const blob=new Blob([await D.save()],{type:'application/pdf'}),f=new File([blob],S.d.n+'.pdf',{type:'application/pdf'});
  if(navigator.canShare&&navigator.canShare({files:[f]}))await navigator.share({files:[f]}).catch(()=>{});
  else{const a=E('a'),u=URL.createObjectURL(blob);a.href=u;a.download=f.name;a.click();setTimeout(()=>URL.revokeObjectURL(u),60000)}
 }catch(e){alert('فشل التصدير: '+e.message)}
 b.textContent='⬇ PDF';b.disabled=false};
// ---------- tap to hide/show bars (hand tool) ----------
let tp=null;const fsT=()=>{const on=$('#ed').classList.toggle('fs'),d=document;try{if(on&&!d.fullscreenElement)0;else if(!on&&d.fullscreenElement)d.exitFullscreen()}catch(e){}};
document.addEventListener('fullscreenchange',()=>{if(!document.fullscreenElement)$('#ed').classList.remove('fs')});
$('#sc').addEventListener('pointerdown',e=>{tp=e.isPrimary&&S.t==='hand'?{x:e.clientX,y:e.clientY,t:Date.now()}:null});
let lt=null,ftm=0;
$('#sc').addEventListener('pointerup',e=>{if(!tp)return;const d=Math.hypot(e.clientX-tp.x,e.clientY-tp.y),ms=Date.now()-tp.t;tp=null;if(d<10&&ms<400&&!S.pd&&!getSelection().toString()){const n=Date.now();
 if(lt&&n-lt.t<320&&Math.hypot(e.clientX-lt.x,e.clientY-lt.y)<35){clearTimeout(ftm);lt=null;const sc=$('#sc'),sr=sc.getBoundingClientRect(),cx=e.clientX-sr.left,cy=e.clientY-sr.top;zoom(S.z>S.fit*1.2?S.fit:S.fit*2,{cx,cy,fx:sc.scrollLeft+cx,fy:sc.scrollTop+cy});return}
 lt={x:e.clientX,y:e.clientY,t:n};clearTimeout(ftm);ftm=setTimeout(()=>{lt=null;fsT()},280)}});
$('#sc').addEventListener('pointercancel',()=>tp=null);
if('serviceWorker'in navigator&&!(window.Capacitor&&Capacitor.isNativePlatform&&Capacitor.isNativePlatform()))navigator.serviceWorker.register('sw.js').catch(()=>{});
home();

// Text extraction from PDF is intentionally not part of Reader All.
// ---------- toolbar position ----------
const lg=()=>{try{return localStorage.getItem('rapos')}catch(e){}};
function pos(){const L=lg()==='L';$('#ed').classList.toggle('L',L);$('#tp').textContent=L?'⬇':'⬅'}
$('#tp').onclick=()=>{try{localStorage.setItem('rapos',lg()==='L'?'B':'L')}catch(e){}pos()};pos();

// ---------- laser pen (مدة الاختفاء بتتسأل كل مرة تختار الليزر) ----------
function lzAsk(){const old=$('#lzq');old&&old.remove();const cur=S.lzT||7;
 const o=E('div','lzq');o.id='lzq';o.innerHTML='<div class="lzb"><b>الليزر يفضل قد إيه قبل ما يختفي؟</b><div class="lzr">'+[3,5,7,10].map(n=>'<button data-s="'+n+'"'+(n===cur?' class="on"':'')+'>'+n+' ثواني</button>').join('')+'</div></div>';
 o.onclick=e=>{const b=e.target.closest('[data-s]');if(b)S.lzT=+b.dataset.s;if(b||e.target===o)o.remove()};document.body.append(o)}
document.querySelectorAll('[data-t="lz"]').forEach(b=>b.addEventListener('click',lzAsk));
let lzR=0;
function lzLoop(){if(lzR)return;lzR=1;const f=()=>{const n=Date.now();let al=0;S.pg.forEach(el=>{if(!el.lz||!el.lz.length)return;el.lz=el.lz.filter(o=>!o.end||n-o.end<(o.d||7000));al+=el.lz.length;ink(el.i)});if(al)requestAnimationFrame(f);else lzR=0};requestAnimationFrame(f)}
function lzd(c,el,k){const n=Date.now();el.lz.forEach(o=>{const D=o.d||7000,F=Math.min(1500,D),P=o.p,T=o.t||[],L=P.length-1;
 const al=j=>T[j]==null||(!o.end&&j>=L-1)?1:Math.max(0,Math.min(1,(D-(n-T[j]))/F));
 c.save();c.lineCap=c.lineJoin='round';c.shadowColor=o.c;
 for(let s=0;s<Math.max(1,L);s+=4){const e=Math.min(L,s+4),a=al(Math.min(L,s+2));if(a<=0)continue;c.globalAlpha=a;c.beginPath();
  for(let j=s;j<=e;j++)j===s?c.moveTo(P[j][0]*k,P[j][1]*k):c.lineTo(P[j][0]*k,P[j][1]*k);
  if(L===0)c.lineTo(P[0][0]*k+.1,P[0][1]*k);
  c.shadowBlur=16*k;c.strokeStyle=o.c;c.lineWidth=(o.w+3)*k;c.stroke();c.shadowBlur=0;c.strokeStyle='#fff';c.lineWidth=Math.max(1,o.w*.45)*k;c.stroke()}
 c.restore()})}
// ---------- pages / bookmarks / highlights panel ----------
let tab='pg';
function updBm(){const b=$('#bm');if(b)b.classList.toggle('act',!!(S.d&&(S.d.bm||[]).includes(S.cp)))}
function tgl(i){const b=S.d.bm=S.d.bm||[],x=b.indexOf(i);x<0?b.push(i):b.splice(x,1);save();updBm()}
function go2(i){$('#pl').hidden=true;S.pg[i].scrollIntoView()}
async function th(d){const i=d.i,c=d.c,k=c.width/S.pg[i].dim[0],x=c.getContext('2d');await bgDraw(i,x,k);(S.d.s[i]||[]).forEach(o=>ds(x,o,k))}
function plR(){const b=$('#plb');b.innerHTML='';b.classList.toggle('l',tab==='hl');document.querySelectorAll('#plh [data-p]').forEach(x=>x.classList.toggle('on',x.dataset.p===tab));
 if(tab==='hl'){const L=[];S.pg.forEach((_,i)=>(S.d.s[i]||[]).forEach(o=>o.t==='hr'&&o.tx&&o.tx.trim()&&L.push([i,o])));
  const cb=E('button');cb.textContent='📋 نسخ الكل ('+L.length+')';cb.onclick=()=>navigator.clipboard.writeText(L.map(([i,o])=>'• '+o.tx.trim()+' (ص'+(i+1)+')').join('\n\n'));b.append(cb);
  if(!L.length){const p=E('p');p.textContent='حدّد نص بزر 🖍 واضغط لون هايلايت، وهيتجمع هنا.';b.append(p)}
  L.forEach(([i,o])=>{const d=E('div','hi');d.style.borderColor=o.c;d.textContent=o.tx.trim()+' — ص'+(i+1);d.onclick=()=>go2(i);b.append(d)});return}
 S.plio&&S.plio.disconnect();const io=new IntersectionObserver(es=>es.forEach(e=>{if(e.isIntersecting){io.unobserve(e.target);th(e.target)}}),{root:b}),bm=S.d.bm||[];S.plio=io;
 S.pg.forEach((el,i)=>{if(tab==='bm'&&!bm.includes(i))return;const d=E('div','th'),c=E('canvas'),s=E('span','st');d.i=i;d.c=c;c.width=220;c.height=Math.round(220*el.dim[1]/el.dim[0]);
  s.textContent=bm.includes(i)?'⭐':'☆';s.onclick=e=>{e.stopPropagation();tgl(i);s.textContent=(S.d.bm||[]).includes(i)?'⭐':'☆'};
  const l=E('div');l.textContent='ص '+(i+1);d.append(c,s,l);d.onclick=()=>go2(i);b.append(d);io.observe(d)});
 if(!b.children.length){const p=E('p');p.textContent='مفيش علامات. اضغط ☆ فوق عشان تعلّم الصفحة.';b.append(p)}}
$('#pb').onclick=()=>{$('#pl').hidden=false;plR()};$('#plx').onclick=()=>$('#pl').hidden=true;
document.querySelectorAll('#plh [data-p]').forEach(x=>x.onclick=()=>{tab=x.dataset.p;plR()});
$('#bm').onclick=()=>tgl(S.cp||0);
// ---------- complete backup / restore (Office files, notes, attachments and app preferences) ----------
async function svF(blob,name){const f=new File([blob],name);if(navigator.canShare&&navigator.canShare({files:[f]}))await navigator.share({files:[f]}).catch(()=>{});else{const a=E('a'),u=URL.createObjectURL(blob);a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),60000)}}
const BACKUP_LOCAL_PREFIXES=['ink','readerall.'];
function backupLocalSettings(){const out={};try{for(let i=0;i<localStorage.length;i++){const k=localStorage.key(i);if(k&&BACKUP_LOCAL_PREFIXES.some(p=>k.startsWith(p)))out[k]=localStorage.getItem(k)}}catch(e){}return out}
const backupKeyOK=k=>typeof k==='string'&&k.length>0&&k.length<=240&&!/[\u0000-\u001f]/.test(k);
async function readBackupFile(key){return await tx('f','readonly',s=>s.get(key))}
$('#bkx').onclick=async()=>{try{
 window.raBusy&&raBusy(true,'جاري تجهيز نسخة كاملة من ملفاتك…');
 const ds=await tx('d','readonly',s=>s.getAll()),records=[],parts=[];
 for(const original of ds){if(!original||original.k==='pdf')continue;
  const d=JSON.parse(JSON.stringify(original)),files=[],seen=new Set();
  const addFile=async key=>{if(!backupKeyOK(key)||seen.has(key))return;const value=await readBackupFile(key);if(value==null)return;
   const blob=value instanceof Blob?value:new Blob([value]);files.push({k:key,n:blob.size});parts.push(blob);seen.add(key)};
  // The primary Office document is stored under its document ID in object store "f".
  await addFile(d.id);
  if(['ppt','doc','xls'].includes(d.k)&&!seen.has(d.id))throw new Error('الملف الأساسي غير موجود: '+(d.n||d.id));
  // Additional binary attachments (including cached MP3 audio) are referenced by d.au.
  for(const a of Array.isArray(d.au)?d.au:[])if(a&&backupKeyOK(a.k))await addFile(a.k);
  records.push({d,files});
 }
 const payload={format:'ReaderAllBackup',version:2,createdAt:new Date().toISOString(),records,local:backupLocalSettings()};
 const header=new TextEncoder().encode(JSON.stringify(payload));if(header.length>64*1048576)throw new Error('بيانات النسخة كبيرة جدًا');
 const prefix=new ArrayBuffer(4);new DataView(prefix).setUint32(0,header.length,true);
 window.raBusy&&raBusy(false);
 await svF(new Blob([prefix,header,...parts],{type:'application/octet-stream'}),'readerall-backup-'+new Date().toISOString().slice(0,10)+'.rabak');
 }catch(e){window.raBusy&&raBusy(false);alert('فشل النسخ الاحتياطي: '+(e&&e.message||e))}};
$('#bki').onchange=async e=>{const f=e.target.files[0];e.target.value='';if(!f)return;
 try{
  if(f.size<6)throw new Error('ملف النسخة فارغ أو تالف');
  const first=await f.slice(0,4).arrayBuffer(),len=new DataView(first).getUint32(0,true);
  if(!(len>1&&len<=f.size-4&&len<64*1048576))throw new Error('حجم بيانات النسخة غير صالح');
  let parsed;try{parsed=JSON.parse(new TextDecoder().decode(await f.slice(4,4+len).arrayBuffer()))}catch(_){throw new Error('تعذر قراءة فهرس النسخة')}
  let records=[],local={},legacy=false;
  if(Array.isArray(parsed)){
   // Backward compatibility with the older notes-only .rabak format.
   legacy=true;records=parsed.map(x=>({d:x&&x.d,files:[...(x&&x.n?[{k:x.d&&x.d.id,n:x.n}]:[]),...((x&&x.au)||[])]}));
  }else if(parsed&&parsed.format==='ReaderAllBackup'&&parsed.version===2&&Array.isArray(parsed.records)){
   records=parsed.records;local=parsed.local&&typeof parsed.local==='object'?parsed.local:{};
  }else throw new Error('صيغة النسخة الاحتياطية غير معروفة');
  let offset=4+len;const prepared=[],allKeys=new Set();
  for(const rec of records){const d=rec&&rec.d;if(!d||typeof d.id!=='string'||!backupKeyOK(d.id)||typeof d.n!=='string'||typeof d.k!=='string')throw new Error('بيانات أحد الملفات غير صالحة');
   if(d.k==='pdf')continue; // Reader All intentionally does not restore PDF documents.
   if(!['note','ppt','doc','xls'].includes(d.k))throw new Error('نوع ملف غير مدعوم داخل النسخة: '+d.k);
   const files=Array.isArray(rec.files)?rec.files:[];let clean=JSON.parse(JSON.stringify(d));clean.au=[];const blobs=[];
   for(const item of files){if(!item||!backupKeyOK(item.k)||!Number.isSafeInteger(item.n)||item.n<0)throw new Error('بيانات مرفق غير صالحة');
    if(allKeys.has(item.k))throw new Error('مفتاح ملف مكرر داخل النسخة');allKeys.add(item.k);
    if(offset+item.n>f.size)throw new Error('بيانات الملفات ناقصة أو تالفة');
    const blob=f.slice(offset,offset+item.n);offset+=item.n;blobs.push({k:item.k,data:await blob.arrayBuffer()});
    if(item.k!==clean.id)clean.au.push({k:item.k});
   }
   // Office records must include their primary binary payload.
   if(['ppt','doc','xls'].includes(clean.k)&&!blobs.some(x=>x.k===clean.id))throw new Error('ملف Office الأساسي غير موجود في النسخة: '+clean.n);
   prepared.push({d:clean,blobs});
  }
  if(offset!==f.size)throw new Error('حجم النسخة لا يطابق البيانات المسجلة');
  window.raBusy&&raBusy(true,'جاري استعادة الملفات والإعدادات…');
  const existing=await tx('d','readonly',s=>s.getAll()),replaced=new Set(prepared.map(x=>x.d.id)),keepKeys=new Set();
  for(const rec of prepared)for(const item of rec.blobs)keepKeys.add(item.k);
  for(const old of existing)if(!replaced.has(old.id)){keepKeys.add(old.id);for(const a of old.au||[])if(a&&a.k)keepKeys.add(a.k)}
  const staleKeys=new Set();for(const old of existing)if(replaced.has(old.id)){staleKeys.add(old.id);for(const a of old.au||[])if(a&&a.k)staleKeys.add(a.k)}
  const db=await DB;
  await new Promise((resolve,reject)=>{const t=db.transaction(['d','f'],'readwrite'),ds=t.objectStore('d'),fs=t.objectStore('f');
   for(const key of staleKeys)if(!keepKeys.has(key))fs.delete(key);
   for(const rec of prepared){ds.put(rec.d);for(const item of rec.blobs)fs.put(item.data,item.k)}
   t.oncomplete=resolve;t.onerror=()=>reject(t.error||new Error('تعذرت كتابة البيانات'));t.onabort=()=>reject(t.error||new Error('تم إلغاء الاستعادة'));
  });
  // Restore only Reader All's own preferences and settings; do not touch unrelated site data.
  if(!legacy){for(const [k,v] of Object.entries(local)){if(typeof k==='string'&&BACKUP_LOCAL_PREFIXES.some(p=>k.startsWith(p))&&typeof v==='string'){try{localStorage.setItem(k,v)}catch(_){}}}}
  window.raBusy&&raBusy(false);alert('تم استرجاع '+prepared.length+' ملف'+(legacy?' من نسخة قديمة':' مع المرفقات والإعدادات'));home();
 }catch(err){window.raBusy&&raBusy(false);alert('تعذر استرجاع النسخة: '+(err&&err.message||err))}
};

// ---------- snap, images, menu, tools ----------
function snap(o){const a=o.p[0],b=o.p[o.p.length-1],xs=o.p.map(q=>q[0]),ys=o.p.map(q=>q[1]),x0=Math.min(...xs),x1=Math.max(...xs),y0=Math.min(...ys),y1=Math.max(...ys);
 if(Math.hypot(a[0]-b[0],a[1]-b[1])<.25*Math.hypot(x1-x0,y1-y0)&&x1-x0>20&&y1-y0>20){o.t='ell';o.p=[[x0,y0],[x1,y1]]}else{o.t='line';o.p=[a,b]}}
const IM={};function im(s){let i=IM[s];if(!i){i=IM[s]=new Image();i.onload=()=>S.pg.forEach(e=>e.ink.width>1&&ink(e.i));i.src=s}return i}
$('#imf').onchange=async e=>{const f=e.target.files[0];e.target.value='';$('#mm').hidden=true;if(!f)return;const u=URL.createObjectURL(f),i=new Image();await new Promise(r=>{i.onload=r;i.onerror=r;i.src=u});URL.revokeObjectURL(u);if(!i.naturalWidth)return alert('الصورة دي مش صالحة');
 const m=Math.min(1,1000/Math.max(i.naturalWidth,i.naturalHeight)),c=E('canvas');c.width=i.naturalWidth*m;c.height=i.naturalHeight*m;c.getContext('2d').drawImage(i,0,0,c.width,c.height);
 placeImg(c.toDataURL('image/jpeg',.82),c.height/c.width)};
// وضع تحديد مكان الصورة: اسحبها لأي مكان، كبّر/صغّر من الركن، وبعدين اضغط تثبيت
function placeImg(src,ar){const old=$('#plc');if(old)old.remove();const sc=$('#sc').getBoundingClientRect(),el0=S.pg[S.cp||0]||S.pg[0];
 let w=Math.min(sc.width*.7,(el0?el0.dim[0]*.5*S.z:200)),x=sc.left+(sc.width-w)/2,y=sc.top+(sc.height-w*ar)/2;
 const b=E('div','',`<img draggable="false"><i class="rz"></i><div class="pht">اسحب الصورة لمكانها · كبّر من الركن ↘ · اسحب برّه الصورة عشان تنزّل الصفحة</div><div class="pb"><button class="p ok">✓ تثبيت هنا</button><button class="no">✕ إلغاء</button></div>`);
 b.id='plc';b.querySelector('img').src=src;
 const put=()=>{b.style.left=x+'px';b.style.top=y+'px';b.style.width=w+'px';b.style.height=w*ar+'px'};put();document.body.append(b);
 let g=null;
 b.onpointerdown=e=>{if(e.target.closest('.pb'))return;e.preventDefault();b.setPointerCapture(e.pointerId);g={r:e.target.classList.contains('rz'),x:e.clientX,y:e.clientY,ox:x,oy:y,ow:w}};
 b.onpointermove=e=>{if(!g)return;const dx=e.clientX-g.x,dy=e.clientY-g.y;if(g.r)w=Math.max(40,g.ow+dx);else{x=g.ox+dx;y=g.oy+dy}put()};
 b.onpointerup=b.onpointercancel=()=>{g=null};
 b.querySelector('.no').onclick=()=>b.remove();
 b.querySelector('.ok').onclick=()=>{const r=b.getBoundingClientRect(),cx=r.left+r.width/2,cy=r.top+r.height/2;
  const pi=S.pg.findIndex(el=>{const q=el.getBoundingClientRect();return cx>=q.left&&cx<=q.right&&cy>=q.top&&cy<=q.bottom});
  if(pi<0){alert('حرّك الصورة فوق صفحة الأول');return}
  const q=S.pg[pi].getBoundingClientRect();b.remove();
  add(pi,{t:'im',src,p:[[(r.left-q.left)/S.z,(r.top-q.top)/S.z]],w:r.width/S.z,h:r.height/S.z});S.t='mv';sync()}}
$('#mn').onclick=()=>$('#mm').hidden=!$('#mm').hidden;
$('#m-po').onclick=()=>{S.po=!S.po;$('#m-po').textContent='🖐 وضع القلم فقط: '+(S.po?'شغال':'إيقاف');$('#m-po').classList.toggle('on',S.po)};
$('#m-nt').onclick=()=>{S.d.nt=!S.d.nt;$('#ed').classList.toggle('nt',S.d.nt);save()};
let tmI=0;$('#m-tm').onclick=()=>{$('#mm').hidden=true;const t=$('#tmr');if(tmI){clearInterval(tmI);tmI=0;t.hidden=true;return}const end=Date.now()+25*60000;t.hidden=false;
 tmI=setInterval(()=>{const s=Math.max(0,Math.round((end-Date.now())/1000));t.textContent='⏱ '+String(s/60|0).padStart(2,'0')+':'+String(s%60).padStart(2,'0');if(!s){clearInterval(tmI);tmI=0;t.hidden=true;navigator.vibrate&&navigator.vibrate([300,100,300]);alert('خلصت 25 دقيقة. خد استراحة 5 دقايق')}},500)};
$('#tmr').onclick=()=>$('#m-tm').click();

function lib(l,a){if(!$('#ls')){const i=E('input');i.id='ls';i.type='search';i.placeholder='🔍 ابحث في ملفاتك';i.oninput=home;l.before(i);const so=E('select');so.id='so';[['u','الأحدث'],['n','الاسم'],['sz','الحجم']].forEach(([v,t])=>{const o=E('option');o.value=v;o.textContent=t;so.append(o)});so.value=SO();so.onchange=()=>{try{localStorage.setItem('raso',so.value)}catch(e){}home()};i.after(so);const t=E('div');t.id='tgs';l.before(t)}
 const t=$('#tgs');t.innerHTML='';[...new Set(a.map(d=>d.tag).filter(Boolean))].forEach(x=>{const b=E('button','chip'+(S.tg===x?' on':''));b.textContent='📁 '+x;b.onclick=()=>{S.tg=S.tg===x?'':x;home()};t.append(b)})}
