// Reader All — Word Pro: tables, images, Arabic/RTL direction, page setup (size, margins),
// headers/footers, page numbers and page breaks. Loaded after office.js; office.js calls window.WordPro.
(()=>{
'use strict';
const E=(t,c,h)=>{const e=document.createElement(t);if(c)e.className=c;if(h!=null)e.innerHTML=h;return e};
const X=()=>window.__of||{};
const toast=t=>{try{X().toast(t)}catch(e){}};
const ask=(q,d)=>window.raAsk?window.raAsk(q,d):Promise.resolve(prompt(q,d));
const AR=/[\u0590-\u08FF\uFB1D-\uFEFF]/;
const strong=t=>{for(const ch of t){if(AR.test(ch))return 1;if(/[A-Za-z]/.test(ch))return 0}return -1};
const lsGet=(k,d)=>{try{const v=localStorage.getItem(k);return v==null?d:v}catch(e){return d}};
const lsSet=(k,v)=>{try{localStorage.setItem(k,v)}catch(e){}};

// ---------------- page model ----------------
const SIZES={A4:[11906,16838],A5:[8391,11906],Letter:[12240,15840],Legal:[12240,20160]};
const cm2tw=v=>Math.round(v*566.929),tw2cm=t=>Math.round(t/566.929*100)/100;
const defPage=()=>({size:'A4',landscape:false,mt:1440,mb:1440,ml:1440,mr:1440,hd:708,fd:708});
const defHF=()=>({header:'',footer:'',ha:'center',fa:'center',pn:'none',pa:'center',pf:'n',nf:'decimal',ps:1,skip1:false});
function normPage(o){const p=Object.assign(defPage(),o||{});if(o&&o.margin&&o.mt==null){p.mt=p.mb=p.ml=p.mr=o.margin}delete p.margin;return p}
function dims(p){let d;if(p.size==='custom'&&p.cw&&p.ch)d=[p.cw,p.ch];else d=SIZES[p.size]||SIZES.A4;return p.landscape?[d[1],d[0]]:[d[0],d[1]]}
let CUR=null; // the mounted editor: {O,ed,pg,...}
const contentPx=()=>{const p=normPage(CUR&&CUR.O.wordPage);return Math.max(100,(dims(p)[0]-p.ml-p.mr)/15)};

// ---------------- number formatting (preview only) ----------------
const hindi=s=>String(s).replace(/\d/g,d=>'٠١٢٣٤٥٦٧٨٩'[d]);
const roman=n=>{let s='';for(const[v,c]of [[1000,'M'],[900,'CM'],[500,'D'],[400,'CD'],[100,'C'],[90,'XC'],[50,'L'],[40,'XL'],[10,'X'],[9,'IX'],[5,'V'],[4,'IV'],[1,'I']])while(n>=v){s+=c;n-=v}return s};
const fnum=(hf,n)=>hf.nf==='hindiNumbers'?hindi(n):hf.nf==='upperRoman'?roman(n):hf.nf==='lowerRoman'?roman(n).toLowerCase():String(n);
const fpn=(hf,ar,n,tot)=>{const a=fnum(hf,n),t=fnum(hf,tot);return hf.pf==='pg'?(ar?'صفحة ':'Page ')+a:hf.pf==='pgof'?(ar?'صفحة '+a+' من '+t:'Page '+a+' of '+t):hf.pf==='dash'?'- '+a+' -':a};

// ---------------- export helpers (used by office.js htmlToDocx) ----------------
const hex=c=>{const m=/rgba?\((\d+)[, ]+(\d+)[, ]+(\d+)(?:[, /]+([\d.]+))?/.exec(c||'');if(!m)return null;if(m[4]!==undefined&&+m[4]===0)return null;return [m[1],m[2],m[3]].map(x=>(+x).toString(16).padStart(2,'0')).join('').toUpperCase()};
const jc=(dx,vis,bidi)=>{if(vis==='center')return dx.AlignmentType.CENTER;if(bidi)return vis==='right'?dx.AlignmentType.LEFT:dx.AlignmentType.RIGHT;return vis==='right'?dx.AlignmentType.RIGHT:dx.AlignmentType.LEFT};
function pnRuns(dx,hf,ar){const C=dx.PageNumber.CURRENT,T=dx.PageNumber.TOTAL_PAGES,t=s=>new dx.TextRun({text:s,rightToLeft:AR.test(s)||undefined}),f=x=>new dx.TextRun({children:[x]});
 if(hf.pf==='pg')return [t(ar?'صفحة ':'Page '),f(C)];
 if(hf.pf==='pgof')return [t(ar?'صفحة ':'Page '),f(C),t(ar?' من ':' of '),f(T)];
 if(hf.pf==='dash')return [t('- '),f(C),t(' -')];
 return [f(C)]}
function hfParas(dx,hf,which,ar){const text=which==='header'?hf.header:hf.footer,al=which==='header'?hf.ha:hf.fa,out=[];
 String(text||'').split(/\r?\n/).forEach(line=>{if(!line.trim())return;const s=strong(line),bidi=s===1||(s<0&&ar);
  out.push(new dx.Paragraph({children:[new dx.TextRun({text:line,rightToLeft:AR.test(line)||undefined})],bidirectional:bidi||undefined,alignment:jc(dx,al,bidi)}))});
 if(hf.pn===which)out.push(new dx.Paragraph({children:pnRuns(dx,hf,ar),bidirectional:ar||undefined,alignment:jc(dx,hf.pa,ar)}));
 if(!out.length)out.push(new dx.Paragraph({}));return out}
// -> {properties, headers?, footers?} for a docx section
function section(dx,O){const p=normPage(O&&O.wordPage),hf=Object.assign(defHF(),(O&&O.wordHF)||{}),docRtl=!!O&&O.wordDir==='rtl';
 const sample=CUR&&O&&CUR.O===O?(CUR.ed.textContent||'').slice(0,400):'';
 const ar=docRtl||AR.test((hf.header||'')+(hf.footer||'')+sample);
 const d0=p.size==='custom'&&p.cw&&p.ch?[p.cw,p.ch]:(SIZES[p.size]||SIZES.A4);
 const props={page:{size:{width:d0[0],height:d0[1],orientation:p.landscape?dx.PageOrientation.LANDSCAPE:dx.PageOrientation.PORTRAIT},margin:{top:p.mt,bottom:p.mb,left:p.ml,right:p.mr,header:p.hd,footer:p.fd}}};
 if(hf.pn!=='none'||hf.nf!=='decimal'||+hf.ps!==1)props.page.pageNumbers={start:Number.isFinite(+hf.ps)?+hf.ps:1,formatType:hf.nf};
 const wantH=!!(hf.header&&hf.header.trim())||hf.pn==='header',wantF=!!(hf.footer&&hf.footer.trim())||hf.pn==='footer',res={properties:props};
 if(hf.skip1&&(wantH||wantF))props.titlePage=true;
 if(wantH){res.headers={default:new dx.Header({children:hfParas(dx,hf,'header',ar)})};if(props.titlePage)res.headers.first=new dx.Header({children:[new dx.Paragraph({})]})}
 if(wantF){res.footers={default:new dx.Footer({children:hfParas(dx,hf,'footer',ar)})};if(props.titlePage)res.footers.first=new dx.Footer({children:[new dx.Paragraph({})]})}
 return res}

// table grid: rows x cols matrix of cells (spans expanded)
function grid(t){const g=[],rows=[...t.rows];rows.forEach((tr,r)=>{g[r]=g[r]||[];let c=0;[...tr.cells].forEach(td=>{while(g[r][c])c++;const rs=Math.max(1,td.rowSpan||1),cs=Math.max(1,td.colSpan||1);for(let i=0;i<rs&&r+i<rows.length;i++){g[r+i]=g[r+i]||[];for(let j=0;j<cs;j++)g[r+i][c+j]=td}c+=cs})});return g}
const gridPos=g=>{const m=new Map();g.forEach((row,r)=>{for(let c=0;c<row.length;c++){const td=row[c];if(td&&!m.has(td))m.set(td,{r,c})}});return m};
const tableRtl=(t)=>{const dc=t.closest('[dir]'),v=dc&&dc.getAttribute('dir');if(v==='rtl')return true;if(v==='ltr')return false;return strong(t.textContent||'')===1};
// column widths from the live layout (twips); null when the table is not on screen
function measure(t,root){if(!t.isConnected||!root||!root.clientWidth)return null;
 const g=grid(t),nc=Math.max(0,...g.map(r=>r.length));if(!nc)return null;
 const w=Array(nc).fill(0);g.forEach(row=>{for(let c=0;c<nc;c++){const td=row[c];if(td&&td.colSpan===1&&td.offsetWidth>w[c])w[c]=td.offsetWidth}});
 const known=w.filter(x=>x>0);if(!known.length)return null;const avg=known.reduce((a,b)=>a+b,0)/known.length;for(let c=0;c<nc;c++)if(!w[c])w[c]=avg;
 const sum=w.reduce((a,b)=>a+b,0),cs=getComputedStyle(root),ew=root.clientWidth-(parseFloat(cs.paddingLeft)||0)-(parseFloat(cs.paddingRight)||0);
 const p=normPage(CUR&&CUR.O.wordPage),ctw=dims(p)[0]-p.ml-p.mr,ratio=t.dataset.tw==='auto'?Math.min(1,sum/Math.max(1,ew)):1,total=Math.round(ctw*ratio);
 const cols=w.map(x=>Math.max(200,Math.round(x/sum*total))),tot=cols.reduce((a,b)=>a+b,0),pos=gridPos(g);
 return {cols,total:tot,cellW:td=>{const q=pos.get(td);if(!q)return 0;let s=0;for(let i=0;i<td.colSpan&&q.c+i<nc;i++)s+=cols[q.c+i];return s}}}
function cellBorders(dx,td,rtl){const cs=getComputedStyle(td),BS=dx.BorderStyle,one=s=>{const w=parseFloat(cs['border'+s+'Width'])||0,st=cs['border'+s+'Style'];
  if(!w||st==='none'||st==='hidden')return {style:BS.NONE,size:0,color:'auto'};
  return {style:st==='dashed'?BS.DASHED:st==='dotted'?BS.DOTTED:st==='double'?BS.DOUBLE:BS.SINGLE,size:Math.max(2,Math.min(96,Math.round(w*6))),color:hex(cs['border'+s+'Color'])||'444444'}};
 const o={top:one('Top'),bottom:one('Bottom')};if(rtl){o.start=one('Right');o.end=one('Left')}else{o.left=one('Left');o.right=one('Right')}return o}
const tableAlign=(dx,t,rtl)=>{const a=t.dataset&&t.dataset.al;if(!a)return undefined;return jc(dx,a==='c'?'center':a==='r'?'right':'left',rtl)};

// ---------------- UI ----------------
function modal(title,build){const w=E('div','shg wpm'),box=E('div','shb wpb','<div class="shh"><b></b><button type="button" aria-label="إغلاق">✕</button></div>');
 box.querySelector('b').textContent=title;const close=()=>w.remove();box.querySelector('button').onclick=close;w.addEventListener('mousedown',e=>{if(e.target===w){e.preventDefault();close()}});
 w.append(box);document.body.append(w);build(box,close);return {w,box,close}}
const btn=(label,fn,cls)=>{const b=E('button','wpbtn'+(cls?' '+cls:''));b.type='button';b.innerHTML=label;b.onmousedown=e=>e.preventDefault();b.onclick=()=>fn(b);return b};
const head=(box,t)=>{const h=E('h5');h.textContent=t;box.append(h);return h};
const gridEl=box=>{const g=E('div','wpgrid');box.append(g);return g};
const row=(parent,label,el)=>{const r=E('label','wpr');r.append(E('span','',label),el);parent.append(r);return el};
const sel=(opts,val,on)=>{const s=E('select');opts.forEach(([v,l])=>{const o=E('option');o.value=v;o.textContent=l;s.append(o)});s.value=val;s.onchange=()=>on(s.value);return s};
const numI=(val,on,step,min,max)=>{const i=E('input');i.type='number';i.inputMode='decimal';i.step=step||'0.01';if(min!=null)i.min=min;if(max!=null)i.max=max;i.value=val;i.oninput=()=>on(parseFloat(i.value));return i};
const chk=(val,on)=>{const i=E('input');i.type='checkbox';i.checked=!!val;i.onchange=()=>on(i.checked);return i};
const BRD_COLORS=['#444444','#000000','#7F7F7F','#2F5597','#C00000','#548235'];
const FILLS=['#FFFFFF','#F2F2F2','#D9E2F3','#FFF2CC','#E2EFDA','#FCE4D6','#FFFF00'];
const isBlank=el=>!(el.textContent||'').replace(/\u00a0/g,' ').trim()&&!el.querySelector('img');

function mount(ctx){
 const {ed,O}=ctx,pg=ed.parentElement;
 CUR={O,ed,pg};
 O.wordPage=normPage(O.wordPage);O.wordHF=Object.assign(defHF(),O.wordHF||{});
 if(!O.wordDir)O.wordDir=ctx.newDoc?lsGet('ra_wdir','auto'):'auto';
 const range=()=>ctx.range&&ctx.range(),elOf=n=>n&&(n.nodeType===1?n:n.parentElement);
 const dirty=()=>{O.dirty=1};
 // ---- header / footer bands ----
 const mk=(cls,ph)=>{const b=E('div','wpband '+cls),t=E('textarea','wpta'),n=E('div','wppn');t.rows=1;t.placeholder=ph;t.dir='auto';t.spellcheck=true;b.append(t,n);return {b,t,n}};
 const H=mk('wphd','✎ الرأس — اكتب هنا (بيظهر أعلى كل صفحة)'),F=mk('wpft','✎ التذييل — اكتب هنا (بيظهر أسفل كل صفحة)');
 pg.insertBefore(H.b,ed);pg.append(F.b);
 const grow=t=>{t.style.height='auto';t.style.height=Math.min(140,t.scrollHeight+2)+'px'};
 H.t.oninput=()=>{O.wordHF.header=H.t.value;grow(H.t);dirty()};F.t.oninput=()=>{O.wordHF.footer=F.t.value;grow(F.t);dirty()};
 const guideOn=()=>lsGet('ra_wguide','1')!=='0';
 let tmr=0;
 function refresh(){const p=O.wordPage,hf=O.wordHF,d=dims(p),w=d[0],h=d[1],small=innerWidth<600,k=small?.5:1;
  pg.style.width='min(100%,'+Math.round(w/15)+'px)';
  const L=(p.ml/w*100*k).toFixed(2),R=(p.mr/w*100*k).toFixed(2);
  ed.style.paddingLeft='max(14px,'+L+'%)';ed.style.paddingRight='max(14px,'+R+'%)';
  [H.b,F.b].forEach(b=>{b.style.paddingLeft='max(14px,'+L+'%)';b.style.paddingRight='max(14px,'+R+'%)'});
  if(H.t.value!==(hf.header||''))H.t.value=hf.header||'';if(F.t.value!==(hf.footer||''))F.t.value=hf.footer||'';H.t.style.textAlign=hf.ha;F.t.style.textAlign=hf.fa;grow(H.t);grow(F.t);
  const sc=(pg.clientWidth||w/15)/(w/15),ch=Math.max(120,(h-p.mt-p.mb)/15*sc);
  ed.style.setProperty('--wpc',ch+'px');ed.classList.toggle('wpguide',guideOn());
  const ar=O.wordDir==='rtl'||AR.test((hf.header||'')+(hf.footer||'')+(ed.textContent||'').slice(0,300)),pages=Math.max(1,Math.ceil(ed.scrollHeight/ch)),start=Number.isFinite(+hf.ps)?+hf.ps:1;
  for(const [b,key] of [[H,'header'],[F,'footer']]){const on=hf.pn===key;b.n.style.display=on?'block':'none';if(on){b.n.textContent=fpn(hf,ar,start,start+pages-1);b.n.style.textAlign=hf.pa}}
  H.b.classList.toggle('has',!!(hf.header||hf.pn==='header'));F.b.classList.toggle('has',!!(hf.footer||hf.pn==='footer'))}
 const later=()=>{clearTimeout(tmr);tmr=setTimeout(()=>{if(CUR&&CUR.ed===ed)refresh()},350)};
 ed.addEventListener('input',later);
  // ---- auto smart direction: each paragraph follows its own language while typing ----
  const autoOn=()=>lsGet('ra_wauto','1')!=='0';
  const leafBlk=n=>{let e=elOf(n);while(e&&e!==ed){if(e.matches&&e.matches(BLOCKS+',td,th')&&!e.querySelector(BLOCKS+',table'))return e;e=e.parentElement}return null};
  ed.addEventListener('input',()=>{if(!autoOn()||O.wordDir!=='auto')return;const s=getSelection(),b=s.rangeCount?leafBlk(s.anchorNode):null;if(!b||b.hasAttribute('data-page-break'))return;
   const d=domDir(b.textContent||'');if(!d)return;const c=b.getAttribute('dir');if(c&&!b.dataset.ad)return;if(c!==d){b.setAttribute('dir',d);b.dataset.ad='1'}});
 let ro=null;try{ro=new ResizeObserver(later);ro.observe(pg)}catch(e){}
 const applyDir=()=>{if(O.wordDir==='rtl'||O.wordDir==='ltr')ed.setAttribute('dir',O.wordDir);else ed.removeAttribute('dir')};
 applyDir();refresh();
 // existing .docx: read page size/margins/header/footer/direction from the file (mammoth drops them)
 if(!ctx.newDoc&&ctx.buf&&!O.wpInit){O.wpInit=1;readDocx(ctx.buf).then(r=>{if(!r||!CUR||CUR.ed!==ed)return;
  if(r.page)O.wordPage=normPage(Object.assign({},O.wordPage,r.page));if(r.hf)O.wordHF=Object.assign(defHF(),O.wordHF,r.hf);
  if(r.dir==='rtl'){O.wordDir='rtl';applyDir()}refresh()})}

 // ---- selection helpers ----
 const BLOCKS='p,li,h1,h2,h3,h4,h5,h6,blockquote,div';
 const blocksSel=()=>{const r=range();if(!r)return [];return [...ed.querySelectorAll(BLOCKS)].filter(b=>!b.hasAttribute('data-page-break')&&!b.querySelector(BLOCKS)&&r.intersectsNode(b))};
 const cellOf=()=>{const r=range(),n=r?r.startContainer:getSelection().anchorNode,el=elOf(n),td=el&&el.closest&&el.closest('td,th');return td&&ed.contains(td)?td:null};

 // ---- direction ----
 function setDir(d){const bs=blocksSel();if(!bs.length)return toast('حط المؤشر داخل فقرة الأول');
  bs.forEach(b=>{delete b.dataset.ad;if(d==='auto')b.removeAttribute('dir');else b.setAttribute('dir',d);if(/^(left|right)$/.test(b.style.textAlign))b.style.textAlign=''});dirty();toast(d==='rtl'?'الفقرة: عربي (يمين → يسار)':d==='ltr'?'الفقرة: إنجليزي (يسار → يمين)':'الفقرة: تلقائي')}
 // ---- smart alignment: direction + alignment follow each paragraph's dominant language ----
 const domDir=t=>{let a=0,l=0;for(const ch of t){if(AR.test(ch))a++;else if(/[A-Za-z]/.test(ch))l++}if(!a&&!l)return null;if(a!==l)return a>l?'rtl':'ltr';return strong(t)===1?'rtl':'ltr'};
 function smartAlign(all){let bs;
  if(all)bs=[...ed.querySelectorAll(BLOCKS+',td,th')];
  else{const r=range();if(!r)return toast('حط المؤشر داخل فقرة الأول');bs=[...ed.querySelectorAll(BLOCKS+',td,th')].filter(b=>r.intersectsNode(b))}
  bs=bs.filter(b=>!b.hasAttribute('data-page-break')&&!b.querySelector(BLOCKS+',table'));
  let n=0,ar=0,en=0;
  for(const b of bs){const d=domDir(b.textContent||'');if(!d)continue;b.setAttribute('dir',d);b.dataset.ad='1';if(/^(left|right|start|end)?$/.test(b.style.textAlign))b.style.textAlign='';n++;d==='rtl'?ar++:en++}
  if(!n)return toast('مفيش نص أطبّق عليه');dirty();toast('✨ محاذاة ذكية: '+ar+' فقرة عربي ➡ و '+en+' إنجليزي ⬅'+(all?' (المستند كله)':''))}
 function setDocDir(d){O.wordDir=d;applyDir();ed.querySelectorAll('[dir]').forEach(e=>e.removeAttribute('dir'));lsSet('ra_wdir',d);dirty();refresh();toast(d==='rtl'?'المستند كله: عربي (RTL)':d==='ltr'?'المستند كله: إنجليزي (LTR)':'المستند: اتجاه تلقائي')}
 function dirPanel(){modal('🌐 اتجاه الكتابة',box=>{
  head(box,'المستند كله');let g=gridEl(box);
  const cur=O.wordDir;
  g.append(btn('عربي RTL ➡',()=>{setDocDir('rtl');mark()},cur==='rtl'?'on':''),btn('English LTR ⬅',()=>{setDocDir('ltr');mark()},cur==='ltr'?'on':''),btn('تلقائي',()=>{setDocDir('auto');mark()},cur==='auto'?'on':''));
  const mark=()=>{[...g.children].forEach((b,i)=>b.classList.toggle('on',['rtl','ltr','auto'][i]===O.wordDir))};
  head(box,'الفقرة الحالية (أو الفقرات المحددة)');g=gridEl(box);
  g.append(btn('عربي RTL ➡',()=>setDir('rtl')),btn('English LTR ⬅',()=>setDir('ltr')),btn('تلقائي',()=>setDir('auto')));
  head(box,'⚡ محاذاة ذكية تلقائية أثناء الكتابة');g=gridEl(box);
   const ga=g,setAuto=v=>{lsSet('ra_wauto',v?'1':'0');[...ga.children].forEach((b,i)=>b.classList.toggle('on',(i===0)===v));toast(v?'المحاذاة الذكية شغّالة تلقائي':'المحاذاة الذكية التلقائية اتقفلت')};
   g.append(btn('شغّال',()=>setAuto(true),autoOn()?'on':''),btn('مقفول',()=>setAuto(false),autoOn()?'':'on'));
   const an=E('p','wpnote');an.textContent='بتشتغل لما اتجاه المستند «تلقائي». الفقرة اللي بتضبط اتجاهها بإيدك (RTL/LTR) مش بتتغير.';box.append(an);
   head(box,'✨ محاذاة ذكية (حسب لغة كل فقرة)');g=gridEl(box);g.append(btn('الفقرة الحالية',()=>smartAlign(false)),btn('المستند كله',()=>smartAlign(true)));
   head(box,'محاذاة الفقرة');g=gridEl(box);
  for(const [l,c] of [['يمين','justifyRight'],['وسط','justifyCenter'],['يسار','justifyLeft'],['ضبط','justifyFull']])g.append(btn(l,()=>{ctx.restore();document.execCommand(c);dirty()}));
  const tp=E('p','wpnote');tp.textContent='اتجاه الجدول من زر ⚙ الجدول. الحفظ بيحوّل الاتجاه لـ Word بشكل سليم (فقرات bidi).';box.append(tp)})}

 // ---- tables ----
 function insertTable(R,C,o){const dir=o.dir==='auto'?'':' dir="'+o.dir+'"',bd=o.border==='none'?'0':'1px solid #444444',
   cs='border:'+bd+';padding:4px 6px;vertical-align:top',hs=cs+';font-weight:bold;background:#D9E2F3';
  let h='<table'+dir+' data-tw="'+o.width+'" style="border-collapse:collapse;width:'+(o.width==='auto'?'auto':'100%')+'">';
  for(let i=0;i<R;i++){h+='<tr'+(i===0&&o.header?' data-hdr="1"':'')+'>';for(let j=0;j<C;j++)h+='<td style="'+(i===0&&o.header?hs:cs)+'">&nbsp;</td>';h+='</tr>'}
  ctx.restore();document.execCommand('insertHTML',false,h+'</table><p><br></p>');dirty()}
 function tableDialog(){modal('▦ إدراج جدول',(box,close)=>{const v={r:3,c:3,header:true,border:'all',width:'full',dir:O.wordDir==='rtl'?'rtl':'auto'};
  row(box,'عدد الصفوف',numI(v.r,x=>{v.r=x},'1',1,50));row(box,'عدد الأعمدة',numI(v.c,x=>{v.c=x},'1',1,12));
  row(box,'صف عنوان ملوّن',chk(v.header,x=>{v.header=x}));
  row(box,'الحدود',sel([['all','كل الحدود'],['none','بدون حدود']],v.border,x=>{v.border=x}));
  row(box,'عرض الجدول',sel([['full','ملء الصفحة'],['auto','حسب المحتوى']],v.width,x=>{v.width=x}));
  row(box,'اتجاه الجدول',sel([['auto','تلقائي'],['rtl','عربي RTL'],['ltr','English LTR']],v.dir,x=>{v.dir=x}));
  const g=gridEl(box);g.append(btn('إدراج',()=>{const R=Math.max(1,Math.min(50,v.r|0)),C=Math.max(1,Math.min(12,v.c|0));close();insertTable(R,C,v)},'pri'),btn('إلغاء',close))})}
 function tablePanel(){let td=cellOf();if(!td)return alert('حط المؤشر داخل خلية في الجدول الأول.');
  const tbl=td.closest('table');if(!tbl)return;let scope='cell';
  const live=()=>td&&td.isConnected&&tbl.isConnected;
  const info=()=>{const g=grid(tbl),p=gridPos(g).get(td);return {g,p,nc:Math.max(0,...g.map(r=>r.length)),pos:gridPos(g)}};
  const scopeCells=()=>{const {g,p}=info();if(scope==='table')return [...new Set(g.flat())];if(scope==='row')return [...td.parentElement.cells];if(scope==='col')return [...new Set(g.map(r=>r[p.c]).filter(Boolean))];return [td]};
  const mkCell=like=>{const c=document.createElement('td');c.innerHTML='&nbsp;';const s=like&&like.getAttribute('style');if(s)c.setAttribute('style',s);c.style.width='';return c};
  const need=()=>{if(!live()){toast('الخلية اتمسحت — اقفل واختار خلية تاني');return false}return true};
  const addRow=after=>{if(!need())return;const {g,pos,nc,p}=info(),rows=[...tbl.rows],k=after?p.r+td.rowSpan:p.r,tr=document.createElement('tr'),seen=new Set();
   for(let c=0;c<nc;c++){const cell=g[k]&&g[k][c];if(cell&&pos.get(cell).r<k){if(!seen.has(cell)){cell.rowSpan++;seen.add(cell)}continue}tr.append(mkCell(cell||td))}
   if(k>=rows.length)rows[rows.length-1].parentElement.append(tr);else rows[k].parentElement.insertBefore(tr,rows[k]);dirty();toast('تمت إضافة صف')};
  const delRow=()=>{if(!need())return;const {g,pos,p}=info();if(tbl.rows.length<2)return alert('لازم يفضل صف واحد على الأقل.');
   const cells=[...new Set(g[p.r])];if(cells.some(c=>c&&pos.get(c).r===p.r&&c.rowSpan>1))return alert('الصف فيه خلايا مدمجة رأسيًا — اعمل «فصل الخلية» الأول.');
   cells.forEach(c=>{if(c&&pos.get(c).r<p.r)c.rowSpan--});td.parentElement.remove();td=null;dirty();toast('اتمسح الصف')};
  const addCol=after=>{if(!need())return;const {g,pos,nc,p}=info(),k=after?p.c+td.colSpan:p.c,seen=new Set();
   [...tbl.rows].forEach((tr,r)=>{const cell=g[r]&&g[r][k];if(cell&&pos.get(cell).c<k){if(!seen.has(cell)){cell.colSpan++;seen.add(cell)}return}
    const ref=[...tr.cells].find(x=>{const q=pos.get(x);return q&&q.c>=k});tr.insertBefore(mkCell(cell||g[r][p.c]||td),ref||null)});dirty();toast('تمت إضافة عمود')};
  const delCol=()=>{if(!need())return;const {g,nc,p}=info();if(nc<2)return alert('لازم يفضل عمود واحد على الأقل.');const seen=new Set();
   g.forEach(r=>{const cell=r[p.c];if(!cell||seen.has(cell))return;seen.add(cell);if(cell.colSpan>1)cell.colSpan--;else cell.remove()});td=null;dirty();toast('اتمسح العمود')};
  const joinInto=(a,b)=>{if(!isBlank(b)){if(isBlank(a))a.innerHTML=b.innerHTML;else a.insertAdjacentHTML('beforeend','<br>'+b.innerHTML)}b.remove()};
  const mergeRight=()=>{if(!need())return;const {g,pos,p}=info(),t=g[p.r][p.c+td.colSpan];
   if(!t)return alert('مفيش خلية على الجنب تتدمج.');const q=pos.get(t);if(q.r!==p.r||t.rowSpan!==td.rowSpan)return alert('الخلايا لازم تكون بنفس الارتفاع عشان تتدمج.');
   td.colSpan+=t.colSpan;joinInto(td,t);dirty();toast('تم الدمج')};
  const mergeDown=()=>{if(!need())return;const {g,pos,p}=info(),t=g[p.r+td.rowSpan]&&g[p.r+td.rowSpan][p.c];
   if(!t)return alert('مفيش خلية تحت تتدمج.');const q=pos.get(t);if(q.c!==p.c||t.colSpan!==td.colSpan||q.r!==p.r+td.rowSpan)return alert('الخلايا لازم تكون بنفس العرض عشان تتدمج.');
   td.rowSpan+=t.rowSpan;joinInto(td,t);dirty();toast('تم الدمج')};
  const split=()=>{if(!need())return;const {pos,p}=info(),cs=td.colSpan,rs=td.rowSpan;if(cs<2&&rs<2)return toast('الخلية دي مش مدمجة');
   for(let i=1;i<rs;i++){const tr=tbl.rows[p.r+i];if(!tr)continue;const ref=[...tr.cells].find(x=>{const q=pos.get(x);return q&&q.c>p.c});for(let j=0;j<cs;j++)tr.insertBefore(mkCell(td),ref||null)}
   td.rowSpan=1;for(let j=1;j<cs;j++)td.after(mkCell(td));td.colSpan=1;dirty();toast('تم فصل الخلية')};
  const toggleHdr=()=>{if(!need())return;const tr=td.parentElement,on=tr.dataset.hdr!=='1';[...tr.cells].forEach(c=>{c.style.fontWeight=on?'bold':'';c.style.background=on?'#D9E2F3':''});if(on)tr.dataset.hdr='1';else delete tr.dataset.hdr;dirty();toast(on?'الصف الأول بقى صف عنوان (بيتكرر في كل صفحة)':'اتلغى صف العنوان')};
  const setBorders=(mode,color,wpx)=>{if(!tbl.isConnected)return;const {g,pos,nc}=info(),nr=g.length,line=wpx+'px solid '+color,none='0 none';
   for(const [cell,q] of pos){const r1=q.r+cell.rowSpan-1,c1=q.c+cell.colSpan-1,top=q.r===0,bot=r1>=nr-1,st=q.c===0,en=c1>=nc-1;let T,B,S,N;
    if(mode==='all')T=B=S=N=1;else if(mode==='none')T=B=S=N=0;else if(mode==='outer'){T=top;B=bot;S=st;N=en}else if(mode==='inner'){T=!top;B=!bot;S=!st;N=!en}else if(mode==='h'){T=B=1;S=N=0}else{T=B=0;S=N=1}
    cell.style.borderTop=T?line:none;cell.style.borderBottom=B?line:none;cell.style.borderInlineStart=S?line:none;cell.style.borderInlineEnd=N?line:none}
   dirty()};
  let bcol='#444444',bw=1;
  const zebra=()=>{if(!need())return;const rows=[...tbl.rows],on=!rows.slice(1).some((tr,i)=>i%2===0&&tr.cells[0]&&tr.cells[0].style.background);rows.forEach((tr,i)=>{if(i===0&&tr.dataset.hdr==='1')return;[...tr.cells].forEach(c=>{c.style.background=on&&i%2===1?'#F2F2F2':''})});dirty()};
  const eq=()=>{const {nc,pos}=info();for(const [cell] of pos)cell.style.width=(100/nc*cell.colSpan).toFixed(2)+'%';tbl.style.tableLayout='fixed';dirty();toast('الأعمدة متساوية')};
  const eqR=()=>{if(!need())return;const rows=[...tbl.rows];if(rows.length<2)return;const h=Math.max(...rows.map(r=>r.offsetHeight));rows.forEach(r=>{r.style.height=h+'px'});dirty();toast('الصفوف متساوية في الارتفاع')};
  const colW=async()=>{if(!need())return;const {p}=info();if(td.colSpan>1)return alert('اختار خلية من عمود غير مدمج.');const v=await ask('عرض العمود بالسنتيمتر (مثال 4)','4');const cm=parseFloat(v);if(!(cm>0))return;
   const {g}=info();new Set(g.map(r=>r[p.c]).filter(c=>c&&c.colSpan===1)).forEach(c=>{c.style.width=Math.round(cm*37.8)+'px'});tbl.style.tableLayout='fixed';dirty();toast('اتضبط عرض العمود')};
  const apply=f=>()=>{if(!need())return;scopeCells().forEach(f);dirty()};
  const tw=m=>()=>{tbl.dataset.tw=m;tbl.style.width=m==='auto'?'auto':'100%';if(m==='full')tbl.style.tableLayout='';dirty()};
  const tal=a=>()=>{tbl.dataset.al=a;tbl.style.marginLeft=a==='l'?'0':'auto';tbl.style.marginRight=a==='r'?'0':'auto';if(tbl.dataset.tw!=='auto'){tbl.dataset.tw='auto';tbl.style.width='auto'}dirty();toast('محاذاة الجدول (بتظهر لما العرض «حسب المحتوى»)')};
  const tdir=d=>()=>{tbl.setAttribute('dir',d);dirty();toast(d==='rtl'?'الجدول: عربي RTL':'الجدول: English LTR')};
  const del=async()=>{if(!window.raConfirm||await window.raConfirm('تحذف الجدول كله؟')){tbl.remove();td=null;dirty();toast('اتمسح الجدول')}};
  modal('⚙ خصائص الجدول',(box,close)=>{
   const ss=row(box,'تطبيق التنسيق على',sel([['cell','الخلية الحالية'],['row','الصف'],['col','العمود'],['table','الجدول كله']],scope,v=>{scope=v}));
   head(box,'الصفوف والأعمدة');let g=gridEl(box);
   g.append(btn('＋ صف تحت',()=>addRow(true)),btn('＋ صف فوق',()=>addRow(false)),btn('− حذف الصف',delRow),btn('＋ عمود بعد',()=>addCol(true)),btn('＋ عمود قبل',()=>addCol(false)),btn('− حذف العمود',delCol));
   head(box,'دمج الخلايا');g=gridEl(box);g.append(btn('دمج مع اللي جنبها',mergeRight),btn('دمج مع اللي تحتها',mergeDown),btn('فصل الخلية',split),btn('صف عنوان ⇄',toggleHdr));
   head(box,'الحدود');g=gridEl(box);
   [['all','كل الحدود'],['none','بدون حدود'],['outer','خارجية فقط'],['inner','داخلية فقط'],['h','أفقية فقط'],['v','رأسية فقط']].forEach(([m,l])=>g.append(btn(l,()=>setBorders(m,bcol,bw))));
   const cr=E('div','wprow');BRD_COLORS.forEach(c=>{const s=E('button','wpsw');s.type='button';s.style.background=c;s.onclick=()=>{bcol=c;[...cr.children].forEach(x=>x.classList.toggle('sel',x===s));toast('لون الحدود اتحدد — اختار نوع الحدود')};if(c===bcol)s.classList.add('sel');cr.append(s)});box.append(cr);
   row(box,'سمك الحدود',sel([['0.5','رفيع'],['1','عادي'],['2','سميك'],['3','سميك جدًا']],'1',v=>{bw=+v}));
   head(box,'محاذاة النص داخل الخلية');g=gridEl(box);
   for(const [l,a] of [['يمين','right'],['وسط','center'],['يسار','left'],['ضبط','justify']])g.append(btn(l,apply(c=>{c.style.textAlign=a})));
   for(const [l,a] of [['أعلى','top'],['منتصف','middle'],['أسفل','bottom']])g.append(btn('↕ '+l,apply(c=>{c.style.verticalAlign=a})));
   head(box,'لون الخلية');const fr=E('div','wprow');FILLS.forEach(c=>{const s=E('button','wpsw');s.type='button';s.style.background=c;s.onclick=apply(x=>{x.style.background=c==='#FFFFFF'?'':c});fr.append(s)});box.append(fr);
   head(box,'الجدول');g=gridEl(box);
   g.append(btn('عرض: ملء الصفحة',tw('full')),btn('عرض: حسب المحتوى',tw('auto')),btn('أعمدة متساوية',eq),btn('صفوف متساوية',eqR),btn('عرض العمود (سم)',colW),btn('تلوين صفوف متبادل',zebra));
   g.append(btn('محاذاة: يمين',tal('r')),btn('محاذاة: وسط',tal('c')),btn('محاذاة: يسار',tal('l')),btn('اتجاه: عربي RTL',tdir('rtl')),btn('اتجاه: English LTR',tdir('ltr')));
   g.append(btn('🗑 حذف الجدول',async()=>{await del();if(!tbl.isConnected)close()},'dng'))})}

 // ---- images ----
 async function prepImage(file){const url=await new Promise((ok,no)=>{const r=new FileReader();r.onload=()=>ok(r.result);r.onerror=no;r.readAsDataURL(file)});
  const im=await new Promise((ok,no)=>{const i=new Image();i.onload=()=>ok(i);i.onerror=()=>no(new Error('img'));i.src=url});
  const MAX=1600,w=im.naturalWidth||300,h=im.naturalHeight||200;
  if(/^data:image\/(png|jpe?g|gif|bmp);/i.test(url)&&Math.max(w,h)<=MAX&&file.size<1.6e6)return {url,w,h};
  const k=Math.min(1,MAX/Math.max(w,h)),cw=Math.max(1,Math.round(w*k)),chh=Math.max(1,Math.round(h*k)),c=document.createElement('canvas');c.width=cw;c.height=chh;const g=c.getContext('2d'),keepAlpha=/png|gif|webp|svg/i.test(file.type);
  if(!keepAlpha){g.fillStyle='#fff';g.fillRect(0,0,cw,chh)}g.drawImage(im,0,0,cw,chh);
  return {url:keepAlpha?c.toDataURL('image/png'):c.toDataURL('image/jpeg',.88),w:cw,h:chh}}
 async function insertImageFile(file){try{const r=await prepImage(file),pct=Math.max(10,Math.min(100,Math.round(r.w/contentPx()*100)));
   ctx.restore();document.execCommand('insertHTML',false,'<img data-new="1" data-wp="'+pct+'" src="'+r.url+'" style="width:'+pct+'%;height:auto">');dirty();
   const im=ed.querySelector('img[data-new]');if(im){im.removeAttribute('data-new');selectImg(im)}}catch(e){toast('تعذر إدراج الصورة (الصيغة مش مدعومة)')}}
 function pickImage(){const i=E('input');i.type='file';i.accept='image/*';i.onchange=()=>{const f=i.files&&i.files[0];if(f)insertImageFile(f)};i.click()}
 ed.addEventListener('paste',e=>{const f=[...((e.clipboardData&&e.clipboardData.files)||[])].find(x=>/^image\//.test(x.type));if(f){e.preventDefault();insertImageFile(f)}});
 let img=null,bar=null,lab=null;
 const edW=()=>{const cs=getComputedStyle(ed);return ed.clientWidth-(parseFloat(cs.paddingLeft)||0)-(parseFloat(cs.paddingRight)||0)};
 const curPct=()=>img?(+img.dataset.wp||Math.round(img.getBoundingClientRect().width/Math.max(1,edW())*100)):0;
 const setPct=p=>{if(!img)return;p=Math.max(5,Math.min(100,Math.round(p)));img.style.width=p+'%';img.style.height='auto';img.dataset.wp=p;if(lab)lab.textContent=p+'%';dirty()};
 const alignImg=a=>{if(!img)return;let b=img.closest('p,div,li,td,th,h1,h2,h3,h4,blockquote');if(!b||b===ed||!ed.contains(b)){const p=document.createElement('p');img.parentNode.insertBefore(p,img);p.append(img);b=p}b.style.textAlign=a;dirty()};
 function deselect(){if(img)img.classList.remove('wpsel');img=null;if(bar){bar.remove();bar=null;lab=null}}
 function selectImg(im){if(img&&img!==im)img.classList.remove('wpsel');img=im;im.classList.add('wpsel');
  if(bar)bar.remove();bar=E('div','wpimgbar');lab=E('span','wplab',curPct()+'%');
  const B=(l,f,t)=>{const b=btn(l,f);if(t)b.title=t;bar.append(b)};
  B('➖',()=>setPct(curPct()-10));bar.append(lab);B('➕',()=>setPct(curPct()+10));
  [25,50,75,100].forEach(p=>B(p+'%',()=>setPct(p)));
  B('◀',()=>alignImg('left'),'يسار');B('◼',()=>alignImg('center'),'وسط');B('▶',()=>alignImg('right'),'يمين');
  B('📏 سم',async()=>{const p=normPage(O.wordPage),cw=(dims(p)[0]-p.ml-p.mr)/566.929,v=parseFloat(await ask('عرض الصورة بالسنتيمتر (أقصى '+cw.toFixed(1)+')','8'));if(v>0)setPct(v/cw*100)});
  B('🗑',()=>{const x=img;deselect();x.remove();dirty()},'حذف الصورة');B('✕',deselect);
  document.body.append(bar)}
 const onClick=e=>{const t=e.target,im=t&&t.closest&&t.closest('img');if(im&&ed.contains(im))selectImg(im);else if(img&&!(t.closest&&t.closest('.wpimgbar')))deselect()};
 ed.addEventListener('click',onClick);

 // ---- page / section breaks ----
 function insertBreak(){ctx.restore();const sel=getSelection(),n=sel.rangeCount?sel.anchorNode:null;let top=n&&ed.contains(n)?n:null;while(top&&top.parentNode!==ed)top=top.parentNode;
  const html='<div class="wpbr" data-page-break="1" contenteditable="false"><span>فاصل صفحة</span></div><p><br></p>',el=elOf(n);
  if(!top){ed.insertAdjacentHTML('beforeend',html)}
  else if(top.nodeType===1&&(/^(TABLE|UL|OL)$/.test(top.tagName)||(el&&el.closest&&el.closest('td,th,li')))){top.insertAdjacentHTML('afterend',html)}
  else document.execCommand('insertHTML',false,html);
  dirty();later();toast('اتضاف فاصل صفحة')}
 function togglePB(){const bs=blocksSel();if(!bs.length)return toast('حط المؤشر داخل الفقرة الأول');const b=bs[0],on=b.dataset.pb!=='1';if(on)b.dataset.pb='1';else delete b.dataset.pb;dirty();toast(on?'الفقرة هتبدأ في صفحة جديدة':'اتلغت «صفحة جديدة قبل الفقرة»')}

 // ---- page dialog ----
 function pageDialog(secName){const p=Object.assign({},O.wordPage),hf=Object.assign({},O.wordHF),secs={};
  modal('📄 إعداد الصفحة',(box,close)=>{
   const S=(k,t)=>{const h=head(box,t);secs[k]=h};
   S('page','حجم واتجاه الصفحة');
   const sizes=[['A4','A4 (21 × 29.7 سم)'],['A5','A5 (14.8 × 21 سم)'],['Letter','Letter'],['Legal','Legal']];if(p.size==='custom')sizes.push(['custom','مخصص (من الملف)']);
   row(box,'المقاس',sel(sizes,p.size,v=>{p.size=v}));row(box,'الاتجاه',sel([['p','رأسي'],['l','أفقي']],p.landscape?'l':'p',v=>{p.landscape=v==='l'}));
   S('margins','الهوامش (سم)');
   const mi={};const set=(k,tw)=>{p[k]=tw;if(mi[k])mi[k].value=tw2cm(tw)};
   const PRE={normal:[1440,1440,1440,1440],narrow:[720,720,720,720],mid:[1440,1440,1080,1080],wide:[1440,1440,2880,2880]};
   row(box,'جاهز',sel([['','— اختار —'],['normal','عادي (2.54)'],['narrow','ضيق (1.27)'],['mid','متوسط (جوانب 1.9)'],['wide','عريض (جوانب 5.08)']],'',v=>{const a=PRE[v];if(a){set('mt',a[0]);set('mb',a[1]);set('ml',a[2]);set('mr',a[3])}}));
   [['mt','أعلى'],['mb','أسفل'],['mr','يمين'],['ml','يسار']].forEach(([k,l])=>{mi[k]=row(box,l,numI(tw2cm(p[k]),v=>{if(v>=0)p[k]=cm2tw(v)},'0.1',0,10))});
   S('hf','الرأس والتذييل');
   const ht=E('textarea','wpta wpbox');ht.rows=2;ht.dir='auto';ht.placeholder='نص الرأس (سطر لكل فقرة)';ht.value=hf.header||'';ht.oninput=()=>{hf.header=ht.value};
   const ft=E('textarea','wpta wpbox');ft.rows=2;ft.dir='auto';ft.placeholder='نص التذييل';ft.value=hf.footer||'';ft.oninput=()=>{hf.footer=ft.value};
   const AL=[['right','يمين'],['center','وسط'],['left','يسار']];
   box.append(ht);row(box,'محاذاة الرأس',sel(AL,hf.ha,v=>{hf.ha=v}));box.append(ft);row(box,'محاذاة التذييل',sel(AL,hf.fa,v=>{hf.fa=v}));
   row(box,'مسافة الرأس من الحافة (سم)',numI(tw2cm(p.hd),v=>{if(v>=0)p.hd=cm2tw(v)},'0.05',0,5));row(box,'مسافة التذييل من الحافة (سم)',numI(tw2cm(p.fd),v=>{if(v>=0)p.fd=cm2tw(v)},'0.05',0,5));
   S('pn','ترقيم الصفحات');
   row(box,'مكان الرقم',sel([['none','بدون ترقيم'],['footer','أسفل الصفحة'],['header','أعلى الصفحة']],hf.pn,v=>{hf.pn=v}));
   row(box,'محاذاة الرقم',sel(AL,hf.pa,v=>{hf.pa=v}));
   row(box,'الشكل',sel([['n','1'],['pg','صفحة 1'],['pgof','صفحة 1 من 5'],['dash','- 1 -']],hf.pf,v=>{hf.pf=v}));
   row(box,'نوع الأرقام',sel([['decimal','1 2 3'],['hindiNumbers','١ ٢ ٣'],['upperRoman','I II III'],['lowerRoman','i ii iii']],hf.nf,v=>{hf.nf=v}));
   row(box,'يبدأ من رقم',numI(hf.ps,v=>{if(Number.isFinite(v))hf.ps=Math.max(0,Math.round(v))},'1',0,9999));
   row(box,'إخفاء الرأس والتذييل والرقم في الصفحة الأولى',chk(hf.skip1,v=>{hf.skip1=v}));
   S('view','العرض');let guide=guideOn();
   row(box,'إظهار حدود الصفحات في المحرر',chk(guide,v=>{guide=v}));
   const nt=E('p','wpnote');nt.textContent='على الموبايل المعاينة بتصغّر الهوامش، لكن الحفظ بيستخدم القيم الحقيقية. الرأس والتذييل بيتكتبوا مباشرة في أعلى وأسفل الصفحة.';box.append(nt);
   const g=gridEl(box);
   g.append(btn('✔ تطبيق',()=>{const d=dims(p);if(p.ml+p.mr>d[0]-1134||p.mt+p.mb>d[1]-1134)return toast('الهوامش كبيرة جدًا — سيب مساحة للكتابة');
     O.wordPage=normPage(p);O.wordHF=Object.assign(defHF(),hf);lsSet('ra_wguide',guide?'1':'0');dirty();refresh();close();toast('اتطبقت إعدادات الصفحة')},'pri'),btn('إلغاء',close));
   if(secName&&secs[secName])setTimeout(()=>secs[secName].scrollIntoView(),60)})}

 const items={
  dir:[{t:'🌐 الاتجاه',f:dirPanel,title:'اتجاه الكتابة'},{t:'✨ ذكية',f:()=>smartAlign(false),title:'محاذاة واتجاه تلقائي حسب لغة الفقرة'},{t:'RTL',f:()=>setDir('rtl'),title:'الفقرة عربي'},{t:'LTR',f:()=>setDir('ltr'),title:'الفقرة إنجليزي'}],
  table:[{t:'▦ جدول',f:tableDialog},{t:'⚙ الجدول',f:tablePanel,title:'خصائص الجدول'}],
  image:[{t:'🖼 صورة',f:pickImage},{t:'⚙ الصورة',f:()=>{if(img&&img.isConnected)selectImg(img);else toast('اضغط على الصورة في المستند الأول')},title:'خصائص الصورة'}],
  page:[{t:'📄 الصفحة',f:()=>pageDialog('page')},{t:'↔ الهوامش',f:()=>pageDialog('margins')},{t:'🔤 رأس/تذييل',f:()=>pageDialog('hf')},{t:'🔢 ترقيم',f:()=>pageDialog('pn')},{t:'⎘ فاصل صفحة',f:insertBreak},{t:'↧ صفحة جديدة قبل الفقرة',f:togglePB}]
 };
 return {items,refresh,cleanup(){clearTimeout(tmr);try{ro&&ro.disconnect()}catch(e){}deselect();ed.removeEventListener('click',onClick);if(CUR&&CUR.ed===ed)CUR=null}}
}

// ---------------- read page setup / header / footer from an existing .docx ----------------
async function readDocx(buf){try{
 const x=X(),z=x.zipIndex(buf),dec=new TextDecoder(),rd=async n=>{const u=await x.zipRead(z,n);return u?dec.decode(u):null};
 const dxml=await rd('word/document.xml');if(!dxml)return null;
 const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',P=s=>new DOMParser().parseFromString(s,'application/xml');
 const d=P(dxml),A=(e,n)=>e?(e.getAttributeNS(W,n)||e.getAttribute('w:'+n)):null,K=(e,n)=>e?[...e.children].find(c=>c.localName===n):null,all=(e,n)=>[...e.getElementsByTagNameNS(W,n)];
 const res={},ps=all(d,'p');let bd=0;
 for(const p of ps){const b=K(K(p,'pPr'),'bidi');if(b){const v=A(b,'val');if(v!=='0'&&v!=='false')bd++}}
 if(ps.length&&bd/ps.length>.5)res.dir='rtl';
 const sps=all(d,'sectPr'),sp=sps[sps.length-1];
 if(sp){const sz=K(sp,'pgSz'),mg=K(sp,'pgMar'),p={};
  if(sz){let w=+A(sz,'w')||0,h=+A(sz,'h')||0;const land=w>h||A(sz,'orient')==='landscape';if(w>h){const t=w;w=h;h=t}
   if(w&&h){p.landscape=land;let nm=null;for(const k in SIZES)if(Math.abs(SIZES[k][0]-w)<80&&Math.abs(SIZES[k][1]-h)<80)nm=k;if(nm)p.size=nm;else{p.size='custom';p.cw=w;p.ch=h}}}
  if(mg)[['mt','top'],['mb','bottom'],['ml','left'],['mr','right'],['hd','header'],['fd','footer']].forEach(([k,n])=>{const s=A(mg,n);if(s!=null&&s!==''&&+s>=0)p[k]=+s});
  res.page=p;const hf={},pn=K(sp,'pgNumType');
  if(pn){const s=A(pn,'start');if(s!=null&&s!=='')hf.ps=+s;const f=A(pn,'fmt');if(f&&['decimal','upperRoman','lowerRoman','hindiNumbers'].includes(f))hf.nf=f}
  if(K(sp,'titlePg'))hf.skip1=true;
  const rels=await rd('word/_rels/document.xml.rels');
  if(rels){const map={};[...P(rels).getElementsByTagName('Relationship')].forEach(r=>{map[r.getAttribute('Id')]=r.getAttribute('Target')});
   for(const [tag,key,jk] of [['headerReference','header','ha'],['footerReference','footer','fa']]){
    const ref=[...sp.children].filter(c=>c.localName===tag).find(c=>(A(c,'type')||'default')==='default');if(!ref)continue;
    const id=ref.getAttributeNS(R,'id')||ref.getAttribute('r:id'),tg=map[id];if(!tg)continue;
    const hx=await rd(tg.startsWith('/')?tg.slice(1):'word/'+tg.replace(/^\.\//,''));if(!hx)continue;
    const hd=P(hx),lines=[];let al=null;
    for(const p of all(hd,'p')){
     const instr=all(p,'instrText').map(e=>e.textContent).join(' ')+' '+all(p,'fldSimple').map(e=>A(e,'instr')||'').join(' '),pr=K(p,'pPr'),bidi=!!K(pr,'bidi'),j=A(K(pr,'jc'),'val');
     const vis=j==='center'?'center':(j==='right'||j==='end')?(bidi?'left':'right'):(bidi?'right':'left');
     if(/\bPAGE\b/.test(instr)){hf.pn=key;hf.pa=vis;hf.pf=/NUMPAGES/.test(instr)?'pgof':'n';const txt=all(p,'t').map(e=>e.textContent).join('');if(hf.pf==='n'&&/(صفحة|page)/i.test(txt))hf.pf='pg';continue}
     const t=all(p,'t').map(e=>e.textContent).join('');if(t.trim()){lines.push(t);if(al==null)al=vis}}
    if(lines.length){hf[key]=lines.join('\n');hf[jk]=al||'center'}}}
  res.hf=hf}
 return res}catch(e){return null}}

// size/margins for PDF export; returns page height / width
function printBox(box){const p=normPage(CUR&&CUR.O.wordPage),d=dims(p);box.style.boxSizing='border-box';box.style.width=Math.round(d[0]/15)+'px';
 box.style.padding=[p.mt,p.mr,p.mb,p.ml].map(v=>Math.round(v/15)+'px').join(' ');const dc=CUR&&CUR.ed.getAttribute('dir');if(dc)box.setAttribute('dir',dc);return d[1]/d[0]}

// ---------------- high-fidelity docx -> html importer (keeps the original formatting) ----------------
async function importDocx(buf){
 const x=X(),z=x.zipIndex(buf),dec=new TextDecoder();
 const rdb=n=>x.zipRead(z,n),rd=async n=>{const u=await rdb(n);return u?dec.decode(u):null};
 const dxml=await rd('word/document.xml');if(!dxml)return null;
 const NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main',RN='http://schemas.openxmlformats.org/officeDocument/2006/relationships',DML='http://schemas.openxmlformats.org/drawingml/2006/main',WPD='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
 const PX=s=>new DOMParser().parseFromString(s,'application/xml');
 const K=(e,n)=>e?[...e.children].find(c=>c.localName===n):null,KS=(e,n)=>e?[...e.children].filter(c=>c.localName===n):[];
 const A=(e,n)=>{if(!e)return null;const v=e.getAttributeNS(NS,n);return v!=null?v:e.getAttribute('w:'+n)};
 const on=e=>{if(!e)return undefined;const v=A(e,'val');return !(v==='0'||v==='false'||v==='off')};
 const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
 const HL={yellow:'#ffff00',green:'#00ff00',cyan:'#00ffff',magenta:'#ff00ff',blue:'#0000ff',red:'#ff0000',darkBlue:'#000080',darkCyan:'#008080',darkGreen:'#008000',darkMagenta:'#800080',darkRed:'#800000',darkYellow:'#808000',darkGray:'#808080',lightGray:'#c0c0c0',black:'#000000',white:'#ffffff'};
 const colr=v=>v&&v!=='auto'&&/^[0-9a-f]{6}$/i.test(v)?'#'+v:null;
 // ----- property parsers
 const rPr=e=>{const o={};if(!e)return o;let t;
  if(t=K(e,'b'))o.b=on(t);if(t=K(e,'bCs'))o.bc=on(t);if(t=K(e,'i'))o.i=on(t);if(t=K(e,'iCs'))o.ic=on(t);
  if(t=K(e,'u')){const v=A(t,'val');o.u=!!v&&v!=='none'}if(t=K(e,'strike'))o.s=on(t);
  if(t=K(e,'color')){const c=colr(A(t,'val'));if(c)o.c=c}
  if(t=K(e,'sz'))o.sz=+A(t,'val')/2;if(t=K(e,'szCs'))o.szc=+A(t,'val')/2;
  if(t=K(e,'rFonts')){const a=A(t,'ascii')||A(t,'hAnsi')||A(t,'eastAsia'),c=A(t,'cs');if(a)o.fa=a;if(c)o.fc=c}
  if(t=K(e,'highlight')){const v=A(t,'val');if(v&&v!=='none')o.hl=HL[v]||null}
  if(t=K(e,'shd')){const c=colr(A(t,'fill'));if(c)o.bg=c}
  if(t=K(e,'caps'))o.caps=on(t);if(t=K(e,'smallCaps'))o.sc=on(t);
  return o};
 const bdr=e=>{const o={};if(!e)return o;for(const s of ['top','bottom','left','right','start','end','insideH','insideV']){const b=K(e,s);if(b)o[s]={v:A(b,'val'),sz:+A(b,'sz')||4,c:A(b,'color')}}return o};
 const bcss=b=>{if(!b||!b.v||/^(nil|none)$/i.test(b.v))return null;const px=Math.max(1,Math.round(b.sz/8*1.333)),sty=/double/i.test(b.v)?'double':/dash/i.test(b.v)?'dashed':/dot/i.test(b.v)?'dotted':'solid';return px+'px '+sty+' '+(colr(b.c)||'#000')};
 const pPr=e=>{const o={};if(!e)return o;let t;
  if(t=K(e,'jc'))o.jc=A(t,'val');if(t=K(e,'bidi'))o.bidi=on(t);
  if(t=K(e,'spacing')){const b=A(t,'before'),a=A(t,'after'),l=A(t,'line');if(b!=null&&b!=='')o.sb=+b;if(a!=null&&a!=='')o.sa=+a;if(l!=null&&l!==''){o.ln=+l;o.lr=A(t,'lineRule')||'auto'}if(A(t,'beforeAutospacing')==='1')o.sb=280;if(A(t,'afterAutospacing')==='1')o.sa=280}
  if(t=K(e,'ind')){const l=A(t,'left')??A(t,'start'),r=A(t,'right')??A(t,'end'),f=A(t,'firstLine'),h=A(t,'hanging');if(l!=null)o.il=+l;if(r!=null)o.ir=+r;if(f!=null)o.fl=+f;if(h!=null)o.hg=+h}
  if(t=K(e,'numPr')){const n=K(t,'numId'),l=K(t,'ilvl');if(n)o.numId=A(n,'val');if(l)o.ilvl=+A(l,'val')}
  if(t=K(e,'shd')){const c=colr(A(t,'fill'));if(c)o.shd=c}
  if(t=K(e,'pBdr'))o.bd=bdr(t);
  if(K(e,'pageBreakBefore'))o.pbb=on(K(e,'pageBreakBefore'));
  return o};
 // ----- styles
 const sx=await rd('word/styles.xml'),ST={},DD={p:{},r:{}};let defP=null;
 if(sx){const sd=PX(sx).documentElement,dd=K(sd,'docDefaults');
  if(dd){DD.r=rPr(K(K(dd,'rPrDefault'),'rPr'));DD.p=pPr(K(K(dd,'pPrDefault'),'pPr'))}
  for(const s of KS(sd,'style')){const id=A(s,'styleId');ST[id]={t:A(s,'type'),base:A(K(s,'basedOn'),'val'),name:A(K(s,'name'),'val')||'',p:pPr(K(s,'pPr')),r:rPr(K(s,'rPr')),tp:K(s,'tblPr')};if(A(s,'default')==='1'&&ST[id].t==='paragraph')defP=id}}
 const eff=(id,k)=>{const ch=[];for(let i=id,n=0;i&&ST[i]&&n<25;i=ST[i].base,n++)ch.unshift(ST[i]);return Object.assign({},...ch.map(s=>s[k]))};
 // ----- numbering
 const nx=await rd('word/numbering.xml'),NUM={},ABS={};
 if(nx){const nd=PX(nx).documentElement;
  for(const a of KS(nd,'abstractNum')){const L={};for(const l of KS(a,'lvl'))L[+A(l,'ilvl')]={fmt:A(K(l,'numFmt'),'val')||'decimal',st:+A(K(l,'start'),'val')||1,txt:A(K(l,'lvlText'),'val')||''};ABS[A(a,'abstractNumId')]=L}
  for(const n of KS(nd,'num'))NUM[A(n,'numId')]=A(K(n,'abstractNumId'),'val')}
 const LST={decimal:'decimal',lowerLetter:'lower-alpha',upperLetter:'upper-alpha',lowerRoman:'lower-roman',upperRoman:'upper-roman',hindiNumbers:'arabic-indic',decimalZero:'decimal-leading-zero'};
 const lvlOf=(numId,l)=>{const a=ABS[NUM[numId]];return a&&(a[l]||a[0])||{fmt:'decimal',st:1,txt:''}};
 // ----- relationships / images
 const relx=await rd('word/_rels/document.xml.rels'),REL={};
 if(relx)for(const r of [...PX(relx).documentElement.children])REL[r.getAttribute('Id')]={t:r.getAttribute('Target')||'',m:r.getAttribute('TargetMode')};
 const MIME={png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',bmp:'image/bmp',svg:'image/svg+xml',webp:'image/webp'},IC={};
 const b64=u=>{let s='';for(let i=0;i<u.length;i+=0x8000)s+=String.fromCharCode.apply(null,u.subarray(i,i+0x8000));return btoa(s)};
 const imgUri=async id=>{const r=REL[id];if(!r||r.m==='External')return null;let p=r.t.replace(/^\//,'');if(!p.startsWith('word/'))p='word/'+p;const q=[];for(const s of p.split('/')){if(s==='..')q.pop();else if(s!=='.')q.push(s)}p=q.join('/');
  if(IC[p]!==undefined)return IC[p];const m=MIME[p.split('.').pop().toLowerCase()];let out=null;if(m){const u=await rdb(p);if(u)out='data:'+m+';base64,'+b64(u)}return IC[p]=out};
 // ----- page content width (twips) for table sizing
 const d=PX(dxml),body=K(d.documentElement,'body'),sps=[...d.getElementsByTagNameNS(NS,'sectPr')],sp=sps[sps.length-1];
 let ctw=9026;if(sp){const sz=K(sp,'pgSz'),mg=K(sp,'pgMar');if(sz&&mg){const w=+A(sz,'w')||11906;ctw=w-(+A(mg,'left')||1440)-(+A(mg,'right')||1440)}}
 // ----- style maps (only differences become inline spans)
 const SM=(r,ar)=>{const m={},fam=ar?(r.fc||r.fa):(r.fa||r.fc),sz=ar?(r.szc||r.sz):r.sz,b=ar&&r.bc!==undefined?r.bc:r.b,i=ar&&r.ic!==undefined?r.ic:r.i;
  if(fam)m['font-family']="'"+String(fam).replace(/['"]/g,'')+"'";if(sz)m['font-size']=sz+'pt';if(r.c)m.color=r.c;
  m['font-weight']=b?'bold':'normal';m['font-style']=i?'italic':'normal';
  m['text-decoration']=[r.u?'underline':'',r.s?'line-through':''].filter(Boolean).join(' ')||'none';
  const bg=r.hl||r.bg;if(bg)m['background-color']=bg;if(r.caps)m['text-transform']='uppercase';if(r.sc)m['font-variant']='small-caps';return m};
 const css=m=>Object.keys(m).map(k=>k+':'+m[k]).join(';');
 const arN=t=>{let a=0,l=0;for(const ch of t){if(AR.test(ch))a++;else if(/[A-Za-z]/.test(ch))l++}return [a,l]};
 // ----- paragraph content
 async function inline(node,ctx,segs){
  const cur=()=>segs[segs.length-1];
  for(const c of node.children){const n=c.localName;
   if(n==='r')await run(c,ctx,segs);
   else if(n==='hyperlink'){const id=c.getAttributeNS(RN,'id')||c.getAttribute('r:id'),rl=id&&REL[id],href=rl&&rl.m==='External'?rl.t:null;
    if(href&&/^(https?:|mailto:|tel:)/i.test(href)){cur().push('<a href="'+esc(href)+'">');await inline(c,ctx,segs);cur().push('</a>')}else await inline(c,ctx,segs)}
   else if(/^(ins|smartTag|fldSimple|customXml|moveTo)$/.test(n))await inline(c,ctx,segs);
   else if(n==='sdt'){const sc=K(c,'sdtContent');if(sc)await inline(sc,ctx,segs)}
   else if(n==='AlternateContent'){const ch=K(c,'Choice');if(ch)await inline(ch,ctx,segs)}}}
 async function run(c,ctx,segs){
  const rp0=K(c,'rPr'),rs=A(K(rp0,'rStyle'),'val'),rp=Object.assign({},ctx.baseRun,eff(rs,'r'),rPr(rp0)),cur=()=>segs[segs.length-1];
  let txt='',html='';
  const flush=()=>{if(!txt&&!html)return;const m=SM(rp,!!(txt&&arN(txt)[0]>0&&arN(txt)[0]>=arN(txt)[1])),diff={};for(const k in m)if(m[k]!==ctx.baseMap[k])diff[k]=m[k];
   const st=css(diff);cur().push(st?'<span style="'+esc(st)+'">'+html+'</span>':html);txt='';html=''};
  const add=t=>{txt+=t;html+=esc(t)};
  for(const e of c.children){const n=e.localName;
   if(n==='t')add(e.textContent);
   else if(n==='tab')add('\u00a0\u00a0\u00a0\u00a0');
   else if(n==='noBreakHyphen')add('-');
   else if(n==='br'||n==='cr'){if(A(e,'type')==='page'){flush();segs.push([])}else html+='<br>'}
   else if(n==='drawing'||n==='pict'||n==='object'){flush();
    let id=null,wpx=0;const bl=e.getElementsByTagNameNS(DML,'blip')[0];
    if(bl){id=bl.getAttributeNS(RN,'embed')||bl.getAttribute('r:embed');const ex=e.getElementsByTagNameNS(WPD,'extent')[0];if(ex)wpx=Math.round((+ex.getAttribute('cx')||0)/9525)}
    else{const im=e.getElementsByTagName('v:imagedata')[0];if(im){id=im.getAttribute('r:id');const sh=e.getElementsByTagName('v:shape')[0],mm=sh&&/width:([\d.]+)pt/.exec(sh.getAttribute('style')||'');if(mm)wpx=Math.round(+mm[1]*1.333)}}
    const u=id&&await imgUri(id);if(u)cur().push('<img src="'+u+'" style="'+(wpx?'width:'+wpx+'px;':'')+'height:auto;max-width:100%">')}}
  flush()}
 let TS=null;
 async function para(p){
  const ppr=K(p,'pPr'),sid=A(K(ppr,'pStyle'),'val')||defP,tsp=TS&&sid===defP?TS.p:null,tsr=TS&&sid===defP?TS.r:null,pr=Object.assign({},DD.p,eff(sid,'p'),tsp,pPr(ppr)),baseRun=Object.assign({},DD.r,eff(sid,'r'),tsr);
  const [a,l]=arN(p.textContent||''),ar=!!pr.bidi||a>l,baseMap=SM(baseRun,ar),ctx={baseRun,baseMap,ar};
  const segs=[[]];await inline(p,ctx,segs);
  const nm=ST[sid]&&ST[sid].name,hm=/^heading\s*(\d)$/i.exec(nm||'')||/^Heading(\d)$/.exec(sid||''),tag=hm?'h'+Math.min(6,+hm[1]):'p';
  const list=pr.numId&&pr.numId!=='0'&&NUM[pr.numId]?{numId:pr.numId,lvl:pr.ilvl||0,lv:lvlOf(pr.numId,pr.ilvl||0)}:null;
  const m=Object.assign({},baseMap);if(tag!=='p'&&!m['font-size'])m['font-size']=(DD.r.sz||11)+'pt';
  const st=['margin-top:'+((pr.sb||0)/20)+'pt','margin-bottom:'+((pr.sa||0)/20)+'pt','white-space:pre-wrap'];
  st.push(pr.ln?(pr.lr==='auto'?'line-height:'+(pr.ln/240).toFixed(3):'line-height:'+(pr.ln/20)+'pt'):'line-height:1.15');
  if(!list){if(pr.il)st.push('margin-inline-start:'+pr.il/20+'pt');if(pr.ir)st.push('margin-inline-end:'+pr.ir/20+'pt');const ti=(pr.fl||0)-(pr.hg||0);if(ti)st.push('text-indent:'+ti/20+'pt')}
  const jc=pr.jc;if(jc){const rt=!!pr.bidi;const v=jc==='center'?'center':jc==='both'||jc==='distribute'?'justify':(jc==='left'||jc==='start')?(rt?'right':'left'):(rt?'left':'right');st.push('text-align:'+v)}
  if(pr.shd)st.push('background-color:'+pr.shd);
  if(pr.bd)for(const [k,s] of [['top','top'],['bottom','bottom'],['left',pr.bidi?'end':'start'],['right',pr.bidi?'start':'end']]){const v=bcss(pr.bd[k]||pr.bd[s]);if(v)st.push('border-'+k+':'+v+';padding-'+k+':2px')}
  const attrs=' dir="'+(pr.bidi?'rtl':'ltr')+'"'+(pr.pbb?' data-pb="1"':'')+' style="'+esc(css(m)+';'+st.join(';'))+'"';
  const BR='<div class="wpbr" data-page-break="1" contenteditable="false"><span>فاصل صفحة</span></div>';
  const empty=s=>!s.join('').replace(/<[^>]*>/g,'').length&&!/<(img|br)/.test(s.join(''));
  if(list)return {list,dir:pr.bidi?'rtl':'ltr',attrs,inner:segs.map(s=>s.join('')).join('<br>')||'<br>'};
  let out='';segs.forEach((s,i)=>{if(i>0)out+=BR;const inner=s.join('');if(segs.length>1&&empty(s)&&(i===0||i===segs.length-1)){return}out+='<'+tag+attrs+'>'+(inner||'<br>')+'</'+tag+'>'});
  return {html:out}}
 // ----- tables
 async function table(t){
  const tp=K(t,'tblPr'),sid=A(K(tp,'tblStyle'),'val'),prevTS=TS;TS=sid&&ST[sid]?{p:eff(sid,'p'),r:eff(sid,'r')}:null;
  const tbl_=await table_(t,tp,sid);TS=prevTS;return tbl_}
 async function table_(t,tp,sid){
  const rtl=!!on(K(tp,'bidiVisual')),jc=A(K(tp,'jc'),'val');
  let tb=bdr(K(tp,'tblBorders'));for(let i=sid,n=0;i&&ST[i]&&n<10;i=ST[i].base,n++)tb=Object.assign({},bdr(K(ST[i].tp,'tblBorders')),tb);
  const cm=K(tp,'tblCellMar'),mL=cm&&K(cm,'left')||cm&&K(cm,'start'),mR=cm&&K(cm,'right')||cm&&K(cm,'end'),padL=mL?(+A(mL,'w')||0)/15:7,padR=mR?(+A(mR,'w')||0)/15:7;
  const tw=K(tp,'tblW'),tt=A(tw,'type'),tv=+A(tw,'w')||0,grid=KS(K(t,'tblGrid'),'gridCol').map(g=>+A(g,'w')||0),gs=grid.reduce((a,b)=>a+b,0);
  const rows=KS(t,'tr').map(tr=>{let col=0;const cells=[];for(const tc of KS(tr,'tc')){const pr=K(tc,'tcPr'),span=+A(K(pr,'gridSpan'),'val')||1,vm=K(pr,'vMerge');cells.push({tc,pr,col,span,vm:vm?(A(vm,'val')||'continue'):null});col+=span}return {tr,cells}});
  const nR=rows.length,nC=Math.max(1,grid.length,...rows.map(r=>r.cells.reduce((a,c)=>a+c.span,0)));
  const full=(tt==='pct'&&tv>=4900)||(tt==='dxa'&&tv>=ctw-100)||(tt!=='pct'&&tt!=='dxa'&&gs>=ctw-100);
  const wst=tt==='pct'?(tv/50)+'%':tt==='dxa'&&tv?Math.round(tv/15)+'px':gs?Math.round(gs/15)+'px':'100%';
  let h='<table dir="'+(rtl?'rtl':'ltr')+'"'+(full?'':' data-tw="auto"')+(jc==='center'?' data-al="c"':jc==='right'||(rtl&&!jc)?' data-al="r"':'')+' style="border-collapse:collapse;table-layout:fixed;max-width:100%;width:'+(full?'100%':wst)+(jc==='center'?';margin-left:auto;margin-right:auto':(jc==='right'&&!rtl)||(rtl&&!jc)?';margin-left:auto;margin-right:0':';margin-left:0;margin-right:auto')+'">';
  const edge=(side,r,c,rs,cs)=>{const first=c===0,last=c+cs>=nC;let b;
   if(side==='top')b=r===0?tb.top:tb.insideH;else if(side==='bottom')b=r+rs>=nR?tb.bottom:tb.insideH;
   else{const st=tb.start||(rtl?tb.right:tb.left),en=tb.end||(rtl?tb.left:tb.right),isStart=(side==='left')!==rtl;b=isStart?(first?st:tb.insideV):(last?en:tb.insideV)}return b};
  for(let r=0;r<nR;r++){const {tr,cells}=rows[r],trp=K(tr,'trPr'),th=A(K(trp,'trHeight'),'val');
   h+='<tr'+(K(trp,'tblHeader')?' data-hdr="1"':'')+(th?' style="height:'+Math.round(+th/15)+'px"':'')+'>';
   for(const cl of cells){if(cl.vm==='continue')continue;
    let rs=1;if(cl.vm==='restart'){for(let r2=r+1;r2<nR;r2++){const k=rows[r2].cells.find(x=>x.col===cl.col&&x.vm==='continue');if(k)rs++;else break}}
    const w=grid.length?grid.slice(cl.col,cl.col+cl.span).reduce((a,b)=>a+b,0)/15:(+A(K(cl.pr,'tcW'),'w')||0)/15;
    const own=bdr(K(cl.pr,'tcBorders')),st=['border:none'];
    for(const [side,dirk] of [['top','top'],['bottom','bottom'],['left',rtl?'end':'start'],['right',rtl?'start':'end']]){const ob=own[side]||own[dirk]||(side==='left'?own.left:side==='right'?own.right:null),b=ob||edge(side,r,cl.col,rs,cl.span),v=bcss(b);if(v)st.push('border-'+side+':'+v)}
    if(w)st.push('width:'+Math.round(w)+'px');
    const sh=K(cl.pr,'shd'),bg=sh&&colr(A(sh,'fill'));if(bg)st.push('background-color:'+bg);
    const va=A(K(cl.pr,'vAlign'),'val');st.push('vertical-align:'+(va==='center'?'middle':va==='bottom'?'bottom':'top'));
    st.push('padding:2px '+Math.round(padR)+'px 2px '+Math.round(padL)+'px');
    h+='<td'+(cl.span>1?' colspan="'+cl.span+'"':'')+(rs>1?' rowspan="'+rs+'"':'')+' style="'+st.join(';')+'">'+(await blocks(cl.tc)||'<p><br></p>')+'</td>'}
   h+='</tr>'}
  return h+'</table>'}
 // ----- block flow (paragraphs, lists, tables)
 async function blocks(parent){
  let h='';const st=[],LC={};
  const closeL=()=>{while(st.length)h+='</li></'+st.pop().tag+'>'};
  for(const c of parent.children){const n=c.localName;
   if(n==='p'){const P=await para(c);
    if(!P.list){closeL();h+=P.html;continue}
    const L=P.list,tag=L.lv.fmt==='bullet'?'ul':'ol',lv=Math.min(L.lvl,8);
    const open=()=>{const key=L.numId+':'+lv,cnt=LC[key]||0;let sty;if(tag==='ul'){const tx=L.lv.txt;sty=tx==='o'?'circle':/[§▪■]/.test(tx)?'square':'disc'}else sty=LST[L.lv.fmt]||'decimal';
     const start=tag==='ol'?(cnt>0?cnt+1:L.lv.st):0;h+='<'+tag+' dir="'+P.dir+'"'+(tag==='ol'&&start!==1?' start="'+start+'"':'')+' style="list-style-type:'+sty+'">';st.push({tag,num:L.numId})};
    while(st.length>lv+1)h+='</li></'+st.pop().tag+'>';
    if(st.length===lv+1){const top=st[st.length-1];h+='</li>';if(top.tag!==tag||top.num!==L.numId){h+='</'+top.tag+'>';st.pop();open()}}
    else while(st.length<lv+1){open();if(st.length<lv+1)h+='<li style="list-style:none">'}
    for(const k of Object.keys(LC))if(k.startsWith(L.numId+':')&&+k.split(':')[1]>lv)delete LC[k];
    LC[L.numId+':'+lv]=(LC[L.numId+':'+lv]||0)+1;
    h+='<li'+P.attrs+'>'+P.inner}
   else if(n==='tbl'){closeL();h+=await table(c)}
   else if(n==='sdt'){const sc=K(c,'sdtContent');if(sc){closeL();h+=await blocks(sc)}}}
  closeL();return h}
 const html=await blocks(body);return html||'<p><br></p>'}

window.WordPro={mount,section,measure,cellBorders,tableRtl,tableAlign,contentPx,printBox,readDocx,importDocx};
})();
