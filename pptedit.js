// Reader All v28 — PowerPoint editor: mobile image styling, expanded templates, text formatting, slides and slideshow.
// Loaded after office.js (uses window.__of). Saves a NEW .pptx copy; the original is never touched.
(()=>{'use strict';
const X=window.__of,$=s=>document.querySelector(s),E=(t,c,h)=>{const e=document.createElement(t);if(c)e.className=c;if(h!=null)e.innerHTML=h;return e};
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',PR='http://schemas.openxmlformats.org/package/2006/relationships',CT='http://schemas.openxmlformats.org/package/2006/content-types';
const NSD=`xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"`,td=new TextDecoder(),te=new TextEncoder(),PT=12700;
const ps=s=>new DOMParser().parseFromString(s,'application/xml');
const ser=d=>new XMLSerializer().serializeToString(d).replace(/^<\?xml[^>]*\?>\s*/,'');
const xml=d=>te.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+ser(d));
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const kids=(n,ns,l)=>n?[...n.childNodes].filter(c=>c.nodeType==1&&c.localName==l&&(!ns||c.namespaceURI==ns)):[];
const k1=(n,ns,l)=>kids(n,ns,l)[0]||null;
const all=(n,ns,l)=>[...n.getElementsByTagNameNS(ns,l)];
const norm=(b,t)=>{if(t[0]=='/')return t.slice(1);const p=b.split('/');p.pop();for(const s of t.split('/')){if(s=='..')p.pop();else if(s&&s!='.')p.push(s)}return p.join('/')};
const relPath=p=>{const i=p.lastIndexOf('/');return p.slice(0,i)+'/_rels/'+p.slice(i+1)+'.rels'};
const nextRid=d=>'rId'+(Math.max(0,...all(d,PR,'Relationship').map(r=>+r.getAttribute('Id').replace(/\D/g,'')||0))+1);
const frag=(s,doc)=>doc.importNode(ps(`<w ${NSD}>${s}</w>`).documentElement.firstElementChild,true);
let S=null;
const cur=()=>S.slides[S.cur],tree=s=>all(s.doc,P,'spTree')[0];
const touch=()=>{cur().t=1;X.O().dirty=1};

// ---------- load ----------
async function load(){const O=X.O(),z=X.zipIndex(O.buf),rd=async n=>{const d=await X.zipRead(z,n);return d?td.decode(d):null};
 const s={z,rd,put:new Map(),del:new Set(),gone:[],blobs:{},docs:{},exts:new Set(),slides:[],cur:0,sel:null,ed:null,k:1,st:0};
 s.pres=ps(await rd('ppt/presentation.xml'));
 try{const tn=Object.keys(z.m).filter(n=>/^ppt\/theme\/theme\d+\.xml$/.test(n)).sort()[0],tt=tn&&await rd(tn);
  if(tt){const th=['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6'].map(k=>{const m=new RegExp('<(?:a:)?'+k+'>([\\s\\S]*?)</(?:a:)?'+k+'>').exec(tt),v=m&&(/srgbClr[^>]*val="([0-9a-fA-F]{6})"/.exec(m[1])||/lastClr="([0-9a-fA-F]{6})"/.exec(m[1]));return v?'#'+v[1]:null});if(th.every(Boolean))s.theme=th}}catch(e){console.warn('theme',e)}s.prel=ps(await rd('ppt/_rels/presentation.xml.rels'));
 const sz=all(s.pres,P,'sldSz')[0];s.W=+sz.getAttribute('cx');s.H=+sz.getAttribute('cy');
 const rm={};for(const r of all(s.prel,PR,'Relationship'))rm[r.getAttribute('Id')]=r;
 for(const sid of all(s.pres,P,'sldId')){const rid=sid.getAttributeNS(R,'id'),r=rm[rid];if(!r||r.getAttribute('TargetMode')=='External')continue;
  const path=norm('ppt/presentation.xml',r.getAttribute('Target')),t=await rd(path);if(!t)continue;
  s.slides.push({path,rid,sid,doc:ps(t),rel:ps(await rd(relPath(path))||`<Relationships xmlns="${PR}"/>`)})}
 if(!s.slides.length)throw new Error('مفيش شرايح في الملف');
 for(const sl of s.slides){const lr=all(sl.rel,PR,'Relationship').find(r=>/slideLayout$/.test(r.getAttribute('Type')));if(!lr)continue;
  sl.lt=lr.getAttribute('Target');const lp=norm(sl.path,sl.lt),get=async p=>s.docs[p]||(s.docs[p]=ps(await rd(p)||'<x/>'));
  sl.lay=await get(lp);sl.lp=lp;sl.lrel=ps(await rd(relPath(lp))||`<Relationships xmlns="${PR}"/>`);const mr=all(sl.lrel,PR,'Relationship').find(r=>/slideMaster$/.test(r.getAttribute('Type')));
  if(mr){sl.mp=norm(lp,mr.getAttribute('Target'));sl.mas=await get(sl.mp);sl.mrel=ps(await rd(relPath(sl.mp))||`<Relationships xmlns="${PR}"/>`)}}
 s.orig=[...s.slides];s.hist=[];s.redo=[];s.lc='#000000';s.lw=2;return s}

// ---------- geometry / colours ----------
const xf=(e,mk)=>{const g=e.localName=='graphicFrame';let x=g?k1(e,P,'xfrm'):k1(k1(e,P,'spPr'),A,'xfrm');
 if(!x&&mk&&!g){const sp=k1(e,P,'spPr');x=frag('<a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm>',e.ownerDocument);sp.insertBefore(x,sp.firstChild)}return x};
function geo(e,s){const x=xf(e);let o=x&&k1(x,A,'off'),c=x&&k1(x,A,'ext');
 if(!o||!c){const ph=all(e,P,'ph')[0];if(ph)for(const d of [s.lay,s.mas]){if(!d)continue;const ix=ph.getAttribute('idx'),ty=ph.getAttribute('type');
  const m=all(d,P,'sp').find(q=>{const h=all(q,P,'ph')[0];return h&&(ix?h.getAttribute('idx')==ix:h.getAttribute('type')==ty)}),mx=m&&xf(m);
  if(mx&&k1(mx,A,'off')){o=k1(mx,A,'off');c=k1(mx,A,'ext');break}}}
 return o&&c?{x:+o.getAttribute('x'),y:+o.getAttribute('y'),w:+c.getAttribute('cx'),h:+c.getAttribute('cy')}:null}
function setGeo(e,g){const x=xf(e,1),o=k1(x,A,'off'),c=k1(x,A,'ext');o.setAttribute('x',Math.round(g.x));o.setAttribute('y',Math.round(g.y));c.setAttribute('cx',Math.round(Math.max(g.w,PT*12)));c.setAttribute('cy',Math.round(Math.max(g.h,PT*12)))}
const clr=n=>{const f=k1(n,A,'solidFill'),c=f&&f.firstElementChild;if(!c)return null;const v=c.getAttribute('val');
 if(c.localName=='srgbClr'){const al=c.firstElementChild;if(al&&al.localName=='alpha'){const n=parseInt(v,16);return `rgba(${n>>16},${(n>>8)&255},${n&255},${(+al.getAttribute('val')||1e5)/1e5})`}return '#'+v}if(c.localName=='schemeClr')return /^(bg1|lt1)$/.test(v)?'#ffffff':/^(tx1|dk1)$/.test(v)?'#111111':null;return null};
const lum=h=>{if(h[0]!='#')return .5;const n=parseInt(h.slice(1),16);return(.299*(n>>16)+.587*((n>>8)&255)+.114*(n&255))/255};
// ---------- التدرجات اللونية (gradFill) ----------
function gcol(gs){const c=gs.firstElementChild;if(!c)return null;const d=gs.ownerDocument,w=d.createElementNS(A,'a:x'),t=d.createElementNS(A,'a:solidFill');t.append(c.cloneNode(true));w.append(t);return clr(w)}
function gradCss(pr){const gf=pr&&k1(pr,A,'gradFill');if(!gf)return null;const gl=k1(gf,A,'gsLst'),gs=gl?kids(gl,A,'gs'):[];
 const st=gs.map(g=>[gcol(g),(+g.getAttribute('pos')||0)/1000]).filter(x=>x[0]).sort((a,b)=>a[1]-b[1]);if(st.length<2)return null;
 const sx=st.map(x=>x[0]+' '+x[1]+'%').join(','),pa=k1(gf,A,'path'),ln=k1(gf,A,'lin');
 if(pa)return 'radial-gradient(circle at center,'+sx+')';
 return 'linear-gradient('+(((ln?(+ln.getAttribute('ang')||0):5400000)/60000)+90)+'deg,'+sx+')'}
function gradFirst(pr){const gf=pr&&k1(pr,A,'gradFill'),gl=gf&&k1(gf,A,'gsLst'),g=gl&&kids(gl,A,'gs')[0];return g?gcol(g):null}
function bgGrad(s){for(const d of [s.doc,s.lay,s.mas]){const b=d&&all(d,P,'bg')[0],pr=b&&k1(b,P,'bgPr');if(pr){if(clr(pr))return null;const g=gradCss(pr);if(g)return g}}return null}
const bgc=s=>{for(const d of [s.doc,s.lay,s.mas]){const b=d&&all(d,P,'bg')[0],pr=b&&k1(b,P,'bgPr'),c=pr&&(clr(pr)||gradFirst(pr));if(c)return c}return '#ffffff'};
const hex=c=>{const m=/rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);return m?[m[1],m[2],m[3]].map(v=>(+v).toString(16).padStart(2,'0')).join(''):/^#[0-9a-f]{6}$/i.test(c)?c.slice(1):null};

// ---------- painting ----------
async function imgUrl(s,rid){const r=all(s.rel,PR,'Relationship').find(x=>x.getAttribute('Id')==rid);if(!r)return null;const p=norm(s.path,r.getAttribute('Target'));
 if(S.blobs[p])return S.blobs[p];const d=S.put.get(p)||await X.zipRead(S.z,p);if(!d)return null;
 const m={png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',svg:'image/svg+xml',webp:'image/webp',bmp:'image/bmp'}[p.split('.').pop().toLowerCase()]||'image/png';
 return S.blobs[p]=URL.createObjectURL(new Blob([d],{type:m}))}
const dc=(df,dark)=>{const c=df.c||'#111111';return dark&&c=='#111111'?'#ffffff':c};
function dflt(e,s){const ph=all(e,P,'ph')[0],ty=(ph&&ph.getAttribute('type'))||'',ix=ph&&ph.getAttribute('idx'),L=[k1(k1(e,P,'txBody'),A,'lstStyle')];
 if(ph)for(const d of [s.lay,s.mas]){const m=d&&all(d,P,'sp').find(q=>{const h=all(q,P,'ph')[0];return h&&(ix?h.getAttribute('idx')==ix:h.getAttribute('type')==ty)});m&&L.push(k1(k1(m,P,'txBody'),A,'lstStyle'))}
 const ts=s.mas&&all(s.mas,P,'txStyles')[0];ts&&L.push(k1(ts,P,/title/i.test(ty)?'titleStyle':ph?'bodyStyle':'otherStyle'));
 const r={sz:null,b:null,i:null,c:null};
 for(const l of L){const d=l&&k1(k1(l,A,'lvl1pPr'),A,'defRPr');if(!d)continue;if(r.sz==null&&d.getAttribute('sz'))r.sz=d.getAttribute('sz')/100;
  if(r.b==null&&d.getAttribute('b'))r.b=d.getAttribute('b')=='1';if(r.i==null&&d.getAttribute('i'))r.i=d.getAttribute('i')=='1';if(!r.c)r.c=clr(d)}
 r.sz=r.sz||(/title/i.test(ty)?32:ty=='subTitle'?24:18);return r}
function drawText(e,d,k,dark,s){const tb=k1(e,P,'txBody');if(!tb)return;const t=E('div','pe-t'),bp=k1(tb,A,'bodyPr'),an=bp&&bp.getAttribute('anchor'),df=dflt(e,s);
 const ins=a=>bp&&bp.getAttribute(a)!=null?+bp.getAttribute(a):({lIns:91440,rIns:91440,tIns:45720,bIns:45720}[a]||0);
 Object.assign(t.style,{display:'flex',flexDirection:'column',justifyContent:an=='ctr'?'center':an=='b'?'flex-end':'flex-start',padding:`${ins('tIns')*k}px ${ins('rIns')*k}px ${ins('bIns')*k}px ${ins('lIns')*k}px`,color:dc(df,dark),fontSize:df.sz*PT*k+'px',fontWeight:df.b?'bold':'normal',fontStyle:df.i?'italic':'normal'});
 for(const p of kids(tb,A,'p')){const q=E('div'),pp=k1(p,A,'pPr'),al=pp&&pp.getAttribute('algn'),ls=pp&&k1(k1(pp,A,'lnSpc'),A,'spcPct');
  if(al)q.style.textAlign={l:'left',r:'right',ctr:'center',just:'justify'}[al]||'';q.dir=pp&&pp.getAttribute('rtl')=='1'?'rtl':'auto';q.style.whiteSpace='pre-wrap';
  if(ls)q.style.lineHeight=ls.getAttribute('val')/100000;if(pp&&k1(pp,A,'buChar'))q.dataset.bu='•';else if(pp&&k1(pp,A,'buAutoNum'))q.dataset.bu='1';
  let n=0;for(const c of p.children){if(c.localName=='br'){q.append(E('br'));continue}if(c.localName!='r'&&c.localName!='fld')continue;
   const rp=k1(c,A,'rPr'),g=a=>rp&&rp.getAttribute(a),sp=E('span'),lt=rp&&k1(rp,A,'latin'),tf=lt&&lt.getAttribute('typeface');sp.textContent=(k1(c,A,'t')||{}).textContent||'';n++;
   sp.style.fontSize=((+g('sz')/100||df.sz)*PT*k)+'px';if(g('b')!=null?g('b')=='1':df.b)sp.style.fontWeight='bold';if(g('i')!=null?g('i')=='1':df.i)sp.style.fontStyle='italic';
   if(g('u')&&g('u')!='none')sp.style.textDecoration='underline';if(g('sp'))sp.style.letterSpacing=(+g('sp')/100)*PT*k+'px';if(rp&&k1(rp,A,'effectLst')&&k1(k1(rp,A,'effectLst'),A,'outerShdw'))sp.style.textShadow='1px 1px 2px #555';sp.style.color=clr(rp)||dc(df,dark);if(tf&&tf[0]!='+')sp.style.fontFamily=`"${tf}",sans-serif`;q.append(sp)}
  if(!n)q.append(E('br'));t.append(q)}
 d.append(t)}
function table(tb,k,dark){const T=E('table','pe-tb'),cg=document.createElement('colgroup');
 for(const c of all(tb,A,'gridCol')){const x=document.createElement('col');x.style.width=c.getAttribute('w')*k+'px';cg.append(x)}T.append(cg);
 for(const tr of kids(tb,A,'tr')){const r=T.insertRow();r.style.height=tr.getAttribute('h')*k+'px';
  for(const tc of kids(tr,A,'tc')){if(tc.getAttribute('hMerge')=='1')continue;const c=r.insertCell(),rp=all(tc,A,'rPr')[0],f=clr(k1(tc,A,'tcPr'));
   c.style.fontSize=((rp&&+rp.getAttribute('sz')/100)||14)*PT*k+'px';c.style.color=dark?'#fff':'#111';c.style.whiteSpace='pre-wrap';c.dir='auto';if(tc.getAttribute('gridSpan'))c.colSpan=+tc.getAttribute('gridSpan');
   if(f)c.style.background=f;if(rp&&rp.getAttribute('b')=='1')c.style.fontWeight='bold';
   c.textContent=all(tc,A,'p').map(p=>all(p,A,'t').map(t=>t.textContent).join('')).join('\n');c._tc=tc;c._o=c.textContent}}
 return T}
// audio/video element -> its media item (so a play button can be put on it while editing)
function mediaOf(e,s){if(!S.media||!S.media.length)return null;const nv=k1(e,P,'nvPicPr'),pv=nv&&k1(nv,P,'nvPr');if(!pv)return null;
 for(const el of pv.getElementsByTagName('*'))if(/^(audioFile|videoFile|wavAudioFile|quickTimeFile|media)$/.test(el.localName))for(const a of ['link','embed']){const id=el.getAttributeNS(R,a);if(!id)continue;
  const r=all(s.rel,PR,'Relationship').find(x=>x.getAttribute('Id')==id);if(!r||r.getAttribute('TargetMode')=='External')continue;const p=norm(s.path,r.getAttribute('Target')),m=S.media.find(x=>x.path==p);if(m)return m}
 return null}
// ---------- قص الصورة على أي شكل: clip-path من مكتبة الأشكال + خبز الشكل قبل تصدير PDF ----------
function picClip(pr){let sv=document.getElementById('pe-defs');
 if(!sv){sv=document.createElementNS('http://www.w3.org/2000/svg','svg');sv.id='pe-defs';sv.setAttribute('width','0');sv.setAttribute('height','0');sv.style.cssText='position:absolute;width:0;height:0;pointer-events:none';sv.innerHTML='<defs></defs>';document.body.append(sv)}
 const id='pecp-'+pr;if(!document.getElementById(id)){const q=SH[pr];sv.querySelector('defs').insertAdjacentHTML('beforeend','<clipPath id="'+id+'" clipPathUnits="objectBoundingBox"><path transform="scale(.01)" d="'+q[1]+'"'+(q[2]?' clip-rule="evenodd"':'')+'/></clipPath>')}return id}
async function bakeClips(st){for(const d of st.querySelectorAll('.pe-s[data-clip]')){try{const im=d.querySelector('img'),pr=d.dataset.clip,q=SH[pr];if(!im||!q)continue;if(im.decode)await im.decode().catch(()=>{});if(!im.naturalWidth)continue;
  const w=Math.max(1,Math.round(d.offsetWidth)),h=Math.max(1,Math.round(d.offsetHeight)),c=document.createElement('canvas');c.width=w;c.height=h;const x=c.getContext('2d');
  const sr=(d.dataset.sr||'0,0,0,0').split(',').map(Number),fl=(d.dataset.fl||'0,0').split(',').map(Number),nw=im.naturalWidth,nh=im.naturalHeight,
   p=new Path2D();p.addPath(new Path2D(q[1]),new DOMMatrix([w/100,0,0,h/100,0,0]));
  x.save();x.clip(p,q[2]?'evenodd':'nonzero');if(fl[0]||fl[1]){x.translate(fl[0]?w:0,fl[1]?h:0);x.scale(fl[0]?-1:1,fl[1]?-1:1)}x.globalAlpha=+(d.dataset.al||1);
  x.drawImage(im,sr[0]*nw,sr[1]*nh,nw*(1-sr[0]-sr[2]),nh*(1-sr[1]-sr[3]),0,0,w,h);x.restore();
  if(d.dataset.lc){x.strokeStyle=d.dataset.lc;x.lineWidth=+d.dataset.lw||1;x.lineJoin='round';x.stroke(p)}
  c.style.cssText='position:absolute;left:0;top:0;width:100%;height:100%';d.innerHTML='';d.append(c)}catch(_){}}}
// ---------- رسم حر: يعرض a:custGeom (قلم / فري هاند) كـ SVG ----------
function inkDraw(d,cg,l,k,f,g){const pa=all(cg,A,'path')[0];if(!pa)return false;const w=+pa.getAttribute('w'),h=+pa.getAttribute('h');if(!(w>0&&h>0))return false;
 const pt=n=>n?n.getAttribute('x')+' '+n.getAttribute('y'):'0 0';let dd='';
 for(const c of pa.children){const n=c.localName,q=kids(c,A,'pt');
  if(n=='moveTo')dd+='M'+pt(q[0]);else if(n=='lnTo')dd+='L'+pt(q[0]);else if(n=='cubicBezTo')dd+='C'+pt(q[0])+' '+pt(q[1])+' '+pt(q[2]);else if(n=='quadBezTo')dd+='Q'+pt(q[0])+' '+pt(q[1]);else if(n=='close')dd+='Z';else return false}
 if(!dd)return false;
 const gw=Math.max(1,g.w),u=w/gw,lc=l&&clr(l),lw=l?(+l.getAttribute('w')||12700):0,sf=l&&k1(l,A,'solidFill'),al=sf&&all(sf,A,'alpha')[0],op=al?(+al.getAttribute('val')||0)/1e5:1,
  fill=pa.getAttribute('fill')!=='none'?f:null,sw=lw*u,hit=Math.max(sw,16*w/(gw*k));
 d.style.background='none';d.style.overflow='visible';d.classList.add('pe-inkd');
 d.insertAdjacentHTML('afterbegin',`<svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" style="position:absolute;left:0;top:0;width:100%;height:100%;overflow:visible;pointer-events:none"><path d="${dd}" fill="${fill||'none'}" stroke="${lc||'none'}" stroke-width="${sw}" stroke-linecap="round" stroke-linejoin="round" opacity="${op}"/><path class="pe-ink-hit" d="${dd}" fill="none" stroke="transparent" stroke-width="${hit}" style="pointer-events:stroke"/></svg>`);return true}
function one(e,s,rs,k,dark,live,h){const n=e.localName;if(e.namespaceURI!=P||!/^(sp|pic|graphicFrame)$/.test(n))return;
 const c=all(e,P,'cNvPr')[0];if(c&&c.getAttribute('hidden')=='1')return;const g=geo(e,s);if(!g)return;
 const d=E('div','pe-s'),xr=xf(e);Object.assign(d.style,{left:g.x*k+'px',top:g.y*k+'px',width:g.w*k+'px',height:g.h*k+'px'});
 if(xr&&xr.getAttribute('rot'))d.style.transform=`rotate(${xr.getAttribute('rot')/60000}deg)`;
 if(n=='sp'){const sp=k1(e,P,'spPr'),f=clr(sp),gm=k1(sp,A,'prstGeom'),l=k1(sp,A,'ln'),lc=l&&clr(l);if(f)d.style.background=f;else{const gg=gradCss(sp);if(gg)d.style.background=gg}
  const pr=gm&&gm.getAttribute('prst'),shp=pr&&SH[pr]&&pr!='rect'&&pr!='ellipse'?pr:null,lw=l?Math.max(1,(+l.getAttribute('w')||12700)*k):0;
  if(k1(sp,A,'custGeom')&&inkDraw(d,k1(sp,A,'custGeom'),l,k,f,g)){}
  else if(pr&&CNR.test(pr)){connDraw(d,g,k,pr,lc,lw,xr,l)}
  else if(shp){d.style.background=shBg(shp,f,lc,lw)}
  else{if(pr=='ellipse')d.style.borderRadius='50%';if(lc)d.style.border=lw+'px solid '+lc}
  if(all(sp,A,'outerShdw')[0]){if(shp)d.style.filter='drop-shadow(3px 3px 3px #0006)';else d.style.boxShadow='3px 3px 6px #0006'}drawText(e,d,k,dark,s)}
 else if(n=='pic'){const im=E('img'),bl=all(e,A,'blip')[0];im.draggable=false;if(bl&&rs.rel)imgUrl(rs,bl.getAttributeNS(R,'embed')).then(u=>u&&(im.src=u));
  {const sr=all(e,A,'srcRect')[0],al=bl&&all(bl,A,'alphaModFix')[0],sp=k1(e,P,'spPr'),gm=sp&&k1(sp,A,'prstGeom'),pr=gm&&gm.getAttribute('prst'),ln=sp&&k1(sp,A,'ln'),lc=ln&&clr(ln),lw=ln?Math.max(1,(+ln.getAttribute('w')||12700)*k):0;let host=im;
   if(pr=='ellipse')d.style.borderRadius='50%';else if(pr=='roundRect')d.style.borderRadius='12%';const cl=pr&&SH[pr]&&!/^(rect|ellipse|roundRect)$/.test(pr)?pr:null;
   if(lc&&!cl){d.style.border=lw+'px '+(ln.getAttribute('cmpd')=='dbl'?'double':'solid')+' '+lc;d.style.boxSizing='border-box'}if(sp&&all(sp,A,'outerShdw').length)d.style.filter='drop-shadow(2px 3px 5px #0006)';
   if(al)im.style.opacity=(+al.getAttribute('amt')||0)/1e5;
   if(sr){const q=['l','t','r','b'].map(a=>(+sr.getAttribute(a)||0)/1e5),vw=Math.max(.01,1-q[0]-q[2]),vh=Math.max(.01,1-q[1]-q[3]);host=E('div');host.style.cssText='position:absolute;left:0;top:0;width:100%;height:100%;overflow:hidden';
    Object.assign(im.style,{position:'absolute',maxWidth:'none',width:100/vw+'%',height:100/vh+'%',left:-q[0]/vw*100+'%',top:-q[1]/vh*100+'%'});host.append(im)}
   if(xr){const fh=xr.getAttribute('flipH')=='1',fv=xr.getAttribute('flipV')=='1';(fh||fv)&&(host.style.transform='scale('+(fh?-1:1)+','+(fv?-1:1)+')')}
   if(cl){host.style.clipPath='url(#'+picClip(cl)+')';d.dataset.clip=cl;d.dataset.sr=sr?['l','t','r','b'].map(a=>(+sr.getAttribute(a)||0)/1e5).join(','):'0,0,0,0';d.dataset.fl=xr?(xr.getAttribute('flipH')=='1'?1:0)+','+(xr.getAttribute('flipV')=='1'?1:0):'0,0';d.dataset.al=al?(+al.getAttribute('amt')||0)/1e5:1;if(lc){d.dataset.lc=lc;d.dataset.lw=lw}}
   d.append(host);
   if(cl&&lc)d.insertAdjacentHTML('beforeend','<svg viewBox="0 0 100 100" preserveAspectRatio="none" style="position:absolute;left:0;top:0;width:100%;height:100%;overflow:visible;pointer-events:none"><path d="'+SH[cl][1]+'" fill="none" stroke="'+lc+'" stroke-width="'+lw+'" vector-effect="non-scaling-stroke" stroke-linejoin="round"/></svg>')}
  const mi=live&&mediaOf(e,rs);if(mi){const b=E('button','pe-mp','<i>'+(mi.kind=='video'?'▶':'🔊')+'</i>');b.setAttribute('aria-label','تشغيل: '+mi.title);b.addEventListener('pointerdown',ev=>ev.stopPropagation());b.onclick=ev=>{ev.stopPropagation();X.media.play(mi)};d.append(b)}}
 else{const tb=all(e,A,'tbl')[0];if(!tb)return;d.append(table(tb,k,dark))}
 if(live)d._e=e;h.append(d)}
function paint(s,h,k,live){h.innerHTML='';const b=bgc(s),dark=lum(b)<.45;Object.assign(h.style,{width:S.W*k+'px',height:S.H*k+'px',background:bgGrad(s)||b});
 const ms=s.doc.documentElement.getAttribute('showMasterSp')!='0',ls=!s.lay||s.lay.documentElement.getAttribute('showMasterSp')!='0';
 if(ms){if(ls&&s.mas)for(const e of tree({doc:s.mas}).children)if(!all(e,P,'ph')[0])one(e,s,{rel:s.mrel,path:s.mp},k,dark,0,h);
  if(s.lay)for(const e of tree({doc:s.lay}).children)if(!all(e,P,'ph')[0])one(e,s,{rel:s.lrel,path:s.lp},k,dark,0,h)}
 for(const e of tree(s).children)one(e,s,s,k,dark,live,h)}

// ---------- text write-back ----------
function runs(n,p,out){for(const c of n.childNodes){if(c.nodeType==3){if(c.data)out.push({...p,t:c.data});continue}if(c.nodeName=='BR')continue;
 const q={...p},s=c.style||{},w=s.fontWeight,td_=s.textDecoration||s.textDecorationLine||'',fc=c.nodeName=='FONT'&&c.getAttribute('color');
 if(/^(B|STRONG)$/.test(c.nodeName)||w=='bold'||+w>=600)q.b=1;else if(w=='normal'||w==='400')q.b=0;
 if(/^(I|EM)$/.test(c.nodeName)||s.fontStyle=='italic')q.i=1;else if(s.fontStyle=='normal')q.i=0;
 if(c.nodeName=='U'||/underline/.test(td_))q.u=1;else if(td_=='none')q.u=0;
 if(s.color)q.c=hex(s.color)||q.c;else if(fc)q.c=hex(fc)||q.c;if(s.fontFamily)q.f=s.fontFamily.split(',')[0].replace(/["']/g,'').trim();if(s.fontSize)q.s=parseFloat(s.fontSize);if(s.letterSpacing&&s.letterSpacing!='normal')q.sp=Math.max(0,Math.round(parseFloat(s.letterSpacing)/(S.k*PT)*100));if(s.textShadow&&s.textShadow!='none')q.sh=1;runs(c,q,out)}}
function writeText(e,el){const tb=k1(e,P,'txBody');if(!tb)return;const bp=k1(tb,A,'bodyPr');if(bp){const p=(el.style.padding||'').split(/\s+/),px=v=>Math.max(0,Math.round(parseFloat(v||0)/S.k));if(p.length===4){['tIns','rIns','bIns','lIns'].forEach((a,i)=>bp.setAttribute(a,px(p[i])))}}const s0=cur(),df=dflt(e,s0),dcl=dc(df,lum(bgc(s0))<.45).slice(1),olds=kids(tb,A,'p'),ps_=[];let cu=null;
 for(const n of el.childNodes){if(n.nodeName=='DIV'||n.nodeName=='P'){ps_.push(n);cu=null}else if(n.nodeName!='BR'||cu){if(!cu){cu=E('div');ps_.push(cu)}cu.append(n.cloneNode(true))}}
 if(!ps_.length)ps_.push(E('div'));let out='';const meta=[];
 for(const d of ps_){const rs=[];runs(d,{},rs);meta.push({d,rtl:d.dir=='rtl'||(d.dir!='ltr'&&/[\u0600-\u06FF]/.test(d.textContent))});
  out+=`<a:p>${rs.map(r=>{const z=r.s&&r.s/(S.k*PT);return `<a:r><a:rPr lang="ar-EG"${z&&Math.abs(z-df.sz)>.3?` sz="${Math.round(z*100)}"`:''}${r.b!=null&&r.b!=(df.b?1:0)?` b="${r.b}"`:''}${r.i!=null&&r.i!=(df.i?1:0)?` i="${r.i}"`:''}${r.u!=null?` u="${r.u?'sng':'none'}"`:''}${r.sp!=null?` sp="${r.sp}"`:''}>${r.sh?'<a:effectLst><a:outerShdw blurRad="63500" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="555555"><a:alpha val="55000"/></a:srgbClr></a:outerShdw></a:effectLst>':''}${r.c&&r.c!=dcl?`<a:solidFill><a:srgbClr val="${r.c}"/></a:solidFill>`:''}${r.f?`<a:latin typeface="${esc(r.f)}"/><a:cs typeface="${esc(r.f)}"/>`:''}</a:rPr><a:t>${esc(r.t)}</a:t></a:r>`}).join('')}${rs.length?'':'<a:endParaRPr lang="en-US"/>'}</a:p>`}
 const doc=e.ownerDocument,w=ps(`<w ${NSD}>${out}</w>`).documentElement;for(const p of olds)p.remove();
 [...w.children].forEach((c,i)=>{const{d,rtl}=meta[i],op=olds[i]&&k1(olds[i],A,'pPr'),al={left:'l',right:'r',center:'ctr',justify:'just'}[d.style.textAlign],bu=d.dataset.bu,ls=parseFloat(d.style.lineHeight),node=doc.importNode(c,true);
  if(op||rtl||al||bu||ls){const pe=op?doc.importNode(op,true):frag('<a:pPr/>',doc),add=n=>{const x=[...pe.children].find(y=>/^(tabLst|defRPr|extLst)$/.test(y.localName));x?pe.insertBefore(n,x):pe.append(n)};
   if(rtl)pe.setAttribute('rtl','1');if(al)pe.setAttribute('algn',al);else if(rtl&&!pe.getAttribute('algn'))pe.setAttribute('algn','r');
   if(ls){for(const x of kids(pe,A,'lnSpc'))x.remove();pe.insertBefore(frag(`<a:lnSpc><a:spcPct val="${Math.round(ls*100000)}"/></a:lnSpc>`,doc),pe.firstChild)}
   if(bu){for(const n of ['buNone','buChar','buAutoNum','buFont'])for(const x of kids(pe,A,n))x.remove();
    if(bu=='none')add(frag('<a:buNone/>',doc));else{if(!pe.getAttribute('marL')){pe.setAttribute('marL','342900');pe.setAttribute('indent','-342900')}add(frag(bu=='1'?'<a:buAutoNum type="arabicPeriod"/>':'<a:buFont typeface="Arial"/><a:buChar char="•"/>',doc))}}
   node.insertBefore(pe,node.firstChild)}
  tb.append(node)})}
function writeCell(tc,txt){const tb=k1(tc,A,'txBody'),rp=all(tb,A,'rPr')[0],rx=rp?ser(rp):'<a:rPr lang="ar-EG"/>';
 for(const p of kids(tb,A,'p'))p.remove();tb.append(frag(`<a:p><a:r>${rx}<a:t>${esc(txt)}</a:t></a:r></a:p>`,tc.ownerDocument))}
function commit(){if(!S)return;if(S.ed){const{e,el,h}=S.ed;if(el.innerHTML!==h||S.ed.changed){snap();writeText(e,el);touch()}el.contentEditable='false';el.parentNode.classList.remove('ed');S.ed=null}
 for(const c of S.stage.querySelectorAll('td'))if(c._tc&&c.textContent!==c._o){snap();writeCell(c._tc,c.textContent);c._o=c.textContent;touch()}}

// ---------- stage / selection ----------
function fit(){const b=S.sc,k=Math.min(b.clientWidth/S.W,b.clientHeight/S.H)*(S.zm||1);S.k=Math.max(k,.00001)}
function draw(){commit();fit();paint(cur(),S.stage,S.k,1);gridApply();const ok=S.sel&&[...S.stage.children].some(d=>d._e===S.sel);select(ok?S.sel:null);(S.tab=='obj'||S.tab=='nts')&&S.panel&&S.panel();thumb(S.cur)}
function thumb(i){const tn=S.th.children[i];if(tn)paint(S.slides[i],tn.querySelector('.pe-stage'),96/S.W,0)}
function thumbs(){S.th.innerHTML='';S.slides.forEach((s,i)=>{const d=E('div','pe-tn'+(i==S.cur?' on':'')),st=E('div','pe-stage'),grip=E('button','pe-drag','⠿');d._slide=s;grip.type='button';grip.title='اضغط واسحب لترتيب الشريحة';grip.setAttribute('aria-label','اسحب لترتيب الشريحة');d.append(E('b','',i+1),st,grip);d.onclick=()=>go(S.slides.indexOf(d._slide));
  let moved=false;grip.onpointerdown=e=>{e.preventDefault();e.stopPropagation();commit();moved=false;S.drag={el:d,slide:s,current:cur(),saved:false,pointer:e.pointerId};grip.setPointerCapture(e.pointerId);d.classList.add('dragging')};
  grip.onpointermove=e=>{const drag=S.drag;if(!drag||drag.el!==d||e.pointerId!==drag.pointer)return;e.preventDefault();const hit=document.elementFromPoint(e.clientX,e.clientY),target=hit&&hit.closest('.pe-tn');if(!target||target===d||target.parentElement!==S.th)return;
   const r=target.getBoundingClientRect(),before=e.clientX>r.left+r.width/2?target.nextSibling:target;if(before===d||before===d.nextSibling)return;
   if(!drag.saved){snap();drag.saved=true}S.th.insertBefore(d,before);S.slides=[...S.th.children].map(el=>el._slide);S.cur=S.slides.indexOf(drag.current);[...S.th.children].forEach((el,j)=>{el.querySelector('b').textContent=j+1;el.classList.toggle('on',el._slide===drag.current)});moved=true;
  };
  const finish=e=>{const drag=S.drag;if(!drag||drag.el!==d||e.pointerId!==drag.pointer)return;d.classList.remove('dragging');S.drag=null;if(moved){S.st=1;X.O().dirty=1;thumbs();draw()}};grip.onpointerup=finish;grip.onpointercancel=finish;S.th.append(d)});
 S.slides.forEach((s,i)=>thumb(i));const c=S.th.children[S.cur];c&&c.scrollIntoView({inline:'center',block:'nearest'});S.trs&&(S.trs.value=getTr(cur()))}
function go(i){commit();S.cur=i;S.sel=null;draw();[...S.th.children].forEach((d,j)=>d.classList.toggle('on',j==i));S.th.children[i].scrollIntoView({inline:'center',block:'nearest'});S.trs.value=getTr(cur())}
function select(e){if(S.sel!==e)S.sub=null;S.sel=e;for(const d of S.stage.children){d.classList.toggle('sel',!!e&&d._e===e);const h=d.querySelector(':scope>.pe-h');h&&h.remove();if(e&&d._e===e)d.append(E('i','pe-h'))}}
function startEdit(d){const t=d.querySelector('.pe-t');if(!t)return false;t.contentEditable='true';d.classList.add('ed');S.setTab&&S.setTab('txt');S.ed={e:d._e,el:t,h:t.innerHTML};t.focus();return true}
function ensureEdit(){if(S.ed)return true;const d=S.sel&&[...S.stage.children].find(x=>x._e===S.sel);if(!d||!startEdit(d))return false;
 const r=document.createRange();r.selectNodeContents(S.ed.el);const g=getSelection();g.removeAllRanges();g.addRange(r);S.rng=r.cloneRange();return true}
function restore(){if(S.ed&&S.rng){S.ed.el.focus();const g=getSelection();g.removeAllRanges();g.addRange(S.rng)}}
function smartAlign(){const was=!!S.ed;if(!ensureEdit())return X.toast('اختار صندوق نص الأول');restore();const el=S.ed.el;
 if(![...el.children].some(c=>/^(DIV|P)$/.test(c.nodeName))){const w=document.createElement('div');while(el.firstChild)w.append(el.firstChild);el.append(w)}
 let ps=[...el.children].filter(c=>/^(DIV|P)$/.test(c.nodeName));
 const g=getSelection(),r=g.rangeCount?g.getRangeAt(0):null;
 if(was&&r&&el.contains(r.commonAncestorContainer)){const s=ps.filter(x=>r.intersectsNode(x));if(s.length)ps=s}
 let ar=0,en=0;
 for(const q of ps){const t=q.textContent||'';let a=0,l=0;for(const ch of t){if(/[\u0590-\u08FF\uFB1D-\uFEFF]/.test(ch))a++;else if(/[A-Za-z]/.test(ch))l++}
  if(!a&&!l)continue;const rtl=a>=l;q.dir=rtl?'rtl':'ltr';const ta=q.style.textAlign;if(ta!=='center'&&ta!=='justify')q.style.textAlign=rtl?'right':'left';rtl?ar++:en++}
 X.toast(ar+en?'✨ محاذاة ذكية: '+ar+' عربي ➡ و '+en+' إنجليزي ⬅':'مفيش نص أطبّق عليه')}
 function fmt(cmd,val){if(!ensureEdit())return;restore();document.execCommand('styleWithCSS',false,true);document.execCommand(cmd,false,val)}
function spanStyle(prop,value){if(!ensureEdit())return false;restore();const g=getSelection();if(!g.rangeCount||g.isCollapsed){X.toast('حدد النص أولاً لتطبيق التنسيق');return false}const r=g.getRangeAt(0),sp=E('span');sp.style[prop]=value;try{sp.append(r.extractContents());r.insertNode(sp);g.removeAllRanges();const nr=document.createRange();nr.selectNodeContents(sp);g.addRange(nr);S.rng=nr.cloneRange();S.ed.changed=1;return true}catch(e){X.toast('تعذر تطبيق التنسيق على التحديد الحالي');return false}}
function charSpace(v){const n=+v;if(!Number.isFinite(n))return;spanStyle('letterSpacing',(n*PT*S.k)+'px')}
function textShadow(){if(!ensureEdit())return;restore();const g=getSelection();if(!g.rangeCount||g.isCollapsed){X.toast('حدد النص أولاً لتطبيق الظل');return}const r=g.getRangeAt(0),n=r.commonAncestorContainer.nodeType===1?r.commonAncestorContainer:r.commonAncestorContainer.parentElement,has=n&&n.closest('span')&&n.closest('span').style.textShadow;spanStyle('textShadow',has?'none':'1px 1px 2px #555')}
function textMargin(v){const e=S.sel;if(!e||!k1(k1(e,P,'txBody'),A,'bodyPr')){X.toast('حدد مربع نص أولاً');return}commit();snap();const bp=k1(k1(e,P,'txBody'),A,'bodyPr'),emu=Math.round(+v*914400);['lIns','rIns','tIns','bIns'].forEach(a=>bp.setAttribute(a,emu));touch();draw()}
function size(dir){if(!ensureEdit())return;restore();document.execCommand('fontSize',false,7);
 for(const f of S.ed.el.querySelectorAll('font[size="7"]')){const sp=E('span'),b=parseFloat(getComputedStyle(f.parentElement).fontSize)||16;sp.style.fontSize=Math.max(4,b*(dir>0?1.15:.87))+'px';while(f.firstChild)sp.append(f.firstChild);f.replaceWith(sp)}}
function sizeTo(pt){pt=+pt;if(!(pt>0)||!ensureEdit())return;restore();document.execCommand('styleWithCSS',false,false);document.execCommand('fontSize',false,7);
 for(const f of S.ed.el.querySelectorAll('font[size="7"]')){const sp=E('span');sp.style.fontSize=Math.min(999,pt)*S.k*PT+'px';while(f.firstChild)sp.append(f.firstChild);f.replaceWith(sp)}document.execCommand('styleWithCSS',false,true)}
function setFill(h){const e=S.sel;if(!e||e.localName!='sp')return;snap();const sp=k1(e,P,'spPr');
 for(const c of kids(sp,A,'solidFill').concat(kids(sp,A,'noFill'),kids(sp,A,'gradFill'),kids(sp,A,'pattFill')))c.remove();
 const f=frag(`<a:solidFill><a:srgbClr val="${h.slice(1)}"/></a:solidFill>`,e.ownerDocument),ln=k1(sp,A,'ln');ln?sp.insertBefore(f,ln):sp.append(f);touch();draw()}

const GK={h:['→ أفقي',0],v:['↓ رأسي',5400000],d:['↘ قطري',2700000],d2:['↙ قطري',8100000],r:['◎ دائري',0]};
const gradXml=(c1,c2,k)=>`<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="${c1}"/></a:gs><a:gs pos="100000"><a:srgbClr val="${c2}"/></a:gs></a:gsLst>${k=='r'?'<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>':`<a:lin ang="${GK[k][1]}" scaled="0"/>`}</a:gradFill>`;
function gradShape(c1,c2,k){const e=S.sel;if(!e||e.localName!='sp')return false;snap();const sp=k1(e,P,'spPr');
 for(const n of ['solidFill','noFill','gradFill','pattFill','blipFill'])kids(sp,A,n).forEach(x=>x.remove());
 const f=frag(gradXml(c1,c2,k),e.ownerDocument),ref=[...sp.children].find(n=>/^(ln|effectLst|effectDag|scene3d|sp3d|extLst)$/.test(n.localName));ref?sp.insertBefore(f,ref):sp.append(f);touch();draw();return true}
function gradBg(c1,c2,k,allSl){snap();for(const s of (allSl?S.slides:[cur()])){const cs=all(s.doc,P,'cSld')[0];if(!cs)continue;for(const b of kids(cs,P,'bg'))b.remove();
 cs.insertBefore(frag(`<p:bg><p:bgPr>${gradXml(c1,c2,k)}<a:effectLst/></p:bgPr></p:bg>`,s.doc),cs.firstChild);s.t=1}X.O().dirty=1;draw();allSl?thumbs():thumb(S.cur)}
function gradDlg(){const old=document.querySelector('.pe-gd');if(old){old.remove();return}
 const canFill=!!(S.sel&&S.sel.localName=='sp');let tg=canFill?'fill':'bg',k='v',c1='#2563eb',c2='#9333ea',allSl=false;
 const GP=[['#2563eb','#9333ea'],['#f97316','#ec4899'],['#06b6d4','#3b82f6'],['#22c55e','#065f46'],['#fde047','#f97316'],['#0f172a','#475569'],['#f43f5e','#fb923c'],['#e0e7ff','#fce7f3']];
 const w=E('div','pe-ds pe-gd','<div class="pds-h"><b>🌈 تدرج لوني</b><button type="button" aria-label="إغلاق">✕</button></div>');
 const row=()=>{const r=E('div');r.style.cssText='display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:8px 0';return r};
 const bt=(t,f)=>{const b=E('button',null,t);b.type='button';b.style.cssText='padding:6px 10px;border-radius:10px;border:1px solid #cfd6e6;background:#f3f6ff;font:inherit;cursor:pointer';b.onclick=f;return b};
 const css=()=>k=='r'?`radial-gradient(circle,${c1},${c2})`:`linear-gradient(${GK[k][1]/60000+90}deg,${c1},${c2})`;
 const pv=E('div');pv.style.cssText='height:34px;border-radius:10px;border:1px solid #cfd6e6;margin-bottom:8px';
 const apply=()=>{pv.style.background=css();const h=x=>x.slice(1).toUpperCase();
  if(tg=='fill'){if(!gradShape(h(c1),h(c2),k))X.toast('اختار شكل الأول')}else gradBg(h(c1),h(c2),k,allSl)};
 let bF,bB;const ar=row();
 const mk=()=>{bF.style.outline=tg=='fill'?'2px solid #2563eb':'';bB.style.outline=tg=='bg'?'2px solid #2563eb':'';ar.style.display=tg=='bg'?'flex':'none';bF.style.opacity=canFill?1:.5};
 const tr=row();bF=bt('🎨 تعبئة الشكل',()=>{if(!canFill)return X.toast('اختار شكل الأول');tg='fill';mk()});bB=bt('🖼 خلفية الشريحة',()=>{tg='bg';mk()});tr.append(bF,bB);
 const cb=E('input');cb.type='checkbox';cb.onchange=()=>{allSl=cb.checked};const cl=E('label');cl.style.cssText='display:flex;gap:6px;align-items:center';cl.append(cb,document.createTextNode('تطبيق على كل الشرائح'));ar.append(cl);
 const i1=E('input'),i2=E('input');
 const pr=row();for(const [a,b] of GP){const s=bt('',()=>{c1=a;c2=b;i1.value=a;i2.value=b;apply()});s.style.cssText='width:46px;height:30px;border-radius:8px;border:1px solid #cfd6e6;cursor:pointer;background:linear-gradient(90deg,'+a+','+b+')';pr.append(s)}
 const cr=row();for(const [i,v,f] of [[i1,c1,x=>{c1=x}],[i2,c2,x=>{c2=x}]]){i.type='color';i.value=v;i.style.cssText='width:44px;height:34px;border:none;padding:0;background:none';i.onchange=()=>{f(i.value);apply()}}
 cr.append(E('span',null,'من'),i1,E('span',null,'إلى'),i2);
 const dr=row();for(const kk of Object.keys(GK))dr.append(bt(GK[kk][0],()=>{k=kk;apply()}));
 const ap=bt('✔ تطبيق',apply);ap.style.background='#2563eb';ap.style.color='#fff';
 w.append(tr,ar,pr,cr,dr,pv,ap);pv.style.background=css();mk();
 w.querySelector('.pds-h button').onclick=()=>w.remove();document.body.append(w)}
function bgSet(h){snap();const cs=all(cur().doc,P,'cSld')[0];for(const b of kids(cs,P,'bg'))b.remove();cs.insertBefore(frag(`<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${h.slice(1)}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`,cur().doc),cs.firstChild);touch();draw();thumb(S.cur)}
const zoom=f=>{S.zm=Math.min(4,Math.max(1,S.zm*f));draw()};
function pinch(){const sc=S.sc,D=e=>Math.hypot(e.touches[0].clientX-e.touches[1].clientX,e.touches[0].clientY-e.touches[1].clientY),mid=e=>[(e.touches[0].clientX+e.touches[1].clientX)/2,(e.touches[0].clientY+e.touches[1].clientY)/2],
 ang=e=>Math.atan2(e.touches[1].clientY-e.touches[0].clientY,e.touches[1].clientX-e.touches[0].clientX)*180/Math.PI,nrm=d=>{while(d>180)d-=360;while(d<-180)d+=360;return d};
 let d0=1,f=1,lx=0,ly=0,pa=0,cum=0,rm=0,rd=null,base=0;
 // two fingers = pinch-zoom + pan; if both fingers sit on the selected element and twist, it rotates the element instead
 sc.addEventListener('touchstart',e=>{if(e.touches.length!=2)return;S.pz=1;d0=D(e)||1;f=1;pa=ang(e);cum=0;rm=0;const sr=S.stage.getBoundingClientRect();[lx,ly]=mid(e);S.stage.style.transformOrigin=(lx-sr.left)+'px '+(ly-sr.top)+'px';
  const d=S.sel&&S.sel.localName!='graphicFrame'&&[...S.stage.children].find(x=>x._e===S.sel);rd=d&&d.contains(e.touches[0].target)&&d.contains(e.touches[1].target)?d:null;
  if(rd){const x=xf(S.sel);base=x&&+x.getAttribute('rot')/60000||0}},{passive:true});
 sc.addEventListener('touchmove',e=>{if(!S.pz||e.touches.length<2)return;e.preventDefault();const a=ang(e);cum+=nrm(a-pa);pa=a;
  if(!rm&&rd&&Math.abs(cum)>15&&Math.abs(D(e)/d0-1)<.25){rm=1;f=1;S.stage.style.transform=''}
  if(rm){rd.style.transform='rotate('+(base+cum)+'deg)';return}
  const[cx,cy]=mid(e);sc.scrollLeft-=cx-lx;sc.scrollTop-=cy-ly;lx=cx;ly=cy;f=Math.min(4/S.zm,Math.max(1/S.zm,D(e)/d0));S.stage.style.transform='scale('+f+')'},{passive:false});
 const end=e=>{if(!S.pz||e.touches.length>1)return;
  if(rm){rm=0;let r=(((base+cum)%360)+360)%360;for(const t of [0,45,90,135,180,225,270,315,360])if(Math.abs(r-t)<4)r=t%360;
   snap();const x=xf(rd._e,1);x.setAttribute('rot',Math.round(r*60000)%21600000);touch();S.stage.style.transformOrigin='';draw();f=1;setTimeout(()=>{S.pz=0},0);return}
  const sr=S.stage.getBoundingClientRect(),fx=(lx-sr.left)/(sr.width||1),fy=(ly-sr.top)/(sr.height||1);S.stage.style.transform='';S.stage.style.transformOrigin='';
  if(Math.abs(f-1)>.01){S.zm=Math.min(4,Math.max(1,S.zm*f));draw();const r2=S.stage.getBoundingClientRect();sc.scrollLeft+=r2.left+fx*r2.width-lx;sc.scrollTop+=r2.top+fy*r2.height-ly}
  f=1;setTimeout(()=>{S.pz=0},0)};
 sc.addEventListener('touchend',end,{passive:true});sc.addEventListener('touchcancel',end,{passive:true})}
// ---------- شبكة التصميم + خطوط المحاذاة الذكية (التقاط أثناء السحب والتكبير) ----------
const GP=(k,d)=>{try{const v=localStorage.getItem(k);return v==null?d:v}catch(_){return d}},GS=(k,v)=>{try{localStorage.setItem(k,v)}catch(_){}};
const GD={smart:GP('raPeSmart','1')=='1',show:GP('raPeGrid','0')=='1',snap:GP('raPeGridSnap','0')=='1',cm:+GP('raPeGridCm','1')||1};
function gridApply(){if(!S||!S.stage)return;const o=S.stage.querySelector(':scope>.pe-grid');o&&o.remove();if(!GD.show)return;
 const z=Math.max(4,GD.cm*360000*S.k),g=E('div','pe-grid');
 g.style.cssText='position:absolute;left:0;top:0;width:100%;height:100%;pointer-events:none;z-index:50;background-image:linear-gradient(to right,rgba(37,99,235,.35) 1px,transparent 1px),linear-gradient(to bottom,rgba(37,99,235,.35) 1px,transparent 1px);background-size:'+z+'px '+z+'px';S.stage.append(g)}
function guidesClear(){S&&S.stage&&S.stage.querySelectorAll(':scope>.pe-gd').forEach(n=>n.remove())}
function snapTargets(skip){const t={x:[0,S.W/2,S.W],y:[0,S.H/2,S.H]};
 for(const e of tree(cur()).children){if(e===skip||e.namespaceURI!=P||!/^(sp|pic|graphicFrame)$/.test(e.localName))continue;const c=all(e,P,'cNvPr')[0];if(c&&c.getAttribute('hidden')=='1')continue;const g=geo(e,cur());if(!g)continue;t.x.push(g.x,g.x+g.w/2,g.x+g.w);t.y.push(g.y,g.y+g.h/2,g.y+g.h)}return t}
function guideSnap(g,dx,dy,rz,skip){guidesClear();if(!GD.smart&&!GD.snap)return{dx,dy};
 const k=S.k,thr=7/k,T=GD.smart?snapTargets(skip):{x:[],y:[]},gs=GD.cm*360000,lines=[];
 const best=(vals,ts)=>{let b=null;if(GD.smart)for(const v of vals)for(const t of ts){const d=t-v;if(Math.abs(d)<=thr&&(!b||Math.abs(d)<Math.abs(b.d)))b={d,line:t}}
  if(!b&&GD.snap){let m=null;for(const v of vals){const d=Math.round(v/gs)*gs-v;if(m==null||Math.abs(d)<Math.abs(m))m=d}b={d:m}}return b};
 const nx=g.x+dx/k,ny=g.y+dy/k,vx=rz?[g.x+g.w+dx/k]:[nx,nx+g.w/2,nx+g.w],vy=rz?[g.y+g.h+dy/k]:[ny,ny+g.h/2,ny+g.h],bx=best(vx,T.x),by=best(vy,T.y);
 if(bx){dx+=bx.d*k;if(bx.line!=null)lines.push(['v',bx.line])}if(by){dy+=by.d*k;if(by.line!=null)lines.push(['h',by.line])}
 for(const[t,v]of lines){const L=E('i','pe-gd');L.style.cssText='position:absolute;pointer-events:none;z-index:60;'+(t=='v'?'left:'+(v*k)+'px;top:0;width:0;height:100%;border-left:1px dashed #e11d9a':'top:'+(v*k)+'px;left:0;height:0;width:100%;border-top:1px dashed #e11d9a');S.stage.append(L)}
 return{dx,dy}}
function bind(){const st=S.stage;st.addEventListener('pointerdown',ev=>{
 const d=ev.target.closest('.pe-s');if(S.ed&&d&&d._e===S.ed.e)return;
 if(!d||!d._e){commit();select(null);return}
 const cell=ev.target.closest('td');if(cell&&S.sel===d._e){cell.contentEditable='true';cell.focus();return}
 const was=S.sel===d._e,rz=!!ev.target.closest('.pe-h');commit();select(d._e);S.setTab(S.tab=='txt'&&d.querySelector('.pe-t')?'txt':'obj');const g=geo(d._e,cur()),x0=ev.clientX,y0=ev.clientY;let mv=0;
 let sx=0,sy=0;const pm=m=>{if(S.pz)return;let dx=m.clientX-x0,dy=m.clientY-y0;if(Math.abs(dx)+Math.abs(dy)>4)mv=1;if(!mv)return;({dx,dy}=guideSnap(g,dx,dy,rz,d._e));sx=dx;sy=dy;
  if(rz){d.style.width=Math.max(12,g.w*S.k+dx)+'px';d.style.height=Math.max(12,g.h*S.k+dy)+'px'}else{d.style.left=g.x*S.k+dx+'px';d.style.top=g.y*S.k+dy+'px'}};
 const pu=m=>{removeEventListener('pointermove',pm);removeEventListener('pointerup',pu);guidesClear();
  if(S.pz);else if(mv){const dx=sx/S.k,dy=sy/S.k;snap();setGeo(d._e,rz?{x:g.x,y:g.y,w:g.w+dx,h:g.h+dy}:{x:g.x+dx,y:g.y+dy,w:g.w,h:g.h});touch();draw()}
  else if(was&&!rz)startEdit(d)};
 addEventListener('pointermove',pm);addEventListener('pointerup',pu)});st.addEventListener('focusin',e=>{const c=e.target.closest&&e.target.closest('td');if(c&&c._tc)S.cell=c})}

// ---------- insert ----------
const nid=()=>Math.max(1,...all(cur().doc,P,'cNvPr').map(x=>+x.getAttribute('id')||0))+1;
const ctr=(w,h)=>({x:Math.round((S.W-w)/2),y:Math.round((S.H-h)/2),w:Math.round(w),h:Math.round(h)});
function addEl(x){commit();snap();S.setTab('obj');const n=frag(x,cur().doc);tree(cur()).append(n);touch();S.sel=n;draw();thumb(S.cur)}
const spX=(g,geom,fill,text)=>`<p:sp><p:nvSpPr><p:cNvPr id="${nid()}" name="Shape ${nid()}"/><p:cNvSpPr${text?' txBox="1"':''}/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${g.x}" y="${g.y}"/><a:ext cx="${g.w}" cy="${g.h}"/></a:xfrm><a:prstGeom prst="${geom}"><a:avLst/></a:prstGeom>${fill?`<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`:'<a:noFill/>'}</p:spPr><p:txBody><a:bodyPr wrap="square" anchor="${text?'t':'ctr'}"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p>${text?`<a:r><a:rPr lang="ar-EG" sz="2400"/><a:t>${esc(text)}</a:t></a:r>`:'<a:endParaRPr lang="en-US"/>'}</a:p></p:txBody></p:sp>`;
const addText=()=>addEl(spX(ctr(S.W*.6,S.H*.18),'rect',null,'نص جديد'));
// ---------- مخططات جاهزة: هياكل تنظيمية / تسلسلات هرمية / عمليات / علاقات ----------
const DPAL=['1F4E79','2E75B6','2E9E6B','E8912D','C0504D','7B5EA7'];
const DG=[
 ['org','هيكل تنظيمي'],['treeh','شجرة هرمية أفقية'],['pyramid','هرم متدرج'],
 ['proch','عملية أفقية'],['procv','عملية رأسية'],['cycle','دورة متكررة'],
 ['hub','علاقة محورية'],['venn','مخطط فن'],['matrix','مصفوفة ٢×٢'],['timeline','خط زمني']];
// يرجع قائمة عناصر بإحداثيات نسبية (0..1) من الشريحة: {k:'b',x,y,w,h,t,prst,fill,tc,alpha,sz,bold} و {k:'l',x1,y1,x2,y2,arrow,both,dash,c}
function dgItems(kind){const L=[],B=[],u=S.H/S.W,
 box=(x,y,w,h,t,o={})=>B.push({k:'b',x,y,w,h,t,prst:'roundRect',fill:DPAL[0],tc:'FFFFFF',...o}),
 cb=(cx,cy,w,h,t,o)=>box(cx-w/2,cy-h/2,w,h,t,o),
 circ=(cx,cy,d,t,o={})=>box(cx-d*u/2,cy-d/2,d*u,d,t,{prst:'ellipse',...o}),
 ln=(x1,y1,x2,y2,o={})=>L.push({k:'l',x1,y1,x2,y2,...o});
 switch(kind){
 case 'org':{const mc=[.2,.5,.8],mt=['المدير المالي','مدير التسويق','مدير العمليات'],st=['المحاسبة','المشتريات','الإعلانات','المبيعات','الإنتاج','الجودة'];
  ln(.5,.35,.5,.4);ln(.2,.4,.8,.4);
  mc.forEach((c,i)=>{ln(c,.4,c,.45);ln(c,.55,c,.615);ln(c-.08,.615,c+.08,.615);
   [-.08,.08].forEach((d,j)=>{ln(c+d,.615,c+d,.68);cb(c+d,.725,.12,.09,st[i*2+j],{fill:DPAL[2],sz:.8})});
   cb(c,.5,.22,.1,mt[i],{fill:DPAL[1]})});
  cb(.5,.295,.24,.11,'المدير العام',{fill:DPAL[0]});break}
 case 'treeh':{const l3=['المحاسبة','المشتريات','الإعلانات','المبيعات','الإنتاج','الجودة'],l2=['الإدارة المالية','الإدارة التسويقية','الإدارة التشغيلية'],c3=[.31,.43,.55,.67,.79,.91],c2=[.37,.61,.85];
  ln(.7,.61,.66,.61);ln(.66,.37,.66,.85);
  c2.forEach((c,i)=>{ln(.66,c,.62,c);ln(.38,c,.34,c);ln(.34,c3[i*2],.34,c3[i*2+1]);
   [0,1].forEach(j=>{ln(.34,c3[i*2+j],.3,c3[i*2+j]);cb(.18,c3[i*2+j],.24,.09,l3[i*2+j],{fill:DPAL[2],sz:.85})});
   cb(.5,c,.24,.1,l2[i],{fill:DPAL[1]})});
  cb(.82,.61,.24,.11,'المدير العام',{fill:DPAL[0]});break}
 case 'pyramid':{const t=['القيادة','الإدارة العليا','الإدارة الوسطى','التنفيذ'],w=[.22,.42,.62,.82],f=[4,3,2,1];
  t.forEach((x,i)=>cb(.5,.33+i*.16,w[i],.14,x,{fill:DPAL[f[i]]}));break}
 case 'proch':{for(let i=0;i<4;i++){const x=.78-i*.24;
   cb(x+.08,.44,.16,.16,'الخطوة '+['الأولى','الثانية','الثالثة','الرابعة'][i],{fill:DPAL[i+1]});
   circ(x+.08,.33,.075,String(i+1),{fill:DPAL[0],sz:.9});
   cb(x+.08,.64,.16,.14,'وصف الخطوة',{prst:'rect',fill:null,tc:'404040',sz:.8,bold:0});
   i<3&&ln(x-.005,.44,x-.075,.44,{arrow:1})}break}
 case 'procv':{const t=['المرحلة الأولى','المرحلة الثانية','المرحلة الثالثة','المرحلة الرابعة'];
  t.forEach((x,i)=>{cb(.5,.3+i*.17,.5,.08,x,{fill:DPAL[i+1]});i<3&&ln(.5,.34+i*.17,.5,.43+i*.17,{arrow:1})});break}
 case 'cycle':{const n=[[.5,.35,'التخطيط'],[.8,.6,'التنفيذ'],[.5,.85,'المتابعة'],[.2,.6,'التحسين']];
  ln(.6,.35,.8,.54,{arrow:1});ln(.8,.66,.6,.85,{arrow:1});ln(.4,.85,.2,.66,{arrow:1});ln(.2,.54,.4,.35,{arrow:1});
  n.forEach((q,i)=>cb(q[0],q[1],.2,.12,q[2],{fill:DPAL[i+1]}));circ(.5,.6,.2,'الهدف',{fill:DPAL[0]});break}
 case 'hub':{const p=[[.5,.34],[.786,.47],[.786,.73],[.5,.86],[.214,.73],[.214,.47]],t=['عنصر ١','عنصر ٢','عنصر ٣','عنصر ٤','عنصر ٥','عنصر ٦'];
  p.forEach(q=>ln(.5,.6,q[0],q[1]));
  p.forEach((q,i)=>cb(q[0],q[1],.2,.1,t[i],{fill:DPAL[i%5+1]}));circ(.5,.6,.22,'الفكرة الرئيسية',{fill:DPAL[0],sz:.9});break}
 case 'venn':{const r=.21,cy=.6,c=[[0,-.11,1],[-.12,.1,3],[.12,.1,2]];
  c.forEach(q=>box(.5+(q[0]-r)*u,cy+q[1]-r,2*r*u,2*r,'',{prst:'ellipse',fill:DPAL[q[2]],alpha:55000}));
  [['المجموعة الأولى',0,-.22],['المجموعة الثانية',-.21,.2],['المجموعة الثالثة',.21,.2],['المشترك',0,.03]].forEach(q=>cb(.5+q[1]*u,cy+q[2],.16,.08,q[0],{prst:'rect',fill:null,tc:'262626',sz:.85}));break}
 case 'matrix':{const t=[['مهم وغير عاجل',.18,.27,1],['مهم وعاجل',.52,.27,4],['غير مهم وغير عاجل',.18,.58,2],['عاجل وغير مهم',.52,.58,3]];
  t.forEach(q=>box(q[1],q[2],.3,.28,q[0],{fill:DPAL[q[3]],sz:1.1}));break}
 case 'timeline':{ln(.94,.55,.06,.55,{arrow:1,c:'7F7F7F'});
  [.8,.58,.36,.14].forEach((x,i)=>{const up=i%2==0;ln(x,.55,x,up?.45:.65,{c:'7F7F7F'});circ(x,.55,.07,String(i+1),{fill:DPAL[i+1],sz:.85});
   cb(x,up?.37:.73,.19,.14,['البداية','التطوير','الإطلاق','النمو'][i],{fill:DPAL[i+1]})});break}
 }
 return L.concat(B)}
function dgXml(kind){const W=S.W,H=S.H,fs=Math.max(1000,Math.round(1600*H/5143500)),items=dgItems(kind);let n=nid();
 return items.map(q=>{if(q.k=='b'){const g={x:Math.round(q.x*W),y:Math.round(q.y*H),w:Math.round(q.w*W),h:Math.round(q.h*H)},id=n++,
   fl=q.fill?`<a:solidFill><a:srgbClr val="${q.fill}">${q.alpha?`<a:alpha val="${q.alpha}"/>`:''}</a:srgbClr></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>`:'<a:noFill/>',
   tx=q.t?`<a:pPr algn="ctr"/><a:r><a:rPr lang="ar-EG" sz="${Math.round(fs*(q.sz||1))}" b="${q.bold===0?0:1}"><a:solidFill><a:srgbClr val="${q.tc}"/></a:solidFill></a:rPr><a:t>${esc(q.t)}</a:t></a:r>`:'<a:endParaRPr lang="ar-EG"/>';
   return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${g.x}" y="${g.y}"/><a:ext cx="${g.w}" cy="${g.h}"/></a:xfrm><a:prstGeom prst="${q.prst}"><a:avLst/></a:prstGeom>${fl}</p:spPr><p:txBody><a:bodyPr wrap="square" lIns="45720" rIns="45720" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p>${tx}</a:p></p:txBody></p:sp>`}
  const x1=q.x1*W,y1=q.y1*H,x2=q.x2*W,y2=q.y2*H,id=n++,fh=x2<x1?' flipH="1"':'',fv=y2<y1?' flipV="1"':'';
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Connector ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${fh}${fv}><a:off x="${Math.round(Math.min(x1,x2))}" y="${Math.round(Math.min(y1,y2))}"/><a:ext cx="${Math.round(Math.abs(x2-x1))}" cy="${Math.round(Math.abs(y2-y1))}"/></a:xfrm><a:prstGeom prst="${q.arrow?'straightConnector1':'line'}"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="28575"><a:solidFill><a:srgbClr val="${q.c||'595959'}"/></a:solidFill><a:round/>${q.arrow?'<a:tailEnd type="triangle" w="med" len="med"/>':''}</a:ln></p:spPr></p:sp>`}) }
function addEls(a){commit();snap();S.setTab('obj');for(const x of a)tree(cur()).append(frag(x,cur().doc));touch();S.sel=null;draw();thumb(S.cur)}
function dgPrev(kind){const it=dgItems(kind),V=(v,m)=>(v*m).toFixed(1);let o='<svg viewBox="0 0 160 90" aria-hidden="true">';
 for(const q of it){if(q.k=='l'){o+=`<path d="M${V(q.x1,160)} ${V(q.y1,90)}L${V(q.x2,160)} ${V(q.y2,90)}" stroke="${q.c?'#'+q.c:'currentColor'}" stroke-opacity=".7" stroke-width="1.6" fill="none"/>`;
   if(q.arrow){const a=Math.atan2((q.y2-q.y1)*90,(q.x2-q.x1)*160),x=q.x2*160,y=q.y2*90,t=(d,r)=>`${(x-Math.cos(a+d)*r).toFixed(1)} ${(y-Math.sin(a+d)*r).toFixed(1)}`;o+=`<path d="M${x.toFixed(1)} ${y.toFixed(1)}L${t(.45,5)}L${t(-.45,5)}Z" fill="${q.c?'#'+q.c:'currentColor'}" fill-opacity=".7"/>`}}
  else{const f=q.fill?'#'+q.fill:'none',op=q.alpha?q.alpha/1e5:1,x=q.x*160,y=q.y*90,w=q.w*160,h=q.h*90;
   if(!q.fill)continue;o+=q.prst=='ellipse'?`<ellipse cx="${(x+w/2).toFixed(1)}" cy="${(y+h/2).toFixed(1)}" rx="${(w/2).toFixed(1)}" ry="${(h/2).toFixed(1)}" fill="${f}" fill-opacity="${op}"/>`:`<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${h.toFixed(1)}" rx="3" fill="${f}"/>`}}
 return o+'</svg>'}
function dgGallery(){const o=document.querySelector('.dgw');if(o){o.remove();return}
 const w=E('div','shg dgw'),box=E('div','shb','<div class="shh"><b>🧩 مخططات جاهزة</b><button type="button" aria-label="إغلاق">✕</button></div>'),close=()=>w.remove();
 box.querySelector('button').onclick=close;w.onpointerdown=e=>{if(e.target===w)close()};
 const h=E('h5');h.textContent='اختار مخطط وبيتضاف على الشريحة وتقدر تعدّل النصوص والألوان وتحرّك كل عنصر';const g=E('div','dgg');
 for(const[id,nm]of DG){const b=E('button','dgt',dgPrev(id)+'<span></span>');b.querySelector('span').textContent=nm;b.title=nm;b.onclick=()=>{close();addEls(dgXml(id))};g.append(b)}
 box.append(h,g);w.append(box);document.body.append(w)}
// موصلات: خطوط وأسهم بين الأشكال (تتحفظ كأشكال PowerPoint عادية بهندسة موصل)
const CN={cnLine:['خط','line','',''],cnArrow:['سهم','straightConnector1','','triangle'],cnBoth:['سهم بالاتجاهين','straightConnector1','triangle','triangle'],
 cnDash:['سهم متقطع','straightConnector1','','triangle',1],cnElbow:['موصل مفصلي','bentConnector3','','triangle'],cnCurve:['موصل منحني','curvedConnector3','','triangle']},
CNR=/^(line|straightConnector1|bentConnector3|curvedConnector3)$/,
CNP={line:'M5 50H95',straightConnector1:'M5 50H95',bentConnector3:'M5 85H50V15H95',curvedConnector3:'M5 85C50 85 50 15 95 15'};
function addConn(id){const q=CN[id],bend=/^(bent|curved)/.test(q[1]),g=ctr(bend?S.W*.25:S.W*.3,bend?S.H*.2:0),
 hd=q[2]?`<a:headEnd type="${q[2]}" w="med" len="med"/>`:'',tl=q[3]?`<a:tailEnd type="${q[3]}" w="med" len="med"/>`:'';
 addEl(`<p:sp><p:nvSpPr><p:cNvPr id="${nid()}" name="Connector ${nid()}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${g.x}" y="${g.y}"/><a:ext cx="${g.w}" cy="${g.h}"/></a:xfrm><a:prstGeom prst="${q[1]}"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="28575"><a:solidFill><a:srgbClr val="404040"/></a:solidFill>${q[4]?'<a:prstDash val="dash"/>':''}<a:round/>${hd}${tl}</a:ln></p:spPr></p:sp>`)}
let CNU=0;
function connDraw(d,g,k,pr,lc,lw,xr,l){const W=g.w*k,H=g.h*k,dw=Math.max(W,18),dh=Math.max(H,18),ox=(dw-W)/2,oy=(dh-H)/2,c=lc||'#404040',u='cn'+(++CNU),
 he=l&&all(l,A,'headEnd')[0],te=l&&all(l,A,'tailEnd')[0],hs=he&&he.getAttribute('type')&&he.getAttribute('type')!='none',ts=te&&te.getAttribute('type')&&te.getAttribute('type')!='none',
 dash=l&&all(l,A,'prstDash')[0]&&all(l,A,'prstDash')[0].getAttribute('val')!='solid',
 path=pr=='bentConnector3'?`M0 0H${W/2}V${H}H${W}`:pr=='curvedConnector3'?`M0 0C${W/2} 0 ${W/2} ${H} ${W} ${H}`:`M0 0L${W} ${H}`,
 mk=`<marker id="${u}" viewBox="0 0 10 10" refX="8" refY="5" markerUnits="strokeWidth" markerWidth="4" markerHeight="4" orient="auto-start-reverse"><path d="M0 0L10 5L0 10Z" fill="${c}"/></marker>`;
 d.style.left=(parseFloat(d.style.left)-ox)+'px';d.style.top=(parseFloat(d.style.top)-oy)+'px';d.style.width=dw+'px';d.style.height=dh+'px';
 const fh=xr&&xr.getAttribute('flipH')=='1',fv=xr&&xr.getAttribute('flipV')=='1',
 sv=E('div');sv.style.cssText='position:absolute;inset:0;pointer-events:none';
 sv.innerHTML=`<svg width="${dw}" height="${dh}" viewBox="0 0 ${dw} ${dh}" style="overflow:visible"><defs>${mk}</defs><g transform="translate(${fh?ox+W:ox},${fv?oy+H:oy}) scale(${fh?-1:1},${fv?-1:1})"><path d="${path}" fill="none" stroke="${c}" stroke-width="${lw||2}" stroke-linejoin="round"${dash?` stroke-dasharray="${(lw||2)*3} ${(lw||2)*2}"`:''}${hs?` marker-start="url(#${u})"`:''}${ts?` marker-end="url(#${u})"`:''}/></g></svg>`;
 d.append(sv)}
// أشكال المخططات الانسيابية: تتضاف بلون حسب نوعها ونص عربي جاهز في النص وحدود بيضا
const FLC={flowChartTerminator:'2E9E6B',flowChartDecision:'E8912D',flowChartSort:'E8912D',flowChartData:'7B5EA7',flowChartManualInput:'7B5EA7',flowChartDocument:'3A8FB7',flowChartMultidocument:'3A8FB7',flowChartMagneticDisk:'6B7A8F',flowChartStoredData:'6B7A8F',flowChartInternalStorage:'6B7A8F',flowChartOfflineStorage:'6B7A8F',flowChartConnector:'C0504D',flowChartOffpageConnector:'C0504D',flowChartOr:'C0504D',flowChartSummingJunction:'C0504D'};
const FLR=/^flowChart(Connector|Or|SummingJunction)$/;
function addFlow(g){const rd=FLR.test(g),w=rd?S.H*.16:S.W*.22,h=rd?S.H*.16:S.H*.16;
 const x=spX(ctr(w,h),g,FLC[g]||'4F81BD').replace('<a:endParaRPr lang="en-US"/>','<a:pPr algn="ctr"/><a:r><a:rPr lang="ar-EG" sz="'+(rd?1400:1800)+'" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>'+esc(SH[g][0])+'</a:t></a:r>').replace('</p:spPr>','<a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></p:spPr>');addEl(x)}
const addShape=g=>/^flowChart/.test(g)?addFlow(g):addEl(spX(ctr(S.W*.25,S.H*.25),g,'4F81BD'));
// ---------- مكتبة الأشكال: كل شكل = اسم OOXML + مسار SVG (0..100) للمعاينة في المحرر ----------
const SH=(()=>{const f=v=>+v.toFixed(1),poly=a=>'M'+a.map(q=>q.map(f).join(' ')).join('L')+'Z',
 star=(n,r)=>poly(Array.from({length:n*2},(_,i)=>{const a=-Math.PI/2+i*Math.PI/n,R=i%2?50*r:50;return[50+R*Math.cos(a),50+R*Math.sin(a)]})),
 sun=()=>{let d='M28 50A22 22 0 1 0 72 50A22 22 0 1 0 28 50Z';for(let i=0;i<8;i++){const a=i*Math.PI/4,t=(r,o)=>[50+r*Math.cos(a+o),50+r*Math.sin(a+o)];d+=poly([t(30,-.22),t(50,0),t(30,.22)])}return d},
 R=(a,b,c,d)=>'M'+a+' '+b+'H'+c+'V'+d+'Z';
 return{
 rect:['مستطيل','M0 0H100V100H0Z'],roundRect:['مستطيل دائري','M14 0H86Q100 0 100 14V86Q100 100 86 100H14Q0 100 0 86V14Q0 0 14 0Z'],
 ellipse:['دائرة','M0 50A50 50 0 1 0 100 50A50 50 0 1 0 0 50Z'],triangle:['مثلث','M50 0L100 100H0Z'],rtTriangle:['مثلث قائم','M0 0L100 100H0Z'],
 diamond:['معيّن','M50 0L100 50L50 100L0 50Z'],parallelogram:['متوازي أضلاع','M25 0H100L75 100H0Z'],trapezoid:['شبه منحرف','M20 0H80L100 100H0Z'],
 pentagon:['خماسي','M50 0L100 38L81 100H19L0 38Z'],hexagon:['سداسي','M25 0H75L100 50L75 100H25L0 50Z'],octagon:['ثماني','M30 0H70L100 30V70L70 100H30L0 70V30Z'],
 plus:['علامة +','M33 0H67V33H100V67H67V100H33V67H0V33H33Z'],
 donut:['حلقة','M0 50A50 50 0 1 0 100 50A50 50 0 1 0 0 50ZM25 50A25 25 0 1 1 75 50A25 25 0 1 1 25 50Z','eo'],
 frame:['إطار','M0 0H100V100H0ZM14 14V86H86V14Z','eo'],can:['أسطوانة','M0 14A50 14 0 0 1 100 14V86A50 14 0 0 1 0 86Z'],
 cube:['مكعب','M0 25L25 0H100V75L75 100H0Z'],foldedCorner:['ورقة مطوية','M0 0H100V70L70 100H0Z'],
 heart:['قلب','M50 95C10 62 0 40 0 25C0 10 12 0 27 0C38 0 46 6 50 14C54 6 62 0 73 0C88 0 100 10 100 25C100 40 90 62 50 95Z'],
 smileyFace:['وجه مبتسم','M0 50A50 50 0 1 0 100 50A50 50 0 1 0 0 50ZM30 36A6 7 0 1 0 42 36A6 7 0 1 0 30 36ZM58 36A6 7 0 1 0 70 36A6 7 0 1 0 58 36ZM24 60Q50 92 76 60Q50 76 24 60Z','eo'],
 moon:['هلال','M62 0A50 50 0 1 0 62 100A36 50 0 1 1 62 0Z'],sun:['شمس',sun()],
 cloud:['سحابة','M25 78C4 78 -2 52 18 46C10 26 36 12 50 26C60 4 94 14 86 40C106 46 102 78 78 78Z'],
 lightningBolt:['برق','M40 0L8 55H38L22 100L92 38H58L78 0Z'],
 rightArrow:['سهم يمين','M0 30H60V0L100 50L60 100V70H0Z'],leftArrow:['سهم يسار','M100 30H40V0L0 50L40 100V70H100Z'],
 upArrow:['سهم لأعلى','M30 100V40H0L50 0L100 40H70V100Z'],downArrow:['سهم لأسفل','M30 0H70V60H100L50 100L0 60H30Z'],
 leftRightArrow:['سهم مزدوج','M0 50L30 15V35H70V15L100 50L70 85V65H30V85Z'],chevron:['شيفرون','M0 0H65L100 50L65 100H0L35 50Z'],homePlate:['بنتاغون سهمي','M0 0H70L100 50L70 100H0Z'],
 star4:['نجمة 4',star(4,.38)],star5:['نجمة 5',star(5,.4)],star6:['نجمة 6',star(6,.55)],star8:['نجمة 8',star(8,.7)],
 wedgeRectCallout:['فقاعة مستطيلة','M0 0H100V75H40L18 100L24 75H0Z'],wedgeRoundRectCallout:['فقاعة دائرية','M12 0H88Q100 0 100 12V63Q100 75 88 75H40L18 100L24 75H12Q0 75 0 63V12Q0 0 12 0Z'],
 flowChartDecision:['قرار','M50 0L100 50L50 100L0 50Z'],flowChartTerminator:['بداية/نهاية','M25 0H75A25 50 0 0 1 75 100H25A25 50 0 0 1 25 0Z'],flowChartData:['بيانات','M20 0H100L80 100H0Z'],
 flowChartProcess:['عملية','M0 0H100V100H0Z'],flowChartAlternateProcess:['عملية بديلة','M17 0H83A17 17 0 0 1 100 17V83A17 17 0 0 1 83 100H17A17 17 0 0 1 0 83V17A17 17 0 0 1 17 0Z'],
 flowChartPredefinedProcess:['عملية معرّفة','M0 0H100V100H0ZM12 0V100M88 0V100'],flowChartInternalStorage:['تخزين داخلي','M0 0H100V100H0ZM15 0V100M0 15H100'],
 flowChartDocument:['مستند','M0 0H100V85C85 75 65 75 50 85C35 95 15 95 0 85Z'],flowChartManualInput:['إدخال يدوي','M0 20L100 0V100H0Z'],
 flowChartPreparation:['تحضير','M20 0H80L100 50L80 100H20L0 50Z'],flowChartDelay:['تأخير','M0 0H50A50 50 0 0 1 50 100H0Z'],
 flowChartConnector:['رابط','M0 50A50 50 0 1 0 100 50A50 50 0 1 0 0 50Z'],flowChartOffpageConnector:['رابط خارج الصفحة','M0 0H100V70L50 100L0 70Z'],
 flowChartSummingJunction:['تجميع','M0 50A50 50 0 1 0 100 50A50 50 0 1 0 0 50ZM15 15L85 85M85 15L15 85'],
 flowChartMerge:['دمج','M0 0H100L50 100Z'],flowChartExtract:['استخراج','M50 0L100 100H0Z'],
 flowChartManualOperation:['عملية يدوية','M0 0H100L80 100H20Z'],flowChartMagneticDisk:['قاعدة بيانات','M0 15A50 15 0 0 1 100 15V85A50 15 0 0 1 0 85ZM0 15A50 15 0 0 0 100 15'],
 flowChartDisplay:['عرض','M0 50L18 0H78A22 50 0 0 1 78 100H18Z'],flowChartCollate:['ترتيب','M0 0H100L0 100H100Z'],
 flowChartSort:['فرز','M50 0L100 50L50 100L0 50ZM0 50H100'],flowChartOr:['أو','M0 50A50 50 0 1 0 100 50A50 50 0 1 0 0 50ZM50 0V100M0 50H100'],
 flowChartPunchedCard:['بطاقة مثقوبة','M20 0H100V100H0V20Z'],flowChartPunchedTape:['شريط مثقوب','M0 12C20 -4 30 -4 50 12S80 28 100 12V88C80 104 70 104 50 88S20 72 0 88Z'],
 flowChartMultidocument:['مستندات متعددة','M0 22H80V88C65 78 50 78 40 88C25 98 12 96 0 88ZM10 11H90V77M20 0H100V66'],
 flowChartOfflineStorage:['تخزين خارجي','M0 0H100L50 100ZM30 18H70'],flowChartStoredData:['بيانات مخزّنة','M12 0H100A12 50 0 0 0 100 100H12A12 50 0 0 1 12 0Z'],
 flowChartMagneticDrum:['تخزين مباشر','M15 0H85A15 50 0 0 1 85 100H15A15 50 0 0 1 15 0ZM85 0A15 50 0 0 0 85 100'],
 upDownArrow:['سهم رأسي مزدوج','M50 0L85 30H65V70H85L50 100L15 70H35V30H15Z'],quadArrow:['سهم رباعي','M50 0L70 22H58V42H78V30L100 50L78 70V58H58V78H70L50 100L30 78H42V58H22V70L0 50L22 30V42H42V22H30Z'],
 notchedRightArrow:['سهم مشقوق','M0 30H60V0L100 50L60 100V70H0L15 50Z'],bentArrow:['سهم منحني','M0 100V45A35 35 0 0 1 35 10H70V0L100 20L70 40V30H35A15 15 0 0 0 20 45V100Z'],
 uturnArrow:['سهم دوران','M0 100V40A40 40 0 0 1 40 0H50A40 40 0 0 1 90 40V50H100L80 80L60 50H70V40A20 20 0 0 0 50 20H40A20 20 0 0 0 20 40V100Z'],rightArrowCallout:['سهم بنص','M0 0H60V30H80V0L100 50L80 100V70H60V100H0Z'],
 mathPlus:['زائد','M35 5H65V35H95V65H65V95H35V65H5V35H35Z'],mathMinus:['ناقص','M5 38H95V62H5Z'],mathMultiply:['ضرب','M20 5L50 35L80 5L95 20L65 50L95 80L80 95L50 65L20 95L5 80L35 50L5 20Z'],
 mathEqual:['يساوي','M5 20H95V42H5ZM5 58H95V80H5Z'],mathDivide:['قسمة','M5 38H95V62H5ZM38 14A12 12 0 1 0 62 14A12 12 0 1 0 38 14ZM38 86A12 12 0 1 0 62 86A12 12 0 1 0 38 86Z']}})();
const SHC=[['أساسية',['rect','roundRect','ellipse','triangle','rtTriangle','diamond','parallelogram','trapezoid','pentagon','hexagon','octagon','plus','donut','frame','can','cube','foldedCorner']],
 ['رموز',['heart','smileyFace','moon','sun','cloud','lightningBolt']],['أسهم',['rightArrow','leftArrow','upArrow','downArrow','leftRightArrow','upDownArrow','quadArrow','notchedRightArrow','bentArrow','uturnArrow','rightArrowCallout','chevron','homePlate']],
 ['رياضيات',['mathPlus','mathMinus','mathMultiply','mathDivide','mathEqual']],
 ['نجوم',['star4','star5','star6','star8']],['فقاعات',['wedgeRectCallout','wedgeRoundRectCallout']],
 ['مخططات انسيابية',['flowChartTerminator','flowChartProcess','flowChartAlternateProcess','flowChartPredefinedProcess','flowChartDecision','flowChartData','flowChartDocument','flowChartManualInput','flowChartPreparation','flowChartDelay','flowChartInternalStorage','flowChartConnector','flowChartOffpageConnector','flowChartSummingJunction','flowChartMerge','flowChartExtract','flowChartManualOperation','flowChartMagneticDisk','flowChartStoredData','flowChartDisplay','flowChartCollate','flowChartSort','flowChartOr','flowChartPunchedCard','flowChartPunchedTape','flowChartMultidocument','flowChartOfflineStorage','flowChartMagneticDrum']]];
const shBg=(id,fill,ln,lw)=>{const q=SH[id];return 'url("data:image/svg+xml,'+encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" preserveAspectRatio="none"><path d="'+q[1]+'" fill="'+(fill||'none')+'"'+(q[2]?' fill-rule="evenodd"':'')+(ln?' stroke="'+ln+'" stroke-width="'+lw+'" vector-effect="non-scaling-stroke" stroke-linejoin="round"':'')+'/></svg>')+'") 0 0/100% 100% no-repeat'};
function shapeGallery(){const o=document.querySelector('.shg');if(o){o.remove();return}
 const w=E('div','shg'),box=E('div','shb','<div class="shh"><b>🔷 مكتبة الأشكال</b><button type="button" aria-label="إغلاق">✕</button></div>'),close=()=>w.remove();
 box.querySelector('button').onclick=close;w.onpointerdown=e=>{if(e.target===w)close()};
 for(const[t,ids]of SHC){const h=E('h5');h.textContent=t;const g=E('div','shgrid');
  for(const id of ids){const b=E('button','sht','<svg viewBox="-5 -5 110 110" aria-hidden="true"><path d="'+SH[id][1]+'"'+(SH[id][2]?' fill-rule="evenodd"':'')+' fill="currentColor" fill-opacity=".28" stroke="currentColor" stroke-width="5" stroke-linejoin="round"/></svg><span></span>');
   b.querySelector('span').textContent=SH[id][0];b.title=SH[id][0];b.onclick=()=>{close();addShape(id)};g.append(b)}
  box.append(h,g)}
 {const h=E('h5');h.textContent='موصلات وخطوط';const g=E('div','shgrid');
  for(const id in CN){const q=CN[id],
   b=E('button','sht','<svg viewBox="-5 -5 110 110" aria-hidden="true"><path d="'+CNP[q[1]]+'" fill="none" stroke="currentColor" stroke-width="6" stroke-linejoin="round"'+(q[4]?' stroke-dasharray="14 8"':'')+'/>'+(q[3]?'<path d="'+(q[1]=='straightConnector1'?'M97 50L80 40V60Z':'M97 15L80 5V25Z')+'" fill="currentColor"/>':'')+(q[2]?'<path d="M3 50L20 40V60Z" fill="currentColor"/>':'')+'</svg><span></span>');
   b.querySelector('span').textContent=q[0];b.title=q[0];b.onclick=()=>{close();addConn(id)};g.append(b)}
  box.append(h,g)}
 w.append(box);document.body.append(w)}

function addImg(){const i=document.createElement('input');i.type='file';i.accept='image/*';i.onchange=async()=>{const f=i.files[0];if(!f)return;
 try{const bm=await createImageBitmap(f);let ext={'image/png':'png','image/jpeg':'jpeg','image/gif':'gif'}[f.type],b;
  if(ext)b=new Uint8Array(await f.arrayBuffer());else{const c=document.createElement('canvas');c.width=bm.width;c.height=bm.height;c.getContext('2d').drawImage(bm,0,0);b=new Uint8Array(await(await new Promise(r=>c.toBlob(r,'image/png'))).arrayBuffer());ext='png'}
  const s=cur(),nm='reader_'+Date.now().toString(36)+'.'+ext,rid=nextRid(s.rel),r=s.rel.createElementNS(PR,'Relationship');
  S.put.set('ppt/media/'+nm,b);S.exts.add(ext);r.setAttribute('Id',rid);r.setAttribute('Type',R+'/image');r.setAttribute('Target','../media/'+nm);s.rel.documentElement.append(r);
  const q=Math.min(S.W*.6/bm.width,S.H*.6/bm.height),g=ctr(bm.width*q,bm.height*q);
  addEl(`<p:pic><p:nvPicPr><p:cNvPr id="${nid()}" name="Picture ${nid()}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${g.x}" y="${g.y}"/><a:ext cx="${g.w}" cy="${g.h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`)
 }catch(e){alert('الصورة دي مش مدعومة')}};i.click()}
// ---------- audio / video ----------
const MAUD={mp3:'audio/mpeg',wav:'audio/wav',m4a:'audio/mp4',wma:'audio/x-ms-wma',aac:'audio/aac',ogg:'audio/ogg',oga:'audio/ogg',opus:'audio/ogg',flac:'audio/flac',aif:'audio/aiff',aiff:'audio/aiff',amr:'audio/amr',mid:'audio/midi',midi:'audio/midi',webm:'audio/webm'},
 MVID={mp4:'video/mp4',m4v:'video/mp4',mov:'video/quicktime',webm:'video/webm',ogv:'video/ogg',mpg:'video/mpeg',mpeg:'video/mpeg','3gp':'video/3gpp'};
const mediaPoster=(kind,vurl)=>new Promise(res=>{
 const draw=(w,h,src)=>{const c=document.createElement('canvas');c.width=w;c.height=h;const x=c.getContext('2d');
  if(src)x.drawImage(src,0,0,w,h);else{x.fillStyle=kind=='video'?'#1f2937':'#1e40af';x.fillRect(0,0,w,h);x.fillStyle='#fff';x.font='bold '+Math.round(h*.45)+'px sans-serif';x.textAlign='center';x.textBaseline='middle';x.fillText(kind=='video'?'▶':'🔊',w/2,h/2)}
  c.toBlob(async b=>res({b:new Uint8Array(await b.arrayBuffer()),w,h}),'image/png')};
 if(kind!='video')return draw(256,256);
 const v=document.createElement('video');v.muted=true;v.preload='auto';v.playsInline=true;let done=0;
 const fin=ok=>{if(done)return;done=1;try{ok&&v.videoWidth?draw(Math.min(640,v.videoWidth),Math.round(Math.min(640,v.videoWidth)*v.videoHeight/v.videoWidth),v):draw(640,360)}catch(e){draw(640,360)}v.removeAttribute('src');v.load()};
 v.onloadeddata=()=>{try{v.currentTime=Math.min(.5,(v.duration||1)/2)}catch(e){fin(1)}};v.onseeked=()=>fin(1);v.onerror=()=>fin(0);setTimeout(()=>fin(v.readyState>=2),4000);v.src=vurl});
function addMedia(recordedFile){const i=document.createElement('input');i.type='file';i.accept='audio/*,video/*,.mp3,.m4a,.wav,.wma,.aac,.ogg,.oga,.opus,.flac,.aif,.aiff,.amr,.mid,.midi,.mp4,.m4v,.mov,.webm,.ogv,.mpg,.mpeg,.3gp';
 i.onchange=async()=>{const f=i.files[0];if(!f)return;
  const ext=((/\.([a-z0-9]+)$/i.exec(f.name)||[])[1]||'').toLowerCase(),kind=MVID[ext]?'video':MAUD[ext]?'audio':/^video\//.test(f.type)?'video':/^audio\//.test(f.type)?'audio':null;
  if(!kind)return alert('الصيغة دي مش مدعومة. استخدم MP3 / M4A / WAV / WMA / AAC / OGG / FLAC للصوت أو MP4 / WEBM للفيديو.');
  const xt=ext&&(MVID[ext]||MAUD[ext])?ext:(kind=='video'?'mp4':'mp3');
  if(f.size>150e6&&!await raConfirm('الملف كبير ('+Math.round(f.size/1e6)+' ميجا) وممكن يبطّأ التطبيق. تكمل؟'))return;
  X.busy(true,'جاري إدراج الملف…');
  try{const data=new Uint8Array(await f.arrayBuffer()),url=URL.createObjectURL(new Blob([data],{type:MVID[xt]||MAUD[xt]}));
   let po;try{po=await mediaPoster(kind,url)}finally{URL.revokeObjectURL(url)}
   const s=cur(),stamp=Date.now().toString(36),mn='reader_'+stamp+'.'+xt,pn='reader_'+stamp+'_poster.png',
    MS='http://schemas.microsoft.com/office/2007/relationships/media',
    rl=nextRid(s.rel);let n=+rl.replace(/\D/g,'');
   const add=(t,tg)=>{const id='rId'+(n++),r=s.rel.createElementNS(PR,'Relationship');r.setAttribute('Id',id);r.setAttribute('Type',t);r.setAttribute('Target',tg);s.rel.documentElement.append(r);return id};
   S.put.set('ppt/media/'+mn,data);S.put.set('ppt/media/'+pn,po.b);S.exts.add('png');S.exts.add(xt);S.mt=S.mt||{};S.mt[xt]=MVID[xt]||MAUD[xt];
   const rLink=add(R+'/'+kind,'../media/'+mn),rEmb=add(MS,'../media/'+mn),rPos=add(R+'/image','../media/'+pn);
   const title=f.name.replace(/\.[^.]+$/,'')||(kind=='video'?'فيديو':'صوت');
   const g=kind=='video'?(()=>{const q=Math.min(S.W*.6/po.w,S.H*.6/po.h);return ctr(po.w*q,po.h*q)})():ctr(Math.min(S.W,S.H)*.14,Math.min(S.W,S.H)*.14);
   const item={slide:S.cur+1,kind,path:'ppt/media/'+mn,target:'../media/'+mn,external:false,rect:null,hid:false,title,file:mn,ext:xt,mime:MVID[xt]||MAUD[xt],size:data.length,data};
   S.media=S.media||[];S.media.push(item);X.media&&X.media.add&&X.media.add(item);
   const id=nid();
   addEl(`<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${kind=='video'?'Video':'Audio'} ${id}" descr="${esc(title)}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:${kind}File r:link="${rLink}"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB1230}"><p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="${rEmb}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="${rPos}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${g.x}" y="${g.y}"/><a:ext cx="${g.w}" cy="${g.h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
   X.toast(kind=='video'?'اتضاف الفيديو للشريحة':'اتضاف الصوت للشريحة')
  }catch(e){console.warn(e);alert('تعذر إدراج الملف')}finally{X.busy(false)}};
 if(recordedFile){try{const dt=new DataTransfer();dt.items.add(recordedFile);i.files=dt.files;i.onchange()}catch(e){alert('تعذر تجهيز التسجيل للإدراج')}}else i.click()}
async function recordSlideAudio(){
 if(!navigator.mediaDevices?.getUserMedia||!window.MediaRecorder)return alert('تسجيل الصوت غير مدعوم على هذا الجهاز أو المتصفح.');
 let stream,rec,chunks=[],stopped=false;
 const box=E('div','pe-rec'),panel=E('div','pe-rec-card'),title=E('b','','🎙 تسجيل صوت للشريحة '+(S.cur+1)),status=E('p','','اضغط بدء التسجيل، ثم إيقاف عند الانتهاء.'),start=E('button','','● بدء التسجيل'),stop=E('button','','■ إيقاف وحفظ'),cancel=E('button','','إلغاء');
 box.append(panel);panel.append(title,status,start,stop,cancel);Object.assign(box.style,{position:'fixed',inset:'0',zIndex:99999,background:'#0009',display:'grid',placeItems:'center',padding:'18px'});Object.assign(panel.style,{background:'white',color:'#111',borderRadius:'18px',padding:'20px',width:'min(360px,100%)',display:'grid',gap:'12px',boxShadow:'0 12px 40px #0005'});[start,stop,cancel].forEach(b=>{b.style.cssText='padding:13px;border:1px solid #ccd3df;border-radius:10px;background:#f4f6fb;font:inherit;color:#111'});stop.disabled=true;stop.style.opacity='.5';document.body.append(box);
 const cleanup=()=>{if(stream)stream.getTracks().forEach(t=>t.stop());box.remove()};cancel.onclick=()=>{stopped=true;try{if(rec&&rec.state!=='inactive')rec.stop()}catch{}cleanup()};
 start.onclick=async()=>{start.disabled=true;try{if(window.raEnsureMic)await window.raEnsureMic();stream=await navigator.mediaDevices.getUserMedia({audio:true});const types=['audio/mp4','audio/webm;codecs=opus','audio/webm'];const mime=types.find(t=>MediaRecorder.isTypeSupported(t));rec=new MediaRecorder(stream,mime?{mimeType:mime}:undefined);rec.ondataavailable=e=>{if(e.data&&e.data.size)chunks.push(e.data)};rec.onerror=()=>{status.textContent='حدث خطأ أثناء التسجيل.';cleanup()};rec.onstop=async()=>{if(stopped)return;try{status.textContent='جاري تجهيز الصوت…';const raw=new Blob(chunks,{type:rec.mimeType||'audio/webm'}),ctx=new (window.AudioContext||window.webkitAudioContext)(),buf=await ctx.decodeAudioData(await raw.arrayBuffer()),samples=buf.getChannelData(0),data=new ArrayBuffer(44+samples.length*2),v=new DataView(data),rate=buf.sampleRate;const put=(o,t)=>{for(let j=0;j<t.length;j++)v.setUint8(o+j,t.charCodeAt(j))};put(0,'RIFF');v.setUint32(4,36+samples.length*2,true);put(8,'WAVE');put(12,'fmt ');v.setUint32(16,16,true);v.setUint16(20,1,true);v.setUint16(22,1,true);v.setUint32(24,rate,true);v.setUint32(28,rate*2,true);v.setUint16(32,2,true);v.setUint16(34,16,true);put(36,'data');v.setUint32(40,samples.length*2,true);for(let j=0;j<samples.length;j++){const q=Math.max(-1,Math.min(1,samples[j]));v.setInt16(44+j*2,q<0?q*32768:q*32767,true)}await ctx.close();const file=new File([data],'slide-'+(S.cur+1)+'-voice.wav',{type:'audio/wav',lastModified:Date.now()});cleanup();addMedia(file)}catch(e){console.warn('audio decode',e);cleanup();alert('تعذر تجهيز التسجيل الصوتي. حاول مرة أخرى.')}};rec.start(500);status.textContent='جاري التسجيل… تحدث الآن.';stop.disabled=false;stop.style.opacity='1'}catch(e){console.warn(e);status.textContent=e&&e.name==='NotFoundError'?'مفيش ميكروفون متاح على الجهاز.':e&&e.name==='NotReadableError'?'الميكروفون مشغول بتطبيق تاني. اقفله وجرّب تاني.':'لم نتمكن من تشغيل الميكروفون. راجع إذن الميكروفون من إعدادات التطبيق.';start.disabled=false}};
 stop.onclick=()=>{if(rec&&rec.state!=='inactive'){status.textContent='جاري تجهيز التسجيل…';stop.disabled=true;rec.stop()}};
}
async function addTable(){const v=await raAsk('عدد الصفوف × الأعمدة','3x3');if(!v)return;const m=/(\d+)\D+(\d+)/.exec(v);if(!m)return;
 const R_=Math.min(12,Math.max(1,+m[1])),C=Math.min(8,Math.max(1,+m[2])),W=S.W*.8,cw=Math.round(W/C),rh=Math.round(Math.min(S.H*.08,PT*40)),g=ctr(cw*C,rh*R_);
 const ln=t=>`<a:${t} w="12700"><a:solidFill><a:srgbClr val="888888"/></a:solidFill></a:${t}>`;let rows='';
 for(let r=0;r<R_;r++){rows+=`<a:tr h="${rh}">`;for(let c=0;c<C;c++)rows+=`<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" sz="1400"/></a:p></a:txBody><a:tcPr>${ln('lnL')}${ln('lnR')}${ln('lnT')}${ln('lnB')}${r?'':'<a:solidFill><a:srgbClr val="DCE6F7"/></a:solidFill>'}</a:tcPr></a:tc>`;rows+='</a:tr>'}
 addEl(`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${nid()}" name="Table ${nid()}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${g.x}" y="${g.y}"/><a:ext cx="${g.w}" cy="${g.h}"/></p:xfrm><a:graphic><a:graphicData uri="${A}/table"><a:tbl><a:tblPr firstRow="1"/><a:tblGrid>${`<a:gridCol w="${cw}"/>`.repeat(C)}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`)}
async function addChart(){
 const raw=await raAsk('اكتب القيم مفصولة بفواصل (مثال: 25,40,65,50)','25,40,65,50');if(!raw)return;
 const vals=raw.split(/[,،\s]+/).map(Number).filter(v=>Number.isFinite(v)&&v>=0).slice(0,8);if(vals.length<2)return alert('اكتب قيمتين على الأقل');
 const max=Math.max(...vals,1),base=Math.round(S.H*.78),top=Math.round(S.H*.28),left=Math.round(S.W*.12),right=Math.round(S.W*.88),usable=base-top,gap=Math.round((right-left)/(vals.length*3)),bw=Math.round((right-left-gap*(vals.length+1))/vals.length);
 commit();snap();const parts=[];parts.push(spX({x:left,y:top,w:Math.max(1,right-left),h:Math.max(1,Math.round(S.H*.002))},'rect','64748b'));
 vals.forEach((v,i)=>{const h=Math.max(1,Math.round(usable*v/max)),x=left+gap+(bw+gap)*i,y=base-h;parts.push(spX({x,y,w:bw,h},'rect',['2563eb','0d9488','d97706','7c3aed','db2777','0891b2','65a30d','ea580c'][i%8]));parts.push(spX({x:x- Math.round(gap/2),y:base+Math.round(S.H*.025),w:bw+gap,h:Math.round(S.H*.07)},'rect',null,'فئة '+(i+1)));parts.push(spX({x:x- Math.round(gap/2),y:Math.max(top,y-Math.round(S.H*.065)),w:bw+gap,h:Math.round(S.H*.06)},'rect',null,String(v)));});
 parts.forEach(x=>tree(cur()).append(frag(x,cur().doc)));touch();S.setTab('ins');draw();thumbs();X.toast('تم إدراج مخطط أعمدة قابل لتعديل عناصره');
}
function delShape(){if(!S.sel)return;snap();const id=all(S.sel,P,'cNvPr')[0],tm=all(cur().doc,P,'timing')[0];
 if(id&&tm&&all(tm,P,'spTgt').some(x=>x.getAttribute('spid')==id.getAttribute('id')))animDrop(cur(),id.getAttribute('id'));S.ed=null;S.sel.remove();S.sel=null;touch();draw()}

// ---------- slides ----------
const newPath=()=>{let n=1;while(S.slides.some(x=>x.path=='ppt/slides/slide'+n+'.xml')||S.z.m['ppt/slides/slide'+n+'.xml'])n++;return 'ppt/slides/slide'+n+'.xml'};
const SKEL=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NSD}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
function addSlide(dup){commit();snap();const a=cur(),n={path:newPath(),isNew:1,t:1,lay:a.lay,mas:a.mas,lt:a.lt,lp:a.lp,lrel:a.lrel,mp:a.mp,mrel:a.mrel};
 if(dup){n.doc=ps(ser(a.doc));n.rel=ps(ser(a.rel));for(const r of all(n.rel,PR,'Relationship'))if(/notesSlide$/.test(r.getAttribute('Type')))r.remove()}
 else{n.doc=ps(SKEL);n.rel=ps(`<Relationships xmlns="${PR}">${a.lt?`<Relationship Id="rId1" Type="${R}/slideLayout" Target="${esc(a.lt)}"/>`:''}</Relationships>`)}
 S.slides.splice(S.cur+1,0,n);S.st=1;S.cur++;S.sel=null;X.O().dirty=1;if(!dup){const e=frag(spX({x:Math.round(S.W*.1),y:Math.round(S.H*.06),w:Math.round(S.W*.8),h:Math.round(S.H*.16)},'rect',null,'عنوان الشريحة'),n.doc);tree(n).append(e)}
 draw();thumbs()}
async function delSlide(){if(S.slides.length<2)return alert('لازم يفضل شريحة واحدة على الأقل.');if(!await raConfirm('حذف الشريحة '+(S.cur+1)+'؟'))return;commit();snap();
 const s=S.slides.splice(S.cur,1)[0];if(!s.isNew){S.gone.push(s);S.del.add(s.path);S.del.add(relPath(s.path))}S.st=1;S.cur=Math.min(S.cur,S.slides.length-1);S.sel=null;X.O().dirty=1;draw();thumbs()}
// ---------- slide templates ----------
const TPLC={b:'1e40af',g:'6b7280',w:'FFFFFF'};
const TPLT=(g,paras,o={})=>{const W=S.W,H=S.H,id=nid(),x=Math.round(g[0]*W),y=Math.round(g[1]*H),w=Math.round(g[2]*W),h=Math.round(g[3]*H);
 const rp=(sz,extra)=>`<a:rPr lang="ar-EG" sz="${sz}"${o.b?' b="1"':''}${o.i?' i="1"':''}${extra||''}>${o.c?`<a:solidFill><a:srgbClr val="${o.c}"/></a:solidFill>`:''}</a:rPr>`;
 const ps_=(paras&&paras.length?paras:[null]).map(t=>{if(t==null)return'<a:p><a:endParaRPr lang="en-US"/></a:p>';
  const pp=`<a:pPr${o.bu?' marL="457200" indent="-457200"':''} algn="${o.al||'r'}" rtl="1">${o.bu?`<a:spcBef><a:spcPts val="900"/></a:spcBef>${o.bu=='n'?'<a:buAutoNum type="arabicPeriod"/>':'<a:buChar char="•"/>'}`:''}</a:pPr>`;
  return `<a:p>${pp}<a:r>${rp(o.sz||2400)}<a:t>${esc(t)}</a:t></a:r></a:p>`}).join('');
 return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${o.n||'Text'} ${id}"/><p:cNvSpPr${paras&&paras.length&&!o.fill?' txBox="1"':''}/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="${o.geom||'rect'}"><a:avLst/></a:prstGeom>${o.fill?`<a:solidFill><a:srgbClr val="${o.fill}"/></a:solidFill>`:'<a:noFill/>'}</p:spPr><p:txBody><a:bodyPr wrap="square" anchor="${o.an||'t'}"><a:normAutofit/></a:bodyPr><a:lstStyle/>${ps_}</p:txBody></p:sp>`};
const TPLH=(t)=>[[.06,.05,.88,.14],[t||'عنوان الشريحة'],{sz:3600,b:1,c:TPLC.b,an:'ctr',n:'Title'}],
 TPLU=()=>[[.06,.2,.2,.008],null,{fill:TPLC.b,n:'Line'}];
const TPL={
 blank:{l:'⬜ فارغة',s:[]},
 title:{l:'🏷 غلاف',bg:'1e3a8a',s:[[[.1,.28,.8,.2],['عنوان العرض'],{sz:5000,b:1,c:TPLC.w,al:'ctr',an:'b',n:'Title'}],[[.42,.5,.16,.012],null,{fill:'fbbf24',n:'Line'}],[[.1,.54,.8,.14],['العنوان الفرعي — الاسم / التاريخ'],{sz:2400,c:'c7d2fe',al:'ctr',n:'Subtitle'}]]},
 bullets:{l:'📋 عنوان + نقاط',s:[TPLH(),TPLU(),[[.06,.25,.88,.65],['النقطة الأولى','النقطة الثانية','النقطة الثالثة'],{sz:2800,bu:'b',n:'Content'}]]},
 agenda:{l:'🔢 قائمة مرقّمة',s:[TPLH('المحتويات'),TPLU(),[[.06,.25,.88,.65],['المقدمة','الموضوع الأساسي','النتائج','الخلاصة'],{sz:2800,bu:'n',n:'Content'}]]},
 two:{l:'◫ عمودين',s:[TPLH(),TPLU(),[[.52,.25,.42,.65],['نص العمود الأول','نقطة','نقطة'],{sz:2400,bu:'b',n:'Left'}],[[.06,.25,.42,.65],['نص العمود الثاني','نقطة','نقطة'],{sz:2400,bu:'b',n:'Right'}]]},
 compare:{l:'⚖ مقارنة',s:[TPLH('مقارنة'),TPLU(),[[.52,.25,.42,.1],['الأول'],{sz:2800,b:1,c:TPLC.w,fill:'1e40af',al:'ctr',an:'ctr',n:'Head 1'}],[[.06,.25,.42,.1],['الثاني'],{sz:2800,b:1,c:TPLC.w,fill:'0f766e',al:'ctr',an:'ctr',n:'Head 2'}],[[.52,.37,.42,.53],['نقطة','نقطة','نقطة'],{sz:2400,bu:'b',n:'Body 1'}],[[.06,.37,.42,.53],['نقطة','نقطة','نقطة'],{sz:2400,bu:'b',n:'Body 2'}]]},
 imgtext:{l:'🖼 صورة + نص',s:[TPLH(),TPLU(),[[.06,.25,.42,.65],['ضع الصورة هنا (إدراج ← صورة)'],{sz:2000,c:TPLC.g,fill:'e5e7eb',al:'ctr',an:'ctr',n:'Picture'}],[[.52,.25,.42,.65],['النص المرافق للصورة','نقطة','نقطة'],{sz:2400,bu:'b',n:'Content'}]]},
 quote:{l:'❝ اقتباس',bg:'f1f5f9',s:[[[.08,.06,.2,.26],['❝'],{sz:9600,c:'93c5fd',al:'l',n:'Mark'}],[[.12,.3,.76,.36],['اكتب الاقتباس أو الفكرة المهمة هنا'],{sz:3600,i:1,c:'1e293b',al:'ctr',an:'ctr',n:'Quote'}],[[.12,.7,.76,.1],['— المصدر'],{sz:2000,c:TPLC.g,al:'ctr',n:'Author'}]]},
 section:{l:'➖ فاصل قسم',bg:'0f766e',s:[[[.1,.35,.8,.22],['عنوان القسم'],{sz:4800,b:1,c:TPLC.w,al:'ctr',an:'ctr',n:'Title'}],[[.1,.58,.8,.12],['وصف مختصر'],{sz:2400,c:'ccfbf1',al:'ctr',n:'Subtitle'}]]},
 timeline:{l:'⏳ خط زمني',s:[TPLH('الخط الزمني'),TPLU(),[[.08,.5,.84,.008],null,{fill:TPLC.b,n:'Line'}],...[.85,.62,.38,.15].flatMap((c,i)=>[[[c-.035,.438,.07,.124],[String(i+1)],{sz:2400,b:1,c:TPLC.w,fill:TPLC.b,geom:'ellipse',al:'ctr',an:'ctr',n:'Dot '+(i+1)}],[[c-.1,.6,.2,.28],['المرحلة '+(i+1)],{sz:2000,al:'ctr',n:'Label '+(i+1)}]])]},
 cards3:{l:'🃏 ثلاث بطاقات',s:[TPLH(),TPLU(),...[[.66,'1e40af'],[.36,'0f766e'],[.06,'b45309']].flatMap(([x,c],i)=>[[[x,.25,.28,.1],['العنوان '+(i+1)],{sz:2400,b:1,c:TPLC.w,fill:c,al:'ctr',an:'ctr',n:'Card head '+(i+1)}],[[x,.35,.28,.5],['النص هنا'],{sz:2000,fill:'f1f5f9',c:'1e293b',al:'ctr',an:'ctr',n:'Card body '+(i+1)}]])]},
 stats:{l:'📊 أرقام وإحصائيات',s:[TPLH('أرقام مهمة'),TPLU(),...[.66,.36,.06].flatMap((x,i)=>[[[x,.28,.28,.25],['٠٠٪'],{sz:6000,b:1,c:TPLC.b,al:'ctr',an:'ctr',n:'Number '+(i+1)}],[[x,.55,.28,.15],['وصف الرقم'],{sz:2400,c:TPLC.g,al:'ctr',n:'Caption '+(i+1)}]])]},
 callout:{l:'💡 ملاحظة مهمة',s:[TPLH('ملاحظة مهمة'),TPLU(),[[.1,.27,.8,.5],['اكتب النقطة الأهم أو التعريف هنا'],{sz:3200,c:'1e293b',fill:'dbeafe',al:'ctr',an:'ctr',n:'Callout'}],[[.9,.27,.012,.5],null,{fill:TPLC.b,n:'Bar'}]]},
 twoimg:{l:'🖼🖼 صورتين',s:[TPLH(),TPLU(),...[.52,.06].flatMap((x,i)=>[[[x,.25,.42,.5],['ضع الصورة '+(i+1)+' هنا'],{sz:2000,c:TPLC.g,fill:'e5e7eb',al:'ctr',an:'ctr',n:'Picture '+(i+1)}],[[x,.77,.42,.1],['التعليق'],{sz:2000,c:TPLC.g,al:'ctr',n:'Caption '+(i+1)}]])]},
 medical:{l:'🩺 طبي / علمي',bg:'f0fdfa',s:[TPLH('عنوان الموضوع الطبي'),TPLU(),[[.66,.27,.28,.48],['التعريف','الأسباب','الأعراض'],{sz:2300,bu:'b',fill:'ccfbf1',c:'134e4a',n:'Key points'}],[[.06,.27,.54,.48],['اكتب شرحًا علميًا مختصرًا هنا'],{sz:2600,c:'1e293b',fill:'ffffff',n:'Explanation'}],[[.06,.8,.88,.1],['المصدر / المرجع العلمي'],{sz:1700,c:'64748b',n:'Reference'}]]},
 process:{l:'🔄 خطوات / عملية',s:[TPLH('مراحل العملية'),TPLU(),...[.76,.53,.30,.07].flatMap((x,i)=>[[[x,.34,.17,.22],[String(i+1)],{sz:3200,b:1,c:TPLC.w,fill:['1e40af','0f766e','b45309','7c3aed'][i],geom:'roundRect',al:'ctr',an:'ctr',n:'Step '+(i+1)}],[[x-.025,.61,.22,.16],[['الخطوة الأولى','الخطوة الثانية','الخطوة الثالثة','الخطوة الرابعة'][i]],{sz:1900,al:'ctr',n:'Step label '+(i+1)}]])]},
 swot:{l:'⚖ تحليل SWOT',s:[TPLH('تحليل SWOT'),TPLU(),...[[.52,.27,'dcfce7','166534','نقاط القوة'],[.06,.27,'fee2e2','991b1b','نقاط الضعف'],[.52,.56,'dbeafe','1e40af','الفرص'],[.06,.56,'fef3c7','92400e','التهديدات']].flatMap(([x,y,bg,c,t],i)=>[[[x,y,.42,.1],[t],{sz:2400,b:1,c:TPLC.w,fill:c,al:'ctr',an:'ctr',n:'SWOT head '+i}],[[x,y+.11,.42,.18],['أضف النقاط هنا','نقطة إضافية'],{sz:1800,c:'1f2937',fill:bg,n:'SWOT body '+i}]])]},
 gallery:{l:'🖼 معرض صور',s:[TPLH('معرض الصور'),TPLU(),...[.66,.36,.06].flatMap((x,i)=>[[[x,.27,.28,.42],['أدرج الصورة'],{sz:2100,c:TPLC.g,fill:'e5e7eb',al:'ctr',an:'ctr',n:'Image placeholder '+i}],[[x,.72,.28,.1],['وصف الصورة'],{sz:1800,c:TPLC.g,al:'ctr',n:'Image caption '+i}]])]},
 lesson:{l:'🎓 درس تعليمي',s:[TPLH('عنوان الدرس'),TPLU(),[[.55,.26,.39,.58],['أهداف التعلم','الهدف الأول','الهدف الثاني','الهدف الثالث'],{sz:2300,bu:'b',fill:'dbeafe',c:'1e3a8a',n:'Objectives'}],[[.06,.26,.43,.58],['المفهوم الأساسي','اكتب شرح الدرس هنا','أضف مثالًا توضيحيًا'],{sz:2200,bu:'b',n:'Lesson content'}]]},
 conclusion:{l:'🏁 خلاصة ونتائج',bg:'eff6ff',s:[TPLH('الخلاصة والنتائج'),TPLU(),[[.08,.27,.84,.42],['النتيجة الأولى','النتيجة الثانية','التوصية الأساسية'],{sz:2700,bu:'b',c:'1e3a8a',fill:'ffffff',n:'Conclusions'}],[[.08,.75,.84,.1],['شكرًا لكم'],{sz:2400,b:1,c:'1e40af',al:'ctr',n:'Thanks'}]]},
 team:{l:'👥 فريق العمل',s:[TPLH('فريق العمل'),TPLU(),...[.66,.36,.06].flatMap((x,i)=>[[[x,.28,.28,.25],['الاسم '+(i+1)],{sz:2600,b:1,c:TPLC.w,fill:['1e40af','0f766e','b45309'][i],al:'ctr',an:'ctr',n:'Member '+i}],[[x,.56,.28,.15],['الدور / التخصص'],{sz:1900,c:TPLC.g,al:'ctr',n:'Role '+i}]])]} ,
 thanks:{l:'🙏 شكرًا',bg:'1e3a8a',s:[[[.1,.3,.8,.22],['شكرًا لكم'],{sz:6000,b:1,c:TPLC.w,al:'ctr',an:'ctr',n:'Title'}],[[.42,.54,.16,.012],null,{fill:'fbbf24',n:'Line'}],[[.1,.58,.8,.12],['أي أسئلة؟'],{sz:2800,c:'c7d2fe',al:'ctr',n:'Subtitle'}]]}
};
function addTpl(k){const t=TPL[k];if(!t)return;commit();snap();const a=cur(),n={path:newPath(),isNew:1,t:1,lay:a.lay,mas:a.mas,lt:a.lt,lp:a.lp,lrel:a.lrel,mp:a.mp,mrel:a.mrel};
 n.doc=ps(SKEL);n.rel=ps(`<Relationships xmlns="${PR}">${a.lt?`<Relationship Id="rId1" Type="${R}/slideLayout" Target="${esc(a.lt)}"/>`:''}</Relationships>`);
 S.slides.splice(S.cur+1,0,n);S.st=1;S.cur++;S.sel=null;X.O().dirty=1;
 if(t.bg){const cs=all(n.doc,P,'cSld')[0];cs.insertBefore(frag(`<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${t.bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`,n.doc),cs.firstChild)}
 for(const d of t.s)tree(n).append(frag(TPLT(d[0],d[1],d[2]),n.doc));
 draw();thumbs()}
// ---------- ready-made presentations (shown when pressing "PowerPoint جديد") ----------
const DK_MAP=t=>({'1e40af':t.p,'1e3a8a':t.d,'0f766e':t.s,'fbbf24':t.a,'b45309':t.t,'7c3aed':t.q||t.d,'dbeafe':t.l2,'eff6ff':t.l,'f1f5f9':t.l,'f0fdfa':t.l,'ccfbf1':t.l2,'c7d2fe':t.i,'93c5fd':t.i2,'134e4a':t.d});
const DK_DARK=pn=>({'f1f5f9':pn,'dbeafe':pn,'eff6ff':pn,'f0fdfa':pn,'ccfbf1':pn,'1e293b':'f1f5f9','6b7280':'94a3b8','64748b':'94a3b8','e5e7eb':'334155'});
const DECKS={
 blank:{n:'فارغ',ic:'⬜',d:'شريحة بيضاء فاضية تبدأ منها بنفسك',s:['blank'],t:{p:'1e40af',d:'e5e7eb',s:'cbd5e1',a:'94a3b8'}},
 biz:{n:'أعمال احترافي',ic:'💼',d:'عروض الشركات والاجتماعات',s:['title','agenda','bullets','two','stats','thanks'],t:{p:'1e40af',d:'1e3a8a',s:'0f766e',a:'fbbf24',t:'b45309',l:'f1f5f9',l2:'dbeafe',i:'c7d2fe',i2:'93c5fd'}},
 nature:{n:'أخضر طبيعي',ic:'🌿',d:'البيئة والصحة والأنشطة',s:['title','bullets','imgtext','cards3','conclusion','thanks'],t:{p:'166534',d:'14532d',s:'0f766e',a:'a3e635',t:'b45309',l:'f0fdf4',l2:'dcfce7',i:'bbf7d0',i2:'86efac'}},
 creative:{n:'بنفسجي إبداعي',ic:'🎨',d:'التسويق والأفكار الإبداعية',s:['title','section','cards3','gallery','quote','thanks'],t:{p:'6d28d9',d:'4c1d95',s:'db2777',a:'fbbf24',t:'ea580c',l:'f5f3ff',l2:'ede9fe',i:'ddd6fe',i2:'c4b5fd'}},
 sales:{n:'برتقالي دافئ',ic:'🧡',d:'المبيعات والخطط والمراحل',s:['title','agenda','timeline','process','stats','thanks'],t:{p:'c2410c',d:'7c2d12',s:'b45309',a:'fbbf24',t:'0f766e',l:'fff7ed',l2:'ffedd5',i:'fed7aa',i2:'fdba74'}},
 medical:{n:'طبي / علمي',ic:'🩺',d:'الطب والعلوم والأبحاث',s:['title','medical','bullets','compare','conclusion','thanks'],t:{p:'0f766e',d:'134e4a',s:'0891b2',a:'5eead4',t:'1e40af',l:'f0fdfa',l2:'ccfbf1',i:'ccfbf1',i2:'99f6e4'}},
 edu:{n:'تعليمي / درس',ic:'🎓',d:'الشرح والمحاضرات والدروس',s:['title','agenda','lesson','imgtext','conclusion','thanks'],t:{p:'4338ca',d:'312e81',s:'0891b2',a:'fbbf24',t:'be185d',l:'eef2ff',l2:'e0e7ff',i:'c7d2fe',i2:'a5b4fc'}},
 report:{n:'تقرير / بحث',ic:'📑',d:'التقارير والأبحاث والمقارنات',s:['title','agenda','bullets','compare','stats','conclusion','thanks'],t:{p:'1e293b',d:'0f172a',s:'475569',a:'38bdf8',t:'b45309',l:'f1f5f9',l2:'e2e8f0',i:'cbd5e1',i2:'94a3b8'}},
 dark:{n:'داكن عصري',ic:'🌙',d:'مظهر داكن أنيق للعروض الحديثة',dark:1,s:['title','bullets','cards3','stats','quote','thanks'],t:{q:'2563eb',p:'3b82f6',d:'0b1220',s:'14b8a6',a:'fbbf24',t:'f97316',l:'1e293b',l2:'1e293b',i:'cbd5e1',i2:'3b82f6'}},
 project:{n:'مشروع / خطة',ic:'🚀',d:'عرض مشروع وتحليل وفريق عمل',s:['title','swot','timeline','team','cards3','thanks'],t:{p:'be123c',d:'881337',s:'0f766e',a:'fbbf24',t:'1e40af',l:'fff1f2',l2:'ffe4e6',i:'fecdd3',i2:'fda4af'}},
 ocean:{n:'محيطي',ic:'🌊',d:'عروض هادئة بألوان البحر',s:['title','section','bullets','imgtext','stats','thanks'],t:{p:'0891b2',d:'164e63',s:'0369a1',a:'fde047',t:'f97316',l:'ecfeff',l2:'cffafe',i:'a5f3fc',i2:'67e8f9'}},
 rose:{n:'وردي أنيق',ic:'🌸',d:'الموضة والجمال والمناسبات',s:['title','quote','gallery','twoimg','cards3','thanks'],t:{p:'be185d',d:'831843',s:'9d174d',a:'fbbf24',t:'f97316',l:'fdf2f8',l2:'fce7f3',i:'fbcfe8',i2:'f9a8d4'}},
 gold:{n:'ذهبي فاخر',ic:'👑',d:'مظهر داكن بلمسات ذهبية',dark:1,bg:'1c1917',pn:'292524',s:['title','agenda','bullets','stats','quote','thanks'],t:{q:'b45309',p:'f59e0b',d:'0c0a09',s:'d97706',a:'fde68a',t:'fbbf24',l:'292524',l2:'292524',i:'e7e5e4',i2:'f59e0b'}},
 crimson:{n:'أحمر قوي',ic:'🔴',d:'العروض الترويجية والحملات',s:['title','section','stats','cards3','callout','thanks'],t:{p:'dc2626',d:'7f1d1d',s:'1e293b',a:'fbbf24',t:'ea580c',l:'fef2f2',l2:'fee2e2',i:'fecaca',i2:'fca5a5'}},
 coffee:{n:'بني قهوة',ic:'☕',d:'ألوان دافئة وتراثية',s:['title','bullets','imgtext','timeline','conclusion','thanks'],t:{p:'92400e',d:'451a03',s:'78350f',a:'fcd34d',t:'0f766e',l:'fffbeb',l2:'fef3c7',i:'fde68a',i2:'fcd34d'}},
 navy:{n:'كحلي رسمي',ic:'🏛',d:'الجهات الحكومية والرسمية',s:['title','agenda','bullets','compare','team','conclusion','thanks'],t:{p:'1e3a8a',d:'172554',s:'475569',a:'fbbf24',t:'b45309',l:'f8fafc',l2:'e2e8f0',i:'bfdbfe',i2:'93c5fd'}},
 startup:{n:'شركة ناشئة',ic:'💡',d:'عرض فكرة ومشروع للمستثمرين',s:['title','callout','imgtext','stats','team','timeline','thanks'],t:{p:'ea580c',d:'1c1917',s:'0ea5e9',a:'fde047',t:'7c3aed',l:'fafaf9',l2:'f5f5f4',i:'e7e5e4',i2:'d6d3d1'}},
 kids:{n:'أطفال ومرح',ic:'🧸',d:'ألوان مبهجة للأطفال',s:['title','imgtext','cards3','gallery','quote','thanks'],t:{p:'f97316',d:'1d4ed8',s:'16a34a',a:'facc15',t:'ec4899',l:'fffbeb',l2:'fef3c7',i:'fed7aa',i2:'fdba74'}},
 tech:{n:'تقني / برمجة',ic:'💻',d:'التقنية والبرمجة والأنظمة',dark:1,bg:'020617',pn:'0f172a',s:['title','process','two','stats','timeline','thanks'],t:{q:'2563eb',p:'0891b2',d:'020617',s:'7c3aed',a:'10b981',t:'fb923c',l:'0f172a',l2:'0f172a',i:'cbd5e1',i2:'22d3ee'}},
 finance:{n:'مالي / ميزانية',ic:'💰',d:'التقارير المالية والميزانيات',s:['title','stats','compare','timeline','swot','conclusion','thanks'],t:{p:'0369a1',d:'082f49',s:'15803d',a:'facc15',t:'b91c1c',l:'f0f9ff',l2:'e0f2fe',i:'bae6fd',i2:'7dd3fc'}},
 lime:{n:'ليموني منعش',ic:'🍋',d:'الصحة والرياضة والعافية',s:['title','bullets','cards3','process','conclusion','thanks'],t:{p:'4d7c0f',d:'365314',s:'0f766e',a:'facc15',t:'c2410c',l:'f7fee7',l2:'ecfccb',i:'d9f99d',i2:'bef264'}},
 slate:{n:'رمادي بسيط',ic:'🩶',d:'تصميم مينيمال هادئ',s:['title','bullets','two','imgtext','quote','thanks'],t:{p:'334155',d:'1e293b',s:'64748b',a:'f59e0b',t:'0ea5e9',l:'f8fafc',l2:'f1f5f9',i:'cbd5e1',i2:'94a3b8'}},
 neon:{n:'نيون داكن',ic:'🔮',d:'مظهر داكن بألوان نيون',dark:1,bg:'13061f',pn:'24103a',s:['title','section','cards3','stats','quote','thanks'],t:{q:'7c3aed',p:'c026d3',d:'13061f',s:'06b6d4',a:'facc15',t:'f43f5e',l:'24103a',l2:'24103a',i:'e9d5ff',i2:'d946ef'}},
 bw:{n:'أبيض وأسود',ic:'⚫',d:'كلاسيكي رسمي بسيط',s:['title','agenda','bullets','two','compare','thanks'],t:{p:'111827',d:'000000',s:'4b5563',a:'ef4444',t:'6b7280',l:'f9fafb',l2:'f3f4f6',i:'d1d5db',i2:'9ca3af'}},
 event:{n:'مناسبات وحفلات',ic:'🎉',d:'الدعوات والفعاليات والحفلات',s:['title','section','gallery','timeline','quote','thanks'],t:{p:'c026d3',d:'701a75',s:'f59e0b',a:'fde047',t:'0ea5e9',l:'fdf4ff',l2:'fae8ff',i:'f5d0fe',i2:'f0abfc'}}
};
function deckSlide(k,a,D){const t=TPL[k],n={path:newPath(),isNew:1,t:1,lay:a.lay,mas:a.mas,lt:a.lt,lp:a.lp,lrel:a.lrel,mp:a.mp,mrel:a.mrel};
 n.doc=ps(SKEL);n.rel=ps(`<Relationships xmlns="${PR}">${a.lt?`<Relationship Id="rId1" Type="${R}/slideLayout" Target="${esc(a.lt)}"/>`:''}</Relationships>`);
 const cs=all(n.doc,P,'cSld')[0],setBg=c=>{for(const b of all(n.doc,P,'bg'))b.remove();cs.insertBefore(frag(`<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${c}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`,n.doc),cs.firstChild)};
 if(t.bg)setBg(t.bg);
 for(const d of t.s)tree(n).append(frag(TPLT(d[0],d[1],d[2]),n.doc));
 if(D.t&&D.t.l){const m=DK_MAP(D.t);if(D.dark)Object.assign(m,DK_DARK(D.pn||'1e293b'));
  for(const x of all(n.doc,A,'srgbClr')){const v=(x.getAttribute('val')||'').toLowerCase();if(m[v])x.setAttribute('val',m[v])}}
 if(D.dark){if(!t.bg||k=='quote')setBg(D.bg||'0f172a');
  for(const r of all(n.doc,A,'rPr'))if(!k1(r,A,'solidFill'))r.append(frag('<a:solidFill><a:srgbClr val="f1f5f9"/></a:solidFill>',n.doc))}
 return n}
function applyDeck(key){const D=DECKS[key];if(!D||!S)return;const old=S.slides.slice(),a=old[0];
 for(const k of D.s)S.slides.push(deckSlide(k,a,D)); // push as we go so newPath() never repeats a name
 S.slides.splice(0,old.length);for(const s of old)if(!s.isNew){S.gone.push(s);S.del.add(s.path);S.del.add(relPath(s.path))}
 S.cur=0;S.sel=null;S.st=1;X.O().dirty=1}
window.pptDeckPicker=()=>new Promise(res=>{
 const ov=E('div','dk-ov');ov.setAttribute('role','dialog');ov.setAttribute('aria-label','اختيار قالب العرض');
 ov.innerHTML='<div class="dk-box"><div class="dk-h"><b>📊 اختر قالب العرض</b><button type="button" class="dk-x" aria-label="إلغاء">✕</button></div><p class="dk-n">25 قالب جاهز للتعديل — أو ابدأ بقالب فارغ.</p><div class="dk-g"></div></div>';
 const g=ov.querySelector('.dk-g'),done=v=>{ov.remove();res(v)};
 const DG=[['الأساسية',['blank','biz','report','navy','slate','bw']],['تعليم وصحة',['edu','medical','nature','lime']],['أعمال ومشاريع',['sales','project','finance','startup','crimson']],['إبداعي ومناسبات',['creative','rose','event','kids','ocean','coffee']],['داكنة',['dark','gold','tech','neon']]];
 for(const[gt,ks]of DG){g.append(E('h6','dk-t',gt));
  for(const k of ks){const D=DECKS[k];if(!D)continue;
   const b=E('button','dk-c'+(k=='blank'?' dk-b':''));b.type='button';
   b.innerHTML=`<span class="dk-p" style="background:${k=='blank'?'#fff':'#'+D.t.d}"><i style="background:#${D.t.p}"></i><i style="background:#${D.t.s}"></i><i style="background:#${D.t.a}"></i></span><b>${D.ic} ${D.n}</b><small>${D.d}</small><em>${k=='blank'?'شريحة واحدة':D.s.length+' شرائح جاهزة'}</em>`;
   b.onclick=()=>done(k);g.append(b)}}
 ov.querySelector('.dk-x').onclick=()=>done(null);ov.addEventListener('click',e=>{if(e.target===ov)done(null)});
 document.body.append(ov)});
// ---------- slide size (16:9 / 4:3 / 9:16) ----------
const SZ={'169':[12192000,6858000],'43':[9144000,6858000],'916':[5143500,9144000]};
function scaleDoc(d,f,dx,dy){const t=all(d,P,'spTree')[0];
 if(t)for(const e of [...t.children]){const n=e.localName;let x=null;
  if(n=='sp'||n=='pic'||n=='cxnSp')x=k1(k1(e,P,'spPr'),A,'xfrm');else if(n=='grpSp')x=k1(k1(e,P,'grpSpPr'),A,'xfrm');else if(n=='graphicFrame')x=k1(e,P,'xfrm');
  if(!x)continue;const o=k1(x,A,'off'),c=k1(x,A,'ext');
  if(o){o.setAttribute('x',Math.round(o.getAttribute('x')*f+dx));o.setAttribute('y',Math.round(o.getAttribute('y')*f+dy))}
  if(c){c.setAttribute('cx',Math.round(c.getAttribute('cx')*f));c.setAttribute('cy',Math.round(c.getAttribute('cy')*f))}}
 for(const r of [...all(d,A,'rPr'),...all(d,A,'defRPr'),...all(d,A,'endParaRPr')])if(r.getAttribute('sz'))r.setAttribute('sz',Math.max(100,Math.round(r.getAttribute('sz')*f)));
 for(const g of all(d,A,'gridCol'))g.setAttribute('w',Math.round(g.getAttribute('w')*f));for(const r of all(d,A,'tr'))r.setAttribute('h',Math.round(r.getAttribute('h')*f))}
async function resize(k){const z=SZ[k];if(!z||(z[0]==S.W&&z[1]==S.H))return;
 if(!await raConfirm('هغيّر مقاس كل الشرايح، والمحتوى هيتصغّر ويتوسّط عشان يناسب المقاس الجديد. التغيير ده مش بيتراجع عنه بزرار ↶. تكمّل؟'))return;
 commit();X.busy(true,'جاري تغيير المقاس…');
 try{const f=Math.min(z[0]/S.W,z[1]/S.H),dx=(z[0]-S.W*f)/2,dy=(z[1]-S.H*f)/2;S.dd=S.dd||new Set();
  for(const n of Object.keys(S.z.m))if(/^ppt\/(slideLayouts\/slideLayout|slideMasters\/slideMaster)\d+\.xml$/.test(n)){const d=S.docs[n]||(S.docs[n]=ps(await S.rd(n)||'<x/>'));scaleDoc(d,f,dx,dy);S.dd.add(n)}
  for(const sl of S.slides){scaleDoc(sl.doc,f,dx,dy);sl.t=1}
  S.W=z[0];S.H=z[1];S.size=z;S.hist=[];S.redo=[];S.sel=null;S.zm=1;X.O().dirty=1;X.busy(false);draw();thumbs();X.toast('اتغيّر المقاس')}
 catch(e){X.fail('تعذّر تغيير المقاس',e)}}
function mv(d){const i=S.cur,j=i+d;if(j<0||j>=S.slides.length)return;commit();snap();[S.slides[i],S.slides[j]]=[S.slides[j],S.slides[i]];S.cur=j;S.st=1;X.O().dirty=1;thumbs()}

// ---------- transitions ----------
const TRS={none:['بدون انتقال'],fade:['تلاشي','fade'],fadeb:['تلاشي عبر الأسود','fade'],push:['دفع','push',1],cover:['تغطية','cover',1],pull:['كشف','pull',1],wipe:['مسح','wipe',1],split:['تقسيم','split'],zoom:['تكبير','zoom'],circle:['دائرة','circle'],diamond:['معيّن','diamond'],dissolve:['تحلل','dissolve']},
 TRD=[['l','⬅ لليسار'],['r','➡ لليمين'],['u','⬆ لأعلى'],['d','⬇ لأسفل']],TRV={l:[-1,0],r:[1,0],u:[0,-1],d:[0,1]};
function getTr(s){const o={t:'none',dir:'l',spd:'med',adv:0},t=all(s.doc,P,'transition')[0];if(!t)return o;
 o.spd=/^(fast|med|slow)$/.test(t.getAttribute('spd')||'')?t.getAttribute('spd'):'med';o.adv=+t.getAttribute('advTm')||0;const c=t.firstElementChild;if(!c)return o;const n=c.localName,d=c.getAttribute('dir')||'';
 if(/^[lrud]$/.test(d))o.dir=d;o.t=n=='fade'?(c.getAttribute('thruBlk')=='1'?'fadeb':'fade'):(TRS[n]&&n!='none'&&n!='fadeb')?n:'fade';return o}
function setTr(s,o,ns){ns||snap();for(const t of all(s.doc,P,'transition')){let n=t;while(n.parentNode&&/^(Choice|Fallback|AlternateContent)$/.test(n.parentNode.localName))n=n.parentNode;n.remove()}
 if(o.t!='none'){const [,el,dirs]=TRS[o.t],ref=all(s.doc,P,'clrMapOvr')[0]||all(s.doc,P,'cSld')[0],
  inner=`<p:${el}${dirs?` dir="${o.dir||'l'}"`:''}${o.t=='fadeb'?' thruBlk="1"':''}${el=='split'?' orient="vert" dir="out"':''}/>`;
  ref.after(frag(`<p:transition spd="${o.spd||'med'}"${o.adv?` advTm="${o.adv}"`:''}>${inner}</p:transition>`,s.doc))}s.t=1;X.O().dirty=1}

// ---------- element animations (entrance effects) ----------
const FXD={appear:['entr',1,0,'appear'],fade:['entr',10,0,'fade'],flyR:['entr',2,2,'flyR'],flyL:['entr',2,8,'flyL'],flyT:['entr',2,1,'flyT'],flyB:['entr',2,4,'flyB'],zoom:['entr',53,16,'zoom'],wipe:['entr',22,8,'wipe'],
 xappear:['exit',1,0,'appear'],xfade:['exit',10,0,'fade'],xflyR:['exit',2,2,'flyR'],xflyL:['exit',2,8,'flyL'],xflyT:['exit',2,1,'flyT'],xflyB:['exit',2,4,'flyB'],xzoom:['exit',53,16,'zoom'],xwipe:['exit',22,8,'wipe'],
 grow:['emph',6,0,'grow'],spin:['emph',8,0,'spin'],pulse:['emph',26,0,'pulse'],
 drift:['entr',42,0,'drift'],bounce:['entr',26,0,'bounce'],turn:['entr',31,0,'turn'],xdrift:['exit',42,0,'drift'],xbounce:['exit',26,0,'bounce'],xturn:['exit',31,0,'turn'],teeter:['emph',32,0,'teeter']},
 FXL=[['none','✨ بدون أنيميشن'],['appear','دخول · ظهور'],['fade','دخول · تلاشي'],['flyR','دخول · انزلاق من اليمين'],['flyL','دخول · انزلاق من اليسار'],['flyT','دخول · انزلاق من فوق'],['flyB','دخول · انزلاق من تحت'],['zoom','دخول · تكبير'],['wipe','دخول · مسح'],
  ['xappear','خروج · اختفاء'],['xfade','خروج · تلاشي'],['xflyR','خروج · انزلاق لليمين'],['xflyL','خروج · انزلاق لليسار'],['xflyT','خروج · انزلاق لفوق'],['xflyB','خروج · انزلاق لتحت'],['xzoom','خروج · تصغير'],['xwipe','خروج · مسح'],
  ['grow','تأكيد · تكبير'],['spin','تأكيد · دوران'],['pulse','تأكيد · نبض'],
  ['drift','دخول · طفو'],['bounce','دخول · ارتداد'],['turn','دخول · تكبير ودوران'],['xdrift','خروج · طفو'],['xbounce','خروج · ارتداد'],['xturn','خروج · تصغير ودوران'],['teeter','تأكيد · اهتزاز']],
 FXK=Object.fromEntries(Object.entries(FXD).map(([k,v])=>[v.slice(0,3).join(':'),k])),
 STL=[['click','🖱 بالضغط'],['with','⏯ مع السابق'],['after','⏭ بعد السابق']],
 DURS=[[250,'⚡ سريع'],[500,'مدة عادية'],[1000,'بطيء'],[2000,'بطيء جدًا']],DLYS=[[0,'بدون تأخير'],[500,'تأخير 0.5 ث'],[1000,'تأخير 1 ث'],[2000,'تأخير 2 ث'],[3000,'تأخير 3 ث']];
function animRead(s){const tm=all(s.doc,P,'timing')[0],list=[];let foreign=false;if(!tm)return{tm:null,list,foreign};
 for(const c of all(tm,P,'cTn')){const nt=c.getAttribute('nodeType')||'';if(nt=='interactiveSeq'){foreign=true;continue}if(!/Effect$/.test(nt))continue;
  const tg=all(c,P,'spTgt')[0],k=FXK[(c.getAttribute('presetClass')||'')+':'+c.getAttribute('presetID')+':'+(c.getAttribute('presetSubtype')||0)];
  if(!tg||k1(tg,P,'txEl')||!k||!/^(click|with|after)Effect$/.test(nt)||list.some(a=>a.id==tg.getAttribute('spid'))){foreign=true;continue}
  const ds=all(c,P,'cTn').filter(x=>x!==c).map(x=>+x.getAttribute('dur')).filter(x=>x>1),sc=k1(c,P,'stCondLst'),cd=sc&&k1(sc,P,'cond');
  list.push({id:tg.getAttribute('spid'),fx:k,st:nt.replace('Effect',''),dur:ds.length?Math.max(...ds):500,dl:(cd&&+cd.getAttribute('delay'))||0})}
 return{tm,list,foreign}}
function animGroups(L){const gs=[];let last=null;for(const a of L){if(!gs.length||a.st=='click'){gs.push([]);last=null}const t=!last?0:a.st=='after'?last.t+(last.a.dl||0)+(last.a.dur||500):last.t,it={a,t};gs[gs.length-1].push(it);last=it}return gs}
function animXml(L,s){let n=2;const id=()=>++n,T=sp=>`<p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl>`,bld=[],
 el=sp=>{const c=all(s.doc,P,'cNvPr').find(x=>x.getAttribute('id')==sp);return c&&c.parentNode.parentNode},
 an=(sp,a,f,t,W)=>`<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${id()}" dur="${W}" fill="hold"/>${T(sp)}<p:attrNameLst><p:attrName>${a}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${f}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${t}"/></p:val></p:tav></p:tavLst></p:anim>`,
 ak=(sp,a,W,k)=>`<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="${id()}" dur="${W}" fill="hold"/>${T(sp)}<p:attrNameLst><p:attrName>${a}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst>${k.map(([t,v])=>`<p:tav tm="${t}"><p:val><p:strVal val="${v}"/></p:val></p:tav>`).join('')}</p:tavLst></p:anim>`,
 ae=(sp,f,W,out)=>`<p:animEffect transition="${out?'out':'in'}" filter="${f}"><p:cBhvr><p:cTn id="${id()}" dur="${W}"/>${T(sp)}</p:cBhvr></p:animEffect>`,
 vis=(sp,v,d)=>`<p:set><p:cBhvr><p:cTn id="${id()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${d}"/></p:stCondLst></p:cTn>${T(sp)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${v}"/></p:to></p:set>`,
 eff=(a,nt)=>{const sp=a.id,[cls,pid,sub,kd]=FXD[a.fx],W=a.dur||500,out=cls=='exit',e=el(sp),isSp=e&&e.localName=='sp',cid=id();
  if(isSp)bld.push([sp,!!k1(k1(e,P,'spPr'),A,'solidFill')]);
  let b='';
  if(kd=='fade')b=ae(sp,'fade',W,out);else if(kd=='wipe')b=ae(sp,'wipe(left)',W,out);
  else if(kd=='zoom'){const z=q=>out?an(sp,q,'#'+q,'0',W):an(sp,q,'0','#'+q,W);b=z('ppt_w')+z('ppt_h')+ae(sp,'fade',W,out)}
  else if(kd[0]=='f'&&kd[1]=='l'){const v={flyR:['1+#ppt_w/2','x'],flyL:['0-#ppt_w/2','x'],flyT:['0-#ppt_h/2','y'],flyB:['1+#ppt_h/2','y']}[kd],X=v[1]=='x'?v[0]:'#ppt_x',Y=v[1]=='y'?v[0]:'#ppt_y';
   b=out?an(sp,'ppt_x','#ppt_x',X,W)+an(sp,'ppt_y','#ppt_y',Y,W):an(sp,'ppt_x',X,'#ppt_x',W)+an(sp,'ppt_y',Y,'#ppt_y',W)}
  else if(kd=='drift')b=(out?an(sp,'ppt_y','#ppt_y','#ppt_y+.1',W):an(sp,'ppt_y','#ppt_y+.1','#ppt_y',W))+ae(sp,'fade',W,out);
  else if(kd=='bounce')b=ak(sp,'ppt_y',W,out?[[0,'#ppt_y'],[20000,'#ppt_y-.08'],[100000,'1+#ppt_h/2']]:[[0,'0-#ppt_h/2'],[50000,'#ppt_y'],[65000,'#ppt_y-.1'],[80000,'#ppt_y'],[90000,'#ppt_y-.03'],[100000,'#ppt_y']]);
  else if(kd=='turn'){const z=q=>out?an(sp,q,'#'+q,'0',W):an(sp,q,'0','#'+q,W);b=z('ppt_w')+z('ppt_h')+(out?an(sp,'r','0','90',W):an(sp,'r','-90','0',W))+ae(sp,'fade',W,out)}
  else if(kd=='teeter')b=ak(sp,'r',W,[[0,'0'],[20000,'-5'],[40000,'5'],[60000,'-5'],[80000,'5'],[100000,'0']]);
  else if(kd=='grow')b=`<p:animScale><p:cBhvr><p:cTn id="${id()}" dur="${W}" fill="hold"/>${T(sp)}</p:cBhvr><p:by x="150000" y="150000"/></p:animScale>`;
  else if(kd=='pulse')b=`<p:animScale><p:cBhvr><p:cTn id="${id()}" dur="${W}" autoRev="1" fill="hold"/>${T(sp)}</p:cBhvr><p:by x="112000" y="112000"/></p:animScale>`;
  else if(kd=='spin')b=`<p:animRot by="21600000"><p:cBhvr><p:cTn id="${id()}" dur="${W}" fill="hold"/>${T(sp)}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`;
  const pre=cls=='entr'?vis(sp,'visible',0):'',post=out?vis(sp,'hidden',kd=='appear'?0:Math.max(0,W-1)):'';
  return `<p:par><p:cTn id="${cid}" presetID="${pid}" presetClass="${cls}" presetSubtype="${sub}" fill="hold"${isSp?' grpId="0"':''} nodeType="${nt}"><p:stCondLst><p:cond delay="${a.dl||0}"/></p:stCondLst><p:childTnLst>${pre}${b}${post}</p:childTnLst></p:cTn></p:par>`};
 let out='';
 for(const g of animGroups(L)){const oid=id();let inner='',i=0;
  while(i<g.length){const t=g[i].t,pid=id();let ch='';
   while(i<g.length&&g[i].t==t){ch+=eff(g[i].a,g[i].a.st=='click'||(i==0&&!inner)?'clickEffect':g[i].a.st+'Effect');i++}
   inner+=`<p:par><p:cTn id="${pid}" fill="hold"><p:stCondLst><p:cond delay="${t}"/></p:stCondLst><p:childTnLst>${ch}</p:childTnLst></p:cTn></p:par>`}
  out+=`<p:par><p:cTn id="${oid}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`}
 const bl=[...new Map(bld.map(x=>[x[0],x])).values()].map(x=>`<p:bldP spid="${x[0]}" grpId="0"${x[1]?' animBg="1"':''}/>`).join('');
 return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${out}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>${bl?`<p:bldLst>${bl}</p:bldLst>`:''}</p:timing>`}
function animWrite(s,L){if(L.length)L[0].st='click';for(const t of all(s.doc,P,'timing'))t.remove();
 if(L.length){let ref=all(s.doc,P,'transition')[0];while(ref&&ref.parentNode&&/^(Choice|Fallback|AlternateContent)$/.test(ref.parentNode.localName))ref=ref.parentNode;ref=ref||all(s.doc,P,'clrMapOvr')[0]||all(s.doc,P,'cSld')[0];ref.after(frag(animXml(L,s),s.doc))}
 s.t=1;X.O().dirty=1}
function animDrop(s,id){const r=animRead(s);if(r.foreign){r.tm.remove();s.t=1;return}animWrite(s,r.list.filter(a=>a.id!=id))}
async function animEdit(f){const s=cur(),r=animRead(s);if(r.foreign&&!await raConfirm('الشريحة دي فيها أنيميشن متقدّم مش مدعوم هنا، وهيتشال لو كمّلت. تكمّل؟'))return;
 snap();animWrite(s,f(r.list)||r.list);S.panel&&S.panel()}
const animSet=(id,fx)=>animEdit(L=>{if(fx=='none')return L.filter(x=>x.id!=id);const a=L.find(x=>x.id==id);if(a)a.fx=fx;else L.push({id,fx,st:'click',dur:500,dl:0});return L}),
 animStart=(id,v)=>animEdit(L=>{const a=L.find(x=>x.id==id);a&&(a.st=v);return L}),
 animProp=(id,k,v)=>animEdit(L=>{const a=L.find(x=>x.id==id);a&&(a[k]=+v);return L}),
 animMove=(id,d)=>animEdit(L=>{const i=L.findIndex(x=>x.id==id),j=i+d;if(i<0||j<0||j>=L.length)return L;[L[i],L[j]]=[L[j],L[i]];return L});
function animCtl(){const e=S.sel,c=e&&all(e,P,'cNvPr')[0];if(!c)return[];const id=c.getAttribute('id'),r=animRead(cur()),i=r.list.findIndex(x=>x.id==id),a=r.list[i];
 const sl=(opts,val,f,label)=>{const x=E('select');x.setAttribute('aria-label',label);if(!opts.some(o=>o[0]==val))opts=[...opts,[val,(val/1000)+' ث']];x.innerHTML=opts.map(o=>`<option value="${o[0]}">${o[1]}</option>`).join('');x.value=val;x.onchange=()=>f(x.value);return{el:x}},
  bt=(t,f,lb)=>{const x=E('button','',t);if(lb)x.title=lb;x.addEventListener('mousedown',ev=>ev.preventDefault());x.onclick=f;return{el:x}};
 const out=[sl(FXL,a?a.fx:'none',v=>animSet(id,v),'أنيميشن')];
 if(a){if(i>0)out.push(sl(STL,a.st,v=>animStart(id,v),'بداية الأنيميشن'));
  if(!/appear$/.test(a.fx))out.push(sl(DURS,a.dur||500,v=>animProp(id,'dur',v),'مدة الأنيميشن'));
  out.push(sl(DLYS,a.dl||0,v=>animProp(id,'dl',v),'تأخير الأنيميشن'),{el:E('span','pe-hint','الترتيب '+(i+1)+'/'+r.list.length)},bt('⏫',()=>animMove(id,-1),'تقديم في الترتيب'),bt('⏬',()=>animMove(id,1),'تأخير في الترتيب'),bt('▶ معاينة',()=>play(S.cur,null,1)))}
 return out}
function anSetup(s,st,full){const by={};for(const d of st.children){const c=d._e&&all(d._e,P,'cNvPr')[0];c&&(by[c.getAttribute('id')]=d)}
 const gs=animGroups(animRead(s).list.filter(a=>by[a.id]));if(!gs.length)return null;let k=0,tm=[];const W=parseFloat(st.style.width)||1,H=parseFloat(st.style.height)||1,
  hasIn=new Set(gs.flat().filter(it=>FXD[it.a.fx][0]=='entr').map(it=>it.a.id));
 const reset=n=>{tm.forEach(clearTimeout);tm=[];const v={};for(const id in by){by[id].getAnimations().forEach(x=>x.cancel());by[id].style.scale='';v[id]=!hasIn.has(id)}
  for(let j=0;j<n;j++)for(const it of gs[j]){const c=FXD[it.a.fx][0];if(c=='entr')v[it.a.id]=true;else if(c=='exit')v[it.a.id]=false;else if(it.a.fx=='grow')by[it.a.id].style.scale='1.5'}
  for(const id in by)by[id].style.visibility=v[id]?'visible':'hidden'};
 const run=it=>{const d=by[it.a.id],[cls,,,kd]=FXD[it.a.fx],b=d.style.transform||'',o={duration:it.a.dur||500,easing:kd=='bounce'?'linear':cls=='exit'?'cubic-bezier(.4,0,1,1)':cls=='emph'?'ease-in-out':(kd=='zoom'||kd=='turn')?'cubic-bezier(.2,.8,.2,1.12)':'cubic-bezier(.22,.61,.36,1)'},
   ol=d.offsetLeft||0,ot=d.offsetTop||0,ow=d.offsetWidth||0,oh=d.offsetHeight||0,dR=Math.max(W-ol,ow,40),dL=Math.max(ol+ow,ow,40),dT=Math.max(ot+oh,oh,40),dB=Math.max(H-ot,oh,40),ty=y=>(`translateY(${y}px) `+b).trim(),tf=x=>[{transform:(x+' '+b).trim()},{transform:b||'none'}];
  const K={fade:[{opacity:0},{opacity:1}],flyR:tf(`translateX(${dR}px)`),flyL:tf(`translateX(${-dL}px)`),flyT:tf(`translateY(${-dT}px)`),flyB:tf(`translateY(${dB}px)`),zoom:[{transform:(b+' scale(0)').trim(),opacity:0},{transform:b||'none',opacity:1}],wipe:[{clipPath:'inset(0 100% 0 0)'},{clipPath:'inset(0 0 0 0)'}],
   drift:[{opacity:0,transform:ty(H*.1)},{opacity:1,transform:b||'none'}],turn:[{transform:(b+' scale(0) rotate(-90deg)').trim(),opacity:0},{transform:b||'none',opacity:1}],
   bounce:[{transform:ty(-dT),easing:'ease-in'},{transform:ty(0),offset:.5,easing:'ease-out'},{transform:ty(-H*.1),offset:.65,easing:'ease-in'},{transform:ty(0),offset:.8,easing:'ease-out'},{transform:ty(-H*.03),offset:.9,easing:'ease-in'},{transform:b||'none'}]}[kd],
   XK={bounce:[{transform:b||'none',easing:'ease-out'},{transform:ty(-H*.08),offset:.2,easing:'ease-in'},{transform:ty(dB)}]}[kd];
  d.style.willChange='transform,opacity';setTimeout(()=>{d.style.willChange=''},(it.a.dur||500)+80);
  if(cls=='entr'){d.style.visibility='visible';K&&d.animate(K,o)}
  else if(cls=='exit'){const hide=()=>{d.style.visibility='hidden'},a=K&&d.animate(XK||[K[1],K[0]],{...o,fill:'forwards'});a?(a.onfinish=hide):hide()}
  else{const em={grow:[{scale:1},{scale:1.5}],spin:[{rotate:'0deg'},{rotate:'360deg'}],pulse:[{scale:1},{scale:1.12},{scale:1}],teeter:[{rotate:'0deg'},{rotate:'-5deg'},{rotate:'5deg'},{rotate:'-5deg'},{rotate:'5deg'},{rotate:'0deg'}]}[kd],a=em&&d.animate(em,o);kd=='grow'&&a&&(a.onfinish=()=>{d.style.scale='1.5'})}};
 reset(full?gs.length:0);k=full?gs.length:0;
 return{next(){if(k>=gs.length)return 0;let ms=1;for(const it of gs[k++]){tm.push(setTimeout(()=>run(it),it.t+(it.a.dl||0)));ms=Math.max(ms,it.t+(it.a.dl||0)+(it.a.dur||500))}return ms},
  prev(){if(k<=0)return false;reset(--k);return true},stop(){tm.forEach(clearTimeout);tm=[]}}}

// ---------- speaker notes ----------
const NHD='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
 NGRP='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>',
 NMX=`<p:notesMaster ${NSD}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${NGRP}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>`,
 NSX=`<p:notes ${NSD}><p:cSld><p:spTree>${NGRP}<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`,
 ntBody=d=>all(d,P,'sp').find(q=>{const h=all(q,P,'ph')[0];return h&&h.getAttribute('type')=='body'}),
 ntText=d=>{const b=ntBody(d),tb=b&&k1(b,P,'txBody');return tb?kids(tb,A,'p').map(p=>[...p.childNodes].map(c=>c.localName=='br'?'\n':/^(r|fld)$/.test(c.localName)?((k1(c,A,'t')||{}).textContent||''):'').join('')).join('\n'):''},
 ntParas=(t,rx)=>t.split('\n').map(l=>l?`<a:p><a:r>${rx}<a:t>${esc(l)}</a:t></a:r></a:p>`:'<a:p><a:endParaRPr lang="en-US"/></a:p>');
function ntSet(d,t){let b=ntBody(d);if(!b){b=frag(`<p:sp><p:nvSpPr><p:cNvPr id="${Math.max(1,...all(d,P,'cNvPr').map(x=>+x.getAttribute('id')||0))+1}" name="Notes Placeholder"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>`,d);tree({doc:d}).append(b)}
 let tb=k1(b,P,'txBody');if(!tb){tb=frag('<p:txBody><a:bodyPr/><a:lstStyle/></p:txBody>',d);b.append(tb)}
 const rp=all(tb,A,'rPr')[0],rx=rp?ser(rp):'<a:rPr lang="ar-EG"/>';for(const p of kids(tb,A,'p'))p.remove();for(const x of ntParas(t,rx))tb.append(frag(x,d))}
async function ntLoad1(s){if(s.nt)return s.nt;const nt={text:'',path:null,doc:null,dirty:0};
 if(!s.isNew){const r=all(s.rel,PR,'Relationship').find(x=>/notesSlide$/.test(x.getAttribute('Type')));
  if(r){nt.path=norm(s.path,r.getAttribute('Target'));const t=await S.rd(nt.path);if(t){nt.doc=ps(t);nt.text=ntText(nt.doc)}}}
 return s.nt=nt}
async function ntBind(nb){if(!S)return;const s=cur();if(S.ntFor===s)return;S.ntFor=null;nb.disabled=true;nb.value='';
 let t;try{t=await ntLoad1(s)}catch(e){t=s.nt={text:'',path:null,doc:null,dirty:0}}
 if(!S||!nb.isConnected)return;if(cur()!==s)return ntBind(nb);nb.value=t.text;nb.disabled=false;S.ntFor=s}
function ntClear(){if(!S||!S.nb||!S.ntFor)return;S.nb.value='';S.nb.oninput()}
async function notesFlush(put,pres,prel,ct){
 const todo=S.slides.filter(s=>s.nt&&s.nt.dirty&&(s.nt.doc||s.nt.text.trim()));if(!todo.length)return;
 try{
  const ov=(p,t)=>{if(all(ct,CT,'Override').some(o=>o.getAttribute('PartName')=='/'+p))return;const o=ct.createElementNS(CT,'Override');o.setAttribute('PartName','/'+p);o.setAttribute('ContentType',t);ct.documentElement.append(o)},
   RT='application/vnd.openxmlformats-officedocument.presentationml.',rels=l=>te.encode(NHD+`<Relationships xmlns="${PR}">${l.map((x,i)=>`<Relationship Id="rId${i+1}" Type="${R}/${x[0]}" Target="${x[1]}"/>`).join('')}</Relationships>`);
  let nmp=null;
  const master=async()=>{if(nmp)return nmp;const mr=all(prel,PR,'Relationship').find(x=>/notesMaster$/.test(x.getAttribute('Type')));
   if(mr)return nmp=norm('ppt/presentation.xml',mr.getAttribute('Target'));
   const nm=S.nm||(S.nm={path:'ppt/notesMasters/notesMaster1.xml'});
   if(!nm.theme){const src=Object.keys(S.z.m).find(n=>/^ppt\/theme\/theme\d+\.xml$/.test(n));if(!src)throw new Error('no theme');let k=1;while(S.z.m['ppt/theme/theme'+k+'.xml'])k++;nm.src=src;nm.theme='ppt/theme/theme'+k+'.xml'}
   put.set(nm.theme,await X.zipRead(S.z,nm.src));ov(nm.theme,'application/vnd.openxmlformats-officedocument.theme+xml');
   put.set(nm.path,te.encode(NHD+NMX));put.set(relPath(nm.path),rels([['theme','../'+nm.theme.slice(4)]]));ov(nm.path,RT+'notesMaster+xml');
   const rid=nextRid(prel),r=prel.createElementNS(PR,'Relationship');r.setAttribute('Id',rid);r.setAttribute('Type',R+'/notesMaster');r.setAttribute('Target',nm.path.slice(4));prel.documentElement.append(r);
   const ml=all(pres,P,'sldMasterIdLst')[0],l=pres.createElementNS(P,'p:notesMasterIdLst'),i=pres.createElementNS(P,'p:notesMasterId');i.setAttributeNS(R,'r:id',rid);l.append(i);ml.after(l);
   if(!all(pres,P,'notesSz')[0]){const z=pres.createElementNS(P,'p:notesSz');z.setAttribute('cx','6858000');z.setAttribute('cy','9144000');all(pres,P,'sldSz')[0].after(z)}
   return nmp=nm.path};
  for(const s of todo){const nt=s.nt;
   if(nt.doc&&!nt.mine){ntSet(nt.doc,nt.text);put.set(nt.path,xml(nt.doc));continue}
   const m=await master(),er=all(s.rel,PR,'Relationship').find(x=>/notesSlide$/.test(x.getAttribute('Type')));
   if(!nt.path){if(er)nt.path=norm(s.path,er.getAttribute('Target'));else{let k=1;const used=new Set(S.slides.map(x=>x.nt&&x.nt.path));while(S.z.m['ppt/notesSlides/notesSlide'+k+'.xml']||used.has('ppt/notesSlides/notesSlide'+k+'.xml'))k++;nt.path='ppt/notesSlides/notesSlide'+k+'.xml'}nt.mine=1}
   if(!nt.doc)nt.doc=ps(NSX);ntSet(nt.doc,nt.text);
   put.set(nt.path,xml(nt.doc));put.set(relPath(nt.path),rels([['notesMaster','../'+m.slice(4)],['slide','../slides/'+s.path.split('/').pop()]]));ov(nt.path,RT+'notesSlide+xml');
   if(!er){const r=s.rel.createElementNS(PR,'Relationship');r.setAttribute('Id',nextRid(s.rel));r.setAttribute('Type',R+'/notesSlide');r.setAttribute('Target','../'+nt.path.slice(4));s.rel.documentElement.append(r)}
   put.set(relPath(s.path),xml(s.rel))}
 }catch(e){X.toast('تعذّر حفظ الملاحظات')}}

// ---------- slideshow ----------
function play(from,done,auto){commit();const ov=E('div','pe-pl'),ob=window.raBack,K=()=>Math.min(innerWidth/S.W,innerHeight/S.H);let i=-1,L=null,lock=0,AN=null,AT=0,ended=0;
 const x=E('button','pe-px','✕'),cn=E('div','pe-pc');const nbt=E('button','pe-nb','📝'),np=E('div','pe-np');np.hidden=true;ov.append(x,cn,nbt,np);
 const ntShow=n=>{if(np.hidden)return;np.textContent='';ntLoad1(S.slides[n]).then(t=>{if(i==n&&!np.hidden)np.textContent=t.text||'(مفيش ملاحظات للشريحة دي)'}).catch(()=>{})};nbt.onclick=e=>{e.stopPropagation();np.hidden=!np.hidden;ntShow(i)};document.body.append(ov);ov.requestFullscreen&&ov.requestFullscreen().catch(()=>{});
 const end=()=>{ended=1;clearTimeout(AT);AN&&AN.stop&&AN.stop();done&&done();removeEventListener('keydown',kd,true);window.raBack=ob;ov.remove();document.fullscreenElement&&document.exitFullscreen&&document.exitFullscreen().catch(()=>{})};
 const arm=ms=>{clearTimeout(AT);const o=S.slides[i]&&getTr(S.slides[i]);if(!o||!o.adv||ended)return;const at=i;AT=setTimeout(()=>{if(ended||i!=at||lock)return;adv()},o.adv+(ms||0))};
 const show=(n,full)=>{if(n<0||lock)return;clearTimeout(AT);AN&&AN.stop&&AN.stop();if(n>=S.slides.length)return end();const w=E('div','pe-pw'),st=E('div');paint(S.slides[n],st,K(),1);w.append(st);AN=anSetup(S.slides[n],st,full);ntShow(n);
  const old=L,o=old?getTr(S.slides[n]):{t:'none'},t=o.t,fw=n>=i?1:-1,Dm={fast:300,slow:1000}[o.spd]||520,D={duration:Dm,easing:'ease'};ov.insertBefore(w,x);L=w;i=n;cn.textContent=(n+1)+' / '+S.slides.length;
  if(!old)return arm();if(t=='none'){old.remove();return arm()}lock=1;
  const v=(TRV[o.dir]||TRV.l).map(c=>c*fw),tr=(a,b)=>`translate(${a*100}%,${b*100}%)`,ins=(a,b,c,d)=>`inset(${a}% ${b}% ${c}% ${d}%)`;
  if(t=='push'){w.animate([{transform:tr(-v[0],-v[1])},{transform:'none'}],D);old.animate([{transform:'none'},{transform:tr(v[0],v[1])}],D)}
  else if(t=='cover'){w.animate([{transform:tr(-v[0],-v[1])},{transform:'none'}],D);old.animate([{transform:'none'},{transform:tr(v[0]*.3,v[1]*.3),opacity:.5}],D)}
  else if(t=='pull'){ov.insertBefore(w,old);old.animate([{transform:'none'},{transform:tr(v[0],v[1])}],D)}
  else if(t=='wipe')w.animate([{clipPath:ins(v[1]<0?100:0,v[0]>0?100:0,v[1]>0?100:0,v[0]<0?100:0)},{clipPath:ins(0,0,0,0)}],D);
  else if(t=='split')w.animate([{clipPath:ins(0,50,0,50)},{clipPath:ins(0,0,0,0)}],D);
  else if(t=='circle')w.animate([{clipPath:'circle(0% at 50% 50%)'},{clipPath:'circle(75% at 50% 50%)'}],D);
  else if(t=='diamond')w.animate([{clipPath:'polygon(50% 50%,50% 50%,50% 50%,50% 50%)'},{clipPath:'polygon(50% -60%,160% 50%,50% 160%,-60% 50%)'}],D);
  else if(t=='dissolve')w.animate([{opacity:0,filter:'blur(16px)'},{opacity:1,filter:'blur(0)'}],D);
  else if(t=='fadeb'){old.animate([{opacity:1},{opacity:0}],{duration:Dm/2,fill:'forwards'});w.animate([{opacity:0},{opacity:1}],{duration:Dm/2,delay:Dm/2,fill:'backwards'})}
  else w.animate(t=='zoom'?[{transform:'scale(.5)',opacity:0},{transform:'none',opacity:1}]:[{opacity:0},{opacity:1}],D);
  setTimeout(()=>{old.remove();lock=0;arm()},Dm+30)};
 const adv=()=>{if(lock)return;const ms=AN&&AN.next();if(ms){arm(ms);return}show(i+1)},bk=()=>{if(lock)return;clearTimeout(AT);AN&&AN.prev()||show(i-1,1)};
 const kd=e=>{if(e.key=='Escape'){e.preventDefault();end()}else if(/^(ArrowRight|ArrowDown| |PageDown)$/.test(e.key))adv();else if(/^(ArrowLeft|ArrowUp|PageUp)$/.test(e.key))bk()};
 addEventListener('keydown',kd,true);window.raBack=()=>{if(ov.isConnected){end();return true}return ob?ob():false};x.onclick=e=>{e.stopPropagation();end()};
 let sx=0;ov.addEventListener('pointerdown',e=>{sx=e.clientX});ov.addEventListener('pointerup',e=>{if(e.target===x||e.target===nbt||np.contains(e.target))return;const d=e.clientX-sx;
  if(d<-50)adv();else if(d>50)bk();else if(e.clientX<innerWidth/3)bk();else adv()});show(from);
 if(auto&&AN){const go=()=>{if(ended||!AN)return;const ms=AN.next();if(ms)setTimeout(go,ms+450);else setTimeout(()=>{ended||end()},1200)};setTimeout(go,700)}}

// ---------- history / clipboard / arrange / table / text extras ----------
const shot=()=>({sl:S.slides.map(s=>[s,ser(s.doc),ser(s.rel)]),cur:S.cur});
function snap(){S.hist.push(shot());if(S.hist.length>40)S.hist.shift();S.redo=[]}
function rest(h){S.slides=h.sl.map(([s,d,r])=>{s.doc=ps(d);s.rel=ps(r);s.t=1;return s});const keep=new Set(S.slides);S.gone=S.orig.filter(x=>!keep.has(x));S.del=new Set();
 for(const g of S.gone){S.del.add(g.path);S.del.add(relPath(g.path))}S.cur=Math.min(h.cur,S.slides.length-1);S.sel=null;S.st=1;X.O().dirty=1;draw();thumbs()}
function undo(){if(S.ed){document.execCommand('undo');return}commit();if(!S.hist.length)return;S.redo.push(shot());rest(S.hist.pop())}
function redo(){if(S.ed){document.execCommand('redo');return}if(!S.redo.length)return;S.hist.push(shot());rest(S.redo.pop())}
function copy(){if(S.sel)S.clip={x:ser(S.sel),s:cur()}}
function paste(){if(!S.clip)return;commit();snap();const s=cur(),src=S.clip.s,n=frag(S.clip.x,s.doc);
 for(const b of all(n,A,'blip')){const r=all(src.rel,PR,'Relationship').find(x=>x.getAttribute('Id')==b.getAttributeNS(R,'embed'));if(!r)continue;
  let t=all(s.rel,PR,'Relationship').find(x=>x.getAttribute('Target')==r.getAttribute('Target'));
  if(!t){t=s.rel.createElementNS(PR,'Relationship');t.setAttribute('Id',nextRid(s.rel));t.setAttribute('Type',r.getAttribute('Type'));t.setAttribute('Target',r.getAttribute('Target'));s.rel.documentElement.append(t)}
  b.setAttributeNS(R,'r:embed',t.getAttribute('Id'))}
 const c=all(n,P,'cNvPr')[0];c&&c.setAttribute('id',nid());const x=xf(n),o=x&&k1(x,A,'off');
 if(o){o.setAttribute('x',Math.round(+o.getAttribute('x')+S.W*.03));o.setAttribute('y',Math.round(+o.getAttribute('y')+S.H*.03))}
 tree(s).append(n);touch();S.sel=n;draw();thumb(S.cur)}

// ---------- تنسيق الصور: أشكال / إطارات / ظلال ----------
const PIC_SHAPES={rect:'مستطيل',roundRect:'زوايا دائرية',ellipse:'دائرة / بيضاوي',hexagon:'سداسي',diamond:'معيّن'};
const PG_SH=[['ellipse',1],['roundRect',0],['hexagon',1],['octagon',1],['pentagon',1],['diamond',1],['triangle',0],['star5',1],['star6',1],['star8',1],['heart',1],['plus',1],['moon',0],['cloud',0],['parallelogram',0],['trapezoid',0],['chevron',0],['homePlate',0],['flowChartTerminator',0],['wedgeRoundRectCallout',0],['donut',1]],
 PG_ST=[{n:'دائرة + إطار أبيض',s:'ellipse',q:1,c:'FFFFFF',w:5,sh:1},{n:'دائرة + إطار أزرق',s:'ellipse',q:1,c:'2563EB',w:6},{n:'دائرة بإطار مزدوج',s:'ellipse',q:1,c:'1E3A8A',w:7,d:1},
  {n:'سداسي ذهبي',s:'hexagon',q:1,c:'F59E0B',w:5},{n:'قلب',s:'heart',q:1,c:'E11D48',w:4,sh:1},{n:'نجمة',s:'star5',q:1,c:'FBBF24',w:3,sh:1},
  {n:'بولارويد',s:'rect',c:'FFFFFF',w:12,sh:1,r:-3},{n:'بولارويد مائل',s:'rect',c:'FFFFFF',w:12,sh:1,r:4},{n:'إطار أنيق',s:'rect',c:'1F2937',w:4,sh:1},
  {n:'زوايا ناعمة + ظل',s:'roundRect',sh:1},{n:'كبسولة',s:'flowChartTerminator',c:'0F766E',w:4},{n:'فقاعة كلام',s:'wedgeRoundRectCallout',c:'7C3AED',w:3},
  {n:'معيّن',s:'diamond',q:1,c:'0891B2',w:4},{n:'بدون تنسيق',s:'rect',reset:1}];
for(const k of['triangle','pentagon','octagon','star4','star5','star6','star8','heart','plus','moon','cloud','parallelogram','trapezoid','chevron','homePlate','flowChartTerminator','wedgeRoundRectCallout','donut'])PIC_SHAPES[k]=SH[k][0];
async function picApply(o){const e=S.sel;if(!e||e.localName!='pic')return;const bl=all(e,A,'blip')[0];let did=false;
 if(o.q&&bl){try{const u=await imgUrl(cur(),bl.getAttributeNS(R,'embed')),im=await new Promise((ok,no)=>{const i=new Image();i.onload=()=>ok(i);i.onerror=no;i.src=u}),a=im.naturalWidth/im.naturalHeight;
   applyCrop(e,a>1?[(1-1/a)/2,0,(1-1/a)/2,0]:[0,(1-a)/2,0,(1-a)/2]);did=true}catch(_){X.toast('تعذر قص الصورة لمربع')}}
 else if(o.reset){applyCrop(e,[0,0,0,0]);did=true}
 if(!did)snap();
 const sp=k1(e,P,'spPr');let gm=k1(sp,A,'prstGeom');if(!gm){gm=frag('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',cur().doc);sp.append(gm)}gm.setAttribute('prst',o.s);
 for(const n of kids(sp,A,'ln'))n.remove();for(const n of kids(sp,A,'effectLst'))n.remove();
 if(o.c)sp.append(frag(`<a:ln w="${Math.round(o.w*PT)}"${o.d?' cmpd="dbl"':''}><a:solidFill><a:srgbClr val="${o.c}"/></a:solidFill><a:prstDash val="solid"/></a:ln>`,cur().doc));
 if(o.sh)sp.append(frag('<a:effectLst><a:outerShdw blurRad="95250" dist="63500" dir="2700000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst>',cur().doc));
 const x=xf(e,1);if(x){if(o.r)x.setAttribute('rot',Math.round((((o.r%360)+360)%360)*60000));else x.removeAttribute('rot')}
 touch();draw();thumb(S.cur)}
async function picGallery(){const e=S.sel;if(!e||e.localName!='pic')return X.toast('اختار صورة الأول');const bl=all(e,A,'blip')[0],u=bl&&await imgUrl(cur(),bl.getAttributeNS(R,'embed'));if(!u)return X.toast('تعذر فتح الصورة');
 const old=document.querySelector('.shg');old&&old.remove();
 const w=E('div','shg'),box=E('div','shb','<div class="shh"><b>🎭 أشكال وتصميمات الصورة</b><button type="button" aria-label="إغلاق">✕</button></div>'),close=()=>w.remove();
 box.querySelector('button').onclick=close;w.onpointerdown=ev=>{if(ev.target===w)close()};
 const pv=o=>{const q=SH[o.s];return `<div class="pgp" style="${o.sh?'filter:drop-shadow(1px 3px 3px #0006);':''}${o.r?`transform:rotate(${o.r}deg);`:''}"><div class="pgi" style="clip-path:url(#${picClip(o.s)});background:url('${u}') center/cover"></div>${o.c?`<svg viewBox="0 0 100 100" preserveAspectRatio="none"><path d="${q[1]}" fill="none" stroke="#${o.c}" stroke-width="${Math.max(1.5,(o.w||3)*.45)}" vector-effect="non-scaling-stroke" stroke-linejoin="round"/></svg>`:''}</div>`};
 const grid=(t,list)=>{const h=E('h5');h.textContent=t;const g=E('div','shgrid');for(const o of list){const b=E('button','sht','<span></span>');b.insertAdjacentHTML('afterbegin',pv(o));b.querySelector('span').textContent=o.n;b.title=o.n;b.onclick=()=>{close();picApply(o)};g.append(b)}box.append(h,g)};
 {const p=E('p','pgn');p.textContent='الدائرة والنجمة والقلب والسداسي… بتقص الصورة لمربع من النص تلقائيًا عشان الشكل يطلع متناسق (تقدر تضبط القص بعدها).';box.append(p)}
 grid('تصميمات جاهزة',PG_ST);grid('الأشكال',PG_SH.map(([id,q])=>({n:SH[id][0],s:id,q})));
 w.append(box);document.body.append(w)}
function picShapeSet(v){const e=S.sel;if(!e||e.localName!='pic')return;const sp=k1(e,P,'spPr');let gm=k1(sp,A,'prstGeom');if(!gm){gm=frag('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',cur().doc);sp.append(gm)}snap();gm.setAttribute('prst',v);touch();draw()}
function picLineSet(hx,w){const e=S.sel;if(!e||e.localName!='pic')return;const sp=k1(e,P,'spPr');snap();for(const n of kids(sp,A,'ln'))n.remove();
 if(w>0){const f=frag(`<a:ln w="${Math.round(w*PT)}"><a:solidFill><a:srgbClr val="${hx.slice(1)}"/></a:solidFill><a:prstDash val="solid"/></a:ln>`,cur().doc);const ef=k1(sp,A,'effectLst');ef?sp.insertBefore(f,ef):sp.append(f)}touch();draw()}
function picShadow(){const e=S.sel;if(!e||e.localName!='pic')return;const sp=k1(e,P,'spPr'),ex=k1(sp,A,'effectLst');snap();if(ex)ex.remove();else sp.append(frag('<a:effectLst><a:outerShdw blurRad="95250" dist="63500" dir="2700000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst>',cur().doc));touch();draw()}
// ---------- الصور: قص / تدوير / قلب / شفافية ----------
const selD=()=>[...S.stage.children].find(d=>d._e===S.sel);
function sliderDlg(t,min,max,v,unit,live,done){const o=document.querySelector('.sld-dlg');o&&o.remove();
 const w=E('div','sld-dlg','<b></b><div class="sr"><input type="range"><output></output></div><p><button type="button" class="ok">✓ تطبيق</button><button type="button" class="no">إلغاء</button></p>');
 w.querySelector('b').textContent=t;const r=w.querySelector('input'),out=w.querySelector('output');r.min=min;r.max=max;r.value=v;out.textContent=v+unit;
 r.oninput=()=>{out.textContent=r.value+unit;live(+r.value)};
 w.querySelector('.ok').onclick=()=>{const x=+r.value;w.remove();done(x)};w.querySelector('.no').onclick=()=>{w.remove();live(null)};
 w.addEventListener('pointerdown',e=>e.stopPropagation());document.body.append(w)}
function rotBy(a){const x=S.sel&&xf(S.sel);if(!x)return;snap();x.setAttribute('rot',(((+x.getAttribute('rot')||0)+a*60000)%21600000+21600000)%21600000);touch();draw()}
function angleDlg(){const x=S.sel&&xf(S.sel);if(!x)return;const d=selD();
 sliderDlg('زاوية التدوير',0,360,Math.round((+x.getAttribute('rot')||0)/60000),'°',v=>{if(v==null)draw();else if(d)d.style.transform='rotate('+v+'deg)'},v=>{snap();x.setAttribute('rot',(Math.round(v)%360)*60000);touch();draw()})}
function flip(a){const x=S.sel&&xf(S.sel);if(!x)return;snap();x.getAttribute(a)=='1'?x.removeAttribute(a):x.setAttribute(a,'1');touch();draw()}
function alphaDlg(){const e=S.sel,bl=e&&all(e,A,'blip')[0];if(!bl)return;const al=all(bl,A,'alphaModFix')[0],t0=al?Math.round(100-(+al.getAttribute('amt')||0)/1000):0,im=selD()&&selD().querySelector('img');
 sliderDlg('شفافية الصورة',0,100,t0,'%',v=>{if(v==null)draw();else if(im)im.style.opacity=1-v/100},v=>{snap();all(bl,A,'alphaModFix').forEach(n=>n.remove());
  if(v>0)bl.insertBefore(frag('<a:alphaModFix amt="'+Math.round((100-v)*1000)+'"/>',cur().doc),k1(bl,A,'extLst')||null);touch();draw()})}
function applyCrop(e,c){const s=cur(),old=all(e,A,'srcRect')[0],o=old?['l','t','r','b'].map(k=>(+old.getAttribute(k)||0)/1e5):[0,0,0,0],g=geo(e,s);if(!g)return;
 const vw0=Math.max(.01,1-o[0]-o[2]),vh0=Math.max(.01,1-o[1]-o[3]),Wf=g.w/vw0,Hf=g.h/vh0,X0=g.x-o[0]*Wf,Y0=g.y-o[1]*Hf,vw=1-c[0]-c[2],vh=1-c[1]-c[3];
 snap();setGeo(e,{x:X0+c[0]*Wf,y:Y0+c[1]*Hf,w:Wf*vw,h:Hf*vh});old&&old.remove();
 if(c.some(v=>v>.0005)){const bf=k1(e,P,'blipFill'),bl=k1(bf,A,'blip');bl.after(frag('<a:srcRect l="'+Math.round(c[0]*1e5)+'" t="'+Math.round(c[1]*1e5)+'" r="'+Math.round(c[2]*1e5)+'" b="'+Math.round(c[3]*1e5)+'"/>',s.doc))}
 touch();draw()}
async function cropDlg(){const e=S.sel;if(!e||e.localName!='pic')return;const bl=all(e,A,'blip')[0],u=bl&&await imgUrl(cur(),bl.getAttributeNS(R,'embed'));if(!u)return X.toast('تعذر فتح الصورة');
 const sr=all(e,A,'srcRect')[0],c=['l','t','r','b'].map(k=>sr?(+sr.getAttribute(k)||0)/1e5:0);
 const w=E('div','crp-dlg','<b>✂ قص الصورة</b><div class="crs"><div class="cri"><img alt=""><div class="crb"><i data-h="nw"></i><i data-h="n"></i><i data-h="ne"></i><i data-h="w"></i><i data-h="e"></i><i data-h="sw"></i><i data-h="s"></i><i data-h="se"></i></div></div></div><p><button type="button" class="ok">✓ تطبيق</button><button type="button" class="rs">↺ إعادة</button><button type="button" class="no">إلغاء</button></p>');
 const im=w.querySelector('img'),box=w.querySelector('.crb');im.src=u;
 const paint=()=>{box.style.left=c[0]*100+'%';box.style.top=c[1]*100+'%';box.style.right=c[2]*100+'%';box.style.bottom=c[3]*100+'%'};paint();
 const M=.04;box.onpointerdown=ev=>{const h=ev.target.dataset&&ev.target.dataset.h||'m';ev.preventDefault();ev.stopPropagation();box.setPointerCapture(ev.pointerId);
  const r=im.getBoundingClientRect(),x0=ev.clientX,y0=ev.clientY,c0=c.slice();
  box.onpointermove=m=>{const dx=(m.clientX-x0)/r.width,dy=(m.clientY-y0)/r.height,cl=(v,a,b)=>Math.max(a,Math.min(b,v));
   if(h=='m'){const w0=1-c0[0]-c0[2],h0=1-c0[1]-c0[3];c[0]=cl(c0[0]+dx,0,1-w0);c[2]=1-w0-c[0];c[1]=cl(c0[1]+dy,0,1-h0);c[3]=1-h0-c[1]}
   else{if(h.includes('w'))c[0]=cl(c0[0]+dx,0,1-c0[2]-M);if(h.includes('e'))c[2]=cl(c0[2]-dx,0,1-c0[0]-M);if(h.includes('n'))c[1]=cl(c0[1]+dy,0,1-c0[3]-M);if(h.includes('s'))c[3]=cl(c0[3]-dy,0,1-c0[1]-M)}paint()};
  box.onpointerup=box.onpointercancel=()=>{box.onpointermove=null}};
 w.querySelector('.ok').onclick=()=>{w.remove();applyCrop(e,c)};w.querySelector('.rs').onclick=()=>{c.fill(0);paint()};w.querySelector('.no').onclick=()=>w.remove();
 w.addEventListener('pointerdown',ev=>ev.stopPropagation());document.body.append(w)}
function rot(){const x=S.sel&&xf(S.sel);if(!x)return;snap();x.setAttribute('rot',((+x.getAttribute('rot')||0)+900000)%21600000);touch();draw()}
function zord(top){const e=S.sel;if(!e)return;snap();const t=tree(cur());if(top)t.append(e);else t.insertBefore(e,[...t.children].find(c=>!/^(nvGrpSpPr|grpSpPr)$/.test(c.localName)));touch();draw()}
function lineSet(hx,w){const e=S.sel;if(!e||e.localName!='sp')return;const sp=k1(e,P,'spPr');snap();for(const c of kids(sp,A,'ln'))c.remove();
 const f=frag(w==0?'<a:ln><a:noFill/></a:ln>':`<a:ln w="${w*PT}"><a:solidFill><a:srgbClr val="${hx.slice(1)}"/></a:solidFill></a:ln>`,e.ownerDocument),ef=k1(sp,A,'effectLst');ef?sp.insertBefore(f,ef):sp.append(f);touch();draw()}
function shadow(){const e=S.sel;if(!e||e.localName!='sp')return;const sp=k1(e,P,'spPr'),ex=k1(sp,A,'effectLst');snap();
 if(ex)ex.remove();else sp.append(frag('<a:effectLst><a:outerShdw blurRad="63500" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>',e.ownerDocument));touch();draw()}
// ---------- محاذاة وتوزيع العناصر بمسافات متساوية (اختيار متعدد من قائمة داخل الشاشة) ----------
function distDlg(){const old=document.querySelector('.pe-ds');if(old){old.remove();S.stage&&[...S.stage.children].forEach(d=>d.classList.remove('pe-ms'));return}
 commit();const ms=new Set();
 const items=()=>[...tree(cur()).children].filter(e=>e.namespaceURI==P&&/^(sp|pic|graphicFrame)$/.test(e.localName)&&geo(e,cur())&&!(all(e,P,'cNvPr')[0]&&all(e,P,'cNvPr')[0].getAttribute('hidden')=='1'));
 if(S.sel&&items().includes(S.sel))ms.add(S.sel);
 const lab=e=>{const t=all(e,A,'t').map(x=>x.textContent).join(' ').trim(),c=all(e,P,'cNvPr')[0],n=e.localName;return(n=='pic'?'🖼 ':n=='graphicFrame'?'▦ ':'🔷 ')+(t?t.slice(0,28):(c&&c.getAttribute('name'))||'عنصر')};
 const live=()=>{const it=items();for(const e of[...ms])if(!it.includes(e))ms.delete(e);return[...ms]};
 const hl=()=>{const a=live();for(const d of S.stage.children)d.classList.toggle('pe-ms',!!d._e&&a.includes(d._e))};
 const w=E('div','pe-ds','<div class="pds-h"><b>📐 محاذاة وتوزيع</b><button type="button" aria-label="إغلاق">✕</button></div><div class="pds-l"></div><div class="pds-n"></div>'),list=w.querySelector('.pds-l'),note=w.querySelector('.pds-n');
 const info=()=>{const n=live().length;note.textContent=n?'محدد '+n+' عنصر':'اختار العناصر اللي عاوز تظبطها'};
 const build=()=>{list.innerHTML='';const all0=E('button','pds-all','☑ تحديد الكل / إلغاء الكل');all0.type='button';all0.onclick=()=>{const it=items();if(live().length==it.length)ms.clear();else it.forEach(e=>ms.add(e));build()};list.append(all0);
  for(const e of items()){const r=E('label'),c=E('input');c.type='checkbox';c.checked=ms.has(e);c.onchange=()=>{c.checked?ms.add(e):ms.delete(e);hl();info()};const sp=E('span');sp.textContent=lab(e);r.append(c,sp);list.append(r)}hl();info()};
 const grp=(t,...els)=>{const g=E('div','pds-g'),b=E('b');b.textContent=t;g.append(b,...els);w.append(g)},
  btn=(t,f,title)=>{const b=E('button','',t);b.type='button';if(title)b.title=title;b.onclick=f;return b};
 const gs=()=>live().map(e=>({e,g:geo(e,cur())})),fin=()=>{touch();draw();hl()};
 const md=E('select');md.setAttribute('aria-label','طريقة التوزيع');md.innerHTML='<option value="span">بين أول عنصر وآخر عنصر (3+)</option><option value="slide">على عرض/ارتفاع الشريحة كلها (2+)</option>';
 const dist=ax=>{const mode=md.value,a=gs();if(a.length<(mode=='slide'?2:3))return X.toast(mode=='slide'?'اختار عنصرين على الأقل':'اختار 3 عناصر على الأقل، أو غيّر الطريقة لـ "على الشريحة كلها"');
  const H=ax=='x',k=H?'x':'y',L=H?'w':'h',tot=H?S.W:S.H;a.sort((p,q)=>p.g[k]-q.g[k]);const sum=a.reduce((t,p)=>t+p.g[L],0);let start,gap;
  if(mode=='slide'){gap=(tot-sum)/(a.length+1);start=gap}else{const f=a[0].g[k],l=a[a.length-1].g;gap=(l[k]+l[L]-f-sum)/(a.length-1);start=f}
  snap();let pos=start;for(const p of a){setGeo(p.e,{...p.g,[k]:pos});pos+=p.g[L]+gap}fin();X.toast('اتوزّعت العناصر بمسافات متساوية')};
 const align=v=>{const a=gs();if(!a.length)return X.toast('اختار عنصر واحد على الأقل');let b={x:0,y:0,w:S.W,h:S.H};
  if(a.length>1){const x0=Math.min(...a.map(p=>p.g.x)),y0=Math.min(...a.map(p=>p.g.y)),x1=Math.max(...a.map(p=>p.g.x+p.g.w)),y1=Math.max(...a.map(p=>p.g.y+p.g.h));b={x:x0,y:y0,w:x1-x0,h:y1-y0}}
  snap();for(const p of a){const g=p.g,n={...g};if(v=='l')n.x=b.x;else if(v=='c')n.x=b.x+(b.w-g.w)/2;else if(v=='r')n.x=b.x+b.w-g.w;else if(v=='t')n.y=b.y;else if(v=='m')n.y=b.y+(b.h-g.h)/2;else n.y=b.y+b.h-g.h;setGeo(p.e,n)}fin()};
 const same=k=>{const a=live();if(a.length<2)return X.toast('اختار عنصرين على الأقل (الأول اللي اخترته هو المرجع)');const ref=geo(a[0],cur());snap();
  for(const e of a.slice(1)){const g=geo(e,cur());setGeo(e,{...g,w:k!='h'?ref.w:g.w,h:k!='w'?ref.h:g.h})}fin()};
 grp('توزيع بمسافات متساوية',md,btn('↔ أفقي',()=>dist('x'),'توزيع أفقي'),btn('↕ رأسي',()=>dist('y'),'توزيع رأسي'));
 grp('محاذاة (لبعض، أو للشريحة لو عنصر واحد)',btn('⬅',()=>align('l'),'يسار'),btn('⬌',()=>align('c'),'وسط أفقي'),btn('➡',()=>align('r'),'يمين'),btn('⬆',()=>align('t'),'فوق'),btn('↕',()=>align('m'),'وسط رأسي'),btn('⬇',()=>align('b'),'تحت'));
 grp('نفس المقاس (على أول عنصر اخترته)',btn('عرض',()=>same('w')),btn('ارتفاع',()=>same('h')),btn('الاتنين',()=>same('wh')));
 w.querySelector('.pds-h button').onclick=()=>{ms.clear();w.remove();[...S.stage.children].forEach(d=>d.classList.remove('pe-ms'))};
 document.body.append(w);build()}
function alignTo(v){const e=S.sel,g=e&&geo(e,cur());if(!g)return;snap();setGeo(e,{...g,...{l:{x:0},c:{x:(S.W-g.w)/2},r:{x:S.W-g.w},t:{y:0},m:{y:(S.H-g.h)/2},b:{y:S.H-g.h}}[v]});touch();draw()}
function pdivs(){if(!ensureEdit())return[];restore();const g=getSelection(),r=g.rangeCount?g.getRangeAt(0):null,t=S.ed.el,ch=[...t.children];if(!r)return ch;
 const f=n=>{while(n&&n.parentNode!==t)n=n.parentNode;return n},a=ch.indexOf(f(r.startContainer)),b=ch.indexOf(f(r.endContainer));return ch.slice(Math.max(0,a),Math.max(a,b)+1)}
const bullet=k=>{for(const d of pdivs())d.dataset.bu=d.dataset.bu==k?'none':k},lsp=v=>{for(const d of pdivs())d.style.lineHeight=v};
const blank1=x=>{const tb=k1(x,A,'txBody');for(const p of kids(tb,A,'p'))p.remove();tb.append(frag('<a:p><a:endParaRPr lang="en-US" sz="1400"/></a:p>',x.ownerDocument));x.removeAttribute('hMerge');x.removeAttribute('gridSpan')};
function setCellFill(h){const c=S.cell;if(!c||!c.isConnected||!c._tc)return alert('اضغط على خلية جوه الجدول الأول.');const tc=c._tc;commit();snap();const d=tc.ownerDocument;let pr=k1(tc,A,'tcPr');if(!pr){pr=d.createElementNS(A,'a:tcPr');tc.append(pr)}
 for(const l of ['solidFill','noFill','gradFill','pattFill','blipFill'])kids(pr,A,l).forEach(n=>n.remove());
 const f=frag('<a:solidFill><a:srgbClr val="'+String(h).replace('#','').toUpperCase()+'"/></a:solidFill>',d);let ref=null;for(const n of pr.children)if(/^(headers|extLst)$/.test(n.localName)){ref=n;break}
 ref?pr.insertBefore(f,ref):pr.append(f);S.cell=null;touch();draw()}
function tabOp(f){const c=S.cell;if(!c||!c.isConnected||!c._tc)return alert('اضغط على خلية جوه الجدول الأول.');const tc=c._tc;commit();snap();const tr=tc.parentNode,t=tr.parentNode;let g=t;while(g&&g.localName!='graphicFrame')g=g.parentNode;
 const ex=g&&k1(k1(g,P,'xfrm'),A,'ext'),fr=(dx,dy)=>{if(ex){ex.setAttribute('cx',Math.max(1,+ex.getAttribute('cx')+dx));ex.setAttribute('cy',Math.max(1,+ex.getAttribute('cy')+dy))}};
 f(tc,tr,t,kids(tr,A,'tc').indexOf(tc),fr);S.cell=null;touch();draw()}
const eqCols=(tc,tr,t)=>{const gc=kids(k1(t,A,'tblGrid'),A,'gridCol');if(gc.length<2)return;const tot=gc.reduce((s,c)=>s+(+c.getAttribute('w')||0),0),w=Math.round(tot/gc.length);gc.forEach(c=>c.setAttribute('w',w));X.toast('الأعمدة اتوزّعت بالتساوي')};
const eqRows=(tc,tr,t)=>{const rs=kids(t,A,'tr');if(rs.length<2)return;const tot=rs.reduce((s,r)=>s+(+r.getAttribute('h')||0),0),h=Math.round(tot/rs.length);rs.forEach(r=>r.setAttribute('h',h));X.toast('الصفوف اتوزّعت بالتساوي')};
const addRow=(tc,tr,t,i,fr)=>{const n=tr.cloneNode(true);kids(n,A,'tc').forEach(blank1);tr.after(n);fr(0,+tr.getAttribute('h'))},
 delRow=(tc,tr,t,i,fr)=>{if(kids(t,A,'tr').length<2)return;fr(0,-tr.getAttribute('h'));tr.remove()},
 addCol=(tc,tr,t,i,fr)=>{for(const r of kids(t,A,'tr')){const c=kids(r,A,'tc')[i];if(!c)continue;const n=c.cloneNode(true);blank1(n);c.after(n)}const gc=kids(k1(t,A,'tblGrid'),A,'gridCol')[i];gc.after(gc.cloneNode(true));fr(+gc.getAttribute('w'),0)},
 delCol=(tc,tr,t,i,fr)=>{if(kids(tr,A,'tc').length<2)return;for(const r of kids(t,A,'tr')){const c=kids(r,A,'tc')[i];c&&c.remove()}const gc=kids(k1(t,A,'tblGrid'),A,'gridCol')[i];fr(-gc.getAttribute('w'),0);gc.remove()},
 merge=(tc,tr,t,i)=>{const nx=kids(tr,A,'tc')[i+1];if(!nx)return;tc.setAttribute('gridSpan',(+tc.getAttribute('gridSpan')||1)+1);nx.setAttribute('hMerge','1')};
async function repl(){const a=await raAsk('ابحث عن:','');if(!a)return;const b=await raAsk('استبدل بـ:','');if(b==null)return;commit();snap();let n=0;
 for(const s of S.slides)for(const t of all(s.doc,A,'t'))if(t.textContent.includes(a)){t.textContent=t.textContent.split(a).join(b);s.t=1;n++}
 X.O().dirty=1;draw();thumbs();X.toast(n?'اتبدّل في '+n+' موضع':'مفيش نتايج')}
async function findInSlides(){const q=(await raAsk('اكتب كلمة أو جملة للبحث في كل الشرائح:',''))?.trim();if(!q)return;commit();const start=S.cur;for(let step=0;step<S.slides.length;step++){const i=(start+step)%S.slides.length,sl=S.slides[i],els=all(tree(sl),P,'sp').concat(all(tree(sl),P,'graphicFrame'),all(tree(sl),P,'pic'));const hit=els.find(e=>all(e,A,'t').some(t=>t.textContent.toLocaleLowerCase().includes(q.toLocaleLowerCase())));if(hit){go(i);S.sel=hit;draw();X.toast('تم العثور على النص في الشريحة '+(i+1));return}}X.toast('ملقيناش النص ده في أي شريحة')}

// ---------- save ----------
// builds the .pptx from the current state (does not touch the UI)
async function build(){
 {const put=new Map(S.put),del=new Set(S.del),pres=ps(await S.rd('ppt/presentation.xml')),prel=ps(await S.rd('ppt/_rels/presentation.xml.rels')),ct=ps(await S.rd('[Content_Types].xml')),lst=all(pres,P,'sldIdLst')[0],by={};
  for(const x of all(lst,P,'sldId'))by[x.getAttributeNS(R,'id')]=x;let mid=Math.max(255,...Object.values(by).map(x=>+x.getAttribute('id')));
  for(const g of S.gone){const r=all(prel,PR,'Relationship').find(x=>x.getAttribute('Id')==g.rid);r&&r.remove();const o=all(ct,CT,'Override').find(x=>x.getAttribute('PartName')=='/'+g.path);o&&o.remove()}
  while(lst.firstChild)lst.firstChild.remove();
  for(const s of S.slides){let e=by[s.rid];if(s.isNew){const rid=nextRid(prel),r=prel.createElementNS(PR,'Relationship');r.setAttribute('Id',rid);r.setAttribute('Type',R+'/slide');r.setAttribute('Target',s.path.replace(/^ppt\//,''));prel.documentElement.append(r);
    const o=ct.createElementNS(CT,'Override');o.setAttribute('PartName','/'+s.path);o.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.presentationml.slide+xml');ct.documentElement.append(o);
    e=pres.createElementNS(P,'p:sldId');e.setAttribute('id',++mid);e.setAttributeNS(R,'r:id',rid)}
   if(!e)continue;lst.append(e);if(s.t||s.isNew){put.set(s.path,xml(s.doc));put.set(relPath(s.path),xml(s.rel))}}
  if(S.st)for(const t of all(pres,'*','sectionLst')){const p=t.parentNode;(p&&p.localName=='ext'?p:t).remove()}
  for(const ext of S.exts)if(!all(ct,CT,'Default').some(d=>d.getAttribute('Extension').toLowerCase()==ext)){const d=ct.createElementNS(CT,'Default');d.setAttribute('Extension',ext);d.setAttribute('ContentType',(S.mt&&S.mt[ext])||'image/'+ext);ct.documentElement.insertBefore(d,ct.documentElement.firstChild)}
  if(S.size){const z=all(pres,P,'sldSz')[0];z.setAttribute('cx',S.size[0]);z.setAttribute('cy',S.size[1]);z.removeAttribute('type')}
  for(const n of S.dd||[])put.set(n,xml(S.docs[n]));
  await notesFlush(put,pres,prel,ct);
  put.set('ppt/presentation.xml',xml(pres));put.set('ppt/_rels/presentation.xml.rels',xml(prel));put.set('[Content_Types].xml',xml(ct));
  return X.zipBuild(S.z,del,put)}}
// keeps the working copy in the app library (originals are never overwritten) so you can continue later
const persist=async out=>X.saveBuf(await out.arrayBuffer(),' (معدّل)',{ls:S.cur,ed:1});
// ---------- export all slides as a PDF copy ----------
async function exportPdf(){
 commit();let host=null;
 try{
  X.busy(true,'جاري تجهيز PDF…');
  const loadLib=async(name,local,cdn)=>{
   if(window[name])return window[name];
   await new Promise((ok,no)=>{const q=document.createElement('script');q.src=local;q.onload=ok;q.onerror=()=>{q.remove();const z=document.createElement('script');z.src=cdn;z.onload=ok;z.onerror=()=>no(new Error('تعذر تحميل مكتبة تصدير PDF'));document.head.append(z)};document.head.append(q)});
   if(!window[name])throw new Error('مكتبة PDF غير متاحة');return window[name]
  };
  const h2c=await loadLib('html2canvas','lib/html2canvas.min.js','https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js');
  const lib=await loadLib('PDFLib','lib/pdf-lib.min.js','https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js');
  const pdf=await lib.PDFDocument.create();host=E('div');host.style.cssText='position:fixed;left:-20000px;top:0;width:1280px;height:720px;background:#fff;z-index:-1';document.body.append(host);
  for(let i=0;i<S.slides.length;i++){
   host.innerHTML='';const stage=E('div');host.append(stage);paint(S.slides[i],stage,1280/S.W,0);
   const imgs=[...stage.querySelectorAll('img')];await Promise.all(imgs.map(im=>im.decode?im.decode().catch(()=>{}):new Promise(r=>{im.onload=im.onerror=r})));
   await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));
   await bakeClips(stage);
   const canvas=await h2c(stage,{backgroundColor:bgc(S.slides[i]),scale:1,useCORS:true,logging:false});
   const bytes=await new Promise(r=>canvas.toBlob(async b=>r(new Uint8Array(await b.arrayBuffer())),'image/jpeg',.92));
   const jpg=await pdf.embedJpg(bytes),page=pdf.addPage([960,540]);page.drawImage(jpg,{x:0,y:0,width:960,height:540});canvas.width=canvas.height=0;
   X.busy(true,`جاري تصدير الشريحة ${i+1} من ${S.slides.length}…`)
  }
  const out=new Blob([await pdf.save()],{type:'application/pdf'});X.dl(out,(X.O().d.n||'presentation')+'.pdf');X.toast('تم تصدير العرض إلى PDF');
 }catch(e){X.fail('فشل تصدير PDF',e)}finally{host&&host.remove();X.busy(false)}
}

async function save(){commit();const O=X.O();X.busy(true,'جاري الحفظ…');
 try{const out=await build();await persist(out);X.dl(out,O.d.n+'.pptx');O.dirty=0;X.busy(false);X.toast('اتحفظ في التطبيق وجاهز للمشاركة')}
 catch(e){X.fail('فشل الحفظ',e)}}
// writes text that is still being typed into the model without ending the edit (used when the app goes to background / on exit)
function flushSoft(){if(!S)return;if(S.ed){const{e,el,h}=S.ed;if(el.innerHTML!==h){writeText(e,el);S.ed.h=el.innerHTML;cur().t=1;X.O().dirty=1}}
 for(const c of S.stage.querySelectorAll('td'))if(c._tc&&c.textContent!==c._o){writeCell(c._tc,c.textContent);c._o=c.textContent;cur().t=1;X.O().dirty=1}}
async function autosave(force){const s=S;if(!s)return;if(s.sp)await s.sp;if(S!==s)return;
 if(force)flushSoft();else if(s.ed||(document.activeElement&&document.activeElement.closest&&document.activeElement.closest('.pe-stage td')))return;
 const O=X.O();if(!O||!O.dirty)return;const v=s.dv;
 s.sp=(async()=>{try{const out=await build();await persist(out);if(s.dv!==v)O.dirty=1}catch(e){console.warn('autosave',e)}})();
 try{await s.sp}finally{s.sp=null}}

// ---------- entry ----------
(()=>{const c=E('style');c.textContent=`.pe-host{background:#eef1f7;direction:ltr}.pe-sc{flex:1;min-height:0;display:flex;overflow:auto;padding:0;touch-action:pan-x pan-y}
.pe-stage{position:relative;flex:none;margin:auto;box-shadow:0 2px 14px #0003;overflow:hidden;user-select:none;-webkit-user-select:none}.pe-s{position:absolute;box-sizing:border-box;overflow:hidden;cursor:move;touch-action:none}
.pe-s img{width:100%;height:100%;display:block;pointer-events:none}.pe-s.sel{outline:2px solid #2563eb;overflow:visible}
.pe-t{width:100%;height:100%;box-sizing:border-box;outline:none;overflow:hidden;word-break:break-word}.pe-s.ed{cursor:text}.pe-s.ed .pe-t{user-select:text;-webkit-user-select:text}
.pe-tb{border-collapse:collapse;table-layout:fixed;width:100%}.pe-tb td{border:1px solid #888;padding:2px 4px;vertical-align:top;word-break:break-word;outline:none}
.pe-h{position:absolute;right:-10px;bottom:-10px;width:20px;height:20px;border-radius:50%;background:#2563eb;border:2px solid #fff;cursor:nwse-resize;z-index:2}
.pe-th{display:flex;gap:8px;overflow-x:auto;padding:8px;background:#fff;border-top:1px solid #d5dbe8;flex:none}.pe-tn{position:relative;flex:none;border:2px solid transparent;border-radius:6px;overflow:hidden;cursor:pointer;background:#fff}
.pe-tn.on{border-color:#2563eb}.pe-tn>b{position:absolute;left:3px;top:2px;font-size:11px;background:#2563ebcc;color:#fff;padding:0 5px;border-radius:8px;z-index:3}.pe-tn .pe-stage{box-shadow:none;pointer-events:none}.pe-drag{position:absolute;right:3px;bottom:3px;z-index:4;width:28px;height:26px;border:1px solid #cbd5e1;border-radius:7px;background:#ffffffed;color:#334155;font-size:19px;line-height:1;touch-action:none;display:grid;place-items:center;padding:0}.pe-tn.dragging{opacity:.65;transform:scale(.96);border-color:#2563eb}
.pe-pl{position:fixed;inset:0;z-index:300;background:#000;overflow:hidden;touch-action:none}.pe-pw{position:absolute;inset:0;display:flex;align-items:center;justify-content:center}.pe-pw>div{position:relative;overflow:hidden;flex:none}
.pe-px{position:absolute;top:calc(8px + env(safe-area-inset-top));right:10px;z-index:5;background:#0008;color:#fff;border:0;border-radius:18px;padding:6px 12px;font-size:16px}
.pe-nt{flex:none;display:block;width:100%;height:120px;box-sizing:border-box;border:0;border-top:1px solid #d5dbe8;padding:8px 10px;font-size:16px;line-height:1.5;font-family:inherit;resize:none;background:#fff;color:#111}.pe-nt[hidden]{display:none}
.pe-nb{position:absolute;top:calc(8px + env(safe-area-inset-top));left:10px;z-index:5;background:#0008;color:#fff;border:0;border-radius:18px;padding:6px 12px;font-size:16px}.pe-np{position:absolute;left:0;right:0;bottom:0;max-height:36%;overflow:auto;z-index:6;background:#000d;color:#fff;padding:10px 14px 26px;font-size:18px;line-height:1.5;white-space:pre-wrap}.pe-np[hidden]{display:none}
.pe-pc{position:absolute;bottom:calc(8px + env(safe-area-inset-bottom));left:50%;transform:translateX(-50%);z-index:5;color:#fffa;font-size:13px;pointer-events:none}.pe-pn{display:flex;gap:6px;overflow-x:auto;padding:8px;background:#fff;border-bottom:1px solid #d5dbe8;flex:none;scrollbar-width:none;align-items:center}.pe-pn button,.pe-pn select,.pe-pn label{flex:none;min-width:46px;height:44px;padding:0 12px;font-size:16px;border-radius:12px;border:1px solid #cfd6e6;background:#f4f6fb;color:#111;display:inline-flex;align-items:center;gap:4px;font-family:inherit;white-space:nowrap}.pe-pn button:active{background:#dbe5fb}.pe-pn i.sep{flex:none;width:1px;height:28px;background:#cfd6e6;margin:0 2px}.pe-mp{position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);width:40px;height:40px;padding:0;border-radius:50%;background:#000b;color:#fff;border:2px solid #fffa;display:grid;place-items:center;font-size:16px;z-index:3;touch-action:manipulation}.pe-mp i{font-style:normal}.pe-hint{color:#667;font-size:14px;padding:0 6px;white-space:nowrap}.pe-tabs{display:flex;background:#f1f4fa;border-bottom:1px solid #d5dbe8;flex:none}.pe-tabs button{flex:1;padding:8px 2px 6px;border:0;border-bottom:3px solid transparent;background:none;color:#667;font-size:12px;display:flex;flex-direction:column;align-items:center}.pe-tabs i{font-style:normal;font-size:20px}.pe-tabs button.on{color:#2563eb;font-weight:700;border-bottom-color:#2563eb;background:#fff}.pe-t{counter-reset:pn}.pe-t>div[data-bu="•"]::before{content:"•";margin-inline-end:.45em}.pe-t>div[data-bu="1"]{counter-increment:pn}.pe-t>div[data-bu="1"]::before{content:counter(pn) ". ";margin-inline-end:.3em}`;document.head.append(c)})();

// view-mode slideshow: loads the deck read-only and plays it (nothing is edited or saved)
window.pptPlay=async()=>{if(S)return;try{X.busy(true,'جاري تجهيز العرض…');const s=await load();s.stage=document.createElement('div');S=s;X.busy(false);
 play(0,()=>{if(S===s){Object.values(s.blobs).forEach(u=>URL.revokeObjectURL(u));S=null}})}catch(e){S=null;X.fail('تعذر تشغيل العرض',e)}};
window.pptEdit=async()=>{try{X.busy(true,'جاري تجهيز المحرر…');S=await load();{const o=X.O();if(o.d&&o.d.ls!=null)S.cur=Math.max(0,Math.min(+o.d.ls||0,S.slides.length-1))}S.media=[];try{if(X.media)S.media=(await X.media.init(X.O().buf)).items}catch(e){console.warn('media scan',e)}X.busy(false)}catch(e){return X.fail('فشل فتح المحرر',e)}
 const h=$('#ofm');h.className='pe-host';h.innerHTML='<div class="pe-sc"><div class="pe-stage"></div></div><div class="pe-th"></div>';
 S.sc=h.firstChild;S.stage=S.sc.firstChild;S.th=h.lastChild;bind();pinch();
 const B=(t,f,title,lbl)=>({t,f,title,lbl}),col=(t,title,f,init)=>X.colorBtn({label:t,title,init:init||'#1e40af',theme:S.theme,onPick:f});
 const ms=(label,opts)=>{const x=E('select');x.setAttribute('aria-label',label);x.innerHTML=opts.map(o=>`<option value="${o[0]}">${o[1]}</option>`).join('');return x},
  tsel=ms('الانتقال',Object.keys(TRS).map(k=>[k,TRS[k][0]])),dsel=ms('اتجاه الانتقال',TRD),ssel=ms('سرعة الانتقال',[['fast','⚡ سريع'],['med','سرعة متوسطة'],['slow','بطيء']]),asel=ms('التقدم التلقائي',[[0,'التقدم: بالضغط'],[2000,'تلقائي بعد 2 ث'],[3000,'تلقائي بعد 3 ث'],[5000,'تلقائي بعد 5 ث'],[10000,'تلقائي بعد 10 ث']]);
 const trRead=()=>({t:tsel.value,dir:dsel.value,spd:ssel.value,adv:+asel.value}),
  trSync=()=>{const o=getTr(cur());tsel.value=o.t;dsel.value=o.dir;ssel.value=o.spd;if(![...asel.options].some(q=>+q.value==o.adv)){const q=E('option');q.value=o.adv;q.textContent='تلقائي بعد '+o.adv/1000+' ث';asel.append(q)}asel.value=o.adv;dsel.hidden=!TRS[o.t][2]};
 const trChg=()=>{dsel.hidden=!TRS[tsel.value][2];setTr(cur(),trRead())};[tsel,dsel,ssel,asel].forEach(q=>q.onchange=trChg);
 S.trs={set value(_){trSync()}};
 const mk=(label,opts,f)=>{const s=E('select');s.setAttribute('aria-label',label);s.innerHTML=`<option value="">${label}</option>`+opts.map(o=>`<option value="${o[0]}">${o[1]}</option>`).join('');s.onchange=()=>{const v=s.value;s.value='';v&&f(v)};return s};
 const szIn=S.szIn=(()=>{const n=document.createElement('input');n.type='number';n.min=1;n.max=999;n.step=1;n.inputMode='numeric';n.placeholder='الحجم';n.title='حجم الخط بالرقم (pt)';n.style.cssText='flex:none;width:64px;height:44px;padding:0 6px;font-size:16px;border-radius:12px;border:1px solid #cfd6e6;background:#f4f6fb;color:#111;text-align:center';const ap=()=>{if(n.value!==''){sizeTo(n.value)}};n.onchange=ap;n.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();ap()}};return n})();
  const fontSel=mk('الخط',['Arial','Calibri','Tahoma','Times New Roman','Courier New','Georgia','Verdana','Traditional Arabic','Simplified Arabic','Sakkal Majalla','Arabic Typesetting','Segoe UI','Cairo','Amiri','Tajawal','Almarai','Noto Naskh Arabic','Noto Kufi Arabic','IBM Plex Sans Arabic','Lateef','Scheherazade New'].map(f=>[f,f]),v=>fmt('fontName',v)),
  charSpSel=mk('تباعد الحروف',[['0','عادي'],['0.5','ضيق قليلًا'],['1','1 pt'],['2','2 pt'],['3','3 pt']],charSpace),
  marginSel=mk('هوامش مربع النص',[['0.05','ضيقة'],['0.1','عادية'],['0.15','متوسطة'],['0.2','واسعة']],textMargin),
  lspSel=mk('تباعد الأسطر',[['1','1'],['1.15','1.15'],['1.5','1.5'],['2','2']],lsp),lwSel=mk('سُمك الحد',[['0','بدون'],['1','1'],['2','2'],['4','4'],['6','6']],v=>{S.lw=+v;lineSet(S.lc,S.lw)}),
  alSel=mk('محاذاة العنصر',[['l','⬅ يسار'],['c','⬌ وسط'],['r','➡ يمين'],['t','⬆ فوق'],['m','↕ منتصف'],['b','⬇ تحت']],alignTo);
 const TB={txt:['✏️','نص'],ins:['➕','إدراج'],obj:['🔷','عنصر'],sld:['🎞','شرايح'],shw:['▶','عرض'],nts:['📝','ملاحظات']},hint=t=>({el:E('span','pe-hint',t)});
 const tplSel=mk('🧩 قوالب',Object.keys(TPL).map(k=>[k,TPL[k].l]),addTpl),szSel=mk('📐 المقاس',[['169','16:9 عرضي'],['43','4:3'],['916','9:16 طولي (تليفون)']],resize);
 const palettes={blue:{l:'🔵 أزرق احترافي',map:{'1e40af':'2563eb','1e3a8a':'1d4ed8','0f766e':'0891b2','fbbf24':'60a5fa'}},green:{l:'🟢 أخضر هادئ',map:{'1e40af':'166534','1e3a8a':'14532d','0f766e':'0f766e','fbbf24':'a3e635'}},purple:{l:'🟣 بنفسجي',map:{'1e40af':'6d28d9','1e3a8a':'4c1d95','0f766e':'7e22ce','fbbf24':'c4b5fd'}},warm:{l:'🟠 دافئ',map:{'1e40af':'c2410c','1e3a8a':'7c2d12','0f766e':'b45309','fbbf24':'fbbf24'}}};
 function applyPalette(k){const p=palettes[k];if(!p)return;commit();snap();for(const sl of S.slides){for(const n of all(sl.doc,A,'srgbClr')){const v=(n.getAttribute('val')||'').toLowerCase();if(p.map[v])n.setAttribute('val',p.map[v]);}sl.t=1;}touch();draw();thumbs();X.toast('تم تطبيق سمة الألوان على العرض كله')}
 const themeSel=mk('🎨 سمة التصميم',Object.keys(palettes).map(k=>[k,palettes[k].l]),applyPalette);
 const picShapeSel=mk('شكل الصورة',Object.entries(PIC_SHAPES).map(([k,v])=>[k,v]),picShapeSet),picLineSel=mk('سُمك إطار الصورة',[['0','بدون إطار'],['1','1 px'],['2','2 px'],['3','3 px'],['5','5 px'],['8','8 px']],v=>{S.picLineW=+v;picLineSet(S.picLineC||'#2563eb',S.picLineW)});
 const objSecs=()=>{const tbl=S.sel.localName=='graphicFrame',pic=S.sel.localName=='pic',secs={pic:[B('🎭 أشكال وتصميمات',picGallery,null,'أشكال وتصميمات الصورة'),'|',B('✂ قص',cropDlg),B('↺ إلغاء القص',()=>applyCrop(S.sel,[0,0,0,0])),'|',{el:picShapeSel},col('▭','لون إطار الصورة',c=>{S.picLineC=c;picLineSet(c,S.picLineW===undefined?2:S.picLineW)},'#2563eb'),{el:picLineSel},B('◪ ظل الصورة',picShadow),'|',B('⟲ 90°',()=>rotBy(-90),null,'تدوير 90° لليسار'),B('⟳ 90°',()=>rotBy(90),null,'تدوير 90° لليمين'),B('🎚 زاوية',angleDlg),B('↔ قلب',()=>flip('flipH'),null,'قلب أفقي'),B('↕ قلب',()=>flip('flipV'),null,'قلب رأسي'),'|',B('👁 شفافية',alphaDlg)],fmt:[col('🎨','تعبئة',setFill),B('🌈 تدرج',gradDlg,null,'تدرج لوني للتعبئة أو الخلفية'),col('▭','لون الحد',c=>{S.lc=c;lineSet(c,S.lw)}),{el:lwSel},B('◪ ظل',shadow)],
     arr:[B('⟳ تدوير',rot),B('🎚 زاوية',angleDlg),B('⤒ للأمام',()=>zord(1)),B('⤓ للخلف',()=>zord(0)),'|',{el:alSel},B('📐 محاذاة وتوزيع',distDlg,null,'محاذاة وتوزيع العناصر بمسافات متساوية')],anim:animCtl(),
     tbl:[B('＋صف',()=>tabOp(addRow)),B('－صف',()=>tabOp(delRow)),'|',B('＋عمود',()=>tabOp(addCol)),B('－عمود',()=>tabOp(delCol)),'|',B('دمج',()=>tabOp(merge)),B('⇔ أعمدة متساوية',()=>tabOp(eqCols),null,'توزيع عرض الأعمدة بالتساوي'),B('⇕ صفوف متساوية',()=>tabOp(eqRows),null,'توزيع ارتفاع الصفوف بالتساوي'),'|',col('🎨 لون الخلية','لون الخلية المحددة',setCellFill,'#FFF2CC')]};return{tbl,pic,secs}};
 const items={
  txt:()=>[B('<b>B</b>',()=>fmt('bold'),null,'غامق'),B('<i>I</i>',()=>fmt('italic'),null,'مائل'),B('<u>U</u>',()=>fmt('underline'),null,'تسطير'),'|',B('A+',()=>size(1),null,'تكبير الخط'),B('A−',()=>size(-1),null,'تصغير الخط'),{el:szIn},col('A','لون النص',c=>fmt('foreColor',c)),{el:fontSel},{el:charSpSel},{el:lspSel},{el:marginSel},B('🌑 ظل النص',textShadow),'|',B('⬅',()=>fmt('justifyLeft'),null,'محاذاة لليسار'),B('⬌',()=>fmt('justifyCenter'),null,'توسيط'),B('➡',()=>fmt('justifyRight'),null,'محاذاة لليمين'),B('☰',()=>fmt('justifyFull'),null,'ضبط'),B('✨',smartAlign,null,'محاذاة ذكية حسب لغة كل فقرة'),'|',B('•',()=>bullet('•'),null,'قائمة نقطية'),B('1.',()=>bullet('1'),null,'قائمة رقمية')],
  ins:()=>[B('🔤 نص',addText),B('🖼 صورة',addImg),B('🎙 تسجيل صوت للسلايد',recordSlideAudio,'سجّل شرحًا صوتيًا للشريحة الحالية'),B('🎞 صوت/فيديو',addMedia,'إدراج صوت أو فيديو'),B('▦ جدول',addTable),B('📊 مخطط بياني',addChart),'|',B('▭',()=>addShape('rect'),null,'مستطيل'),B('◯',()=>addShape('ellipse'),null,'دائرة'),B('△',()=>addShape('triangle'),null,'مثلث'),B('☆',()=>addShape('star5'),null,'نجمة'),B('🔷 كل الأشكال',shapeGallery,'مكتبة الأشكال'),B('🧩 مخططات جاهزة',dgGallery,'مخططات جاهزة')],
  obj:()=>{if(!S.sel)return[hint('اضغط على أي عنصر في الشريحة'),B('📌 لصق',paste)];
   const{tbl,pic,secs}=objSecs(),sb=(k,t)=>B(t,()=>{S.sub=k;S.panel()}),bk=B('‹ رجوع',()=>{S.sub=null;S.panel()});   if(S.sub&&secs[S.sub]&&(S.sub!='tbl'||tbl))return[bk,'|',...secs[S.sub]];
   return[B('🗑',delShape,null,'حذف العنصر'),B('📋',copy,null,'نسخ'),B('📌',paste,null,'لصق'),'|',sb('fmt','🎨 تنسيق'),sb('arr','📐 ترتيب'),sb('anim','✨ حركة'),...(pic?[B('🎭 أشكال وتصميمات',picGallery,null,'أشكال وتصميمات الصورة'),sb('pic','🖼 الصورة')]:[]),...(tbl?[sb('tbl','▦ جدول')]:[])]},
  sld:()=>[B('＋ جديدة',()=>addSlide(0)),{el:tplSel},B('⧉ تكرار',()=>addSlide(1)),B('🗑 حذف',delSlide),'|',B('⬆',()=>mv(-1),null,'نقل الشريحة لأعلى'),B('⬇',()=>mv(1),null,'نقل الشريحة لأسفل'),'|',{el:themeSel},col('🖼','خلفية',bgSet),B('🌈 تدرج',gradDlg,null,'تدرج لوني للتعبئة أو الخلفية'),B('🔎 بحث في الشرائح',findInSlides),B('🔍 استبدال',repl),{el:szSel}],
  shw:()=>{trSync();return[B('▶ تشغيل',()=>play(S.cur)),{el:tsel},{el:dsel},{el:ssel},{el:asel},B('🎬 للكل',()=>{const o=trRead();snap();for(const s of S.slides)setTr(s,o,1);X.toast('الانتقال اتطبّق على كل الشرايح')})]},
  nts:()=>[hint('ملاحظات المتحدث للشريحة دي — بتظهر في العرض بزرار 📝'),B('🗑 مسح',ntClear)]};
 const pn=E('div','pe-pn'),tb=E('div','pe-tabs');S.tab='ins';S.zm=1;

 // ---------- vertical tools list (☰ الأدوات): every tool, grouped and labelled, one under another ----------
 const VM={open:false,sec:new Set(['txt']),el:null};
 const hasLetters=t=>/[A-Za-z؀-ۿ]/.test(String(t).replace(/<[^>]*>/g,''));
 const GT={txt:['الشرائح','تنسيق الخط','الحجم والخط واللون والتباعد','المحاذاة','القوائم'],ins:['إضافة محتوى','الأشكال والمخططات'],sld:['إدارة الشرائح','الترتيب','التصميم والبحث'],shw:['التشغيل والانتقالات'],nts:['الملاحظات']};
 const OT={fmt:'🎨 تنسيق',arr:'📐 ترتيب',anim:'✨ حركة',pic:'🖼 الصورة',tbl:'▦ جدول'};
 const segs=a=>{const o=[[]];for(const x of a){if(x==='|')o.push([]);else o[o.length-1].push(x)}return o.filter(g=>g.length)};
 const vmClose=()=>{if(!VM.open)return;VM.open=false;if(VM.el){VM.el.remove();VM.el=null}S.panel()};
 const vmRow=(it,key)=>{
  if(it==='|')return document.createDocumentFragment();
  if(it.el){const el=it.el,tg=el.tagName;
   if(tg==='SELECT'||tg==='INPUT'||el.classList.contains('ofcol')){
    const first=tg==='SELECT'&&el.options[0]&&el.options[0].value==='';
    const lb=first?'':(el.getAttribute('aria-label')||el.title||el.placeholder||'');
    const r=E('div','pv-row pv-w');if(lb){const sp=E('span','pv-l');sp.textContent=lb;r.append(sp)}
    if(!lb&&el.classList.contains('ofcol')&&el.title){const sp=E('span','pv-l');sp.textContent=el.title;r.append(sp)}
    r.append(el);return r}
   if(tg==='BUTTON'){el.classList.add('pv-row');if(!hasLetters(el.textContent)&&el.title)el.textContent=el.textContent+' '+el.title;return el}
   el.classList.add('pv-hint');return el}
  const b=E('button','pv-row');b.type='button';
  const label=it.lbl?it.t+' '+it.lbl:(hasLetters(it.t)?it.t:(it.title?it.t+' '+it.title:it.t));
  b.innerHTML=label;if(it.title&&hasLetters(it.t)&&!it.lbl&&!String(it.t).includes(it.title)){const sm=E('small','pv-s');sm.textContent=it.title;b.append(sm)}
  b.addEventListener('mousedown',e=>e.preventDefault());
  b.onclick=()=>{if(key!=='txt')vmClose();it.f()};return b};
 const vmGroups=key=>{
  if(key==='obj'){
   if(!S.sel)return[['العنصر المحدد',items.obj()]];
   const{tbl,pic,secs}=objSecs(),out=[['العنصر المحدد',[B('🗑',delShape,null,'حذف العنصر'),B('📋',copy,null,'نسخ'),B('📌',paste,null,'لصق')]]];
   for(const k of['fmt','arr','anim','pic','tbl']){if(k==='pic'&&!pic)continue;if(k==='tbl'&&!tbl)continue;if(secs[k]&&secs[k].length)segs(secs[k]).forEach((g,i)=>out.push([i?'':OT[k],g]))}
   return out}
  return segs(items[key]()).map((g,i)=>[(GT[key]||[])[i]||'',g])};
 const vmBuild=()=>{
  const m=VM.el;if(!m)return;const body=m.querySelector('.pv-b'),top=body.scrollTop;body.innerHTML='';
  for(const[k,v]of Object.entries(TB)){
   const d=E('details','pv-sec');d.dataset.k=k;d.open=VM.sec.has(k);
   const sm=E('summary','',`<i>${v[0]}</i><span>${v[1]}</span>`);d.append(sm);
   d.ontoggle=()=>{d.open?VM.sec.add(k):VM.sec.delete(k)};
   if(d.open){for(const[t,g]of vmGroups(k)){if(t){d.append(E('h6','pv-h',t))}for(const it of g)d.append(vmRow(it,k))}}
   else d.addEventListener('toggle',()=>{if(d.open)vmBuild()},{once:true});
   body.append(d)}
  body.scrollTop=top};
 const vmOpen=()=>{if(VM.open)return vmClose();VM.open=true;VM.sec=new Set([S.tab]);
  const m=E('div','pv-m','<div class="pv-t"><b>🧰 الأدوات</b><button type="button" class="pv-x" aria-label="إغلاق">✕</button></div><div class="pv-b"></div>');
  m.querySelector('.pv-x').onclick=vmClose;VM.el=m;($('#of')||document.body).append(m);vmBuild()};
 S.panel=()=>{pn.innerHTML='';for(const it of items[S.tab]()){if(it==='|'){pn.append(E('i','sep'));continue}if(it.el){pn.append(it.el);continue}const x=E('button','',it.t);if(it.title)x.setAttribute('aria-label',it.title);x.addEventListener('mousedown',e=>e.preventDefault());x.onclick=()=>it.f();pn.append(x)}[...tb.children].forEach(b=>b.classList.toggle('on',b.dataset.t==S.tab))};
 S.setTab=t=>{if(S.tab!=t)S.sub=null;S.tab=t;S.panel()};
 for(const[k,v]of Object.entries(TB)){const b=E('button','',`<i>${v[0]}</i><span>${v[1]}</span>`);b.dataset.t=k;b.onclick=()=>S.setTab(k);tb.append(b)}

 // ---------- dropdown ribbon menu (Home / Insert / Draw / Design / Transitions / Animations / Slide Show / Review) — PowerPoint editor only ----------
 {
 const selTools=()=>S.sel?[col('🎨','تعبئة',setFill),B('🌈 تدرج',gradDlg,null,'تدرج لوني للتعبئة أو الخلفية'),col('▭','لون الحد',c=>{S.lc=c;lineSet(c,S.lw)}),{el:lwSel},B('◪ ظل',shadow)]:[hint('اختار عنصر عشان تغيّر التعبئة والحد')];

 // ---------- القلم (رسم حر): طبقة شفافة فوق الشريحة، وكل خط بيتحفظ كشكل a:custGeom داخل ملف pptx ----------
 const PEN={m:null,c:'e11d48',w:3};
 const penW=E('select');penW.setAttribute('aria-label','سُمك القلم');penW.innerHTML=[1,2,3,5,8,12].map(n=>`<option value="${n}">سُمك ${n}</option>`).join('');penW.value=PEN.w;penW.onchange=()=>{PEN.w=+penW.value};
 let ov=null,raf=0;
 const lwPt=()=>PEN.m==='hl'?PEN.w*4:PEN.w;
 const penSync=()=>{if(!ov)return;const st=S&&S.stage;if(st&&st.isConnected){const r=st.getBoundingClientRect(),v=S.sc.getBoundingClientRect(),l=Math.max(r.left,v.left),t=Math.max(r.top,v.top),rr=Math.min(r.right,v.right),b=Math.min(r.bottom,v.bottom);Object.assign(ov.style,{left:l+'px',top:t+'px',width:Math.max(0,rr-l)+'px',height:Math.max(0,b-t)+'px'})}raf=requestAnimationFrame(penSync)};
 const penStop=()=>{PEN.m=null;if(ov){ov.remove();ov=null}cancelAnimationFrame(raf)};
 const penOff=()=>{penStop();S.panel()};
 const smooth=q=>{const o=[`<a:moveTo><a:pt x="${q[0].x}" y="${q[0].y}"/></a:moveTo>`];
  if(q.length==2)return o[0]+`<a:lnTo><a:pt x="${q[1].x}" y="${q[1].y}"/></a:lnTo>`;
  for(let i=0;i<q.length-1;i++){const p0=q[i-1]||q[i],p1=q[i],p2=q[i+1],p3=q[i+2]||p2,R=Math.round,
   c1=[R(p1.x+(p2.x-p0.x)/6),R(p1.y+(p2.y-p0.y)/6)],c2=[R(p2.x-(p3.x-p1.x)/6),R(p2.y-(p3.y-p1.y)/6)];
   o.push(`<a:cubicBezTo><a:pt x="${c1[0]}" y="${c1[1]}"/><a:pt x="${c2[0]}" y="${c2[1]}"/><a:pt x="${p2.x}" y="${p2.y}"/></a:cubicBezTo>`)}
  return o.join('')};
 const addInk=(a,mode,color,pt)=>{if(!a.length)return;if(a.length<2)a=[a[0],{x:a[0].x+1,y:a[0].y}];
  const lwE=Math.round(pt*12700),pad=lwE,xs=a.map(p=>p.x),ys=a.map(p=>p.y),mx=Math.floor(Math.min(...xs)),my=Math.floor(Math.min(...ys)),
   W=Math.max(1,Math.ceil(Math.max(...xs))-mx)+2*pad,H=Math.max(1,Math.ceil(Math.max(...ys))-my)+2*pad,
   q=a.map(p=>({x:Math.round(p.x-mx+pad),y:Math.round(p.y-my+pad)})),id=nid();
  const x=`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Ink ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${mx-pad}" y="${my-pad}"/><a:ext cx="${W}" cy="${H}"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="${W}" h="${H}" fill="none">${smooth(q)}</a:path></a:pathLst></a:custGeom><a:noFill/><a:ln w="${lwE}" cap="rnd"><a:solidFill><a:srgbClr val="${color}">${mode==='hl'?'<a:alpha val="40000"/>':''}</a:srgbClr></a:solidFill><a:round/></a:ln></p:spPr></p:sp>`;
  commit();snap();const n=frag(x,cur().doc);tree(cur()).append(n);touch();draw();thumb(S.cur)};
 const penEraseAt=e=>{const hit=document.elementsFromPoint(e.clientX,e.clientY).find(n=>n.classList&&n.classList.contains('pe-ink-hit')),d=hit&&hit.closest('.pe-s'),el=d&&d._e;
  if(el&&el.parentNode){commit();snap();el.remove();touch();draw();thumb(S.cur)}};
 const penOn=m=>{if(PEN.m===m)return penOff();commit();select(null);PEN.m=m;
  if(!ov){ov=E('div','pe-ink');ov.innerHTML='<svg style="position:absolute;left:0;top:0;width:100%;height:100%;overflow:visible;pointer-events:none"><path fill="none" stroke-linecap="round" stroke-linejoin="round"/></svg>';document.body.append(ov);
   const live=ov.querySelector('path');let pts=null,pid=null;
   const toS=e=>{const r=S.stage.getBoundingClientRect();return{x:(e.clientX-r.left)/r.width*S.W,y:(e.clientY-r.top)/r.height*S.H,cx:e.clientX,cy:e.clientY}};
   const show=()=>{const r=ov.getBoundingClientRect(),st=S.stage.getBoundingClientRect();live.setAttribute('stroke','#'+PEN.c);live.setAttribute('stroke-width',lwPt()*12700*st.width/S.W);live.setAttribute('stroke-opacity',PEN.m==='hl'?.4:1);
    live.setAttribute('d',pts.map((p,i)=>(i?'L':'M')+(p.cx-r.left)+' '+(p.cy-r.top)).join(''))};
   ov.onpointerdown=e=>{if(pid!=null||!PEN.m)return;e.preventDefault();pid=e.pointerId;try{ov.setPointerCapture(pid)}catch(_){}
    if(PEN.m==='er'){penEraseAt(e);return}pts=[toS(e)];pts.push({...pts[0],cx:pts[0].cx+.01});show()};
   ov.onpointermove=e=>{if(e.pointerId!==pid)return;e.preventDefault();if(PEN.m==='er'){penEraseAt(e);return}if(!pts)return;
    for(const ev of(e.getCoalescedEvents?e.getCoalescedEvents():[e])){const p=toS(ev),l=pts[pts.length-1];if(Math.hypot(p.cx-l.cx,p.cy-l.cy)>=1.5)pts.push(p)}show()};
   const up=e=>{if(e.pointerId!==pid)return;pid=null;const a=pts;pts=null;live.setAttribute('d','');if(a&&a.length&&PEN.m&&PEN.m!=='er'){const same=a.length==2&&a[1].cx-a[0].cx<.05;addInk(same?[a[0]]:a.map(p=>({x:p.x,y:p.y})),PEN.m,PEN.c,lwPt())}};
   ov.onpointerup=up;ov.onpointercancel=up;penSync()}
  X.toast(m==='er'?'اضغط على أي خط عشان تمسحه':'ارسم بصباعك على الشريحة');S.panel()};
 const gridSz=E('select');gridSz.setAttribute('aria-label','حجم مربعات الشبكة');gridSz.innerHTML=[.5,1,1.5,2,3].map(n=>`<option value="${n}">مربع ${n} سم</option>`).join('');gridSz.value=String(GD.cm);gridSz.onchange=()=>{GD.cm=+gridSz.value;GS('raPeGridCm',GD.cm);gridApply()};
 Object.assign(items,{
  drw:()=>[B((PEN.m=='pen'?'✅ ':'')+'🖊 قلم',()=>penOn('pen'),null,'رسم حر بالقلم'),B((PEN.m=='hl'?'✅ ':'')+'🖍 ماركر',()=>penOn('hl'),null,'قلم تحديد شفاف'),B((PEN.m=='er'?'✅ ':'')+'🧽 ممحاة',()=>penOn('er'),null,'امسح خط بالضغط عليه'),...(PEN.m?[B('✋ إيقاف الرسم',penOff)]:[]),col('🎨','لون القلم',c=>{PEN.c=c.replace('#','')},'#'+PEN.c),{el:penW},'|',B('▭',()=>addShape('rect'),null,'مستطيل'),B('◯',()=>addShape('ellipse'),null,'دائرة'),B('△',()=>addShape('triangle'),null,'مثلث'),B('☆',()=>addShape('star5'),null,'نجمة'),B('➡',()=>addShape('rightArrow'),null,'سهم'),B('🔷 كل الأشكال',shapeGallery,'مكتبة الأشكال'),'|',...selTools()],
  dsn:()=>[{el:tplSel},{el:themeSel},col('🖼','خلفية',bgSet),B('🌈 تدرج',gradDlg,null,'تدرج لوني للتعبئة أو الخلفية'),{el:szSel},'|',B((GD.show?'✅ ':'')+'▦ شبكة التصميم',()=>{GD.show=!GD.show;GS('raPeGrid',GD.show?1:0);gridApply();S.panel()},null,'إظهار شبكة التصميم'),{el:gridSz},B((GD.snap?'✅ ':'')+'🧲 التقاط للشبكة',()=>{GD.snap=!GD.snap;GS('raPeGridSnap',GD.snap?1:0);S.panel();X.toast(GD.snap?'العناصر هتلتقط على الشبكة وإنت بتسحب':'التقاط الشبكة اتقفل')},null,'التقاط للشبكة أثناء السحب'),B((GD.smart?'✅ ':'')+'⊹ خطوط ذكية',()=>{GD.smart=!GD.smart;GS('raPeSmart',GD.smart?1:0);S.panel();X.toast(GD.smart?'الخطوط الذكية شغالة':'الخطوط الذكية اتقفلت')},null,'خطوط المحاذاة الذكية')],
  trn:()=>items.shw(),
  ani:()=>{const a=animCtl();return a.length?a:[hint('اختار عنصر في الشريحة عشان تضيف له حركة')]},
  sho:()=>[B('▶ من الأول',()=>play(0)),B('▶ من الشريحة الحالية',()=>play(S.cur))],
  rev:()=>[B('📝 ملاحظات',()=>S.setTab('nts')),B('🔎 بحث في الشرائح',findInSlides),B('🔍 استبدال',repl)]});
 const MN=[['Home','txt'],['Insert','ins'],['Draw','drw'],['Design','dsn'],['Transitions','trn'],['Animations','ani'],['Slide Show','sho'],['Review','rev']];
 const mnOf=t=>{const m=MN.find(x=>x[1]===t);return m?m[0]:(t==='nts'?'Review':'Object')};
 if(!document.getElementById('pe-mn-css')){const c=E('style');c.id='pe-mn-css';c.textContent=`.pe-tabs .pe-mn{flex:0 0 auto;min-width:96px;flex-direction:row;justify-content:center;gap:6px;align-self:center;margin:4px 6px;min-height:40px;padding:6px 12px;border:0;border-radius:10px;background:#262626;color:#fff;font-size:15px;font-weight:600;direction:ltr}.pe-mn i{font-size:13px;color:#e8552b}.pe-mm{position:fixed;z-index:400;direction:ltr;width:176px;background:#262626;border-radius:10px;box-shadow:0 8px 28px #0007;overflow:hidden;padding:2px 0}.pe-mm button{display:block;width:100%;text-align:left;border:0;border-bottom:1px solid #3a3a3a;background:none;color:#fff;font-size:18px;padding:0 18px;height:52px;font-family:inherit}.pe-mm button:last-child{border-bottom:0}.pe-mm button.on{color:#e8552b;background:#111}.pe-mm button:active{background:#3a3a3a}.pe-tabs{min-height:48px!important;justify-content:flex-start}.pe-bb{flex:none;display:flex;align-items:center;background:#f1f4fa;border-top:1px solid #d5dbe8;padding-bottom:env(safe-area-inset-bottom)}.pe-bb .pe-tabs{flex:none;min-height:0!important;border:0;background:none;overflow:visible}.pe-bb .pe-tabs .pe-mn{margin:3px 4px 3px 6px;min-height:34px!important;padding:4px 10px;font-size:13px;min-width:84px}.pe-bb .pe-pn{flex:1;min-width:0;min-height:0!important;max-height:none!important;padding:3px 6px!important;gap:4px!important;border:0;background:none}.pe-bb .pe-pn button,.pe-bb .pe-pn select,.pe-bb .pe-pn label,.pe-bb .pe-pn .ofcol{min-height:34px!important;height:34px!important;min-width:36px!important;padding:0 8px!important;font-size:13px!important;border-radius:9px!important}.pe-bb .pe-pn input{height:34px!important;font-size:13px!important}.pe-bb .pe-pn i.sep{height:20px}.pe-th{min-height:0!important;max-height:none!important;padding:3px 6px!important;border-top:0}.pe-nt{height:80px!important;min-height:80px!important}.pe-mm button{height:44px;font-size:16px}.pgp{position:relative;width:60px;height:60px}.pgi{position:absolute;left:0;top:0;width:100%;height:100%}.pgp svg{position:absolute;left:0;top:0;width:100%!important;height:100%!important;overflow:visible;pointer-events:none}.pgn{margin:4px 0 0;font-size:12px;color:var(--fg2)}.pe-ds{position:fixed;left:8px;right:8px;bottom:calc(8px + env(safe-area-inset-bottom));z-index:90;max-width:520px;margin:0 auto;background:#fff;color:#111;border:1px solid #cfd6e6;border-radius:16px;box-shadow:0 10px 34px #0007;direction:rtl;padding:10px 12px;max-height:64vh;overflow:auto}.pds-h{display:flex;justify-content:space-between;align-items:center;font-size:16px;margin-bottom:6px}.pds-h button{border:0;background:none;font-size:20px;color:#111}.pds-l{max-height:22vh;overflow:auto;border:1px solid #e2e6f0;border-radius:10px;padding:2px 8px;margin-bottom:4px}.pds-l label{display:flex;gap:8px;align-items:center;padding:7px 2px;font-size:14px;border-bottom:1px solid #eef1f7}.pds-l input{width:20px;height:20px;flex:none}.pds-all{border:0;background:none;color:#2563eb;font-size:13px;padding:6px 0}.pds-n{font-size:12px;color:#667;margin:2px 0}.pds-g{display:flex;flex-wrap:wrap;gap:6px;margin:6px 0}.pds-g b{flex:0 0 100%;font-size:12px;color:#667;font-weight:600}.pds-g button,.pds-g select{height:38px;padding:0 12px;font-size:14px;border-radius:10px;border:1px solid #cfd6e6;background:#f4f6fb;color:#111;font-family:inherit}.pds-g select{flex:1 1 100%}.pe-ms{outline:2px dashed #f97316!important;outline-offset:1px}.pe-ink{position:fixed;z-index:250;touch-action:none;cursor:crosshair;background:transparent}.pe-inkd{pointer-events:none}.pe-inkd>.pe-h{pointer-events:auto}`;document.head.append(c)}
 const mb=E('button','pe-mn');mb.type='button';mb.setAttribute('aria-label','قائمة الأقسام');
 let mpop=null;
 const mOut=e=>{if(mpop&&!mpop.contains(e.target)&&!mb.contains(e.target))mClose()};
 const mClose=()=>{if(mpop){mpop.remove();mpop=null;document.removeEventListener('pointerdown',mOut,true)}};
 mb.onclick=()=>{if(mpop)return mClose();const r=mb.getBoundingClientRect();mpop=E('div','pe-mm');mpop.style.bottom=(innerHeight-r.top+4)+'px';mpop.style.left=Math.max(6,Math.min(r.left,innerWidth-182))+'px';
  for(const[l,t]of MN){const b=E('button','',l);b.type='button';if(mnOf(S.tab)===l)b.classList.add('on');b.onclick=()=>{mClose();S.setTab(t)};mpop.append(b)}
  document.body.append(mpop);document.addEventListener('pointerdown',mOut,true)};
 const mbSet=()=>{mb.innerHTML='<span>'+mnOf(S.tab)+'</span><i>▾</i>'};
 const pnM=S.panel;S.panel=()=>{if(S.tab!=='drw'&&PEN.m)penStop();pnM();mbSet()};
 [...tb.children].forEach(b=>b.remove());
 {const txt0=items.txt;items.txt=()=>[B('＋ جديدة',()=>addSlide(0)),B('⧉ تكرار',()=>addSlide(1)),B('🗑 حذف',delSlide),B('⬆',()=>mv(-1),null,'نقل الشريحة لأعلى'),B('⬇',()=>mv(1),null,'نقل الشريحة لأسفل'),B('📐 توزيع',distDlg,null,'محاذاة وتوزيع العناصر بمسافات متساوية'),'|',...txt0()]}
 tb.prepend(mb);mbSet();S.mClose=mClose;S.penStop=penStop;
 }
 const ntb=E('textarea','pe-nt');ntb.hidden=true;ntb.dir='auto';ntb.placeholder='اكتب ملاحظات المتحدث للشريحة دي هنا…';ntb.oninput=()=>{const s=S.ntFor;if(!s||!s.nt)return;s.nt.text=ntb.value;s.nt.dirty=1;X.O().dirty=1};S.nb=ntb;
 const pn0=S.panel;S.panel=()=>{pn0();ntb.hidden=S.tab!='nts';if(!ntb.hidden)ntBind(ntb);if(VM.open&&VM.el)vmBuild()};
 {const bb=E('div','pe-bb');bb.append(tb,pn);h.append(ntb);h.append(bb)}S.panel();
 X.setBar([B('☰ الأدوات',vmOpen,'قائمة كل الأدوات'),B('💾 حفظ',save),B('⬇ PDF',exportPdf,'تصدير كل الشرائح إلى PDF'),B('▶',()=>play(S.cur)),B('↶',undo,'تراجع'),B('↷',redo,'إعادة'),...(S.media.length?[B('🎬 '+S.media.length,()=>X.media.list(),'الوسائط')]:[]),B('－',()=>zoom(1/1.25),'تصغير'),B('＋',()=>zoom(1.25),'تكبير')]);
 const sc=()=>{const g=getSelection();if(S&&S.ed&&g.rangeCount&&S.ed.el.contains(g.anchorNode))S.rng=g.getRangeAt(0).cloneRange();if(S.szIn&&document.activeElement!==S.szIn){const n=g.anchorNode,el=n&&(n.nodeType==1?n:n.parentElement),px=el&&parseFloat(getComputedStyle(el).fontSize);if(px&&S.k)S.szIn.value=Math.round(px/(S.k*PT)*10)/10}};document.addEventListener('selectionchange',sc);
const ro=new ResizeObserver(()=>{S&&S.stage&&S.stage.isConnected&&!S.ed&&draw()});ro.observe(S.sc);
 {const o=X.O();let dd=o.dirty;Object.defineProperty(o,'dirty',{get:()=>dd,set:v=>{dd=v;if(v&&S)S.dv=(S.dv||0)+1},configurable:true});o.autosave=autosave}
 const iv=setInterval(()=>autosave(),4000),vis=()=>{if(document.hidden)autosave(1)},ph=()=>autosave(1);document.addEventListener('visibilitychange',vis);addEventListener('pagehide',ph);
 X.O().cleanup=()=>{S&&S.mClose&&S.mClose();S&&S.penStop&&S.penStop();document.querySelectorAll('.pe-ds').forEach(n=>n.remove());VM.open=false;VM.el&&VM.el.remove();clearInterval(iv);document.removeEventListener('visibilitychange',vis);removeEventListener('pagehide',ph);S&&X.saveMeta&&X.saveMeta({ls:S.cur});X.media&&X.media.close();ro.disconnect();document.removeEventListener('selectionchange',sc);if(S)Object.values(S.blobs).forEach(u=>URL.revokeObjectURL(u));S=null};
 {const o=X.O();if(o.opt&&o.opt.deck){const dk=o.opt.deck;delete o.opt.deck;try{applyDeck(dk)}catch(e){console.warn('deck',e)}}}
 fit();thumbs();draw();trSync()};
})();
