// Reader All — Office module (Word / Excel / PowerPoint). Loaded after plus.js.
(()=>{
'use strict';
const $=s=>document.querySelector(s),E=(t,c,h)=>{const e=document.createElement(t);if(c)e.className=c;if(h!=null)e.innerHTML=h;return e};
const KIND={doc:'Word',xls:'Excel',ppt:'PowerPoint'};
window.raKind=d=>KIND[d.k]||'ملف';

// ---------- lazy library loader (local lib/ first, CDN fallback) ----------
const LIB={
 mammoth:{f:'mammoth.browser.min.js',c:'https://cdn.jsdelivr.net/npm/mammoth@1.8.0/mammoth.browser.min.js',g:'mammoth'},
 dp:{f:'docx-preview.min.js',c:'https://cdn.jsdelivr.net/npm/docx-preview@0.3.2/dist/docx-preview.min.js',g:'docx'},
 dx:{f:'docx.min.js',c:'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.js',g:'docx'},
 xl:{f:'exceljs.min.js',c:'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js',g:'ExcelJS'},
 sj:{f:'xlsx.full.min.js',c:'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',g:'XLSX'},
 fp:{f:'formula-parser.min.js',c:'https://cdn.jsdelivr.net/npm/hot-formula-parser@4.0.0/dist/formula-parser.min.js',g:'formulaParser'},
 pp:{f:'pptx-preview.umd.js',c:'https://cdn.jsdelivr.net/npm/pptx-preview@1.0.5/dist/pptx-preview.umd.js',g:'pptxPreview'},
 ff:{f:'ffmpeg/ffmpeg.js',c:'https://cdn.jsdelivr.net/npm/@ffmpeg/ffmpeg@0.12.10/dist/umd/ffmpeg.js',g:'FFmpegWASM'},
 h2c:{f:'html2canvas.min.js',c:'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js',g:'html2canvas'}
};
const G={};let LQ=Promise.resolve();
const scr=src=>new Promise((ok,no)=>{const s=document.createElement('script');s.src=src;s.onload=ok;s.onerror=()=>{s.remove();no(new Error('load '+src))};document.head.append(s)});
function need(k){const p=LQ.then(async()=>{if(G[k])return G[k];const L=LIB[k];try{window[L.g]=undefined}catch(e){}
 try{await scr('lib/'+L.f)}catch(e){await scr(L.c)}
 if(!window[L.g]){try{await scr(L.c)}catch(e){}}
 G[k]=window[L.g];if(!G[k])throw new Error('المكتبة '+k+' ما اتحمّلتش');return G[k]});
 LQ=p.catch(()=>{});return p}

// ---------- helpers ----------
const busy=(on,t)=>window.raBusy&&window.raBusy(on,t);
const getBuf=id=>tx('f','readonly',s=>s.get(id));
const putBuf=(id,b)=>tx('f','readwrite',s=>s.put(b,id));
const baseName=n=>n.replace(/\.(docx|xlsx|xls|csv|pptx)$/i,'');
const MIME={doc:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',xls:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',ppt:'application/vnd.openxmlformats-officedocument.presentationml.presentation'};
const EXT={doc:'.docx',xls:'.xlsx',ppt:'.pptx'};
function dl(blob,name){const a=E('a'),u=URL.createObjectURL(blob);a.href=u;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(u),60000)}
const printOffice=()=>{document.body.classList.add('ra-print');window.print();setTimeout(()=>document.body.classList.remove('ra-print'),1200)};
async function encryptedExport(){
 const pass=await raAsk('اختار كلمة مرور لتشفير نسخة الملف (8 أحرف على الأقل):');if(pass===null)return;if(pass.length<8)return alert('كلمة المرور لازم تكون 8 أحرف على الأقل');
 try{busy(true,'جاري التشفير…');const salt=crypto.getRandomValues(new Uint8Array(16)),iv=crypto.getRandomValues(new Uint8Array(12)),mat=await crypto.subtle.importKey('raw',new TextEncoder().encode(pass),'PBKDF2',false,['deriveKey']);
 const key=await crypto.subtle.deriveKey({name:'PBKDF2',salt,iterations:250000,hash:'SHA-256'},mat,{name:'AES-GCM',length:256},false,['encrypt']);
 const meta=new TextEncoder().encode(JSON.stringify({name:O.d.n+EXT[O.d.k],kind:O.d.k})),raw=new Uint8Array(O.buf),plain=new Uint8Array(8+meta.length+raw.length);plain.set(new TextEncoder().encode('RAE1'));new DataView(plain.buffer).setUint32(4,meta.length);plain.set(meta,8);plain.set(raw,8+meta.length);
 const cipher=new Uint8Array(await crypto.subtle.encrypt({name:'AES-GCM',iv},key,plain)),out=new Uint8Array(32+cipher.length);out.set(new TextEncoder().encode('RAC1'));out.set(salt,4);out.set(iv,20);out.set(cipher,32);dl(new Blob([out]),O.d.n+'.raenc');busy(false);toast('تم إنشاء نسخة مشفرة AES-GCM');
 }catch(e){fail('تعذر التشفير',e)}
}
async function decryptOfficeFile(file){
 try{busy(true,'جاري فك التشفير…');const a=new Uint8Array(await file.arrayBuffer());if(a.length<48||new TextDecoder().decode(a.slice(0,4))!=='RAC1')throw 0;const pass=await raAsk('اكتب كلمة مرور الملف المشفر:');if(pass===null){busy(false);return}
 const salt=a.slice(4,20),iv=a.slice(20,32),mat=await crypto.subtle.importKey('raw',new TextEncoder().encode(pass),'PBKDF2',false,['deriveKey']),key=await crypto.subtle.deriveKey({name:'PBKDF2',salt,iterations:250000,hash:'SHA-256'},mat,{name:'AES-GCM',length:256},false,['decrypt']);
 const plain=new Uint8Array(await crypto.subtle.decrypt({name:'AES-GCM',iv},key,a.slice(32)));if(new TextDecoder().decode(plain.slice(0,4))!=='RAE1')throw 0;const ml=new DataView(plain.buffer,plain.byteOffset).getUint32(4),meta=JSON.parse(new TextDecoder().decode(plain.slice(8,8+ml))),buf=plain.slice(8+ml);if(!['doc','xls','ppt'].includes(meta.kind))throw 0;
 const d=await createRec(meta.kind,(meta.name||'ملف').replace(/\.(docx|xlsx|pptx)$/i,'')+' (مفكوك',buf.buffer.slice(buf.byteOffset,buf.byteOffset+buf.byteLength),1);busy(false);openOffice(d,{edit:1,newDoc:meta.kind==='doc'});
 }catch(e){busy(false);alert('فشل فك التشفير. تأكد من كلمة المرور وأن الملف سليم.')}
}
document.getElementById('encin')?.addEventListener('change',e=>{const f=e.target.files[0];e.target.value='';if(f)decryptOfficeFile(f)});

const toast=t=>{const d=E('div','oftoast');d.textContent=t;document.body.append(d);setTimeout(()=>d.remove(),2600)};
const fail=(m,e)=>{busy(false);console.warn(e);alert(m+(e&&e.message?': '+e.message:''))};

// ---------- shell ----------
const sec=E('section');sec.id='of';sec.hidden=true;
sec.innerHTML='<div id="oft"><button id="ofb" aria-label="رجوع">←</button><b id="oftt"></b></div><div id="ofbar"></div><div id="ofm"></div>';
document.body.append(sec);
let O=null;
function setBar(items){const b=$('#ofbar');b.innerHTML='';b.hidden=!items.length;for(const it of items){if(it==='|'){b.append(E('i','sep'));continue}
 if(it.el){b.append(it.el);continue}
 const x=E('button',it.on?'on':'',it.t);if(it.title)x.setAttribute('aria-label',it.title);x.addEventListener('mousedown',e=>e.preventDefault());x.onclick=()=>it.f(x);b.append(x)}}
async function back(){if(!O)return;if(O.dirty&&O.autosave){try{await O.autosave(1)}catch(e){}}if(O&&O.dirty&&!await raConfirm('في تعديلات مش محفوظة. تخرج من غير حفظ؟'))return;close()}
function close(){if(O&&O.cleanup)O.cleanup();sec.hidden=true;O=null;$('#ofm').innerHTML='';$('#ofm').className='';$('#home').hidden=false;home()}
$('#ofb').onclick=back;
addEventListener('load',()=>{const o=window.raBack;window.raBack=()=>{if(O&&!sec.hidden){back();return true}return o?o():false}});

// ---------- open / import ----------
const OFK={doc:1,xls:1,ppt:1};
const _open=window.openDoc;
const UNL={};
window.openDoc=async function(id){let d;try{d=await tx('d','readonly',s=>s.get(id))}catch(e){}
 if(d&&OFK[d.k]){if(d.lk&&window.raVerify&&!(UNL[id]&&Date.now()-UNL[id]<180000)){if(!await window.raVerify({title:'الملف مقفول 🔒',sub:d.n}))return;UNL[id]=Date.now()}return openOffice(d)}
 return _open.apply(this,arguments)};
async function openOffice(d,opt={}){
 if(O&&O.cleanup)O.cleanup();O={d,opt,dirty:0,z:1};$('#home').hidden=true;sec.hidden=false;$('#oftt').textContent=d.n;$('#ofm').innerHTML='';$('#ofm').className='';setBar([]);
 try{busy(true,'جاري فتح الملف…');
  const buf=await getBuf(d.id);if(!buf){busy(false);alert('الملف مش موجود (ممكن يكون اتمسح).');return close()}
  O.buf=buf;
  if(d.k==='doc')opt.edit?await wordEdit():await wordView();else if(d.k==='xls')await xlsOpen();else if((opt.edit||d.ed)&&window.pptEdit)await window.pptEdit();else await pptView();
  busy(false)}catch(e){fail('تعذر فتح الملف',e)}}
async function createRec(k,name,buf,own){const id='o'+Date.now()+Math.floor(Math.random()*99);await putBuf(id,buf);const d={id,n:name,k,s:{},u:Date.now(),own:own?1:0};await dbPut(d);return d}
async function saveBuf(buf,suffix,extra){ // never overwrite an imported original: first save makes a copy
 if(O.d.own){if(extra)Object.assign(O.d,extra);O.d.u=Date.now();await putBuf(O.d.id,buf);await dbPut(O.d)}
 else{const d=await createRec(O.d.k,O.d.n+(suffix||' (معدّل)'),buf,1);if(extra){Object.assign(d,extra);await dbPut(d)}O.d=d;$('#oftt').textContent=d.n;toast('اتحفظ كنسخة جديدة — الأصل زي ما هو')}
 O.buf=buf;O.dirty=0}
['fo','fod','fox','fop'].forEach(id=>{const el=document.getElementById(id);el&&(el.onchange=async e=>{const f=e.target.files[0];e.target.value='';if(f)importFile(f)})});
async function importFile(f){
 const ext=(/\.([a-z0-9]+)$/i.exec(f.name)||[])[1]||'';let k=null;const x=ext.toLowerCase();
 if(x==='docx')k='doc';else if(x==='xlsx'||x==='xls'||x==='csv')k='xls';else if(x==='pptx')k='ppt';
 if(!k)return alert(/^(doc|ppt)$/i.test(x)?'صيغة .'+x+' القديمة مش مدعومة. حوّل الملف لـ docx / pptx وجرّب تاني.':'نوع الملف مش مدعوم.');
 try{busy(true,'جاري الاستيراد…');let buf=await f.arrayBuffer();if(['docx','xlsx','pptx'].includes(x)&&(new TextDecoder().decode(new Uint8Array(buf).slice(0,2))!=='PK'))throw new Error('الملف ليس حزمة Office Open XML سليمة أو قد يكون محميًا بكلمة مرور.');
  if(x==='xls'){const S=await need('sj');const wb=S.read(buf);buf=S.write(wb,{bookType:'xlsx',type:'array'})}
  if(x==='csv')buf=await csvToXlsx(new TextDecoder('utf-8').decode(buf));
  const d=await createRec(k,baseName(f.name),buf);busy(false);openOffice(d)}catch(e){fail('تعذر استيراد الملف',e)}}
window.raImportOffice=importFile;
function wire(id,fn){const b=document.getElementById(id);b&&(b.onclick=fn)}
// blank .docx / .xlsx written by hand (no library download needed, instant)
function zipStore(files){const te=new TextEncoder(),parts=[],cd=[];let off=0;
 for(const [nm,txt] of files){const nb=te.encode(nm),d=te.encode(txt),crc=crc32(d),h=new DataView(new ArrayBuffer(30)),c=new DataView(new ArrayBuffer(46));
  h.setUint32(0,0x04034b50,true);h.setUint16(4,20,true);h.setUint16(6,0x0800,true);h.setUint16(10,0,true);h.setUint16(12,0x21,true);h.setUint32(14,crc,true);h.setUint32(18,d.length,true);h.setUint32(22,d.length,true);h.setUint16(26,nb.length,true);
  c.setUint32(0,0x02014b50,true);c.setUint16(4,20,true);c.setUint16(6,20,true);c.setUint16(8,0x0800,true);c.setUint16(12,0,true);c.setUint16(14,0x21,true);c.setUint32(16,crc,true);c.setUint32(20,d.length,true);c.setUint32(24,d.length,true);c.setUint16(28,nb.length,true);c.setUint32(42,off,true);
  parts.push(new Uint8Array(h.buffer),nb,d);cd.push(new Uint8Array(c.buffer),nb);off+=30+nb.length+d.length}
 let n=0;cd.forEach(x=>n+=x.length);const e=new DataView(new ArrayBuffer(22)),k=cd.length/2;
 e.setUint32(0,0x06054b50,true);e.setUint16(8,k,true);e.setUint16(10,k,true);e.setUint32(12,n,true);e.setUint32(16,off,true);
 return new Blob([...parts,...cd,new Uint8Array(e.buffer)]).arrayBuffer()}
const XH='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',RELNS='http://schemas.openxmlformats.org/package/2006/relationships',ODR='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const blankDocx=()=>zipStore([
 ['[Content_Types].xml',XH+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'],
 ['_rels/.rels',XH+'<Relationships xmlns="'+RELNS+'"><Relationship Id="rId1" Type="'+ODR+'/officeDocument" Target="word/document.xml"/></Relationships>'],
 ['word/document.xml',XH+'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>']]);
const blankXlsx=()=>zipStore([
 ['[Content_Types].xml',XH+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>'],
 ['_rels/.rels',XH+'<Relationships xmlns="'+RELNS+'"><Relationship Id="rId1" Type="'+ODR+'/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
 ['xl/workbook.xml',XH+'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="'+ODR+'"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'],
 ['xl/_rels/workbook.xml.rels',XH+'<Relationships xmlns="'+RELNS+'"><Relationship Id="rId1" Type="'+ODR+'/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'],
 ['xl/worksheets/sheet1.xml',XH+'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>']]);
const WORD_TPL={
 cv:'<h1>السيرة الذاتية</h1><h2>البيانات الشخصية</h2><p>الاسم: ....................</p><p>الهاتف: ....................</p><p>البريد الإلكتروني: ....................</p><h2>الملخص المهني</h2><p>اكتب نبذة مختصرة عن خبراتك ومهاراتك.</p><h2>الخبرات العملية</h2><p><b>المسمى الوظيفي — جهة العمل</b></p><ul><li>المسؤولية أو الإنجاز الأول</li><li>المسؤولية أو الإنجاز الثاني</li></ul><h2>التعليم</h2><p>المؤهل — المؤسسة التعليمية — السنة</p><h2>المهارات</h2><ul><li>مهارة أولى</li><li>مهارة ثانية</li></ul>',
 report:'<h1>عنوان التقرير</h1><p>إعداد: ....................　 التاريخ: ....................</p><h2>الملخص التنفيذي</h2><p>اكتب ملخصًا موجزًا لأهم النتائج.</p><h2>المقدمة</h2><p>خلفية الموضوع وهدف التقرير.</p><h2>المنهجية</h2><p>طريقة جمع المعلومات وتحليلها.</p><h2>النتائج والمناقشة</h2><p>اعرض النتائج الرئيسية هنا.</p><h2>التوصيات</h2><ol><li>التوصية الأولى</li><li>التوصية الثانية</li></ol><h2>المراجع</h2><p>أضف المصادر المستخدمة.</p>',
 letter:'<h1>عنوان الخطاب</h1><p>التاريخ: ....................</p><p>السيد/السيدة: ....................</p><p>تحية طيبة وبعد،</p><p>اكتب نص الخطاب هنا...</p><p>وتفضلوا بقبول فائق الاحترام،</p><p>الاسم: ....................</p><p>التوقيع: ....................</p>'
};
const WORD_BASE=[{id:'cv',ic:'👤',nm:'سيرة ذاتية'},{id:'report',ic:'📄',nm:'تقرير'},{id:'letter',ic:'✉️',nm:'خطاب'}];
async function makeWordTpl(t){try{busy(true,'جاري تجهيز القالب…');const dx=await need('dx'),tmp=E('div');tmp.innerHTML=t.html;const blob=await htmlToDocx(dx,tmp),buf=await blob.arrayBuffer(),d=await createRec('doc',t.nm,buf,1);busy(false);openOffice(d,{edit:1,newDoc:1})}catch(e){fail('تعذر إنشاء القالب',e)}}
wire('tpl-doc',()=>{const o=document.querySelector('.twg');if(o){o.remove();return}
 const all=[...WORD_BASE.map(b=>({...b,html:WORD_TPL[b.id]})),...(window.WORD_TPLS||[])];
 const w=E('div','shg twg'),box=E('div','shb','<div class="shh"><b>📚 قوالب Word ('+all.length+')</b><button type="button" aria-label="إغلاق">✕</button></div>'),close=()=>w.remove();
 box.querySelector('button').onclick=close;w.onpointerdown=e=>{if(e.target===w)close()};
 const h=E('h5');h.textContent='اختار قالب وهيتفتح جاهز للتعديل';const g=E('div','tpg');
 for(const t of all){const b=E('button','tpt','<i>'+t.ic+'</i><span></span>');b.querySelector('span').textContent=t.nm;b.onclick=()=>{close();makeWordTpl(t)};g.append(b)}
 box.append(h,g);w.append(box);document.body.append(w)});
wire('nw-doc',async()=>{try{busy(true,'جاري إنشاء المستند…');const buf=await blankDocx();
 const d=await createRec('doc','مستند '+new Date().toLocaleDateString('ar-EG'),buf,1);busy(false);openOffice(d,{edit:1,newDoc:1})}catch(e){fail('تعذر إنشاء المستند',e)}});
// ---------- new blank PowerPoint (built from scratch: 16:9, one title slide) ----------
function pptNew(){
 const H='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',te=new TextEncoder(),
  PN='xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"',
  RT='http://schemas.openxmlformats.org/officeDocument/2006/relationships/',PK='http://schemas.openxmlformats.org/package/2006/relationships',
  CTP='application/vnd.openxmlformats-officedocument.presentationml.',
  rels=l=>H+'<Relationships xmlns="'+PK+'">'+l.map((x,i)=>'<Relationship Id="rId'+(i+1)+'" Type="'+RT+x[0]+'" Target="'+x[1]+'"/>').join('')+'</Relationships>',
  grp='<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>',
  W=12192000,Hh=6858000,
  tx=(id,nm,x,y,w,h,t,sz,b,c,an,fill)=>'<p:sp><p:nvSpPr><p:cNvPr id="'+id+'" name="'+nm+'"/><p:cNvSpPr'+(t?' txBox="1"':'')+'/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="'+x+'" y="'+y+'"/><a:ext cx="'+w+'" cy="'+h+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'+(fill?'<a:solidFill><a:srgbClr val="'+fill+'"/></a:solidFill>':'<a:noFill/>')+'</p:spPr><p:txBody><a:bodyPr wrap="square" anchor="'+an+'"><a:normAutofit/></a:bodyPr><a:lstStyle/>'+(t?'<a:p><a:pPr algn="ctr" rtl="1"/><a:r><a:rPr lang="ar-EG" sz="'+sz+'"'+(b?' b="1"':'')+'><a:solidFill><a:srgbClr val="'+c+'"/></a:solidFill></a:rPr><a:t>'+t+'</a:t></a:r></a:p>':'<a:p><a:endParaRPr lang="en-US"/></a:p>')+'</p:txBody></p:sp>',
  fill3=t=>t.repeat(3),sf='<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>',
  clrs=[['dk2','44546A'],['lt2','E7E6E6'],['accent1','4472C4'],['accent2','ED7D31'],['accent3','A5A5A5'],['accent4','FFC000'],['accent5','5B9BD5'],['accent6','70AD47'],['hlink','0563C1'],['folHlink','954F72']],
  font='<a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface="Arial"/>',
  theme=H+'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Reader All"><a:themeElements><a:clrScheme name="Reader All"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'+clrs.map(c=>'<a:'+c[0]+'><a:srgbClr val="'+c[1]+'"/></a:'+c[0]+'>').join('')+'</a:clrScheme><a:fontScheme name="Reader"><a:majorFont>'+font+'</a:majorFont><a:minorFont>'+font+'</a:minorFont></a:fontScheme><a:fmtScheme name="Reader"><a:fillStyleLst>'+fill3(sf)+'</a:fillStyleLst><a:lnStyleLst>'+fill3('<a:ln w="6350">'+sf+'</a:ln>')+'</a:lnStyleLst><a:effectStyleLst>'+fill3('<a:effectStyle><a:effectLst/></a:effectStyle>')+'</a:effectStyleLst><a:bgFillStyleLst>'+fill3(sf)+'</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>',
  master=H+'<p:sldMaster '+PN+'><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>'+grp+'</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>',
  layout=H+'<p:sldLayout '+PN+' type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>'+grp+'</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>',
  slide=H+'<p:sld '+PN+'><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="1E3A8A"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>'+grp
   +tx(2,'Title',Math.round(W*.1),Math.round(Hh*.28),Math.round(W*.8),Math.round(Hh*.2),'عنوان العرض',5000,1,'FFFFFF','b')
   +tx(3,'Line',Math.round(W*.42),Math.round(Hh*.5),Math.round(W*.16),Math.round(Hh*.012),'',0,0,'',  'ctr','FBBF24')
   +tx(4,'Subtitle',Math.round(W*.1),Math.round(Hh*.54),Math.round(W*.8),Math.round(Hh*.14),'العنوان الفرعي — الاسم / التاريخ',2400,0,'C7D2FE','t')
   +'</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>',
  ct=H+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
   +'<Override PartName="/ppt/presentation.xml" ContentType="'+CTP+'presentation.main+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="'+CTP+'slideMaster+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="'+CTP+'slideLayout+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="'+CTP+'slide+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
  pres=H+'<p:presentation '+PN+'><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="'+W+'" cy="'+Hh+'"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>';
 const put=new Map([['[Content_Types].xml',ct],['_rels/.rels',rels([['officeDocument','ppt/presentation.xml']])],['ppt/presentation.xml',pres],
  ['ppt/_rels/presentation.xml.rels',rels([['slideMaster','slideMasters/slideMaster1.xml'],['slide','slides/slide1.xml'],['theme','theme/theme1.xml']])],
  ['ppt/slideMasters/slideMaster1.xml',master],['ppt/slideMasters/_rels/slideMaster1.xml.rels',rels([['slideLayout','../slideLayouts/slideLayout1.xml'],['theme','../theme/theme1.xml']])],
  ['ppt/slideLayouts/slideLayout1.xml',layout],['ppt/slideLayouts/_rels/slideLayout1.xml.rels',rels([['slideMaster','../slideMasters/slideMaster1.xml']])],
  ['ppt/slides/slide1.xml',slide],['ppt/slides/_rels/slide1.xml.rels',rels([['slideLayout','../slideLayouts/slideLayout1.xml']])],['ppt/theme/theme1.xml',theme]].map(([k,v])=>[k,te.encode(v)]));
 return zipBuild({m:{}},new Set(),put).arrayBuffer()}
wire('nw-ppt',async()=>{try{const deck=window.pptDeckPicker?await window.pptDeckPicker():'biz';if(!deck)return;const buf=await pptNew(),d=await createRec('ppt','عرض '+new Date().toLocaleDateString('ar-EG'),buf,1);openOffice(d,{edit:1,deck})}catch(e){fail('تعذر إنشاء العرض',e)}});
wire('tpl-xls',async()=>{const c=await raAsk('اختر قالب Excel: budget للميزانية أو attendance للحضور أو inventory للمخزون','budget');if(!c)return;const k=c.trim().toLowerCase();if(!['budget','attendance','inventory'].includes(k))return toast('استخدم budget أو attendance أو inventory');try{busy(true,'جاري إنشاء القالب…');const ex=await need('xl'),wb=new ex.Workbook(),ws=wb.addWorksheet(({budget:'الميزانية',attendance:'الحضور',inventory:'المخزون'}[k]));const rows={budget:[['البند','المبلغ المخطط','المبلغ الفعلي','الفرق'],['إيجار',0,0,{formula:'B2-C2'}],['طعام',0,0,{formula:'B3-C3'}],['مواصلات',0,0,{formula:'B4-C4'}],['الإجمالي',{formula:'SUM(B2:B4)'},{formula:'SUM(C2:C4)'},{formula:'B5-C5'}]],attendance:[['الاسم','التاريخ','الحالة','ملاحظات'],['','','حاضر',''],['','','غائب','']],inventory:[['الصنف','الكمية','سعر الوحدة','الإجمالي'],['',0,0,{formula:'B2*C2'}],['',0,0,{formula:'B3*C3'}],['الإجمالي','','',{formula:'SUM(D2:D3)'}]]};ws.addRows(rows[k]);ws.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};ws.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1E40AF'}};ws.columns.forEach(col=>col.width=20);ws.views=[{state:'frozen',ySplit:1}];const b=await wb.xlsx.writeBuffer(),d=await createRec('xls',({budget:'ميزانية',attendance:'سجل حضور',inventory:'مخزون'}[k]),b,1);busy(false);openOffice(d)}catch(e){fail('تعذر إنشاء قالب Excel',e)}});
wire('nw-xls',async()=>{try{busy(true,'جاري إنشاء الجدول…');const buf=await blankXlsx();const d=await createRec('xls','جدول '+new Date().toLocaleDateString('ar-EG'),buf,1);busy(false);openOffice(d)}catch(e){fail('تعذر إنشاء الجدول',e)}});

// ---------- PDF export (raster pages) ----------
const AR=/[\u0590-\u08FF\uFB1D-\uFEFF]/;
async function elsToPdf(els){const h2c=await need('h2c'),L=window.PDFLib,pdf=await L.PDFDocument.create();
 for(const el of els){const cv=await h2c(el,{scale:2,backgroundColor:'#ffffff',useCORS:true,logging:false});
  const jb=await new Promise(r=>cv.toBlob(r,'image/jpeg',.92)),im=await pdf.embedJpg(await jb.arrayBuffer());
  const w=el.offsetWidth*.75,h=el.offsetHeight*.75,pg=pdf.addPage([w,h]);pg.drawImage(im,{x:0,y:0,width:w,height:h});cv.width=cv.height=0}
 return new Blob([await pdf.save()],{type:'application/pdf'})}
async function longToPdf(el,ratio){const h2c=await need('h2c'),L=window.PDFLib,pdf=await L.PDFDocument.create();
 const cv=await h2c(el,{scale:2,backgroundColor:'#ffffff',useCORS:true,logging:false});
 const W=cv.width,PH=Math.floor(W*ratio),pw=595*(ratio<1?1.414:1),ph=pw*ratio;
 for(let y=0;y<cv.height;y+=PH){const h=Math.min(PH,cv.height-y),s=E('canvas');s.width=W;s.height=PH;const c=s.getContext('2d');c.fillStyle='#fff';c.fillRect(0,0,W,PH);c.drawImage(cv,0,y,W,h,0,0,W,h);
  const jb=await new Promise(r=>s.toBlob(r,'image/jpeg',.92)),im=await pdf.embedJpg(await jb.arrayBuffer());const pg=pdf.addPage([pw,ph]);pg.drawImage(im,{x:0,y:0,width:pw,height:ph});s.width=s.height=0}
 cv.width=cv.height=0;return new Blob([await pdf.save()],{type:'application/pdf'})}
const withZoom=async(el,f)=>{const z=el.style.zoom;el.style.zoom=1;try{return await f()}finally{el.style.zoom=z}};
function zoomBtns(getEl){return [{t:'＋',f:()=>{O.z=Math.min(3,O.z*1.2);getEl().style.zoom=O.z},title:'تكبير'},{t:'－',f:()=>{O.z=Math.max(.3,O.z/1.2);getEl().style.zoom=O.z},title:'تصغير'}]}

// ====================== WORD ======================
async function wordView(){
 const dp=await need('dp');$('#ofm').className='ofscroll';$('#ofm').innerHTML='<div class="ofw"></div>';const w=$('#ofm').firstChild;
 await dp.renderAsync(O.buf.slice(0),w,null,{inWrapper:true,breakPages:true,useBase64URL:true,renderHeaders:true,renderFooters:true,ignoreLastRenderedPageBreak:false});
 const secs=[...w.querySelectorAll('section.docx')],mw=Math.max(1,...secs.map(s=>s.offsetWidth));O.z=Math.min(1,($('#ofm').clientWidth-8)/mw);w.style.zoom=O.z;O.w=w;
 setBar([{t:'✏️ تعديل',f:()=>openOffice(O.d,{edit:1})},{t:'⬇ DOCX',f:()=>dl(new Blob([O.buf],{type:MIME.doc}),O.d.n+'.docx')},'|',...zoomBtns(()=>w)])}

async function wordEdit(){
 let html='<p><br></p>';
 if(!O.opt.newDoc){let hi=null;
   try{if(window.WordPro&&window.WordPro.importDocx)hi=await window.WordPro.importDocx(O.buf.slice(0))}catch(e){console.warn('importDocx failed, falling back',e)}
   if(hi){html=hi;toast('اتحافظ على تنسيق المستند الأصلي (خطوط، ألوان، مسافات، جداول، صور). حفظك هيتعمل كنسخة جديدة.')}
   else{const mm=await need('mammoth');const r=await mm.convertToHtml({arrayBuffer:O.buf.slice(0)},{convertImage:mm.images.dataUri});html=r.value||html;toast('وضع التعديل بيبسّط التنسيق. حفظك هيتعمل كنسخة جديدة.')}}
 O.newDoc=!!O.opt.newDoc;$('#ofm').className='ofscroll';$('#ofm').innerHTML='<div class="ofpg"><div id="oed" contenteditable="true" spellcheck="true"></div></div>';
 const ed=$('#oed');ed.innerHTML=html;ed.addEventListener('input',()=>O.dirty=1);
 let rg=null;document.addEventListener('selectionchange',O.sc=()=>{const s=getSelection();if(s.rangeCount&&ed.contains(s.anchorNode))rg=s.getRangeAt(0).cloneRange()});
 const restore=()=>{ed.focus();if(rg){const s=getSelection();s.removeAllRanges();s.addRange(rg)}};
 const wp=window.WordPro?window.WordPro.mount({ed,O,restore,range:()=>rg,buf:O.buf,newDoc:O.newDoc}):{items:{dir:[],table:[],image:[],page:[]}};
 const ex=(c,v)=>()=>{restore();document.execCommand(c,false,v);O.dirty=1};
 const sel=(opts,ph,f)=>{const s=E('select');s.innerHTML='<option value="">'+ph+'</option>'+opts.map(o=>'<option>'+o+'</option>').join('');s.onchange=()=>{if(s.value){f(s.value);s.value=''}};return {el:s}};
 const col=(ph,cmd)=>colorBtn({label:ph,title:cmd==='hiliteColor'||cmd==='backColor'?'لون التظليل':'لون النص',init:cmd==='hiliteColor'||cmd==='backColor'?'#FFFF00':'#FF0000',onPick:c=>{restore();document.execCommand(cmd,false,c);O.dirty=1}});
 const FONTS=['Calibri','Arial','Times New Roman','Tahoma','Verdana','Georgia','Courier New','Segoe UI','Cambria','Trebuchet MS','Traditional Arabic','Simplified Arabic','Sakkal Majalla','Arabic Typesetting','Cairo','Amiri','Tajawal','Almarai','Noto Naskh Arabic','Noto Kufi Arabic'];
 const SIZES=[8,9,10,11,12,14,16,18,20,24,28,32,36,48,72];
 const blockOf=()=>{let n=getSelection().anchorNode;while(n&&n!==ed&&!(n.nodeType===1&&/^(P|DIV|LI|H[1-6]|TD|TH|BLOCKQUOTE)$/.test(n.tagName)))n=n.parentNode;return n&&n!==ed?n:null};
 const replaceText=async()=>{const a=await raAsk('ابحث عن النص');if(!a)return;const b=await raAsk('استبدال بـ');if(b===null)return;let count=0;const w=document.createTreeWalker(ed,NodeFilter.SHOW_TEXT);const ns=[];while(w.nextNode())ns.push(w.currentNode);ns.forEach(n=>{if(n.nodeValue.includes(a)){const parts=n.nodeValue.split(a);count+=parts.length-1;n.nodeValue=parts.join(b)}});if(count)O.dirty=1;toast(count?'تم استبدال '+count+' موضع':'ملقتش النص')};
 const stats=()=>{const t=(ed.innerText||'').trim(),words=t? t.split(/\s+/).length:0;toast('الكلمات: '+words+' · الأحرف: '+t.length)};
 const insertTOC=()=>{const hs=[...ed.querySelectorAll('h1,h2,h3')].map(h=>({tag:h.tagName,text:(h.innerText||'').trim()})).filter(x=>x.text);if(!hs.length){toast('استخدم أنماط عنوان 1 أو 2 أو 3 أولاً');return}const html='<div data-reader-toc="1"><h2>جدول المحتويات</h2><ol>'+hs.map(x=>'<li style="margin-right:'+((+x.tag.slice(1)-1)*18)+'px">'+esc(x.text)+'</li>').join('')+'</ol></div><p><br></p>';restore();document.execCommand('insertHTML',false,html);O.dirty=1;toast('تم إدراج جدول المحتويات؛ حدّثه بإعادة إدراجه بعد تعديل العناوين')};
 const insertDate=()=>{restore();document.execCommand('insertText',false,new Date().toLocaleDateString('ar-EG'));O.dirty=1};
 setBar([
  {t:'💾',title:'حفظ',f:save},{t:'⬇ DOCX',f:async()=>{await save(1)}},'|',
  {t:'↶',f:ex('undo'),title:'تراجع'},{t:'↷',f:ex('redo'),title:'إعادة'},'|',
  {t:'<b>B</b>',f:ex('bold')},{t:'<i>I</i>',f:ex('italic')},{t:'<u>U</u>',f:ex('underline')},{t:'<s>S</s>',f:ex('strikeThrough')},'|',
  sel(FONTS,'الخط',v=>{restore();document.execCommand('fontName',false,v);O.dirty=1}),
  sel(SIZES,'الحجم',v=>{restore();document.execCommand('fontSize',false,'7');ed.querySelectorAll('font[size="7"]').forEach(f=>{f.removeAttribute('size');f.style.fontSize=v+'pt'});O.dirty=1}),
  (()=>{const n=E('input');n.type='number';n.min=1;n.max=999;n.step=1;n.inputMode='numeric';n.placeholder='حجم';n.title='حجم الخط بالرقم (pt)';n.style.cssText='flex:none;width:64px;padding:6px 6px;font-size:14px;border-radius:10px;border:1px solid var(--line);background:var(--card);color:var(--fg);text-align:center';const ap=()=>{const v=+n.value;if(!(v>0))return;restore();document.execCommand('fontSize',false,'7');ed.querySelectorAll('font[size="7"]').forEach(f=>{f.removeAttribute('size');f.style.fontSize=Math.min(999,v)+'pt'});O.dirty=1};n.onchange=ap;n.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();ap()}};return {el:n}})(),
  col('A لون','foreColor'),col('🖍 تظليل','hiliteColor'),colorBtn({label:'▦ لون الخلية',title:'لون الخلية الحالية',init:'#FFF2CC',onPick:c=>{restore();const n=getSelection().anchorNode,el=n&&(n.nodeType===1?n:n.parentElement),td=el&&el.closest&&el.closest('td,th');if(!td||!ed.contains(td)){alert('اضغط جوه خلية في الجدول الأول.');return}td.style.background=c;O.dirty=1}}),'|',
  {t:'⬅',f:ex('justifyLeft'),title:'يسار'},{t:'↔',f:ex('justifyCenter'),title:'وسط'},{t:'➡',f:ex('justifyRight'),title:'يمين'},{t:'☰',f:ex('justifyFull'),title:'ضبط'},
  ...wp.items.dir,'|',
  {t:'• قائمة',f:ex('insertUnorderedList')},{t:'1. قائمة',f:ex('insertOrderedList')},
  sel(['عنوان 1','عنوان 2','عنوان 3','عنوان 4','فقرة','اقتباس'],'النمط',v=>{restore();document.execCommand('formatBlock',false,{'عنوان 1':'H1','عنوان 2':'H2','عنوان 3':'H3','عنوان 4':'H4','فقرة':'P','اقتباس':'BLOCKQUOTE'}[v]);O.dirty=1}),'|',
  {t:'x²',title:'مرتفع',f:ex('superscript')},{t:'x₂',title:'منخفض',f:ex('subscript')},{t:'مسافة بادئة +',f:ex('indent')},{t:'مسافة بادئة −',f:ex('outdent')},{t:'مسح التنسيق',f:ex('removeFormat')},
  ...wp.items.table,...wp.items.image,{t:'🎙 إملاء صوتي',f:()=>{const SR=window.SpeechRecognition||window.webkitSpeechRecognition;if(!SR){toast('الإملاء الصوتي غير مدعوم على هذا الجهاز أو المتصفح');return}const rec=new SR();rec.lang='ar-EG';rec.interimResults=false;rec.maxAlternatives=1;restore();toast('اتكلم الآن…');rec.onresult=e=>{const t=e.results&&e.results[0]&&e.results[0][0]&&e.results[0][0].transcript;if(t){restore();document.execCommand('insertText',false,t+' ');O.dirty=1}};rec.onerror=e=>toast('تعذر الإملاء الصوتي: '+(e.error||'حاول مرة أخرى'));rec.start()}},'|',
  ...wp.items.page,{t:'📑 فهرس تلقائي',f:insertTOC},{t:'📅 تاريخ',f:insertDate},{t:'بحث/استبدال',f:replaceText},{t:'🔢 إحصاء',f:stats},{t:'⬇ PDF',f:pdf},{t:'🖨 طباعة',f:printOffice},{t:'🔐 نسخة مشفرة',f:encryptedExport}]);
 async function save(dlOnly){try{busy(true,'جاري الحفظ…');const dx=await need('dx');const blob=await htmlToDocx(dx,ed,{live:1}),buf=await blob.arrayBuffer();await saveBuf(buf);busy(false);if(dlOnly===1)dl(new Blob([buf],{type:MIME.doc}),O.d.n+'.docx');else toast('اتحفظ ✅')}catch(e){fail('فشل الحفظ',e)}}
 async function pdf(){let box;try{busy(true,'جاري تحويل PDF…');box=E('div','ofprint');box.innerHTML=ed.innerHTML;const ratio=window.WordPro?window.WordPro.printBox(box):1.4142;document.body.append(box);const b=await longToPdf(box,ratio);box.remove();busy(false);dl(b,O.d.n+'.pdf')}catch(e){box&&box.remove();fail('فشل التصدير',e)}}
 O.cleanup=()=>{document.removeEventListener('selectionchange',O.sc);wp.cleanup&&wp.cleanup()}}

// ---- HTML -> DOCX ----
const hex=c=>{const m=/rgba?\((\d+)[, ]+(\d+)[, ]+(\d+)(?:[, /]+([\d.]+))?/.exec(c||'');if(!m)return null;if(m[4]!==undefined&&+m[4]===0)return null;return [m[1],m[2],m[3]].map(x=>(+x).toString(16).padStart(2,'0')).join('').toUpperCase()};
const GEN=/^(serif|sans-serif|monospace|system-ui|cursive|fantasy|ui-.*)$/i;
function firstStrong(t){for(const ch of t){if(AR.test(ch))return 1;if(/[A-Za-z]/.test(ch))return 0}return -1}
function runStyle(el,root){const cs=getComputedStyle(el),st={};
 if(+cs.fontWeight>=600)st.bold=true;if(cs.fontStyle==='italic')st.italics=true;
 let u=0,s=0,bg=null;for(let e=el;e&&e!==root;e=e.parentElement){const c=getComputedStyle(e);if(/underline/.test(c.textDecorationLine))u=1;if(/line-through/.test(c.textDecorationLine))s=1;if(!bg&&/^inline/.test(c.display)){const b=hex(c.backgroundColor);if(b)bg=b}}
 if(u)st.underline={};if(s)st.strike=true;if(bg)st.shading={type:'clear',fill:bg,color:'auto'};
 const c=hex(cs.color);if(c&&c!=='000000')st.color=c;
 const sz=parseFloat(cs.fontSize);if(sz)st.size=Math.max(2,Math.round(sz*1.5));
 const fn=cs.fontFamily.split(',')[0].replace(/["']/g,'').trim();if(fn&&!GEN.test(fn))st.font={ascii:fn,hAnsi:fn,cs:fn,eastAsia:fn};return st}
function imgRun(dx,im,root){const m=/^data:image\/(png|jpe?g|gif|bmp);base64,(.*)$/i.exec(im.src||'');if(!m)return null;
 const bin=atob(m[2]),u=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)u[i]=bin.charCodeAt(i);
 const nw=im.naturalWidth||im.width||300,nh=im.naturalHeight||im.height||200,W=window.WordPro;let w=nw,h=nh;
 if(W&&root&&im.isConnected&&root.clientWidth){const cs=getComputedStyle(root),cw=root.clientWidth-(parseFloat(cs.paddingLeft)||0)-(parseFloat(cs.paddingRight)||0),rw=im.getBoundingClientRect().width;if(cw>0&&rw>0){w=rw/cw*W.contentPx();h=nh*w/nw}}
 else if(W&&im.dataset&&im.dataset.wp){w=+im.dataset.wp/100*W.contentPx();h=nh*w/nw}
 const cap=W?W.contentPx():520;if(w>cap){h=h*cap/w;w=cap}
 const t=/jpe?g/i.test(m[1])?'jpg':m[1].toLowerCase();return new dx.ImageRun({data:u,transformation:{width:Math.round(w),height:Math.round(h)},type:t})}
function runs(dx,nodes,root,out){
 for(const n of nodes){
  if(n.nodeType===3){const t=n.nodeValue.replace(/\u00a0/g,' ').replace(/[\r\n\t]+/g,' ');if(!t)continue;const p=n.parentElement;out.push(new dx.TextRun({text:t,...runStyle(p,root),rightToLeft:AR.test(t)||undefined}))}
  else if(n.nodeType===1){const tg=n.tagName;if(tg==='BR'){out.push(new dx.TextRun({break:1}));continue}
   if(tg==='IMG'){const r=imgRun(dx,n,root);r&&out.push(r);continue}
   if(/^(UL|OL|TABLE)$/.test(tg))continue;
   runs(dx,n.childNodes,root,out)}}}
const ALIGN=(dx,el,bidi)=>{const cs=getComputedStyle(el);let a=cs.textAlign,vis=a==='center'?'center':a==='justify'?'justify':a==='right'||(a==='start'&&bidi)||(a==='end'&&!bidi)?'right':'left';
 const def=bidi?'right':'left';if(vis===def)return undefined;if(vis==='center')return dx.AlignmentType.CENTER;if(vis==='justify')return dx.AlignmentType.BOTH;
 return bidi?(vis==='right'?dx.AlignmentType.LEFT:dx.AlignmentType.RIGHT):(vis==='right'?dx.AlignmentType.RIGHT:dx.AlignmentType.LEFT)};
function para(dx,el,nodes,root,x){
 nodes=[...nodes];while(nodes.length&&nodes[nodes.length-1].nodeName==='BR')nodes.pop();
 const out=[];runs(dx,nodes,root,out);const txt=nodes.map(n=>n.textContent||'').join(''),fs=firstStrong(txt);
 const dc=el.closest&&el.closest('[dir]'),dv=dc&&dc.getAttribute('dir'),bidi=dv==='rtl'?true:dv==='ltr'?false:fs===1;
 const cs=getComputedStyle(el),H={1:dx.HeadingLevel.HEADING_1,2:dx.HeadingLevel.HEADING_2,3:dx.HeadingLevel.HEADING_3,4:dx.HeadingLevel.HEADING_4,5:dx.HeadingLevel.HEADING_5,6:dx.HeadingLevel.HEADING_6};
 const ist=el.style||{},px=v=>Math.round((parseFloat(v)||0)*15),isLi=el.tagName==='LI',spc={after:px(cs.marginBottom)};
  const o={children:out,bidirectional:bidi||undefined,alignment:ALIGN(dx,el,bidi),spacing:spc,pageBreakBefore:(PB||(el&&el.dataset&&el.dataset.pb==='1'))||undefined,...x};
  if(ist.marginTop&&!isLi)spc.before=px(cs.marginTop);
  if(ist.lineHeight&&cs.lineHeight!=='normal'){const fs=parseFloat(cs.fontSize)||16;if(/pt$/.test(ist.lineHeight)){spc.line=Math.round(parseFloat(ist.lineHeight)*20);spc.lineRule='exact'}else{spc.line=Math.max(120,Math.round(parseFloat(cs.lineHeight)/fs*240));spc.lineRule='auto'}}
  if(!isLi){const ind={};if(ist.marginInlineStart||ist.marginInlineEnd||ist.marginLeft||ist.marginRight){const ml=px(cs.marginLeft),mr=px(cs.marginRight);if(bidi){ind.start=mr;ind.end=ml}else{ind.start=ml;ind.end=mr}}
   if(ist.textIndent){const ti=px(cs.textIndent);if(ti>0)ind.firstLine=ti;else if(ti<0)ind.hanging=-ti}if(Object.keys(ind).length)o.indent=ind}
  const pbg=(ist.backgroundColor||ist.background)&&hex(cs.backgroundColor);if(pbg)o.shading={type:'clear',fill:pbg,color:'auto'};
  {const bd={};for(const [S,k] of [['Top','top'],['Bottom','bottom'],['Left','left'],['Right','right']]){const w=parseFloat(cs['border'+S+'Width'])||0;if(w&&cs['border'+S+'Style']!=='none'&&ist['border'+S])bd[k]={style:'single',size:Math.max(2,Math.round(w*6)),color:hex(cs['border'+S+'Color'])||'000000',space:2}}if(Object.keys(bd).length)o.border=bd}
 const m=/^H([1-6])$/.exec(el.tagName||'');if(m&&!x.noHead)o.heading=H[+m[1]];delete o.noHead;PB=false;return new dx.Paragraph(o)}
const BLK=/^(P|DIV|H[1-6]|UL|OL|LI|TABLE|BLOCKQUOTE|PRE|SECTION)$/;
const hasBlk=n=>[...n.children].some(c=>BLK.test(c.tagName));
let LI=0,PB=false;
function lst(dx,el,root,out,lvl,inst,ord){for(const li of el.children){if(li.tagName!=='LI')continue;const sub=[],nodes=[];
  for(const c of li.childNodes){if(c.nodeType===1&&/^(UL|OL)$/.test(c.tagName))sub.push(c);else nodes.push(c)}
  out.push(para(dx,li,nodes,root,ord?{numbering:{reference:'num',level:Math.min(lvl,3),instance:inst}}:{bullet:{level:Math.min(lvl,3)}}));
  sub.forEach(s=>{const o2=s.tagName==='OL';lst(dx,s,root,out,lvl+1,o2?++LI:0,o2)})}}
function blocks(dx,parent,root,out){let pend=[];
 const flush=()=>{if(pend.some(n=>n.nodeType===1||/\S/.test(n.nodeValue||'')))out.push(para(dx,parent,pend,root,{}));pend=[]};
 for(const n of parent.childNodes){
  if(n.nodeType===1&&n.dataset&&n.dataset.pageBreak==='1'){flush();PB=true;continue}
  if(n.nodeType===1&&BLK.test(n.tagName)){flush();const tg=n.tagName;
   if(tg==='UL'||tg==='OL')lst(dx,n,root,out,0,tg==='OL'?++LI:0,tg==='OL');
   else if(tg==='TABLE'){if(PB){out.push(new dx.Paragraph({pageBreakBefore:true}));PB=false}out.push(tbl(dx,n,root))}
   else if(hasBlk(n)&&!/^H/.test(tg))blocks(dx,n,root,out);
   else out.push(para(dx,n,n.childNodes,root,{}))}
  else if(n.nodeType===1&&n.tagName==='TABLE'){flush();out.push(tbl(dx,n,root))}
  else pend.push(n)}
 flush()}
function tbl(dx,t,root){const W=window.WordPro,live=!!W&&t.isConnected,M=live?W.measure(t,root):null,hint=AR.test(t.textContent||'')&&firstStrong(t.textContent||'')===1,
  rtl=W?(()=>{const dc=t.closest('[dir]'),v=dc&&dc.getAttribute('dir');return v==='rtl'?true:v==='ltr'?false:hint})():hint;
 const rows=[...t.rows].map(tr=>{const cells=[...tr.cells].map(td=>{const ch=[];blocks(dx,td,root,ch);if(!ch.length)ch.push(new dx.Paragraph({}));
  const bg=hex(getComputedStyle(td).backgroundColor),o={children:ch};if(td.colSpan>1)o.columnSpan=td.colSpan;if(td.rowSpan>1)o.rowSpan=td.rowSpan;if(bg)o.shading={type:'clear',fill:bg,color:'auto'};
  if(live){o.borders=W.cellBorders(dx,td,rtl);const va=td.style.verticalAlign;if(va==='middle')o.verticalAlign=dx.VerticalAlign.CENTER;else if(va==='bottom')o.verticalAlign=dx.VerticalAlign.BOTTOM;if(M){const w=M.cellW(td);if(w)o.width={size:w,type:dx.WidthType.DXA}}}
  return new dx.TableCell(o)});const ro={children:cells};if(tr.dataset&&tr.dataset.hdr==='1')ro.tableHeader=true;return new dx.TableRow(ro)});
 const o={rows,visuallyRightToLeft:rtl};
 if(M){o.width={size:M.total,type:dx.WidthType.DXA};o.columnWidths=M.cols;o.layout=dx.TableLayoutType.FIXED}else o.width={size:100,type:dx.WidthType.PERCENTAGE};
 if(live){o.borders=dx.TableBorders.NONE;const al=W.tableAlign(dx,t,rtl);if(al)o.alignment=al}
 return new dx.Table(o)}
async function htmlToDocx(dx,ed,opt){LI=0;PB=false;const ch=[];blocks(dx,ed,ed,ch);if(PB){ch.push(new dx.Paragraph({pageBreakBefore:true}));PB=false}if(!ch.length)ch.push(new dx.Paragraph({}));
 const doc=new dx.Document({styles:{default:{document:{run:{font:{ascii:'Calibri',hAnsi:'Calibri',cs:'Arial',eastAsia:'Calibri'},size:24}}}},
  numbering:{config:[{reference:'num',levels:[0,1,2,3].map(l=>({level:l,format:dx.LevelFormat.DECIMAL,text:'%'+(l+1)+'.',alignment:dx.AlignmentType.START,style:{paragraph:{indent:{left:720*(l+1),hanging:360}}}}))}]},
  sections:[Object.assign({children:ch},window.WordPro?window.WordPro.section(dx,opt&&opt.live?O:null):{properties:{}})]});
 return dx.Packer.toBlob(doc)}

// ====================== EXCEL ======================
const RH=26,HW=46,HH=26,DCW=84;
const colName=i=>{let s='';i++;while(i>0){const m=(i-1)%26;s=String.fromCharCode(65+m)+s;i=((i-1)/26)|0}return s};
const colIdx=s=>{let n=0;for(const ch of s)n=n*26+ch.charCodeAt(0)-64;return n-1};
const K=(r,c)=>r+','+c;
const parseA=a=>{const m=/^\$?([A-Z]+)\$?(\d+)$/.exec(a);return m?{r:+m[2]-1,c:colIdx(m[1])}:null};
const esc=s=>String(s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const newSheet=n=>({name:n,cells:{},cw:{},merges:[],NR:100,NC:20});
const THEME=['#FFFFFF','#000000','#E7E6E6','#44546A','#4472C4','#ED7D31','#A5A5A5','#FFC000','#5B9BD5','#70AD47'];
const colr=c=>{if(!c)return null;if(c.argb)return '#'+String(c.argb).slice(-6);if(c.theme!=null)return THEME[c.theme]||null;return null};
const d2s=d=>Math.round((d.getTime()/864e5+25569)*1e6)/1e6;
let XS=null;const curSh=()=>XS.X.sheets[XS.cur];
function convCell(ex,cell){const V=ex.ValueType,t=cell.type,o={};if(t===V.Merge)return null;
 if(t===V.Formula){o.f=cell.formula;let v=cell.result;if(v&&typeof v==='object'){v=v.error?v.error:v instanceof Date?d2s(v):undefined}o.v=v;o.cv=v}
 else if(t===V.Date)o.v=d2s(cell.value);
 else if(t===V.RichText||t===V.Hyperlink)o.v=cell.text;
 else if(t===V.Error)o.v=cell.value.error;
 else if(t!==V.Null)o.v=cell.value;
 const s={},f=cell.font||{};
 if(f.bold)s.b=1;if(f.italic)s.i=1;if(f.underline)s.u=1;if(f.strike)s.st=1;const fc=colr(f.color);if(fc&&fc!=='#000000')s.c=fc;if(f.size&&f.size!==11)s.fs=f.size;if(f.name&&f.name!=='Calibri')s.fn=f.name;
 const fl=cell.fill;if(fl&&fl.type==='pattern'&&fl.pattern==='solid'){const bg=colr(fl.fgColor);if(bg&&bg!=='#FFFFFF')s.bg=bg}
 const al=cell.alignment||{};if(al.horizontal)s.al=al.horizontal==='centerContinuous'?'center':al.horizontal;if(al.vertical)s.va=al.vertical;if(al.wrapText)s.w=1;
 if(cell.numFmt&&cell.numFmt!=='General')s.nf=cell.numFmt;
 const b=cell.border;if(b&&(b.top||b.bottom||b.left||b.right))s.bd=1;
 if(Object.keys(s).length)o.s=s;
 if(o.v===undefined&&o.f===undefined&&!o.s)return null;return o}
async function xlsParse(buf){const ex=await need('xl'),wb=new ex.Workbook();await wb.xlsx.load(buf);const M={sheets:[]};
 wb.eachSheet(ws=>{const sh=newSheet(ws.name);let R=0,C=0;
  ws.eachRow({includeEmpty:false},(row,r)=>{row.eachCell({includeEmpty:true},(cell,c)=>{const o=convCell(ex,cell);if(o){sh.cells[K(r-1,c-1)]=o;R=Math.max(R,r);C=Math.max(C,c)}})});
  for(let c=1;c<=C;c++){const w=ws.getColumn(c).width;if(w)sh.cw[c-1]=Math.round(w*7+5)}
  ((ws.model&&ws.model.merges)||[]).forEach(m=>{const [a,b]=m.split(':'),p=parseA(a),q=parseA(b);if(p&&q)sh.merges.push([p.r,p.c,q.r,q.c])});
  sh.NR=Math.max(R+60,100);sh.NC=Math.min(300,Math.max(C+6,20));M.sheets.push(sh)});
 if(!M.sheets.length)M.sheets.push(newSheet('Sheet1'));return M}
function applySt(cell,s){const f={};if(s.b)f.bold=true;if(s.i)f.italic=true;if(s.u)f.underline=true;if(s.st)f.strike=true;if(s.c)f.color={argb:'FF'+s.c.slice(1)};
 if(Object.keys(f).length||s.fs||s.fn){f.name=s.fn||'Calibri';f.size=s.fs||11;cell.font=f}
 if(s.bg)cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF'+s.bg.slice(1)}};
 if(s.al||s.w||s.va)cell.alignment={horizontal:s.al,vertical:s.va,wrapText:!!s.w};
 if(s.nf)cell.numFmt=s.nf;
 if(s.bd){const t={style:'thin',color:{argb:'FF000000'}};cell.border={top:t,left:t,bottom:t,right:t}}}
async function xlsBuild(M){const ex=await need('xl'),wb=new ex.Workbook();
 for(const sh of M.sheets){const ws=wb.addWorksheet(sh.name);
  for(const k in sh.cells){const [r,c]=k.split(',').map(Number),o=sh.cells[k],cell=ws.getCell(r+1,c+1);
   if(o.f!==undefined){const v=o.v;cell.value={formula:o.f,result:typeof v==='string'&&v[0]==='#'?{error:v}:v}}else if(o.v!==undefined)cell.value=o.v;
   if(o.s)applySt(cell,o.s)}
  for(const c in sh.cw)ws.getColumn(+c+1).width=Math.max(2,(sh.cw[c]-5)/7);
  for(const m of sh.merges){try{ws.mergeCells(m[0]+1,m[1]+1,m[2]+1,m[3]+1)}catch(e){}}}
 const b=await wb.xlsx.writeBuffer();return b instanceof ArrayBuffer?b:b.buffer.slice(b.byteOffset,b.byteOffset+b.byteLength)}
function parseCsv(t){t=t.replace(/^\uFEFF/,'');const l1=t.split('\n')[0],D=[',',';','\t'].sort((a,b)=>l1.split(b).length-l1.split(a).length)[0];
 const rows=[];let r=[],f='',q=0;for(let i=0;i<t.length;i++){const ch=t[i];
  if(q){if(ch==='"'){if(t[i+1]==='"'){f+='"';i++}else q=0}else f+=ch}
  else if(ch==='"')q=1;else if(ch===D){r.push(f);f=''}else if(ch==='\n'){r.push(f);rows.push(r);r=[];f=''}else if(ch!=='\r')f+=ch}
 if(f!==''||r.length){r.push(f);rows.push(r)}return rows}
async function csvToXlsx(text){const sh=newSheet('Sheet1');let C=0;const rows=parseCsv(text);
 rows.forEach((r,i)=>{r.forEach((v,j)=>{if(v==='')return;const p=parseInput(v);sh.cells[K(i,j)]=p.f!==undefined?{f:p.f}:{v:p.v,...(p.nf?{s:{nf:p.nf}}:{})};C=Math.max(C,j+1)})});
 sh.NR=Math.max(rows.length+60,100);sh.NC=Math.max(C+6,20);return xlsBuild({sheets:[sh]})}
function parseInput(raw){raw=String(raw);if(raw[0]==='='&&raw.length>1)return {f:raw.slice(1)};if(raw==='')return {v:undefined};
 const t=raw.trim().replace(/,/g,'').replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d));
 if(/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(t))return {v:+t};
 if(/^[+-]?(\d+\.?\d*|\.\d+)%$/.test(t))return {v:parseFloat(t)/100,nf:'0%'};
 if(/^(true|false)$/i.test(raw))return {v:raw.toLowerCase()==='true'};return {v:raw}}
// ---- formulas ----
async function recalc(sh,full){const fp=await need('fp'),P=new fp.Parser(),memo={},stk={};
 const val=(r,c)=>{const k=K(r,c),o=sh.cells[k];if(!o)return null;if(o.f===undefined)return o.v===undefined?null:o.v;
  if(!full&&o.v!==undefined)return o.v;if(k in memo)return memo[k];if(stk[k])return '#REF!';
  stk[k]=1;let x=P.parse(o.f),v=x.error?x.error:x.result;stk[k]=0;
  if(Array.isArray(v))v=Array.isArray(v[0])?v[0][0]:v[0];if(v instanceof Error)v='#VALUE!';
  if(x.error&&o.cv!==undefined&&o.cv!==null)v=o.cv;if(v===undefined)v=null;
  o.v=v===null?undefined:v;memo[k]=v;return v};
 P.on('callCellValue',(cc,done)=>done(val(cc.row.index,cc.column.index)));
 P.on('callRangeValue',(s,e,done)=>{const A=[];for(let r=s.row.index;r<=e.row.index;r++){const row=[];for(let c=s.column.index;c<=e.column.index;c++)row.push(val(r,c));A.push(row)}done(A)});
 for(const k in sh.cells){const o=sh.cells[k];if(o.f!==undefined&&(full||o.v===undefined)){const [r,c]=k.split(',').map(Number);val(r,c)}}}
const REF=/(?<![A-Za-z_.\d])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])/g;
const mapF=(f,fn)=>f.split(/("(?:[^"]|"")*")/).map((p,i)=>i%2?p:p.replace(REF,(m,a,c,b,r)=>fn(a,colIdx(c),b,+r-1))).join('');
const shiftF=(f,dr,dc)=>mapF(f,(a,ci,b,ri)=>{ci+=a?0:dc;ri+=b?0:dr;if(ci<0||ri<0)return '#REF!';return a+colName(ci)+b+(ri+1)});
// ---- number formats ----
function fmtNum(v,nf){const m=nf.split(';')[0],plain=m.replace(/\[[^\]]*\]|"[^"]*"/g,'');
 if(/[ymdhs]/i.test(plain)&&!/[#0]/.test(plain)){const d=new Date(Math.round((v-25569)*864e5)),p=n=>String(n).padStart(2,'0'),MN=['January','February','March','April','May','June','July','August','September','October','November','December'];
  return plain.replace(/yyyy|yy|mmmm|mmm|mm|m|dd|d|hh|h|ss|s|AM\/PM/gi,t=>{switch(t.toLowerCase()){case 'yyyy':return d.getUTCFullYear();case 'yy':return p(d.getUTCFullYear()%100);case 'mmmm':return MN[d.getUTCMonth()];case 'mmm':return MN[d.getUTCMonth()].slice(0,3);case 'mm':return p(d.getUTCMonth()+1);case 'm':return d.getUTCMonth()+1;case 'dd':return p(d.getUTCDate());case 'd':return d.getUTCDate();case 'hh':return p(d.getUTCHours());case 'h':return d.getUTCHours();case 'ss':return p(d.getUTCSeconds());case 's':return d.getUTCSeconds();default:return d.getUTCHours()<12?'AM':'PM'}})}
 if(m==='@')return String(v);
 const pct=/%/.test(plain),x=pct?v*100:v,dm=/\.([0#]+)/.exec(plain),dec=dm?dm[1].length:0;let s=Math.abs(x).toFixed(dec);
 if(/,/.test(plain)){const [i,f]=s.split('.');s=i.replace(/\B(?=(\d{3})+(?!\d))/g,',')+(f?'.'+f:'')}
 const lit=m.replace(/\[[^\]]*\]/g,'').replace(/"([^"]*)"/g,'$1').replace(/[_*\\]./g,''),a=lit.search(/[#0]/),b=lit.length-1-[...lit].reverse().findIndex(ch=>/[#0.,]/.test(ch));
 const pre=a>0?lit.slice(0,a).replace(/%/g,''):'',suf=a>=0?lit.slice(b+1).replace(/%/g,''):'';
 return (x<0?'-':'')+pre+s+(pct?'%':'')+suf}
function disp(o){if(!o)return '';const v=o.v;if(v===undefined||v===null)return '';if(typeof v==='boolean')return v?'TRUE':'FALSE';
 const nf=o.s&&o.s.nf;if(typeof v==='number'){if(nf)try{return fmtNum(v,nf)}catch(e){}return String(Number(v.toPrecision(12)))}return String(v)}
function cstyle(o,ar){const s=(o&&o.s)||{};let t='';
 if(s.bg)t+='background:'+s.bg+';';if(s.c)t+='color:'+s.c+';';if(s.b)t+='font-weight:700;';if(s.i)t+='font-style:italic;';
 if(s.u||s.st)t+='text-decoration:'+(s.u?'underline ':'')+(s.st?'line-through':'')+';';
 if(s.fs)t+='font-size:'+Math.round(s.fs*1.33)+'px;';if(s.fn)t+='font-family:"'+s.fn+'",sans-serif;';
 const al=s.al&&s.al!=='general'?s.al:(o&&typeof o.v==='number'?'right':'');if(al)t+='text-align:'+al+';';
 if(s.va)t+='vertical-align:'+(s.va==='middle'?'middle':s.va)+';';if(s.w&&ar)t+='white-space:normal;';
 if(s.bd)t+='border:1px solid #333;';return t}
// ---- open ----
async function xlsOpen(){const M=await xlsParse(O.buf.slice(0));
 XS={X:M,cur:0,sel:{r1:0,c1:0,r2:0,c2:0},undo:[],redo:[],clip:null,ext:0};
 for(const sh of M.sheets)await recalc(sh,false);
 const m=$('#ofm');m.className='ofx';m.innerHTML='<div id="xfb"><span id="xad">A1</span><input id="xin" dir="auto" autocomplete="off" enterkeyhint="done"><button id="xok" aria-label="تأكيد">✓</button></div><div id="xg"><div id="xi"></div></div><div id="xtabs"></div>';
 const g=$('#xg');g.addEventListener('scroll',()=>{if(XS.raf)return;XS.raf=requestAnimationFrame(()=>{XS.raf=0;const sh=curSh();if(g.scrollTop+g.clientHeight>HH+sh.NR*RH-300&&sh.NR<5000)sh.NR+=50;xrender()})},{passive:true});
 g.addEventListener('click',xclick);
 const inp=$('#xin');inp.addEventListener('focus',()=>{XS.edit={r:XS.sel.r1,c:XS.sel.c1,raw:inp.value}});
 inp.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();inp.blur();moveSel(1,0)}});
 inp.addEventListener('blur',async()=>{const ed=XS.edit;XS.edit=null;if(ed&&inp.value!==ed.raw)await commit(inp.value,ed.r,ed.c)});
 $('#xok').onclick=()=>{inp.blur();moveSel(1,0)};
 xbar();xtabs();xrender();loadFB()}
function rng(){const s=XS.sel;return {r1:Math.min(s.r1,s.r2),r2:Math.max(s.r1,s.r2),c1:Math.min(s.c1,s.c2),c2:Math.max(s.c1,s.c2)}}
const addrOf=()=>{const q=rng();return q.r1===q.r2&&q.c1===q.c2?colName(q.c1)+(q.r1+1):colName(q.c1)+(q.r1+1)+':'+colName(q.c2)+(q.r2+1)};
function loadFB(){const sh=curSh(),o=sh.cells[K(XS.sel.r1,XS.sel.c1)];$('#xad').textContent=addrOf();const i=$('#xin');if(document.activeElement!==i)i.value=o?(o.f!==undefined?'='+o.f:o.v===undefined?'':String(o.v)):''}
function moveSel(dr,dc){const sh=curSh(),s=XS.sel;s.r1=s.r2=Math.max(0,Math.min(sh.NR-1,s.r1+dr));s.c1=s.c2=Math.max(0,Math.min(sh.NC-1,s.c1+dc));xrender();loadFB()}
function xclick(e){const sh=curSh(),s=XS.sel;const td=e.target.closest('td[data-c]'),rh=e.target.closest('th[data-rh]'),ch=e.target.closest('[data-ch]');
 if(td){const r=+td.dataset.r,c=+td.dataset.c;
  if(XS.ext){s.r2=r;s.c2=c}else{const same=s.r1===r&&s.c1===c&&s.r2===r&&s.c2===c;s.r1=s.r2=r;s.c1=s.c2=c;if(same){loadFB();$('#xin').focus();return xrender()}}}
 else if(rh){s.r1=s.r2=+rh.dataset.rh;s.c1=0;s.c2=sh.NC-1}
 else if(ch){s.c1=s.c2=+ch.dataset.ch;s.r1=0;s.r2=sh.NR-1}
 else return;xrender();loadFB()}
function xrender(){const sh=curSh(),g=$('#xg'),xi=$('#xi');if(!g||!xi)return;
 const cw=c=>sh.cw[c]||DCW,xs=[0];for(let c=0;c<sh.NC;c++)xs.push(xs[c]+cw(c));const tw=HW+xs[sh.NC];
 xi.style.width=tw+'px';xi.style.height=(HH+sh.NR*RH)+'px';
 let r0=Math.max(0,Math.floor((g.scrollTop-HH)/RH)-4);const r1=Math.min(sh.NR,Math.ceil((g.scrollTop+g.clientHeight)/RH)+4);
 const mg={},cov={};for(const m of sh.merges){if(m[2]<r0||m[0]>=r1)continue;if(m[0]<r0&&m[2]>=r0)r0=m[0]}
 for(const m of sh.merges){if(m[2]<r0||m[0]>=r1)continue;mg[K(m[0],m[1])]=m;for(let r=m[0];r<=m[2];r++)for(let c=m[1];c<=m[3];c++)if(r!==m[0]||c!==m[1])cov[K(r,c)]=1}
 const q=rng(),a=XS.sel;
 let h='<div class="xh" style="width:'+tw+'px"><span class="xcorner" style="width:'+HW+'px"></span>';
 for(let c=0;c<sh.NC;c++)h+='<span data-ch="'+c+'" class="'+(c>=q.c1&&c<=q.c2?'xs':'')+'" style="width:'+cw(c)+'px">'+colName(c)+'</span>';
 h+='</div><table class="xt" style="top:'+(HH+r0*RH)+'px;width:'+tw+'px"><colgroup><col style="width:'+HW+'px">';
 for(let c=0;c<sh.NC;c++)h+='<col style="width:'+cw(c)+'px">';h+='</colgroup><tbody>';
 for(let r=r0;r<r1;r++){h+='<tr style="height:'+RH+'px"><th class="rh'+(r>=q.r1&&r<=q.r2?' xs':'')+'" data-rh="'+r+'">'+(r+1)+'</th>';
  for(let c=0;c<sh.NC;c++){const k=K(r,c);if(cov[k])continue;const o=sh.cells[k],m=mg[k];
   const on=r>=q.r1&&r<=q.r2&&c>=q.c1&&c<=q.c2,act=r===a.r1&&c===a.c1;
   h+='<td data-r="'+r+'" data-c="'+c+'" dir="auto" class="'+(on?'xs ':'')+(act?'xa ':'')+(o&&typeof o.v==='string'&&o.v[0]==='#'?'xe':'')+'"'+(m?' colspan="'+(m[3]-m[1]+1)+'" rowspan="'+(Math.min(m[2],r1-1)-m[0]+1)+'"':'')+(o&&o.s?' style="'+cstyle(o)+'"':o&&typeof o.v==='number'?' style="text-align:right"':'')+'>'+esc(disp(o))+'</td>'}
  h+='</tr>'}
 xi.innerHTML=h+'</tbody></table>'}
// ---- editing ----
const snap=()=>JSON.stringify({sheets:XS.X.sheets,cur:XS.cur});
function pushU(){const s=snap();XS.undo.push(s);while(XS.undo.length>(s.length>1.5e6?3:30))XS.undo.shift();XS.redo=[];O.dirty=1}
async function restoreSnap(s){const o=JSON.parse(s);XS.X.sheets=o.sheets;XS.cur=Math.min(o.cur,o.sheets.length-1);xtabs();xrender();loadFB()}
async function undo(){if(!XS.undo.length)return;XS.redo.push(snap());await restoreSnap(XS.undo.pop())}
async function redo(){if(!XS.redo.length)return;XS.undo.push(snap());await restoreSnap(XS.redo.pop())}
const cellAt=(sh,r,c)=>sh.cells[K(r,c)]||(sh.cells[K(r,c)]={});
async function commit(raw,r,c){const sh=curSh();pushU();const o=cellAt(sh,r,c),p=parseInput(raw);
 if(p.f!==undefined){o.f=p.f;delete o.v;delete o.cv}else{delete o.f;delete o.cv;if(p.v===undefined)delete o.v;else o.v=p.v;if(p.nf&&!(o.s&&o.s.nf))(o.s=o.s||{}).nf=p.nf}
 if(!o.f&&o.v===undefined&&!o.s)delete sh.cells[K(r,c)];
 await recalc(sh,true);xrender();loadFB()}
function eachSel(f){const sh=curSh(),q=rng();for(let r=q.r1;r<=q.r2;r++)for(let c=q.c1;c<=q.c2;c++)f(cellAt(sh,r,c),r,c)}
function style(fn){pushU();const sh=curSh(),a=cellAt(sh,rng().r1,rng().c1);a.s=a.s||{};const st=fn.pre?fn.pre(a.s):null;eachSel(o=>{o.s=o.s||{};fn(o.s,st)});xrender()}
const tog=k=>style(Object.assign((s,on)=>{if(on)s[k]=1;else delete s[k]},{pre:s=>!s[k]}));
const setS=(k,v)=>style((s)=>{if(v===null)delete s[k];else s[k]=v});
function merge(un){pushU();const sh=curSh(),q=rng();sh.merges=sh.merges.filter(m=>!(m[0]<=q.r2&&m[2]>=q.r1&&m[1]<=q.c2&&m[3]>=q.c1));
 if(!un&&(q.r1!==q.r2||q.c1!==q.c2)){sh.merges.push([q.r1,q.c1,q.r2,q.c2]);for(let r=q.r1;r<=q.r2;r++)for(let c=q.c1;c<=q.c2;c++)if(r!==q.r1||c!==q.c1){const o=sh.cells[K(r,c)];o&&(delete o.v,delete o.f)}}xrender()}
function fill(dir){pushU();const sh=curSh(),q=rng();
 if(dir==='d'){for(let c=q.c1;c<=q.c2;c++){const s=sh.cells[K(q.r1,c)];for(let r=q.r1+1;r<=q.r2;r++){if(!s){delete sh.cells[K(r,c)];continue}const o=JSON.parse(JSON.stringify(s));if(o.f!==undefined){o.f=shiftF(o.f,r-q.r1,0);delete o.v;delete o.cv}sh.cells[K(r,c)]=o}}}
 else{for(let r=q.r1;r<=q.r2;r++){const s=sh.cells[K(r,q.c1)];for(let c=q.c1+1;c<=q.c2;c++){if(!s){delete sh.cells[K(r,c)];continue}const o=JSON.parse(JSON.stringify(s));if(o.f!==undefined){o.f=shiftF(o.f,0,c-q.c1);delete o.v;delete o.cv}sh.cells[K(r,c)]=o}}}
 recalc(sh,true).then(xrender)}
function copy(cut){const sh=curSh(),q=rng(),cl={h:q.r2-q.r1+1,w:q.c2-q.c1+1,r0:q.r1,c0:q.c1,cells:{},cut:!!cut,sheet:XS.cur};
 for(let r=q.r1;r<=q.r2;r++)for(let c=q.c1;c<=q.c2;c++){const o=sh.cells[K(r,c)];if(o)cl.cells[K(r-q.r1,c-q.c1)]=JSON.parse(JSON.stringify(o))}XS.clip=cl;toast(cut?'اتقصّ':'اتنسخ')}
async function paste(){const cl=XS.clip;if(!cl)return toast('مفيش حاجة منسوخة');pushU();const sh=curSh(),q=rng();
 if(cl.cut&&cl.sheet===XS.cur){for(let r=0;r<cl.h;r++)for(let c=0;c<cl.w;c++)delete sh.cells[K(cl.r0+r,cl.c0+c)]}
 for(let r=0;r<cl.h;r++)for(let c=0;c<cl.w;c++){const s=cl.cells[K(r,c)],k=K(q.r1+r,q.c1+c);if(!s){delete sh.cells[k];continue}const o=JSON.parse(JSON.stringify(s));if(o.f!==undefined){o.f=shiftF(o.f,q.r1-cl.r0,q.c1-cl.c0);delete o.v;delete o.cv}sh.cells[k]=o}
 if(cl.cut)XS.clip=null;await recalc(sh,true);xrender()}
function struct(axis,at,n,del){pushU();const sh=curSh(),nc={},R=axis==='r';
 const mv=i=>del?(i>=at+n?i-n:i>=at?-1:i):(i>=at?i+n:i);
 for(const k in sh.cells){const [r,c]=k.split(',').map(Number),i=mv(R?r:c);if(i<0)continue;const o=sh.cells[k];
  if(o.f!==undefined)o.f=mapF(o.f,(a,ci,b,ri)=>{const x=R?ri:ci,y=mv(x);if(y<0)return '#REF!';ri=R?y:ri;ci=R?ci:y;return a+colName(ci)+b+(ri+1)});
  nc[R?K(i,c):K(r,i)]=o}
 sh.cells=nc;
 sh.merges=sh.merges.map(m=>{const a=mv(R?m[0]:m[1]),b=mv(R?m[2]:m[3]);if(a<0||b<0)return null;return R?[a,m[1],b,m[3]]:[m[0],a,m[2],b]}).filter(Boolean);
 if(!R){const cw={};for(const c in sh.cw){const i=mv(+c);if(i>=0)cw[i]=sh.cw[c]}sh.cw=cw;sh.NC=Math.max(20,sh.NC+(del?-n:n))}else sh.NR=Math.max(100,sh.NR+(del?-n:n));
 recalc(sh,true).then(()=>{xrender();loadFB()})}
function colw(d){pushU();const sh=curSh(),q=rng();for(let c=q.c1;c<=q.c2;c++)sh.cw[c]=Math.max(30,(sh.cw[c]||DCW)+d);xrender()}
function sortCol(asc){pushU();const sh=curSh(),q=rng(),rows=[];let last=-1;for(const k in sh.cells){const r=+k.split(',')[0];if(r>last)last=r}
 const c0=q.c1,r1=q.r1===q.r2?q.r1:q.r1,r2=q.r1===q.r2?last:q.r2;
 for(let r=r1;r<=r2;r++){const row={};for(let c=0;c<sh.NC;c++){const o=sh.cells[K(r,c)];if(o)row[c]=o}rows.push({row,key:(sh.cells[K(r,c0)]||{}).v})}
 rows.sort((a,b)=>{const x=a.key,y=b.key;if(x===undefined)return 1;if(y===undefined)return -1;const d=typeof x==='number'&&typeof y==='number'?x-y:String(x).localeCompare(String(y),'ar');return asc?d:-d});
 for(let r=r1;r<=r2;r++)for(let c=0;c<sh.NC;c++)delete sh.cells[K(r,c)];
 rows.forEach((x,i)=>{for(const c in x.row)sh.cells[K(r1+i,+c)]=x.row[c]});recalc(sh,true).then(xrender)}
async function xformula(){const q=rng(),range=colName(q.c1)+(q.r1+1)+':'+colName(q.c2)+(q.r2+1);const opts=['SUM','AVERAGE','MIN','MAX','COUNT','COUNTA','COUNTBLANK','PRODUCT','MEDIAN','MODE','STDEV','VAR','ABS','ROUND','ROUNDUP','ROUNDDOWN','INT','MOD','SQRT','POWER','IF','IFERROR','AND','OR','NOT','SUMIF','COUNTIF','AVERAGEIF','CONCATENATE','LEFT','RIGHT','MID','LEN','UPPER','LOWER','TRIM'];const f=await raAsk('اكتب اسم الدالة:\n'+opts.join(' · '),'SUM');if(!f)return;const fn=f.trim().toUpperCase();if(!opts.includes(fn))return toast('اسم الدالة غير موجود بالقائمة');let args=range;if(!['SUM','AVERAGE','MIN','MAX','COUNT','COUNTA','COUNTBLANK','PRODUCT','MEDIAN','MODE','STDEV','VAR'].includes(fn)){const ex={IF:'A1>0,\"نعم\",\"لا\"',IFERROR:'A1/B1,0',ROUND:'A1,2',ROUNDUP:'A1,0',ROUNDDOWN:'A1,0',MOD:'A1,2',POWER:'A1,2',LEFT:'A1,3',RIGHT:'A1,3',MID:'A1,2,3',SUMIF:'A1:A10,\">0\",B1:B10',COUNTIF:'A1:A10,\">0\"',AVERAGEIF:'A1:A10,\">0\"',CONCATENATE:'A1,B1',AND:'A1>0,B1>0',OR:'A1>0,B1>0',NOT:'A1>0'};const a=await raAsk('اكتب معاملات '+fn+' بدون علامة ='+fn+' (مثال: '+(ex[fn]||'A1')+')',ex[fn]||'A1');if(a===null)return;args=a.trim()}commit('='+fn+'('+args+')',q.r1,q.c1)}
function xclear(){const q=rng(),sh=curSh();pushU();for(let r=q.r1;r<=q.r2;r++)for(let c=q.c1;c<=q.c2;c++){const k=K(r,c),o=sh.cells[k];if(o){delete o.v;delete o.f;delete o.cv;if(!o.s||!Object.keys(o.s).length)delete sh.cells[k]}}recalc(sh,true).then(()=>{xrender();loadFB()})}
async function xfindreplace(){const find=await raAsk('ابحث عن:', '');if(!find)return;const rep=await raAsk('استبدال بـ (اتركه فارغًا للحذف):','');const sh=curSh();let n=0;pushU();for(const k in sh.cells){const o=sh.cells[k];if(typeof o.v==='string'&&o.v.includes(find)){o.v=o.v.split(find).join(rep);n++}}xrender();loadFB();toast('تم استبدال '+n+' خلية')}
function xinsertDate(){const q=rng();commit(String(d2s(new Date())),q.r1,q.c1).then(()=>setS('nf','dd/mm/yyyy'))}
function xduplicateSheet(){const sh=curSh(),copy=JSON.parse(JSON.stringify(sh));let name=(sh.name+' Copy').slice(0,31),base=name,i=2;while(XS.X.sheets.some(x=>x.name===name)){name=(base.slice(0,27)+' '+i++).slice(0,31)}pushU();copy.name=name;XS.X.sheets.splice(XS.cur+1,0,copy);XS.cur++;xtabs();xrender();loadFB()}
function xtabs(){const t=$('#xtabs');t.innerHTML='';XS.X.sheets.forEach((sh,i)=>{const b=E('button',i===XS.cur?'on':'');b.textContent=sh.name;b.onclick=()=>{XS.cur=i;XS.sel={r1:0,c1:0,r2:0,c2:0};$('#xg').scrollTop=0;xtabs();xrender();loadFB()};t.append(b)});
 const add=E('button','',' ＋ ');add.onclick=()=>{pushU();let n=XS.X.sheets.length+1;while(XS.X.sheets.some(s=>s.name==='Sheet'+n))n++;XS.X.sheets.push(newSheet('Sheet'+n));XS.cur=XS.X.sheets.length-1;xtabs();xrender();loadFB()};
 const ren=E('button','','✎');ren.onclick=async()=>{const n=await raAsk('اسم الشيت:',curSh().name);if(n&&n.trim()&&!XS.X.sheets.some((s,i)=>i!==XS.cur&&s.name===n.trim())){pushU();curSh().name=n.trim().slice(0,31);xtabs()}};
 const del=E('button','','🗑');del.onclick=async()=>{if(XS.X.sheets.length<2)return toast('لازم يفضل شيت واحد على الأقل');if(await raConfirm('حذف الشيت "'+curSh().name+'"؟')){pushU();XS.X.sheets.splice(XS.cur,1);XS.cur=Math.max(0,XS.cur-1);xtabs();xrender();loadFB()}};
 t.append(add,ren,del)}
function xbar(){
 const nfs={'عام':'General','رقم 0.00':'0.00','1,234':'#,##0','1,234.00':'#,##0.00','نسبة %':'0.00%','تاريخ':'yyyy-mm-dd','dd/mm/yyyy':'dd/mm/yyyy','نص':'@','عملة $':'"$"#,##0.00'};
 const nf=E('select');nf.innerHTML='<option value="">التنسيق</option>'+Object.keys(nfs).map(k=>'<option>'+k+'</option>').join('');nf.onchange=()=>{if(nf.value){const v=nfs[nf.value];setS('nf',v==='General'?null:v);nf.value=''}};
 const xcol=(ph,key)=>colorBtn({label:ph,title:key==='bg'?'لون الخلفية':'لون الخط',init:key==='bg'?'#FFFF00':'#FF0000',onPick:c=>setS(key,c)});
 setBar([{t:'💾',title:'حفظ',f:xsave},{t:'⬇ XLSX',f:()=>xsave(1)},'|',
  {t:'↶',f:undo,title:'تراجع'},{t:'↷',f:redo,title:'إعادة'},'|',
  {t:'ƒx دوال',f:xformula},{t:'بحث/استبدال',f:xfindreplace},{t:'تاريخ اليوم',f:xinsertDate},{t:'مسح المحتوى',f:xclear},'|',
  {t:'<b>B</b>',f:()=>tog('b')},{t:'<i>I</i>',f:()=>tog('i')},{t:'<u>U</u>',f:()=>tog('u')},{t:'S̶',f:()=>tog('st'),title:'يتوسطه خط'},xcol('A لون','c'),xcol('🎨 خلفية','bg'),
  {t:'⬅',f:()=>setS('al','left')},{t:'↔',f:()=>setS('al','center')},{t:'➡',f:()=>setS('al','right')},{t:'⏎ التفاف',f:()=>tog('w')},{t:'▦ حدود',f:()=>tog('bd')},{el:nf},'|',
  {t:'⬚ نطاق',f:b=>{XS.ext=!XS.ext;b.classList.toggle('on',XS.ext)}},{t:'⊞ دمج',f:()=>merge()},{t:'⊟ فك',f:()=>merge(1)},'|',
  {t:'نسخ',f:()=>copy()},{t:'قص',f:()=>copy(1)},{t:'لصق',f:paste},{t:'⬇ تعبئة',f:()=>fill('d')},{t:'➡ تعبئة',f:()=>fill('r')},'|',
  {t:'نسخ الشيت',f:xduplicateSheet},{t:'＋صف',f:()=>struct('r',rng().r1,1,0)},{t:'＋عمود',f:()=>struct('c',rng().c1,1,0)},{t:'－صف',f:()=>{const q=rng();struct('r',q.r1,q.r2-q.r1+1,1)}},{t:'－عمود',f:()=>{const q=rng();struct('c',q.c1,q.c2-q.c1+1,1)}},'|',
  {t:'↔＋',f:()=>colw(16)},{t:'↔－',f:()=>colw(-16)},{t:'A→Z',f:()=>sortCol(1)},{t:'Z→A',f:()=>sortCol(0)},{t:'🖨 طباعة',f:printOffice},{t:'🔐 نسخة مشفرة',f:encryptedExport}])}
async function xsave(dlOnly){try{busy(true,'جاري الحفظ…');if(document.activeElement===$('#xin'))$('#xin').blur();const buf=await xlsBuild(XS.X);await saveBuf(buf);busy(false);if(dlOnly===1)dl(new Blob([buf],{type:MIME.xls}),O.d.n+'.xlsx');else toast('اتحفظ ✅')}catch(e){fail('فشل الحفظ',e)}}
async function xpdf(){try{busy(true,'جاري تحويل PDF…');const sh=curSh();let R=0,C=0;for(const k in sh.cells){const [r,c]=k.split(',').map(Number);R=Math.max(R,r);C=Math.max(C,c)}
 const cov={},mg={};for(const m of sh.merges){mg[K(m[0],m[1])]=m;for(let r=m[0];r<=m[2];r++)for(let c=m[1];c<=m[3];c++)if(r!==m[0]||c!==m[1])cov[K(r,c)]=1}
 let h='<table class="xp">';for(let r=0;r<=R;r++){h+='<tr>';for(let c=0;c<=C;c++){const k=K(r,c);if(cov[k])continue;const o=sh.cells[k],m=mg[k];h+='<td dir="auto" style="min-width:'+(sh.cw[c]||DCW)+'px;'+cstyle(o,1)+'"'+(m?' colspan="'+(m[3]-m[1]+1)+'" rowspan="'+(m[2]-m[0]+1)+'"':'')+'>'+esc(disp(o))+'</td>'}h+='</tr>'}
 const box=E('div','ofprint x');box.innerHTML=h+'</table>';document.body.append(box);const b=await longToPdf(box,.7071);box.remove();busy(false);dl(b,O.d.n+'.pdf')}catch(e){fail('فشل التصدير',e)}}

// ====================== POWERPOINT (view) ======================
async function zipText(buf,name){const u=new Uint8Array(buf),dv=new DataView(u.buffer,u.byteOffset,u.byteLength);let e=u.length-22;while(e>=0&&dv.getUint32(e,true)!==0x06054b50)e--;if(e<0)throw new Error('zip');
 const n=dv.getUint16(e+10,true);let p=dv.getUint32(e+16,true);
 for(let i=0;i<n;i++){const cm=dv.getUint16(p+10,true),cs=dv.getUint32(p+20,true),nl=dv.getUint16(p+28,true),el=dv.getUint16(p+30,true),cl=dv.getUint16(p+32,true),lo=dv.getUint32(p+42,true),nm=new TextDecoder().decode(u.subarray(p+46,p+46+nl));
  if(nm===name){const s=lo+30+dv.getUint16(lo+26,true)+dv.getUint16(lo+28,true),d=u.subarray(s,s+cs);if(cm===0)return new TextDecoder().decode(d);
   return await new Response(new Blob([d]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).text()}
  p+=46+nl+el+cl}throw new Error('entry')}
// ---------- PowerPoint: play embedded audio / video ----------
const NSR='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const AUD='mp3 wav m4a wma aac ogg oga flac aif aiff mid midi amr opus'.split(' '),VID='mp4 m4v mov wmv avi mpg mpeg webm mkv 3gp asf flv ogv'.split(' ');
const MM={mp3:'audio/mpeg',wav:'audio/wav',m4a:'audio/mp4',aac:'audio/aac',ogg:'audio/ogg',oga:'audio/ogg',opus:'audio/ogg',flac:'audio/flac',aif:'audio/aiff',aiff:'audio/aiff',mid:'audio/midi',midi:'audio/midi',wma:'audio/x-ms-wma',amr:'audio/amr',
 mp4:'video/mp4',m4v:'video/mp4',mov:'video/quicktime',webm:'video/webm',wmv:'video/x-ms-wmv',avi:'video/x-msvideo',mpg:'video/mpeg',mpeg:'video/mpeg',mkv:'video/x-matroska','3gp':'video/3gpp',asf:'video/x-ms-asf',flv:'video/x-flv',ogv:'video/ogg'};
const extOf=n=>(/\.([a-z0-9]+)$/i.exec(n)||[])[1]?.toLowerCase()||'';
function zipIndex(buf){const u=new Uint8Array(buf),dv=new DataView(u.buffer,u.byteOffset,u.byteLength);let e=u.length-22;while(e>=0&&dv.getUint32(e,true)!==0x06054b50)e--;if(e<0)throw new Error('zip');
 const n=dv.getUint16(e+10,true),td=new TextDecoder(),m={};let p=dv.getUint32(e+16,true);
 for(let i=0;i<n&&dv.getUint32(p,true)===0x02014b50;i++){const cm=dv.getUint16(p+10,true),cs=dv.getUint32(p+20,true),us=dv.getUint32(p+24,true),nl=dv.getUint16(p+28,true),el=dv.getUint16(p+30,true),cl=dv.getUint16(p+32,true),lo=dv.getUint32(p+42,true);
  m[td.decode(u.subarray(p+46,p+46+nl))]={cm,cs,us,lo,fl:dv.getUint16(p+8,true),tm:dv.getUint16(p+12,true),dt:dv.getUint16(p+14,true),crc:dv.getUint32(p+16,true)};p+=46+nl+el+cl}
 return {u,dv,m}}
async function zipRead(z,name){const f=z.m[name];if(!f)return null;const s=f.lo+30+z.dv.getUint16(f.lo+26,true)+z.dv.getUint16(f.lo+28,true),d=z.u.subarray(s,s+f.cs);
 if(f.cm===0)return d;return new Uint8Array(await new Response(new Blob([d]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer())}
const kids=(el,n)=>[...el.children].filter(c=>c.localName===n),first=(el,n)=>el&&kids(el,n)[0];
const num=(el,a,d)=>{const v=el&&el.getAttribute(a);return v==null||v===''?d:+v};

// finds every <p:pic> that is really an audio / video object (also inside groups)
function walkPics(tree,tf,out){for(const c of tree.children){const n=c.localName;
 if(n==='grpSp'){const xf=first(first(c,'grpSpPr'),'xfrm'),off=first(xf,'off'),ext=first(xf,'ext'),co=first(xf,'chOff'),ce=first(xf,'chExt');let t=tf;
  if(off&&ext&&co&&ce){const sx=num(ext,'cx',1)/(num(ce,'cx',1)||1),sy=num(ext,'cy',1)/(num(ce,'cy',1)||1);
   t={sx:tf.sx*sx,sy:tf.sy*sy,tx:tf.tx+tf.sx*(num(off,'x',0)-num(co,'x',0)*sx),ty:tf.ty+tf.sy*(num(off,'y',0)-num(co,'y',0)*sy)}}
  walkPics(c,t,out)}
 else if(n==='AlternateContent'){const ch=first(c,'Choice');ch&&walkPics(ch,tf,out)}
 else if(n==='pic'){const nv=first(c,'nvPicPr'),cn=first(nv,'cNvPr'),pv=first(nv,'nvPr');if(!pv)continue;
  let kind=null;const r=[];
  const af=first(pv,'audioFile'),vf=first(pv,'videoFile'),wf=first(pv,'wavAudioFile'),qf=first(pv,'quickTimeFile');
  if(af){kind='audio';r.push(af.getAttributeNS(NSR,'link'))}
  if(wf){kind='audio';r.push(wf.getAttributeNS(NSR,'embed'))}
  if(vf){kind='video';r.push(vf.getAttributeNS(NSR,'link'))}
  if(qf){kind='video';r.push(qf.getAttributeNS(NSR,'link'))}
  const md=pv.getElementsByTagNameNS('*','media')[0];if(md){r.unshift(md.getAttributeNS(NSR,'embed')||md.getAttributeNS(NSR,'link'));kind=kind||'video'}
  if(!kind)continue;
  const xf=first(first(c,'spPr'),'xfrm'),off=first(xf,'off'),ext=first(xf,'ext');let rect=null;
  if(off&&ext)rect={x:tf.tx+tf.sx*num(off,'x',0),y:tf.ty+tf.sy*num(off,'y',0),w:tf.sx*num(ext,'cx',0),h:tf.sy*num(ext,'cy',0)};
  out.push({kind,rids:r.filter(Boolean),rect,name:(cn&&(cn.getAttribute('descr')||cn.getAttribute('name')))||'',hid:cn&&cn.getAttribute('hidden')==='1'})}}}

async function pptMediaScan(buf){
 const z=zipIndex(buf),td=new TextDecoder();
 const xml=async n=>{const d=await zipRead(z,n);return d?new DOMParser().parseFromString(td.decode(d),'application/xml'):null};
 const norm=(base,t)=>{if(t[0]==='/')return t.slice(1);const p=base.split('/');p.pop();for(const s of t.split('/')){if(s==='..')p.pop();else if(s&&s!=='.')p.push(s)}return p.join('/')};
 const rels=async path=>{const d=await xml(path.replace(/([^/]+)$/,'_rels/$1.rels')),m={};if(d)for(const r of d.getElementsByTagNameNS('*','Relationship'))m[r.getAttribute('Id')]={t:r.getAttribute('Target')||'',x:r.getAttribute('TargetMode')==='External'};return m};
 const pres=await xml('ppt/presentation.xml'),pr=await rels('ppt/presentation.xml');
 const paths=[];for(const s of pres.getElementsByTagNameNS('*','sldId')){const r=pr[s.getAttributeNS(NSR,'id')];if(r&&!r.x)paths.push(norm('ppt/presentation.xml',r.t))}
 const sz=pres.getElementsByTagNameNS('*','sldSz')[0],SW=num(sz,'cx',12192000),SH=num(sz,'cy',6858000);
 const items=[],used=new Set(),hidden=[];
 const mk=(o)=>{const f=o.path?o.path.split('/').pop():(o.target||'').split(/[\\/]/).pop(),ext=extOf(f);
  return Object.assign({file:f,ext,mime:MM[ext]||(o.kind==='video'?'video/mp4':'audio/mpeg'),size:o.path&&z.m[o.path]?z.m[o.path].us:0},o)};
 for(let i=0;i<paths.length;i++){const sp=paths[i],doc=await xml(sp);if(!doc)continue;const rl=await rels(sp);
  hidden.push(doc.documentElement.getAttribute('show')==='0');
  const tree=doc.getElementsByTagNameNS('*','spTree')[0],pics=[];
  if(tree)walkPics(tree,{sx:1,sy:1,tx:0,ty:0},pics);
  for(const p of pics){const cand=p.rids.map(id=>rl[id]).filter(Boolean),ok=cand.find(c=>!c.x)||cand[0];if(!ok)continue;
   const path=ok.x?null:norm(sp,ok.t);if(path)used.add(path);
   items.push(mk({slide:i+1,kind:p.kind,path,target:ok.t,external:ok.x||(path&&!z.m[path]),rect:p.rect,hid:p.hid,title:p.name}))}
  // transition / click sounds
  const seen=new Set(items.filter(m=>m.slide===i+1).map(m=>m.path));
  for(const s of doc.getElementsByTagNameNS('*','snd')){const r=rl[s.getAttributeNS(NSR,'embed')];if(!r||r.x)continue;const path=norm(sp,r.t);if(seen.has(path))continue;seen.add(path);used.add(path);
   items.push(mk({slide:i+1,kind:'audio',path,rect:null,title:s.getAttribute('name')||'صوت انتقال / نقر'}))}}
 // anything else stored in ppt/media (layouts, masters, unused)
 for(const nm of Object.keys(z.m)){if(!/^ppt\/media\//.test(nm)||used.has(nm))continue;const e=extOf(nm);
  if(AUD.includes(e)||VID.includes(e))items.push(mk({slide:0,kind:AUD.includes(e)?'audio':'video',path:nm,rect:null,title:''}))}
 items.forEach(m=>{if(!m.title||/^(Audio|Video|Media|Picture|Content Placeholder)\b/i.test(m.title))m.title=m.file||'ملف وسائط'});
 return {z,items,SW,SH,total:paths.length,hidden}}

// ----- convert unsupported audio (wma, aif, amr…) to mp3 -----
const CT_=(()=>{const t=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[n]=c>>>0}return t})();
const crc32=d=>{let c=0xFFFFFFFF;for(let i=0;i<d.length;i++)c=CT_[(c^d[i])&255]^(c>>>8);return (c^0xFFFFFFFF)>>>0};
const AP_=document.createElement('audio');
const needConv=m=>m.kind==='audio'&&!m.external&&!!m.path&&m.ext!=='mp3'&&!/^midi?$/.test(m.ext)&&!AP_.canPlayType(m.mime);
const FFB='https://cdn.jsdelivr.net/npm/@ffmpeg/';
async function getFF(){if(MP.ff)return MP.ff;
 await need('ff');
 const urls=[],errs=[];
 const blob=async(local,cdn,type)=>{let r;try{r=await fetch(local)}catch(e){}if(!r||!r.ok)r=await fetch(cdn);if(!r.ok)throw new Error('تحميل '+local+' فشل ('+r.status+')');const u=URL.createObjectURL(new Blob([await r.arrayBuffer()],{type}));urls.push(u);return u};
 const core=FFB+'core@0.12.6/dist/',worker=FFB+'ffmpeg@0.12.10/dist/umd/';
 const withTimeout=(p,ms)=>Promise.race([p,new Promise((_,no)=>setTimeout(()=>no(new Error('انتهت مهلة تحميل المحوّل')),ms))]);
 // Attempt 1: default classic worker (ffmpeg.js and 814.ffmpeg.js sit side by side in lib/ffmpeg/) + UMD core.
 // Attempt 2: module worker through classWorkerURL — this REQUIRES the ESM core (a UMD core cannot be import()ed from a module worker).
 const attempts=[
  ['classic',async()=>({coreURL:await blob('lib/ffmpeg/ffmpeg-core.js',core+'umd/ffmpeg-core.js','text/javascript'),wasmURL:await blob('lib/ffmpeg/ffmpeg-core.wasm',core+'umd/ffmpeg-core.wasm','application/wasm')})],
  ['module',async()=>({classWorkerURL:await blob('lib/ffmpeg/814.ffmpeg.js',worker+'814.ffmpeg.js','text/javascript'),coreURL:await blob('lib/ffmpeg/esm/ffmpeg-core.js',core+'esm/ffmpeg-core.js','text/javascript'),wasmURL:await blob('lib/ffmpeg/esm/ffmpeg-core.wasm',core+'esm/ffmpeg-core.wasm','application/wasm')})]
 ];
 for(const [name,mk] of attempts){
  const ff=new window.FFmpegWASM.FFmpeg();
  try{await withTimeout(ff.load(await mk()),120000);MP.ff=ff;break}
  catch(e){errs.push(name+': '+(e&&e.message||e));console.warn('ffmpeg load failed ('+name+')',e);try{ff.terminate()}catch(_){}}
 }
 urls.forEach(u=>{try{URL.revokeObjectURL(u)}catch(_){}});
 if(!MP.ff)throw new Error('تعذر تشغيل محوّل الصوت. ['+errs.join(' | ')+']');
 return MP.ff}
async function toMp3(bytes,ext){const ff=await getFF(),i='in.'+(ext||'bin'),o='out.mp3';
 await ff.writeFile(i,new Uint8Array(bytes)); // copy: writeFile transfers (detaches) its buffer
 try{const rc=await ff.exec(['-y','-i',i,'-vn','-map_metadata','-1','-c:a','libmp3lame','-q:a','4',o]);if(rc!==0)throw new Error('ffmpeg رجّع '+rc);return new Uint8Array(await ff.readFile(o))}
 finally{try{await ff.deleteFile(i)}catch(e){}try{await ff.deleteFile(o)}catch(e){}}}
// WMA playback: decode to an in-memory PCM/WAV stream only when the device cannot
// decode WMA natively. The original PPTX media and downloaded WMA are never changed.
async function decodeWmaForPlayback(bytes){const ff=await getFF(),i='playback.wma',o='playback.wav';
 await ff.writeFile(i,new Uint8Array(bytes));
 try{const rc=await ff.exec(['-y','-i',i,'-vn','-map_metadata','-1','-acodec','pcm_s16le',o]);if(rc!==0)throw new Error('WMA decoder returned '+rc);return new Uint8Array(await ff.readFile(o))}
 finally{try{await ff.deleteFile(i)}catch(e){}try{await ff.deleteFile(o)}catch(e){}}}

// writes a new zip: copies untouched entries byte-for-byte, drops `del`, adds/replaces `put`
function zipBuild(z,del,put){const parts=[],cd=[],te=new TextEncoder();let off=0;
 const add=(nm,tm,dt,fl,cm,crc,cs,us,raw)=>{const nb=te.encode(nm),h=new DataView(new ArrayBuffer(30)),c=new DataView(new ArrayBuffer(46));
  h.setUint32(0,0x04034b50,true);h.setUint16(4,20,true);h.setUint16(6,fl,true);h.setUint16(8,cm,true);h.setUint16(10,tm,true);h.setUint16(12,dt,true);h.setUint32(14,crc,true);h.setUint32(18,cs,true);h.setUint32(22,us,true);h.setUint16(26,nb.length,true);
  c.setUint32(0,0x02014b50,true);c.setUint16(4,20,true);c.setUint16(6,20,true);c.setUint16(8,fl,true);c.setUint16(10,cm,true);c.setUint16(12,tm,true);c.setUint16(14,dt,true);c.setUint32(16,crc,true);c.setUint32(20,cs,true);c.setUint32(24,us,true);c.setUint16(28,nb.length,true);c.setUint32(42,off,true);
  parts.push(new Uint8Array(h.buffer),nb,raw);cd.push(new Uint8Array(c.buffer),nb);off+=30+nb.length+raw.length};
 for(const [nm,f] of Object.entries(z.m)){if(del.has(nm))continue;
  if(put.has(nm)){const d=put.get(nm);add(nm,f.tm,f.dt,0x0800,0,crc32(d),d.length,d.length,d);continue}
  if(f.cs===0xFFFFFFFF||f.lo===0xFFFFFFFF)throw new Error('ملف zip64 مش مدعوم');
  const s=f.lo+30+z.dv.getUint16(f.lo+26,true)+z.dv.getUint16(f.lo+28,true);
  add(nm,f.tm,f.dt,f.fl&0x0800,f.cm,f.crc,f.cs,f.us,z.u.subarray(s,s+f.cs))}
 for(const [nm,d] of put)if(!(nm in z.m))add(nm,0,0x21,0x0800,0,crc32(d),d.length,d.length,d);
 let n=0;cd.forEach(x=>n+=x.length);const e=new DataView(new ArrayBuffer(22)),k=cd.length/2;
 e.setUint32(0,0x06054b50,true);e.setUint16(8,k,true);e.setUint16(10,k,true);e.setUint32(12,n,true);e.setUint32(16,off,true);
 return new Blob([...parts,...cd,new Uint8Array(e.buffer)],{type:MIME.ppt})}

// ----- player (one shared overlay: list / audio dock / video fullscreen) -----
const MP={z:null,items:[],url:null,el:null,cur:null,tk:0,cv:new Map(),ff:null,autoPromise:null,doc:null};
// Convert every embedded audio track to MP3 silently in the background as soon as a presentation is opened.
async function autoConvertPptAudio(){
 const z=MP.z,doc=MP.doc,todo=[...new Map(MP.items.filter(m=>m.kind==='audio'&&!m.external&&m.path&&m.ext!=='mp3'&&!/^midi?$/.test(m.ext)).map(m=>[m.path,m])).values()];
 if(!z||!todo.length)return;
 toast('جاري تجهيز '+todo.length+' ملف صوت في الخلفية…');
 MP.autoPromise=(async()=>{let done=0,failed=0,cached=0;
  for(const m of todo){if(MP.z!==z)break;if(MP.cv.has(m.path))continue;
   try{
    const src=await zipRead(z,m.path);if(!src)throw new Error('ملف الصوت غير موجود');
    // Cache by document ID + source checksum/length so reopening a presentation never reconverts unchanged audio.
    const cacheKey=doc&&doc.id?('ra_mp3_'+String(doc.id).replace(/[^a-z0-9_.-]/gi,'_')+'_'+crc32(src).toString(16)+'_'+src.length):null;
    let mp3=cacheKey?await getBuf(cacheKey):null;
    if(mp3){cached++}else{
     mp3=await toMp3(src,m.ext);if(MP.z!==z)break;
     if(cacheKey&&doc){await putBuf(cacheKey,mp3);doc.au=Array.isArray(doc.au)?doc.au:[];if(!doc.au.some(a=>a&&a.k===cacheKey))doc.au.push({k:cacheKey});doc.u=Date.now();await dbPut(doc)}
    }
    if(MP.z!==z)break;MP.cv.set(m.path,mp3);done++;
   }catch(e){failed++;console.warn('Background audio conversion failed:',m.file,e)}
  }
  if(MP.z===z&&(done||cached))toast('الصوت جاهز للتشغيل: '+done+' جديد، '+cached+' من الذاكرة'+(failed?'، تعذر تحويل '+failed:''));
 })();
 return MP.autoPromise;
}
const fmtSize=b=>b>1048576?(b/1048576).toFixed(1)+' MB':Math.max(1,Math.round(b/1024))+' KB';
function mpUI(){let r=$('#mpl');if(r)return r;r=E('div','','<div class="mph"><b class="mpt"></b><button class="mpg" hidden aria-label="القائمة">☰</button><button class="mps" hidden aria-label="حفظ الملف">⬇</button><button class="mx" aria-label="إغلاق">✕</button></div><div class="mpbd"></div>');
 r.id='mpl';r.hidden=true;document.body.append(r);
 r.querySelector('.mx').onclick=mpClose;r.querySelector('.mpg').onclick=mpList;
 r.querySelector('.mps').onclick=async()=>{const m=MP.cur;if(!m||!MP.z)return;try{const c=MP.cv.get(m.path),d=c||await zipRead(MP.z,m.path);dl(new Blob([d],{type:c?'audio/mpeg':m.mime}),c?m.file.replace(/\.[^.]+$/,'')+'.mp3':m.file)}catch(e){fail('تعذر الحفظ',e)}};
 return r}
function mpStop(){MP.tk++;if(MP.el){try{MP.el.pause();MP.el.removeAttribute('src');MP.el.load()}catch(e){}MP.el=null}if(MP.url){URL.revokeObjectURL(MP.url);MP.url=null}const r=$('#mpl');if(r)r.querySelector('.mpbd').innerHTML=''}
function mpClose(){mpStop();MP.cur=null;const r=$('#mpl');if(r){r.hidden=true;r.className=''}}
function mpList(){const r=mpUI();mpStop();MP.cur=null;r.className='ls';r.hidden=false;
 r.querySelector('.mpt').textContent='🎬 الوسائط ('+MP.items.length+')';r.querySelector('.mpg').hidden=true;r.querySelector('.mps').hidden=true;
 const b=r.querySelector('.mpbd');b.innerHTML='';
 for(const m of MP.items){const x=E('button','mpr','<i>'+(m.kind==='video'?'🎬':'🔊')+'</i><span><b></b><small></small></span><em>'+(m.external?'⚠':'▶')+'</em>');
  x.querySelector('b').textContent=m.title;
  const v=document.createElement(m.kind),weak=!m.external&&!v.canPlayType(m.mime);
  x.querySelector('small').textContent=(m.slide?'شريحة '+m.slide:'مش مربوط بشريحة')+' · '+(m.file||'')+(m.size?' · '+fmtSize(m.size):'')+(m.external?' · ملف خارجي مش جوه العرض':m.ext==='wma'?' · ▶ تشغيل WMA':weak?(needConv(m)?' · 🔄 قابل للتحويل لـ mp3':' · صيغة ممكن متشتغلش'):'');
  x.onclick=()=>mpPlay(m);b.append(x)}}
async function mpPlay(m){
 if(m.external||!m.path)return toast('الملف ده مربوط من برّه العرض ومش موجود جواه'+(m.file?': '+m.file:''));
 const r=mpUI();mpStop();const tk=MP.tk;
 try{
  let d,mime=m.mime,original=null;
  if(m.data){d=m.data}
  else if(MP.cv.has(m.path)){d=MP.cv.get(m.path);mime='audio/mpeg'}
  else{
   if(m.size>8e6)busy(true,'جاري تجهيز الملف…');
   d=await zipRead(MP.z,m.path);busy(false);if(tk!==MP.tk)return;
   // Wait for the automatic background conversion; no manual conversion prompt.
   if(m.kind==='audio'&&m.ext!=='mp3'&&MP.autoPromise){await MP.autoPromise;if(tk!==MP.tk)return;const converted=MP.cv.get(m.path);if(converted){d=converted;mime='audio/mpeg'}else if(m.ext==='wma')original=d}
   else if(m.ext==='wma')original=d;
   if(tk!==MP.tk)return
  }
  const makeUrl=bytes=>URL.createObjectURL(new Blob([bytes],{type:mime}));
  MP.url=makeUrl(d);MP.cur=m;
  r.className=m.kind==='video'?'vd':'au';r.hidden=false;
  r.querySelector('.mpt').textContent=m.title;r.querySelector('.mpg').hidden=MP.items.length<2;r.querySelector('.mps').hidden=false;
  const el=document.createElement(m.kind);el.controls=true;el.playsInline=true;el.preload='auto';el.setAttribute('dir','ltr');
  let fallbackStarted=false;
  el.onerror=async()=>{
   if(MP.el!==el)return;
   if(m.ext==='wma'&&original&&!fallbackStarted){
    fallbackStarted=true;
    try{
     busy(true,'الجهاز لا يدعم WMA مباشرة… جاري تجهيزه للتشغيل');
     const wav=await decodeWmaForPlayback(original);
     if(tk!==MP.tk||MP.el!==el)return;
     const old=MP.url;MP.url=URL.createObjectURL(new Blob([wav],{type:'audio/wav'}));
     el.src=MP.url;el.load();if(old)URL.revokeObjectURL(old);
     const play=el.play();if(play&&play.catch)play.catch(()=>{});
    }catch(e){
     console.warn('WMA playback fallback failed',e);
     if(tk===MP.tk&&MP.el===el){const p=E('p','mpe');p.textContent='تعذر تشغيل ملف WMA تلقائيًا. اضغط ⬇ احفظ الملف وشغّله ببرنامج تاني. ('+String(e&&e.message||e).slice(0,300)+')';r.querySelector('.mpbd').append(p)}
    }finally{busy(false)}
    return
   }
   const p=E('p','mpe');p.textContent='الجهاز مش قادر يشغّل صيغة ‎.'+m.ext+'‎ دي. اضغط ⬇ احفظ الملف وشغّله ببرنامج تاني.';r.querySelector('.mpbd').append(p)
  };
  el.src=MP.url;MP.el=el;const b=r.querySelector('.mpbd');b.innerHTML='';b.append(el);const pl=el.play();pl&&pl.catch&&pl.catch(()=>{})
 }catch(e){busy(false);fail('تعذر تشغيل الملف',e)}}

// builds the on-slide play buttons and returns the toolbar items for the media list
async function pptMediaSetup(w,slideEls){
 MP.z=null;MP.items=[];MP.cv.clear();MP.autoPromise=null;MP.doc=O&&O.d||null;let sc;
 try{sc=await pptMediaScan(O.buf)}catch(e){console.warn('media scan',e);return[]}
 MP.z=sc.z;MP.items=sc.items;if(!sc.items.length)return[];
 // Start conversion immediately without blocking slide rendering.
 autoConvertPptAudio().catch(e=>console.warn('Background audio conversion:',e));
 const els=slideEls();let map=null; // slide number -> rendered element
 if(els.length===sc.total)map=i=>els[i-1];
 else{const vis=[];sc.hidden.forEach((h,i)=>{if(!h)vis.push(i+1)});
  if(els.length===vis.length)map=i=>els[vis.indexOf(i)]}
 if(map)for(const m of sc.items){if(!m.slide||!m.rect||m.hid||m.external)continue;const el=map(m.slide);if(!el)continue;
  const r=m.rect,cx=(r.x+r.w/2)/sc.SW*100,cy=(r.y+r.h/2)/sc.SH*100;if(cx<0||cx>100||cy<0||cy>100)continue;
  let L=[...el.children].find(c=>c.classList.contains('mpl-ov'));
  if(!L){L=E('div','mpl-ov');if(getComputedStyle(el).position==='static')el.style.position='relative';el.append(L)}
  const b=E('button','mpo','<i>'+(m.kind==='video'?'▶':'🔊')+'</i>');
  b.style.cssText='left:'+cx+'%;top:'+cy+'%;width:max('+r.w/sc.SW*100+'%,40px);height:max('+r.h/sc.SH*100+'%,40px)';
  b.setAttribute('aria-label','تشغيل: '+m.title);b.onclick=e=>{e.stopPropagation();mpPlay(m)};L.append(b)}
 O.cleanup=()=>{mpClose();MP.z=null;MP.items=[];MP.cv.clear();MP.autoPromise=null;MP.doc=null};
 return['|',{t:'🎬 الوسائط ('+sc.items.length+')',f:mpList}]}

async function pptView(){const pv=await need('pp');$('#ofm').className='ofscroll';$('#ofm').innerHTML='<div class="ofw ppt"></div>';const w=$('#ofm').firstChild;
 let ar=16/9;try{const x=await zipText(O.buf,'ppt/presentation.xml'),m=/<p:sldSz[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/.exec(x);if(m)ar=+m[1]/+m[2]}catch(e){}
 const W=Math.max(300,$('#ofm').clientWidth),H=Math.round(W/ar);
 const p=pv.init(w,{width:W,height:H});await p.preview(O.buf.slice(0));O.w=w;
 const slides=()=>{let s=[...w.querySelectorAll('.pptx-preview-slide-wrapper')];if(!s.length)s=[...w.querySelectorAll('[class*="slide-wrapper"]')];return s};
 const mbtn=await pptMediaSetup(w,slides);
 setBar([{t:'▶ تشغيل',f:()=>window.pptPlay&&window.pptPlay()},{t:'✏️ تعديل',f:()=>window.pptEdit&&window.pptEdit()},{t:'⬇ PPTX',f:()=>dl(new Blob([O.buf],{type:MIME.ppt}),O.d.n+'.pptx')},{t:'🖨 طباعة',f:printOffice},{t:'🔐 نسخة مشفرة',f:encryptedExport},'|',...zoomBtns(()=>w),...mbtn])}

// ---------- منتقي الألوان (ألوان الثيم + قياسية + آخر المستخدمة) ----------
/*CP-START*/
const CP_THEME=['#FFFFFF','#000000','#E7E6E6','#44546A','#4472C4','#ED7D31','#A5A5A5','#FFC000','#5B9BD5','#70AD47'],
CP_STD=['#C00000','#FF0000','#FFC000','#FFFF00','#92D050','#00B050','#00B0F0','#0070C0','#002060','#7030A0'],
CP_TN=['أبيض','أسود','رمادي فاتح','أزرق داكن','تمييز 1','تمييز 2','تمييز 3','تمييز 4','تمييز 5','تمييز 6'];
const cpShade=(h,t)=>{const n=parseInt(h.slice(1),16),c=[n>>16,(n>>8)&255,n&255].map(v=>Math.max(0,Math.min(255,Math.round(t>0?v+(255-v)*t:v*(1+t)))));return '#'+c.map(v=>v.toString(16).padStart(2,'0')).join('').toUpperCase()};
const cpRecent=()=>{try{return JSON.parse(localStorage.getItem('racol')||'[]')}catch(e){return[]}};
const cpPush=c=>{try{localStorage.setItem('racol',JSON.stringify([c,...cpRecent().filter(x=>x!==c)].slice(0,10)))}catch(e){}};
let cpOpen=null;
function cpClose(){if(cpOpen){cpOpen.remove();cpOpen=null}}
document.addEventListener('pointerdown',e=>{if(cpOpen&&!e.target.closest('.cpop,.cpb'))cpClose()},true);
addEventListener('resize',cpClose);
// opt: {label,title,init,theme,onPick}  ->  {el}
function colorBtn(opt){
 const theme=opt.theme&&opt.theme.length>=10?opt.theme.map(c=>c.toUpperCase()):CP_THEME;let cur=(opt.init||'#FF0000').toUpperCase();
 const b=E('button','ofcol cpb','<span>'+(opt.label||'')+'</span><i class="chip"></i>');b.type='button';b.title=opt.title||'';
 const chip=b.querySelector('.chip'),paint=()=>{chip.style.background=cur};paint();
 b.onmousedown=e=>e.preventDefault();
 const pick=c=>{cur=c.toUpperCase();paint();cpPush(cur);cpClose();opt.onPick(c.toLowerCase())};
 const sw=(c,t)=>{const x=E('button','sw');x.type='button';x.style.background=c;x.title=(t?t+' · ':'')+c;x.setAttribute('aria-label',x.title);if(c.toUpperCase()===cur)x.classList.add('cur');x.onclick=()=>pick(c);return x};
 const sec=(t,k)=>{const h=E('h5');h.textContent=t;return h};
 b.onclick=()=>{if(cpOpen){const same=cpOpen.owner===b;cpClose();if(same)return}
  const p=E('div','cpop');p.owner=b;p.onmousedown=e=>{if(!e.target.closest('input'))e.preventDefault()};
  p.append(sec('ألوان الثيم'));
  const g=E('div','cg');theme.forEach((c,i)=>g.append(sw(c,CP_TN[i])));
  [.8,.6,.4,-.25,-.5].forEach(t=>theme.forEach(c=>g.append(sw(cpShade(c,t),''))));p.append(g);
  p.append(sec('ألوان قياسية'));const g2=E('div','cg');CP_STD.forEach(c=>g2.append(sw(c,'')));p.append(g2);
  const rc=cpRecent();if(rc.length){p.append(sec('آخر المستخدمة'));const g3=E('div','cg');rc.forEach(c=>g3.append(sw(c,'')));p.append(g3)}
  const cu=E('label','cc','＋ لون مخصص'),inp=E('input');inp.type='color';inp.value=cur.toLowerCase();inp.onchange=()=>pick(inp.value);cu.append(inp);p.append(cu);
  document.body.append(p);cpOpen=p;
  const r=b.getBoundingClientRect(),w=p.offsetWidth,h=p.offsetHeight;
  p.style.left=Math.max(6,Math.min(innerWidth-w-6,r.left+r.width/2-w/2))+'px';
  p.style.top=(r.bottom+6+h>innerHeight?Math.max(6,r.top-h-6):r.bottom+6)+'px'};
 return {el:b}}
/*CP-END*/
async function createWordFromText(text,name){
 const dx=await need('dx'),tmp=E('div');
 const lines=String(text||'').replace(/\r/g,'').split('\n');
 tmp.innerHTML=lines.map(line=>'<p dir=\"auto\">'+(line.trim()?esc(line):'<br>')+'</p>').join('')||'<p><br></p>';
 const blob=await htmlToDocx(dx,tmp),buf=await blob.arrayBuffer();
 const d=await createRec('doc',name||('مسح ضوئي '+new Date().toLocaleDateString('ar-EG')),buf,1);
 openOffice(d,{edit:1,newDoc:1});return d;
}
window.__of={colorBtn,zipIndex,zipRead,zipBuild,setBar,dl,toast,busy,fail,back,O:()=>O,saveBuf,createWordFromText,saveMeta:async x=>{const d=O&&O.d;if(d&&d.own){Object.assign(d,x);await dbPut(d)}},media:{init:async buf=>{const sc=await pptMediaScan(buf);MP.z=sc.z;MP.items=sc.items;return sc},play:mpPlay,list:mpList,add:m=>{MP.items.push(m)},close:()=>{mpClose();MP.z=null;MP.items=[];MP.cv.clear();MP.autoPromise=null;MP.doc=null}}};
})();
