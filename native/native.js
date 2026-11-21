// Reader All — Android (Capacitor) glue: share/save files, hardware back button.
(()=>{
const C=window.Capacitor;if(!C||!C.isNativePlatform||!C.isNativePlatform())return;
const P=C.Plugins||{},FS=P.Filesystem,SH=P.Share,AP=P.App,CAM=P.Camera;
const q=s=>document.querySelector(s);

// ---- microphone: ask for the Android permission, and native Arabic dictation (WebView has no SpeechRecognition) ----
const SRP=P.SpeechRecognition;
window.raEnsureMic=async()=>{if(SRP&&SRP.requestPermissions){try{await SRP.requestPermissions()}catch(e){}}};
if(SRP)window.raNativeDictate=async lang=>{
 const av=await SRP.available();if(!av||!av.available)throw new Error('خدمة التعرف على الصوت مش متاحة. ثبّت أو فعّل تطبيق Google.');
 await window.raEnsureMic();
 const r=await SRP.start({language:lang||'ar-EG',maxResults:1,prompt:'اتكلم الآن',partialResults:false,popup:true});
 return r&&r.matches&&r.matches[0]||''};

// ---- receive Office files from Android Open with / Share intents ----
// Java sends large files in base64 chunks; reassemble them, then dispatch the same event as before.
const chunks={};
window.__raChunk=(id,name,type,i,total,data)=>{
 const c=chunks[id]||(chunks[id]={parts:new Array(total),got:0});
 if(c.parts[i]===undefined){c.parts[i]=data;c.got++}
 if(c.got===total){delete chunks[id];window.dispatchEvent(new CustomEvent('readerall-file-received',{detail:{name,type,data:c.parts.join('')}}))}
};
let incomingQueue=Promise.resolve();
window.addEventListener('readerall-file-received',ev=>{
 const d=ev.detail||{};
 incomingQueue=incomingQueue.then(async()=>{
  try{
   if(!d.data||!d.name)return;
   const raw=atob(d.data),bytes=new Uint8Array(raw.length);
   for(let i=0;i<raw.length;i++)bytes[i]=raw.charCodeAt(i);
   const mime=d.type||'application/octet-stream';
   const extByMime={
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document':'.docx',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':'.xlsx',
    'application/vnd.ms-excel':'.xls','text/csv':'.csv',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation':'.pptx'
   };
   const name=/\.(docx|xlsx|xls|csv|pptx)$/i.test(d.name)?d.name:d.name+(extByMime[mime]||'');
   const file=new File([bytes],name,{type:mime});
   if(typeof window.raImportOffice==='function')await window.raImportOffice(file);
   else throw new Error('لم يتم تجهيز قارئ الملفات بعد');
  }catch(e){alert('تعذر فتح الملف المستلم: '+(e&&e.message||e))}
 });
});
window.__raReady=true; // lets MainActivity know the receiver exists
// Ask for camera permission only when the scan-to-Word control is used.
const scanInput=document.querySelector('#scan-word');
if(scanInput&&CAM){scanInput.addEventListener('click',async ev=>{
 ev.preventDefault();ev.stopImmediatePropagation();
 try{
  const perm=await CAM.requestPermissions({permissions:['camera']});
  if(perm.camera!=='granted'){alert('لازم تسمح باستخدام الكاميرا علشان تصور الورقة.');return}
  const photo=await CAM.getPhoto({quality:90,allowEditing:false,resultType:'uri',source:'CAMERA'});
  if(!photo.webPath)return;
  const blob=await (await fetch(photo.webPath)).blob();
  const file=new File([blob],'scan-'+Date.now()+'.jpg',{type:blob.type||'image/jpeg'});
  const dt=new DataTransfer();dt.items.add(file);scanInput.files=dt.files;
  scanInput.dispatchEvent(new Event('change',{bubbles:true}));
 }catch(e){if(e&&e.message)alert('تعذر تشغيل الكاميرا: '+e.message)}
},true)}

// ---- save / share files (blob downloads and Web Share don't work inside the WebView) ----
async function out(blob,name){
 if(!FS||!SH)throw new Error('الإضافات الأصلية مش متاحة');
 const b64=await new Promise((ok,no)=>{const r=new FileReader();r.onload=()=>ok(String(r.result).split(',')[1]);r.onerror=()=>no(r.error);r.readAsDataURL(blob)});
 const w=await FS.writeFile({path:name,data:b64,directory:'CACHE'});
 await SH.share({title:name,url:w.uri,dialogTitle:name})}
const def=(k,v)=>{try{Object.defineProperty(navigator,k,{value:v,configurable:true,writable:true})}catch(e){}};
def('canShare',d=>!!(d&&d.files&&d.files.length));
def('share',async d=>{for(const f of (d&&d.files)||[])await out(f,f.name)});
const click=HTMLAnchorElement.prototype.click;
HTMLAnchorElement.prototype.click=function(){
 if(this.download&&/^blob:/.test(this.href)){const n=this.download;
  fetch(this.href).then(r=>r.blob()).then(b=>out(b,n)).catch(e=>alert('تعذر الحفظ: '+(e&&e.message||e)));return}
 return click.call(this)};
// ---- Android back button ----
if(AP&&AP.addListener)AP.addListener('backButton',()=>{
 const d=q('.dlg .dc')||q('.dlg .do');if(d){d.click();return}
 for(const s of['.dk-ov .dk-x','.pv-m .pv-x','#mpl .mx','#trc .trx','#plc .no','#mgo .x']){const b=q(s);if(b){b.click();return}}
 if(q('#pl')&&!q('#pl').hidden){q('#plx').click();return}
 if(!q('#mm').hidden||!q('#pc2').hidden){q('#mm').hidden=true;q('#pc2').hidden=true;return}
 if(window.raBack&&window.raBack())return;
 if(!q('#ed').hidden){q('#bk').click();return}
 AP.exitApp()});
})();
