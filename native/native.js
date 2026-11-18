// Reader All — Android (Capacitor) glue: share/save files, hardware back button.
(()=>{
const C=window.Capacitor;if(!C||!C.isNativePlatform||!C.isNativePlatform())return;
const P=C.Plugins||{},FS=P.Filesystem,SH=P.Share,AP=P.App;
const q=s=>document.querySelector(s);
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
 for(const s of['#mpl .mx','#trc .trx','#plc .no','#mgo .x']){const b=q(s);if(b){b.click();return}}
 if(q('#pl')&&!q('#pl').hidden){q('#plx').click();return}
 if(!q('#mm').hidden||!q('#pc2').hidden){q('#mm').hidden=true;q('#pc2').hidden=true;return}
 if(window.inkBack&&window.inkBack())return;
 if(!q('#ed').hidden){q('#bk').click();return}
 AP.exitApp()});
})();
