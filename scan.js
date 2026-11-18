// Reader All — free on-device OCR (Tesseract.js). First use downloads the free Arabic/English language data;
// Tesseract caches trained data in the browser for later use where storage permits.
(()=>{
'use strict';
const $=s=>document.querySelector(s);
const input=$('#scan-word');if(!input)return;
const loadScript=src=>new Promise((ok,no)=>{const s=document.createElement('script');s.src=src;s.onload=ok;s.onerror=()=>{s.remove();no(new Error('تعذر تحميل مكتبة التعرف على النص'))};document.head.append(s)});
const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function panel(){let w=$('#ocr-panel');if(w)return w;w=document.createElement('div');w.id='ocr-panel';w.innerHTML='<div class="ocr-card"><div class="ocr-head"><b>📷 تحويل الصورة إلى Word</b><button type="button" class="ocr-close">✕</button></div><div class="ocr-state" role="status">جاري تجهيز التعرف على النص…</div><div class="ocr-progress"><i></i></div><textarea dir="auto" placeholder="النص المستخرج هيظهر هنا وتقدر تراجعه وتعدله…" spellcheck="true"></textarea><div class="ocr-actions"><button type="button" class="ocr-cancel">إلغاء</button><button type="button" class="ocr-create" disabled>إنشاء ملف Word</button></div></div>';document.body.append(w);const close=()=>w.remove();w.querySelector('.ocr-close').onclick=close;w.querySelector('.ocr-cancel').onclick=close;return w}
input.addEventListener('change',async()=>{
 const file=input.files&&input.files[0];input.value='';if(!file)return;
 if(!file.type.startsWith('image/')){alert('اختار صورة من فضلك');return}
 const w=panel(),state=w.querySelector('.ocr-state'),bar=w.querySelector('.ocr-progress i'),area=w.querySelector('textarea'),create=w.querySelector('.ocr-create');
 try{
  state.textContent='تحميل محرك OCR المجاني…';
  if(!window.Tesseract)await loadScript('https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js');
  if(!window.Tesseract)throw new Error('محرك OCR غير متاح');
  const worker=await Tesseract.createWorker('ara+eng',1,{
   workerPath:'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/worker.min.js',
   corePath:'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.1.1',
   langPath:'https://tessdata.projectnaptha.com/4.0.0',
   logger:m=>{if(m.status){state.textContent=m.status==='recognizing text'?'جاري قراءة الكلام…':m.status==='loading language traineddata'?'تحميل ملفات اللغة المجانية…':m.status==='initializing tesseract'?'تشغيل محرك القراءة…':m.status}if(typeof m.progress==='number')bar.style.width=Math.round(m.progress*100)+'%'}
  });
  const result=await worker.recognize(file);await worker.terminate();
  area.value=(result.data&&result.data.text||'').trim();
  if(!area.value)throw new Error('لم يتم العثور على نص واضح. جرّب صورة أوضح وبإضاءة جيدة.');
  state.textContent='تم استخراج النص. راجعه وعدّله قبل إنشاء ملف Word.';bar.style.width='100%';create.disabled=false;
 }catch(e){state.textContent=e&&e.message?e.message:'حدث خطأ أثناء قراءة الصورة';bar.style.width='0%'}
 create.onclick=async()=>{const text=area.value.trim();if(!text){state.textContent='اكتب أو استخرج نصًا أولاً';return}try{create.disabled=true;state.textContent='جاري إنشاء ملف Word…';await window.__of.createWordFromText(text,'مستند من صورة '+new Date().toLocaleDateString('ar-EG'));w.remove()}catch(e){create.disabled=false;state.textContent='تعذر إنشاء Word: '+(e.message||e)}};
});
})();
