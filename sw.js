// Reader All service worker (web/PWA only — the APK ships its files and skips this).
// Bump V when you want to force every cache to be rebuilt.
const V='reader-all-v50';
const CORE=['./','index.html','app.js','ui.js','wordtpl.js','office.js','wordpro.js','pptedit.js','scan.js','office.css','style.css','fonts.css','fonts-ar.css','manifest.json','icon-192.png','icon-512.png','icon-maskable-512.png'];
const CDN=[];
self.addEventListener('install',e=>{e.waitUntil((async()=>{const c=await caches.open(V);
 await Promise.all([...CORE.map(u=>c.add(u).catch(()=>{})),...CDN.map(u=>c.add(new Request(u,{mode:'cors'})).catch(()=>{}))]);
 self.skipWaiting()})())});
self.addEventListener('activate',e=>{e.waitUntil((async()=>{for(const k of await caches.keys())if(k!==V)await caches.delete(k);await clients.claim()})())});
const keep=(r,res)=>{if(res&&(res.ok||res.type==='opaque'))caches.open(V).then(c=>c.put(r,res.clone())).catch(()=>{});return res};
self.addEventListener('fetch',e=>{const r=e.request;if(r.method!=='GET')return;
 if(/translate\.googleapis\.com|mymemory\.translated\.net/.test(r.url))return; // never cache translations
 const same=new URL(r.url).origin===location.origin;
 // our own files: network first (so updates show up immediately), cache when offline
 // libraries from CDNs: cache first, and remember them the first time they load (so Word/Excel/PowerPoint work offline later)
 e.respondWith(same
  ?fetch(r).then(res=>keep(r,res)).catch(()=>caches.match(r,{ignoreSearch:true}).then(h=>h||Response.error()))
  :caches.match(r).then(h=>h||fetch(r).then(res=>keep(r,res))))});
