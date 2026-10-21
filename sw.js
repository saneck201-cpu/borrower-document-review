const CACHE='borrower-review-v8';
const FILES=['./','./index.html','./styles.css','./app.js?v=10','./manifest.webmanifest','./icon.svg'];
self.addEventListener('install',event=>event.waitUntil(caches.open(CACHE).then(cache=>cache.addAll(FILES)).then(()=>self.skipWaiting())));
self.addEventListener('activate',event=>event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(key=>key!==CACHE).map(key=>caches.delete(key)))).then(()=>self.clients.claim())));
self.addEventListener('fetch',event=>{const req=event.request;if(req.method!=='GET'||new URL(req.url).origin!==self.location.origin)return;event.respondWith((async()=>{const cached=await caches.match(req,{ignoreSearch:true});try{const fresh=await fetch(req);if(fresh.ok){const cache=await caches.open(CACHE);cache.put(req,fresh.clone())}return fresh}catch{return cached||Response.error()}})())});



