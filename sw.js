const CACHE = "sinvena-v1";
const URLS = ["/sinvena/","/sinvena/index.html","/sinvena/manifest.json","/sinvena/icon-192.png","/sinvena/icon-512.png","/sinvena/cropped-iconhka.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(URLS)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.map(k => { if(k!==CACHE) return caches.delete(k); }))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  if(e.request.method!=="GET") return;
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request).then(r => r || caches.match("/sinvena/index.html")))
  );
});
