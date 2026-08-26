/* DepoApp — Service Worker: instalable + push + auto-actualizacion */
try { importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDKWorker.js"); } catch (e) {}

const CACHE = "depoapp-v18";
const ASSETS = ["./manifest.webmanifest", "./icon-192.png", "./icon-512.png", "./apple-touch-icon.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).catch(()=>{}));
  self.skipWaiting();
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener("message", (e) => { if (e.data === "skipWaiting") self.skipWaiting(); });

/* El HTML SIEMPRE de la red y SIN CACHÉ (cache:"no-store"), así cada vez que se abre
   la app toma la última versión sola, sin reinstalar. Solo cae al cache si no hay internet. */
self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  const url = new URL(e.request.url);
  /* NO tocar pedidos a otros dominios (el backend de Apps Script, OneSignal, etc.).
     El SW solo maneja los archivos propios de la app. Así el backend siempre trae JSON fresco. */
  if (url.origin !== self.location.origin) return;
  const esDoc = e.request.mode === "navigate" || url.pathname.endsWith("DepoApp.html") || url.pathname.endsWith("/");
  if (esDoc) {
    e.respondWith(
      fetch(e.request.url, { cache: "no-store" })
        .then((res) => { const c = res.clone(); caches.open(CACHE).then(x => x.put("./DepoApp.html", c)).catch(()=>{}); return res; })
        .catch(() => caches.match("./DepoApp.html"))
    );
    return;
  }
  e.respondWith(
    fetch(e.request).then((res) => { const c=res.clone(); caches.open(CACHE).then(x=>x.put(e.request,c)).catch(()=>{}); return res; })
    .catch(() => caches.match(e.request))
  );
});

self.addEventListener("push", (e) => {
  let data = { title: "DepoApp", body: "Tenés una novedad en tus partidos." };
  try { if (e.data) data = Object.assign(data, e.data.json()); } catch (err) {}
  e.waitUntil(self.registration.showNotification(data.title, { body: data.body, icon: "./icon.svg", badge: "./icon.svg", data: data.url || "./DepoApp.html" }));
});
self.addEventListener("notificationclick", (e) => { e.notification.close(); e.waitUntil(clients.openWindow(e.notification.data || "./DepoApp.html")); });
