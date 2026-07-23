/* DepoApp — Service Worker: instalable + offline + push de OneSignal */

/* OneSignal: se integra acá para no tener dos service workers peleando */
try { importScripts("https://cdn.onesignal.com/sdks/web/v16/OneSignalSDKWorker.js"); } catch (e) {}

const CACHE = "depoapp-v9";
const ASSETS = ["./DepoApp.html", "./manifest.webmanifest", "./icon.svg"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).catch(()=>{}));
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
  );
  self.clients.claim();
});

/* Red primero, cae al cache si no hay conexión (así siempre trae lo último cuando hay internet) */
self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  e.respondWith(
    fetch(e.request).then((res) => {
      const copy = res.clone();
      caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(()=>{});
      return res;
    }).catch(() => caches.match(e.request).then((r) => r || caches.match("./DepoApp.html")))
  );
});

/* Notificaciones push (se activa en la Fase 3, cuando conectemos OneSignal/Web Push) */
self.addEventListener("push", (e) => {
  let data = { title: "DepoApp", body: "Tenés una novedad en tus partidos." };
  try { if (e.data) data = Object.assign(data, e.data.json()); } catch (err) {}
  e.waitUntil(
    self.registration.showNotification(data.title, {
      body: data.body,
      icon: "./icon.svg",
      badge: "./icon.svg",
      data: data.url || "./DepoApp.html"
    })
  );
});

self.addEventListener("notificationclick", (e) => {
  e.notification.close();
  e.waitUntil(clients.openWindow(e.notification.data || "./DepoApp.html"));
});
