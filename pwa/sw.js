// Service Worker fuer MyWaterSessions.
//
// BEWUSST MINIMAL, und das ist keine Faulheit: Streamlit rendert die Seite
// serverseitig ueber eine WebSocket-Verbindung. Ohne Server gibt es keine
// Oberflaeche - auch keine zuletzt gesehene. Ein Worker, der hier die App-Huelle
// cachte, wuerde offline eine leere weisse Flaeche ausliefern und damit etwas
// versprechen, was dahinter nicht steht.
//
// Er hat also genau zwei Aufgaben:
//   1. Android verlangt fuer die Installationsaufforderung einen Worker MIT
//      fetch-Handler. Ohne ihn erscheint der Knopf nicht.
//   2. Wer die installierte App ohne Netz oeffnet, soll eine Seite sehen, die
//      das erklaert, statt des Browser-Dinosauriers.
//
// Der Dateiname traegt eine Version (sw-<hash>.js, siehe install_pwa.py). Das
// ist noetig, weil Streamlits Auslieferung alles unter "/" mit
// "immutable, max-age=31536000" ausliefert - ein Worker unter festem Namen waere
// ein Jahr lang nicht zu aktualisieren. Ueber den Namen geht es trotzdem: Die
// Registrierung steht im HTML, und das wird nicht gecacht.

const CACHE = 'ws-shell-v1';
const OFFLINE = '/offline.html';

// Nur die Notfallseite und ihr Icon. Mehr waere geraten.
const VORRAT = [OFFLINE, '/app/static/pwa/icon-192.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE).then((c) => c.addAll(VORRAT)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  // Alte Staende wegraeumen, sonst sammeln sich mit jeder Version neue an.
  e.waitUntil(
    caches.keys()
      .then((namen) => Promise.all(
        namen.filter((n) => n !== CACHE).map((n) => caches.delete(n))
      ))
      .then(() => self.clients.claim())
  );
});

// --- Benachrichtigungen ---------------------------------------------------
// Der Push-Dienst weckt den Worker, auch wenn keine Seite offen ist - genau das
// ist der Sinn. Der Inhalt kommt verschluesselt an und ist nur hier lesbar.
//
// showNotification ist PFLICHT: Wir haben mit userVisibleOnly abonniert, und
// ein Push ohne sichtbare Benachrichtigung fuehrt dazu, dass der Browser die
// Erlaubnis nach mehreren Verstoessen von sich aus entzieht.
self.addEventListener('push', (e) => {
  let d = { title: 'MyWaterSessions', body: '', url: '/' };
  try { if (e.data) { d = Object.assign(d, e.data.json()); } }
  catch (err) { try { d.body = e.data.text(); } catch (e2) {} }
  e.waitUntil(self.registration.showNotification(d.title, {
    body: d.body,
    icon: '/app/static/pwa/icon-192.png',
    badge: '/app/static/pwa/icon-192.png',
    data: { url: d.url || '/' },
    // Gleiches tag = die neue Nachricht ERSETZT die alte, statt sich zu
    // stapeln. Wer drei Tage nicht hineingesehen hat, soll nicht drei
    // Erinnerungen vorfinden.
    tag: 'ws-note',
    renotify: true
  }));
});

// Tippen soll die App oeffnen - und zwar die BEREITS offene, wenn es eine gibt.
// Sonst sammelt jeder Fingertipp einen weiteren Tab an.
self.addEventListener('notificationclick', (e) => {
  e.notification.close();
  const ziel = (e.notification.data && e.notification.data.url) || '/';
  e.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true })
      .then((liste) => {
        for (const c of liste) {
          if ('focus' in c) { c.navigate(ziel); return c.focus(); }
        }
        return self.clients.openWindow(ziel);
      })
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  // NUR Seitenaufrufe abfangen. Alles andere - die WebSocket-Verbindung, die
  // JS-Bundles, die Bilder - geht unveraendert ans Netz: Dazwischenzufunken
  // waere der sicherste Weg, eine funktionierende App zu zerlegen.
  if (req.mode !== 'navigate') { return; }
  e.respondWith(
    fetch(req).catch(() => caches.match(OFFLINE))
  );
});
