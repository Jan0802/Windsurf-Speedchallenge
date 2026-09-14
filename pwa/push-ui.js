// Abonnieren von Benachrichtigungen: eine Leiste, ein Tippen.
//
// WARUM DAS NICHT AUS STREAMLIT HERAUS GEHT: Notification.requestPermission()
// und pushManager.subscribe() muessen im OBERSTEN Dokument laufen und aus einer
// Nutzergeste kommen. Streamlit-Komponenten stecken in einem iframe. Darum
// liegt dieser Code - wie die Installationsleiste - im <head> der Seite.
//
// WIE DIE ANMELDUNG HIERHER KOMMT: Der Server darf nicht jedem erlauben, ein
// Abonnement auf einen fremden Namen einzutragen. Die Web-App schreibt darum
// ein unsichtbares Element mit Name, Zeichen und Gueltigkeit in die Seite
// (#ws-push-auth), und das wird hier gelesen. Streamlit rendert
// unsafe_allow_html in das oberste Dokument, nicht in ein iframe - deshalb
// funktioniert das.
//
// AUF DEM iPHONE gibt es Push NUR fuer eine installierte PWA. Im Safari-Tab
// existiert window.PushManager gar nicht. Die Faehigkeitspruefung deckt das
// mit ab, ohne nach dem Geraet zu fragen - der Schalter im Profil fragt dann
// doch danach, aber erst, um den GRUND zu nennen statt nur leer zu bleiben.

(function () {
  'use strict';

  if (window.top !== window.self) { return; }

  // FAEHIG statt Ausstieg: Frueher endete das Skript hier, wenn der Browser
  // kein Push kann. Damit blieb aber auch der Schalter im Profil leer - und
  // genau dort will jemand ja erfahren, WARUM nichts geht. Auf dem iPhone ist
  // die Antwort fast immer "weil die Seite im Safari-Tab laeuft statt als
  // installierte App", und das sagt einem sonst niemand.
  var FAEHIG = ('serviceWorker' in navigator) && ('PushManager' in window)
    && ('Notification' in window);

  var SPEICHER = 'ws-push-ask';
  var WARTEN_MS = 4000;      // nach der Installationsleiste, nicht daneben
  var INGEST = 'https://spots.mywatersessions.com';
  var SLOT = 'ws-push-slot';  // das leere Element, das die App ins Profil legt
  var NACHSEHEN_MS = 800;     // Streamlit baut die Seitenleiste staendig neu

  function erledigt(grund) {
    try { localStorage.setItem(SPEICHER, grund); } catch (e) {}
  }
  function schonGefragt() {
    try { return !!localStorage.getItem(SPEICHER); } catch (e) { return true; }
  }

  // base64url -> Uint8Array. Der VAPID-Schluessel kommt als Text, subscribe()
  // will Bytes.
  function schluesselBytes(b64) {
    var rein = (b64 + '='.repeat((4 - b64.length % 4) % 4))
      .replace(/-/g, '+').replace(/_/g, '/');
    var roh = atob(rein);
    var a = new Uint8Array(roh.length);
    for (var i = 0; i < roh.length; i++) { a[i] = roh.charCodeAt(i); }
    return a;
  }

  // Die Anmeldedaten, die die Web-App in die Seite schreibt. Sie stehen erst da,
  // wenn Streamlit fertig gerendert hat - also wird kurz gewartet statt einmal
  // vergeblich nachgesehen.
  function authHolen(cb, versuche) {
    var el = document.getElementById('ws-push-auth');
    if (el && el.dataset.token) {
      cb({ name: el.dataset.name || '', token: el.dataset.token,
           exp: parseInt(el.dataset.exp, 10) });
      return;
    }
    if ((versuche || 0) > 20) { cb(null); return; }
    setTimeout(function () { authHolen(cb, (versuche || 0) + 1); }, 500);
  }

  function el(tag, css, text) {
    var n = document.createElement(tag);
    n.style.cssText = css;
    if (text) { n.textContent = text; }
    return n;
  }

  function zeigen(auth) {
    var huelle = el('div',
      'position:fixed;left:0;right:0;bottom:0;z-index:2147483000;' +
      'background:#0a2338;color:#eaf4ff;border-top:1px solid rgba(255,255,255,.14);' +
      'box-shadow:0 -8px 30px rgba(0,0,0,.45);' +
      'padding:16px 16px calc(16px + env(safe-area-inset-bottom));' +
      'font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif;' +
      'transform:translateY(110%);transition:transform .28s ease;');
    var kopf = el('div', 'display:flex;gap:12px;align-items:center;');
    kopf.appendChild(el('div', 'font-size:26px;flex:0 0 auto;', '🔔'));
    var text = el('div', 'flex:1 1 auto;');
    text.appendChild(el('div', 'font-weight:700;', 'Get a nudge when it matters'));
    text.appendChild(el('div', 'font-size:13.5px;opacity:.8;',
      'A spot record falls, or your safety check-in is overdue. Nothing else.'));
    kopf.appendChild(text);
    huelle.appendChild(kopf);

    var reihe = el('div', 'display:flex;gap:10px;margin-top:14px;');
    var nein = el('button',
      'flex:0 0 auto;background:transparent;color:#9fc4cf;border:0;' +
      'padding:12px 14px;font-size:15px;cursor:pointer;', 'No thanks');
    var ja = el('button',
      'flex:1 1 auto;background:#2bd4d9;color:#02162b;border:0;' +
      'border-radius:999px;padding:13px 20px;font-size:15px;font-weight:700;' +
      'cursor:pointer;', 'Turn on');
    reihe.appendChild(nein);
    reihe.appendChild(ja);
    huelle.appendChild(reihe);

    function weg(grund) {
      erledigt(grund);
      huelle.style.transform = 'translateY(110%)';
      setTimeout(function () {
        if (huelle.parentNode) { huelle.parentNode.removeChild(huelle); }
      }, 300);
    }

    nein.addEventListener('click', function () { weg('no'); });
    ja.addEventListener('click', function () {
      // requestPermission() MUSS in der Nutzergeste beginnen. Alles Weitere
      // haengt an dessen Zusage und darf danach kommen.
      Notification.requestPermission().then(function (antwort) {
        if (antwort !== 'granted') { weg('denied'); return; }
        weg('granted');
        abonnieren(auth);
      }).catch(function () { weg('error'); });
    });

    document.body.appendChild(huelle);
    void huelle.offsetHeight;
    huelle.style.transform = 'translateY(0)';
  }

  // fertig(ok, grund) ist freiwillig: Die Leiste braucht keine Rueckmeldung,
  // der Schalter im Profil schon - dort ist der Grund das ganze Produkt.
  function abonnieren(auth, fertig) {
    fetch(INGEST + '/push_key').then(function (r) { return r.json(); })
      .then(function (k) {
        if (!k.configured || !k.key) { throw new Error('kein VAPID-Schluessel'); }
        return navigator.serviceWorker.ready.then(function (reg) {
          return reg.pushManager.subscribe({
            // Pflicht: Jeder Push muss eine sichtbare Benachrichtigung
            // erzeugen. Ohne das entziehen Browser die Erlaubnis wieder.
            userVisibleOnly: true,
            applicationServerKey: schluesselBytes(k.key)
          });
        });
      })
      .then(function (sub) {
        var j = sub.toJSON();
        return fetch(INGEST + '/push_subscribe', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            exp: auth.exp, token: auth.token, name: auth.name,
            endpoint: j.endpoint,
            p256dh: j.keys.p256dh, auth: j.keys.auth,
            ua: navigator.userAgent
          })
        });
      })
      .then(function (r) {
        if (!r.ok) { throw new Error('Server ' + r.status); }
        // Eine Rueckmeldung gehoert dazu: Wer gerade "Erlauben" getippt hat,
        // will wissen, dass es angekommen ist - sonst tippt er noch einmal.
        try {
          new Notification('Notifications are on',
            { body: 'You will hear from us when something happens.',
              icon: '/app/static/pwa/icon-192.png' });
        } catch (e) {}
        if (fertig) { fertig(true, ''); }
      })
      .catch(function (e) {
        // Fehlschlag darf nicht als "erledigt" gelten - sonst fragt die App nie
        // wieder, obwohl nichts eingerichtet ist.
        try { localStorage.removeItem(SPEICHER); } catch (e2) {}
        if (window.console) { console.warn('Push nicht eingerichtet:', e); }
        if (fertig) { fertig(false, String((e && e.message) || e)); }
      });
  }

  // ---------------------------------------------------------------------
  // Der Schalter im Profil
  //
  // WARUM NICHT ALS STREAMLIT-KNOPF: requestPermission() und subscribe()
  // verlangen eine echte Nutzergeste im obersten Dokument. Ein st.button loest
  // einen Rerun aus - die Geste waere vorbei, bevor gefragt wird. Und
  // components.html steckt in einem iframe, in dem der Browser
  // Benachrichtigungen gar nicht erst erlaubt. Also legt die App ein leeres
  // Element, und dieses Skript macht den Schalter daraus.
  //
  // WOZU UEBERHAUPT: Die Leiste beim Start fragt genau EINMAL. Wer sie
  // wegtippt oder bei wem das Abonnieren scheitert, haette sonst keinen Weg
  // zurueck - ausser die App zu loeschen und neu zu installieren.
  // ---------------------------------------------------------------------

  function istApple() {
    return /iPad|iPhone|iPod/.test(navigator.userAgent || '');
  }
  function alsAppGeoeffnet() {
    try {
      return window.matchMedia('(display-mode: standalone)').matches
        || navigator.standalone === true;
    } catch (e) { return false; }
  }

  function knopf(text, haupt) {
    return el('button',
      'flex:1 1 auto;border-radius:999px;padding:10px 14px;cursor:pointer;' +
      'font:600 14px/1.2 system-ui,-apple-system,"Segoe UI",sans-serif;' +
      (haupt ? 'background:#2bd4d9;color:#02162b;border:0;'
             : 'background:transparent;color:#9fc4cf;' +
               'border:1px solid rgba(159,196,207,.45);'),
      text);
  }

  function montieren(ziel) {
    ziel.dataset.wsMounted = '1';
    ziel.textContent = '';
    var lage = el('div', 'margin-bottom:10px;font:14px/1.5 system-ui,' +
      '-apple-system,"Segoe UI",sans-serif;');
    var reihe = el('div', 'display:flex;gap:8px;flex-wrap:wrap;');
    ziel.appendChild(lage);
    ziel.appendChild(reihe);

    function sagen(text, farbe) {
      lage.textContent = text;
      lage.style.color = farbe || '';
    }
    function neu() { delete ziel.dataset.wsMounted; montieren(ziel); }

    if (!FAEHIG) {
      // Der haeufigste Fall, und der einzige, den man selbst beheben kann.
      sagen(istApple() && !alsAppGeoeffnet()
        ? 'On iPhone and iPad this only works when the app runs from your '
          + 'home screen. In Safari: Share → scroll down → Add to Home Screen, '
          + 'then open it from there.'
        : 'This browser cannot do notifications.', '#ffc09f');
      return;
    }
    if (Notification.permission === 'denied') {
      sagen('Blocked. Allow notifications for MyWaterSessions in your device '
        + 'settings, then reopen the app.', '#ffc09f');
      return;
    }

    sagen('Checking…');
    authHolen(function (auth) {
      if (!auth) { sagen('Sign in to turn notifications on.'); return; }
      navigator.serviceWorker.ready.then(function (reg) {
        return reg.pushManager.getSubscription();
      }).then(function (vorhanden) {
        if (vorhanden) { istAn(auth, vorhanden); } else { istAus(auth); }
      }).catch(function (e) {
        sagen('Could not read the state: ' + ((e && e.message) || e), '#ffc09f');
      });
    });

    function istAus(auth) {
      sagen('Notifications are off.');
      var an = knopf('Turn on', true);
      reihe.appendChild(an);
      an.addEventListener('click', function () {
        // Eine frueher gespeicherte Absage zuruecknehmen - sonst bliebe sie
        // liegen und die Startleiste kaeme auch nach dem Einschalten nie.
        try { localStorage.removeItem(SPEICHER); } catch (e) {}
        an.disabled = true;
        sagen('Waiting for your permission…');
        Notification.requestPermission().then(function (antwort) {
          if (antwort !== 'granted') {
            an.disabled = false;
            sagen('Permission was not granted.', '#ffc09f');
            return;
          }
          abonnieren(auth, function (ok, grund) {
            if (ok) { neu(); } else {
              an.disabled = false;
              sagen('Could not turn on: ' + grund, '#ffc09f');
            }
          });
        }).catch(function () {
          an.disabled = false;
          sagen('Could not ask for permission.', '#ffc09f');
        });
      });
    }

    function istAn(auth, sub) {
      sagen('Notifications are on. ✅');
      var pruefen = knopf('Send test', true);
      var aus = knopf('Turn off', false);
      reihe.appendChild(pruefen);
      reihe.appendChild(aus);

      pruefen.addEventListener('click', function () {
        pruefen.disabled = true;
        sagen('Sending…');
        senden('/push_selftest', { exp: auth.exp, token: auth.token,
                                   name: auth.name })
          .then(function (a) {
            pruefen.disabled = false;
            if (a.sent > 0) {
              sagen('Sent to ' + a.sent + ' device'
                + (a.sent === 1 ? '' : 's') + '. If nothing shows up, check '
                + 'notification settings for this app.');
            } else if (!a.devices) {
              sagen('No device is registered – turn it off and on again.',
                    '#ffc09f');
            } else {
              sagen('Delivery failed: ' + (a.reasons && a.reasons[0] || '?'),
                    '#ffc09f');
            }
          })
          .catch(function (e) {
            pruefen.disabled = false;
            sagen('Test failed: ' + ((e && e.message) || e), '#ffc09f');
          });
      });

      aus.addEventListener('click', function () {
        aus.disabled = true;
        sagen('Turning off…');
        // Erst beim Browser kuendigen, dann beim Server austragen. Andersherum
        // bliebe bei einem Abbruch ein Abonnement, das niemand mehr kennt.
        sub.unsubscribe().then(function () {
          return senden('/push_unsubscribe', {
            exp: auth.exp, token: auth.token, name: auth.name,
            endpoint: sub.endpoint });
        }).then(function () {
          // Bewusst ausgeschaltet heisst: nicht beim naechsten Start wieder
          // fragen.
          erledigt('off');
          neu();
        }).catch(function (e) {
          aus.disabled = false;
          sagen('Could not turn off: ' + ((e && e.message) || e), '#ffc09f');
        });
      });
    }
  }

  function senden(pfad, daten) {
    return fetch(INGEST + pfad, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(daten)
    }).then(function (r) {
      if (!r.ok) { throw new Error('Server ' + r.status); }
      return r.json();
    });
  }

  // Schon erlaubt? Dann still nachziehen, falls das Abonnement fehlt (nach
  // Browserdaten-Loeschung, nach Neuinstallation). Ohne das steht die Erlaubnis
  // auf "erteilt" und es kommt trotzdem nie etwas an.
  function nachziehen(auth) {
    navigator.serviceWorker.ready.then(function (reg) {
      reg.pushManager.getSubscription().then(function (vorhanden) {
        if (!vorhanden) { abonnieren(auth); }
      });
    });
  }

  // Nachsehen statt einmal suchen: Streamlit baut die Seitenleiste bei jedem
  // Rerun neu auf. Das Element ist danach ein anderes - ohne die Markierung,
  // also wird es wieder bestueckt. Ein einmaliger Fund waere nach dem ersten
  // Klick irgendwo in der App verschwunden.
  setInterval(function () {
    var s = document.getElementById(SLOT);
    if (s && !s.dataset.wsMounted) { montieren(s); }
  }, NACHSEHEN_MS);

  window.addEventListener('load', function () {
    if (!FAEHIG) { return; }   // die Leiste haette hier nichts anzubieten
    setTimeout(function () {
      if (Notification.permission === 'denied') { return; }
      authHolen(function (auth) {
        if (!auth) { return; }            // nicht angemeldet - dann kein Abo
        if (Notification.permission === 'granted') { nachziehen(auth); return; }
        if (schonGefragt()) { return; }
        zeigen(auth);
      });
    }, WARTEN_MS);
  });
})();
