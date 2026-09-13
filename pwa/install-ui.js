// Fragt auf dem Handy einmal nach, ob die App auf den Startbildschirm soll.
//
// WAS AUF WELCHEM SYSTEM GEHT - das ist der Kern, und die beiden Seiten sind
// NICHT gleich:
//
//   Android/Chrome: Der Browser meldet sich selbst mit 'beforeinstallprompt'.
//     Wir fangen das Ereignis ab, zeigen unsere eigene Leiste, und auf Tippen
//     oeffnet prompt() den ECHTEN Installationsdialog des Systems. Danach sagt
//     'appinstalled', dass es geklappt hat.
//
//   iPhone/iPad: Es gibt KEINE programmatische Installation. Safari bietet
//     kein Ereignis und keine Methode an - wir koennen nur die drei Schritte
//     zeigen. Und wir koennen aus Safari heraus AUCH NICHT ERKENNEN, ob die App
//     schon auf dem Startbildschirm liegt: navigator.standalone ist nur wahr,
//     wenn die Seite AUS der installierten App heraus laeuft. Wer bei uns
//     "Verstanden" tippt, gilt darum als erledigt - mehr ist ehrlich nicht
//     moeglich.
//
// Gefragt wird genau einmal. Ein Nein wird gemerkt und nie wieder gestellt.

(function () {
  'use strict';

  // In einem Streamlit-Komponenten-iframe hat die Leiste nichts zu suchen.
  if (window.top !== window.self) { return; }

  var SPEICHER = 'ws-install-ask';
  // Kurz warten, damit der Nutzer zuerst die Seite sieht und nicht eine
  // Aufforderung auf leerem Grund. Siehe auch den Hinweis unten zum Zeitpunkt.
  var WARTEN_MS = 2500;

  function erledigt(grund) {
    try { localStorage.setItem(SPEICHER, grund); } catch (e) {}
  }
  function schonGefragt() {
    try { return !!localStorage.getItem(SPEICHER); } catch (e) { return true; }
  }

  // Laeuft die Seite bereits als installierte App? Dann nie fragen.
  function installiert() {
    return (window.matchMedia &&
            window.matchMedia('(display-mode: standalone)').matches) ||
           window.navigator.standalone === true;
  }

  var ua = navigator.userAgent || '';
  var istIOS = /iPad|iPhone|iPod/.test(ua) ||
               // iPadOS meldet sich als Mac - am Touch erkennbar.
               (/Macintosh/.test(ua) && navigator.maxTouchPoints > 1);
  // In fremden In-App-Browsern (Instagram, Facebook, LinkedIn) gibt es
  // "Zum Home-Bildschirm" gar nicht. Dort zu fragen waere eine Anleitung ins
  // Leere.
  var inAppBrowser = /FBAN|FBAV|Instagram|Line\/|LinkedInApp|Twitter/.test(ua);
  var istHandy = window.matchMedia &&
                 window.matchMedia('(max-width: 860px)').matches;

  if (installiert()) { erledigt('installed'); return; }
  if (schonGefragt() || !istHandy || inAppBrowser) { return; }

  // --- Aussehen ----------------------------------------------------------
  // Alles inline: Der Streamlit-Stil laedt spaeter und wuerde eigene Klassen
  // ueberschreiben.
  function el(tag, css, text) {
    var n = document.createElement(tag);
    n.style.cssText = css;
    if (text) { n.textContent = text; }
    return n;
  }

  function zeigen(istApple, aufInstallieren) {
    var huelle = el('div',
      'position:fixed;left:0;right:0;bottom:0;z-index:2147483000;' +
      'background:#0a2338;color:#eaf4ff;border-top:1px solid rgba(255,255,255,.14);' +
      'box-shadow:0 -8px 30px rgba(0,0,0,.45);padding:16px 16px calc(16px + env(safe-area-inset-bottom));' +
      'font:15px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif;' +
      'transform:translateY(110%);transition:transform .28s ease;');

    var kopf = el('div', 'display:flex;gap:12px;align-items:center;');
    var bild = document.createElement('img');
    bild.src = '/app/static/pwa/icon-192.png';
    bild.alt = '';
    bild.style.cssText = 'width:42px;height:42px;border-radius:11px;flex:0 0 auto;';
    var text = el('div', 'flex:1 1 auto;');
    text.appendChild(el('div', 'font-weight:700;', 'Add to your home screen'));
    text.appendChild(el('div', 'font-size:13.5px;opacity:.8;',
      'Opens full screen, like an app. No app store, no download.'));
    kopf.appendChild(bild);
    kopf.appendChild(text);
    huelle.appendChild(kopf);

    if (istApple) {
      var schritte = document.createElement('ol');
      schritte.style.cssText =
        'margin:12px 0 0;padding-left:20px;font-size:14px;opacity:.9;';
      [['Tap the ', 'Share', ' button below.'],
       ['', 'Scroll down', ' — “Add to Home Screen” sits further down.'],
       ['Tap ', 'Add', ' at the top right.']].forEach(function (t) {
        var li = document.createElement('li');
        li.style.margin = '5px 0';
        li.appendChild(document.createTextNode(t[0]));
        var b = document.createElement('b');
        b.textContent = t[1];
        li.appendChild(b);
        li.appendChild(document.createTextNode(t[2]));
        schritte.appendChild(li);
      });
      huelle.appendChild(schritte);
    }

    var reihe = el('div', 'display:flex;gap:10px;margin-top:14px;');
    var nein = el('button',
      'flex:0 0 auto;background:transparent;color:#9fc4cf;border:0;' +
      'padding:12px 14px;font-size:15px;cursor:pointer;',
      istApple ? 'Not now' : 'Not now');
    var ja = el('button',
      'flex:1 1 auto;background:#2bd4d9;color:#02162b;border:0;border-radius:999px;' +
      'padding:13px 20px;font-size:15px;font-weight:700;cursor:pointer;',
      istApple ? 'Got it' : 'Install');
    reihe.appendChild(nein);
    reihe.appendChild(ja);
    huelle.appendChild(reihe);

    function schliessen(grund) {
      erledigt(grund);
      huelle.style.transform = 'translateY(110%)';
      setTimeout(function () {
        if (huelle.parentNode) { huelle.parentNode.removeChild(huelle); }
      }, 300);
    }

    nein.addEventListener('click', function () { schliessen('no'); });
    ja.addEventListener('click', function () {
      if (istApple) {
        // Mehr als die Anleitung geht hier nicht - siehe Kopf der Datei.
        schliessen('ios-shown');
      } else {
        // Das ECHTE Systemfenster. Danach nicht mehr fragen, egal wie der
        // Nutzer dort entscheidet: Ein zweites Mal waere Draengeln.
        schliessen('prompted');
        aufInstallieren();
      }
    });

    document.body.appendChild(huelle);
    // Erzwingt einen Layoutdurchlauf, damit die Einblendung wirklich animiert.
    void huelle.offsetHeight;
    huelle.style.transform = 'translateY(0)';
  }

  // Eine Installation zaehlt immer als erledigt - auch wenn sie ueber das
  // Browsermenue lief und nicht ueber unsere Leiste.
  window.addEventListener('appinstalled', function () { erledigt('installed'); });

  if (istIOS) {
    window.addEventListener('load', function () {
      setTimeout(function () {
        if (!schonGefragt() && !installiert()) { zeigen(true, null); }
      }, WARTEN_MS);
    });
    return;
  }

  // Android: auf das Angebot des Browsers warten. Kommt es nicht (Seite nicht
  // installierbar, schon installiert, Browser kann es nicht), fragen wir auch
  // nicht - eine Leiste mit einem Knopf, der nichts tut, waere schlimmer als
  // keine Leiste.
  window.addEventListener('beforeinstallprompt', function (e) {
    e.preventDefault();
    var zurueckgehalten = e;
    setTimeout(function () {
      if (schonGefragt() || installiert()) { return; }
      zeigen(false, function () {
        zurueckgehalten.prompt();
      });
    }, WARTEN_MS);
  });
})();
