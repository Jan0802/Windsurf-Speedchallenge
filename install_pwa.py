#!/usr/bin/env python
"""Macht aus der Streamlit-App eine installierbare PWA.

AUFRUF: einmal je Deploy, NACH pip install, VOR dem Start:

    pip install -r requirements.txt && python install_pwa.py

WARUM DAS NOETIG IST
--------------------
Eine PWA braucht drei Dinge an Stellen, an die eine Streamlit-App von sich aus
nicht herankommt:

  1. Ein Manifest, das im <head> des Dokuments verlinkt ist. Streamlit liefert
     eine feste index.html aus; Komponenten laufen in einem iframe und kommen
     dort nicht heran.
  2. Einen Service Worker, der unter "/" ausgeliefert wird. Sein
     GELTUNGSBEREICH ist der Pfad, unter dem er liegt - ein Worker unter
     /app/static/ kontrolliert nur /app/static/ und damit nichts, was zaehlt.
  3. Die Apple-Zusaetze, die iOS braucht, um das Ding als App zu behandeln.

Gemessen (Streamlit 1.58): Das Paketverzeichnis static/ wird an der WURZEL
gemountet. Was dort liegt, erscheint unter "/". Damit ist 2. loesbar - aber nur,
indem Dateien in das installierte Paket gelegt werden. Genau das tut dieses
Skript.

DREI FALLEN, die hier schon eingebaut sind
------------------------------------------
* "manifest.json" ist BELEGT. Unter diesem Namen liegt Vites Build-Manifest
  (eine Asset-Landkarte, kein Web-App-Manifest). Wer sein Manifest so nennt,
  bekommt es nie zu sehen - der Browser laedt still das falsche JSON und die
  Installation scheitert ohne Fehlermeldung. Darum "pwa.webmanifest".

* Streamlit liefert alles unter "/" mit "immutable, max-age=31536000" aus. Ein
  Service Worker unter festem Namen waere ein Jahr lang nicht zu aktualisieren.
  Darum traegt er den Hash seines Inhalts im Namen: sw-<hash>.js. Aendert sich
  der Inhalt, aendert sich der Name, und die Registrierung im (nicht gecachten)
  HTML zeigt auf die neue Datei.

* index.html gehoert dem Paket und wird bei JEDEM Streamlit-Update
  ueberschrieben. Dieses Skript muss also bei jedem Deploy laufen. Findet es
  seinen Ankerpunkt nicht mehr, BRICHT ES AB - ein stilles Durchlaufen waere
  schlimmer, weil die App dann aussieht wie immer und die Installation
  wortlos fehlt.
"""

from __future__ import annotations

import hashlib
import io
import os
import re
import shutil
import sys

HIER = os.path.dirname(os.path.abspath(__file__))
QUELLE = os.path.join(HIER, "pwa")

# Anfang und Ende des eingefuegten Blocks. Sie machen den Lauf wiederholbar:
# Ein vorhandener Block wird ersetzt, nicht ein zweiter angehaengt.
MARKE_AUF = "<!-- pwa:anfang (install_pwa.py) -->"
MARKE_ZU = "<!-- pwa:ende -->"
# Der Ankerpunkt im Streamlit-<head>. Aendert Streamlit ihn, scheitert der Lauf
# sichtbar - siehe Modulkommentar.
ANKER = "<title>Streamlit</title>"


def fehler(text: str) -> None:
    print(f"install_pwa: FEHLER - {text}", file=sys.stderr)
    raise SystemExit(1)


def streamlit_static() -> str:
    try:
        import streamlit
    except ImportError:
        fehler("streamlit ist nicht installiert - laeuft das Skript vor pip?")
    pfad = os.path.join(os.path.dirname(streamlit.__file__), "static")
    if not os.path.isdir(pfad):
        fehler(f"Streamlits static/ nicht gefunden: {pfad}")
    return pfad


def kopf_block(sw_name: str) -> str:
    """Der Block, der in den <head> kommt.

    Die Registrierung steht INLINE und nicht in einer eigenen Datei: Eine
    zusaetzliche Datei unter "/" bekaeme wieder ein Jahr Cache, und dann zeigte
    sie womoeglich auf einen Worker, den es nicht mehr gibt.
    """
    return f"""{MARKE_AUF}
    <link rel="manifest" href="/pwa.webmanifest" />
    <meta name="theme-color" content="#02162b" />
    <meta name="apple-mobile-web-app-capable" content="yes" />
    <meta name="apple-mobile-web-app-status-bar-style" content="black" />
    <meta name="apple-mobile-web-app-title" content="WaterSessions" />
    <link rel="apple-touch-icon" href="/app/static/pwa/apple-touch-icon.png" />
    <script>
      if ('serviceWorker' in navigator) {{
        window.addEventListener('load', function () {{
          navigator.serviceWorker.register('/{sw_name}', {{ scope: '/' }})
            .catch(function () {{ /* ohne Worker laeuft die App normal weiter */ }});
        }});
      }}
    </script>
    {MARKE_ZU}
    """


def main() -> None:
    ziel = streamlit_static()
    index = os.path.join(ziel, "index.html")
    print(f"install_pwa: Ziel {ziel}")

    # --- 1. Service Worker mit Inhalts-Hash im Namen ----------------------
    sw_quelle = os.path.join(QUELLE, "sw.js")
    if not os.path.isfile(sw_quelle):
        fehler(f"{sw_quelle} fehlt")
    sw_text = io.open(sw_quelle, encoding="utf-8").read()
    sw_hash = hashlib.sha256(sw_text.encode("utf-8")).hexdigest()[:10]
    sw_name = f"sw-{sw_hash}.js"

    # Alte Worker-Dateien entfernen, sonst bleibt bei jedem Deploy eine liegen.
    for alt in os.listdir(ziel):
        if re.fullmatch(r"sw-[0-9a-f]{10}\.js", alt) and alt != sw_name:
            os.remove(os.path.join(ziel, alt))
            print(f"install_pwa: alter Worker entfernt: {alt}")
    io.open(os.path.join(ziel, sw_name), "w", encoding="utf-8").write(sw_text)
    print(f"install_pwa: {sw_name} geschrieben")

    # --- 2. Manifest und Offline-Seite ------------------------------------
    for name in ("pwa.webmanifest", "offline.html"):
        q = os.path.join(QUELLE, name)
        if not os.path.isfile(q):
            fehler(f"{q} fehlt")
        shutil.copy2(q, os.path.join(ziel, name))
        print(f"install_pwa: {name} kopiert")

    # --- 3. index.html ergaenzen ------------------------------------------
    if not os.path.isfile(index):
        fehler(f"{index} fehlt")
    html = io.open(index, encoding="utf-8").read()

    # Wiederholbar: einen frueheren Block zuerst herausnehmen.
    vorher = html
    html = re.sub(re.escape(MARKE_AUF) + r".*?" + re.escape(MARKE_ZU),
                  "", html, flags=re.S)
    if html != vorher:
        print("install_pwa: frueheren Block ersetzt")

    if ANKER not in html:
        fehler(
            f"Ankerpunkt {ANKER!r} steht nicht mehr in index.html. Streamlit hat "
            "seinen <head> geaendert - das Skript muss angepasst werden. Es "
            "bricht hier bewusst ab, statt eine App auszuliefern, die aussieht "
            "wie immer und sich nicht installieren laesst."
        )
    html = html.replace(ANKER, kopf_block(sw_name) + ANKER, 1)
    io.open(index, "w", encoding="utf-8", newline="").write(html)
    print("install_pwa: index.html ergaenzt")

    # --- 4. Nachpruefen ----------------------------------------------------
    # Nicht "hat geklappt" behaupten, sondern nachsehen.
    neu = io.open(index, encoding="utf-8").read()
    pruef = [
        ('rel="manifest"', 'rel="manifest"' in neu),
        ("apple-touch-icon", "apple-touch-icon" in neu),
        (f"Registrierung auf /{sw_name}", f"/{sw_name}" in neu),
        ("genau EIN Block", neu.count(MARKE_AUF) == 1),
        (f"{sw_name} liegt da", os.path.isfile(os.path.join(ziel, sw_name))),
        ("pwa.webmanifest liegt da",
         os.path.isfile(os.path.join(ziel, "pwa.webmanifest"))),
        ("offline.html liegt da",
         os.path.isfile(os.path.join(ziel, "offline.html"))),
    ]
    for text, ok in pruef:
        print(f"install_pwa:   {'OK  ' if ok else 'NEIN'} {text}")
    if not all(ok for _, ok in pruef):
        fehler("Nachpruefung fehlgeschlagen")
    print("install_pwa: fertig")


if __name__ == "__main__":
    main()
