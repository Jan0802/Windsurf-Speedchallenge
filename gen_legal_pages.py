#!/usr/bin/env python3
"""Erzeugt die statischen Rechtsseiten in ./landing/ aus legal_texts.py.

Warum statisch: Impressum und Datenschutz lagen nur in der Streamlit-App. Damit
hingen zwei Pflichtseiten daran, dass ein Server hochfährt, sie waren für
Suchmaschinen praktisch unsichtbar, und sie liessen sich nicht auf `noindex`
setzen, ohne unerreichbar zu werden. Jetzt liegen sie als eigene Seiten auf der
Landing – und die App zeigt weiter denselben Text aus derselben Quelle.

Ablauf bei Textänderungen:  legal_texts.py pflegen  ->  python gen_legal_pages.py
->  Landing neu deployen.  LEGAL_STAND nicht vergessen.
"""
import os
import re
import sys

from legal_texts import LEGAL_STAND, datenschutz_html, impressum_html

BASE = "https://mywatersessions.com"
APP = "https://app.mywatersessions.com"
SPOTS = "https://spots.mywatersessions.com"

# Aufbau und Farben wie changelog.html, damit die Seiten nicht wie Fremdkörper
# wirken. Absichtlich ohne Google Fonts: die Landing nutzt Systemschriften, das
# spart einen Drittanbieter genau auf den Seiten, die Datenschutz erklären.
TEMPLATE = """<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="icon" href="/favicon.ico" sizes="any">
<link rel="icon" type="image/png" sizes="32x32" href="/favicon-32.png">
<link rel="apple-touch-icon" href="/apple-touch-icon-180.png">
<title>{title}</title>
<meta name="description" content="{desc}">
<link rel="canonical" href="{canon}">
<meta name="theme-color" content="#06303a">
<meta property="og:type" content="website">
<meta property="og:site_name" content="MyWaterSessions">
<meta property="og:title" content="{title}">
<meta property="og:description" content="{desc}">
<meta property="og:url" content="{canon}">
<style>
  :root{{ --aqua:#2bd4d9; --ink:#06303a; }}
  *{{ box-sizing:border-box; }}
  body{{ margin:0; font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;
    color:#eaf4ff; background:#06222e; line-height:1.65; }}
  a{{ color:var(--aqua); }}
  .wrap{{ max-width:820px; margin:0 auto; padding:0 20px; }}
  header.hero{{ background:linear-gradient(180deg, rgba(6,34,46,.55), rgba(6,34,46,.92)),
    radial-gradient(1200px 500px at 70% -10%, #0b6b8c, #06222e 60%);
    padding:44px 0 30px; }}
  .logo{{ margin:0; font-size:clamp(26px,4.4vw,40px); font-weight:800;
    letter-spacing:-.5px; }}
  .logo .dot{{ color:var(--aqua); }}
  .kicker{{ font-size:13px; letter-spacing:3px; text-transform:uppercase;
    color:#8fd9e3; margin-top:10px; }}
  main{{ padding:30px 0 10px; }}
  main h2{{ font-size:clamp(18px,2.6vw,23px); margin:30px 0 8px; }}
  main p, main li{{ color:#dcecf1; }}
  main ul{{ padding-left:1.2rem; }}
  main li{{ margin:6px 0; }}
  code{{ background:rgba(255,255,255,.08); padding:1px 5px; border-radius:4px;
    font-size:.92em; }}
  .lead{{ color:#bfe7ee; }}
  .stand{{ margin-top:34px; padding-top:14px; border-top:1px solid rgba(255,255,255,.14);
    font-size:14px; color:#9fc4cf; }}
  .backline{{ margin:26px 0 0; font-size:15px; }}
  footer{{ padding:30px 0 48px; color:#7fa6b2; font-size:13px; text-align:center; }}
  footer a{{ color:#9fd9e3; margin:0 8px; }}
</style>
</head>
<body>

<header class="hero">
  <div class="wrap">
    <h1 class="logo">MyWaterSessions<span class="dot">.</span></h1>
    <div class="kicker">{kicker}</div>
  </div>
</header>

<main>
  <div class="wrap">
{body}
    <p class="backline"><a href="/">&larr; Zur Startseite</a> &middot;
       <a href="{other_href}">{other_label}</a></p>
  </div>
</main>

<footer>
  <div class="wrap">
    <a href="/">Start</a> ·
    <a href="/guide-de.html">Anleitung</a> ·
    <a href="{spots}/spots">Spots</a> ·
    <a href="/impressum.html">Impressum</a> ·
    <a href="/datenschutz.html">Datenschutz</a> ·
    <a href="{app}/">App öffnen</a>
  </div>
</footer>

</body>
</html>
"""

PAGES = {
    "impressum.html": {
        "title": "Impressum – MyWaterSessions",
        "desc": ("Impressum von MyWaterSessions: Diensteanbieter, Kontakt und "
                 "Verantwortlicher nach § 5 DDG und § 18 Abs. 2 MStV."),
        "kicker": "Impressum",
        "body": impressum_html,
        "other_href": "/datenschutz.html",
        "other_label": "Datenschutzerklärung",
    },
    "datenschutz.html": {
        "title": "Datenschutzerklärung – MyWaterSessions",
        "desc": ("Datenschutzerklärung von MyWaterSessions: welche Daten wir "
                 "verarbeiten, was öffentlich sichtbar ist, welche Dienste "
                 "eingebunden sind und wie du deine Daten löschst."),
        "kicker": "Datenschutz",
        "body": datenschutz_html,
        "other_href": "/impressum.html",
        "other_label": "Impressum",
    },
}


def build():
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "landing")
    if not os.path.isdir(out_dir):
        print(f"FEHLER: {out_dir} gibt es nicht", file=sys.stderr)
        return 1
    for fname, cfg in PAGES.items():
        body = cfg["body"]()
        # Inhalt um zwei Ebenen einruecken, damit die Quelle lesbar bleibt.
        body = "\n".join(("    " + ln) if ln.strip() else ln
                         for ln in body.strip().splitlines())
        html = TEMPLATE.format(
            title=cfg["title"], desc=cfg["desc"], kicker=cfg["kicker"],
            canon=f"{BASE}/{fname}", body=body, app=APP, spots=SPOTS,
            other_href=cfg["other_href"], other_label=cfg["other_label"],
        )
        with open(os.path.join(out_dir, fname), "w", encoding="utf-8") as fh:
            fh.write(html)
        # Kurze Selbstpruefung: keine offenen Platzhalter, Tags ausgeglichen.
        leftover = re.findall(r"\{[a-z_]+\}", html)
        print(f"{fname}: {len(html)} Zeichen, h1={html.count('<h1')}, "
              f"h2={html.count('<h2')}, p={html.count('<p')}/{html.count('</p>')}, "
              f"offene Platzhalter={leftover or 'keine'}")
    print(f"\nStand laut legal_texts.py: {LEGAL_STAND}")
    print("Nicht vergessen: Landing neu deployen (Manual Deploy).")
    return 0


if __name__ == "__main__":
    sys.exit(build())
