#!/usr/bin/env python3
"""Erzeugt die statischen Rechtsseiten in ./landing/ aus legal_texts.py.

Vier Seiten, zwei Sprachen:
    /impressum.html   /datenschutz.html   (deutsch)
    /imprint.html     /privacy.html       (englisch)

Warum statisch: Impressum und Datenschutz lagen nur in der Streamlit-App. Damit
hingen zwei Pflichtseiten daran, dass ein Server hochfährt, sie waren für
Suchmaschinen unsichtbar, und sie liessen sich nicht auf `noindex` setzen, ohne
unerreichbar zu werden.

Die Namen folgen der Sprache, nicht dem Suffix-Schema der uebrigen Seiten
(index-de.html): "impressum" und "datenschutz" sind die Begriffe, nach denen
deutschsprachige Besucher suchen, "imprint" und "privacy" die englischen. Die
Paare sind ueber hreflang verknuepft, x-default zeigt auf Englisch.

Ablauf bei Textänderungen:  legal_texts.py pflegen (BEIDE Sprachen)
->  python gen_legal_pages.py  ->  Landing neu deployen.  LEGAL_STAND nicht
vergessen.
"""
import os
import re
import sys

from legal_texts import (
    LEGAL_STAND,
    datenschutz_html,
    imprint_html_en,
    impressum_html,
    privacy_html_en,
)

BASE = "https://mywatersessions.com"
APP = "https://app.mywatersessions.com"
SPOTS = "https://spots.mywatersessions.com"

# Aufbau und Farben wie changelog.html, damit die Seiten nicht wie Fremdkörper
# wirken. Absichtlich ohne Google Fonts: die Landing nutzt Systemschriften, das
# spart einen Drittanbieter genau auf den Seiten, die Datenschutz erklären.
TEMPLATE = """<!doctype html>
<html lang="{lang}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="icon" href="/favicon.ico" sizes="any">
<link rel="icon" type="image/png" sizes="32x32" href="/favicon-32.png">
<link rel="apple-touch-icon" href="/apple-touch-icon-180.png">
<title>{title}</title>
<meta name="description" content="{desc}">
<link rel="canonical" href="{canon}">
{hreflang}
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
  .langbar{{ font-size:13px; margin-top:14px; }}
  .langbar a{{ margin-right:10px; text-decoration:none; color:#bfe7ee; }}
  .langbar .active{{ color:#fff; font-weight:700; }}
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
    <div class="langbar">{langbar}</div>
  </div>
</header>

<main>
  <div class="wrap">
{body}
    <p class="backline"><a href="{home}">{back}</a> &middot;
       <a href="{other_href}">{other_label}</a></p>
  </div>
</main>

<footer>
  <div class="wrap">
    <a href="{home}">{f_home}</a> ·
    <a href="{f_guide_href}">{f_guide}</a> ·
    <a href="{spots}/spots">Spots</a> ·
    <a href="{f_imp_href}">{f_imp}</a> ·
    <a href="{f_priv_href}">{f_priv}</a> ·
    <a href="{app}/">{f_app}</a>
  </div>
</footer>

</body>
</html>
"""

# Beschriftungen je Sprache – an einer Stelle, damit die Seiten nicht halb
# uebersetzt herauskommen.
LANG = {
    "de": {
        "home": "/index-de.html", "back": "← Zur Startseite",
        "f_home": "Start", "f_guide": "Anleitung", "f_guide_href": "/guide-de.html",
        "f_imp": "Impressum", "f_imp_href": "/impressum.html",
        "f_priv": "Datenschutz", "f_priv_href": "/datenschutz.html",
        "f_app": "App öffnen",
    },
    "en": {
        "home": "/", "back": "← Back to the start page",
        "f_home": "Home", "f_guide": "Guide", "f_guide_href": "/guide.html",
        "f_imp": "Imprint", "f_imp_href": "/imprint.html",
        "f_priv": "Privacy", "f_priv_href": "/privacy.html",
        "f_app": "Open the app",
    },
}

# (Datei, Sprache, Partnerdatei in der anderen Sprache)
PAGES = {
    "impressum.html": {
        "lang": "de", "twin": "imprint.html",
        "title": "Impressum – MyWaterSessions",
        "desc": ("Impressum von MyWaterSessions: Diensteanbieter, Kontakt und "
                 "Verantwortlicher nach § 5 DDG und § 18 Abs. 2 MStV."),
        "kicker": "Impressum",
        "body": impressum_html,
        "other_href": "/datenschutz.html", "other_label": "Datenschutzerklärung",
    },
    "datenschutz.html": {
        "lang": "de", "twin": "privacy.html",
        "title": "Datenschutzerklärung – MyWaterSessions",
        "desc": ("Datenschutzerklärung von MyWaterSessions: welche Daten wir "
                 "verarbeiten, was öffentlich sichtbar ist, welche Dienste "
                 "eingebunden sind und wie du deine Daten löschst."),
        "kicker": "Datenschutz",
        "body": datenschutz_html,
        "other_href": "/impressum.html", "other_label": "Impressum",
    },
    "imprint.html": {
        "lang": "en", "twin": "impressum.html",
        "title": "Imprint – MyWaterSessions",
        "desc": ("Imprint of MyWaterSessions: service provider, contact and the "
                 "person responsible for the content."),
        "kicker": "Imprint",
        "body": imprint_html_en,
        "other_href": "/privacy.html", "other_label": "Privacy policy",
    },
    "privacy.html": {
        "lang": "en", "twin": "datenschutz.html",
        "title": "Privacy policy – MyWaterSessions",
        "desc": ("Privacy policy of MyWaterSessions: what data we process, what is "
                 "publicly visible, which third-party services are involved and "
                 "how to delete your data."),
        "kicker": "Privacy",
        "body": privacy_html_en,
        "other_href": "/imprint.html", "other_label": "Imprint",
    },
}

# Sprachumschalter-Beschriftung je Zielsprache
LANG_NAME = {"de": "DE", "en": "EN"}


def build():
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "landing")
    if not os.path.isdir(out_dir):
        print(f"FEHLER: {out_dir} gibt es nicht", file=sys.stderr)
        return 1
    problems = 0
    for fname, cfg in PAGES.items():
        lang = cfg["lang"]
        twin = cfg["twin"]
        twin_lang = PAGES[twin]["lang"]
        # hreflang-Paar + x-default auf Englisch: die deutsche und die englische
        # Fassung sind derselbe Inhalt, sonst haelt Google sie fuer Dubletten.
        en_file = fname if lang == "en" else twin
        hreflang = "\n".join([
            f'<link rel="alternate" hreflang="{lang}" href="{BASE}/{fname}">',
            f'<link rel="alternate" hreflang="{twin_lang}" href="{BASE}/{twin}">',
            f'<link rel="alternate" hreflang="x-default" href="{BASE}/{en_file}">',
        ])
        langbar = " ".join(
            (f'<a class="active">{LANG_NAME[lang]}</a>' if lg == lang
             else f'<a href="/{twin}">{LANG_NAME[lg]}</a>')
            for lg in ("en", "de")
        )
        body = cfg["body"]()
        body = "\n".join(("    " + ln) if ln.strip() else ln
                         for ln in body.strip().splitlines())
        html = TEMPLATE.format(
            lang=lang, title=cfg["title"], desc=cfg["desc"], kicker=cfg["kicker"],
            canon=f"{BASE}/{fname}", hreflang=hreflang, langbar=langbar, body=body,
            app=APP, spots=SPOTS,
            other_href=cfg["other_href"], other_label=cfg["other_label"],
            **LANG[lang],
        )
        with open(os.path.join(out_dir, fname), "w", encoding="utf-8") as fh:
            fh.write(html)

        leftover = re.findall(r"\{[a-z_]+\}", html)
        h1 = html.count("<h1")
        po, pc = html.count("<p"), html.count("</p>")
        ok = not leftover and h1 == 1 and po == pc
        if not ok:
            problems += 1
        print(f"{'ok ' if ok else '!! '}{fname:<20} {len(html):>6} Zeichen  "
              f"h1={h1}  h2={html.count('<h2')}  p={po}/{pc}  "
              f"hreflang={html.count('hreflang=')}  "
              f"Platzhalter={leftover or 'keine'}")
    print(f"\nStand laut legal_texts.py: {LEGAL_STAND}")
    if problems:
        print(f"WARNUNG: {problems} Seite(n) mit Auffaelligkeiten")
        return 1
    print("Nicht vergessen: Landing neu deployen (Manual Deploy).")
    return 0


if __name__ == "__main__":
    sys.exit(build())
