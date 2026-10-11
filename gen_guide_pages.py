#!/usr/bin/env python3
"""Generiert Ratgeber-Seiten ("GPS-Tracking beim <Sport>") in ./landing/.

WARUM EIGENE SEITEN NEBEN DEN SPORT-LANDINGPAGES: Die Landingpages verkaufen
("Windsurf Rangliste & GPS-Tracking"), aber wer "GPS Tracking Windsurfen"
sucht, will oft erst einmal WISSEN, wie das geht - welche Kennzahl zaehlt,
welche Uhr taugt, warum zwei Geraete verschiedene Zahlen zeigen. Diese Seiten
beantworten genau das und verlinken dann auf App und Landingpage. Ratgeber-
Inhalt trifft die Suchintention solcher Anfragen besser als eine Produktseite
(Befund 11.10.2026: bei "GPS Tracking Windsurfen" ranken Magazin-Artikel und
Foren vorn, keine Produktseiten).

Der Inhalt kommt aus echter Messarbeit an diesem Projekt (2-s-Regel,
Doppler-Speed vs. Positionsrauschen, ausgeduennte Tracks, Streckenfehler der
Speed-Integration) - das unterscheidet die Seite von zusammengegoogeltem
SEO-Text, und genau deshalb darf sie das auch behaupten.

START: Windsurf in DE + EN (die Suche, um die es geht, ist deutsch). Weitere
Sportarten/Sprachen = ein weiterer Eintrag in GUIDES, der Rest ist Vorlage.

Aufruf: python gen_guide_pages.py  -> schreibt die HTML-Dateien und druckt
die Sitemap-Zeilen zum Einfuegen in landing/sitemap.xml.
"""
import datetime
import json
import os
import re

BASE = "https://mywatersessions.com"
APP = "https://app.mywatersessions.com"
SPOTS = "https://spots.mywatersessions.com/spots"
BEACON = ("<script type='module' src='https://static.cloudflareinsights.com/beacon.min.js' "
          "data-cf-beacon='{\"token\": \"c16ea8dadd7749e1826d6ea48f80ca6d\"}'></script>")

# je Seite: (dateiname, sprache, titel, meta, h1, untertitel,
#            [(ueberschrift, html-absatz), ...], [(frage, antwort), ...], landing-href)
GUIDES = {
 ("windsurf", "de"): {
  "file": "gps-tracking-windsurfen.html",
  "html": "de",
  "title": "GPS-Tracking beim Windsurfen: Speed richtig messen — kostenlos | MyWaterSessions",
  "meta": ("GPS-Tracking beim Windsurfen erklärt: warum der 2-Sekunden-Schnitt zählt, was 500 m "
           "und Seemeile bedeuten, welche Uhr taugt — und warum zwei Geräte verschiedene Zahlen "
           "zeigen. Von den Machern einer kostenlosen Ranglisten-Plattform."),
  "h1": "GPS-Tracking beim Windsurfen: Speed richtig messen",
  "sub": ("Was die Zahlen bedeuten, welche Uhr taugt und woran Messungen scheitern — "
          "erklärt von Leuten, die die Messfehler selbst ausgemessen haben."),
  "cta": "🌊 Kostenlos tracken & in die Rangliste",
  "landing": "windsurf-de.html",
  "landing_text": "Mehr über die kostenlose Windsurf-Rangliste",
  "sections": [
   ("Warum niemand den Topspeed wertet",
    "Die eine Zahl, die deine Uhr groß anzeigt — der maximale Momentanwert — ist die "
    "unzuverlässigste von allen. GPS misst Geschwindigkeit über die Dopplerverschiebung des "
    "Satellitensignals, und ein einzelner Messwert kann durch Abschattung, einen Schwenk des "
    "Handgelenks oder einen Spritzer aufs Gehäuse nach oben ausreißen. Deshalb werten alle "
    "ernsthaften Speed-Wertungen <b>Mittelwerte über eine Dauer oder eine Strecke</b>: Ein "
    "Ausreißer von einer halben Sekunde verschwindet im Schnitt, gefahren werden muss trotzdem. "
    "Der Standard ist der <b>beste 2-Sekunden-Schnitt</b> — kurz genug, um den einen perfekten "
    "Lauf zu belohnen, lang genug, um Messfehler auszusieben."),
   ("Die Disziplinen: 2 s, 30 s, 500 m, Seemeile, Alpha",
    "Hinter den Kürzeln stecken verschiedene Fragen an denselben Lauf: Der <b>2-s-Schnitt</b> "
    "misst den Spitzenmoment, der <b>30-s-Schnitt</b>, ob du das Tempo halten kannst. "
    "<b>500 m</b> und die <b>Seemeile</b> (1852 m) sind die klassischen Streckendisziplinen des "
    "Speedsurfens — hier zählt nicht der Moment, sondern der komplette Run. <b>Alpha 500</b> "
    "verlangt 500 m mit einer Halse darin und misst damit, was ein Slalomfahrer wirklich kann. "
    "Und <b>avg 5×10</b> (der Schnitt deiner fünf besten 10-Sekunden-Läufe, ohne Überlappung) "
    "belohnt die konstante Session statt des einen Glückstreffers."),
   ("Welche Uhr taugt dafür? Fast jede.",
    "Für belastbare Zahlen braucht es keinen Spezial-Logger mehr. Eine <b>Garmin</b> mit unserer "
    "kostenlosen Connect-IQ-App rechnet die Disziplinen direkt auf der Uhr und lädt die Session "
    "automatisch hoch. Eine <b>Polar</b> verbindet sich per Konto. Und jede andere GPS-Uhr oder "
    "Handy-App — Apple Watch, Suunto, COROS, Garmin Connect, Strava — exportiert eine "
    "<b>FIT-, GPX- oder TCX-Datei</b>, aus der sich alle Werte berechnen lassen. Wichtiger als "
    "die Marke ist die <b>Aufzeichnungsrate</b>: einmal pro Sekunde ist die Basis für saubere "
    "2-s-Werte. „Smart Recording“, das nur alle paar Sekunden einen Punkt setzt, verschenkt "
    "Spitzenwerte."),
   ("Warum zwei Geräte verschiedene Zahlen zeigen",
    "Der häufigste Verdacht — „meine Uhr misst falsch“ — hat meist eine harmlosere Erklärung: "
    "Zwei Geräte messen <b>nicht dasselbe</b>. Die Geschwindigkeit kommt aus dem Doppler-Signal "
    "und ist bei niedrigem Tempo sehr zuverlässig; die <b>Strecke</b> dagegen wird mal aus den "
    "Positionen summiert, mal aus der Geschwindigkeit integriert — und beides driftet "
    "auseinander. Wir haben es nachgemessen: Dieselbe Session kam je nach Rechenweg auf 4,78 "
    "oder 5,38 km, elf Prozent Unterschied, beide Geräte „richtig“. Dazu kommt die Darstellung: "
    "Ein ausgedünnter Track (ein Punkt alle 10 s) schneidet Kurven ab und misst systematisch zu "
    "kurz. Wer Zahlen vergleicht, muss den Rechenweg vergleichen — oder eine Plattform nutzen, "
    "die für alle gleich rechnet."),
   ("Woran erkennt man eine unseriöse Zahl?",
    "An drei Widersprüchen: Der 2-s-Wert liegt <i>unter</i> dem 30-s-Wert (unmöglich, das "
    "kürzere Fenster gewinnt immer). Die Durchschnittsgeschwindigkeit der Session übersteigt "
    "den Spitzenwert (unmöglich). Oder der Track liegt gar nicht auf dem Wasser. Eine faire "
    "Rangliste prüft genau das automatisch — bei uns heißt das <b>Trust-Score</b>: Unplausibles "
    "fliegt aus der Wertung, und wer betroffen ist, sieht den Grund. So bleibt die Bestenliste "
    "glaubwürdig, ohne dass jemand von Hand zensiert."),
  ],
  "faq_title": "Häufige Fragen",
  "faq": [
   ("Welche GPS-Tracking-App für Windsurfen ist kostenlos?",
    "MyWaterSessions ist komplett kostenlos — Tracking, Spot-Ranglisten, Session-Analyse und "
    "Windvorhersage, ohne Premium-Stufe und ohne Paywall. Die Garmin-Uhren-App ist ebenfalls "
    "gratis, und Dateien von jeder anderen Uhr (FIT/GPX/TCX) lassen sich kostenlos hochladen."),
   ("Geht GPS-Tracking beim Windsurfen auch ohne Garmin?",
    "Ja. Jede GPS-Uhr und jede Handy-App, die FIT, GPX oder TCX exportiert, funktioniert — "
    "Apple Watch, Suunto, COROS, Polar (auch automatisch per Konto-Verbindung) oder ein Export "
    "aus Garmin Connect bzw. Strava."),
   ("Warum zeigt meine Uhr eine andere Strecke als die Auswertung?",
    "Weil Strecke auf zwei Wegen berechnet wird: aus den GPS-Positionen oder durch Aufsummieren "
    "der Geschwindigkeit. Die Wege können um rund zehn Prozent auseinanderliegen — vor allem "
    "bei langsamer Fahrt. Die Positionen sind der verlässlichere Weg."),
   ("Was ist ein guter Windsurf-Speed?",
    "Gleiten beginnt je nach Material um 15–20 km/h. Ambitionierte Freerider erreichen "
    "50–60 km/h im 2-Sekunden-Schnitt, Slalom- und Speedfahrer deutlich mehr — der Weltrekord "
    "über 500 m liegt bei über 98 km/h. Aussagekräftiger als jede absolute Zahl ist die "
    "Rangliste am eigenen Spot bei vergleichbarem Wind."),
   ("Reicht das Handy statt einer Uhr?",
    "Fürs Protokoll ja, fürs Speed-Messen nur bedingt: Das Handy steckt meist in Neopren oder "
    "Trockentasche, dämpft das GPS-Signal und zeichnet oft nur alle paar Sekunden auf. Eine Uhr "
    "am Handgelenk mit 1-Sekunden-Aufzeichnung liefert die saubereren Spitzenwerte."),
  ],
  "foot": "Speed-Bestenliste &amp; Spot-Guide für Windsurf, Kite, Wing, SUP, Wakeboard &amp; Surf",
  "nav": ("Start", "index-de.html", "Anleitung", "guide-de.html", "Spots", "App öffnen"),
 },
 ("windsurf", "en"): {
  "file": "gps-tracking-windsurfing.html",
  "html": "en",
  "title": "GPS tracking for windsurfing: how to measure speed properly — free | MyWaterSessions",
  "meta": ("GPS tracking for windsurfing explained: why the 2-second average counts, what 500 m "
           "and the nautical mile mean, which watch works — and why two devices show different "
           "numbers. By the makers of a free ranking platform."),
  "h1": "GPS tracking for windsurfing: measuring speed properly",
  "sub": ("What the numbers mean, which watch works and where measurements go wrong — "
          "explained by people who measured the errors themselves."),
  "cta": "🌊 Track for free & join the ranking",
  "landing": "windsurf.html",
  "landing_text": "More about the free windsurf ranking",
  "sections": [
   ("Why nobody ranks the top speed",
    "The one number your watch shows big — the maximum instantaneous value — is the least "
    "reliable of all. GPS measures speed via the Doppler shift of the satellite signal, and a "
    "single reading can spike from shading, a flick of the wrist or spray on the case. That is "
    "why every serious speed discipline ranks <b>averages over a duration or a distance</b>: a "
    "half-second outlier disappears in the average, but the run still has to be ridden. The "
    "standard is the <b>best 2-second average</b> — short enough to reward one perfect run, "
    "long enough to filter measurement noise."),
   ("The disciplines: 2 s, 30 s, 500 m, nautical mile, alpha",
    "The abbreviations ask different questions of the same run: the <b>2 s average</b> measures "
    "the peak moment, the <b>30 s average</b> whether you can hold the pace. <b>500 m</b> and "
    "the <b>nautical mile</b> (1852 m) are the classic distance disciplines of speedsurfing — "
    "here the whole run counts, not the moment. <b>Alpha 500</b> demands 500 m with a gybe in "
    "the middle, measuring what a slalom sailor can actually do. And <b>avg 5×10</b> (the "
    "average of your five best 10-second runs, non-overlapping) rewards the consistent session "
    "over the one lucky hit."),
   ("Which watch is good enough? Almost any.",
    "You no longer need a dedicated logger for solid numbers. A <b>Garmin</b> with our free "
    "Connect IQ app computes the disciplines on the watch and uploads the session "
    "automatically. A <b>Polar</b> connects via your account. And every other GPS watch or "
    "phone app — Apple Watch, Suunto, COROS, Garmin Connect, Strava — exports a "
    "<b>FIT, GPX or TCX file</b> from which all values can be computed. More important than "
    "the brand is the <b>recording rate</b>: one point per second is the basis for clean 2 s "
    "values. “Smart recording”, which stores a point only every few seconds, throws peaks away."),
   ("Why two devices show different numbers",
    "The usual suspicion — “my watch is wrong” — mostly has a more innocent explanation: two "
    "devices do <b>not measure the same thing</b>. Speed comes from the Doppler signal and is "
    "very reliable even at low pace; <b>distance</b>, however, is sometimes summed from "
    "positions and sometimes integrated from speed — and the two drift apart. We measured it: "
    "the same session came out at 4.78 or 5.38 km depending on the method, eleven percent "
    "apart, both devices “correct”. Add presentation: a thinned track (one point every 10 s) "
    "cuts corners and systematically under-measures. If you compare numbers, compare the "
    "method — or use a platform that computes them the same way for everyone."),
   ("How do you spot a bogus number?",
    "Three contradictions give it away: a 2 s value <i>below</i> the 30 s value (impossible — "
    "the shorter window always wins). A session average above the peak (impossible). Or a "
    "track that never touches water. A fair ranking checks exactly this automatically — we "
    "call it the <b>trust score</b>: implausible sessions drop out of the ranking, and the "
    "rider affected sees the reason. The leaderboard stays credible without anyone censoring "
    "by hand."),
  ],
  "faq_title": "Frequently asked questions",
  "faq": [
   ("Which GPS tracking app for windsurfing is free?",
    "MyWaterSessions is completely free — tracking, per-spot rankings, session analysis and "
    "wind forecast, with no premium tier and no paywall. The Garmin watch app is free as well, "
    "and files from any other watch (FIT/GPX/TCX) can be uploaded at no cost."),
   ("Does GPS tracking for windsurfing work without a Garmin?",
    "Yes. Any GPS watch or phone app that exports FIT, GPX or TCX works — Apple Watch, Suunto, "
    "COROS, Polar (also automatically via account link) or an export from Garmin Connect or "
    "Strava."),
   ("Why does my watch show a different distance than the analysis?",
    "Because distance is computed two ways: from the GPS positions, or by summing up speed. "
    "The two can differ by around ten percent — especially at slow pace. The positions are the "
    "more reliable method."),
   ("What is a good windsurfing speed?",
    "Planing starts around 15–20 km/h depending on gear. Ambitious freeriders reach 50–60 km/h "
    "as a 2-second average, slalom and speed sailors considerably more — the 500 m world "
    "record is above 98 km/h. More meaningful than any absolute number is the ranking at your "
    "own spot in comparable wind."),
   ("Is a phone enough instead of a watch?",
    "For a logbook, yes; for speed measurement only partly: a phone usually sits in neoprene "
    "or a dry bag, damps the GPS signal and often records only every few seconds. A watch on "
    "the wrist recording once per second delivers the cleaner peaks."),
  ],
  "foot": "speed rankings &amp; spot guide for windsurf, kite, wing, SUP, wakeboard &amp; surf",
  "nav": ("Home", "", "Guide", "guide.html", "Spots", "Open the app"),
 },
}

TEMPLATE = """<!doctype html>
<html lang="{html}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="icon" href="/favicon.ico" sizes="any">
<link rel="icon" type="image/png" sizes="32x32" href="/favicon-32.png">
<link rel="apple-touch-icon" href="/apple-touch-icon-180.png">
<link rel="manifest" href="/site.webmanifest">
<title>{title}</title>
<meta name="description" content="{meta}">
<link rel="canonical" href="{canon}">
{hreflang}
<meta name="theme-color" content="#06303a">
<meta property="og:type" content="article">
<meta property="og:site_name" content="MyWaterSessions">
<meta property="og:title" content="{h1}">
<meta property="og:description" content="{meta}">
<meta property="og:url" content="{canon}">
<meta property="og:image" content="{base}/og.png">
<meta name="twitter:card" content="summary_large_image">
<script type="application/ld+json">
{article_ld}
</script>
<script type="application/ld+json">
{faq_ld}
</script>
<style>
  :root{{ --aqua:#2bd4d9; --ink:#06303a; }}
  *{{ box-sizing:border-box; }}
  body{{ margin:0; font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif; color:#eaf4ff; background:#06222e; line-height:1.7; }}
  a{{ color:var(--aqua); }}
  .wrap{{ max-width:780px; margin:0 auto; padding:0 20px; }}
  .hero{{ background: linear-gradient(180deg, rgba(6,34,46,.55), rgba(6,34,46,.92)), radial-gradient(1200px 500px at 70% -10%, #0b6b8c, #06222e 60%); padding:54px 0 36px; text-align:center; }}
  .logo{{ font-size:clamp(22px,4vw,30px); font-weight:800; }}
  .logo .dot{{ color:var(--aqua); }}
  h1{{ font-size:clamp(25px,4.4vw,38px); margin:18px 0 8px; }}
  .sub{{ color:#bfe7ee; font-size:clamp(16px,2.2vw,18px); max-width:640px; margin:0 auto; }}
  .btn{{ display:inline-block; margin-top:20px; background:var(--aqua); color:var(--ink); font-weight:800; text-decoration:none; padding:13px 26px; border-radius:14px; font-size:16px; }}
  article{{ padding:30px 0 6px; }}
  h2{{ font-size:clamp(20px,3vw,25px); margin:30px 0 8px; }}
  article p{{ color:#dcecf1; font-size:16.5px; margin:0 0 10px; }}
  .faq{{ margin:26px 0 8px; }}
  .faq details{{ background:rgba(255,255,255,.05); border:1px solid rgba(255,255,255,.12); border-radius:14px; padding:12px 18px; margin:10px 0; }}
  .faq summary{{ font-weight:700; cursor:pointer; font-size:16.5px; }}
  .faq p{{ margin:10px 0 2px; }}
  .more{{ margin:26px 0 10px; font-size:16px; }}
  footer{{ padding:30px 0 48px; color:#7fa6b2; font-size:13px; text-align:center; }}
  footer a{{ color:#9fd9e3; margin:0 8px; }}
</style>
{beacon}
</head>
<body>
<header class="hero">
  <div class="wrap">
    <div class="logo">MyWaterSessions<span class="dot">.</span></div>
    <h1>{h1}</h1>
    <p class="sub">{sub}</p>
    <a class="btn" href="{app}/?sport=windsurf">{cta}</a>
  </div>
</header>
<article>
  <div class="wrap">
    {sections}
    <div class="faq">
      <h2>{faq_title}</h2>
      {faq}
    </div>
    <p class="more">→ <a href="/{landing}">{landing_text}</a> · <a href="{spots}">{spotsname}</a></p>
    <p><a class="btn" href="{app}/?sport=windsurf">{cta}</a></p>
  </div>
</article>
<footer>
  <div class="wrap">
    <div>© MyWaterSessions · {foot}</div>
    <div style="margin-top:8px;">
      <a href="/{nav1h}">{nav1}</a> ·
      <a href="/{nav2h}">{nav2}</a> ·
      <a href="{spots}">{nav3}</a> ·
      <a href="{app}/">{nav4}</a>
    </div>
  </div>
</footer>
</body>
</html>
"""


def strip_tags(s):
    return re.sub(r"<[^>]+>", "", s)


def build():
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "landing")
    today = datetime.date.today().isoformat()
    # hreflang-Paare je Sportart
    by_slug = {}
    for (slug, lang), g in GUIDES.items():
        by_slug.setdefault(slug, {})[lang] = g["file"]

    for (slug, lang), g in GUIDES.items():
        canon = f"{BASE}/{g['file']}"
        pair = by_slug[slug]
        hreflang = "\n".join(
            f'<link rel="alternate" hreflang="{lg}" href="{BASE}/{f}">'
            for lg, f in sorted(pair.items())
        )
        if "en" in pair:
            hreflang += f'\n<link rel="alternate" hreflang="x-default" href="{BASE}/{pair["en"]}">'

        article_ld = json.dumps({
            "@context": "https://schema.org",
            "@type": "Article",
            "headline": strip_tags(g["h1"]),
            "description": g["meta"],
            "inLanguage": g["html"],
            "mainEntityOfPage": canon,
            "image": f"{BASE}/og.png",
            "dateModified": today,
            "author": {"@type": "Organization", "name": "MyWaterSessions",
                       "url": BASE},
            "publisher": {"@type": "Organization", "name": "MyWaterSessions",
                          "url": BASE},
        }, ensure_ascii=False, indent=2)
        faq_ld = json.dumps({
            "@context": "https://schema.org",
            "@type": "FAQPage",
            "mainEntity": [
                {"@type": "Question", "name": strip_tags(q),
                 "acceptedAnswer": {"@type": "Answer", "text": strip_tags(a)}}
                for q, a in g["faq"]
            ],
        }, ensure_ascii=False, indent=2)

        sections = "".join(f"<h2>{h}</h2>\n<p>{p}</p>\n" for h, p in g["sections"])
        faq = "".join(
            f"<details><summary>{q}</summary><p>{a}</p></details>\n"
            for q, a in g["faq"]
        )
        nav1, nav1h, nav2, nav2h, nav3, nav4 = g["nav"]
        html = TEMPLATE.format(
            html=g["html"], title=g["title"], meta=g["meta"], canon=canon,
            base=BASE, hreflang=hreflang, h1=g["h1"], sub=g["sub"],
            cta=g["cta"], app=APP, spots=SPOTS, spotsname=nav3,
            article_ld=article_ld, faq_ld=faq_ld, sections=sections,
            faq=faq, faq_title=g["faq_title"], landing=g["landing"],
            landing_text=g["landing_text"], foot=g["foot"],
            nav1=nav1, nav1h=nav1h, nav2=nav2, nav2h=nav2h, nav3=nav3, nav4=nav4,
            beacon=BEACON,
        )
        with open(os.path.join(out_dir, g["file"]), "w", encoding="utf-8") as fh:
            fh.write(html)
        print("wrote", g["file"])

    print("--- sitemap urls ---")
    for (slug, lang), g in GUIDES.items():
        print(f"  <url><loc>{BASE}/{g['file']}</loc>"
              f"<lastmod>{today}</lastmod>"
              f"<changefreq>monthly</changefreq><priority>0.7</priority></url>")


if __name__ == "__main__":
    build()
