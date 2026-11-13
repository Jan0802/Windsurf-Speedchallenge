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

START: Windsurf in DE + EN (die Suche, um die es geht, ist deutsch). Seit
13.11.2026 auch NL/FR/ES, mit Sprachleiste wie auf den Sportseiten (Jan:
"gibt es scheinbar nur auf deutsch" - die EN-Seite gab es, war aber von der
DE-Seite aus unsichtbar). Weitere Sportarten/Sprachen = ein weiterer Eintrag
in GUIDES, der Rest ist Vorlage. Die Dateinamen tragen den Suchbegriff der
jeweiligen Sprache; NL heisst wie DE "windsurfen", darum dort ein -nl-Suffix.

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

# Reihenfolge und Beschriftung der Sprachleiste (wie gen_sport_pages.py).
LANGS = ["en", "de", "nl", "fr", "es"]
LANG_NAME = {"en": "EN", "de": "DE", "nl": "NL", "fr": "FR", "es": "ES"}

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
    "Für belastbare Zahlen braucht es keinen Spezial-Logger mehr. <b>Garmin</b> und <b>Polar</b> "
    "verbinden sich per Konto, und jede Wassersport-Session kommt von selbst an. Auf einer Garmin "
    "rechnet unsere kostenlose Connect-IQ-App die Disziplinen zusätzlich direkt auf der Uhr, "
    "Sekunde für Sekunde. Und jede andere GPS-Uhr oder "
    "Handy-App — Apple Watch, Suunto, COROS, Strava — exportiert eine "
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
    "You no longer need a dedicated logger for solid numbers. <b>Garmin</b> and <b>Polar</b> "
    "connect via your account, and every water session arrives on its own. On a Garmin our "
    "free Connect IQ app also computes the disciplines on the watch, second by second. And "
    "every other GPS watch or phone app — Apple Watch, Suunto, COROS, Strava — exports a "
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
 ("windsurf", "nl"): {
  "file": "gps-tracking-windsurfen-nl.html",
  "html": "nl",
  "title": "GPS-tracking bij windsurfen: snelheid goed meten — gratis | MyWaterSessions",
  "meta": ("GPS-tracking bij windsurfen uitgelegd: waarom het 2-seconden-gemiddelde telt, wat 500 m "
           "en zeemijl betekenen, welk horloge geschikt is — en waarom twee apparaten verschillende "
           "cijfers laten zien. Van de makers van een gratis ranglijstplatform."),
  "h1": "GPS-tracking bij windsurfen: snelheid goed meten",
  "sub": ("Wat de cijfers betekenen, welk horloge geschikt is en waar metingen misgaan — "
          "uitgelegd door mensen die de meetfouten zelf hebben nagemeten."),
  "cta": "🌊 Gratis tracken & de ranglijst in",
  "landing": "windsurf-nl.html",
  "landing_text": "Meer over de gratis windsurf-ranglijst",
  "sections": [
   ("Waarom niemand de topsnelheid rangschikt",
    "Het ene getal dat je horloge groot laat zien — de maximale momentwaarde — is het "
    "onbetrouwbaarste van allemaal. GPS meet snelheid via de dopplerverschuiving van het "
    "satellietsignaal, en één meetwaarde kan omhoogschieten door afschaduwing, een draai van "
    "de pols of een spat water op de kast. Daarom rangschikken alle serieuze snelheidsklassementen "
    "<b>gemiddelden over een duur of een afstand</b>: een uitschieter van een halve seconde "
    "verdwijnt in het gemiddelde, maar de run moet nog steeds gevaren worden. De standaard is het "
    "<b>beste 2-seconden-gemiddelde</b> — kort genoeg om die ene perfecte run te belonen, lang "
    "genoeg om meetruis weg te filteren."),
   ("De disciplines: 2 s, 30 s, 500 m, zeemijl, alpha",
    "Achter de afkortingen zitten verschillende vragen aan dezelfde run: het <b>2 s-gemiddelde</b> "
    "meet het piekmoment, het <b>30 s-gemiddelde</b> of je het tempo kunt vasthouden. <b>500 m</b> "
    "en de <b>zeemijl</b> (1852 m) zijn de klassieke afstandsdisciplines van het speedsurfen — "
    "hier telt niet het moment, maar de complete run. <b>Alpha 500</b> vraagt 500 m met een gijp "
    "erin en meet daarmee wat een slalomsurfer echt kan. En <b>avg 5×10</b> (het gemiddelde van "
    "je vijf beste runs van 10 seconden, zonder overlap) beloont de constante sessie in plaats van "
    "die ene gelukstreffer."),
   ("Welk horloge is geschikt? Bijna elk.",
    "Voor betrouwbare cijfers heb je geen speciale logger meer nodig. <b>Garmin</b> en <b>Polar</b> "
    "koppel je via je account, en elke watersportsessie komt vanzelf binnen. Op een Garmin rekent "
    "onze gratis Connect IQ-app de disciplines bovendien direct op het horloge uit, seconde voor "
    "seconde. En elk ander gps-horloge of elke telefoon-app — Apple Watch, Suunto, COROS, Strava — "
    "exporteert een <b>FIT-, GPX- of TCX-bestand</b> waaruit alle waarden berekend kunnen worden. "
    "Belangrijker dan het merk is de <b>opnamefrequentie</b>: één punt per seconde is de basis voor "
    "schone 2 s-waarden. „Smart recording”, dat maar om de paar seconden een punt zet, gooit "
    "pieken weg."),
   ("Waarom twee apparaten verschillende cijfers laten zien",
    "Het gebruikelijke vermoeden — „mijn horloge meet verkeerd” — heeft meestal een onschuldiger "
    "verklaring: twee apparaten meten <b>niet hetzelfde</b>. De snelheid komt uit het dopplersignaal "
    "en is ook bij laag tempo heel betrouwbaar; de <b>afstand</b> daarentegen wordt soms uit de "
    "posities opgeteld en soms uit de snelheid geïntegreerd — en die twee lopen uiteen. We hebben "
    "het nagemeten: dezelfde sessie kwam afhankelijk van de rekenmethode op 4,78 of 5,38 km uit, "
    "elf procent verschil, beide apparaten „correct”. Daar komt de weergave bij: een uitgedund "
    "track (één punt per 10 s) snijdt bochten af en meet systematisch te kort. Wie cijfers "
    "vergelijkt, moet de rekenmethode vergelijken — of een platform gebruiken dat voor iedereen "
    "hetzelfde rekent."),
   ("Hoe herken je een onbetrouwbaar cijfer?",
    "Aan drie tegenstrijdigheden: de 2 s-waarde ligt <i>onder</i> de 30 s-waarde (onmogelijk — het "
    "kortere venster wint altijd). De gemiddelde snelheid van de sessie ligt boven de piek "
    "(onmogelijk). Of het track ligt helemaal niet op het water. Een eerlijke ranglijst controleert "
    "precies dat automatisch — bij ons heet dat de <b>trust-score</b>: onwaarschijnlijke sessies "
    "vallen uit de ranglijst, en wie het treft, ziet de reden. Zo blijft de ranglijst geloofwaardig "
    "zonder dat iemand met de hand censureert."),
  ],
  "faq_title": "Veelgestelde vragen",
  "faq": [
   ("Welke GPS-tracking-app voor windsurfen is gratis?",
    "MyWaterSessions is helemaal gratis — tracking, spotranglijsten, sessieanalyse en "
    "windverwachting, zonder premium-niveau en zonder paywall. De Garmin-horloge-app is ook gratis, "
    "en bestanden van elk ander horloge (FIT/GPX/TCX) kun je kosteloos uploaden."),
   ("Werkt GPS-tracking bij windsurfen ook zonder Garmin?",
    "Ja. Elk gps-horloge en elke telefoon-app die FIT, GPX of TCX exporteert werkt — Apple Watch, "
    "Suunto, COROS, Polar (ook automatisch via accountkoppeling) of een export uit Garmin Connect "
    "of Strava."),
   ("Waarom laat mijn horloge een andere afstand zien dan de analyse?",
    "Omdat afstand op twee manieren berekend wordt: uit de gps-posities of door de snelheid op te "
    "tellen. Die twee kunnen zo'n tien procent uit elkaar liggen — vooral bij langzaam varen. De "
    "posities zijn de betrouwbaarste methode."),
   ("Wat is een goede windsurfsnelheid?",
    "Planeren begint afhankelijk van het materiaal rond 15–20 km/u. Ambitieuze freeriders halen "
    "50–60 km/u als 2-seconden-gemiddelde, slalom- en speedsurfers aanzienlijk meer — het "
    "wereldrecord over 500 m ligt boven 98 km/u. Zinvoller dan elk absoluut getal is de ranglijst "
    "op je eigen spot bij vergelijkbare wind."),
   ("Is een telefoon genoeg in plaats van een horloge?",
    "Voor een logboek wel, voor snelheidsmetingen maar beperkt: een telefoon zit meestal in neopreen "
    "of een drybag, dempt het gps-signaal en neemt vaak maar om de paar seconden op. Een horloge om "
    "de pols dat elke seconde opneemt, levert de schonere pieken."),
  ],
  "foot": "snelheidsranglijst &amp; spotgids voor windsurf, kite, wing, SUP, wakeboard &amp; surf",
  "nav": ("Home", "index-nl.html", "Handleiding", "guide-nl.html", "Spots", "App openen"),
 },
 ("windsurf", "fr"): {
  "file": "suivi-gps-windsurf.html",
  "html": "fr",
  "title": "Suivi GPS en windsurf : bien mesurer sa vitesse — gratuit | MyWaterSessions",
  "meta": ("Le suivi GPS en windsurf expliqué : pourquoi la moyenne sur 2 secondes compte, ce que "
           "signifient 500 m et mille nautique, quelle montre convient — et pourquoi deux appareils "
           "affichent des chiffres différents. Par les créateurs d'une plateforme de classement gratuite."),
  "h1": "Suivi GPS en windsurf : bien mesurer sa vitesse",
  "sub": ("Ce que signifient les chiffres, quelle montre convient et où les mesures se trompent — "
          "expliqué par des gens qui ont mesuré les erreurs eux-mêmes."),
  "cta": "🌊 Tracker gratuitement & rejoindre le classement",
  "landing": "windsurf-fr.html",
  "landing_text": "En savoir plus sur le classement windsurf gratuit",
  "sections": [
   ("Pourquoi personne ne classe la vitesse max",
    "Le seul chiffre que votre montre affiche en grand — la valeur instantanée maximale — est le "
    "moins fiable de tous. Le GPS mesure la vitesse par l'effet Doppler du signal satellite, et une "
    "seule mesure peut s'envoler à cause d'un masquage, d'un mouvement du poignet ou d'une "
    "éclaboussure sur le boîtier. C'est pourquoi tous les classements de vitesse sérieux retiennent "
    "des <b>moyennes sur une durée ou une distance</b> : un pic d'une demi-seconde disparaît dans la "
    "moyenne, mais le run doit quand même être navigué. La référence est la <b>meilleure moyenne sur "
    "2 secondes</b> — assez courte pour récompenser le run parfait, assez longue pour filtrer le "
    "bruit de mesure."),
   ("Les disciplines : 2 s, 30 s, 500 m, mille nautique, alpha",
    "Derrière les abréviations se cachent des questions différentes posées au même run : la "
    "<b>moyenne 2 s</b> mesure le pic, la <b>moyenne 30 s</b> si vous tenez le rythme. <b>500 m</b> "
    "et le <b>mille nautique</b> (1852 m) sont les disciplines de distance classiques du speed "
    "windsurf — ici, c'est le run complet qui compte, pas l'instant. L'<b>Alpha 500</b> exige 500 m "
    "avec un jibe au milieu et mesure ainsi ce qu'un slalomeur sait vraiment faire. Et l'<b>avg "
    "5×10</b> (la moyenne de vos cinq meilleurs runs de 10 secondes, sans chevauchement) récompense "
    "la session régulière plutôt que le coup de chance."),
   ("Quelle montre convient ? Presque toutes.",
    "Plus besoin d'un enregistreur spécialisé pour obtenir des chiffres solides. <b>Garmin</b> et "
    "<b>Polar</b> se connectent via votre compte, et chaque session nautique arrive toute seule. Sur "
    "une Garmin, notre appli Connect IQ gratuite calcule en plus les disciplines directement sur la "
    "montre, seconde par seconde. Et toute autre montre GPS ou appli mobile — Apple Watch, Suunto, "
    "COROS, Strava — exporte un <b>fichier FIT, GPX ou TCX</b> à partir duquel toutes les valeurs se "
    "calculent. Plus important que la marque : la <b>fréquence d'enregistrement</b>. Un point par "
    "seconde est la base de valeurs 2 s propres. L'« enregistrement intelligent », qui ne pose un "
    "point que toutes les quelques secondes, jette les pics."),
   ("Pourquoi deux appareils affichent des chiffres différents",
    "Le soupçon habituel — « ma montre mesure mal » — a le plus souvent une explication plus "
    "innocente : deux appareils ne mesurent <b>pas la même chose</b>. La vitesse vient du signal "
    "Doppler et reste très fiable même à faible allure ; la <b>distance</b>, elle, est tantôt sommée "
    "à partir des positions, tantôt intégrée à partir de la vitesse — et les deux divergent. Nous "
    "l'avons mesuré : la même session donnait 4,78 ou 5,38 km selon la méthode, onze pour cent "
    "d'écart, les deux appareils « corrects ». S'y ajoute l'affichage : une trace allégée (un point "
    "toutes les 10 s) coupe les virages et mesure systématiquement trop court. Qui compare des "
    "chiffres doit comparer la méthode — ou utiliser une plateforme qui calcule de la même façon "
    "pour tout le monde."),
   ("Comment reconnaître un chiffre douteux ?",
    "À trois contradictions : la valeur 2 s est <i>inférieure</i> à la valeur 30 s (impossible — la "
    "fenêtre la plus courte gagne toujours). La vitesse moyenne de la session dépasse le pic "
    "(impossible). Ou la trace n'est pas du tout sur l'eau. Un classement équitable vérifie "
    "exactement cela automatiquement — chez nous, c'est le <b>score de confiance</b> : les sessions "
    "invraisemblables sortent du classement, et le rider concerné voit la raison. Le classement "
    "reste crédible sans que personne ne censure à la main."),
  ],
  "faq_title": "Questions fréquentes",
  "faq": [
   ("Quelle appli de suivi GPS pour le windsurf est gratuite ?",
    "MyWaterSessions est entièrement gratuite — suivi, classements par spot, analyse de session et "
    "prévisions de vent, sans niveau premium ni paywall. L'appli pour montre Garmin est gratuite "
    "elle aussi, et les fichiers de toute autre montre (FIT/GPX/TCX) se téléversent sans frais."),
   ("Le suivi GPS en windsurf fonctionne-t-il sans Garmin ?",
    "Oui. Toute montre GPS et toute appli mobile qui exporte du FIT, GPX ou TCX fonctionne — Apple "
    "Watch, Suunto, COROS, Polar (aussi automatiquement via la connexion du compte) ou un export "
    "depuis Garmin Connect ou Strava."),
   ("Pourquoi ma montre affiche-t-elle une distance différente de l'analyse ?",
    "Parce que la distance se calcule de deux façons : à partir des positions GPS, ou en sommant la "
    "vitesse. Les deux peuvent s'écarter d'environ dix pour cent — surtout à faible allure. Les "
    "positions sont la méthode la plus fiable."),
   ("Qu'est-ce qu'une bonne vitesse en windsurf ?",
    "Le planing commence vers 15–20 km/h selon le matériel. Les freeriders ambitieux atteignent "
    "50–60 km/h en moyenne sur 2 secondes, les slalomeurs et riders de speed nettement plus — le "
    "record du monde sur 500 m dépasse 98 km/h. Plus parlant que tout chiffre absolu : le classement "
    "sur votre propre spot, par vent comparable."),
   ("Un téléphone suffit-il à la place d'une montre ?",
    "Pour un carnet de bord, oui ; pour mesurer la vitesse, seulement en partie : le téléphone est "
    "généralement dans du néoprène ou une pochette étanche, atténue le signal GPS et n'enregistre "
    "souvent que toutes les quelques secondes. Une montre au poignet qui enregistre chaque seconde "
    "donne des pics plus propres."),
  ],
  "foot": "classement de vitesse &amp; guide des spots pour windsurf, kite, wing, SUP, wakeboard &amp; surf",
  "nav": ("Accueil", "index-fr.html", "Guide", "guide-fr.html", "Spots", "Ouvrir l'app"),
 },
 ("windsurf", "es"): {
  "file": "seguimiento-gps-windsurf.html",
  "html": "es",
  "title": "Seguimiento GPS en windsurf: cómo medir bien la velocidad — gratis | MyWaterSessions",
  "meta": ("El seguimiento GPS en windsurf explicado: por qué cuenta la media de 2 segundos, qué "
           "significan 500 m y milla náutica, qué reloj sirve — y por qué dos dispositivos muestran "
           "cifras distintas. De los creadores de una plataforma de ranking gratuita."),
  "h1": "Seguimiento GPS en windsurf: cómo medir bien la velocidad",
  "sub": ("Qué significan las cifras, qué reloj sirve y dónde fallan las mediciones — "
          "explicado por gente que midió los errores por sí misma."),
  "cta": "🌊 Registra gratis y entra en el ranking",
  "landing": "windsurf-es.html",
  "landing_text": "Más sobre el ranking de windsurf gratuito",
  "sections": [
   ("Por qué nadie puntúa la velocidad máxima",
    "La única cifra que tu reloj muestra en grande — el valor instantáneo máximo — es la menos "
    "fiable de todas. El GPS mide la velocidad por el efecto Doppler de la señal del satélite, y "
    "una sola lectura puede dispararse por una sombra, un giro de muñeca o una salpicadura en la "
    "caja. Por eso todos los rankings de velocidad serios puntúan <b>medias sobre una duración o "
    "una distancia</b>: un pico de medio segundo desaparece en la media, pero el tramo hay que "
    "navegarlo igualmente. El estándar es la <b>mejor media de 2 segundos</b> — lo bastante corta "
    "para premiar ese tramo perfecto, lo bastante larga para filtrar el ruido de medición."),
   ("Las disciplinas: 2 s, 30 s, 500 m, milla náutica, alpha",
    "Tras las abreviaturas hay preguntas distintas al mismo tramo: la <b>media de 2 s</b> mide el "
    "momento punta, la <b>media de 30 s</b> si puedes mantener el ritmo. <b>500 m</b> y la "
    "<b>milla náutica</b> (1852 m) son las disciplinas de distancia clásicas del speed windsurf — "
    "aquí no cuenta el instante, sino el tramo completo. <b>Alpha 500</b> exige 500 m con una "
    "trasluchada en medio y mide así lo que un regatista de slalom sabe hacer de verdad. Y "
    "<b>avg 5×10</b> (la media de tus cinco mejores tramos de 10 segundos, sin solaparse) premia la "
    "sesión constante en lugar del golpe de suerte."),
   ("¿Qué reloj sirve? Casi cualquiera.",
    "Para cifras sólidas ya no hace falta un registrador especial. <b>Garmin</b> y <b>Polar</b> se "
    "conectan con tu cuenta, y cada sesión en el agua llega sola. En un Garmin, nuestra app "
    "gratuita de Connect IQ calcula además las disciplinas directamente en el reloj, segundo a "
    "segundo. Y cualquier otro reloj GPS o app del móvil — Apple Watch, Suunto, COROS, Strava — "
    "exporta un <b>archivo FIT, GPX o TCX</b> del que se calculan todos los valores. Más importante "
    "que la marca es la <b>frecuencia de grabación</b>: un punto por segundo es la base para valores "
    "de 2 s limpios. La « grabación inteligente », que solo guarda un punto cada pocos segundos, "
    "tira los picos."),
   ("Por qué dos dispositivos muestran cifras distintas",
    "La sospecha habitual — « mi reloj mide mal » — suele tener una explicación más inocente: dos "
    "dispositivos <b>no miden lo mismo</b>. La velocidad sale de la señal Doppler y es muy fiable "
    "incluso a poca velocidad; la <b>distancia</b>, en cambio, a veces se suma a partir de las "
    "posiciones y a veces se integra a partir de la velocidad — y ambas se separan. Lo medimos: la "
    "misma sesión daba 4,78 o 5,38 km según el método, un once por ciento de diferencia, ambos "
    "dispositivos « correctos ». A eso se suma la presentación: un track aligerado (un punto cada "
    "10 s) recorta las curvas y mide sistemáticamente de menos. Quien compara cifras debe comparar "
    "el método — o usar una plataforma que calcule igual para todos."),
   ("¿Cómo se reconoce una cifra poco fiable?",
    "Por tres contradicciones: el valor de 2 s está <i>por debajo</i> del de 30 s (imposible — la "
    "ventana más corta siempre gana). La velocidad media de la sesión supera el pico (imposible). O "
    "el track ni siquiera está sobre el agua. Un ranking justo comprueba justo eso automáticamente — "
    "en nuestro caso se llama <b>puntuación de confianza</b>: lo inverosímil sale del ranking, y el "
    "afectado ve el motivo. Así el ranking sigue siendo creíble sin que nadie censure a mano."),
  ],
  "faq_title": "Preguntas frecuentes",
  "faq": [
   ("¿Qué app de seguimiento GPS para windsurf es gratis?",
    "MyWaterSessions es completamente gratis — seguimiento, rankings por spot, análisis de sesión y "
    "previsión de viento, sin nivel premium ni muro de pago. La app para reloj Garmin también es "
    "gratuita, y los archivos de cualquier otro reloj (FIT/GPX/TCX) se suben sin coste."),
   ("¿Funciona el seguimiento GPS en windsurf sin Garmin?",
    "Sí. Cualquier reloj GPS y cualquier app del móvil que exporte FIT, GPX o TCX funciona — Apple "
    "Watch, Suunto, COROS, Polar (también automáticamente al conectar la cuenta) o una exportación "
    "desde Garmin Connect o Strava."),
   ("¿Por qué mi reloj muestra una distancia distinta a la del análisis?",
    "Porque la distancia se calcula de dos maneras: a partir de las posiciones GPS o sumando la "
    "velocidad. Ambas pueden diferir en torno a un diez por ciento — sobre todo a poca velocidad. "
    "Las posiciones son el método más fiable."),
   ("¿Cuál es una buena velocidad en windsurf?",
    "El planeo empieza, según el material, en torno a 15–20 km/h. Los freeriders ambiciosos "
    "alcanzan 50–60 km/h de media en 2 segundos, los regatistas de slalom y speed bastante más — el "
    "récord mundial en 500 m supera los 98 km/h. Más significativo que cualquier cifra absoluta es "
    "el ranking en tu propio spot con un viento comparable."),
   ("¿Basta el móvil en lugar de un reloj?",
    "Para un registro sí, para medir velocidad solo a medias: el móvil suele ir en neopreno o en una "
    "bolsa estanca, atenúa la señal GPS y a menudo solo graba cada pocos segundos. Un reloj en la "
    "muñeca que graba cada segundo da los picos más limpios."),
  ],
  "foot": "ranking de velocidad y guía de spots para windsurf, kite, wing, SUP, wakeboard y surf",
  "nav": ("Inicio", "index-es.html", "Guía", "guide-es.html", "Spots", "Abrir la app"),
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
  .langbar{{ font-size:13px; margin:0 0 10px; }}
  .langbar a{{ margin:0 6px; text-decoration:none; color:#bfe7ee; }}
  .langbar .active{{ color:#fff; font-weight:700; }}
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
    <div class="langbar">{langbar}</div>
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
        # Sprachleiste: nur Sprachen, die es fuer diese Sportart wirklich gibt.
        langbar = "".join(
            (f'<a class="active">{LANG_NAME[lg]}</a>' if lg == lang
             else f'<a href="/{pair[lg]}">{LANG_NAME[lg]}</a>')
            for lg in LANGS if lg in pair
        )

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
            base=BASE, hreflang=hreflang, langbar=langbar, h1=g["h1"], sub=g["sub"],
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
