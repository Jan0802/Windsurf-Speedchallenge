#!/usr/bin/env python3
"""Generiert die SEO-Sportart-Landingpages (5 Sportarten x 5 Sprachen) in ./landing/.
Ein Build-Skript: hier den Inhalt pflegen, dann `python gen_sport_pages.py` laufen
lassen -> schreibt <slug>[-<lang>].html. Kopf/CSS/Footer sind Vorlage (einmalig)."""
import os

LANGS = ["en", "de", "nl", "fr", "es"]
LANG_NAME = {"en": "EN", "de": "DE", "nl": "NL", "fr": "FR", "es": "ES"}
BASE = "https://mywatersessions.com"
APP = "https://app.mywatersessions.com"
SPOTS = "https://spots.mywatersessions.com/spots"
BEACON = ("<script type='module' src='https://static.cloudflareinsights.com/beacon.min.js' "
          "data-cf-beacon='{\"token\": \"c16ea8dadd7749e1826d6ea48f80ca6d\"}'></script>")

SLUGS = ["windsurf", "kitesurf", "wingfoil", "sup", "wakeboard"]


def fname(slug, lang):
    return f"{slug}.html" if lang == "en" else f"{slug}-{lang}.html"


# --- Sprach-übergreifende Bausteine je Sprache -----------------------------
T = {
 "en": {
  "html": "en",
  "s_track": "What gets tracked", "s_how": "How it works — with any watch",
  "s_why": "Why riders use it", "cta": "🌊 Launch the app",
  "how": ("Record on a <b>Garmin</b> with our free Connect IQ app and the session uploads "
          "automatically. No Garmin? Just upload a <b>FIT / GPX / TCX</b> file from any other "
          "watch or phone app — Apple Watch, Suunto, COROS, or a Garmin Connect / Strava export. "
          "Speed, GPS track and every metric above are computed for you."),
  "why": ["<b>100% free</b> — no premium, no paywall on the ranking.",
          "<b>Per-spot rankings</b> and a personal <b>performance index</b>.",
          "<b>Spot guide + 3-day wind forecast</b> with a “worth it” rating — <a href='%SPOTS%'>browse spots</a>.",
          "<b>Groups &amp; community</b>, plus a Spot-TV live screen for your club or café."],
  "other": "Other sports:", "home": "Home", "guide": "Guide", "spotsw": "Spots",
  "chg": "Changelog", "openapp": "Open the app",
  "foot": "speed rankings &amp; spot guide for windsurf, kite, wing, SUP &amp; wakeboard",
 },
 "de": {
  "html": "de",
  "s_track": "Was getrackt wird", "s_how": "So funktioniert's — mit jeder Uhr",
  "s_why": "Warum Rider es nutzen", "cta": "🌊 App öffnen",
  "how": ("Mit einer <b>Garmin</b> und unserer kostenlosen Connect-IQ-App aufzeichnen – die "
          "Session lädt automatisch hoch. Keine Garmin? Einfach eine <b>FIT / GPX / TCX</b>-Datei "
          "von jeder anderen Uhr oder App hochladen – Apple Watch, Suunto, COROS oder ein Export "
          "aus Garmin Connect / Strava. Speed, GPS-Track und alle Werte oben werden automatisch berechnet."),
  "why": ["<b>100% kostenlos</b> – kein Premium, keine Paywall auf der Rangliste.",
          "<b>Ranglisten pro Spot</b> und ein persönlicher <b>Performance-Index</b>.",
          "<b>Spot-Guide + 3-Tage-Windvorhersage</b> mit „Lohnt-sich\"-Bewertung — <a href='%SPOTS%'>Spots ansehen</a>.",
          "<b>Gruppen &amp; Community</b>, plus ein Spot-TV-Live-Screen für Club oder Café."],
  "other": "Weitere Sportarten:", "home": "Start", "guide": "Anleitung", "spotsw": "Spots",
  "chg": "Changelog", "openapp": "App öffnen",
  "foot": "Speed-Bestenliste &amp; Spot-Guide für Windsurf, Kite, Wing, SUP &amp; Wakeboard",
 },
 "nl": {
  "html": "nl",
  "s_track": "Wat wordt gemeten", "s_how": "Zo werkt het — met elk horloge",
  "s_why": "Waarom riders het gebruiken", "cta": "🌊 App openen",
  "how": ("Neem op met een <b>Garmin</b> en onze gratis Connect IQ-app – de sessie wordt "
          "automatisch geüpload. Geen Garmin? Upload gewoon een <b>FIT / GPX / TCX</b>-bestand "
          "van elk ander horloge of elke app – Apple Watch, Suunto, COROS of een export uit "
          "Garmin Connect / Strava. Snelheid, GPS-track en alle waarden hierboven worden automatisch berekend."),
  "why": ["<b>100% gratis</b> – geen premium, geen paywall op de ranglijst.",
          "<b>Ranglijsten per spot</b> en een persoonlijke <b>performance-index</b>.",
          "<b>Spotgids + 3-daagse windverwachting</b> met een „de moeite waard\"-score — <a href='%SPOTS%'>bekijk spots</a>.",
          "<b>Groepen &amp; community</b>, plus een Spot-TV live-scherm voor je club of café."],
  "other": "Andere sporten:", "home": "Home", "guide": "Handleiding", "spotsw": "Spots",
  "chg": "Changelog", "openapp": "App openen",
  "foot": "snelheidsranglijst &amp; spotgids voor windsurf, kite, wing, SUP &amp; wakeboard",
 },
 "fr": {
  "html": "fr",
  "s_track": "Ce qui est mesuré", "s_how": "Comment ça marche — avec n'importe quelle montre",
  "s_why": "Pourquoi les riders l'utilisent", "cta": "🌊 Ouvrir l'app",
  "how": ("Enregistrez sur une <b>Garmin</b> avec notre appli Connect IQ gratuite – la session "
          "se téléverse automatiquement. Pas de Garmin ? Téléversez un fichier <b>FIT / GPX / TCX</b> "
          "depuis n'importe quelle montre ou appli – Apple Watch, Suunto, COROS ou un export "
          "Garmin Connect / Strava. Vitesse, trace GPS et toutes les mesures ci-dessus sont calculées pour vous."),
  "why": ["<b>100% gratuit</b> – pas de premium, pas de paywall sur le classement.",
          "<b>Classements par spot</b> et un <b>indice de performance</b> personnel.",
          "<b>Guide des spots + prévisions de vent 3 jours</b> avec une note « ça vaut le coup » — <a href='%SPOTS%'>voir les spots</a>.",
          "<b>Groupes &amp; communauté</b>, plus un écran Spot-TV en direct pour votre club ou café."],
  "other": "Autres sports :", "home": "Accueil", "guide": "Guide", "spotsw": "Spots",
  "chg": "Changelog", "openapp": "Ouvrir l'app",
  "foot": "classement de vitesse &amp; guide des spots pour windsurf, kite, wing, SUP &amp; wakeboard",
 },
 "es": {
  "html": "es",
  "s_track": "Qué se mide", "s_how": "Cómo funciona — con cualquier reloj",
  "s_why": "Por qué lo usan los riders", "cta": "🌊 Abrir la app",
  "how": ("Graba con un <b>Garmin</b> y nuestra app Connect IQ gratuita: la sesión se sube "
          "automáticamente. ¿Sin Garmin? Sube un archivo <b>FIT / GPX / TCX</b> desde cualquier "
          "otro reloj o app: Apple Watch, Suunto, COROS o una exportación de Garmin Connect / Strava. "
          "La velocidad, el track GPS y todas las métricas de arriba se calculan por ti."),
  "why": ["<b>100% gratis</b> — sin premium, sin muro de pago en el ranking.",
          "<b>Rankings por spot</b> y un <b>índice de rendimiento</b> personal.",
          "<b>Guía de spots + previsión de viento 3 días</b> con una valoración « vale la pena » — <a href='%SPOTS%'>ver spots</a>.",
          "<b>Grupos y comunidad</b>, además de una pantalla Spot-TV en vivo para tu club o café."],
  "other": "Otros deportes:", "home": "Inicio", "guide": "Guía", "spotsw": "Spots",
  "chg": "Changelog", "openapp": "Abrir la app",
  "foot": "ranking de velocidad y guía de spots para windsurf, kite, wing, SUP y wakeboard",
 },
}

# --- Metrik-Karten je Sprache (title, desc) --------------------------------
M = {
 "en": {
  "speed": ("⚡ Speed", "Your best <b>2 s</b> and <b>30 s</b> speeds — the classic benchmarks."),
  "dist": ("📏 Distance disciplines", "Fastest <b>500 m</b> and <b>nautical mile</b>."),
  "jump": ("🚀 Jumps &amp; airtime", "Jump <b>height</b> and <b>airtime</b> of every jump."),
  "cadence": ("🌊 Cadence &amp; strokes", "Strokes per minute and total paddle strokes."),
  "runs": ("🚩 Runs &amp; distance", "Longest run and total distance per session."),
  "trust": ("🛡️ Trust score", "A fair plausibility check on every GPS session."),
 },
 "de": {
  "speed": ("⚡ Speed", "Deine besten <b>2 s</b>- und <b>30 s</b>-Speeds — die Klassiker."),
  "dist": ("📏 Distanz-Disziplinen", "Schnellste <b>500 m</b> und <b>Seemeile</b>."),
  "jump": ("🚀 Sprünge &amp; Airtime", "<b>Sprunghöhe</b> und <b>Airtime</b> jedes Sprungs."),
  "cadence": ("🌊 Kadenz &amp; Schläge", "Schläge pro Minute und Gesamt-Paddelschläge."),
  "runs": ("🚩 Runs &amp; Distanz", "Längster Run und Gesamtstrecke pro Session."),
  "trust": ("🛡️ Trust-Score", "Faire Plausibilitätsprüfung jeder GPS-Session."),
 },
 "nl": {
  "speed": ("⚡ Snelheid", "Je beste <b>2 s</b>- en <b>30 s</b>-snelheden — de klassiekers."),
  "dist": ("📏 Afstand-disciplines", "Snelste <b>500 m</b> en <b>zeemijl</b>."),
  "jump": ("🚀 Sprongen &amp; airtime", "<b>Spronghoogte</b> en <b>airtime</b> van elke sprong."),
  "cadence": ("🌊 Cadans &amp; slagen", "Slagen per minuut en totaal aantal peddelslagen."),
  "runs": ("🚩 Runs &amp; afstand", "Langste run en totale afstand per sessie."),
  "trust": ("🛡️ Trust-score", "Eerlijke plausibiliteitscheck op elke GPS-sessie."),
 },
 "fr": {
  "speed": ("⚡ Vitesse", "Vos meilleures vitesses sur <b>2 s</b> et <b>30 s</b> — les classiques."),
  "dist": ("📏 Disciplines de distance", "Meilleur <b>500 m</b> et <b>mille nautique</b>."),
  "jump": ("🚀 Sauts &amp; airtime", "<b>Hauteur</b> et <b>airtime</b> de chaque saut."),
  "cadence": ("🌊 Cadence &amp; coups", "Coups par minute et total de coups de pagaie."),
  "runs": ("🚩 Runs &amp; distance", "Plus long run et distance totale par session."),
  "trust": ("🛡️ Score de confiance", "Contrôle de plausibilité équitable sur chaque session GPS."),
 },
 "es": {
  "speed": ("⚡ Velocidad", "Tus mejores velocidades en <b>2 s</b> y <b>30 s</b> — las clásicas."),
  "dist": ("📏 Disciplinas de distancia", "Mejor <b>500 m</b> y <b>milla náutica</b>."),
  "jump": ("🚀 Saltos y airtime", "<b>Altura</b> y <b>airtime</b> de cada salto."),
  "cadence": ("🌊 Cadencia y remadas", "Remadas por minuto y total de remadas."),
  "runs": ("🚩 Runs y distancia", "Run más largo y distancia total por sesión."),
  "trust": ("🛡️ Puntuación de confianza", "Comprobación de plausibilidad justa en cada sesión GPS."),
 },
}

# --- Sportart-Definition: Karten-Reihenfolge + Anzeigename je Sprache -------
SPORTS = {
 "windsurf":  {"cards": ["speed", "dist", "runs", "trust"],
               "name": {"en": "Windsurf", "de": "Windsurf", "nl": "Windsurf", "fr": "Windsurf", "es": "Windsurf"}},
 "kitesurf":  {"cards": ["jump", "speed", "runs", "trust"],
               "name": {"en": "Kitesurf", "de": "Kitesurf", "nl": "Kitesurf", "fr": "Kitesurf", "es": "Kitesurf"}},
 "wingfoil":  {"cards": ["speed", "dist", "runs", "trust"],
               "name": {"en": "Wingfoil", "de": "Wingfoil", "nl": "Wingfoil", "fr": "Wingfoil", "es": "Wingfoil"}},
 "sup":       {"cards": ["cadence", "dist", "speed", "trust"],
               "name": {"en": "SUP", "de": "SUP", "nl": "SUP", "fr": "SUP", "es": "SUP"}},
 "wakeboard": {"cards": ["jump", "speed", "runs", "trust"],
               "name": {"en": "Wakeboard", "de": "Wakeboard", "nl": "Wakeboard", "fr": "Wakeboard", "es": "Wakeboard"}},
}

# app ?sport= key je slug
SPORT_KEY = {"windsurf": "windsurf", "kitesurf": "kitesurf", "wingfoil": "wingsurf",
             "sup": "sup", "wakeboard": "wakeboard"}

# --- Unikater Kurztext je (slug, lang): title, meta, h1, sub, lead ----------
# H1/sub/lead sind bewusst sportspezifisch (kein Thin/Duplicate Content).
P = {
 "windsurf": {
  "en": ("Windsurf speed ranking & GPS tracking — free | MyWaterSessions",
         "Free windsurf speed ranking & GPS tracking: 2 s, 30 s, 500 m and nautical-mile speeds, per spot, with a trust score. Works with any watch (Garmin auto-upload or FIT/GPX/TCX).",
         "Windsurf speed ranking &amp; GPS tracking",
         "Track your windsurf sessions, see your top speeds and climb a fair community ranking — 100% free, for every level, with any watch.",
         "The free community <b>speed ranking &amp; spot guide</b> for windsurfing — part of a platform that also covers kite, wing, SUP and wakeboard. Your results land in the ranking automatically, per spot."),
  "de": ("Windsurf Rangliste & GPS-Tracking — kostenlos | MyWaterSessions",
         "Kostenlose Windsurf-Rangliste & GPS-Tracking: 2 s, 30 s, 500 m und Seemeile, pro Spot, mit Trust-Score. Mit jeder Uhr (Garmin-Auto-Upload oder FIT/GPX/TCX).",
         "Windsurf Rangliste &amp; GPS-Tracking",
         "Zeichne deine Windsurf-Sessions auf, sieh deine Top-Speeds und steig in einer fairen Community-Rangliste auf — 100% kostenlos, für jedes Level, mit jeder Uhr.",
         "Die kostenlose Community-<b>Speed-Bestenliste &amp; der Spot-Guide</b> fürs Windsurfen — Teil einer Plattform für Kite, Wing, SUP und Wakeboard. Deine Ergebnisse landen automatisch in der Rangliste, pro Spot."),
  "nl": ("Windsurf ranglijst & GPS-tracking — gratis | MyWaterSessions",
         "Gratis windsurf-ranglijst & GPS-tracking: 2 s, 30 s, 500 m en zeemijl, per spot, met trust-score. Werkt met elk horloge (Garmin auto-upload of FIT/GPX/TCX).",
         "Windsurf ranglijst &amp; GPS-tracking",
         "Registreer je windsurfsessies, zie je topsnelheden en klim in een eerlijke community-ranglijst — 100% gratis, voor elk niveau, met elk horloge.",
         "De gratis community-<b>snelheidsranglijst &amp; spotgids</b> voor windsurfen — onderdeel van een platform voor kite, wing, SUP en wakeboard. Je resultaten komen automatisch in de ranglijst, per spot."),
  "fr": ("Classement de vitesse windsurf & suivi GPS — gratuit | MyWaterSessions",
         "Classement de vitesse windsurf gratuit & suivi GPS : 2 s, 30 s, 500 m et mille nautique, par spot, avec score de confiance. Avec n'importe quelle montre (Garmin ou FIT/GPX/TCX).",
         "Classement de vitesse windsurf &amp; suivi GPS",
         "Enregistrez vos sessions de windsurf, voyez vos meilleures vitesses et grimpez dans un classement communautaire équitable — 100% gratuit, pour tous les niveaux, avec n'importe quelle montre.",
         "Le <b>classement de vitesse &amp; guide des spots</b> gratuit et communautaire pour le windsurf — au sein d'une plateforme qui couvre aussi kite, wing, SUP et wakeboard. Vos résultats arrivent automatiquement dans le classement, par spot."),
  "es": ("Ranking de velocidad de windsurf y seguimiento GPS — gratis | MyWaterSessions",
         "Ranking de velocidad de windsurf gratis y seguimiento GPS: 2 s, 30 s, 500 m y milla náutica, por spot, con puntuación de confianza. Con cualquier reloj (Garmin o FIT/GPX/TCX).",
         "Ranking de velocidad de windsurf y seguimiento GPS",
         "Registra tus sesiones de windsurf, mira tus mejores velocidades y sube en un ranking comunitario justo — 100% gratis, para todos los niveles, con cualquier reloj.",
         "El <b>ranking de velocidad y guía de spots</b> gratuito y comunitario para el windsurf — parte de una plataforma que también cubre kite, wing, SUP y wakeboard. Tus resultados entran automáticamente en el ranking, por spot."),
 },
 "kitesurf": {
  "en": ("Kitesurf jump tracker & speed leaderboard — free | MyWaterSessions",
         "Free kitesurf tracker & leaderboard: jump height, airtime and 2 s / 30 s speed, per spot, with a trust score. Works with any watch (Garmin auto-upload or FIT/GPX/TCX).",
         "Kitesurf jump tracker &amp; speed leaderboard",
         "Track your kite sessions — jump height, airtime and top speed — and climb a fair community leaderboard. 100% free, any watch.",
         "The free community <b>leaderboard &amp; spot guide</b> for kitesurfing — jumps, airtime and speed. Part of a platform that also covers windsurf, wing, SUP and wakeboard."),
  "de": ("Kitesurf Rangliste & Sprunghöhe-Tracking — kostenlos | MyWaterSessions",
         "Kostenlose Kitesurf-Rangliste & Tracking: Sprunghöhe, Airtime und 2 s / 30 s Speed, pro Spot, mit Trust-Score. Mit jeder Uhr (Garmin oder FIT/GPX/TCX).",
         "Kitesurf Rangliste &amp; Sprunghöhe-Tracking",
         "Tracke deine Kite-Sessions — Sprunghöhe, Airtime und Top-Speed — und steig in einer fairen Community-Rangliste auf. 100% kostenlos, mit jeder Uhr.",
         "Die kostenlose Community-<b>Rangliste &amp; der Spot-Guide</b> fürs Kitesurfen — Sprünge, Airtime und Speed. Teil einer Plattform für Windsurf, Wing, SUP und Wakeboard."),
  "nl": ("Kitesurf ranglijst & spronghoogte-tracking — gratis | MyWaterSessions",
         "Gratis kitesurf-ranglijst & tracking: spronghoogte, airtime en 2 s / 30 s snelheid, per spot, met trust-score. Werkt met elk horloge (Garmin of FIT/GPX/TCX).",
         "Kitesurf ranglijst &amp; spronghoogte-tracking",
         "Track je kitesessies — spronghoogte, airtime en topsnelheid — en klim in een eerlijke community-ranglijst. 100% gratis, met elk horloge.",
         "De gratis community-<b>ranglijst &amp; spotgids</b> voor kitesurfen — sprongen, airtime en snelheid. Onderdeel van een platform voor windsurf, wing, SUP en wakeboard."),
  "fr": ("Classement kitesurf & suivi de hauteur de saut — gratuit | MyWaterSessions",
         "Classement kitesurf gratuit & suivi : hauteur de saut, airtime et vitesse 2 s / 30 s, par spot, avec score de confiance. Avec n'importe quelle montre (Garmin ou FIT/GPX/TCX).",
         "Classement kitesurf &amp; suivi de hauteur de saut",
         "Suivez vos sessions de kite — hauteur de saut, airtime et vitesse max — et grimpez dans un classement communautaire équitable. 100% gratuit, avec n'importe quelle montre.",
         "Le <b>classement &amp; guide des spots</b> gratuit et communautaire pour le kitesurf — sauts, airtime et vitesse. Au sein d'une plateforme qui couvre aussi windsurf, wing, SUP et wakeboard."),
  "es": ("Ranking de kitesurf y seguimiento de altura de salto — gratis | MyWaterSessions",
         "Ranking de kitesurf gratis y seguimiento: altura de salto, airtime y velocidad 2 s / 30 s, por spot, con puntuación de confianza. Con cualquier reloj (Garmin o FIT/GPX/TCX).",
         "Ranking de kitesurf y seguimiento de altura de salto",
         "Registra tus sesiones de kite — altura de salto, airtime y velocidad máxima — y sube en un ranking comunitario justo. 100% gratis, con cualquier reloj.",
         "El <b>ranking y guía de spots</b> gratuito y comunitario para el kitesurf — saltos, airtime y velocidad. Parte de una plataforma que también cubre windsurf, wing, SUP y wakeboard."),
 },
 "wingfoil": {
  "en": ("Wingfoil ranking & GPS tracker — free | MyWaterSessions",
         "Free wingfoil (wing surfing) ranking & GPS tracker: 2 s / 30 s speed, 500 m, nautical mile and distance, per spot, with a trust score. Works with any watch.",
         "Wingfoil ranking &amp; GPS tracker",
         "Track your wingfoil / wing surfing sessions, see your top speeds and distance, and climb a fair community ranking — 100% free, any watch.",
         "The free community <b>ranking &amp; spot guide</b> for wing foiling — part of a platform that also covers windsurf, kite, SUP and wakeboard."),
  "de": ("Wingfoil Rangliste & GPS-Tracker — kostenlos | MyWaterSessions",
         "Kostenlose Wingfoil-/Wingsurf-Rangliste & GPS-Tracker: 2 s / 30 s Speed, 500 m, Seemeile und Distanz, pro Spot, mit Trust-Score. Mit jeder Uhr.",
         "Wingfoil Rangliste &amp; GPS-Tracker",
         "Tracke deine Wingfoil-/Wingsurf-Sessions, sieh Top-Speeds und Distanz und steig in einer fairen Community-Rangliste auf — 100% kostenlos, mit jeder Uhr.",
         "Die kostenlose Community-<b>Rangliste &amp; der Spot-Guide</b> fürs Wingfoilen — Teil einer Plattform für Windsurf, Kite, SUP und Wakeboard."),
  "nl": ("Wingfoil ranglijst & GPS-tracker — gratis | MyWaterSessions",
         "Gratis wingfoil-/wingsurf-ranglijst & GPS-tracker: 2 s / 30 s snelheid, 500 m, zeemijl en afstand, per spot, met trust-score. Werkt met elk horloge.",
         "Wingfoil ranglijst &amp; GPS-tracker",
         "Track je wingfoil-/wingsurfsessies, zie je topsnelheden en afstand en klim in een eerlijke community-ranglijst — 100% gratis, met elk horloge.",
         "De gratis community-<b>ranglijst &amp; spotgids</b> voor wingfoilen — onderdeel van een platform voor windsurf, kite, SUP en wakeboard."),
  "fr": ("Classement wingfoil & tracker GPS — gratuit | MyWaterSessions",
         "Classement wingfoil (wing) gratuit & tracker GPS : vitesse 2 s / 30 s, 500 m, mille nautique et distance, par spot, avec score de confiance. Avec n'importe quelle montre.",
         "Classement wingfoil &amp; tracker GPS",
         "Suivez vos sessions de wingfoil / wing, voyez vos meilleures vitesses et distances, et grimpez dans un classement communautaire équitable — 100% gratuit, avec n'importe quelle montre.",
         "Le <b>classement &amp; guide des spots</b> gratuit et communautaire pour le wingfoil — au sein d'une plateforme qui couvre aussi windsurf, kite, SUP et wakeboard."),
  "es": ("Ranking de wingfoil y tracker GPS — gratis | MyWaterSessions",
         "Ranking de wingfoil (wing) gratis y tracker GPS: velocidad 2 s / 30 s, 500 m, milla náutica y distancia, por spot, con puntuación de confianza. Con cualquier reloj.",
         "Ranking de wingfoil y tracker GPS",
         "Registra tus sesiones de wingfoil / wing, mira tus mejores velocidades y distancia, y sube en un ranking comunitario justo — 100% gratis, con cualquier reloj.",
         "El <b>ranking y guía de spots</b> gratuito y comunitario para el wingfoil — parte de una plataforma que también cubre windsurf, kite, SUP y wakeboard."),
 },
 "sup": {
  "en": ("SUP GPS tracker & speed ranking — free | MyWaterSessions",
         "Free stand-up paddle (SUP) GPS tracker & ranking: speed, distance, cadence and strokes, per spot, with a trust score. Works with any watch (Garmin or FIT/GPX/TCX).",
         "SUP GPS tracker &amp; speed ranking",
         "Track your stand-up paddle sessions — speed, distance, cadence and strokes — and compare in a fair community ranking. 100% free, any watch.",
         "The free community <b>ranking &amp; spot guide</b> for stand-up paddling — part of a platform that also covers windsurf, kite, wing and wakeboard."),
  "de": ("SUP GPS-Tracker & Geschwindigkeit — kostenlos | MyWaterSessions",
         "Kostenloser Stand-Up-Paddle-(SUP)-GPS-Tracker & Rangliste: Speed, Distanz, Kadenz und Schläge, pro Spot, mit Trust-Score. Mit jeder Uhr (Garmin oder FIT/GPX/TCX).",
         "SUP GPS-Tracker &amp; Geschwindigkeit",
         "Tracke deine Stand-Up-Paddle-Sessions — Speed, Distanz, Kadenz und Schläge — und vergleiche dich in einer fairen Community-Rangliste. 100% kostenlos, mit jeder Uhr.",
         "Die kostenlose Community-<b>Rangliste &amp; der Spot-Guide</b> fürs Stand-Up-Paddeln — Teil einer Plattform für Windsurf, Kite, Wing und Wakeboard."),
  "nl": ("SUP GPS-tracker & snelheidsranglijst — gratis | MyWaterSessions",
         "Gratis stand-up paddle (SUP) GPS-tracker & ranglijst: snelheid, afstand, cadans en slagen, per spot, met trust-score. Werkt met elk horloge (Garmin of FIT/GPX/TCX).",
         "SUP GPS-tracker &amp; snelheidsranglijst",
         "Track je stand-up-paddlesessies — snelheid, afstand, cadans en slagen — en vergelijk in een eerlijke community-ranglijst. 100% gratis, met elk horloge.",
         "De gratis community-<b>ranglijst &amp; spotgids</b> voor stand-up paddling — onderdeel van een platform voor windsurf, kite, wing en wakeboard."),
  "fr": ("Tracker GPS SUP & classement de vitesse — gratuit | MyWaterSessions",
         "Tracker GPS stand-up paddle (SUP) gratuit & classement : vitesse, distance, cadence et coups, par spot, avec score de confiance. Avec n'importe quelle montre.",
         "Tracker GPS SUP &amp; classement de vitesse",
         "Suivez vos sessions de stand-up paddle — vitesse, distance, cadence et coups — et comparez-vous dans un classement communautaire équitable. 100% gratuit, avec n'importe quelle montre.",
         "Le <b>classement &amp; guide des spots</b> gratuit et communautaire pour le stand-up paddle — au sein d'une plateforme qui couvre aussi windsurf, kite, wing et wakeboard."),
  "es": ("Tracker GPS de SUP y ranking de velocidad — gratis | MyWaterSessions",
         "Tracker GPS de stand-up paddle (SUP) gratis y ranking: velocidad, distancia, cadencia y remadas, por spot, con puntuación de confianza. Con cualquier reloj.",
         "Tracker GPS de SUP y ranking de velocidad",
         "Registra tus sesiones de stand-up paddle — velocidad, distancia, cadencia y remadas — y compárate en un ranking comunitario justo. 100% gratis, con cualquier reloj.",
         "El <b>ranking y guía de spots</b> gratuito y comunitario para el stand-up paddle — parte de una plataforma que también cubre windsurf, kite, wing y wakeboard."),
 },
 "wakeboard": {
  "en": ("Wakeboard tracking app — jumps & speed, free | MyWaterSessions",
         "Free wakeboard tracking: jumps, airtime, speed and distance per session, with a trust score. Works with any watch (Garmin auto-upload or FIT/GPX/TCX).",
         "Wakeboard tracking — jumps &amp; speed",
         "Track your wakeboard sessions — jumps, airtime and speed — and compare in a fair community ranking. 100% free, any watch (cable or boat).",
         "The free community <b>ranking &amp; session tracker</b> for wakeboarding — part of a platform that also covers windsurf, kite, wing and SUP."),
  "de": ("Wakeboard Tracking-App — Sprünge & Speed, kostenlos | MyWaterSessions",
         "Kostenloses Wakeboard-Tracking: Sprünge, Airtime, Speed und Distanz pro Session, mit Trust-Score. Mit jeder Uhr (Garmin-Auto-Upload oder FIT/GPX/TCX).",
         "Wakeboard Tracking — Sprünge &amp; Speed",
         "Tracke deine Wakeboard-Sessions — Sprünge, Airtime und Speed — und vergleiche dich in einer fairen Community-Rangliste. 100% kostenlos, mit jeder Uhr (Cable oder Boot).",
         "Die kostenlose Community-<b>Rangliste &amp; der Session-Tracker</b> fürs Wakeboarden — Teil einer Plattform für Windsurf, Kite, Wing und SUP."),
  "nl": ("Wakeboard tracking-app — sprongen & snelheid, gratis | MyWaterSessions",
         "Gratis wakeboard-tracking: sprongen, airtime, snelheid en afstand per sessie, met trust-score. Werkt met elk horloge (Garmin auto-upload of FIT/GPX/TCX).",
         "Wakeboard tracking — sprongen &amp; snelheid",
         "Track je wakeboardsessies — sprongen, airtime en snelheid — en vergelijk in een eerlijke community-ranglijst. 100% gratis, met elk horloge (kabel of boot).",
         "De gratis community-<b>ranglijst &amp; sessietracker</b> voor wakeboarden — onderdeel van een platform voor windsurf, kite, wing en SUP."),
  "fr": ("Appli de suivi wakeboard — sauts & vitesse, gratuit | MyWaterSessions",
         "Suivi wakeboard gratuit : sauts, airtime, vitesse et distance par session, avec score de confiance. Avec n'importe quelle montre (Garmin ou FIT/GPX/TCX).",
         "Suivi wakeboard — sauts &amp; vitesse",
         "Suivez vos sessions de wakeboard — sauts, airtime et vitesse — et comparez-vous dans un classement communautaire équitable. 100% gratuit, avec n'importe quelle montre (câble ou bateau).",
         "Le <b>classement &amp; tracker de sessions</b> gratuit et communautaire pour le wakeboard — au sein d'une plateforme qui couvre aussi windsurf, kite, wing et SUP."),
  "es": ("App de seguimiento de wakeboard — saltos y velocidad, gratis | MyWaterSessions",
         "Seguimiento de wakeboard gratis: saltos, airtime, velocidad y distancia por sesión, con puntuación de confianza. Con cualquier reloj (Garmin o FIT/GPX/TCX).",
         "Seguimiento de wakeboard — saltos y velocidad",
         "Registra tus sesiones de wakeboard — saltos, airtime y velocidad — y compárate en un ranking comunitario justo. 100% gratis, con cualquier reloj (cable o barco).",
         "El <b>ranking y tracker de sesiones</b> gratuito y comunitario para el wakeboard — parte de una plataforma que también cubre windsurf, kite, wing y SUP."),
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
<meta property="og:type" content="website">
<meta property="og:site_name" content="MyWaterSessions">
<meta property="og:title" content="{h1_plain}">
<meta property="og:description" content="{meta}">
<meta property="og:url" content="{canon}">
<meta property="og:image" content="{base}/og.png">
<meta name="twitter:card" content="summary_large_image">
<script type="application/ld+json">
{{
  "@context": "https://schema.org",
  "@type": "WebApplication",
  "name": "MyWaterSessions — {name}",
  "url": "{canon}",
  "applicationCategory": "SportsApplication",
  "operatingSystem": "Web, Garmin Connect IQ",
  "offers": {{ "@type": "Offer", "price": "0", "priceCurrency": "EUR" }},
  "description": "{meta}",
  "inLanguage": ["en", "de", "nl", "fr", "es"]
}}
</script>
<script type="application/ld+json">
{{
  "@context": "https://schema.org",
  "@type": "BreadcrumbList",
  "itemListElement": [
    {{ "@type": "ListItem", "position": 1, "name": "MyWaterSessions", "item": "{base}/" }},
    {{ "@type": "ListItem", "position": 2, "name": "{name}", "item": "{canon}" }}
  ]
}}
</script>
<style>
  :root{{ --aqua:#2bd4d9; --ink:#06303a; }}
  *{{ box-sizing:border-box; }}
  body{{ margin:0; font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif; color:#eaf4ff; background:#06222e; line-height:1.65; }}
  a{{ color:var(--aqua); }}
  .wrap{{ max-width:900px; margin:0 auto; padding:0 20px; }}
  .hero{{ background: linear-gradient(180deg, rgba(6,34,46,.55), rgba(6,34,46,.92)), radial-gradient(1200px 500px at 70% -10%, #0b6b8c, #06222e 60%); padding:60px 0 40px; text-align:center; }}
  .logo{{ font-size:clamp(24px,4vw,34px); font-weight:800; }}
  .logo .dot{{ color:var(--aqua); }}
  h1{{ font-size:clamp(26px,4.5vw,40px); margin:14px 0 6px; }}
  .sub{{ color:#bfe7ee; font-size:clamp(16px,2.2vw,19px); max-width:680px; margin:0 auto; }}
  .langbar{{ font-size:13px; margin:0 0 10px; }}
  .langbar a{{ margin:0 6px; text-decoration:none; color:#bfe7ee; }}
  .langbar .active{{ color:#fff; font-weight:700; }}
  .btn{{ display:inline-block; margin-top:20px; background:var(--aqua); color:var(--ink); font-weight:800; text-decoration:none; padding:14px 28px; border-radius:14px; font-size:17px; }}
  section{{ padding:34px 0; }}
  h2{{ font-size:clamp(20px,3.2vw,27px); margin:0 0 10px; }}
  p.lead{{ color:#cfe3ea; font-size:18px; }}
  .grid{{ display:grid; grid-template-columns:repeat(2,1fr); gap:16px; margin-top:8px; }}
  @media(max-width:680px){{ .grid{{ grid-template-columns:1fr; }} }}
  .card{{ background:rgba(255,255,255,.05); border:1px solid rgba(255,255,255,.12); border-radius:16px; padding:18px 20px; }}
  .card h3{{ margin:0 0 6px; font-size:18px; }}
  .card p{{ margin:0; color:#cfe3ea; font-size:15px; }}
  ul{{ color:#dcecf1; }} li{{ margin:6px 0; }}
  .sports-links{{ margin-top:10px; font-size:15px; }}
  .sports-links a{{ display:inline-block; margin:4px 10px 4px 0; }}
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
    <a class="btn" href="{app}/?sport={sportkey}">{cta}</a>
  </div>
</header>
<section>
  <div class="wrap">
    <p class="lead">{lead}</p>
    <h2>{s_track}</h2>
    <div class="grid">{cards}</div>
    <h2>{s_how}</h2>
    <p class="lead">{how}</p>
    <h2>{s_why}</h2>
    <ul>{why}</ul>
    <p style="margin-top:22px;"><a class="btn" href="{app}/?sport={sportkey}">{cta}</a></p>
    <div class="sports-links"><b>{other}</b> {otherlinks}</div>
  </div>
</section>
<footer>
  <div class="wrap">
    <div>© MyWaterSessions · {foot}</div>
    <div style="margin-top:8px;">
      <a href="/{homehref}">{home}</a> ·
      <a href="/{guidehref}">{guide}</a> ·
      <a href="{spots}">{spotsw}</a> ·
      <a href="/changelog.html">{chg}</a> ·
      <a href="{app}/">{openapp}</a>
    </div>
  </div>
</footer>
</body>
</html>
"""


def strip_tags(s):
    import re
    return re.sub(r"<[^>]+>", "", s)


def build():
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "landing")
    home = {"en": "", "de": "index-de.html", "nl": "index-nl.html",
            "fr": "index-fr.html", "es": "index-es.html"}
    guide = {"en": "guide.html", "de": "guide-de.html", "nl": "guide-nl.html",
             "fr": "guide-fr.html", "es": "guide-es.html"}
    count = 0
    for slug in SLUGS:
        for lang in LANGS:
            t = T[lang]
            title, meta, h1, sub, lead = P[slug][lang]
            canon = f"{BASE}/{fname(slug, lang)}"
            hreflang = "\n".join(
                f'<link rel="alternate" hreflang="{lg}" href="{BASE}/{fname(slug, lg)}">'
                for lg in LANGS
            ) + f'\n<link rel="alternate" hreflang="x-default" href="{BASE}/{fname(slug, "en")}">'
            langbar = "".join(
                (f'<a class="active">{LANG_NAME[lg]}</a>' if lg == lang
                 else f'<a href="/{fname(slug, lg)}">{LANG_NAME[lg]}</a>')
                for lg in LANGS
            )
            cards = "".join(
                f'<div class="card"><h3>{M[lang][c][0]}</h3><p>{M[lang][c][1]}</p></div>'
                for c in SPORTS[slug]["cards"]
            )
            why = "".join(f"<li>{w.replace('%SPOTS%', SPOTS)}</li>" for w in t["why"])
            others = [s for s in SLUGS if s != slug]
            otherlinks = " · ".join(
                f'<a href="/{fname(s, lang)}">{SPORTS[s]["name"][lang]}</a>' for s in others
            )
            html = TEMPLATE.format(
                html=t["html"], title=title, meta=meta, canon=canon, base=BASE,
                hreflang=hreflang, h1=h1, h1_plain=strip_tags(h1), sub=sub, lead=lead,
                name=SPORTS[slug]["name"][lang], langbar=langbar, cards=cards,
                s_track=t["s_track"], s_how=t["s_how"], s_why=t["s_why"],
                how=t["how"], why=why, cta=t["cta"], app=APP, spots=SPOTS,
                sportkey=SPORT_KEY[slug], other=t["other"], otherlinks=otherlinks,
                foot=t["foot"], homehref=home[lang], home=t["home"],
                guidehref=guide[lang], guide=t["guide"], spotsw=t["spotsw"],
                chg=t["chg"], openapp=t["openapp"], beacon=BEACON,
            )
            with open(os.path.join(out_dir, fname(slug, lang)), "w", encoding="utf-8") as fh:
                fh.write(html)
            count += 1
    # Sitemap-Fragment ausgeben (zum Einfügen).
    print("generated", count, "pages")
    print("--- sitemap urls ---")
    for slug in SLUGS:
        for lang in LANGS:
            print(f"  <url><loc>{BASE}/{fname(slug, lang)}</loc>"
                  f"<changefreq>monthly</changefreq><priority>0.7</priority></url>")


if __name__ == "__main__":
    build()
