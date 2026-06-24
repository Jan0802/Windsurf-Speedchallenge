import base64
import glob
import hashlib
import json
import logging
import os
import re
import secrets
import tempfile
import time
from datetime import datetime, timedelta
from html import unescape
from urllib.error import HTTPError
from urllib.parse import urlencode, urljoin
from urllib.request import Request, urlopen
import io

import numpy as np
import pandas as pd
import pydeck as pdk
import streamlit as st
import streamlit.components.v1 as components
from fitparse import FitFile
from sqlalchemy import (
    Boolean,
    Column,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    LargeBinary,
    MetaData,
    String,
    Table,
    UniqueConstraint,
    create_engine,
    delete,
    func,
    insert,
    inspect,
    select,
    text,
    update,
)

# QR-Code fuer das Spot-TV ("Join today's ranking"). Optionales Paket – fehlt
# es, zeigt das Dashboard stattdessen den Beitritts-Link als Text.
try:
    import qrcode

    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False

# MTP-Zugriff (Garmin-Uhren ohne Laufwerksbuchstaben, z.B. Fenix 6 Pro).
# Nur unter Windows mit pywin32 verfügbar – auf einem Server fällt das
# sauber weg, ohne dass die App abstürzt.
try:
    import pythoncom
    import win32com.client

    MTP_AVAILABLE = True
except ImportError:
    MTP_AVAILABLE = False

# Persistenter Login per Cookie ("Angemeldet bleiben"). Optionales Paket –
# fehlt es, läuft die App normal weiter (nur ohne Cookie-Login).
try:
    import extra_streamlit_components as stx

    COOKIES_AVAILABLE = True
except ImportError:
    COOKIES_AVAILABLE = False

# Uhr-Auslesen (USB/MTP) ist nur lokal unter Windows sinnvoll.
IS_WINDOWS = os.name == "nt"


# Basisverzeichnis der App – alle Pfade werden relativ zur Skript-Datei
# aufgelöst, damit die App unabhängig vom Arbeitsverzeichnis läuft
# (wichtig fürs Deployment auf einem Server).
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def app_path(*parts):
    return os.path.join(BASE_DIR, *parts)


MIN_RUN_SPEED_KMH = 10
END_RUN_SPEED_KMH = 5
MIN_RUN_DISTANCE_M = 50
RANKING_FILE = app_path("ranking.csv")
PROFILES_FILE = app_path("profiles.json")
SPOTS_FILE = app_path("spots.json")

NEW_ENTRY = "➕ Add new..."

# --- Sport-Modus (Windsurf / Kitesurf) -------------------------------------
# Eine DB, eine sessions-Tabelle mit Spalte `sport`. Der aktive Sport kommt aus
# dem URL-Parameter ?sport= (bleibt so über Reload/Link erhalten). Standard:
# Windsurf. So bleiben Caching & Performance unverändert – nur ein WHERE mehr.
SPORTS = ("windsurf", "kitesurf", "wingsurf", "sup")

# Beta-Hinweis neben dem Logo: verirrte Besucher sollen wissen, dass die Seite
# noch im Aufbau ist (volle Geschwindigkeit erst nach Umzug auf einen Pay-Server).
# em-basiert -> bleibt relativ zum jeweiligen Logo klein, weiss, dezent.
BETA_BADGE = (
    "<span style=\"font-size:.34em;font-weight:700;color:#fff;opacity:.85;"
    "vertical-align:super;margin-left:.45em;letter-spacing:0;white-space:nowrap;\">"
    "Beta 0.7</span>"
)

SPORT_META = {
    "windsurf": {
        "label": "🏄 Windsurf",
        "emoji": "🏄",
        "title": "WINDSURF",
        "gear_label": "Sail",                       # 2. Material (neben Board)
        "gear_size_unit": "m²",                     # Größeneinheit des 2. Materials
        "gear_type_label": "Foil / Fin",            # Label des gear_type-Felds
        "gear_type_options": ["Fin", "Foil"],
        # Profil-Keys (Autovervollständigung). Windsurf nutzt die bestehenden
        # Schlüssel (Abwärtskompatibilität), die anderen eigene.
        "boards_key": "boards",
        "gear_key": "sails",
        "bg_stem": "background",                    # assets/background.*
    },
    "kitesurf": {
        "label": "🪁 Kitesurf",
        "emoji": "🪁",
        "title": "KITESURF",
        "gear_label": "Kite",
        "gear_size_unit": "m²",
        "gear_type_label": "Foil / Twintip",
        "gear_type_options": ["Twintip", "Foil"],
        "boards_key": "kite_boards",
        "gear_key": "kites",
        "bg_stem": "background_kite",               # assets/background_kite.*
    },
    "wingsurf": {
        "label": "🪽 Wingsurf",
        "emoji": "🪽",
        "title": "WINGSURF",
        "gear_label": "Wing",
        "gear_size_unit": "m²",
        "gear_type_label": "Foil / Fin",
        "gear_type_options": ["Foil", "Fin"],
        "boards_key": "wing_boards",
        "gear_key": "wings",
        "bg_stem": "background_wing",               # assets/background_wing.*
    },
    "sup": {
        "label": "🛶 SUP",
        "emoji": "🛶",
        "title": "SUP",
        "gear_label": "Paddle",
        "gear_size_unit": "",                       # Paddel hat keine m²-Größe
        "gear_type_label": "Foil / Fin",
        "gear_type_options": ["Fin", "Foil"],
        "boards_key": "sup_boards",
        "gear_key": "paddles",
        "bg_stem": "background_sup",                # assets/background_sup.*
    },
}


def active_sport():
    """Aktiver Sport aus ?sport= (validiert), Standard 'windsurf'."""
    s = st.query_params.get("sport", "windsurf")
    return s if s in SPORTS else "windsurf"


def format_board(brand, model, volume=0):
    """Einheitlicher Anzeige-String fürs Board (Session-Upload UND Profil), damit
    identische Eingaben denselben String ergeben (keine Duplikate)."""
    s = f"{(brand or '').strip()} {(model or '').strip()}".strip()
    try:
        vol = float(volume or 0)
    except (TypeError, ValueError):
        vol = 0
    if vol > 0:
        s = f"{s} {int(vol)}L".strip()
    return s


def format_gear(brand, model, size=0.0, unit="m²"):
    """Einheitlicher Anzeige-String fürs 2. Material (Sail/Kite/Wing/Paddle)."""
    s = f"{(brand or '').strip()} {(model or '').strip()}".strip()
    try:
        sz = float(size or 0)
    except (TypeError, ValueError):
        sz = 0
    if unit and sz > 0:
        s = f"{s} {sz:.1f} {unit}".strip()
    return s


# WMO-Wettercodes -> (Emoji, Beschreibung)
WEATHER_CODES = {
    0: ("☀️", "Clear"),
    1: ("🌤️", "Mainly clear"),
    2: ("⛅", "Partly cloudy"),
    3: ("☁️", "Overcast"),
    45: ("🌫️", "Fog"),
    48: ("🌫️", "Rime fog"),
    51: ("🌦️", "Light drizzle"),
    53: ("🌦️", "Drizzle"),
    55: ("🌧️", "Heavy drizzle"),
    61: ("🌦️", "Light rain"),
    63: ("🌧️", "Rain"),
    65: ("🌧️", "Heavy rain"),
    66: ("🌧️", "Freezing rain"),
    67: ("🌧️", "Heavy freezing rain"),
    71: ("🌨️", "Light snowfall"),
    73: ("🌨️", "Snowfall"),
    75: ("❄️", "Heavy snowfall"),
    77: ("🌨️", "Snow grains"),
    80: ("🌦️", "Light showers"),
    81: ("🌧️", "Showers"),
    82: ("⛈️", "Violent showers"),
    85: ("🌨️", "Snow showers"),
    86: ("❄️", "Heavy snow showers"),
    95: ("⛈️", "Thunderstorm"),
    96: ("⛈️", "Thunderstorm with hail"),
    99: ("⛈️", "Severe thunderstorm with hail"),
}

WEEKDAYS = [
    "Monday", "Tuesday", "Wednesday", "Thursday",
    "Friday", "Saturday", "Sunday",
]

# 16-Punkte-Kompass (O = Ost im Deutschen -> E im Englischen)
COMPASS = [
    "N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
    "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW",
]

# Pfeile in Strömungsrichtung (wohin der Wind weht)
WIND_ARROWS = ["⬆️", "↗️", "➡️", "↘️", "⬇️", "↙️", "⬅️", "↖️"]


def degrees_to_compass(degrees):
    if degrees is None:
        return None

    return COMPASS[int(degrees / 22.5 + 0.5) % 16]


def wind_arrow(degrees):
    if degrees is None:
        return ""

    toward = (degrees + 180) % 360

    return WIND_ARROWS[int(toward / 45 + 0.5) % 8]


def kmh_to_beaufort(kmh):
    if kmh is None:
        return None

    # Obergrenzen der Beaufort-Stufen 0..11 in km/h
    limits = [1, 5, 11, 19, 28, 38, 49, 61, 74, 88, 102, 117]

    for beaufort, upper in enumerate(limits):
        if kmh <= upper:
            return beaufort

    return 12


st.set_page_config(
    page_title="MyWaterSessions",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def load_css(path):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)


def image_to_base64(path):
    if not os.path.exists(path):
        return None

    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


# Endung -> MIME-Typ für base64-Data-URIs (CSS/HTML).
_IMAGE_MIME = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".webp": "image/webp",
    ".gif": "image/gif",
}


def image_data_uri(path):
    """base64-Data-URI eines Bildes inkl. passendem MIME-Typ (oder None).

    Anders als image_to_base64 setzt diese Funktion den MIME-Typ anhand der
    Dateiendung – damit funktionieren JPEG, PNG, WebP usw. gleichermaßen.
    """
    if not path or not os.path.exists(path):
        return None

    ext = os.path.splitext(path)[1].lower()
    mime = _IMAGE_MIME.get(ext, "application/octet-stream")

    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    return f"data:{mime};base64,{b64}"


@st.cache_data(show_spinner=False)
def _bg_uri_cached(path, _mtime):
    """base64-Kodierung des Hintergrundbilds, gecacht je (Pfad, mtime).

    Die mtime ist Teil des Cache-Keys: Tauscht/aktualisierst du die Bilddatei,
    ändert sich die mtime → der Cache lädt automatisch neu. So bleibt das teure
    base64-Kodieren gecacht, ohne bei einem Bildwechsel das alte Bild zu zeigen.
    """
    return image_data_uri(path)


def background_data_uri(sport="windsurf"):
    """Findet das Hintergrundbild (je Sport) und liefert die Data-URI.

    Kitesurf nutzt zuerst assets/background_kite.* und fällt – falls (noch) nicht
    vorhanden – auf das Windsurf-Bild zurück. Du kannst das Bild als .webp ODER
    .jpg/.jpeg/.png ablegen (letzter Fallback: header.*). Die Existenz-/mtime-
    Prüfung läuft ungecacht (billig); nur das Kodieren ist gecacht und
    invalidiert beim Bildwechsel automatisch.
    """
    stem = SPORT_META.get(sport, {}).get("bg_stem", "background")
    stems = [stem] if stem == "background" else [stem, "background"]
    exts = [".webp", ".jpg", ".jpeg", ".png"]

    candidates = [("assets", f"{stem}{ext}") for stem in stems for ext in exts]
    candidates += [
        ("assets", "header.webp"),
        ("assets", "header.jpg"),
        ("assets", "header.jpeg"),
    ]

    for parts in candidates:
        path = app_path(*parts)
        if os.path.exists(path):
            return _bg_uri_cached(path, os.path.getmtime(path))

    return None


# =====================================================================
#  Rechtsseiten (Impressum / Datenschutz)
#  Erreichbar über ?seite=impressum bzw. ?seite=datenschutz – BEWUSST vor
#  dem Login-Gate, damit das Impressum (Pflicht!) ohne Anmeldung erreichbar
#  ist. Bitte die [PLATZHALTER] im LEGAL_OPERATOR ausfüllen.
# =====================================================================

LEGAL_OPERATOR = {
    "name": "Jan Brinkman",
    "street": "Thorner Strasse 12",
    "city": "51469 Bergisch Gladbach",
    "country": "Deutschland",
    "email": "Windsurfspeedchallenge@outlook.de",
    }

LEGAL_STAND = "Juni 2026"


def _legal_back_button(key):
    if st.button("← Zurück zur Startseite", key=key):
        st.query_params.clear()
        st.rerun()


def render_impressum():
    op = LEGAL_OPERATOR

    st.markdown("## 📄 Impressum")
    st.caption("Angaben gemäß § 5 Digitale-Dienste-Gesetz (DDG)")

    # Optionale Felder (z.B. USt-IdNr.) werden nur angezeigt, wenn gesetzt –
    # so stürzt nichts ab, wenn eine Zeile in LEGAL_OPERATOR entfernt wird.
    vat_block = ""
    if op.get("vat"):
        vat_block = f"**Umsatzsteuer-Identifikationsnummer**<br>\n{op['vat']}\n\n"

    st.markdown(
        f"""
**Diensteanbieter**<br>
{op.get('name', '')}<br>
{op.get('street', '')}<br>
{op.get('city', '')}<br>
{op.get('country', '')}

**Kontakt**<br>
E-Mail: {op.get('email', '')}

{vat_block}**Verantwortlich für den Inhalt** nach § 18 Abs. 2 MStV<br>
{op.get('name', '')}, Anschrift wie oben.

---

**Haftung für Inhalte**<br>
Als Diensteanbieter sind wir für eigene Inhalte auf diesen Seiten nach den
allgemeinen Gesetzen verantwortlich. Wir sind jedoch nicht verpflichtet,
übermittelte oder gespeicherte fremde Informationen zu überwachen.

**Haftung für Links**<br>
Unser Angebot enthält Links zu externen Websites Dritter, auf deren Inhalte
wir keinen Einfluss haben. Für diese fremden Inhalte ist stets der jeweilige
Anbieter der Seiten verantwortlich.
""",
        unsafe_allow_html=True,
    )

    _legal_back_button("back_impressum")


def render_datenschutz():
    op = LEGAL_OPERATOR

    st.markdown("## 🔒 Datenschutzerklärung")

    st.markdown(
        f"""
### 1. Verantwortlicher

Verantwortlich für die Datenverarbeitung auf dieser Website ist:<br>
{op.get('name', '')}, {op.get('street', '')}, {op.get('city', '')}, {op.get('country', '')}<br>
E-Mail: {op.get('email', '')}

### 2. Welche Daten wir verarbeiten

**Konto:** Benutzername und Passwort. Das Passwort wird ausschließlich als
gesalzener Hash (PBKDF2-HMAC-SHA256) gespeichert – im Klartext ist es uns
nicht bekannt.

**Session-/Leistungsdaten:** Pro hochgeladener Windsurf-Session speichern wir
Datum, Surfspot, Board, Segel sowie die berechneten Kennzahlen
(Höchstgeschwindigkeit über 1 s und 30 s, längster Run, Gesamtstrecke) und die
zur Session passenden Wetterdaten.

**GPS-Daten:** Deine hochgeladene FIT-Datei enthält GPS-Punkte. Diese werden
nur kurzzeitig zur Berechnung deiner Kennzahlen und zur Anzeige der Karte
**für dich selbst** verwendet. **Der GPS-Track wird nicht dauerhaft
gespeichert und niemals öffentlich angezeigt oder mit anderen Nutzern
geteilt.**

**Gruppen:** Von dir erstellte oder beigetretene Gruppen und der jeweilige
Mitgliedsstatus.

### 3. Öffentliche Sichtbarkeit (Ranking)

Diese App ist eine Community-Bestenliste. Wenn du eine Session speicherst,
werden **Benutzername, Datum, Surfspot, Board, Segel, die Speed-Werte und die
Wetterangaben** im Ranking für andere angemeldete Nutzer sichtbar. In privaten
Gruppen sind die Ergebnisse nur für bestätigte Mitglieder sichtbar.

### 4. Zwecke und Rechtsgrundlagen

- **Konto & Anmeldung** zur Bereitstellung der App – Art. 6 Abs. 1 lit. b DSGVO
  (Nutzungsverhältnis).
- **Veröffentlichung deiner Ergebnisse im Ranking** – Art. 6 Abs. 1 lit. a DSGVO
  (deine Einwilligung, die du bei der Registrierung erteilst und jederzeit mit
  Wirkung für die Zukunft widerrufen kannst).

### 5. Cookies

Setzt du beim Login „Angemeldet bleiben", speichern wir ein funktional
notwendiges Cookie (`surf_auth`) mit einem zufälligen Anmelde-Token, damit du
nicht bei jedem Besuch neu eingeben musst. Ohne diese Option werden keine
Cookies gesetzt. Es findet kein Tracking und keine Werbung statt.

### 6. Hosting

Die App wird auf **Streamlit Community Cloud** (Snowflake Inc., USA) betrieben.
Dabei können technisch bedingt Verbindungsdaten (z. B. IP-Adresse) in den USA
verarbeitet werden. Grundlage für die Übermittlung sind die
Standardvertragsklauseln der EU bzw. das EU-US Data Privacy Framework.

### 7. Externe Dienste (Wetter)

Zur Anzeige von Wetter und Vorhersage rufen wir die Schnittstellen von
**Open-Meteo** ab. Dabei werden die **Koordinaten des ausgewählten Spots** sowie
technische Verbindungsdaten an Open-Meteo übermittelt. Es werden keine Konto-
oder Session-Daten von dir übertragen.

### 8. Speicherdauer

Konto-, Session- und Gruppendaten speichern wir, bis du dein Konto bzw. die
jeweiligen Einträge löschst oder die Löschung verlangst.

### 9. Deine Rechte

Du hast das Recht auf Auskunft, Berichtigung, Löschung, Einschränkung der
Verarbeitung, Datenübertragbarkeit sowie auf Widerruf erteilter Einwilligungen
mit Wirkung für die Zukunft. Außerdem steht dir ein Beschwerderecht bei einer
Datenschutz-Aufsichtsbehörde zu.

Zur Ausübung deiner Rechte oder zur Löschung deines Kontos genügt eine formlose
Nachricht an: {op.get('email', '')}

---

*Stand: {LEGAL_STAND}. Diese Erklärung wird angepasst, wenn sich die
Datenverarbeitung ändert.*
""",
        unsafe_allow_html=True,
    )

    _legal_back_button("back_datenschutz")


def render_legal_page():
    """Zeigt eine Rechtsseite, falls per ?seite=... angefragt. True = behandelt."""
    page = st.query_params.get("seite")

    if page == "impressum":
        render_impressum()
        return True

    if page == "datenschutz":
        render_datenschutz()
        return True

    return False


# =====================================================================
#  Datenbank-Layer
#  Standard: lokale SQLite-Datei. Fürs Cloud-Deployment einfach eine
#  DATABASE_URL (z.B. kostenlose Postgres bei Neon/Supabase) in den
#  Streamlit-Secrets oder als Umgebungsvariable setzen – kein Code-Umbau.
# =====================================================================

DB_METADATA = MetaData()

users_table = Table(
    "users", DB_METADATA,
    Column("id", Integer, primary_key=True),
    Column("username", String(80), unique=True, nullable=False),
    Column("password_hash", String(255), nullable=False),
    Column("salt", String(64), nullable=False),
    Column("email", String(255)),       # Pflicht bei Neu-Registrierung
    Column("weight_kg", Float),         # optionales Profilfeld
    Column("height_cm", Float),         # optionales Profilfeld
    Column("created_at", DateTime, server_default=func.now()),
)

groups_table = Table(
    "groups", DB_METADATA,
    Column("id", Integer, primary_key=True),
    Column("name", String(120), unique=True, nullable=False),
    Column("is_private", Boolean, nullable=False, default=False),
    Column("owner_id", Integer, ForeignKey("users.id"), nullable=False),
    Column("created_at", DateTime, server_default=func.now()),
)

memberships_table = Table(
    "memberships", DB_METADATA,
    Column("id", Integer, primary_key=True),
    Column("user_id", Integer, ForeignKey("users.id"), nullable=False),
    Column("group_id", Integer, ForeignKey("groups.id"), nullable=False),
    Column("status", String(20), nullable=False, default="member"),  # member | pending
    Column("created_at", DateTime, server_default=func.now()),
    UniqueConstraint("user_id", "group_id", name="uq_membership"),
)

sessions_table = Table(
    "sessions", DB_METADATA,
    Column("id", Integer, primary_key=True),
    Column("sport", String(20)),  # "windsurf" | "kitesurf" (Altdaten: windsurf)
    Column("name", String(80)),  # = Benutzername des Fahrers
    Column("date", String(20)),
    Column("surfspot", String(200)),
    Column("board", String(200)),
    Column("sail", String(200)),
    Column("gear_type", String(10)),  # "Fin" / "Foil" / "Twintip"
    Column("fin_size_cm", Float),         # optional, nur bei gear_type == "Fin"
    Column("foil_front_cm2", Float),      # optional, nur bei gear_type == "Foil"
    Column("filename", String(255)),
    Column("total_distance_km", Float),
    Column("longest_run_km", Float),
    Column("longest_run_m", Float),
    Column("speed_1s_kmh", Float),
    Column("speed_1s_kn", Float),
    Column("speed_30s_kmh", Float),
    Column("speed_30s_kn", Float),
    Column("wind_kmh", Float),
    Column("gust_kmh", Float),
    Column("wind_dir_deg", Float),
    Column("temp_c", Float),
    Column("precip_mm", Float),
    Column("weather_code", Integer),
    Column("trust_score", Float),
    # Von der WaterSession-Uhr gelieferte Zusatzwerte (NULL bei FIT-Uploads):
    Column("jumps", Integer),            # Sprungzahl (Wind)
    Column("max_airtime_s", Float),      # laengste Airtime in Sekunden
    Column("max_jump_m", Float),         # geschaetzte hoechste Sprunghoehe (m)
    Column("strokes", Integer),          # Paddelschlaege (SUP)
    Column("cadence_spm", Integer),      # Paddelkadenz beim Stoppen (Schlaege/Minute)
    Column("max_cadence_spm", Integer),  # hoechste Kadenz der Session
    Column("duration_s", Integer),       # Aufzeichnungsdauer in Sekunden
    Column("source", String(20)),        # Herkunft, z.B. "watch"
    Column("start_lat", Float),          # Startposition (von der Uhr)
    Column("start_lon", Float),
    Column("track", String),             # GPS-Route als JSON [[lat,lon],...]
    Column("created_at", DateTime, server_default=func.now()),
)

profiles_table = Table(
    "profiles", DB_METADATA,
    Column("name", String(80), primary_key=True),  # Benutzername
    Column("data", String),  # JSON: {"spots": [], "boards": [], "sails": []}
)

spots_table = Table(
    "spots", DB_METADATA,
    Column("name", String(200), primary_key=True),
    Column("lat", Float),
    Column("lon", Float),
)

auth_tokens_table = Table(
    "auth_tokens", DB_METADATA,
    Column("token", String(64), primary_key=True),
    Column("user_id", Integer, ForeignKey("users.id"), nullable=False),
    Column("created_at", DateTime, server_default=func.now()),
)

# Langlebige Geraete-Tokens fuer den Upload von der WaterSession-Uhr. Bewusst
# getrennt von auth_tokens, da Letztere beim Logout geloescht werden – ein
# Geraete-Token soll den Logout ueberleben.
device_tokens_table = Table(
    "device_tokens", DB_METADATA,
    Column("token", String(64), primary_key=True),
    Column("user_id", Integer, ForeignKey("users.id"), nullable=False),
    Column("created_at", DateTime, server_default=func.now()),
)

# Freigabe-Codes zum Teilen des Equipments (Spots/Boards/Segel) zwischen
# Konten – z.B. Familie, die dasselbe Material nutzt. Code -> Besitzer.
equip_shares_table = Table(
    "equip_shares", DB_METADATA,
    Column("code", String(32), primary_key=True),
    Column("owner", String(80), nullable=False),
    Column("created_at", DateTime, server_default=func.now()),
)

# Werbung/Sponsor je Spot fuers Spot-TV – im Admin-Backoffice verwaltet. Logo
# als Bytes IN der DB, weil das Streamlit-Cloud-Dateisystem fluechtig ist
# (assets/-Uploads ueberleben keinen Neustart).
spot_ads_table = Table(
    "spot_ads", DB_METADATA,
    Column("spot", String(200), primary_key=True),
    Column("sponsor_name", String(200)),
    Column("sponsor_url", String(500)),
    Column("logo", LargeBinary),         # Bilddaten (PNG/JPG/WebP)
    Column("logo_mime", String(50)),     # z.B. image/png
    Column("active", Boolean, nullable=False, default=True),
    Column("updated_at", DateTime, server_default=func.now()),
)

# Beworbene Produkte/Angebote je Spot (Shop/Cafe), mit Bild + Link. Werden auf
# dem Spot-TV als Leiste ausgespielt.
spot_products_table = Table(
    "spot_products", DB_METADATA,
    Column("id", Integer, primary_key=True),
    Column("spot", String(200), nullable=False),
    Column("title", String(200), nullable=False),
    Column("price", String(40)),
    Column("url", String(500)),
    Column("image", LargeBinary),
    Column("image_mime", String(50)),
    Column("active", Boolean, nullable=False, default=True),
    Column("sort_order", Integer, default=0),
    Column("created_at", DateTime, server_default=func.now()),
)

# Spot-Infos fuers Spot-TV (unten): Beschreibungstext + Bild (Bytes) ODER eine
# Webcam-/Bild-URL. Im Admin-Backoffice editierbar.
spot_info_table = Table(
    "spot_info", DB_METADATA,
    Column("spot", String(200), primary_key=True),
    Column("description", String),
    Column("image", LargeBinary),
    Column("image_mime", String(50)),
    Column("webcam_url", String(500)),
    Column("country", String(80)),       # Land (fuer den Filter der Spots-Seite)
    Column("best_winds", String(120)),   # beste Windrichtungen, z.B. "SW, W, NW"
    Column("auto_filled", Boolean, nullable=False, default=False),  # KI-Entwurf?
    Column("updated_at", DateTime, server_default=func.now()),
)

# Mehrere Bilder je Spot (Galerie auf der Spots-Seite). Bytes in der DB.
spot_images_table = Table(
    "spot_images", DB_METADATA,
    Column("id", Integer, primary_key=True),
    Column("spot", String(200), nullable=False),
    Column("image", LargeBinary),
    Column("image_mime", String(50)),
    Column("sort_order", Integer, default=0),
    Column("created_at", DateTime, server_default=func.now()),
)

# Gruppen-Ereignisse (Rekorde / Top-3) – andere Mitglieder sehen sie beim
# nächsten Öffnen als Banner. Wird von create_all automatisch angelegt.
group_events_table = Table(
    "group_events", DB_METADATA,
    Column("id", Integer, primary_key=True),
    Column("group_id", Integer, ForeignKey("groups.id"), nullable=False),
    Column("username", String(80), nullable=False),
    Column("metric", String(40)),
    Column("rank", Integer),
    Column("value", Float),
    Column("message", String(255)),
    Column("created_at", DateTime, server_default=func.now()),
)

# Pro Nutzer: bis zu welcher Ereignis-ID wurde das News-Banner gelesen.
event_reads_table = Table(
    "event_reads", DB_METADATA,
    Column("user_id", Integer, primary_key=True),
    Column("last_seen_event_id", Integer, nullable=False, default=0),
)

# Pro Nutzer: persönliche Einstellungen (z.B. das Ranking-Filter-Preset) als JSON.
user_prefs_table = Table(
    "user_prefs", DB_METADATA,
    Column("username", String(80), primary_key=True),
    Column("data", String),
)

SESSION_FIELDS = [
    c.name for c in sessions_table.columns if c.name not in ("id", "created_at")
]

# track (GPS-Route als JSON) ist potenziell gross und wird nur in der Einzel-
# Detailansicht (Karte) gebraucht. Daher in den Massen-Ladevorgaengen (Rankings,
# Fahrer-Sessions) NICHT mitladen -> weniger Transfer/Speicher/Kopien. In der
# Detailansicht wird sie per ID nachgeholt (load_session_track).
_SESSION_COLS_NO_TRACK = [c for c in sessions_table.c if c.name != "track"]


def _database_url():
    try:
        url = st.secrets.get("DATABASE_URL")
        if url:
            return url
    except Exception:
        pass

    return os.environ.get("DATABASE_URL") or f"sqlite:///{app_path('surfapp.db')}"


def _py(value):
    """numpy/pandas-Skalar -> natives Python (für DB-Bindings).

    Postgres/psycopg2 kann numpy-Typen (z.B. numpy.float64) nicht binden –
    SQLite schluckt sie, Postgres nicht. Daher hier zu nativem Python wandeln
    und NaN zu None machen.
    """
    if value is None:
        return None
    if hasattr(value, "item"):
        try:
            value = value.item()
        except Exception:
            return value
    if isinstance(value, float) and value != value:  # NaN
        return None
    return value


def _migrate_legacy(engine):
    """Einmalige Übernahme alter CSV/JSON-Daten in die Datenbank."""
    with engine.begin() as conn:
        sessions_count = conn.execute(
            select(func.count()).select_from(sessions_table)
        ).scalar()

        if not sessions_count and os.path.exists(RANKING_FILE):
            try:
                df = pd.read_csv(RANKING_FILE)
                for _, row in df.iterrows():
                    values = {
                        c: _py(row[c])
                        for c in SESSION_FIELDS
                        if c in df.columns and pd.notna(row[c])
                    }
                    if values:
                        conn.execute(insert(sessions_table).values(**values))
            except Exception:
                pass

        profiles_count = conn.execute(
            select(func.count()).select_from(profiles_table)
        ).scalar()

        if not profiles_count and os.path.exists(PROFILES_FILE):
            try:
                with open(PROFILES_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for pname, pdata in data.items():
                    conn.execute(insert(profiles_table).values(
                        name=pname, data=json.dumps(pdata, ensure_ascii=False)
                    ))
            except Exception:
                pass

        spots_count = conn.execute(
            select(func.count()).select_from(spots_table)
        ).scalar()

        if not spots_count and os.path.exists(SPOTS_FILE):
            try:
                with open(SPOTS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for sname, coords in data.items():
                    conn.execute(insert(spots_table).values(
                        name=sname, lat=coords.get("lat"), lon=coords.get("lon")
                    ))
            except Exception:
                pass


@st.cache_resource(show_spinner=False)
def get_engine():
    url = _database_url()
    is_sqlite = url.startswith("sqlite")
    connect_args = {"check_same_thread": False} if is_sqlite else {}
    # Remote-Postgres (Neon/Supabase) kappt Leerlauf-Verbindungen. pool_pre_ping
    # fängt tote Verbindungen ab, pool_recycle erneuert sie vor dem Timeout –
    # so vermeiden wir den langsamen "stale connection"-Fehler + Retry, der die
    # App in der Cloud träge wirken lässt. Bei SQLite (lokal) irrelevant.
    engine_kwargs = {"pool_pre_ping": True}
    if not is_sqlite:
        engine_kwargs["pool_recycle"] = 1800
    engine = create_engine(url, connect_args=connect_args, **engine_kwargs)
    DB_METADATA.create_all(engine)
    _migrate_legacy(engine)
    return engine


@st.cache_resource(show_spinner=False)
def ensure_schema():
    """Legt fehlende Tabellen UND fehlende Spalten an – EINMALIG PRO PROZESS.

    Wichtig nach einem Deploy mit NEUEN Tabellen/Spalten: get_engine() ist mit
    @st.cache_resource gecacht, sodass dessen create_all nach dem Deploy ggf.
    nicht erneut läuft. Diese idempotente Prüfung stellt sicher, dass z.B.
    group_events/event_reads und neue Spalten wie sessions.trust_score
    existieren. create_all legt nur fehlende TABELLEN an – fehlende SPALTEN
    ergänzen wir per ALTER TABLE.

    PERFORMANCE: Mit @st.cache_resource gegated (prozessweit, über alle Sessions
    geteilt) statt über st.session_state. Sonst lief die teure Schema-Inspektion
    (`get_columns` je Tabelle = ein Remote-Roundtrip) bei JEDEM neuen Seiten-
    aufruf – auf einer entfernten Neon-DB mehrere Sekunden Grundladezeit.
    """
    engine = get_engine()
    DB_METADATA.create_all(engine, checkfirst=True)

    try:
        inspector = inspect(engine)
    except Exception:
        logging.exception("Schema-Inspektion fehlgeschlagen")
        return True

    # WICHTIG: jede ALTER-Anweisung in EIGENEM try – sonst blockiert ein einzelner
    # Fehler alle weiteren Spalten (z.B. die users-Spalten, die zuerst drankommen).
    for table in DB_METADATA.tables.values():
        try:
            existing = {c["name"] for c in inspector.get_columns(table.name)}
        except Exception:
            logging.exception("get_columns fehlgeschlagen: %s", table.name)
            continue

        for column in table.columns:
            if column.name in existing:
                continue

            col_type = column.type.compile(engine.dialect)

            try:
                with engine.begin() as conn:
                    conn.execute(text(
                        f'ALTER TABLE {table.name} ADD COLUMN "{column.name}" {col_type}'
                    ))
            except Exception:
                logging.exception(
                    "ALTER fehlgeschlagen: %s.%s (%s)", table.name, column.name, col_type
                )

    # Altdaten ohne Sport gehören zu Windsurf. Idempotent, eigener try-Block.
    try:
        with engine.begin() as conn:
            conn.execute(text(
                "UPDATE sessions SET sport = 'windsurf' WHERE sport IS NULL"
            ))
    except Exception:
        logging.exception("Sport-Backfill fehlgeschlagen")

    return True


_WATCH_COLUMNS = {
    "jumps": "INTEGER",
    "max_airtime_s": "DOUBLE PRECISION",
    "max_jump_m": "DOUBLE PRECISION",
    "strokes": "INTEGER",
    "cadence_spm": "INTEGER",
    "max_cadence_spm": "INTEGER",
    "duration_s": "INTEGER",
    "source": "VARCHAR(20)",
    "start_lat": "DOUBLE PRECISION",
    "start_lon": "DOUBLE PRECISION",
    "track": "TEXT",
}

@st.cache_resource(show_spinner=False)
def ensure_watch_columns():
    """Ergaenzt die von der WaterSession-Uhr genutzten sessions-Spalten.

    PERFORMANCE: @st.cache_resource -> laeuft genau EINMAL pro Prozess, nicht bei
    jedem Rerun. (Frueher per Modul-Flag abgesichert -> das wird bei Streamlits
    Skript-Reruns zurueckgesetzt, sodass die 7 ALTER-Roundtrips zur Neon-DB bei
    JEDEM Klick liefen = deutliche Verlangsamung.) Nach einem neuen Deploy laeuft
    ein frischer Prozess -> einmaliges Nachziehen. Zusaetzliche Absicherung: der
    Ingest-Dienst legt dieselben Spalten per ADD COLUMN IF NOT EXISTS an.
    Postgres nutzt ADD COLUMN IF NOT EXISTS; lokales SQLite erledigt ensure_schema.
    """
    engine = get_engine()
    if engine.dialect.name == "sqlite":
        return True

    for name, coltype in _WATCH_COLUMNS.items():
        try:
            with engine.begin() as conn:
                conn.execute(text(
                    f'ALTER TABLE sessions ADD COLUMN IF NOT EXISTS "{name}" {coltype}'
                ))
        except Exception:
            logging.exception("ALTER sessions ADD %s fehlgeschlagen", name)

    return True


# ---- Sessions ----

def save_session(entry):
    # _py: numpy-Typen -> natives Python (Postgres kann numpy nicht binden)
    values = {k: _py(entry.get(k)) for k in SESSION_FIELDS if k in entry}

    with get_engine().begin() as conn:
        conn.execute(insert(sessions_table).values(**values))

    clear_data_caches()


def update_session(session_id, fields):
    """Aktualisiert einzelne Felder einer Session (z.B. Spot/Board/Segel beim
    Nachpflegen einer von der Uhr hochgeladenen Session)."""
    clean = {k: _py(v) for k, v in fields.items() if k in SESSION_FIELDS}
    if not clean:
        return
    with get_engine().begin() as conn:
        conn.execute(
            update(sessions_table)
            .where(sessions_table.c.id == int(session_id))
            .values(**clean)
        )
    clear_data_caches()


def delete_session(session_id):
    """Loescht eine einzelne Session (z.B. fehlerhafte/Test-Sessions)."""
    with get_engine().begin() as conn:
        conn.execute(delete(sessions_table).where(sessions_table.c.id == int(session_id)))
    clear_data_caches()


@st.cache_data(ttl=3600, show_spinner=False)
def load_sessions(sport=None):
    """Sessions als DataFrame. sport=None -> alle Sportarten (z.B. für die
    geteilte Spot-Liste); sonst nur der angegebene Sport. sport ist Teil des
    Cache-Keys, damit Windsurf/Kite getrennt gecacht werden."""
    with get_engine().connect() as conn:
        query = select(*_SESSION_COLS_NO_TRACK)
        if sport:
            query = query.where(sessions_table.c.sport == sport)
        rows = conn.execute(query).mappings().all()

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame([dict(r) for r in rows])
    # Datum EINMAL hier (gecacht) in datetime wandeln -> die vielen
    # pd.to_datetime-Aufrufe weiter unten laufen dann auf datetime (schnell),
    # statt bei jedem Rerun Strings mit format="mixed" zu parsen.
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce", format="mixed")
    return df


@st.cache_data(ttl=3600, show_spinner=False)
def load_session_track(session_id):
    """Holt die (potenziell grosse) GPS-Route einer einzelnen Session per ID –
    bewusst getrennt von den Massen-Ladevorgaengen, nur fuer die Detailkarte."""
    if session_id is None:
        return None
    with get_engine().connect() as conn:
        row = conn.execute(
            select(sessions_table.c.track).where(
                sessions_table.c.id == int(session_id)
            )
        ).first()
    return row[0] if row else None


def current_username():
    user = st.session_state.get("user")
    return user["username"] if user else None


def session_exists(filename, name=None, sport=None):
    if not filename:
        return False

    if name is None:
        name = current_username()

    with get_engine().connect() as conn:
        query = (
            select(func.count())
            .select_from(sessions_table)
            .where(sessions_table.c.filename == filename)
        )

        if name is not None:
            query = query.where(sessions_table.c.name == name)

        if sport:
            query = query.where(sessions_table.c.sport == sport)

        return bool(conn.execute(query).scalar())


@st.cache_data(ttl=600, show_spinner=False)
def load_rider_sessions(name, sport=None):
    """Alle gespeicherten Sessions eines Fahrers, neueste zuerst (optional auf
    einen Sport eingegrenzt; sport ist Teil des Cache-Keys)."""
    with get_engine().connect() as conn:
        query = select(*_SESSION_COLS_NO_TRACK).where(sessions_table.c.name == name)
        if sport:
            query = query.where(sessions_table.c.sport == sport)
        rows = conn.execute(query).mappings().all()

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame([dict(r) for r in rows])

    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce", format="mixed")
        df = df.sort_values("date", ascending=False)

    return df.reset_index(drop=True)


@st.cache_data(ttl=3600, show_spinner=False)
def load_user_weights():
    """{username: weight_kg} für den Gewichts-Filter in den Rankings (nur User
    mit hinterlegtem Gewicht)."""
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(users_table.c.username, users_table.c.weight_kg)
        ).mappings().all()

    return {r["username"]: r["weight_kg"] for r in rows if r["weight_kg"] is not None}


# ---- Profile (Spots/Boards/Segel je Fahrer) ----

@st.cache_data(ttl=3600, show_spinner=False)
def load_profiles():
    with get_engine().connect() as conn:
        rows = conn.execute(select(profiles_table)).mappings().all()

    profiles = {}

    for row in rows:
        try:
            profiles[row["name"]] = json.loads(row["data"]) if row["data"] else {}
        except Exception:
            profiles[row["name"]] = {}

    return profiles


def update_profile(name, spot, board, gear, sport="windsurf"):
    if not name:
        return

    boards_key = SPORT_META[sport]["boards_key"]
    gear_key = SPORT_META[sport]["gear_key"]

    with get_engine().begin() as conn:
        # Profil-Zeile frisch lesen (nicht aus dem Cache), damit der Merge
        # keine zwischenzeitlich ergänzten Einträge verliert.
        row = conn.execute(
            select(profiles_table.c.data).where(profiles_table.c.name == name)
        ).first()

        rider = {}
        if row and row[0]:
            try:
                rider = json.loads(row[0])
            except Exception:
                rider = {}

        if not rider:
            rider = {}

        # Spots sind sportübergreifend geteilt; Board/Sail bzw. Board/Kite je Sport.
        for key, value in (("spots", spot), (boards_key, board), (gear_key, gear)):
            items = rider.setdefault(key, [])

            if value and value not in items:
                items.append(value)

        data = json.dumps(rider, ensure_ascii=False)

        if row:
            conn.execute(
                update(profiles_table).where(profiles_table.c.name == name).values(data=data)
            )
        else:
            conn.execute(insert(profiles_table).values(name=name, data=data))

    clear_data_caches()


# ---- Spots (Koordinaten-Cache) ----

@st.cache_data(ttl=3600, show_spinner=False)
def load_spots():
    with get_engine().connect() as conn:
        rows = conn.execute(select(spots_table)).mappings().all()

    return {r["name"]: {"lat": r["lat"], "lon": r["lon"]} for r in rows}


def update_spot_coords(spot, lat, lon):
    if not spot or lat is None or lon is None:
        return

    values = {"lat": round(float(lat), 5), "lon": round(float(lon), 5)}

    with get_engine().begin() as conn:
        exists = conn.execute(
            select(spots_table.c.name).where(spots_table.c.name == spot)
        ).first()

        if exists:
            conn.execute(
                update(spots_table).where(spots_table.c.name == spot).values(**values)
            )
        else:
            conn.execute(insert(spots_table).values(name=spot, **values))

    clear_data_caches()


def all_known_spots():
    names = set(load_spots().keys())

    for rider in load_profiles().values():
        for spot in rider.get("spots", []):
            names.add(spot)

    sessions = load_sessions()

    if not sessions.empty and "surfspot" in sessions.columns:
        for spot in sessions["surfspot"].dropna().astype(str):
            names.add(spot)

    return sorted(name for name in names if name)


# =====================================================================
#  Authentifizierung
# =====================================================================

def _hash_password(password, salt):
    return hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt.encode("utf-8"), 200_000
    ).hex()


def get_user(username):
    with get_engine().connect() as conn:
        row = conn.execute(
            select(users_table).where(users_table.c.username == username)
        ).mappings().first()

    return dict(row) if row else None


def _valid_email(email):
    """Einfache, pragmatische E-Mail-Prüfung (kein voller RFC-Check)."""
    email = (email or "").strip()
    if "@" not in email or len(email) < 5:
        return False
    local, _, domain = email.partition("@")
    return bool(local) and "." in domain and not domain.startswith(".") and not domain.endswith(".")


def register_user(username, password, email=""):
    username = (username or "").strip()
    email = (email or "").strip()

    if not username or not password:
        return False, "Username and password must not be empty."

    if not _valid_email(email):
        return False, "Please enter a valid email address."

    if len(password) < 6:
        return False, "The password must be at least 6 characters long."

    if get_user(username):
        return False, "This username is already taken."

    salt = secrets.token_hex(16)

    with get_engine().begin() as conn:
        conn.execute(insert(users_table).values(
            username=username,
            password_hash=_hash_password(password, salt),
            salt=salt,
            email=email,
        ))

    return True, "Registration successful – you can now log in."


def update_user_account(user_id, email=None, weight_kg=None, height_cm=None):
    """Aktualisiert E-Mail/Gewicht/Größe eines Users (nur übergebene Felder)."""
    if user_id is None:
        return

    values = {}
    if email is not None:
        values["email"] = email.strip()
    if weight_kg is not None:
        values["weight_kg"] = float(weight_kg) if weight_kg else None
    if height_cm is not None:
        values["height_cm"] = float(height_cm) if height_cm else None

    if not values:
        return

    with get_engine().begin() as conn:
        conn.execute(
            update(users_table).where(users_table.c.id == int(user_id)).values(**values)
        )


def change_password(user_id, old_password, new_password):
    """Ändert das Passwort, wenn das alte stimmt. (ok, message)."""
    if user_id is None:
        return False, "Not logged in."

    if not new_password or len(new_password) < 6:
        return False, "The new password must be at least 6 characters long."

    with get_engine().connect() as conn:
        row = conn.execute(
            select(users_table.c.password_hash, users_table.c.salt)
            .where(users_table.c.id == int(user_id))
        ).mappings().first()

    if not row:
        return False, "User not found."

    if not secrets.compare_digest(
        _hash_password(old_password, row["salt"]), row["password_hash"]
    ):
        return False, "The current password is incorrect."

    salt = secrets.token_hex(16)
    with get_engine().begin() as conn:
        conn.execute(
            update(users_table).where(users_table.c.id == int(user_id)).values(
                password_hash=_hash_password(new_password, salt),
                salt=salt,
            )
        )

    return True, "Password changed."


def set_profile_list(name, key, values):
    """Überschreibt eine Profil-Liste (z.B. Spots/Boards/Sails) – zum Entfernen
    von Einträgen im Profil-Bereich. Andere Keys bleiben unangetastet."""
    if not name:
        return

    with get_engine().begin() as conn:
        row = conn.execute(
            select(profiles_table.c.data).where(profiles_table.c.name == name)
        ).first()

        rider = {}
        if row and row[0]:
            try:
                rider = json.loads(row[0])
            except Exception:
                rider = {}

        rider[key] = list(values)
        data = json.dumps(rider, ensure_ascii=False)

        if row:
            conn.execute(
                update(profiles_table).where(profiles_table.c.name == name).values(data=data)
            )
        else:
            conn.execute(insert(profiles_table).values(name=name, data=data))

    clear_data_caches()


# ---- Equipment teilen (Freigabe-Code) ----

@st.cache_resource(show_spinner=False)
def _ensure_equip_shares_table():
    try:
        equip_shares_table.create(get_engine(), checkfirst=True)
    except Exception:
        logging.exception("equip_shares-Tabelle konnte nicht angelegt werden")
    return True


def get_or_create_share_code(username):
    """Liefert den bestehenden Freigabe-Code des Nutzers oder legt einen an."""
    _ensure_equip_shares_table()
    with get_engine().begin() as conn:
        row = conn.execute(
            select(equip_shares_table.c.code)
            .where(equip_shares_table.c.owner == username).limit(1)
        ).first()
        if row:
            return row[0]
        code = secrets.token_hex(4)   # 8 Zeichen, leicht abzutippen
        conn.execute(insert(equip_shares_table).values(code=code, owner=username))
        return code


def regenerate_share_code(username):
    """Verwirft den alten Code (entzieht damit Zugriff) und erzeugt einen neuen."""
    _ensure_equip_shares_table()
    with get_engine().begin() as conn:
        conn.execute(delete(equip_shares_table).where(equip_shares_table.c.owner == username))
        code = secrets.token_hex(4)
        conn.execute(insert(equip_shares_table).values(code=code, owner=username))
        return code


def owner_for_share_code(code):
    if not code:
        return None
    _ensure_equip_shares_table()
    code = code.strip().lower()
    with get_engine().connect() as conn:
        row = conn.execute(
            select(equip_shares_table.c.owner).where(equip_shares_table.c.code == code)
        ).first()
    return row[0] if row else None


def copy_equipment(from_name, to_name):
    """Führt das Equipment (alle Listen: Spots/Boards/Segel/… aller Sportarten)
    von from_name in to_name zusammen – ohne Duplikate. Gibt die Anzahl neu
    hinzugefügter Einträge zurück."""
    if not from_name or not to_name or from_name == to_name:
        return 0
    with get_engine().begin() as conn:
        src_row = conn.execute(
            select(profiles_table.c.data).where(profiles_table.c.name == from_name)
        ).first()
        if not src_row or not src_row[0]:
            return 0
        try:
            src = json.loads(src_row[0])
        except Exception:
            return 0

        dst_row = conn.execute(
            select(profiles_table.c.data).where(profiles_table.c.name == to_name)
        ).first()
        dst = {}
        if dst_row and dst_row[0]:
            try:
                dst = json.loads(dst_row[0])
            except Exception:
                dst = {}

        added = 0
        for key, vals in src.items():
            if isinstance(vals, list):
                existing = dst.get(key)
                if not isinstance(existing, list):
                    existing = []
                for v in vals:
                    if v and v not in existing:
                        existing.append(v)
                        added += 1
                dst[key] = existing

        data = json.dumps(dst, ensure_ascii=False)
        if dst_row:
            conn.execute(
                update(profiles_table).where(profiles_table.c.name == to_name).values(data=data)
            )
        else:
            conn.execute(insert(profiles_table).values(name=to_name, data=data))

    clear_data_caches()
    return added


def verify_login(username, password):
    user = get_user((username or "").strip())

    if not user:
        return False

    return secrets.compare_digest(
        _hash_password(password, user["salt"]), user["password_hash"]
    )


# ---- Persistenter Login per Cookie ("Angemeldet bleiben") ----

def create_auth_token(user_id):
    token = secrets.token_urlsafe(32)

    with get_engine().begin() as conn:
        conn.execute(insert(auth_tokens_table).values(token=token, user_id=user_id))

    return token


def user_for_token(token):
    if not token:
        return None

    with get_engine().connect() as conn:
        row = conn.execute(
            select(users_table.c.id, users_table.c.username)
            .select_from(
                auth_tokens_table.join(
                    users_table, auth_tokens_table.c.user_id == users_table.c.id
                )
            )
            .where(auth_tokens_table.c.token == token)
        ).mappings().first()

    return dict(row) if row else None


def delete_auth_token(token):
    if not token:
        return

    with get_engine().begin() as conn:
        conn.execute(delete(auth_tokens_table).where(auth_tokens_table.c.token == token))


# ---- Geraete-Tokens (Upload von der WaterSession-Uhr) ----

@st.cache_resource(show_spinner=False)
def _ensure_device_tokens_table():
    """Legt die device_tokens-Tabelle bei Bedarf an. @st.cache_resource ->
    genau einmal pro Prozess (checkfirst macht sonst pro Aufruf einen
    DB-Roundtrip). Nach einem neuen Deploy laeuft ein frischer Prozess."""
    try:
        device_tokens_table.create(get_engine(), checkfirst=True)
    except Exception:
        logging.exception("device_tokens-Tabelle konnte nicht angelegt werden")
    return True


def get_or_create_device_token(user_id):
    """Liefert den bestehenden Geraete-Token des Users oder legt einen an."""
    _ensure_device_tokens_table()
    with get_engine().begin() as conn:
        row = conn.execute(
            select(device_tokens_table.c.token)
            .where(device_tokens_table.c.user_id == user_id)
            .limit(1)
        ).first()
        if row:
            return row[0]

        token = secrets.token_urlsafe(32)
        conn.execute(insert(device_tokens_table).values(token=token, user_id=user_id))
        return token


def regenerate_device_token(user_id):
    """Verwirft alte Geraete-Tokens des Users und erzeugt einen neuen."""
    _ensure_device_tokens_table()
    with get_engine().begin() as conn:
        conn.execute(
            delete(device_tokens_table).where(device_tokens_table.c.user_id == user_id)
        )
        token = secrets.token_urlsafe(32)
        conn.execute(insert(device_tokens_table).values(token=token, user_id=user_id))
        return token


# ---- Spot-Werbung (Admin-Backoffice) ----

@st.cache_resource(show_spinner=False)
def _ensure_ad_tables():
    try:
        spot_ads_table.create(get_engine(), checkfirst=True)
        spot_products_table.create(get_engine(), checkfirst=True)
        spot_info_table.create(get_engine(), checkfirst=True)
        spot_images_table.create(get_engine(), checkfirst=True)
    except Exception:
        logging.exception("Werbe-Tabellen konnten nicht angelegt werden")
    return True


def _clear_ad_caches():
    for fn in (load_spot_ad, load_spot_products, load_spot_info, load_all_spot_info,
               load_spot_images, load_spot_image_ids, _spot_thumb_uri):
        try:
            fn.clear()
        except Exception:
            pass


@st.cache_data(ttl=60, show_spinner=False)
def load_spot_images(spot):
    """Alle Galerie-Bilder eines Spots inkl. Bytes (sortiert) – fuers Backoffice."""
    if not spot:
        return []
    _ensure_ad_tables()
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(spot_images_table)
            .where(spot_images_table.c.spot == spot)
            .order_by(spot_images_table.c.sort_order, spot_images_table.c.id)
        ).mappings().all()
    return [dict(r) for r in rows]


@st.cache_data(ttl=60, show_spinner=False)
def load_spot_image_ids(spot):
    """Nur die Bild-IDs eines Spots (leichtgewichtig, ohne Bytes) – fuer die
    Galerie-Anzeige, die ueber _spot_thumb_uri gecachte Thumbnails laedt."""
    if not spot:
        return []
    _ensure_ad_tables()
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(spot_images_table.c.id)
            .where(spot_images_table.c.spot == spot)
            .order_by(spot_images_table.c.sort_order, spot_images_table.c.id)
        ).all()
    return [r[0] for r in rows]


@st.cache_data(ttl=3600, show_spinner=False)
def _spot_thumb_uri(image_id, max_dim=560):
    """Kleines, gecachtes Thumbnail (data-URI) eines Galerie-Bilds. Spart enorm
    Seitenlast gegenueber dem Voll-Bild-base64 in der Galerie."""
    _ensure_ad_tables()
    with get_engine().connect() as conn:
        row = conn.execute(
            select(spot_images_table.c.image, spot_images_table.c.image_mime)
            .where(spot_images_table.c.id == int(image_id))
        ).first()
    if not row or not row[0]:
        return None
    data, mime = _optimize_image(row[0], row[1], max_dim=max_dim, quality=72)
    return _bytes_to_data_uri(data, mime)


def add_spot_image(spot, image_bytes, image_mime):
    if not spot or not image_bytes:
        return
    image_bytes, image_mime = _optimize_image(image_bytes, image_mime)
    _ensure_ad_tables()
    with get_engine().begin() as conn:
        mx = conn.execute(
            select(func.coalesce(func.max(spot_images_table.c.sort_order), 0))
            .where(spot_images_table.c.spot == spot)
        ).scalar() or 0
        conn.execute(insert(spot_images_table).values(
            spot=spot, image=image_bytes, image_mime=image_mime or "image/jpeg",
            sort_order=int(mx) + 1,
        ))
    _clear_ad_caches()


def delete_spot_image(image_id):
    _ensure_ad_tables()
    with get_engine().begin() as conn:
        conn.execute(delete(spot_images_table).where(spot_images_table.c.id == int(image_id)))
    _clear_ad_caches()


@st.cache_data(ttl=60, show_spinner=False)
def load_all_spot_info():
    """Alle Spots MIT Beschreibung (ohne Bild-Bytes – leichtgewichtig) fuer die
    Spots-Seite: spot, country, description."""
    _ensure_ad_tables()
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(spot_info_table.c.spot, spot_info_table.c.country,
                   spot_info_table.c.description)
            .where(spot_info_table.c.description.isnot(None))
            .order_by(spot_info_table.c.spot)
        ).mappings().all()
    return [dict(r) for r in rows if (r["description"] or "").strip()]


@st.cache_data(ttl=60, show_spinner=False)
def load_spot_ad(spot):
    """Sponsor-Eintrag eines Spots (oder None)."""
    if not spot:
        return None
    _ensure_ad_tables()
    with get_engine().connect() as conn:
        row = conn.execute(
            select(spot_ads_table).where(spot_ads_table.c.spot == spot)
        ).mappings().first()
    return dict(row) if row else None


def save_spot_ad(spot, sponsor_name, sponsor_url, active,
                 logo_bytes=None, logo_mime=None, clear_logo=False):
    if not spot:
        return
    _ensure_ad_tables()
    values = {
        "sponsor_name": (sponsor_name or "").strip() or None,
        "sponsor_url": (sponsor_url or "").strip() or None,
        "active": bool(active),
    }
    if clear_logo:
        values["logo"] = None
        values["logo_mime"] = None
    elif logo_bytes is not None:
        logo_bytes, logo_mime = _optimize_image(logo_bytes, logo_mime, max_dim=600)
        values["logo"] = logo_bytes
        values["logo_mime"] = logo_mime or "image/png"

    with get_engine().begin() as conn:
        exists = conn.execute(
            select(spot_ads_table.c.spot).where(spot_ads_table.c.spot == spot)
        ).first()
        if exists:
            conn.execute(
                update(spot_ads_table).where(spot_ads_table.c.spot == spot).values(**values)
            )
        else:
            conn.execute(insert(spot_ads_table).values(spot=spot, **values))
    _clear_ad_caches()


def delete_spot_ad(spot):
    _ensure_ad_tables()
    with get_engine().begin() as conn:
        conn.execute(delete(spot_ads_table).where(spot_ads_table.c.spot == spot))
    _clear_ad_caches()


@st.cache_data(ttl=60, show_spinner=False)
def load_spot_products(spot, only_active=False):
    if not spot:
        return []
    _ensure_ad_tables()
    query = select(spot_products_table).where(spot_products_table.c.spot == spot)
    if only_active:
        query = query.where(spot_products_table.c.active == True)  # noqa: E712
    query = query.order_by(spot_products_table.c.sort_order, spot_products_table.c.id)
    with get_engine().connect() as conn:
        rows = conn.execute(query).mappings().all()
    return [dict(r) for r in rows]


def save_spot_product(product_id, spot, title, price, url, active, sort_order,
                      image_bytes=None, image_mime=None, clear_image=False):
    _ensure_ad_tables()
    values = {
        "spot": spot,
        "title": (title or "").strip(),
        "price": (price or "").strip() or None,
        "url": (url or "").strip() or None,
        "active": bool(active),
        "sort_order": int(sort_order or 0),
    }
    if clear_image:
        values["image"] = None
        values["image_mime"] = None
    elif image_bytes is not None:
        image_bytes, image_mime = _optimize_image(image_bytes, image_mime, max_dim=900)
        values["image"] = image_bytes
        values["image_mime"] = image_mime or "image/png"

    with get_engine().begin() as conn:
        if product_id:
            conn.execute(
                update(spot_products_table)
                .where(spot_products_table.c.id == int(product_id))
                .values(**values)
            )
        else:
            conn.execute(insert(spot_products_table).values(**values))
    _clear_ad_caches()


def delete_spot_product(product_id):
    _ensure_ad_tables()
    with get_engine().begin() as conn:
        conn.execute(
            delete(spot_products_table).where(spot_products_table.c.id == int(product_id))
        )
    _clear_ad_caches()


@st.cache_data(ttl=60, show_spinner=False)
def load_spot_info(spot):
    if not spot:
        return None
    _ensure_ad_tables()
    with get_engine().connect() as conn:
        row = conn.execute(
            select(spot_info_table).where(spot_info_table.c.spot == spot)
        ).mappings().first()
    return dict(row) if row else None


def save_spot_info(spot, description, webcam_url, country="", best_winds="",
                   image_bytes=None, image_mime=None, clear_image=False):
    if not spot:
        return
    _ensure_ad_tables()
    values = {
        "description": (description or "").strip() or None,
        "webcam_url": (webcam_url or "").strip() or None,
        "country": (country or "").strip() or None,
        "best_winds": (best_winds or "").strip() or None,
        "auto_filled": False,   # manuell gespeichert = geprueft
    }
    if clear_image:
        values["image"] = None
        values["image_mime"] = None
    elif image_bytes is not None:
        values["image"] = image_bytes
        values["image_mime"] = image_mime or "image/jpeg"

    with get_engine().begin() as conn:
        exists = conn.execute(
            select(spot_info_table.c.spot).where(spot_info_table.c.spot == spot)
        ).first()
        if exists:
            conn.execute(
                update(spot_info_table).where(spot_info_table.c.spot == spot).values(**values)
            )
        else:
            conn.execute(insert(spot_info_table).values(spot=spot, **values))
    _clear_ad_caches()


def _is_image_url(url):
    """True, wenn die URL direkt auf ein Bild zeigt (Snapshot-Webcam)."""
    return bool(re.search(r"\.(jpe?g|png|webp|gif)(\?|$)", (url or "").split("#")[0], re.I))


def _bytes_to_data_uri(data, mime):
    if not data:
        return None
    # Postgres (Neon) liefert bytea als memoryview -> in bytes wandeln.
    b64 = base64.b64encode(bytes(data)).decode("ascii")
    return f"data:{mime or 'image/png'};base64,{b64}"


def _optimize_image(data, mime, max_dim=1600, quality=82):
    """Macht ein hochgeladenes Bild webfaehig: skaliert auf max_dim herunter und
    komprimiert (JPEG, bzw. PNG bei Transparenz). Spart Speicher/Transfer und
    macht Galerie/TV schneller. Bei Fehlern bleibt das Original erhalten."""
    if not data:
        return data, mime
    try:
        from PIL import Image
        im = Image.open(io.BytesIO(bytes(data)))
        im.load()
        has_alpha = im.mode in ("RGBA", "LA") or (
            im.mode == "P" and "transparency" in im.info
        )
        im.thumbnail((max_dim, max_dim))  # nur verkleinern, Seitenverhaeltnis bleibt
        buf = io.BytesIO()
        if has_alpha:
            im.convert("RGBA").save(buf, format="PNG", optimize=True)
            return buf.getvalue(), "image/png"
        im.convert("RGB").save(buf, format="JPEG", quality=quality, optimize=True)
        return buf.getvalue(), "image/jpeg"
    except Exception:  # noqa: BLE001 – im Zweifel Original behalten
        return data, mime


# ---- Produkt-Metadaten aus einer Shop-URL ziehen (Open Graph / JSON-LD) ----

_BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/*;q=0.8,*/*;q=0.7",
    "Accept-Language": "de,en;q=0.8",
}


def _http_get(url, max_bytes=3_000_000, timeout=12):
    req = Request(url, headers=_BROWSER_HEADERS)
    with urlopen(req, timeout=timeout) as resp:
        ctype = resp.headers.get("Content-Type", "") or ""
        data = resp.read(max_bytes)
    return data, ctype


def _clean_text(value):
    if value is None:
        return None
    text = unescape(str(value)).strip()
    text = re.sub(r"\s+", " ", text)
    return text or None


def _format_price(price, currency):
    if price in (None, ""):
        return ""
    try:
        val = float(str(price).replace(",", "."))
        num = str(int(val)) if val == int(val) else f"{val:.2f}"
    except (TypeError, ValueError):
        num = str(price).strip()
    symbol = {"EUR": "€", "USD": "$", "GBP": "£", "CHF": "CHF "}.get((currency or "").upper())
    if symbol:
        return f"{symbol}{num}"
    return f"{num} {currency}".strip() if currency else num


def _iter_jsonld_products(data):
    """Sucht (rekursiv) alle schema.org-Product-Knoten in JSON-LD-Daten."""
    found = []

    def walk(node):
        if isinstance(node, list):
            for item in node:
                walk(item)
        elif isinstance(node, dict):
            if "@graph" in node:
                walk(node["@graph"])
            types = node.get("@type")
            types = types if isinstance(types, list) else [types]
            if any(isinstance(t, str) and "product" in t.lower() for t in types):
                found.append(node)

    walk(data)
    return found


def _jsonld_image(img):
    if isinstance(img, list) and img:
        img = img[0]
    if isinstance(img, dict):
        return img.get("url")
    return img if isinstance(img, str) else None


def fetch_product_meta(url):
    """Liest Titel, Hauptbild und (falls vorhanden) Preis aus einer Produktseite.
    Reihenfolge: JSON-LD (schema.org Product) > Open-Graph/Meta > <title>."""
    url = (url or "").strip()
    if not url.lower().startswith(("http://", "https://")):
        return {"error": "Bitte eine vollständige URL mit http(s):// angeben."}

    try:
        raw, _ctype = _http_get(url)
    except Exception as exc:  # noqa: BLE001 – Netzwerk-/Parse-Fehler sauber melden
        return {"error": f"Seite nicht erreichbar ({type(exc).__name__})."}

    page = raw.decode("utf-8", errors="replace")
    title = price = currency = image_url = None

    # 1) JSON-LD (zuverlaessigste Quelle fuer Preis)
    for block in re.findall(
        r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        page, re.S | re.I,
    ):
        try:
            data = json.loads(block.strip())
        except Exception:
            continue
        for node in _iter_jsonld_products(data):
            title = title or _clean_text(node.get("name"))
            image_url = image_url or _jsonld_image(node.get("image"))
            offers = node.get("offers")
            offer = offers[0] if isinstance(offers, list) and offers else offers
            if isinstance(offer, dict):
                if price is None and offer.get("price") is not None:
                    price = offer.get("price")
                currency = currency or offer.get("priceCurrency")

    # 2) Open Graph / Meta-Tags
    def meta(prop):
        m = re.search(
            r'<meta[^>]+(?:property|name)=["\']' + re.escape(prop) + r'["\'][^>]*>',
            page, re.I,
        )
        if not m:
            return None
        c = re.search(r'content=["\'](.*?)["\']', m.group(0), re.I | re.S)
        return _clean_text(c.group(1)) if c else None

    title = title or meta("og:title") or meta("twitter:title")
    image_url = (image_url or meta("og:image") or meta("og:image:secure_url")
                 or meta("twitter:image") or meta("twitter:image:src"))
    if price is None:
        price = meta("product:price:amount") or meta("og:price:amount")
    currency = currency or meta("product:price:currency") or meta("og:price:currency")

    # 3) <title> als letzter Ausweg
    if not title:
        tm = re.search(r"<title[^>]*>(.*?)</title>", page, re.S | re.I)
        if tm:
            title = _clean_text(tm.group(1))

    if image_url:
        image_url = urljoin(url, _clean_text(image_url) or "")

    return {
        "title": title or "",
        "price": _format_price(price, currency),
        "image_url": image_url or "",
    }


def download_image_bytes(url):
    """Laedt ein Bild herunter -> (bytes, mime) oder (None, None)."""
    url = (url or "").strip()
    if not url.lower().startswith(("http://", "https://")):
        return None, None
    try:
        data, ctype = _http_get(url, max_bytes=8_000_000)
    except Exception:
        return None, None
    mime = ctype.split(";")[0].strip().lower() if ctype else ""
    if not mime.startswith("image/"):
        mime = "image/jpeg"
    return data, mime


def fetch_page_description(url):
    """Holt einen Beschreibungstext von einer Seite (og:description / meta
    description / erster laengerer Absatz) – fuer die Spot-Info-Vorbefuellung."""
    url = (url or "").strip()
    if not url.lower().startswith(("http://", "https://")):
        return {"error": "Bitte eine vollständige URL mit http(s):// angeben."}
    try:
        raw, _ctype = _http_get(url)
    except Exception as exc:  # noqa: BLE001
        return {"error": f"Seite nicht erreichbar ({type(exc).__name__})."}

    page = raw.decode("utf-8", errors="replace")

    def meta(prop):
        m = re.search(
            r'<meta[^>]+(?:property|name)=["\']' + re.escape(prop) + r'["\'][^>]*>',
            page, re.I,
        )
        if not m:
            return None
        c = re.search(r'content=["\'](.*?)["\']', m.group(0), re.I | re.S)
        return _clean_text(c.group(1)) if c else None

    desc = meta("og:description") or meta("description") or meta("twitter:description")
    if not desc:
        for para in re.findall(r"<p[^>]*>(.*?)</p>", page, re.S | re.I):
            txt = _clean_text(re.sub(r"<[^>]+>", "", para))
            if txt and len(txt) > 60:
                desc = txt
                break
    return {"description": desc or ""}


# ---- Admin: Profil-/Konto-Verwaltung (fremde Daten – nur Backoffice) ----

def list_all_users():
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(users_table.c.id, users_table.c.username, users_table.c.email)
            .order_by(users_table.c.username)
        ).mappings().all()
    return [dict(r) for r in rows]


def get_device_token(user_id):
    """Aktuellen Geraete-Token eines Users lesen (ohne anzulegen)."""
    _ensure_device_tokens_table()
    with get_engine().connect() as conn:
        row = conn.execute(
            select(device_tokens_table.c.token)
            .where(device_tokens_table.c.user_id == int(user_id)).limit(1)
        ).first()
    return row[0] if row else None


def admin_set_device_token(user_id, token):
    """Setzt einen bestimmten Geraete-Token fuer einen User (z.B. den fest in der
    Uhr hinterlegten). Entfernt den Token von einem evtl. anderen Besitzer."""
    token = (token or "").strip()
    if not token:
        return False, "Token must not be empty."
    _ensure_device_tokens_table()
    with get_engine().begin() as conn:
        conn.execute(delete(device_tokens_table).where(device_tokens_table.c.token == token))
        conn.execute(
            delete(device_tokens_table).where(device_tokens_table.c.user_id == int(user_id))
        )
        conn.execute(insert(device_tokens_table).values(token=token, user_id=int(user_id)))
    return True, "Token assigned."


def admin_set_password(user_id, new_password):
    if not new_password or len(new_password) < 6:
        return False, "The password must be at least 6 characters long."
    salt = secrets.token_hex(16)
    with get_engine().begin() as conn:
        conn.execute(
            update(users_table).where(users_table.c.id == int(user_id)).values(
                password_hash=_hash_password(new_password, salt), salt=salt
            )
        )
    return True, "Password updated."


def admin_rename_profile(old_name, new_name):
    """Benennt einen Fahrer/ein Konto um – inkl. aller verknuepften Zeilen, die
    ueber den Namen (statt user_id) referenzieren."""
    old_name = (old_name or "").strip()
    new_name = (new_name or "").strip()
    if not old_name or not new_name or old_name == new_name:
        return False, "Choose a different new name."
    with get_engine().begin() as conn:
        clash = conn.execute(
            select(users_table.c.id).where(users_table.c.username == new_name)
        ).first()
        if clash:
            return False, "A user with that name already exists."
        conn.execute(
            update(users_table).where(users_table.c.username == old_name).values(username=new_name)
        )
        conn.execute(
            update(sessions_table).where(sessions_table.c.name == old_name).values(name=new_name)
        )
        prow = conn.execute(
            select(profiles_table.c.data).where(profiles_table.c.name == old_name)
        ).first()
        if prow:
            conn.execute(delete(profiles_table).where(profiles_table.c.name == new_name))
            conn.execute(
                update(profiles_table).where(profiles_table.c.name == old_name).values(name=new_name)
            )
        conn.execute(
            update(equip_shares_table).where(equip_shares_table.c.owner == old_name)
            .values(owner=new_name)
        )
        conn.execute(
            update(user_prefs_table).where(user_prefs_table.c.username == old_name)
            .values(username=new_name)
        )
        conn.execute(
            update(group_events_table).where(group_events_table.c.username == old_name)
            .values(username=new_name)
        )
    clear_data_caches()
    return True, f"Renamed '{old_name}' to '{new_name}'."


def admin_delete_profile(name):
    """Loescht ein Profil/Konto samt aller eigenen Daten (Admin – auch fremde)."""
    if not name:
        return False
    with get_engine().begin() as conn:
        urow = conn.execute(
            select(users_table.c.id).where(users_table.c.username == name)
        ).first()
        user_id = urow[0] if urow else None

        conn.execute(delete(sessions_table).where(sessions_table.c.name == name))
        conn.execute(delete(profiles_table).where(profiles_table.c.name == name))
        conn.execute(delete(equip_shares_table).where(equip_shares_table.c.owner == name))
        conn.execute(delete(user_prefs_table).where(user_prefs_table.c.username == name))
        conn.execute(delete(group_events_table).where(group_events_table.c.username == name))

        if user_id is not None:
            conn.execute(delete(auth_tokens_table).where(auth_tokens_table.c.user_id == user_id))
            conn.execute(delete(device_tokens_table).where(device_tokens_table.c.user_id == user_id))
            conn.execute(delete(memberships_table).where(memberships_table.c.user_id == user_id))
            owned = [
                r[0] for r in conn.execute(
                    select(groups_table.c.id).where(groups_table.c.owner_id == user_id)
                ).all()
            ]
            if owned:
                conn.execute(delete(memberships_table).where(memberships_table.c.group_id.in_(owned)))
                conn.execute(delete(group_events_table).where(group_events_table.c.group_id.in_(owned)))
                conn.execute(delete(groups_table).where(groups_table.c.id.in_(owned)))
            conn.execute(delete(users_table).where(users_table.c.id == user_id))
    clear_data_caches()
    return True


def admin_merge_profiles(from_name, to_name):
    """Verschiebt Sessions + Equipment von from_name nach to_name und loescht
    danach das Quellprofil. Gibt (verschobene_sessions, meldung) zurueck."""
    from_name = (from_name or "").strip()
    to_name = (to_name or "").strip()
    if not from_name or not to_name or from_name == to_name:
        return 0, "Choose two different profiles."
    copy_equipment(from_name, to_name)
    with get_engine().begin() as conn:
        res = conn.execute(
            update(sessions_table).where(sessions_table.c.name == from_name)
            .values(name=to_name)
        )
    moved = res.rowcount or 0
    admin_delete_profile(from_name)
    clear_data_caches()
    return moved, f"Merged '{from_name}' into '{to_name}' ({moved} sessions moved)."


# ---- Konto-/Datenlöschung (immer nur eigene Daten) ----

def count_user_sessions(name, sport=None):
    if not name:
        return 0

    with get_engine().connect() as conn:
        query = (
            select(func.count())
            .select_from(sessions_table)
            .where(sessions_table.c.name == name)
        )
        if sport:
            query = query.where(sessions_table.c.sport == sport)
        return conn.execute(query).scalar()


def delete_user_sessions(name, sport=None):
    """Löscht die hochgeladenen Sessions des Fahrers (optional nur eines Sports).
    Gibt die Anzahl zurück."""
    if not name:
        return 0

    with get_engine().begin() as conn:
        stmt = delete(sessions_table).where(sessions_table.c.name == name)
        if sport:
            stmt = stmt.where(sessions_table.c.sport == sport)
        result = conn.execute(stmt)

    clear_data_caches()

    return result.rowcount or 0


def delete_session(session_id, name):
    """Löscht eine einzelne Session – nur wenn sie dem Fahrer gehört (kein Fremdlöschen)."""
    if session_id is None or not name:
        return False

    with get_engine().begin() as conn:
        result = conn.execute(
            delete(sessions_table).where(
                (sessions_table.c.id == int(session_id))
                & (sessions_table.c.name == name)
            )
        )

    clear_data_caches()

    return (result.rowcount or 0) > 0


def delete_account(user_id, username):
    """Löscht das Konto und ALLE eigenen Daten – aber keine Fremddaten.

    Eigene Sessions, Profil, Anmelde-Tokens und Mitgliedschaften werden
    entfernt. Selbst erstellte Gruppen werden mitsamt ihrer Mitgliedschaften
    gelöscht; die Sessions der anderen Mitglieder bleiben dabei unberührt.
    """
    with get_engine().begin() as conn:
        conn.execute(delete(sessions_table).where(sessions_table.c.name == username))
        conn.execute(delete(profiles_table).where(profiles_table.c.name == username))
        conn.execute(
            delete(auth_tokens_table).where(auth_tokens_table.c.user_id == user_id)
        )
        conn.execute(
            delete(device_tokens_table).where(device_tokens_table.c.user_id == user_id)
        )
        conn.execute(
            delete(memberships_table).where(memberships_table.c.user_id == user_id)
        )

        owned = [
            row[0]
            for row in conn.execute(
                select(groups_table.c.id).where(groups_table.c.owner_id == user_id)
            ).all()
        ]

        if owned:
            conn.execute(
                delete(memberships_table).where(memberships_table.c.group_id.in_(owned))
            )
            conn.execute(delete(groups_table).where(groups_table.c.id.in_(owned)))

        conn.execute(delete(users_table).where(users_table.c.id == user_id))

    clear_data_caches()


def _cookie_manager():
    """Eine Cookie-Manager-Instanz pro Skriptlauf (oder None ohne Paket)."""
    if not COOKIES_AVAILABLE:
        return None

    try:
        return stx.CookieManager(key="surf_cookies")
    except Exception:
        return None


def _persist_auth_cookie(token, days=30):
    """Setzt das „Angemeldet bleiben"-Cookie direkt im Browser.

    Bewusst per JS auf ``window.parent.document`` (das echte App-Dokument), NICHT
    über eine Streamlit-Cookie-Komponente: ``components.html`` löst keinen zweiten
    Skriptlauf aus (anders als extra-streamlit-components), und das Cookie liegt
    auf der App-Domain – so kann es beim nächsten Laden direkt über
    ``st.context.cookies`` gelesen werden. Same-Origin-Zugriff aufs Eltern-
    Dokument funktioniert hier (siehe auch autocollapse_sidebar).
    """
    max_age = int(days * 86400)
    safe = json.dumps(token)  # sicheres JS-String-Literal

    components.html(
        f"""
        <script>
          try {{
            window.parent.document.cookie =
              "surf_auth=" + {safe} + "; max-age={max_age}; path=/; SameSite=Lax";
          }} catch (e) {{}}
        </script>
        """,
        height=0,
    )


def login_session(user, remember):
    st.session_state["user"] = {"id": user["id"], "username": user["username"]}

    if remember:
        # Das Cookie wird bewusst erst NACH dem Rerun gesetzt (siehe Gate
        # weiter unten). Ein set() direkt vor st.rerun() wird vom Browser
        # häufig verworfen, sodass das Cookie nie gespeichert würde.
        st.session_state["_pending_token"] = create_auth_token(user["id"])


def logout_session():
    # ALLE „Angemeldet bleiben"-Tokens des Users serverseitig löschen. Wichtig:
    # Wir gehen über die user_id aus der Session, NICHT über das Cookie –
    # st.context.cookies liefert auf der Streamlit Community Cloud nichts, daher
    # würde ein Cookie-basiertes Löschen den Token nie treffen (Folge: nach dem
    # Logout loggt das noch gültige Cookie sofort wieder ein). So ist das Cookie
    # nach dem Logout wertlos (user_for_token gibt None) – auch ohne es im
    # Browser zu entfernen.
    user = st.session_state.get("user")

    if user and user.get("id") is not None:
        try:
            with get_engine().begin() as conn:
                conn.execute(
                    delete(auth_tokens_table).where(
                        auth_tokens_table.c.user_id == int(user["id"])
                    )
                )
        except Exception:
            pass

    st.session_state.pop("user", None)


# =====================================================================
#  Gruppen
# =====================================================================

ALL_GROUP = "All"


@st.cache_data(ttl=3600, show_spinner=False)
def list_groups():
    with get_engine().connect() as conn:
        rows = conn.execute(select(groups_table)).mappings().all()

    return [dict(r) for r in rows]


def create_group(name, owner_id, is_private):
    name = (name or "").strip()

    if not name:
        return False, "Please enter a group name."

    if name.lower() == ALL_GROUP.lower():
        return False, "This name is reserved."

    with get_engine().begin() as conn:
        exists = conn.execute(
            select(groups_table.c.id).where(func.lower(groups_table.c.name) == name.lower())
        ).first()

        if exists:
            return False, "A group with this name already exists."

        result = conn.execute(insert(groups_table).values(
            name=name, is_private=bool(is_private), owner_id=owner_id
        ))
        group_id = result.inserted_primary_key[0]

        conn.execute(insert(memberships_table).values(
            user_id=owner_id, group_id=group_id, status="member"
        ))

    clear_data_caches()

    return True, f"Group \"{name}\" was created."


@st.cache_data(ttl=3600, show_spinner=False)
def my_memberships(user_id):
    """group_id -> status ('member' | 'pending')."""
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(memberships_table).where(memberships_table.c.user_id == user_id)
        ).mappings().all()

    return {r["group_id"]: r["status"] for r in rows}


@st.cache_data(ttl=3600, show_spinner=False)
def my_member_groups(user_id):
    """Gruppen, in denen der Nutzer bestätigtes Mitglied ist."""
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(groups_table)
            .select_from(
                groups_table.join(
                    memberships_table, groups_table.c.id == memberships_table.c.group_id
                )
            )
            .where(
                (memberships_table.c.user_id == user_id)
                & (memberships_table.c.status == "member")
            )
        ).mappings().all()

    return [dict(r) for r in rows]


def join_or_request_group(user_id, group_id):
    with get_engine().begin() as conn:
        group = conn.execute(
            select(groups_table).where(groups_table.c.id == group_id)
        ).mappings().first()

        if not group:
            return False, "Group not found."

        existing = conn.execute(
            select(memberships_table).where(
                (memberships_table.c.user_id == user_id)
                & (memberships_table.c.group_id == group_id)
            )
        ).first()

        if existing:
            return False, "You are already part of this group."

        status = "pending" if group["is_private"] else "member"

        conn.execute(insert(memberships_table).values(
            user_id=user_id, group_id=group_id, status=status
        ))

    clear_data_caches()

    if status == "pending":
        return True, "Join requested – the owner has to approve you."

    return True, "You have joined the group."


def leave_group(user_id, group_id):
    with get_engine().begin() as conn:
        conn.execute(delete(memberships_table).where(
            (memberships_table.c.user_id == user_id)
            & (memberships_table.c.group_id == group_id)
        ))

    clear_data_caches()


def group_membership_count(group_id):
    """Anzahl aller Zuordnungen (bestätigte Mitglieder + offene Anfragen)."""
    with get_engine().connect() as conn:
        return conn.execute(
            select(func.count())
            .select_from(memberships_table)
            .where(memberships_table.c.group_id == group_id)
        ).scalar()


def delete_group(group_id, owner_id):
    """Löscht eine Gruppe – nur durch den Ersteller und nur, wenn er allein drin ist.

    Die Prüfung erfolgt innerhalb der Transaktion, damit niemand fremde Gruppen
    oder eine Gruppe mit weiteren Mitgliedern löschen kann. (ok, message).
    """
    with get_engine().begin() as conn:
        group = conn.execute(
            select(groups_table).where(groups_table.c.id == group_id)
        ).mappings().first()

        if not group:
            return False, "Group not found."

        if group["owner_id"] != owner_id:
            return False, "Only the owner can delete the group."

        count = conn.execute(
            select(func.count())
            .select_from(memberships_table)
            .where(memberships_table.c.group_id == group_id)
        ).scalar()

        if count and count > 1:
            return False, (
                "You can only delete the group when you are the only one in it "
                "(remove other members or pending requests first)."
            )

        conn.execute(
            delete(memberships_table).where(memberships_table.c.group_id == group_id)
        )
        conn.execute(delete(groups_table).where(groups_table.c.id == group_id))

    clear_data_caches()

    return True, f"Group {group['name']} was deleted."


def pending_requests(group_id):
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(memberships_table.c.user_id, users_table.c.username)
            .select_from(
                memberships_table.join(
                    users_table, memberships_table.c.user_id == users_table.c.id
                )
            )
            .where(
                (memberships_table.c.group_id == group_id)
                & (memberships_table.c.status == "pending")
            )
        ).mappings().all()

    return [dict(r) for r in rows]


def set_membership_status(user_id, group_id, status):
    with get_engine().begin() as conn:
        conn.execute(
            update(memberships_table)
            .where(
                (memberships_table.c.user_id == user_id)
                & (memberships_table.c.group_id == group_id)
            )
            .values(status=status)
        )

    clear_data_caches()


def invite_user(group_id, username):
    user = get_user((username or "").strip())

    if not user:
        return False, "No user found with this username."

    with get_engine().begin() as conn:
        existing = conn.execute(
            select(memberships_table).where(
                (memberships_table.c.user_id == user["id"])
                & (memberships_table.c.group_id == group_id)
            )
        ).mappings().first()

        if existing:
            if existing["status"] == "member":
                return False, f"{user['username']} is already a member."

            conn.execute(
                update(memberships_table)
                .where(
                    (memberships_table.c.user_id == user["id"])
                    & (memberships_table.c.group_id == group_id)
                )
                .values(status="member")
            )
            message = f"{user['username']} was approved."
        else:
            conn.execute(insert(memberships_table).values(
                user_id=user["id"], group_id=group_id, status="member"
            ))
            message = f"{user['username']} was added to the group."

    clear_data_caches()

    return True, message


@st.cache_data(ttl=3600, show_spinner=False)
def group_member_names(group_id):
    """Benutzernamen aller bestätigten Mitglieder einer Gruppe."""
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(users_table.c.username)
            .select_from(
                memberships_table.join(
                    users_table, memberships_table.c.user_id == users_table.c.id
                )
            )
            .where(
                (memberships_table.c.group_id == group_id)
                & (memberships_table.c.status == "member")
            )
        ).all()

    return [r[0] for r in rows]


def clear_data_caches():
    """Leert die DB-Lese-Caches nach Schreibvorgängen.

    Bewusst NUR die Datenbank-Caches – die Wetter-Caches (Open-Meteo) bleiben
    erhalten, damit kein unnötiger erneuter Abruf (429-Risiko) ausgelöst wird.
    """
    for fn in (
        load_sessions, load_rider_sessions, load_profiles, load_spots,
        list_groups, my_memberships, my_member_groups, group_member_names,
        load_user_pref,
    ):
        try:
            fn.clear()
        except Exception:
            pass


# ---- Persönliche Einstellungen (Ranking-Filter-Preset) ----

@st.cache_data(ttl=3600, show_spinner=False)
def load_user_pref(username):
    if not username:
        return {}

    with get_engine().connect() as conn:
        row = conn.execute(
            select(user_prefs_table.c.data).where(user_prefs_table.c.username == username)
        ).first()

    if not row or not row[0]:
        return {}

    try:
        return json.loads(row[0])
    except Exception:
        return {}


def save_user_pref(username, preset):
    if not username:
        return

    data = json.dumps(preset, ensure_ascii=False)

    with get_engine().begin() as conn:
        exists = conn.execute(
            select(user_prefs_table.c.username).where(user_prefs_table.c.username == username)
        ).first()

        if exists:
            conn.execute(
                update(user_prefs_table)
                .where(user_prefs_table.c.username == username)
                .values(data=data)
            )
        else:
            conn.execute(insert(user_prefs_table).values(username=username, data=data))

    clear_data_caches()


def delete_user_pref(username):
    if not username:
        return

    with get_engine().begin() as conn:
        conn.execute(delete(user_prefs_table).where(user_prefs_table.c.username == username))

    clear_data_caches()


# =====================================================================
#  Rekorde & Bestleistungen
# =====================================================================

# Alle vier Wertungen – jeweils „höher ist besser".
RECORD_METRICS = [
    {"key": "speed_1s_kmh", "label": "Top speed 2 s", "unit": "km/h", "decimals": 2},
    {"key": "speed_30s_kmh", "label": "Top speed 30 s", "unit": "km/h", "decimals": 2},
    {"key": "longest_run_km", "label": "Longest run", "unit": "km", "decimals": 3},
    {"key": "total_distance_km", "label": "Total distance", "unit": "km", "decimals": 2},
]


def _series_max(df, key, mask=None):
    """Maximum einer Kennzahl-Spalte (optional gefiltert) oder None."""
    if df is None or df.empty or key not in df.columns:
        return None

    series = df[key] if mask is None else df.loc[mask, key]
    series = pd.to_numeric(series, errors="coerce").dropna()

    return float(series.max()) if not series.empty else None


def detect_records(entry, all_sessions, member_groups):
    """Vergleicht die neue (noch NICHT gespeicherte) Session mit dem Bestand.

    Liefert dict mit Listen: personal, spot, year, group_events.
    """
    username = str(entry.get("name") or "")
    spot = entry.get("surfspot")
    year = (entry.get("date") or "")[:4]

    df = all_sessions if all_sessions is not None else pd.DataFrame()
    have_data = not df.empty

    if have_data:
        names = df["name"].astype(str) if "name" in df.columns else None
        spots = df["surfspot"].astype(str) if "surfspot" in df.columns else None
        years = (
            pd.to_datetime(df["date"], errors="coerce", format="mixed").dt.year
            if "date" in df.columns else None
        )
    else:
        names = spots = years = None

    personal, spot_records, year_records, group_events = [], [], [], []

    for m in RECORD_METRICS:
        key, label, unit, dec = m["key"], m["label"], m["unit"], m["decimals"]
        new_val = entry.get(key)

        if new_val is None:
            continue

        new_val = float(new_val)
        base = {"label": label, "value": new_val, "unit": unit, "decimals": dec}

        # --- Persönlicher Rekord (alle eigenen Sessions) ---
        prev_personal = (
            _series_max(df, key, names == username) if have_data and names is not None else None
        )
        is_personal = prev_personal is None or new_val > prev_personal

        if is_personal:
            personal.append({**base, "previous": prev_personal})

        # --- Jahresbestleistung (nur wenn KEIN All-time-Rekord) ---
        if year and have_data and names is not None and years is not None:
            mask_year = (names == username) & (years == int(year) if year.isdigit() else False)
            prev_year = _series_max(df, key, mask_year)
            is_year_best = prev_year is None or new_val > prev_year

            if is_year_best and not is_personal:
                year_records.append({**base, "previous": prev_year, "year": year})

        # --- Spotrekord (alle Fahrer am selben Spot) ---
        if spot and have_data and spots is not None:
            prev_spot = _series_max(df, key, spots == str(spot))

            if prev_spot is None or new_val > prev_spot:
                spot_records.append({**base, "previous": prev_spot, "spot": str(spot)})

        # --- Gruppen: Rang & Rekord ---
        for g in member_groups:
            member_names = set(group_member_names(g["id"]))
            member_names.add(username)

            # Nur sinnvolle Gruppen (mind. 2 Fahrer mit Sessions) – sonst Spam.
            if have_data and names is not None:
                mask_grp = names.isin(member_names)
                distinct_riders = df.loc[mask_grp, "name"].astype(str).nunique() if mask_grp.any() else 0
            else:
                mask_grp = None
                distinct_riders = 0

            if distinct_riders < 2:
                continue

            grp_vals = pd.to_numeric(df.loc[mask_grp, key], errors="coerce").dropna()
            prev_group = float(grp_vals.max()) if not grp_vals.empty else None
            rank = int((grp_vals > new_val).sum()) + 1
            is_group_record = prev_group is None or new_val > prev_group

            if rank <= 3:
                group_events.append({
                    **base,
                    "group_id": g["id"],
                    "group_name": g["name"],
                    "rank": rank,
                    "is_record": is_group_record,
                })

    return {
        "personal": personal,
        "spot": spot_records,
        "year": year_records,
        "group_events": group_events,
    }


def record_group_events(username, group_events):
    """Persistiert Top-3/Rekord-Ereignisse, damit Mitglieder sie später sehen."""
    if not group_events:
        return

    rows = []

    for e in group_events:
        value_str = f"{e['value']:.{e['decimals']}f} {e['unit']}"

        if e.get("is_record"):
            message = (
                f"🏆 {username} set a new group record: "
                f"{e['label']} {value_str} (group {e['group_name']})."
            )
        else:
            message = (
                f"🏅 {username} is in rank {e['rank']} for {e['label']} "
                f"in group {e['group_name']} ({value_str})."
            )

        rows.append({
            "group_id": e["group_id"],
            "username": username,
            "metric": e["label"],
            "rank": e["rank"],
            "value": e["value"],
            "message": message,
        })

    with get_engine().begin() as conn:
        conn.execute(insert(group_events_table), rows)


@st.cache_data(ttl=60, show_spinner=False)
def unseen_group_events(user_id, username):
    """Noch ungelesene Gruppen-Ereignisse (nicht die eigenen) für das Banner.
    Kurz gecacht (60 s), damit nicht jeder Rerun 2 DB-Abfragen ausloest."""
    group_ids = [g["id"] for g in my_member_groups(user_id)]

    if not group_ids:
        return []

    with get_engine().connect() as conn:
        seen_row = conn.execute(
            select(event_reads_table.c.last_seen_event_id)
            .where(event_reads_table.c.user_id == user_id)
        ).first()
        last_seen = seen_row[0] if seen_row else 0

        rows = conn.execute(
            select(group_events_table)
            .where(
                group_events_table.c.group_id.in_(group_ids),
                group_events_table.c.id > last_seen,
                group_events_table.c.username != username,
            )
            .order_by(group_events_table.c.id.desc())
            .limit(20)
        ).mappings().all()

    return [dict(r) for r in rows]


def mark_group_events_seen(user_id):
    with get_engine().begin() as conn:
        max_id = conn.execute(select(func.max(group_events_table.c.id))).scalar() or 0

        exists = conn.execute(
            select(event_reads_table.c.user_id).where(event_reads_table.c.user_id == user_id)
        ).first()

        if exists:
            conn.execute(
                update(event_reads_table)
                .where(event_reads_table.c.user_id == user_id)
                .values(last_seen_event_id=max_id)
            )
        else:
            conn.execute(insert(event_reads_table).values(
                user_id=user_id, last_seen_event_id=max_id
            ))

    # Cache leeren, damit der Banner nach "gelesen" sofort verschwindet.
    unseen_group_events.clear()


# Pflichtfelder, damit eine Session im Ranking/in den Personal Bests zählt.
RANKING_REQUIRED = ["surfspot", "board", "sail"]


def complete_sessions(df):
    """Nur Sessions, die fürs Ranking vollständig sind: Spot + Board + Segel/
    Kite gesetzt. Unvollständige (z.B. frisch von der Uhr, ohne Equipment)
    bleiben aus Ranking & Personal Bests draußen, bis sie nachgepflegt wurden."""
    if df is None or df.empty:
        return df
    if not all(col in df.columns for col in RANKING_REQUIRED):
        return df.iloc[0:0]
    mask = pd.Series(True, index=df.index)
    for col in RANKING_REQUIRED:
        s = df[col].astype(str).str.strip()
        mask &= s.ne("") & ~s.str.lower().isin(["none", "nan", "null"])
    return df[mask]


def personal_bests(name, spot=None, year=None):
    """Beste Werte je Metrik für einen Fahrer, optional nach Spot/Jahr gefiltert."""
    df = load_rider_sessions(name, active_sport())
    df = complete_sessions(df)

    if df.empty:
        return []

    if spot and spot != "All" and "surfspot" in df.columns:
        df = df[df["surfspot"].astype(str) == spot]

    if year and year != "All" and "date" in df.columns:
        df = df[pd.to_datetime(df["date"], errors="coerce", format="mixed").dt.year == int(year)]

    return [
        {
            "label": m["label"],
            "value": _series_max(df, m["key"]),
            "unit": m["unit"],
            "decimals": m["decimals"],
        }
        for m in RECORD_METRICS
    ]


def personal_best_table(df, spot="All", year="All", board="All", max_bft=None, limit=10):
    """Bis zu `limit` beste Sessions eines Fahrers als Speed-Tabelle.

    Sortiert nach Topspeed 1 s (absteigend). Optional gefiltert nach Spot, Jahr,
    Board und maximaler Windstärke (`max_bft` = Beaufort-Obergrenze, z.B. 5 für
    "höchstens 5 Bft" – Sessions ohne Winddaten fallen dann raus). Gibt
    (Anzeige-DataFrame, Filter-Beschriftung) zurück; der DataFrame ist leer, wenn
    keine passenden Sessions vorhanden sind.
    """
    data = complete_sessions(df.copy())

    if spot and spot != "All" and "surfspot" in data.columns:
        data = data[data["surfspot"].astype(str) == spot]

    if year and year != "All" and "date" in data.columns:
        data = data[pd.to_datetime(data["date"], errors="coerce", format="mixed").dt.year == int(year)]

    if board and board != "All" and "board" in data.columns:
        data = data[data["board"].astype(str) == board]

    if max_bft is not None and "wind_kmh" in data.columns:
        wind = pd.to_numeric(data["wind_kmh"], errors="coerce")
        bft = wind.apply(lambda v: kmh_to_beaufort(v) if pd.notna(v) else float("nan"))
        data = data[bft <= max_bft]  # NaN <= x ist False -> Sessions ohne Wind raus

    # Gesamtzahl der Sessions in dieser Filter-Auswahl (vor Top-10-Begrenzung).
    total = len(data)

    if "speed_1s_kmh" not in data.columns:
        return pd.DataFrame(), "", total

    data = data.copy()
    data["_s1"] = pd.to_numeric(data["speed_1s_kmh"], errors="coerce")
    data = data.dropna(subset=["_s1"]).sort_values("_s1", ascending=False).head(limit)

    if data.empty:
        return pd.DataFrame(), "", total

    out = pd.DataFrame()
    out["Rank"] = range(1, len(data) + 1)

    if "date" in data.columns:
        out["Date"] = [
            "" if pd.isna(d) else d.strftime("%Y-%m-%d")
            for d in pd.to_datetime(data["date"], errors="coerce", format="mixed")
        ]
    if "surfspot" in data.columns:
        out["Spot"] = data["surfspot"].astype(str).values
    if "board" in data.columns:
        out["Board"] = data["board"].astype(str).values

    if "wind_kmh" in data.columns:
        wind = pd.to_numeric(data["wind_kmh"], errors="coerce")
        out["Wind (km/h)"] = wind.round(0).values
        out["Bft"] = [
            None if pd.isna(v) else kmh_to_beaufort(v) for v in wind
        ]

    out["Top 2 s (km/h)"] = data["_s1"].round(2).values
    out["2 s (kn)"] = (data["_s1"] / 1.852).round(2).values

    if "speed_30s_kmh" in data.columns:
        s30 = pd.to_numeric(data["speed_30s_kmh"], errors="coerce")
        out["Top 30 s (km/h)"] = s30.round(2).values
        out["30 s (kn)"] = (s30 / 1.852).round(2).values

    parts = []
    if spot and spot != "All":
        parts.append(f"Spot: {spot}")
    if year and year != "All":
        parts.append(f"Year: {year}")
    if board and board != "All":
        parts.append(f"Board: {board}")
    if max_bft is not None:
        parts.append(f"≤ {max_bft} Bft")
    caption = " · ".join(parts) if parts else "All spots · all years"

    return out, caption, total


def render_personal_best_filter(name):
    """Bestleistungs-Filter (Spot/Jahr/Board/Wind) – wird im Konto-Bereich der
    Sidebar gerendert (unter „Konto & Daten löschen", einfacher zu finden).

    Gibt (Anzeige-DataFrame, Caption) zurück; die eigentliche Tabelle wird separat
    im Hauptfenster über render_personal_best_table() angezeigt. Bewusst KEIN
    Fragment mehr (eine Filteränderung löst einen normalen Rerun aus) – die
    kleine Bestleistungs-Tabelle macht das unkritisch, und so lässt sich der
    Filter frei im Konto-Bereich platzieren.
    """
    pb_df = load_rider_sessions(name, active_sport())

    with st.expander("🏅 Personal Bests", expanded=False):
        if pb_df.empty:
            st.info("No sessions yet – upload a FIT file.")
            return None, "", 0

        pb_spots = sorted(
            {str(s) for s in pb_df["surfspot"].dropna().astype(str) if str(s).strip()}
            if "surfspot" in pb_df.columns else set()
        )
        pb_years = sorted(
            {int(y) for y in pd.to_datetime(pb_df["date"], errors="coerce", format="mixed").dt.year.dropna()}
            if "date" in pb_df.columns else set(),
            reverse=True,
        )
        pb_boards = sorted(
            {str(b) for b in pb_df["board"].dropna().astype(str) if str(b).strip()}
            if "board" in pb_df.columns else set()
        )

        cpb1, cpb2 = st.columns(2)
        spot_pb = cpb1.selectbox("Spot", ["All"] + pb_spots, key=f"pb_spot_{name}")
        year_pb = cpb2.selectbox(
            "Year", ["All"] + [str(y) for y in pb_years], key=f"pb_year_{name}"
        )

        cpb3, cpb4 = st.columns(2)
        board_pb = cpb3.selectbox(
            "Board", ["All"] + pb_boards, key=f"pb_board_{name}"
        )
        wind_pb = cpb4.selectbox(
            "Max. wind",
            ["All"] + [f"≤ {b} Bft" for b in range(2, 11)],
            key=f"pb_wind_{name}",
            help="Shows only sessions up to this wind force – e.g. \"≤ 5 Bft\" "
                 "for: How fast was I at no more than 5 Beaufort?",
        )
        max_bft_pb = None if wind_pb == "All" else int(wind_pb.split()[1])

        pb_table, pb_table_caption, pb_total = personal_best_table(
            pb_df, spot_pb, year_pb, board_pb, max_bft_pb
        )
        st.caption("➡️ Your top 10 speed table is shown in the main window.")

    return pb_table, pb_table_caption, pb_total


def render_personal_best_table(pb_table, pb_table_caption, total=0):
    """Zeigt die Top-10-Bestleistungs-Tabelle im Hauptfenster (gefiltert über
    den Filter im Konto-Bereich). `total` = Anzahl Sessions in der Auswahl."""
    if pb_table is None:
        return

    st.markdown("## 🏅 Personal Bests")
    st.caption(f"📊 {total} session{'' if total == 1 else 's'} in this selection")

    if pb_table.empty:
        st.info("No sessions for this selection.")
    else:
        if pb_table_caption:
            st.caption(pb_table_caption)

        st.dataframe(pb_table, width="stretch", hide_index=True, height=df_height(len(pb_table)))

    st.markdown("---")


def render_session_history(name):
    """Verlauf der eigenen Sessions (Tabelle + Detail-Auswahl) – wird im Konto-
    Bereich der Sidebar gerendert (unter „Meine Bestleistungen"). Gibt den für
    die Detailansicht im Hauptfenster gewählten Datensatz zurück (oder None)."""
    selected = None
    gear_label = SPORT_META[active_sport()]["gear_label"]

    with st.expander("📅 View my sessions", expanded=False):
        history = load_rider_sessions(name, active_sport())

        if history.empty:
            st.info("No saved sessions for this rider yet.")
            return None

        valid_dates = history["date"].dropna() if "date" in history.columns else pd.Series(dtype="datetime64[ns]")

        if not valid_dates.empty:
            min_date = valid_dates.min().date()
            max_date = valid_dates.max().date()

            date_range = st.date_input(
                "Filter by date range",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date,
                key=f"history_dates_{name}",
            )

            if isinstance(date_range, (tuple, list)) and len(date_range) == 2:
                start, end = date_range
                history = history[
                    (history["date"].dt.date >= start)
                    & (history["date"].dt.date <= end)
                ]

        # Vollständige (datums-gefilterte) Liste behalten – die Tabelle zeigt die
        # letzten 10, im Dropdown sind aber alle wählbar.
        history_full = history
        history = history_full.head(10)

        if history_full.empty:
            st.info("No sessions in the selected date range.")
        else:
            columns = [
                "date", "surfspot", "board", "sail",
                "speed_30s_kmh", "speed_1s_kmh",
                "longest_run_km", "total_distance_km",
            ]
            show_history = history[[c for c in columns if c in history.columns]].copy()

            if "date" in show_history.columns:
                show_history["date"] = show_history["date"].dt.strftime("%Y-%m-%d")

            show_history = show_history.rename(columns={
                "date": "Date",
                "surfspot": "Surf spot",
                "board": "Board",
                "sail": gear_label,
                "speed_30s_kmh": "30s km/h",
                "speed_1s_kmh": "2s km/h",
                "longest_run_km": "Run km",
                "total_distance_km": "Distance km",
            })

            st.dataframe(
                show_history,
                width="stretch",
                hide_index=True,
                height=df_height(len(show_history)),
            )

            if len(history_full) > len(history):
                st.caption(
                    f"Table shows the last {len(history)} – all {len(history_full)} "
                    f"are selectable in the dropdown. "
                    "(Delete under \"Account & delete data\".)"
                )

            # Dropdown über ALLE (gefilterten) Sessions, nicht nur die 10 in der
            # Tabelle. Label mit Speed zur besseren Unterscheidung.
            record_by_label = {}

            for i, (_, row) in enumerate(history_full.iterrows()):
                if "date" in history_full.columns and pd.notna(row["date"]):
                    date_str = row["date"].strftime("%Y-%m-%d")
                else:
                    date_str = "?"

                spot_label = str(row.get("surfspot", "") or "")
                speed_val = row.get("speed_1s_kmh")
                speed_label = (
                    "" if speed_val is None or pd.isna(speed_val)
                    else f" · {float(speed_val):.1f} km/h"
                )
                label = (
                    f"{date_str} · {spot_label}{speed_label}".strip(" ·")
                    or f"Session {i + 1}"
                )

                if label in record_by_label:
                    label = f"{label} ({i + 1})"

                record_by_label[label] = row

            chosen_label = st.selectbox(
                "Show session details on the right",
                ["—"] + list(record_by_label.keys()),
                key=f"history_pick_{name}",
            )

            if chosen_label != "—":
                selected = record_by_label[chosen_label]

    return selected


def _http_get_json(url, timeout, retries=2):
    """GET einer JSON-API mit explizitem User-Agent und Retry.

    Wichtig fürs Cloud-Deployment (z.B. Streamlit Community Cloud): Der
    Default-UA "Python-urllib/x.y" wird von manchen CDNs/WAFs ausgehend von
    geteilten Cloud-IPs blockiert, lokal aber durchgelassen. Zudem teilen sich
    viele Apps dieselbe Ausgangs-IP, sodass Open-Meteo sporadisch mit HTTP 429
    antwortet – ein kurzer Retry fängt solche Minuten-Spitzen ab. Wirft, wenn
    alle Versuche scheitern; die Aufrufer fangen das ab.
    """
    request = Request(url, headers={"User-Agent": "WindsurfSpeedChallenge/1.0"})
    last_error = None

    for attempt in range(retries + 1):
        try:
            with urlopen(request, timeout=timeout) as response:
                return json.load(response)
        except HTTPError as e:
            # 4xx (v.a. 429 Rate-Limit) lösen sich durch sofortiges Wiederholen
            # nicht – daher gleich aufgeben, statt die Seite mit Warte-Retries
            # zu blockieren. Nur 5xx/Netzwerkfehler werden erneut versucht.
            if 400 <= e.code < 500:
                raise
            last_error = e

            if attempt < retries:
                time.sleep(1.5 * (attempt + 1))
        except Exception as e:
            last_error = e

            if attempt < retries:
                time.sleep(1.5 * (attempt + 1))

    raise last_error


@st.cache_data(show_spinner=False)
def geocode_spot(name):
    if not name:
        return None

    params = {"name": name, "count": 1, "language": "de", "format": "json"}

    try:
        url = _open_meteo_url("geocoding-api", "/v1/search", params)
        data = _http_get_json(url, timeout=10)
    except Exception:
        return None

    results = data.get("results") or []

    if not results:
        return None

    return {"lat": results[0]["latitude"], "lon": results[0]["longitude"]}


def resolve_spot_coords(spot):
    spots = load_spots()

    if spot in spots:
        return spots[spot]["lat"], spots[spot]["lon"]

    geo = geocode_spot(spot)

    if geo:
        update_spot_coords(spot, geo["lat"], geo["lon"])
        return geo["lat"], geo["lon"]

    return None, None


def find_watch_fit_files(extra_folder=None):
    folders = []

    for letter in "DEFGHIJKLMNOPQRSTUVWXYZ":
        activity = f"{letter}:\\GARMIN\\Activity"

        if os.path.isdir(activity):
            folders.append(activity)

    if extra_folder and os.path.isdir(extra_folder):
        folders.append(extra_folder)

    files = []

    for folder in folders:
        files.extend(glob.glob(os.path.join(folder, "*.fit")))

    files = sorted(set(files), key=os.path.getmtime, reverse=True)

    return files


# Zielordner für vom MTP-Gerät kopierte FIT-Dateien
MTP_TEMP_DIR = os.path.join(tempfile.gettempdir(), "surfapp_mtp")
SSF_DRIVES = 17  # "Dieser PC" im Shell-Namespace


def _mtp_subfolder(folder, name):
    for item in folder.Items():
        if item.IsFolder and item.Name == name:
            return item.GetFolder

    return None


def _mtp_activity_folder(device):
    """Navigiert device -> [Primary] -> GARMIN -> Activity."""
    dev_folder = device.GetFolder
    primary = _mtp_subfolder(dev_folder, "Primary") or dev_folder
    garmin = _mtp_subfolder(primary, "GARMIN")

    if garmin is None:
        return None

    return _mtp_subfolder(garmin, "Activity")


def find_mtp_fit_files():
    """FIT-Dateien auf per MTP angeschlossenen Garmin-Uhren (kein Laufwerk).

    Gibt eine Liste von Dicts mit device, name und modified zurück.
    Bei Fehlern oder ohne pywin32 schlicht eine leere Liste.
    """
    if not MTP_AVAILABLE:
        return []

    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    entries = []

    try:
        shell = win32com.client.Dispatch("Shell.Application")
        pc = shell.Namespace(SSF_DRIVES)

        for device in pc.Items():
            if not device.IsFolder:
                continue

            # Laufwerksbuchstaben (C:, D: …) überspringen – die deckt die
            # normale Ordner-Erkennung ab; hier nur echte MTP-Geräte.
            if device.Path and len(device.Path) <= 3 and device.Path[1:2] == ":":
                continue

            try:
                activity = _mtp_activity_folder(device)
            except Exception:
                continue

            if activity is None:
                continue

            for f in activity.Items():
                if f.IsFolder:
                    continue

                is_fit = f.Name.lower().endswith(".fit") or "fit" in (f.Type or "").lower()

                if not is_fit:
                    continue

                entries.append({
                    "device": device.Name,
                    "name": f.Name,
                    "modified": str(f.ExtendedProperty("System.DateModified"))[:19],
                })
    except Exception:
        entries = []
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    return entries


def copy_mtp_file(device_name, file_name, dest_dir=MTP_TEMP_DIR):
    """Kopiert eine FIT-Datei vom MTP-Gerät in einen lokalen Ordner.

    Gibt den lokalen Pfad zurück (oder None bei Fehler). Bereits kopierte
    Dateien werden wiederverwendet (idempotent über Streamlit-Reruns).
    """
    if not MTP_AVAILABLE:
        return None

    os.makedirs(dest_dir, exist_ok=True)

    expected_name = file_name if file_name.lower().endswith(".fit") else f"{file_name}.fit"
    expected_path = os.path.join(dest_dir, expected_name)

    if os.path.exists(expected_path) and os.path.getsize(expected_path) > 0:
        return expected_path

    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    try:
        shell = win32com.client.Dispatch("Shell.Application")
        pc = shell.Namespace(SSF_DRIVES)

        device = None

        for d in pc.Items():
            if d.Name == device_name:
                device = d
                break

        if device is None:
            return None

        activity = _mtp_activity_folder(device)

        if activity is None:
            return None

        src_item = None

        for f in activity.Items():
            if f.Name == file_name:
                src_item = f
                break

        if src_item is None:
            return None

        before = set(os.listdir(dest_dir))
        dest_folder = shell.Namespace(dest_dir)
        dest_folder.CopyHere(src_item, 16)  # 16 = ohne Rückfrage-Dialoge

        import time

        for _ in range(150):
            new_files = [
                n for n in os.listdir(dest_dir)
                if n not in before and n.lower().endswith(".fit")
            ]

            if new_files:
                path = os.path.join(dest_dir, new_files[0])

                if os.path.getsize(path) > 0:
                    return path

            time.sleep(0.1)

        return None
    except Exception:
        return None
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


@st.cache_data(show_spinner=False)
def _fetch_weather(lat, lon, when_iso):
    date = when_iso[:10]
    target_hour = when_iso[:13]

    common = {
        "latitude": round(float(lat), 4),
        "longitude": round(float(lon), 4),
        "hourly": "temperature_2m,precipitation,weather_code,wind_speed_10m,wind_gusts_10m,wind_direction_10m",
        "wind_speed_unit": "kmh",
        "timezone": "GMT",
        "start_date": date,
        "end_date": date,
    }

    services = (
        ("archive-api", "/v1/archive"),
        ("api", "/v1/forecast"),
    )

    any_response = False

    for service, path in services:
        try:
            data = _http_get_json(_open_meteo_url(service, path, common), timeout=25)
        except Exception:
            continue

        any_response = True
        hourly = data.get("hourly") or {}
        times = hourly.get("time") or []

        for i, t in enumerate(times):
            if t[:13] != target_hour:
                continue

            def value(key):
                series = hourly.get(key) or []
                return series[i] if i < len(series) else None

            temp = value("temperature_2m")

            if temp is None:
                continue

            return {
                "temp": temp,
                "precip": value("precipitation"),
                "code": value("weather_code"),
                "wind": value("wind_speed_10m"),
                "gust": value("wind_gusts_10m"),
                "dir": value("wind_direction_10m"),
            }

    # Kein Endpunkt erreichbar -> Fehler werfen, damit st.cache_data das
    # Ergebnis NICHT zwischenspeichert und beim nächsten Mal neu versucht.
    if not any_response:
        raise RuntimeError("Wetterdienst nicht erreichbar")

    # Endpunkte haben geantwortet, aber kein Treffer -> echtes "keine Daten"
    return None


def get_weather(lat, lon, when_iso):
    """Wetter zur Session-Zeit (UTC) von Open-Meteo holen.

    when_iso: "YYYY-MM-DDTHH:MM" in UTC (aus dem FIT-Zeitstempel).
    Gibt ein Dict (Temperatur, Wind, Böen, Niederschlag, Wettercode) zurück
    oder None, wenn keine Daten verfügbar sind bzw. der Dienst gerade streikt.
    """
    try:
        return _fetch_weather(lat, lon, when_iso)
    except Exception:
        logging.exception("Historischer Wetter-Abruf fehlgeschlagen")
        return None


def _open_meteo_key():
    """Optionaler Open-Meteo-API-Key aus Streamlit-Secrets oder Umgebung.

    Mit Key laufen die Aufrufe über den Kunden-Endpunkt und sind NICHT mehr
    vom geteilten IP-Limit der Cloud betroffen (der eigentliche Grund für die
    429-Fehler). Ohne Key bleibt alles wie bisher (kostenloser Endpunkt).
    """
    try:
        key = st.secrets.get("OPENMETEO_API_KEY")
        if key:
            return key
    except Exception:
        pass

    return os.environ.get("OPENMETEO_API_KEY")


def _open_meteo_url(service, path, params):
    """Baut eine Open-Meteo-URL und schaltet mit API-Key auf den Kunden-Endpunkt.

    service: "api" (Forecast), "archive-api" (Historie), "geocoding-api".
    Mit Key läuft der Aufruf über "customer-<service>.open-meteo.com" und ist
    nicht mehr vom geteilten IP-Limit (429) der Cloud betroffen.
    """
    params = dict(params)
    key = _open_meteo_key()

    if key:
        params["apikey"] = key
        host = f"https://customer-{service}.open-meteo.com"
    else:
        host = f"https://{service}.open-meteo.com"

    return f"{host}{path}?{urlencode(params)}"


@st.cache_data(show_spinner=False, ttl=1800)
def _fetch_forecast(lat, lon):
    params = {
        "latitude": round(float(lat), 4),
        "longitude": round(float(lon), 4),
        "current": "temperature_2m,precipitation,weather_code,wind_speed_10m,wind_gusts_10m,wind_direction_10m",
        "hourly": "temperature_2m,precipitation,weather_code,wind_speed_10m,wind_gusts_10m,wind_direction_10m",
        "wind_speed_unit": "kmh",
        "timezone": "auto",
        "forecast_days": 4,
    }

    url = _open_meteo_url("api", "/v1/forecast", params)
    # Großzügiger Timeout: Open-Meteo ist zeitweise langsam (>20 s), liefert
    # aber. Erfolgreiche Antworten werden 30 Min gecacht.
    return _http_get_json(url, timeout=30)


@st.cache_resource(show_spinner=False)
def _forecast_fallback_store():
    """Letzter erfolgreicher Forecast je Spot (überlebt Reruns/Sessions).

    Dient als Notfall-Anzeige, wenn ein frischer Abruf gerade an einem 429
    scheitert – dann sehen Nutzer den letzten Stand statt einer leeren Warnung.
    """
    return {}


def get_forecast(lat, lon):
    """Aktuelles Wetter + Stundenvorhersage von Open-Meteo (Forecast-API).

    Fehler werden NICHT gecacht – nur erfolgreiche Antworten landen im Cache,
    damit eine kurze Dienststörung nicht 30 Minuten hängen bleibt.
    """
    store = _forecast_fallback_store()
    key = (round(float(lat), 4), round(float(lon), 4))

    try:
        result = _fetch_forecast(lat, lon)
        store[key] = result
        st.session_state.pop("_weather_error", None)
        st.session_state["_weather_stale"] = False
        return result
    except Exception as e:
        logging.exception("Forecast-Abruf fehlgeschlagen")
        st.session_state["_weather_error"] = f"{type(e).__name__}: {e}"

        # Fallback: letzter erfolgreicher Stand für diesen Spot (falls vorhanden)
        stale = store.get(key)
        st.session_state["_weather_stale"] = stale is not None
        return stale


# HTML/JS-Panel, das das Wetter DIREKT IM BROWSER des Besuchers von Open-Meteo
# holt (fetch). Dadurch zählt der Abruf auf die IP/das Tageslimit des Besuchers
# – nicht auf die geteilte Streamlit-Cloud-IP, die sonst 429-Fehler verursacht.
_WEATHER_HTML = """
<!DOCTYPE html><html><head><meta charset="utf-8"><style>
  /* KEIN color-scheme:dark – das zwang den Browser, hinter dem transparenten
     Body eine dunkle Fläche zu zeichnen (der „schwarze Kasten"). Ohne das und
     mit transparentem html/body ist der Iframe wirklich durchsichtig, sodass
     der Seitenhintergrund durchscheint. Textfarbe setzen wir selbst hell. */
  html, body { background: transparent !important; }
  body { margin:0; color:#eaf4ff;
         font-family: "Source Sans Pro", system-ui, -apple-system, sans-serif; }
  /* Kein dunkler Wrapper: die aktuellen Karten schweben direkt auf dem
     Hintergrund (wie der Bereich „Aktuell & Vorhersage" darüber). */
  .current { background:transparent; border:none; padding:0; margin:0 0 14px 0; }
  .now { display:flex; flex-wrap:wrap; gap:10px; margin:0; }
  .card { background:rgba(255,255,255,.10); border:1px solid rgba(255,255,255,.22);
          border-radius:14px; padding:9px 14px; min-width:120px; flex:1 1 130px; }
  .card .lbl { font-size:12px; opacity:.8; }
  .card .val { font-size:20px; font-weight:800; }
  .card .sub { font-size:12px; opacity:.85; min-height:1em; }
  .dir { font-size:13px; opacity:.92; margin:10px 0 0; }
  h4 { margin:6px 0 8px; font-weight:800; }
  table { width:100%; border-collapse:collapse; font-size:13px;
          background:rgba(8,28,52,.35); border-radius:12px; overflow:hidden; }
  th,td { padding:6px 8px; text-align:left;
          border-bottom:1px solid rgba(255,255,255,.12); white-space:nowrap; }
  th { font-weight:700; opacity:.85; }
  tr.dsep td { border-top:2px solid rgba(255,255,255,.30); }
  .warn { background:rgba(255,200,0,.18); border:1px solid rgba(255,200,0,.4);
          border-radius:12px; padding:12px 14px; }
  .muted { opacity:.7; font-size:12px; margin-top:6px; }
</style></head><body>
<div id="w">Loading weather…</div>
<script>
const LAT=__LAT__, LON=__LON__;
const WMO={0:["☀️","Clear"],1:["🌤️","Mainly clear"],2:["⛅","Partly cloudy"],3:["☁️","Overcast"],45:["🌫️","Fog"],48:["🌫️","Rime fog"],51:["🌦️","Light drizzle"],53:["🌦️","Drizzle"],55:["🌧️","Heavy drizzle"],61:["🌦️","Light rain"],63:["🌧️","Rain"],65:["🌧️","Heavy rain"],66:["🌧️","Freezing rain"],67:["🌧️","Heavy freezing rain"],71:["🌨️","Light snowfall"],73:["🌨️","Snowfall"],75:["❄️","Heavy snowfall"],77:["🌨️","Snow grains"],80:["🌦️","Light showers"],81:["🌧️","Showers"],82:["⛈️","Violent showers"],85:["🌨️","Snow showers"],86:["❄️","Heavy snow showers"],95:["⛈️","Thunderstorm"],96:["⛈️","Thunderstorm with hail"],99:["⛈️","Severe thunderstorm with hail"]};
const COMPASS=["N","NNE","NE","ENE","E","ESE","SE","SSE","S","SSW","SW","WSW","W","WNW","NW","NNW"];
const ARROWS=["⬆️","↗️","➡️","↘️","⬇️","↙️","⬅️","↖️"];
const WD=["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
function comp(d){ if(d==null) return ""; return COMPASS[Math.round(d/22.5)%16]; }
function arr(d){ if(d==null) return ""; return ARROWS[Math.round(((d+180)%360)/45)%8]; }
function bft(k){ if(k==null) return ""; const L=[1,5,11,19,28,38,49,61,74,88,102,117]; for(let i=0;i<L.length;i++){ if(k<=L[i]) return i; } return 12; }
function wc(c){ return WMO[c]||["❓","Unknown"]; }
function f1(x){ return (x==null)?"–":(Math.round(x*10)/10).toFixed(1); }
function r0(x){ return (x==null)?"–":Math.round(x); }
const URL="https://api.open-meteo.com/v1/forecast?latitude="+LAT+"&longitude="+LON+"&current=temperature_2m,precipitation,weather_code,wind_speed_10m,wind_gusts_10m,wind_direction_10m&hourly=temperature_2m,precipitation,weather_code,wind_speed_10m,wind_gusts_10m,wind_direction_10m&wind_speed_unit=kmh&timezone=auto&forecast_days=4";
fetch(URL).then(function(r){ if(!r.ok) throw new Error("HTTP "+r.status); return r.json(); })
  .then(render)
  .catch(function(e){ document.getElementById("w").innerHTML='<div class="warn">Weather service (Open-Meteo) is currently unavailable – please try again later.</div><div class="muted">Technical reason: '+e.message+'</div>'; });
function render(data){
  const c=data.current||{}; const cc=wc(c.weather_code); const rain=(c.precipitation||0);
  let out='<div class="current"><div class="now">';
  out+='<div class="card"><div class="lbl">Wind</div><div class="val">'+r0(c.wind_speed_10m)+' km/h</div><div class="sub">'+(c.wind_gusts_10m==null?'':'Gusts '+r0(c.wind_gusts_10m)+' km/h')+'</div></div>';
  out+='<div class="card"><div class="lbl">Temperature</div><div class="val">'+f1(c.temperature_2m)+' °C</div><div class="sub"></div></div>';
  out+='<div class="card"><div class="lbl">Rain</div><div class="val">'+(rain>0?'Yes':'No')+'</div><div class="sub">'+(rain>0?f1(rain)+' mm':'')+'</div></div>';
  out+='<div class="card"><div class="lbl">Condition</div><div class="val">'+cc[0]+'</div><div class="sub">'+cc[1]+'</div></div>';
  out+='</div>';
  if(c.wind_direction_10m!=null){ out+='<div class="dir">🧭 Wind from <b>'+comp(c.wind_direction_10m)+'</b> '+arr(c.wind_direction_10m)+' ('+r0(c.wind_direction_10m)+'°)</div>'; }
  out+='</div>';
  const h=data.hourly||{}; const t=h.time||[]; const nowt=c.time||"";
  let start=0; for(let i=0;i<t.length;i++){ if(t[i]>=nowt){ start=i; break; } }
  const dayH={9:1,12:1,15:1,18:1,21:1}; const seen=[]; let rows=""; let lastDay=null;
  for(let i=start;i<t.length;i++){
    const hh=parseInt(t[i].slice(11,13),10); if(!dayH[hh]) continue;
    const day=t[i].slice(0,10);
    if(seen.indexOf(day)<0){ if(seen.length>=3) break; seen.push(day); }
    const wci=wc(h.weather_code[i]); const pr=(h.precipitation[i]||0);
    const wknd=WD[new Date(day+"T00:00").getDay()];
    const sep=(lastDay!==null && day!==lastDay)?' class="dsep"':''; lastDay=day;
    rows+='<tr'+sep+'><td>'+wknd+'</td><td>'+t[i].slice(5,10)+'</td><td>'+t[i].slice(11,16)+'</td><td>'+wci[0]+'</td><td>'+f1(h.temperature_2m[i])+'</td><td>'+r0(h.wind_speed_10m[i])+'</td><td>'+bft(h.wind_speed_10m[i])+'</td><td>'+r0(h.wind_gusts_10m[i])+'</td><td>'+bft(h.wind_gusts_10m[i])+'</td><td>'+comp(h.wind_direction_10m[i])+'</td><td>'+(pr>0?'Yes':'No')+'</td></tr>';
  }
  out+='<h4>Forecast (3 days, daytime 09–21h)</h4>';
  out+='<table><thead><tr><th>Day</th><th>Date</th><th>Time</th><th>Weather</th><th>Temp °C</th><th>Wind</th><th>Bft</th><th>Gusts</th><th>Gust Bft</th><th>Direction</th><th>Rain</th></tr></thead><tbody>'+rows+'</tbody></table>';
  document.getElementById("w").innerHTML=out;
}
</script></body></html>
"""


def render_weather_browser(lat, lon):
    """Rendert das Wetter-Panel clientseitig (fetch im Browser des Besuchers)."""
    html = (
        _WEATHER_HTML
        .replace("__LAT__", f"{float(lat):.4f}")
        .replace("__LON__", f"{float(lon):.4f}")
    )
    components.html(html, height=640, scrolling=True)


def _preset_index(options, value):
    """Position von value in options (für selectbox-Vorbelegung), sonst 0."""
    try:
        return options.index(value) if value in options else 0
    except Exception:
        return 0


@st.fragment
def render_rankings(results_container):
    # Als Fragment gekapselt: Ändert der Nutzer einen Filter (Gruppe/Lokation/
    # Jahr/Monat/Tag), läuft NUR diese Funktion neu – nicht das gesamte Skript.
    # WICHTIG: Das Fragment wird IN der Sidebar verankert (Aufruf via
    # `with sidebar_tab_filter:`), denn ein Fragment darf Widgets nur an seinem
    # EIGENEN Anker erzeugen, nicht in einen externen Container. Die Filter
    # rendern daher hier (Sidebar); die Tabellen (keine Widgets) schreiben wir
    # in den separaten Haupt-Container `results_container`.
    user = st.session_state.get("user")
    username = user["username"] if user else None
    preset = load_user_pref(username)

    sport = active_sport()
    ranking = complete_sessions(load_sessions(sport))

    months = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December",
    ]

    member_groups = my_member_groups(user["id"]) if user else []

    # Aktuelle Filter-Auswahl bestimmen: aus session_state (vom letzten Lauf)
    # oder beim ERSTEN Laden aus dem gespeicherten Preset. Dadurch können wir die
    # Tabellen rendern, BEVOR die Filter-Widgets gebaut werden – die Rankings
    # erscheinen so zuerst, der teurere Optionen-Aufbau der Dropdowns läuft erst
    # danach.
    # Filter-Keys pro Sport getrennt (z.B. "rank_spot_kitesurf"), sonst würde ein
    # in Windsurf gewählter Spot beim Wechsel zu Kite einen ungültigen Selectbox-
    # Wert erzeugen. Der gemeinsame Preset dient als Erstbelegung (ungültige Werte
    # fängt _preset_index ab).
    group_choice = st.session_state.get(f"rank_group_{sport}", preset.get("group") or ALL_GROUP)
    spot_filter = st.session_state.get(f"rank_spot_{sport}", preset.get("spot") or "Overall")
    year_filter = st.session_state.get(f"rank_year_{sport}", preset.get("year") or "All years")
    month_filter = st.session_state.get(f"rank_month_{sport}", preset.get("month") or "Whole year")
    day_filter = st.session_state.get(f"rank_day_{sport}", preset.get("day") or "Whole month")
    gear_filter = st.session_state.get(f"rank_gear_{sport}", preset.get("gear") or "All")

    # Erweiterte (optionale) Filter – 0 bzw. (0,0) bedeutet „aus".
    front_max = st.session_state.get(f"rank_front_{sport}", 0)
    fin_max = st.session_state.get(f"rank_finmax_{sport}", 0.0)
    weight_from = st.session_state.get(f"rank_wfrom_{sport}", 0.0)
    weight_to = st.session_state.get(f"rank_wto_{sport}", 0.0)

    extra = {
        "front_max": front_max,
        "fin_max": fin_max,
        "weight_from": weight_from,
        "weight_to": weight_to,
    }

    # ---- Tabellen ZUERST (Hauptinhalt) in den Haupt-Container ----
    # Reine Anzeige (keine Widgets) -> aus dem Fragment in externen Container ok;
    # .container() im st.empty()-Platzhalter ersetzt den Inhalt bei jedem Rerun.
    with results_container.container():
        _render_ranking_tables(
            ranking, group_choice, member_groups, months,
            spot_filter, year_filter, month_filter, day_filter, gear_filter, extra,
        )

    # ---- Filter-UI DANACH am Fragment-Anker (= Sidebar-Tab „Filter") ----
    # Eigener Container (im Fragment erzeugt) -> Widgets sind hier erlaubt. Die
    # Selectboxen lesen ihren Wert aus session_state (oben bereits ausgelesen);
    # index= dient nur der Erstbelegung beim allerersten Laden.
    with st.container():
        with st.expander("⭐ My start (filter preset)", expanded=False):
            st.caption(
                "Save the current filter selection as your start – it will be "
                "preselected automatically every time you open the app."
            )

            if st.button("💾 Save current filters", use_container_width=True):
                save_user_pref(username, {
                    "group": st.session_state.get(f"rank_group_{sport}", ALL_GROUP),
                    "spot": st.session_state.get(f"rank_spot_{sport}", "Overall"),
                    "year": st.session_state.get(f"rank_year_{sport}", "All years"),
                    "month": st.session_state.get(f"rank_month_{sport}", "Whole year"),
                    "day": st.session_state.get(f"rank_day_{sport}", "Whole month"),
                    "gear": st.session_state.get(f"rank_gear_{sport}", "All"),
                })
                st.success("Saved – will be loaded on start from now on.")

            if preset and st.button("↺ Reset", use_container_width=True):
                delete_user_pref(username)
                for _k in ("rank_group", "rank_spot", "rank_year", "rank_month", "rank_day",
                           "rank_gear", "rank_front", "rank_finmax", "rank_wfrom", "rank_wto"):
                    st.session_state.pop(f"{_k}_{sport}", None)
                st.rerun()

            # Einstieg ins Spot-TV (Vollbild-Live-Screen) fuer den aktuellen Spot.
            st.markdown("---")
            if spot_filter and spot_filter != "Overall":
                _tv_url = "?" + urlencode(
                    {"tv": "1", "sport": sport, "spot": spot_filter, "mode": "today"}
                )
                st.link_button(
                    f"📺 Open Spot TV · {spot_filter}", _tv_url, use_container_width=True
                )
                st.caption("Full-screen live screen for café/shop. Open in a new tab, then F11.")
            else:
                st.caption("📺 Spot TV: pick a Location first – the live-screen link appears here.")

        group_options = [ALL_GROUP] + [g["name"] for g in member_groups]
        st.selectbox(
            "👥 Group",
            group_options,
            index=_preset_index(group_options, group_choice),
            key=f"rank_group_{sport}",
            help="\"All\" shows every rider. You only see group results as a member.",
        )

        # Auswahl-Optionen für Lokation/Jahr abhängig von der gewählten Gruppe.
        if not ranking.empty:
            opt_df = ranking.copy()
            opt_df["_date"] = (
                pd.to_datetime(opt_df["date"], errors="coerce", format="mixed")
                if "date" in opt_df.columns else pd.NaT
            )

            if group_choice != ALL_GROUP:
                _gid = next((g["id"] for g in member_groups if g["name"] == group_choice), None)
                if _gid is not None and "name" in opt_df.columns:
                    opt_df = opt_df[opt_df["name"].astype(str).isin(set(group_member_names(_gid)))]

            spot_values = (
                opt_df["surfspot"].dropna().astype(str)
                if "surfspot" in opt_df.columns else pd.Series(dtype=str)
            )
            spots = sorted(s for s in spot_values.unique() if s.strip())
            years = sorted(opt_df["_date"].dropna().dt.year.unique(), reverse=True)
        else:
            spots, years = [], []

        spot_options = ["Overall"] + spots
        year_options = ["All years"] + [str(y) for y in years]
        month_options = ["Whole year"] + months
        day_options = ["Whole month"] + [str(d) for d in range(1, 32)]

        st.selectbox(
            "📍 Location", spot_options,
            index=_preset_index(spot_options, spot_filter), key=f"rank_spot_{sport}",
        )
        st.selectbox(
            "📅 Year", year_options,
            index=_preset_index(year_options, year_filter), key=f"rank_year_{sport}",
        )
        st.selectbox(
            "📆 Month", month_options,
            index=_preset_index(month_options, month_filter), key=f"rank_month_{sport}",
        )
        st.selectbox(
            "🗓️ Day", day_options,
            index=_preset_index(day_options, day_filter), key=f"rank_day_{sport}",
        )

        gear_options = ["All"] + SPORT_META[sport]["gear_type_options"]
        st.selectbox(
            f"🪙 {SPORT_META[sport]['gear_type_label']}", gear_options,
            index=_preset_index(gear_options, gear_filter), key=f"rank_gear_{sport}",
        )

        # --- Erweiterte Filter (optional; 0 = aus) ---
        st.caption("Advanced filters (0 = off)")
        st.number_input(
            "Max. front wing (cm²)", min_value=0, step=10,
            key=f"rank_front_{sport}",
            help="Only foil sessions with a front wing size up to this value.",
        )
        st.number_input(
            "Max. fin (cm)", min_value=0.0, step=0.5,
            key=f"rank_finmax_{sport}",
            help="Only fin sessions with a fin size up to this value.",
        )
        wcol1, wcol2 = st.columns(2)
        wcol1.number_input(
            "Weight from (kg)", min_value=0.0, step=1.0, key=f"rank_wfrom_{sport}",
            help="Filter riders by their profile weight (0 = no limit).",
        )
        wcol2.number_input(
            "Weight to (kg)", min_value=0.0, step=1.0, key=f"rank_wto_{sport}",
        )


@st.cache_data(show_spinner=False, max_entries=64)
def _enrich_ranking(ranking):
    """Ergänzt das (bereits gefilterte) Ranking um die teuren Spalten „Wetter"
    und „Trust" (zeilenweises apply).

    Gecacht und selbst-invalidierend: Bei gleichem Eingangs-DataFrame wird nicht
    neu gerechnet; ändern sich die Sessions (load_sessions liefert andere Daten)
    oder der Filter, ändert sich der Cache-Key automatisch. max_entries begrenzt
    den Speicher über viele Filter-Kombinationen hinweg.
    """
    def weather_summary(row):
        parts = []

        if pd.notna(row.get("weather_code")):
            emoji = WEATHER_CODES.get(int(row["weather_code"]), ("", ""))[0]

            if emoji:
                parts.append(emoji)

        if pd.notna(row.get("wind_kmh")):
            wind = f"{row['wind_kmh']:.0f} km/h"

            if pd.notna(row.get("wind_dir_deg")):
                wind += f" {degrees_to_compass(row['wind_dir_deg'])}"

            parts.append(wind)

        if pd.notna(row.get("temp_c")):
            parts.append(f"{row['temp_c']:.0f}°C")

        return " · ".join(parts) if parts else "–"

    ranking = ranking.copy()
    ranking["Weather"] = ranking.apply(weather_summary, axis=1)
    ranking["Trust"] = ranking["trust_score"].apply(_trust_badge)
    return ranking


def df_height(n_rows, max_rows=15):
    """Exakte Pixel-Höhe für st.dataframe, damit die Box ohne Leerraum gefüllt
    ist (kein Default-Leerraum unter wenigen Zeilen).

    Streamlit rendert Kopf- und Datenzeilen mit ~35 px. Bei mehr als `max_rows`
    Zeilen wird gedeckelt – dann scrollt die Tabelle innerhalb der gefüllten Box.
    """
    rows = min(max(int(n_rows), 1), max_rows)
    return (rows + 1) * 35 + 3


def _render_ranking_tables(ranking, group_choice, member_groups, months,
                           spot_filter, year_filter, month_filter, day_filter,
                           gear_filter="All", extra=None):
    """Rendert die vier Ranking-Tabellen – reine Anzeige (keine Widgets)."""
    extra = extra or {}
    gear_label = SPORT_META[active_sport()]["gear_label"]  # "Sail" / "Kite"
    st.markdown("## 🏆 Online rankings")

    flash = st.session_state.pop("ranking_flash", None)

    if flash:
        st.success(flash)

    if ranking.empty:
        st.info("No online ranking entries yet.")
        return

    if "date" not in ranking.columns:
        ranking["date"] = ""

    for column in ("wind_kmh", "wind_dir_deg", "temp_c", "weather_code", "trust_score",
                   "gear_type", "fin_size_cm", "foil_front_cm2",
                   "max_airtime_s", "max_jump_m", "strokes", "max_cadence_spm"):
        if column not in ranking.columns:
            ranking[column] = None

    ranking["_date"] = pd.to_datetime(ranking["date"], errors="coerce", format="mixed")

    if group_choice != ALL_GROUP:
        group_id = next((g["id"] for g in member_groups if g["name"] == group_choice), None)

        if group_id is not None:
            member_names = set(group_member_names(group_id))
            ranking = ranking[ranking["name"].astype(str).isin(member_names)].copy()

            if ranking.empty:
                st.info(f"The group \"{group_choice}\" has no sessions yet.")
                return

    # Defensiv gegen veraltete/ungültige Filterwerte (z.B. ein altes Preset mit
    # deutschen Sentinels „Alle Jahre"/„Ganzes Jahr"): ungültige Werte werden wie
    # „kein Filter" behandelt statt zu crashen.
    if spot_filter and spot_filter != "Overall":
        ranking = ranking[ranking["surfspot"].astype(str) == spot_filter]

    if str(year_filter).isdigit():
        ranking = ranking[ranking["_date"].dt.year == int(year_filter)]

    if month_filter in months:
        ranking = ranking[ranking["_date"].dt.month == months.index(month_filter) + 1]

        if str(day_filter).isdigit():
            ranking = ranking[ranking["_date"].dt.day == int(day_filter)]

    if gear_filter and gear_filter != "All":
        # Ältere Sessions ohne gear_type (None) fallen bei Fin/Foil-Auswahl raus.
        ranking = ranking[ranking["gear_type"].astype(str) == gear_filter]

    # --- Erweiterte Filter (Frontwing / Finne / Gewicht). Sessions ohne den
    # jeweiligen Wert fallen bei aktivem Filter raus (NaN-Vergleich = False). ---
    front_max = extra.get("front_max") or 0
    if front_max:
        fw = pd.to_numeric(ranking["foil_front_cm2"], errors="coerce")
        ranking = ranking[fw <= front_max]

    fin_max = extra.get("fin_max") or 0
    if fin_max:
        fs = pd.to_numeric(ranking["fin_size_cm"], errors="coerce")
        ranking = ranking[fs <= fin_max]

    weight_from = extra.get("weight_from") or 0
    weight_to = extra.get("weight_to") or 0
    if weight_from or weight_to:
        weights = load_user_weights()
        w = pd.to_numeric(ranking["name"].astype(str).map(weights), errors="coerce")
        mask = w.notna()
        if weight_from:
            mask &= w >= weight_from
        if weight_to:
            mask &= w <= weight_to
        ranking = ranking[mask]

    ranking = ranking.copy()

    if ranking.empty:
        st.info("No entries for the selected filters.")
        return

    # Teure Spalten (zeilenweises apply für Wetter-Text + Trust-Badge) gecacht
    # ergänzen – siehe _enrich_ranking. Vermeidet die Neuberechnung bei jedem
    # (Fragment-)Rerun, solange sich das gefilterte Ranking nicht ändert.
    ranking = _enrich_ranking(ranking)

    # 2x2-Raster. width="stretch" (moderne API, ersetzt das veraltete
    # use_container_width) lässt jede Tabelle ihre Box voll ausfüllen; bei vielen
    # Spalten ist sie innerhalb der Box horizontal scrollbar.
    rcol1, rcol2 = st.columns(2)

    with rcol1:
        st.markdown("### 🏆 Best 30 seconds")

        r30 = ranking[[
            "date",
            "name",
            "speed_30s_kmh",
            "speed_30s_kn",
            "surfspot",
            "board",
            "sail",
            "Weather",
            "Trust",
        ]].copy()

        # Pro Fahrer nur die beste Session (kein Mehrfach-Platzieren).
        r30 = (
            r30.sort_values("speed_30s_kmh", ascending=False)
            .drop_duplicates(subset="name", keep="first")
            .reset_index(drop=True)
        )
        r30.insert(0, "Rank", r30.index + 1)

        r30 = r30.rename(columns={
            "date": "Date",
            "name": "Name",
            "surfspot": "Surf spot",
            "board": "Board",
            "sail": gear_label,
            "speed_30s_kmh": "30s km/h",
            "speed_30s_kn": "30s kn",
        })

        st.dataframe(r30, width="stretch", hide_index=True, height=df_height(len(r30)))

    with rcol2:
        st.markdown("### ⚡ Top speed 2 seconds")

        r1 = ranking[[
            "date",
            "name",
            "speed_1s_kmh",
            "speed_1s_kn",
            "surfspot",
            "board",
            "sail",
            "Weather",
            "Trust",
        ]].copy()

        # Pro Fahrer nur die beste Session.
        r1 = (
            r1.sort_values("speed_1s_kmh", ascending=False)
            .drop_duplicates(subset="name", keep="first")
            .reset_index(drop=True)
        )
        r1.insert(0, "Rank", r1.index + 1)

        r1 = r1.rename(columns={
            "date": "Date",
            "name": "Name",
            "surfspot": "Surf spot",
            "board": "Board",
            "sail": gear_label,
            "speed_1s_kmh": "2s km/h",
            "speed_1s_kn": "2s kn",
        })

        st.dataframe(r1, width="stretch", hide_index=True, height=df_height(len(r1)))

    rcol3, rcol4 = st.columns(2)

    with rcol3:
        st.markdown("### 🚩 Longest run")

        rrun = ranking[[
            "date",
            "name",
            "longest_run_km",
            "longest_run_m",
            "surfspot",
            "board",
            "sail",
            "Weather",
            "Trust",
        ]].copy()

        # Pro Fahrer nur der beste (längste) Run.
        rrun = (
            rrun.sort_values("longest_run_m", ascending=False)
            .drop_duplicates(subset="name", keep="first")
            .reset_index(drop=True)
        )
        rrun.insert(0, "Rank", rrun.index + 1)

        rrun = rrun.rename(columns={
            "date": "Date",
            "name": "Name",
            "surfspot": "Surf spot",
            "board": "Board",
            "sail": gear_label,
            "longest_run_km": "Run km",
            "longest_run_m": "Run m",
        })

        st.dataframe(rrun, width="stretch", hide_index=True, height=df_height(len(rrun)))

    with rcol4:
        st.markdown("### 👥 Longest total distance per rider")

        rtotal = (
            ranking
            .groupby("name", as_index=False)
            .agg(
                total_distance_km=("total_distance_km", "sum"),
                last_date=("date", "max"),
            )
            .sort_values("total_distance_km", ascending=False)
            .reset_index(drop=True)
        )

        rtotal.insert(0, "Rank", rtotal.index + 1)

        rtotal = rtotal.rename(columns={
            "name": "Name",
            "total_distance_km": "Total distance km",
            "last_date": "Last session",
        })

        st.dataframe(rtotal, width="stretch", hide_index=True, height=df_height(len(rtotal)))

    # Zusatz-Rankings aus den Uhr-Daten: Wind -> Sprünge, SUP -> Paddeln.
    # Sessions ohne die jeweiligen Werte fallen raus.
    def _metric_table(metric, title, col_label, decimals=1, empty_msg="No data yet."):
        tbl = ranking[[
            "date", "name", metric, "surfspot", "board", "sail", "Weather", "Trust",
        ]].copy()
        tbl[metric] = pd.to_numeric(tbl[metric], errors="coerce")
        tbl = tbl[tbl[metric] > 0].dropna(subset=[metric])
        tbl = (
            tbl.sort_values(metric, ascending=False)
            .drop_duplicates(subset="name", keep="first")
            .reset_index(drop=True)
        )
        st.markdown(title)
        if tbl.empty:
            st.caption(empty_msg)
            return
        tbl.insert(0, "Rank", tbl.index + 1)
        if decimals == 0:
            tbl[metric] = tbl[metric].round(0).astype(int)
        else:
            tbl[metric] = tbl[metric].round(decimals)
        tbl = tbl.rename(columns={
            "date": "Date",
            "name": "Name",
            "surfspot": "Surf spot",
            "board": "Board",
            "sail": gear_label,
            metric: col_label,
        })
        st.dataframe(tbl, width="stretch", hide_index=True, height=df_height(len(tbl)))

    if active_sport() == "sup":
        scol1, scol2 = st.columns(2)
        with scol1:
            _metric_table("strokes", "### 🛶 Most paddle strokes", "Strokes",
                          decimals=0, empty_msg="No paddle data yet.")
        with scol2:
            _metric_table("max_cadence_spm", "### ⏱️ Max strokes / minute", "Max spm",
                          decimals=0, empty_msg="No paddle data yet.")
    else:
        jcol1, jcol2 = st.columns(2)
        with jcol1:
            _metric_table("max_airtime_s", "### 🪂 Best airtime", "Airtime s",
                          empty_msg="No jump data yet – record a session with jumps on the watch.")
        with jcol2:
            _metric_table("max_jump_m", "### 🚀 Highest jump", "Jump m",
                          empty_msg="No jump data yet – record a session with jumps on the watch.")


def semicircles_to_degrees(value):
    if value is None:
        return None

    return value * (180 / 2**31)


def read_fit_file(uploaded_file):
    fitfile = FitFile(uploaded_file)
    records = []

    for message in fitfile.get_messages("record"):
        row = {}

        for field in message:
            row[field.name] = field.value

        records.append(row)

    df = pd.DataFrame(records)

    if df.empty:
        return df

    if "timestamp" in df.columns:
        df["timestamp"] = pd.to_datetime(df["timestamp"])

    if "enhanced_speed" in df.columns:
        df["speed_kmh"] = df["enhanced_speed"] * 3.6
    elif "speed" in df.columns:
        df["speed_kmh"] = df["speed"] * 3.6

    if "position_lat" in df.columns:
        df["lat"] = df["position_lat"].apply(semicircles_to_degrees)

    if "position_long" in df.columns:
        df["lon"] = df["position_long"].apply(semicircles_to_degrees)

    return df


def surf_weather_time(df):
    """Repräsentative Zeit (UTC) fürs Wetter.

    Nimmt den Median der Zeitpunkte, an denen tatsächlich gesurft wurde
    (Speed über der Run-Grenze) – so trifft das stündliche Wetter die echte
    Surfzeit und nicht das Aufriggen/Warten davor. Fällt auf die zeitliche
    Mitte der Aufzeichnung und schließlich den Beginn zurück.
    """
    if "timestamp" not in df.columns:
        return None

    stamps = df["timestamp"].dropna()

    if stamps.empty:
        return None

    if "speed_kmh" in df.columns:
        planing = df.loc[df["speed_kmh"] >= MIN_RUN_SPEED_KMH, "timestamp"].dropna()

        if not planing.empty:
            return planing.median()

    return stamps.min() + (stamps.max() - stamps.min()) / 2


def best_average_speed(df, seconds):
    if "timestamp" not in df.columns or "speed_kmh" not in df.columns:
        return None

    data = df[["timestamp", "speed_kmh"]].dropna().copy()

    if len(data) < 2:
        return None

    data = data.sort_values("timestamp").set_index("timestamp")
    best = data["speed_kmh"].rolling(f"{seconds}s").mean().max()

    if pd.isna(best):
        return None

    return float(best)


def detect_runs(df):
    required_columns = {"timestamp", "speed_kmh", "distance"}

    if not required_columns.issubset(df.columns):
        return pd.DataFrame()

    data = df[["timestamp", "speed_kmh", "distance"]].dropna().copy()
    data = data.sort_values("timestamp").reset_index(drop=True)

    runs = []
    in_run = False
    start_index = None

    for i, row in data.iterrows():
        speed = row["speed_kmh"]

        if not in_run and speed >= MIN_RUN_SPEED_KMH:
            in_run = True
            start_index = i

        elif in_run and speed <= END_RUN_SPEED_KMH:
            run = data.iloc[start_index:i + 1]
            distance_m = run["distance"].iloc[-1] - run["distance"].iloc[0]

            if distance_m >= MIN_RUN_DISTANCE_M:
                runs.append({
                    "Start": run["timestamp"].iloc[0],
                    "End": run["timestamp"].iloc[-1],
                    "Duration": run["timestamp"].iloc[-1] - run["timestamp"].iloc[0],
                    "Distance m": distance_m,
                    "Distance km": distance_m / 1000,
                    "Max Speed km/h": run["speed_kmh"].max(),
                    "Ø Speed km/h": run["speed_kmh"].mean(),
                })

            in_run = False
            start_index = None

    if in_run and start_index is not None:
        run = data.iloc[start_index:]
        distance_m = run["distance"].iloc[-1] - run["distance"].iloc[0]

        if distance_m >= MIN_RUN_DISTANCE_M:
            runs.append({
                "Start": run["timestamp"].iloc[0],
                "End": run["timestamp"].iloc[-1],
                "Duration": run["timestamp"].iloc[-1] - run["timestamp"].iloc[0],
                "Distance m": distance_m,
                "Distance km": distance_m / 1000,
                "Max Speed km/h": run["speed_kmh"].max(),
                "Ø Speed km/h": run["speed_kmh"].mean(),
            })

    return pd.DataFrame(runs)


# =====================================================================
#  Trust Score – Plausibilitätsprüfung einer Aufzeichnung (Anti-Cheat)
# =====================================================================

def _haversine_m(lat1, lon1, lat2, lon2):
    """Distanz in Metern zwischen Punktfolgen (vektorisiert)."""
    radius = 6371000.0
    p1 = np.radians(lat1)
    p2 = np.radians(lat2)
    dphi = np.radians(lat2 - lat1)
    dlmb = np.radians(lon2 - lon1)
    a = np.sin(dphi / 2.0) ** 2 + np.cos(p1) * np.cos(p2) * np.sin(dlmb / 2.0) ** 2
    return 2.0 * radius * np.arcsin(np.sqrt(np.clip(a, 0, 1)))


def _bearing_deg(lat1, lon1, lat2, lon2):
    """Kurs (0–360°) von Punkt 1 nach Punkt 2 (vektorisiert)."""
    p1 = np.radians(lat1)
    p2 = np.radians(lat2)
    dl = np.radians(lon2 - lon1)
    x = np.sin(dl) * np.cos(p2)
    y = np.cos(p1) * np.sin(p2) - np.sin(p1) * np.cos(p2) * np.cos(dl)
    return (np.degrees(np.arctan2(x, y)) + 360.0) % 360.0


def _penalty(value, good, bad, max_pen):
    """Linearer Abzug: <=good → 0, >=bad → max_pen, dazwischen interpoliert."""
    if value is None or value != value:  # None oder NaN
        return 0.0
    if value <= good:
        return 0.0
    if value >= bad:
        return float(max_pen)
    return float(max_pen) * (value - good) / (bad - good)


def _status(pen, max_pen):
    if pen >= 0.6 * max_pen:
        return "bad"
    if pen >= 0.2 * max_pen:
        return "warn"
    return "ok"


def compute_trust_score(df, spot_top_kmh=None):
    """Plausibilitäts-/Trust-Score 0–100 aus der GPS-Physik einer Session.

    Prüft maximale Beschleunigung, Kursänderung bei hoher Geschwindigkeit,
    GPS-Rauschen, Punktdichte und (optional) den Vergleich mit dem Spot-Bestwert.
    Kann eine Kennzahl nicht berechnet werden (z.B. ohne GPS), wird sie neutral
    gewertet (kein Abzug). Rückgabe: {'score', 'components', 'note'}.
    """
    if df is None or df.empty or not {"timestamp", "speed_kmh"}.issubset(df.columns):
        return {"score": None, "components": [], "note": "Not enough data for a rating."}

    d = df.dropna(subset=["timestamp", "speed_kmh"]).sort_values("timestamp").reset_index(drop=True)

    if len(d) < 10:
        return {"score": None, "components": [], "note": "Recording too short for a rating."}

    dt = d["timestamp"].diff().dt.total_seconds()
    dt_valid = dt.where(dt > 0)
    speed_ms = d["speed_kmh"] / 3.6
    has_gps = (
        {"lat", "lon"}.issubset(d.columns)
        and int(d[["lat", "lon"]].notna().all(axis=1).sum()) > 10
    )

    components = []
    score = 100.0

    # 1) Maximale Beschleunigung (m/s²), robust über 99,5-Perzentil
    accel = (speed_ms.diff() / dt_valid).abs()
    max_accel = float(accel.quantile(0.995)) if accel.notna().any() else None
    pen = _penalty(max_accel, good=4.0, bad=9.0, max_pen=30.0)
    score -= pen
    components.append({
        "label": "Max. acceleration",
        "value": "–" if max_accel is None else f"{max_accel:.1f} m/s²",
        "status": _status(pen, 30.0),
    })

    # 2) Kursänderung bei hoher Geschwindigkeit (°/s, nur > 25 km/h)
    turn_val = None
    if has_gps:
        brg = pd.Series(
            _bearing_deg(d["lat"], d["lon"], d["lat"].shift(-1), d["lon"].shift(-1)),
            index=d.index,
        )
        dbrg = brg.diff().abs()
        dbrg = dbrg.where(dbrg <= 180, 360 - dbrg)
        turn_rate = (dbrg / dt_valid).where(speed_ms > (25 / 3.6))
        if turn_rate.notna().any():
            turn_val = float(turn_rate.quantile(0.99))
    pen = _penalty(turn_val, good=40.0, bad=150.0, max_pen=25.0)
    score -= pen
    components.append({
        "label": "Course change at speed",
        "value": "–" if turn_val is None else f"{turn_val:.0f} °/s",
        "status": _status(pen, 25.0),
    })

    # 3) GPS-Rauschen: Anteil „unmöglicher" Punkt-Sprünge
    noise_frac = None
    if has_gps:
        seg = _haversine_m(d["lat"], d["lon"], d["lat"].shift(-1), d["lon"].shift(-1))
        implied = seg / dt_valid
        impossible = (implied > (speed_ms + 5)) & (implied > speed_ms * 1.5)
        denom = len(d) - 1
        if denom > 0:
            noise_frac = float(impossible.sum()) / denom * 100.0
    pen = _penalty(noise_frac, good=1.0, bad=15.0, max_pen=25.0)
    score -= pen
    components.append({
        "label": "GPS noise",
        "value": "–" if noise_frac is None else f"{noise_frac:.1f} % outliers",
        "status": _status(pen, 25.0),
    })

    # 4) Punktdichte: Meter pro Trackpunkt
    total_m = None
    if "distance" in d.columns and d["distance"].notna().any():
        total_m = float(d["distance"].max() - d["distance"].min())
    elif has_gps:
        total_m = float(
            _haversine_m(d["lat"], d["lon"], d["lat"].shift(-1), d["lon"].shift(-1)).sum()
        )
    m_per_point = total_m / (len(d) - 1) if total_m and len(d) > 1 else None
    pen = _penalty(m_per_point, good=15.0, bad=60.0, max_pen=20.0)
    score -= pen
    components.append({
        "label": "Point density",
        "value": "–" if m_per_point is None else f"{m_per_point:.0f} m/point",
        "status": _status(pen, 20.0),
    })

    # 5) Vergleich mit typischem Spot-Bestwert (falls vorhanden)
    if spot_top_kmh and spot_top_kmh > 0:
        ratio = float(d["speed_kmh"].max()) / float(spot_top_kmh)
        pen = _penalty(ratio, good=1.3, bad=2.2, max_pen=20.0)
        score -= pen
        components.append({
            "label": "Comparison with spot best",
            "value": f"{ratio * 100:.0f} % of the spot top",
            "status": _status(pen, 20.0),
        })

    return {"score": int(max(0, min(100, round(score)))), "components": components, "note": ""}


def render_trust_score(result):
    """Zeigt den Trust Score samt Teilbewertungen an."""
    score = result.get("score")

    if score is None:
        st.caption(f"🔍 Trust Score: {result.get('note', 'not available')}")
        return

    dot = "🟢" if score >= 80 else "🟡" if score >= 55 else "🔴"
    st.markdown(f"### {dot} Trust Score: {score}/100")
    st.progress(score / 100)

    icons = {"ok": "✅", "warn": "⚠️", "bad": "❌"}

    for c in result.get("components", []):
        st.caption(f"{icons.get(c['status'], '•')} {c['label']}: {c['value']}")

    st.caption(
        "Heuristic plausibility check of the GPS data – not proof of cheating, "
        "but a hint at unusual or noisy recordings."
    )


def _trust_badge(score):
    """Kompaktes Trust-Symbol für die Ranking-Tabellen (🟢/🟡/🔴 + Wert)."""
    if score is None or (isinstance(score, float) and pd.isna(score)):
        return "–"

    try:
        s = int(round(float(score)))
    except Exception:
        return "–"

    dot = "🟢" if s >= 80 else "🟡" if s >= 55 else "🔴"
    return f"{dot} {s}"


def show_map(df):
    if "lat" not in df.columns or "lon" not in df.columns:
        st.warning("No GPS data found.")
        return

    gps_df = df.dropna(subset=["lat", "lon"]).copy()

    if gps_df.empty:
        st.warning("No GPS data found.")
        return

    view_state = pdk.ViewState(
        latitude=gps_df["lat"].mean(),
        longitude=gps_df["lon"].mean(),
        zoom=13,
    )

    has_speed = "speed_kmh" in gps_df.columns and gps_df["speed_kmh"].notna().any()
    speed_lo = speed_hi = None

    if has_speed:
        # Track nach Geschwindigkeit einfärben: blau (langsam) -> orange (schnell).
        # Jede Teilstrecke (Punkt i -> i+1) bekommt eine eigene Farbe.
        pts = gps_df[["lon", "lat", "speed_kmh"]].to_numpy(dtype=float)

        # Sehr lange Tracks ausdünnen, damit die Browser-Last klein bleibt.
        step = max(1, len(pts) // 4000)
        pts = pts[::step]

        speeds = pts[:, 2]
        speed_lo = float(np.nanmin(speeds))
        # 95.-Perzentil als robuster Maximalwert (ein GPS-Ausreißer soll die
        # Farbskala nicht ruinieren).
        speed_hi = float(np.nanpercentile(speeds, 95))
        span = speed_hi - speed_lo if speed_hi > speed_lo else 1.0

        def _color(v):
            t = 0.0 if np.isnan(v) else min(1.0, max(0.0, (v - speed_lo) / span))
            # linear blau [25,100,230] -> orange [255,125,0]
            return [
                int(25 + t * (255 - 25)),
                int(100 + t * (125 - 100)),
                int(230 + t * (0 - 230)),
            ]

        segments = [
            {
                "path": [
                    [pts[i, 0], pts[i, 1]],
                    [pts[i + 1, 0], pts[i + 1, 1]],
                ],
                "color": _color((speeds[i] + speeds[i + 1]) / 2.0),
            }
            for i in range(len(pts) - 1)
        ]

        layers = [pdk.Layer(
            "PathLayer",
            data=segments,
            get_path="path",
            get_color="color",
            get_width=4,
            width_min_pixels=3,
        )]
    else:
        # Kein Speed in den Daten -> einfarbiger Track in kräftigem Orange.
        layers = [pdk.Layer(
            "PathLayer",
            data=[{"path": gps_df[["lon", "lat"]].values.tolist()}],
            get_path="path",
            get_color=[255, 125, 0],
            get_width=4,
            width_min_pixels=3,
        )]

    st.pydeck_chart(
        pdk.Deck(
            # Heller CARTO-Basemap "Voyager" – ohne Mapbox-Token nutzbar und
            # zeigt das Meer in hellem Blau (Positron färbt Wasser nur grau).
            # Vorher map_style=None -> pydeck fiel auf den dunklen Default zurück.
            map_provider="carto",
            map_style=pdk.map_styles.CARTO_ROAD,
            initial_view_state=view_state,
            layers=layers,
        ),
        use_container_width=True,
    )

    if has_speed:
        st.caption(
            f"🟦 slow → 🟧 fast · color scale {speed_lo:.0f}–{speed_hi:.0f} km/h"
        )


# =====================================================================
#  Spot Live Dashboard ("Spot TV") – Vollbild-Ansicht fuer Cafe/Shop/Club.
#  Aufruf per Query-Parameter, z.B.:
#    ?tv=1&spot=Koeln&mode=today&sport=windsurf
#    &period=week&group=MeinClub&sponsor=Surfshop%20XY&logo=https://...&event=...
# =====================================================================

_TV_CSS = """
<style>
  [data-testid="stSidebar"], [data-testid="stHeader"], #MainMenu, footer {display:none !important;}
  .block-container {padding-top:1rem !important; max-width:100% !important;}
  .tv-header {display:grid; grid-template-columns:1fr auto 1fr; align-items:center;
              gap:12px; margin-bottom:12px;}
  .tv-brand {font-size:34px; font-weight:900; letter-spacing:-0.4px; white-space:nowrap;}
  .tv-brand .dot {color:#2bd4d9;}
  .tv-spot {text-align:center;}
  .tv-spot .name {font-size:52px; font-weight:900; line-height:1.0;}
  .tv-spot .event {font-size:22px; opacity:.9; margin-top:4px;}
  .tv-header .sponsor {text-align:right;}
  .tv-sponsor-chip {display:inline-block; background:#ffffff; border-radius:14px;
                    padding:8px 16px; box-shadow:0 2px 10px rgba(0,0,0,.3);}
  .tv-sponsor-chip img {max-height:56px; max-width:300px; display:block;}
  .tv-presented {font-size:16px; opacity:.85; margin-top:5px;}
  .tv-cards {display:flex; flex-wrap:wrap; gap:16px; margin:6px 0 18px;}
  .tv-card {flex:1 1 200px; background:rgba(255,255,255,.08);
            border:1px solid rgba(255,255,255,.18); border-radius:18px; padding:14px 20px;}
  .tv-card .lbl {font-size:19px; opacity:.85;}
  .tv-card .val {font-size:56px; font-weight:800; line-height:1.1;}
  .tv-card .sub {font-size:18px; opacity:.8; min-height:1em;}
  .tv-rank-title {font-size:28px; font-weight:800; margin:12px 0 8px;}
  .tv-grid {display:grid; grid-template-columns:repeat(2, 1fr); gap:6px 30px;}
  .tv-rk {display:flex; align-items:center; gap:14px; padding:8px 10px;
          border-bottom:1px solid rgba(255,255,255,.12); font-size:27px;}
  .tv-rk .pos {min-width:48px; text-align:center; font-weight:800; opacity:.9;}
  .tv-rk .nm {flex:1; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;}
  .tv-rk .sp {white-space:nowrap; font-weight:800;}
  .tv-rk .sp small {font-size:16px; opacity:.6; font-weight:600;}
  .tv-update {font-size:20px; opacity:.8; margin-top:12px;}
  .tv-msg {font-size:26px; opacity:.85; padding:24px 0;}

  /* Balken-Leaderboard: links Plaetze 1..N/2, rechts der Rest. */
  .lb {display:grid; grid-template-columns:1fr 1fr; gap:0 44px; margin-top:4px;}
  .lb-col {display:flex; flex-direction:column;}
  .lb-row {display:flex; align-items:center; gap:14px; padding:9px 6px; font-size:30px;
           border-bottom:1px solid rgba(255,255,255,.08); animation: tvFade .45s ease both;}
  .lb-pos {min-width:54px; text-align:center; font-weight:800;}
  .lb-name {flex:0 0 210px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
            font-weight:700;}
  .lb-bar {flex:1; height:30px; background:rgba(255,255,255,.12); border-radius:10px;
           overflow:hidden; position:relative;}
  .lb-fill {display:block; height:100%; border-radius:10px;}
  .lb-pct {position:absolute; left:12px; top:50%; transform:translateY(-50%);
           font-size:18px; font-weight:800; color:#10243f; z-index:2;}
  .lb-val {flex:0 0 150px; text-align:right; font-weight:800; white-space:nowrap;}
  .lb-val small {font-size:16px; opacity:.6; font-weight:600;}

  /* Sanfter Einblende-Effekt bei jedem Live-Refresh (Kacheln + Ranking). */
  @keyframes tvFade {from {opacity:0; transform:translateY(10px);} to {opacity:1; transform:none;}}
  @keyframes tvPop  {from {opacity:0; transform:scale(.95);} to {opacity:1; transform:scale(1);}}
  .tv-cards {animation: tvFade .45s ease both;}
  .tv-card  {animation: tvPop .5s cubic-bezier(.2,.7,.3,1) both;}
  .tv-card:nth-child(2){animation-delay:.06s;}
  .tv-card:nth-child(3){animation-delay:.12s;}
  .tv-card:nth-child(4){animation-delay:.18s;}
  .tv-card:nth-child(5){animation-delay:.24s;}
  .tv-rank-title {animation: tvFade .45s ease both;}
  .tv-grid {animation: tvFade .5s ease both;}
  .tv-rk {animation: tvFade .45s ease both;}
  .tv-rk:nth-child(2){animation-delay:.04s;}
  .tv-rk:nth-child(3){animation-delay:.08s;}
  .tv-rk:nth-child(4){animation-delay:.12s;}
  .tv-rk:nth-child(5){animation-delay:.16s;}
  .tv-rk:nth-child(6){animation-delay:.20s;}
  .tv-rk:nth-child(7){animation-delay:.24s;}
  .tv-rk:nth-child(8){animation-delay:.28s;}
  .tv-rk:nth-child(9){animation-delay:.32s;}
  .tv-rk:nth-child(10){animation-delay:.36s;}
</style>
"""


def _spot_tv_config():
    """Liest die Spot-TV-Parameter aus der URL. None, wenn kein TV-Modus."""
    qp = st.query_params
    if "tv" not in qp:
        return None
    try:
        trust = float(qp.get("trust", "0") or 0)
    except ValueError:
        trust = 0.0
    return {
        "spot": qp.get("spot", ""),
        "mode": qp.get("mode", "today"),       # today | week | month | year
        "group": qp.get("group", ""),
        "sport": active_sport(),
        "sponsor": qp.get("sponsor", ""),
        "logo": qp.get("logo", ""),
        "event": qp.get("event", ""),
        "trust": trust,
        "base": qp.get("base", "https://mywatersessions.com"),
    }


@st.cache_data(ttl=3600, show_spinner=False)
def _all_spot_names():
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(spots_table.c.name).order_by(spots_table.c.name)
        ).all()
    return [r[0] for r in rows if r[0]]


@st.cache_data(ttl=30, show_spinner=False)
def _tv_load_sessions(sport):
    """Eigener Loader fuers Live-Dashboard mit kurzer TTL (30 s), damit neue
    Sessions (Uhr-Uploads, Demo) zeitnah erscheinen – der normale load_sessions
    cached 1 h und ist fuer einen Live-Screen ungeeignet. Ohne track-Spalte."""
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(*_SESSION_COLS_NO_TRACK).where(sessions_table.c.sport == sport)
        ).mappings().all()
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame([dict(r) for r in rows])
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce", format="mixed")
    return df


@st.cache_data(ttl=3600, show_spinner=False)
def _spot_coords(name):
    if not name:
        return None
    with get_engine().connect() as conn:
        row = conn.execute(
            select(spots_table.c.lat, spots_table.c.lon).where(spots_table.c.name == name)
        ).first()
    if row and row[0] is not None and row[1] is not None:
        return (float(row[0]), float(row[1]))
    return None


def _tv_card(label, value, sub=""):
    return (f"<div class='tv-card'><div class='lbl'>{label}</div>"
            f"<div class='val'>{value}</div><div class='sub'>{sub}</div></div>")


def _tv_period_scope(df, now, period):
    """Sessions im gewuenschten Zeitraum (Woche ab Montag / Monat / Jahr)."""
    if period == "year":
        start = now.normalize().replace(month=1, day=1)
    elif period == "month":
        start = now.normalize().replace(day=1)
    else:
        start = (now - pd.Timedelta(days=int(now.dayofweek))).normalize()
    return df[df["_date"] >= start]


def _tv_leaderboard(scope, metric):
    """Balken-Leaderboard (Top 10, best <metric> je Fahrer), 2-spaltig:
    Plaetze 1..N/2 links, der Rest rechts. metric = '1s' oder '30s'."""
    kmh_col = "speed_1s_kmh" if metric == "1s" else "speed_30s_kmh"
    if scope is None or scope.empty or kmh_col not in scope.columns:
        return "<div class='tv-msg'>No entries yet.</div>"
    t = scope.copy()
    t["_v"] = pd.to_numeric(t[kmh_col], errors="coerce")
    g = (t.groupby("name", as_index=False).agg(v=("_v", "max"))
         .dropna(subset=["v"]).sort_values("v", ascending=False)
         .reset_index(drop=True).head(10))
    if g.empty:
        return "<div class='tv-msg'>No entries yet.</div>"

    maxv = float(g["v"].iloc[0]) or 1.0
    medals = {1: "🥇", 2: "🥈", 3: "🥉"}
    colors = {1: "#f4c430", 2: "#cfd4da", 3: "#cd7f32"}
    rows = []
    for i, (_, r) in enumerate(g.iterrows(), start=1):
        pos = medals.get(i, str(i))
        col = colors.get(i, "#3aa0ff")
        pct = round(r["v"] / maxv * 100)   # Platz 1 = 100 %, Rest anteilig
        w = max(5, pct)
        rows.append(
            f"<div class='lb-row'><span class='lb-pos'>{pos}</span>"
            f"<span class='lb-name'>{r['name']}</span>"
            f"<span class='lb-bar'><span class='lb-fill' style='width:{w}%;background:{col}'></span>"
            f"<span class='lb-pct'>{pct}%</span></span>"
            f"<span class='lb-val'>{r['v']:.1f}<small> km/h</small></span></div>"
        )
    half = (len(rows) + 1) // 2  # 8 -> links 1-4 / rechts 5-8; 10 -> 1-5 / 6-10
    left = "".join(rows[:half])
    right = "".join(rows[half:])
    return (f"<div class='lb' translate='no'><div class='lb-col'>{left}</div>"
            f"<div class='lb-col'>{right}</div></div>")


def _join_url(cfg):
    params = {"sport": cfg["sport"], "spot": cfg["spot"], "join": "1"}
    return cfg["base"].rstrip("/") + "/?" + urlencode(params)


def _render_join_qr(cfg):
    url = _join_url(cfg)

    # QR als data-URI (statt st.image), damit er in DERSELBEN Flex-Reihe wie die
    # Produktkarten sitzt -> Produkte direkt rechts neben dem QR, dann Umbruch.
    if QR_AVAILABLE:
        img = qrcode.make(url)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        qr_uri = "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii")
        qr_html = f"<div class='tv-join-qr'><img src='{qr_uri}' alt='QR'/></div>"
    else:
        qr_html = f"<div class='tv-join-qr tv-join-qr-text'>{url}</div>"

    cards = _product_cards_html(cfg["spot"])
    deals_html = ""
    if cards:
        ad = load_spot_ad(cfg["spot"]) or {}
        sponsor = ad.get("sponsor_name")
        head = f"{sponsor} · Top Deals" if sponsor else "Top Deals"
        deals_html = (
            "<div class='tv-deals'>"
            f"<div class='tv-deals-head'>{head}</div>"
            f"<div class='tv-deals-cards'>{''.join(cards)}</div>"
            "</div>"
        )

    st.markdown(
        "<style>"
        ".tv-join-row{display:flex;gap:22px;align-items:flex-start;margin-top:6px;}"
        # Titel + QR sind eine eigene Spalte, die um 1,5 cm nach unten versetzt ist
        # (Produkte bleiben oben). gap steuert den Abstand Titel<->QR.
        ".tv-join-qcol{display:flex;flex-direction:column;gap:0.35cm;margin-top:1.5cm;}"
        ".tv-join-title{font-size:28px;font-weight:800;margin:0;line-height:1.1;}"
        ".tv-join-qr{flex:0 0 auto;align-self:flex-start;background:#fff;"
        "border-radius:16px;padding:10px;line-height:0;box-shadow:0 6px 18px rgba(0,0,0,.18);}"
        ".tv-join-qr img{width:200px;height:200px;display:block;}"
        ".tv-join-qr-text{line-height:1.3;padding:14px;color:#111;max-width:220px;"
        "word-break:break-all;font-size:13px;}"
        ".tv-deals{flex:1 1 auto;min-width:0;display:flex;flex-direction:column;"
        "justify-content:space-between;}"
        ".tv-deals-head{text-align:center;font-size:22px;font-weight:800;color:#eaf4ff;"
        "margin:0 0 10px;}"
        ".tv-deals-cards{display:flex;flex-wrap:wrap;gap:16px;justify-content:center;}"
        ".tv-prod-card{width:184px;background:#ffffff;border-radius:16px;overflow:hidden;"
        "box-shadow:0 6px 18px rgba(0,0,0,.18);text-decoration:none;color:#111;display:block;}"
        ".tv-prod-img{height:172px;background-size:cover;background-position:center;}"
        ".tv-prod-noimg{display:flex;align-items:center;justify-content:center;"
        "font-size:48px;background:#eef2f6;}"
        ".tv-prod-title{padding:9px 12px 2px;font-weight:700;font-size:16px;line-height:1.18;}"
        ".tv-prod-price{padding:0 12px 11px;color:#0a7;font-weight:800;font-size:17px;}"
        "</style>"
        "<div class='tv-join-row'>"
        "<div class='tv-join-qcol'>"
        "<div class='tv-join-title'>📲 Join today’s ranking</div>"
        f"{qr_html}"
        "</div>"
        f"{deals_html}"
        "</div>",
        unsafe_allow_html=True,
    )


def _tv_weather_html(lat, lon):
    html = """
<!DOCTYPE html><html><head><meta charset='utf-8'><style>
 html,body{background:transparent!important;margin:0;color:#eaf4ff;
   font-family:system-ui,-apple-system,sans-serif;}
 .row{display:flex;gap:16px;}
 .c{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.18);
    border-radius:18px;padding:10px 20px;flex:1;}
 .c .l{font-size:17px;opacity:.8;} .c .v{font-size:44px;font-weight:800;}
</style></head><body>
<div id='w' class='row'>…</div>
<script>
const LAT=__LAT__,LON=__LON__;
const C=["N","NNE","NE","ENE","E","ESE","SE","SSE","S","SSW","SW","WSW","W","WNW","NW","NNW"];
const W={0:"☀️",1:"🌤️",2:"⛅",3:"☁️",45:"🌫️",51:"🌦️",61:"🌦️",63:"🌧️",65:"🌧️",80:"🌦️",81:"🌧️",95:"⛈️"};
function comp(d){return d==null?"":C[Math.round(d/22.5)%16];}
function load(){
fetch("https://api.open-meteo.com/v1/forecast?latitude="+LAT+"&longitude="+LON+"&current=temperature_2m,weather_code,wind_speed_10m,wind_gusts_10m,wind_direction_10m&wind_speed_unit=kmh&timezone=auto")
.then(r=>r.json()).then(d=>{const c=d.current||{};
 document.getElementById('w').innerHTML=
 "<div class='c'><div class='l'>🌬️ Wind</div><div class='v'>"+Math.round(c.wind_speed_10m)+" km/h</div><div class='l'>Gusts "+Math.round(c.wind_gusts_10m)+" km/h · "+comp(c.wind_direction_10m)+"</div></div>"+
 "<div class='c'><div class='l'>🌡️ Temperature</div><div class='v'>"+(Math.round(c.temperature_2m*10)/10)+" °C</div><div class='l'>"+(W[c.weather_code]||"")+"</div></div>";
}).catch(e=>{document.getElementById('w').innerHTML="<div class='c'>Weather unavailable</div>";});
}
load(); setInterval(load, 600000);
</script></body></html>
"""
    return html.replace("__LAT__", str(lat)).replace("__LON__", str(lon))


@st.fragment(run_every=30)
def _spot_tv_live(cfg):
    """Dynamischer Teil des Dashboards – aktualisiert sich alle 30 s selbst."""
    sport = cfg["sport"]
    spot = cfg["spot"]

    df = _tv_load_sessions(sport)
    if df is None or df.empty or "surfspot" not in df.columns:
        st.markdown("<div class='tv-msg'>No sessions yet for this spot.</div>", unsafe_allow_html=True)
        return

    df = df[df["surfspot"].astype(str) == spot].copy()

    # Trust-Filter: NULL (Uhr-Sessions ohne Trust) gilt als ok.
    if cfg["trust"] > 0 and "trust_score" in df.columns:
        ts = pd.to_numeric(df["trust_score"], errors="coerce")
        df = df[ts.isna() | (ts >= cfg["trust"])]

    df["_date"] = pd.to_datetime(df.get("date"), errors="coerce", format="mixed")
    if "created_at" in df.columns:
        df["_created"] = pd.to_datetime(df["created_at"], errors="coerce")

    now = pd.Timestamp(datetime.now())
    today = now.normalize()
    today_df = df[df["_date"].dt.normalize() == today]

    # Nur VOLLSTAENDIGE Sessions (Spot+Board+Segel) zaehlen fuer Bestzeiten/Rangliste.
    ranked = complete_sessions(df)

    # Zeitraum (mode) bestimmt Header-Kacheln UND Leaderboard.
    mode = cfg["mode"]
    period_word = {"today": "today", "week": "this week",
                   "month": "this month", "year": "this year"}.get(mode, "today")
    if mode == "today":
        scope = ranked[ranked["_date"].dt.normalize() == today] if not ranked.empty else ranked
        scope_title = "Today"
    else:
        scope = _tv_period_scope(ranked, now, mode)
        scope_title = {"week": "This week", "month": "This month",
                       "year": "This year"}.get(mode, mode.capitalize())

    if cfg["group"]:
        gid = next((g["id"] for g in list_groups() if g["name"] == cfg["group"]), None)
        members = set(group_member_names(gid)) if gid else set()
        scope = scope[scope["name"].astype(str).isin(members)]
        scope_title += f" · {cfg['group']}"

    def _mx(d, col):
        if col not in d.columns:
            return None
        v = pd.to_numeric(d[col], errors="coerce").max()
        return None if pd.isna(v) else float(v)

    # Header-Bestwerte aus dem gewaehlten Zeitraum.
    top1 = _mx(scope, "speed_1s_kmh")
    top1kn = _mx(scope, "speed_1s_kn")
    top30 = _mx(scope, "speed_30s_kmh")

    leader = "–"
    if not scope.empty and "speed_1s_kmh" in scope.columns:
        t = scope.copy()
        t["_s"] = pd.to_numeric(t["speed_1s_kmh"], errors="coerce")
        t = t.dropna(subset=["_s"])
        if not t.empty:
            leader = str(t.loc[t["_s"].idxmax(), "name"])

    # Aktivitaets-Kacheln zaehlen ALLE heutigen Sessions (auch unvollstaendige).
    n_sessions = len(today_df)
    n_riders = today_df["name"].nunique() if "name" in today_df.columns else 0
    last_txt = "–"
    if "_created" in today_df.columns and not today_df["_created"].dropna().empty:
        mins = int(max(0, (now - today_df["_created"].max()).total_seconds() // 60))
        last_txt = "just now" if mins == 0 else f"{mins} min ago"

    leader_label = "👑 Rider of the Day" if mode == "today" else f"👑 Top rider {period_word}"
    cards = "".join([
        _tv_card(f"🏆 Top 2s {period_word}", f"{top1:.1f}" if top1 else "–",
                 ("km/h" + (f" · {top1kn:.1f} kn" if top1kn else "")) if top1 else ""),
        _tv_card(f"🔥 Top 30s {period_word}", f"{top30:.1f}" if top30 else "–", "km/h" if top30 else ""),
        _tv_card("🏄 Sessions today", f"{n_sessions}", f"{n_riders} riders"),
        _tv_card(leader_label, leader),
        _tv_card("🧭 Last activity", last_txt),
    ])
    rk = now.strftime("%H%M%S")  # wechselt je Refresh -> erzwingt Re-Mount (Animation)
    st.markdown(f"<div class='tv-cards' translate='no' data-r='{rk}'>{cards}</div>",
                unsafe_allow_html=True)

    # Leaderboard zeigt 1s ODER 30s und wechselt automatisch (~alle 30 s).
    metric = "30s" if (int(now.timestamp()) // 30) % 2 else "1s"
    metric_lbl = "Top 2 s" if metric == "1s" else "Top 30 s"
    st.markdown(
        f"<div class='tv-rank-title' translate='no'>🏁 {scope_title} leaderboard · {metric_lbl}</div>",
        unsafe_allow_html=True)
    st.markdown(f"<div data-r='{rk}{metric}'>" + _tv_leaderboard(scope, metric) + "</div>",
                unsafe_allow_html=True)

    st.markdown(f"<div class='tv-update' translate='no'>⏱️ Last update: {now.strftime('%H:%M')} "
                f"· auto-refresh 30 s · switches 2 s / 30 s</div>", unsafe_allow_html=True)


# Sponsor je Spot. logo = Dateiname in assets/, name/url optional.
# Erweiterbar: hier Eintraege ergaenzen ODER einfach assets/sponsor_<slug>.png
# ablegen (slug = klein, Leerzeichen->_, Umlaute ae/oe/ue, z.B. "Strand Horst"
# -> sponsor_strand_horst.png).
SPOT_SPONSORS = {
    "Strand Horst": {"logo": "tv_logo.png", "name": "Telstar Surf",
                     "url": "https://www.telstar-surf.com"},
    # "Brouwersdam": {"logo": "sponsor_brouwersdam.png", "name": "Surfshop XY"},
}


def _slug(s):
    s = (s or "").lower().strip()
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss")):
        s = s.replace(a, b)
    out = "".join(ch if ch.isalnum() else "_" for ch in s)
    while "__" in out:
        out = out.replace("__", "_")
    return out.strip("_")


def _spot_sponsor_img(cfg):
    """(img_src, name) fuer die Sponsor-Anzeige des Spots, sonst (None, name).
    Prioritaet: ?logo= URL > DB (Admin-Backoffice) > SPOT_SPONSORS > sponsor_<slug>.png."""
    if cfg["logo"]:
        return cfg["logo"], (cfg["sponsor"] or None)

    # Im Backoffice gepflegter Eintrag hat Vorrang (sofern aktiv).
    ad = load_spot_ad(cfg["spot"])
    if ad and ad.get("active", True):
        name = ad.get("sponsor_name") or (cfg["sponsor"] or None)
        uri = _bytes_to_data_uri(ad.get("logo"), ad.get("logo_mime"))
        if uri:
            return uri, name
        if name:
            return None, name

    entry = SPOT_SPONSORS.get(cfg["spot"], {})
    name = entry.get("name") or (cfg["sponsor"] or None)
    logo_file = entry.get("logo")
    if not logo_file:
        cand = "sponsor_" + _slug(cfg["spot"]) + ".png"
        if os.path.exists(app_path("assets", cand)):
            logo_file = cand
    if logo_file:
        b64 = image_to_base64(app_path("assets", logo_file))
        if b64:
            return f"data:image/png;base64,{b64}", name
    return None, name


def _spot_tv_controls(cfg):
    """Auf dem Screen verdrahtete Bedienung: Zeitraum, Spot, Gruppe, Exit.
    Aenderungen schreiben in die URL (?mode/?spot/?group) -> bookmarkbar."""
    modes = [("today", "Today"), ("week", "Week"), ("month", "Month"), ("year", "Year")]
    spots = _all_spot_names()
    spot_opts = list(spots)
    if not cfg["spot"]:
        spot_opts = ["– Spot –"] + spot_opts
    elif cfg["spot"] not in spot_opts:
        spot_opts = [cfg["spot"]] + spot_opts
    groups = ["All"] + [g["name"] for g in list_groups()]

    cols = st.columns([1, 1, 1, 1, 2, 2, 1.3])
    for i, (mkey, mlabel) in enumerate(modes):
        if cols[i].button(mlabel, key=f"tvmode_{mkey}", use_container_width=True,
                          type="primary" if cfg["mode"] == mkey else "secondary"):
            st.query_params["mode"] = mkey
            st.rerun()

    with cols[4]:
        idx = spot_opts.index(cfg["spot"]) if cfg["spot"] in spot_opts else 0
        sel = st.selectbox("Spot", spot_opts, index=idx,
                           label_visibility="collapsed", key="tv_spot_sel")
        if sel and sel != "– Spot –" and sel != cfg["spot"]:
            st.query_params["spot"] = sel
            st.rerun()

    with cols[5]:
        gidx = groups.index(cfg["group"]) if cfg["group"] in groups else 0
        gsel = st.selectbox("Group", groups, index=gidx,
                            label_visibility="collapsed", key="tv_group_sel")
        newg = "" if gsel == "All" else gsel
        if newg != cfg["group"]:
            if newg:
                st.query_params["group"] = newg
            elif "group" in st.query_params:
                del st.query_params["group"]
            st.rerun()

    with cols[6]:
        if st.button("← Exit", use_container_width=True, key="tv_exit"):
            for k in ("tv", "mode", "group"):
                if k in st.query_params:
                    del st.query_params[k]
            st.rerun()


def render_spot_tv(cfg):
    """Vollbild-Dashboard fuer einen Spot. Statischer Kopf + Live-Fragment."""
    st.markdown(_TV_CSS, unsafe_allow_html=True)

    # Browser-Auto-Uebersetzung (z.B. Chrome -> Deutsch) fuer die ganze Seite
    # abschalten, damit die Beschriftungen englisch bleiben UND die 30s-Refreshes
    # stabil laufen (uebersetzte Textknoten -> React removeChild-Fehler).
    components.html(
        """
        <script>
          try {
            var d = window.parent.document;
            d.documentElement.setAttribute('translate', 'no');
            d.documentElement.classList.add('notranslate');
          } catch (e) {}
        </script>
        """,
        height=0,
    )

    event = f"<div class='event'>🏁 {cfg['event']}</div>" if cfg["event"] else ""
    title = cfg["spot"] or "Spot TV"

    # Sponsor/Werbung oben rechts, je Spot (auf hellem Chip, damit dunkle Logos
    # sichtbar bleiben). Logo wenn vorhanden, sonst "Presented by <Name>".
    ad_src, ad_name = _spot_sponsor_img(cfg)
    ad_row = load_spot_ad(cfg["spot"]) or {}
    ad_url = ad_row.get("sponsor_url") or ""
    if ad_src:
        chip = f"<div class='tv-sponsor-chip'><img src='{ad_src}' alt='sponsor'/></div>"
        sponsor = f"<a href='{ad_url}' target='_blank' rel='noopener'>{chip}</a>" if ad_url else chip
    elif ad_name:
        sponsor = f"<div class='tv-presented'>Presented by <b>{ad_name}</b></div>"
    else:
        sponsor = ""

    st.markdown(
        "<div class='tv-header'>"
        f"<div class='tv-brand'>MyWaterSessions<span class='dot'>.</span>{BETA_BADGE}</div>"
        f"<div class='tv-spot'><div class='name'>{title}</div>{event}</div>"
        f"<div class='sponsor'>{sponsor}</div>"
        "</div>",
        unsafe_allow_html=True,
    )

    _spot_tv_controls(cfg)

    if not cfg["spot"]:
        st.markdown("<div class='tv-msg'>Add <code>&spot=YourSpot</code> to the URL "
                    "to choose a spot for this screen.</div>", unsafe_allow_html=True)
        return

    # Wetter (iFrame) statisch zeichnen – NICHT im 30s-Refresh, sonst wird der
    # iFrame staendig neu erzeugt (React removeChild-Fehler). Es aktualisiert
    # sich ueber sein eigenes JS.
    coords = _spot_coords(cfg["spot"])
    if coords:
        components.html(_tv_weather_html(coords[0], coords[1]), height=130)

    _spot_tv_live(cfg)

    # QR-Code + beworbene Produkte zusammen in EINER umbrechenden Reihe (statisch).
    _render_join_qr(cfg)

    # Ganz unten: wechselt synchron zum Ranking (2s/30s) zwischen Spot-Info
    # und 3-Tage-Wettervorhersage.
    _tv_bottom_info(cfg)


def _tv_spot_info(cfg):
    """Spot-Infobereich am unteren Rand des TV: Text + Webcam/Bild."""
    info = load_spot_info(cfg["spot"])
    if not info:
        return
    desc = (info.get("description") or "").strip()
    webcam = (info.get("webcam_url") or "").strip()
    img_ids = load_spot_image_ids(cfg["spot"])
    if img_ids:
        img_uri = _spot_thumb_uri(img_ids[0], max_dim=520)
    else:
        img_uri = _bytes_to_data_uri(info.get("image"), info.get("image_mime"))
    if not desc and not webcam and not img_uri:
        return

    st.markdown(
        "<style>"
        ".tv-info-title{font-size:26px;font-weight:800;margin:30px 0 10px;}"
        ".tv-info-text{font-size:20px;line-height:1.5;opacity:.92;}"
        ".tv-info-img{width:100%;border-radius:16px;box-shadow:0 6px 18px rgba(0,0,0,.18);"
        "display:block;}"
        "</style>"
        f"<div class='tv-info-title'>ℹ️ {cfg['spot']}</div>",
        unsafe_allow_html=True,
    )

    col_text, col_media = st.columns([1.3, 1])
    with col_text:
        if desc:
            st.markdown(f"<div class='tv-info-text'>{desc}</div>", unsafe_allow_html=True)
    with col_media:
        if webcam and _is_image_url(webcam):
            # Snapshot-Webcam: Bild alle 60 s mit Cache-Buster neu laden.
            components.html(
                f"<img id='cam' src='{webcam}' "
                "style='width:100%;height:240px;object-fit:cover;border-radius:16px;display:block;'>"
                "<script>setInterval(function(){var c=document.getElementById('cam');"
                "var u=c.src.split('?')[0];c.src=u+'?t='+Date.now();},60000);</script>",
                height=248,
            )
        elif webcam:
            # Einbettbare Seite (YouTube-Live/Windy/…): als iFrame.
            components.html(
                f"<iframe src='{webcam}' allow='autoplay; fullscreen' "
                "style='width:100%;height:240px;border:0;border-radius:16px;'></iframe>",
                height=248,
            )
        elif img_uri:
            st.markdown(
                f"<img class='tv-info-img' src='{img_uri}' alt='Spot'>",
                unsafe_allow_html=True,
            )


_WCODE_EMOJI = {
    0: "☀️", 1: "🌤️", 2: "⛅", 3: "☁️", 45: "🌫️", 48: "🌫️", 51: "🌦️", 53: "🌦️",
    55: "🌦️", 61: "🌦️", 63: "🌧️", 65: "🌧️", 71: "🌨️", 73: "🌨️", 75: "❄️",
    80: "🌦️", 81: "🌧️", 82: "⛈️", 85: "🌨️", 86: "❄️", 95: "⛈️", 96: "⛈️", 99: "⛈️",
}
_COMPASS = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
            "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"]
_WEEKDAY_EN = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


_DIR_DEG = {
    # Englische Kompass-Kürzel (so liefert die KI) ...
    "N": 0, "NNE": 22.5, "NE": 45, "ENE": 67.5, "E": 90, "ESE": 112.5, "SE": 135,
    "SSE": 157.5, "S": 180, "SSW": 202.5, "SW": 225, "WSW": 247.5, "W": 270,
    "WNW": 292.5, "NW": 315, "NNW": 337.5,
    # ... plus deutsche Varianten (Ost/Südost): O, OSO, SO, SSO, NO, ONO, NNO.
    "O": 90, "ONO": 67.5, "OSO": 112.5, "SO": 135, "SSO": 157.5, "NO": 45, "NNO": 22.5,
}


def _parse_best_dirs(text):
    """'SW, W, NW' / 'SW-NW' -> Liste von Grad-Mittelwerten. Unbekanntes wird ignoriert."""
    if not text:
        return []
    degs = []
    for tok in re.split(r"[^A-Za-zÄÖÜäöü]+", str(text)):
        d = _DIR_DEG.get(tok.strip().upper())
        if d is not None and d not in degs:
            degs.append(d)
    return degs


def _ang_diff(a, b):
    d = abs((a - b) % 360)
    return min(d, 360 - d)


def _assess_forecast_day(wind, dir_deg, best_degs):
    """(emoji, kurztext) – bewertet einen Vorhersagetag fuer Wind-Wassersport.
    Windstaerke (km/h) x Richtungs-Match mit den besten Windrichtungen des Spots."""
    if wind is None:
        return "", ""
    if wind < 12:
        return "🔴", "too little wind"
    if wind > 60:
        return "🔴", "too strong"
    dir_ok = None  # unknown
    if best_degs and dir_deg is not None:
        dir_ok = any(_ang_diff(dir_deg, b) <= 34 for b in best_degs)
    if dir_ok is False:
        return "🟡", "wrong direction"
    if 16 <= wind <= 45:
        return "🟢", "worth it"
    if wind > 45:
        return "🟡", "lots of wind (pros)"
    return "🟡", "light wind"


@st.cache_resource(show_spinner=False)
def _forecast3d_store():
    """Letzter erfolgreicher 3-Tage-Abruf je (Spot, Block). Überlebt Reruns und
    Sessions (cache_resource) → dient als Fallback, wenn ein frischer Abruf an
    einem 429 scheitert. So bleiben die Kacheln stehen statt zu flackern."""
    return {}


def _fetch_open_meteo_block(params, block, attempts=3):
    """Holt einen Open-Meteo-Block ('daily'/'hourly') robust.

    _http_get_json wirft bei 4xx sofort – inkl. 429 (Rate-Limit der geteilten
    Render-IP). Genau das ist hier aber der häufigste, vorübergehende Fehler,
    daher wird 429 ein paar Mal mit kurzer Pause wiederholt. Wirft, wenn alle
    Versuche scheitern (der Aufrufer fällt dann auf den letzten Stand zurück)."""
    url = _open_meteo_url("api", "/v1/forecast", params)
    last = None
    for i in range(attempts):
        try:
            return (_http_get_json(url, timeout=20) or {}).get(block)
        except HTTPError as e:
            last = e
            if e.code == 429 and i < attempts - 1:
                time.sleep(1.2 * (i + 1))
                continue
            if 400 <= e.code < 500:   # andere 4xx sind dauerhaft
                raise
        except Exception as e:  # noqa: BLE001
            last = e
            if i < attempts - 1:
                time.sleep(1.2 * (i + 1))
    raise last if last else RuntimeError("forecast fetch failed")


@st.cache_data(ttl=1800, show_spinner=False)
def _fetch_forecast_3d_raw(lat, lon):
    """3-Tage-Tagesvorhersage – cached NUR bei Erfolg (wirft sonst, damit
    st.cache_data den Fehler nicht 30 Min festhält)."""
    return _fetch_open_meteo_block({
        "latitude": round(float(lat), 4),
        "longitude": round(float(lon), 4),
        "daily": (
            "weather_code,temperature_2m_max,temperature_2m_min,"
            "wind_speed_10m_max,wind_gusts_10m_max,wind_direction_10m_dominant"
        ),
        "wind_speed_unit": "kmh",
        "timezone": "auto",
        "forecast_days": 3,
    }, "daily")


@st.cache_data(ttl=1800, show_spinner=False)
def _fetch_hourly_forecast_raw(lat, lon):
    """Stuendliche 3-Tage-Vorhersage – cached nur bei Erfolg (wirft sonst)."""
    return _fetch_open_meteo_block({
        "latitude": round(float(lat), 4),
        "longitude": round(float(lon), 4),
        "hourly": "wind_speed_10m,wind_gusts_10m,wind_direction_10m",
        "wind_speed_unit": "kmh",
        "timezone": "auto",
        "forecast_days": 3,
    }, "hourly")


def _forecast_with_fallback(kind, fetch, lat, lon):
    """Erfolg cachen (über fetch) + letzten guten Stand merken; bei Fehler den
    letzten Stand zurückgeben statt None. Behebt das „mal da, mal nicht"."""
    store = _forecast3d_store()
    key = (kind, round(float(lat), 4), round(float(lon), 4))
    try:
        result = fetch(lat, lon)
        if result:
            store[key] = result
        return result
    except Exception:  # noqa: BLE001
        return store.get(key)


def _fetch_forecast_3d(lat, lon):
    return _forecast_with_fallback("daily", _fetch_forecast_3d_raw, lat, lon)


def _fetch_hourly_forecast(lat, lon):
    return _forecast_with_fallback("hourly", _fetch_hourly_forecast_raw, lat, lon)


def _wind_color(v):
    """Fliessender Farbverlauf fuer die Windstaerke (km/h):
    gelb (zu wenig) -> gruen (gut) -> orange/rot (zu stark). Gibt 'r,g,b'."""
    stops = [
        (0, (245, 197, 24)),    # gelb – zu wenig Wind
        (15, (46, 204, 113)),   # gruen – guter Wind beginnt
        (42, (46, 204, 113)),   # gruen – Sweet Spot
        (55, (243, 156, 18)),   # orange – viel Wind
        (70, (231, 76, 60)),    # rot – zu stark/Sturm
    ]
    if v <= stops[0][0]:
        r, g, b = stops[0][1]
        return f"{r},{g},{b}"
    if v >= stops[-1][0]:
        r, g, b = stops[-1][1]
        return f"{r},{g},{b}"
    for (x0, c0), (x1, c1) in zip(stops, stops[1:]):
        if x0 <= v <= x1:
            t = (v - x0) / (x1 - x0) if x1 > x0 else 0
            r, g, b = (round(c0[k] + (c1[k] - c0[k]) * t) for k in range(3))
            return f"{r},{g},{b}"
    r, g, b = stops[-1][1]
    return f"{r},{g},{b}"


_FORECAST_CSS = (
    "<style>"
    ".tv-fc-row{display:flex;gap:18px;flex-wrap:wrap;}"
    ".tv-fc-card{flex:1 1 0;min-width:160px;background:rgba(255,255,255,.08);"
    "border:1px solid rgba(255,255,255,.18);border-radius:18px;padding:14px 18px;text-align:center;}"
    ".tv-fc-day{font-size:22px;font-weight:800;opacity:.9;}"
    ".tv-fc-emoji{font-size:46px;line-height:1.1;margin:2px 0;}"
    ".tv-fc-temp{font-size:30px;font-weight:800;}"
    ".tv-fc-temp span{font-size:20px;font-weight:600;opacity:.6;}"
    ".tv-fc-wind{font-size:26px;font-weight:800;color:#7fd4ff;margin-top:4px;}"
    ".tv-fc-wind small{font-size:15px;opacity:.7;font-weight:600;}"
    ".tv-fc-gust{font-size:16px;opacity:.75;}"
    ".tv-fc-rate{margin-top:8px;font-size:18px;font-weight:800;}"
    "</style>"
)


def _forecast_card_html(daily, i, best_degs):
    """HTML einer einzelnen Tageskachel (Emoji/Temp/Wind/Ampel)."""
    def at(key):
        arr = daily.get(key) or []
        return arr[i] if i < len(arr) else None

    def num(x):
        return str(round(x)) if x is not None else "–"

    try:
        label = "Today" if i == 0 else _WEEKDAY_EN[pd.to_datetime(daily["time"][i]).weekday()]
    except Exception:  # noqa: BLE001
        label = str(daily["time"][i])
    wspeed = at("wind_speed_10m_max")
    wdir = at("wind_direction_10m_dominant")
    comp = _COMPASS[round(wdir / 22.5) % 16] if wdir is not None else ""
    emoji, verdict = _assess_forecast_day(wspeed, wdir, best_degs)
    rate = f"<div class='tv-fc-rate'>{emoji} {verdict}</div>" if emoji else ""
    return (
        "<div class='tv-fc-card'>"
        f"<div class='tv-fc-day'>{label}</div>"
        f"<div class='tv-fc-emoji'>{_WCODE_EMOJI.get(at('weather_code'), '')}</div>"
        f"<div class='tv-fc-temp'>{num(at('temperature_2m_max'))}°"
        f"<span> / {num(at('temperature_2m_min'))}°</span></div>"
        f"<div class='tv-fc-wind'>🌬️ {num(wspeed)}<small> km/h</small></div>"
        f"<div class='tv-fc-gust'>Gusts {num(at('wind_gusts_10m_max'))} · {comp}</div>"
        f"{rate}"
        "</div>"
    )


def _forecast_title(spot, info):
    bw = info.get("best_winds")
    bw_note = f" · best winds: {bw}" if bw else " · <i>best wind directions not set yet</i>"
    return (f"<div class='tv-info-title'>🌤️ Forecast · {spot}<span "
            f"style='font-size:16px;font-weight:600;opacity:.7;'>{bw_note}</span></div>")


def _tv_forecast(cfg, coords):
    """3-Tage-Vorhersage als HTML-Karten (Spot-TV, nicht interaktiv)."""
    daily = _fetch_forecast_3d(coords[0], coords[1])
    if not daily or not daily.get("time"):
        return
    info = load_spot_info(cfg["spot"]) or {}
    best_degs = _parse_best_dirs(info.get("best_winds"))
    cards = "".join(_forecast_card_html(daily, i, best_degs)
                    for i in range(min(3, len(daily["time"]))))
    st.markdown(
        _FORECAST_CSS + _forecast_title(cfg["spot"], info)
        + f"<div class='tv-fc-row'>{cards}</div>",
        unsafe_allow_html=True,
    )


def render_spots_forecast(spot, coords):
    """Spots-Seite: 3-Tage-Vorhersage als Kacheln mit Button je Tag -> Klick
    oeffnet die Stundenansicht per Streamlit-Rerun (KEIN voller Seiten-Reload)."""
    daily = _fetch_forecast_3d(coords[0], coords[1])
    if not daily or not daily.get("time"):
        st.caption("⛅ Forecast is temporarily unavailable – please try again in a moment.")
        return
    info = load_spot_info(spot) or {}
    best_degs = _parse_best_dirs(info.get("best_winds"))
    n = min(3, len(daily["time"]))

    st.markdown(_FORECAST_CSS + _forecast_title(spot, info), unsafe_allow_html=True)

    cols = st.columns(n)
    for i in range(n):
        with cols[i]:
            st.markdown(_forecast_card_html(daily, i, best_degs), unsafe_allow_html=True)
            try:
                lbl = "Today" if i == 0 else _WEEKDAY_EN[pd.to_datetime(daily["time"][i]).weekday()]
            except Exception:  # noqa: BLE001
                lbl = f"Day {i + 1}"
            if st.button(f"🕘 {lbl} · hourly", key=f"fcbtn_{spot}_{i}", use_container_width=True):
                st.session_state["spots_fcday"] = i
                st.rerun()

    di = st.session_state.get("spots_fcday")
    if di is not None and di < n:
        _render_hourly(spot, coords, di)
        if st.button("✕ Close hourly view", key="close_hourly"):
            st.session_state.pop("spots_fcday", None)
            st.rerun()


def _render_hourly(spot, coords, day_index):
    """Stundenansicht (9–21 Uhr) eines Tages als Balken: Hoehe=Windstaerke,
    Farbe fliessend (gelb zu wenig -> gruen gut -> rot zu stark), Zahl in km/h."""
    daily = _fetch_forecast_3d(coords[0], coords[1])
    hourly = _fetch_hourly_forecast(coords[0], coords[1])
    if not daily or not hourly or not daily.get("time") or day_index >= len(daily["time"]):
        return
    target = daily["time"][day_index]   # 'YYYY-MM-DD'
    try:
        dt = pd.to_datetime(target)
        day_label = "Today" if day_index == 0 else _WEEKDAY_EN[dt.weekday()]
    except Exception:  # noqa: BLE001
        day_label = str(target)

    times = hourly.get("time") or []
    ws = hourly.get("wind_speed_10m") or []
    gs = hourly.get("wind_gusts_10m") or []
    ds = hourly.get("wind_direction_10m") or []
    rows = []
    for i, t in enumerate(times):
        if t[:10] != target:
            continue
        hh = int(t[11:13])
        if hh < 9 or hh > 21:
            continue
        rows.append((
            hh,
            ws[i] if i < len(ws) else None,
            gs[i] if i < len(gs) else None,
            ds[i] if i < len(ds) else None,
        ))
    if not rows:
        st.info("No hourly data for this day.")
        return

    gvals = [g for _, _, g, _ in rows if g is not None]
    wvals = [w for _, w, _, _ in rows if w is not None]
    ref = max(70.0, max(gvals + wvals, default=0))   # auf max. Boe skalieren
    bars = []
    for hh, w, g, d in rows:
        wv = w or 0
        gv = g if g is not None else wv
        w_pct = max(4, min(100, wv / ref * 100))
        g_pct = max(w_pct, min(100, gv / ref * 100))   # Boe >= Wind -> hoeher
        comp = _COMPASS[round(d / 22.5) % 16] if d is not None else ""
        bars.append(
            "<div class='hb-col'>"
            f"<div class='hb-val'>{round(wv)}</div>"
            "<div class='hb-track'>"
            f"<div class='hb-gust' style='height:{g_pct}%'></div>"
            f"<div class='hb-bar' style='height:{w_pct}%;background:rgb({_wind_color(wv)})'></div>"
            "</div>"
            f"<div class='hb-hr'>{hh}h</div>"
            f"<div class='hb-dir'>{comp}</div>"
            f"<div class='hb-gust-val'>⤴ {round(gv)}</div>"
            "</div>"
        )

    st.markdown(
        "<style>"
        ".hb-wrap{background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.15);"
        "border-radius:18px;padding:16px 18px;margin-top:6px;}"
        ".hb-title{font-size:22px;font-weight:800;margin-bottom:4px;}"
        ".hb-legend{font-size:14px;opacity:.7;margin-bottom:12px;}"
        ".hb-row{display:flex;gap:6px;align-items:flex-end;}"
        ".hb-col{flex:1 1 0;text-align:center;}"
        ".hb-val{font-size:15px;font-weight:800;margin-bottom:4px;}"
        ".hb-track{position:relative;height:180px;}"
        # Gläserner Böen-Balken (frosted) hinter dem farbigen Wind-Balken.
        ".hb-gust{position:absolute;left:0;right:0;bottom:0;background:rgba(255,255,255,.16);"
        "border:1px solid rgba(255,255,255,.32);border-bottom:none;border-radius:9px 9px 0 0;"
        "backdrop-filter:blur(2px);-webkit-backdrop-filter:blur(2px);z-index:1;}"
        ".hb-bar{position:absolute;left:0;right:0;bottom:0;border-radius:7px 7px 0 0;"
        "min-height:4px;z-index:2;}"
        ".hb-hr{font-size:13px;opacity:.75;margin-top:5px;}"
        ".hb-dir{font-size:12px;opacity:.55;}"
        ".hb-gust-val{font-size:11px;opacity:.5;}"
        "</style>"
        "<div class='hb-wrap'>"
        f"<div class='hb-title'>⏱️ {spot} · {day_label} · 9–21h <small style='font-weight:600;"
        "opacity:.6;'>(wind km/h)</small></div>"
        "<div class='hb-legend'>🟡 too little · 🟢 good · 🔴 too strong &nbsp;·&nbsp; "
        "⤴ gusts = frosted bar</div>"
        f"<div class='hb-row'>{''.join(bars)}</div>"
        "</div>",
        unsafe_allow_html=True,
    )


@st.fragment(run_every=30)
def _tv_bottom_info(cfg):
    """Unterer Bereich: wechselt synchron zum Ranking (2s/30s) zwischen Spot-Info
    und 3-Tage-Vorhersage. Gleiche Zeitformel wie der Metrik-Umschalter -> synchron.
    Gibt es nur eins von beiden, wird dieses dauerhaft gezeigt."""
    info = load_spot_info(cfg["spot"]) or {}
    has_info = bool(
        (info.get("description") or "").strip() or info.get("webcam_url")
        or info.get("image") or load_spot_image_ids(cfg["spot"])
    )
    coords = _spot_coords(cfg["spot"])
    has_fc = coords is not None and _fetch_forecast_3d(coords[0], coords[1]) is not None

    if has_info and has_fc:
        # Phase 0 (= Ranking "2 s") -> Spot-Info, Phase 1 (= "30 s") -> Vorhersage.
        phase = (int(datetime.now().timestamp()) // 30) % 2
        if phase:
            _tv_forecast(cfg, coords)
        else:
            _tv_spot_info(cfg)
    elif has_info:
        _tv_spot_info(cfg)
    elif has_fc:
        _tv_forecast(cfg, coords)


def render_spots_page():
    """Reine Spot-Seite (Revierführer): Filter Land/Spot -> Beschreibung, Webcam/
    Bild und Wetter (aktuell + 3-Tage) des gewählten Spots."""
    st.markdown("## 🗺️ Spots")
    all_info = load_all_spot_info()
    if not all_info:
        st.info(
            "No spots with a description yet. They appear here automatically once "
            "the AI enrichment has run (or when you save a text in the backoffice)."
        )
        return

    countries = sorted({
        (s.get("country") or "").strip() for s in all_info if (s.get("country") or "").strip()
    })
    fcol1, fcol2 = st.columns(2)
    country = fcol1.selectbox("Country", ["All countries"] + countries, key="spots_country")
    pool = [
        s for s in all_info
        if country == "All countries" or (s.get("country") or "").strip() == country
    ]
    names = [s["spot"] for s in pool]
    if not names:
        st.info("No spots for this country.")
        return
    # Spot in der URL halten (bookmarkbar / direkt verlinkbar).
    url_spot = st.query_params.get("spot")
    default_idx = names.index(url_spot) if url_spot in names else 0
    spot = fcol2.selectbox("Spot", names, index=default_idx)
    if spot != st.query_params.get("spot"):
        st.query_params["spot"] = spot
        st.session_state.pop("spots_fcday", None)   # anderer Spot -> Stundenansicht zu
        st.rerun()

    chosen = next((s for s in pool if s["spot"] == spot), {})
    info = load_spot_info(spot) or {}   # voll, inkl. Bild/Webcam

    heading = f"### {spot}"
    if chosen.get("country"):
        heading += f"  ·  📍 {chosen['country']}"
    st.markdown(heading)

    desc = (chosen.get("description") or info.get("description") or "").strip()
    if desc:
        st.markdown(
            f"<div style='font-size:18px;line-height:1.6;'>{desc}</div>",
            unsafe_allow_html=True,
        )

    webcam = (info.get("webcam_url") or "").strip()
    if webcam and _is_image_url(webcam):
        components.html(
            f"<img src='{webcam}' style='width:100%;max-height:380px;object-fit:cover;"
            "border-radius:16px;display:block;'>", height=388)
    elif webcam:
        components.html(
            f"<iframe src='{webcam}' allow='autoplay; fullscreen' "
            "style='width:100%;height:380px;border:0;border-radius:16px;'></iframe>", height=388)

    # Bilder-Galerie als gleichmaessiges Kachel-Raster (alle gleich gross, cover).
    # Gecachte Thumbnails statt Voll-Bild-base64 -> viel kleinere Seitenlast.
    img_ids = load_spot_image_ids(spot)
    if img_ids:
        tiles = []
        for iid in img_ids:
            uri = _spot_thumb_uri(iid)
            if uri:
                tiles.append(
                    f"<div class='sp-tile' style=\"background-image:url('{uri}')\"></div>"
                )
        if tiles:
            st.markdown(
                "<style>"
                # 5 Kacheln nebeneinander (breit); auf schmalen Screens automatisch weniger.
                ".sp-gallery{display:grid;gap:14px;margin-top:10px;"
                "grid-template-columns:repeat(5,1fr);}"
                ".sp-tile{height:240px;background-size:cover;background-position:center;"
                "border-radius:16px;box-shadow:0 6px 18px rgba(0,0,0,.18);}"
                "@media (max-width:1200px){.sp-gallery{grid-template-columns:repeat(3,1fr);}}"
                "@media (max-width:680px){.sp-gallery{grid-template-columns:repeat(2,1fr);}}"
                "</style>"
                f"<div class='sp-gallery'>{''.join(tiles)}</div>",
                unsafe_allow_html=True,
            )
    elif not webcam:
        img_uri = _bytes_to_data_uri(info.get("image"), info.get("image_mime"))
        if img_uri:
            st.markdown(
                f"<img src='{img_uri}' style='width:100%;border-radius:16px;'>",
                unsafe_allow_html=True,
            )

    # resolve_spot_coords liest aus der DB ODER geocodet einmalig und speichert ->
    # auch Spots ohne GPS-Session bekommen Wetter (und danach profitiert das TV davon).
    lat, lon = resolve_spot_coords(spot)
    if lat is not None and lon is not None:
        coords = (lat, lon)
        st.markdown("#### 🌬️ Weather")
        components.html(_tv_weather_html(coords[0], coords[1]), height=130)
        # 3-Tage-Kacheln mit Button je Tag -> Stundenansicht per Rerun (kein Reload).
        render_spots_forecast(spot, coords)
    else:
        st.caption("No coordinates could be determined for this spot – no weather.")


def _product_cards_html(spot):
    """HTML-Karten (Bild + Titel + Preis, je Shop/Cafe-Link) der aktiven Produkte."""
    cards = []
    for p in load_spot_products(spot, only_active=True):
        img = _bytes_to_data_uri(p.get("image"), p.get("image_mime"))
        img_html = (
            f"<div class='tv-prod-img' style=\"background-image:url('{img}')\"></div>"
            if img else "<div class='tv-prod-img tv-prod-noimg'>🛍️</div>"
        )
        price = f"<div class='tv-prod-price'>{p['price']}</div>" if p.get("price") else ""
        inner = f"{img_html}<div class='tv-prod-title'>{p['title']}</div>{price}"
        if p.get("url"):
            cards.append(
                f"<a class='tv-prod-card' href='{p['url']}' target='_blank' "
                f"rel='noopener'>{inner}</a>"
            )
        else:
            cards.append(f"<div class='tv-prod-card'>{inner}</div>")
    return cards


def _parse_track(raw):
    """JSON-Track ([[lat,lon],...]) -> Liste von (lat, lon)-Tupeln oder None."""
    if raw is None:
        return None
    if isinstance(raw, float) and pd.isna(raw):
        return None
    data = raw
    if isinstance(raw, str):
        try:
            data = json.loads(raw)
        except (ValueError, TypeError):
            return None
    pts = []
    if isinstance(data, list):
        for p in data:
            if isinstance(p, (list, tuple)) and len(p) >= 2:
                try:
                    pts.append((float(p[0]), float(p[1])))
                except (TypeError, ValueError):
                    continue
    return pts or None


def render_history_overview(record):
    """Session-Übersicht aus den gespeicherten Ranking-Werten (ohne Roh-FIT)."""

    def num(key):
        value = record.get(key)
        return None if value is None or pd.isna(value) else float(value)

    st.markdown("## 🌊 Session overview")

    meta_bits = []
    date_value = record.get("date")

    if date_value is not None and pd.notna(date_value):
        _dt = pd.to_datetime(date_value, errors="coerce")
        if pd.notna(_dt):
            # Uhrzeit nur zeigen, wenn vorhanden (Uhr-Sessions); Altdaten nur Datum.
            if _dt.hour or _dt.minute or _dt.second:
                meta_bits.append(f"📅 {_dt:%Y-%m-%d %H:%M}")
            else:
                meta_bits.append(f"📅 {_dt:%Y-%m-%d}")

    for icon, key in (("📍", "surfspot"), ("🏄", "board"), ("⛵", "sail")):
        value = record.get(key)

        if value is not None and pd.notna(value) and str(value).strip():
            meta_bits.append(f"{icon} {value}")

    if meta_bits:
        st.caption(" · ".join(meta_bits))

    best_1s = num("speed_1s_kmh")
    best_30s = num("speed_30s_kmh")
    distance_km = num("total_distance_km")
    longest_run_km = num("longest_run_km")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Top 2 s", "–" if best_1s is None else f"{best_1s:.2f} km/h")
    c2.metric("Top 30 s", "–" if best_30s is None else f"{best_30s:.2f} km/h")
    c3.metric("Total distance", "–" if distance_km is None else f"{distance_km:.2f} km")
    c4.metric("Longest run", "–" if longest_run_km is None else f"{longest_run_km:.2f} km")

    st.markdown("## 🌦️ Weather during session")

    wind = num("wind_kmh")
    gust = num("gust_kmh")
    temp = num("temp_c")
    precip = num("precip_mm")
    wdir = num("wind_dir_deg")
    code = record.get("weather_code")
    has_code = code is not None and pd.notna(code)

    if all(v is None for v in (wind, gust, temp, precip)) and not has_code:
        st.info("No weather data saved for this session.")
    else:
        emoji, description = (
            WEATHER_CODES.get(int(code), ("❓", "Unknown")) if has_code else ("❓", "Unknown")
        )

        w1, w2, w3, w4 = st.columns(4)

        w1.metric(
            "Wind",
            "–" if wind is None else f"{wind:.0f} km/h",
            None if gust is None else f"Gusts {gust:.0f} km/h",
        )
        w2.metric("Temperature", "–" if temp is None else f"{temp:.1f} °C")
        w3.metric("Precipitation", "–" if precip is None else f"{precip:.1f} mm")
        w4.metric("Condition", emoji, description, delta_color="off")

        if wdir is not None:
            st.caption(
                f"🧭 Wind from **{degrees_to_compass(wdir)}** "
                f"{wind_arrow(wdir)} ({wdir:.0f}°)"
            )

    st.markdown("## ⚡ Speed values")

    speed_table = pd.DataFrame([
        {
            "Category": "2 seconds",
            "Speed km/h": best_1s,
            "Speed kn": None if best_1s is None else best_1s / 1.852,
        },
        {
            "Category": "30 seconds",
            "Speed km/h": best_30s,
            "Speed kn": None if best_30s is None else best_30s / 1.852,
        },
    ])

    st.dataframe(
        speed_table.round(2),
        width="stretch",
        hide_index=True,
        height=df_height(len(speed_table)),
    )

    # --- Sprünge / SUP (nur wenn von der Uhr geliefert) ---
    jumps = record.get("jumps")
    max_air = num("max_airtime_s")
    max_jump = num("max_jump_m")
    strokes = record.get("strokes")
    cadence = record.get("cadence_spm")
    max_cadence = record.get("max_cadence_spm")

    def _has(value):
        return value is not None and pd.notna(value) and value

    if _has(jumps) or max_air or max_jump:
        st.markdown("## 🪂 Jumps")
        j1, j2, j3 = st.columns(3)
        j1.metric("Jumps", "–" if not _has(jumps) else f"{int(jumps)}")
        j2.metric("Max airtime", "–" if max_air is None else f"{max_air:.1f} s")
        j3.metric("Highest jump", "–" if max_jump is None else f"{max_jump:.1f} m")

    if _has(strokes) or _has(cadence) or _has(max_cadence):
        st.markdown("## 🛶 Paddling")
        p1, p2, p3 = st.columns(3)
        p1.metric("Strokes", "–" if not _has(strokes) else f"{int(strokes)}")
        p2.metric("Max cadence", "–" if not _has(max_cadence) else f"{int(max_cadence)} spm")
        p3.metric("Cadence (end)", "–" if not _has(cadence) else f"{int(cadence)} spm")

    # Karte aus der von der Uhr gesendeten Route (per ID nachgeladen, falls da).
    track_pts = _parse_track(load_session_track(record.get("id")))
    if track_pts:
        st.markdown("## 🗺️ Track")
        show_map(pd.DataFrame(track_pts, columns=["lat", "lon"]))

    st.caption(
        "ℹ️ Individual runs and max/avg speed are only available right after the "
        "upload – for saved sessions the stored metrics (and the track, if the "
        "watch sent one) are shown."
    )


# =====================================================================
#  Admin-Backoffice (Werbung pro Spot + Profil-Verwaltung)
#  Aufruf per ?admin=1, geschuetzt durch ADMIN_PASSWORD aus den Secrets.
# =====================================================================

def _admin_password():
    try:
        pw = st.secrets.get("ADMIN_PASSWORD")
    except Exception:
        pw = None
    return pw or os.environ.get("ADMIN_PASSWORD")


def _admin_flash(msg):
    st.session_state["_admin_flash"] = msg
    st.rerun()


def render_admin():
    st.markdown("# 🔧 Backoffice")

    admin_pw = _admin_password()
    if not admin_pw:
        st.error(
            "Kein **ADMIN_PASSWORD** gesetzt. Lege es in den Streamlit-Secrets an "
            "(oder als Umgebungsvariable), dann ist das Backoffice nutzbar."
        )
        st.code('ADMIN_PASSWORD = "dein-geheimes-passwort"', language="toml")
        if st.button("← Zurück zur App"):
            st.query_params.clear()
            st.rerun()
        return

    if not st.session_state.get("is_admin"):
        st.caption("Bitte mit dem Admin-Passwort anmelden.")
        with st.form("admin_login"):
            pw = st.text_input("Admin-Passwort", type="password")
            ok = st.form_submit_button("Anmelden")
        if ok:
            if pw == admin_pw:
                st.session_state["is_admin"] = True
                st.rerun()
            else:
                st.error("Falsches Passwort.")
        if st.button("← Zurück zur App", key="admin_back_login"):
            st.query_params.clear()
            st.rerun()
        return

    flash = st.session_state.pop("_admin_flash", None)
    if flash:
        st.success(flash)

    top = st.columns([1.4, 1.4, 5])
    if top[0].button("← Zurück zur App", key="admin_back"):
        st.query_params.clear()
        st.rerun()
    if top[1].button("Admin abmelden", key="admin_logout"):
        st.session_state["is_admin"] = False
        st.rerun()

    tab_ads, tab_profiles = st.tabs(["📣 Werbung pro Spot", "👤 Profile"])
    with tab_ads:
        render_admin_ads()
    with tab_profiles:
        render_admin_profiles()


def render_admin_ads():
    st.caption(
        "Sponsor-Logo/Name und Produkte (verlinkt zu Shop oder Café) verwalten, "
        "die auf dem Spot-TV erscheinen."
    )
    spots = _all_spot_names()

    col_a, col_b = st.columns(2)
    choice = col_a.selectbox("Spot wählen", ["– auswählen –"] + spots, key="admin_ad_spot")
    new_spot = col_b.text_input("…oder neuer Spot-Name", key="admin_ad_newspot").strip()
    spot = new_spot or (choice if choice != "– auswählen –" else "")

    if not spot:
        st.info("Spot wählen oder eingeben, um die Werbung zu verwalten.")
        return

    ad = load_spot_ad(spot) or {}

    st.markdown(f"### Sponsor für **{spot}**")
    if ad.get("logo"):
        st.image(bytes(ad["logo"]), width=200, caption="aktuelles Logo")
    with st.form(f"ad_form_{spot}"):
        name = st.text_input("Sponsor-Name", value=ad.get("sponsor_name") or "")
        url = st.text_input("Link (Shop/Café-Webseite)", value=ad.get("sponsor_url") or "")
        active = st.checkbox("Aktiv (auf dem TV anzeigen)", value=bool(ad.get("active", True)))
        logo_up = st.file_uploader("Logo (PNG/JPG/WebP)", type=["png", "jpg", "jpeg", "webp"])
        clear_logo = st.checkbox("Aktuelles Logo entfernen") if ad.get("logo") else False
        saved = st.form_submit_button("💾 Sponsor speichern")
    if saved:
        save_spot_ad(
            spot, name, url, active,
            logo_bytes=logo_up.getvalue() if logo_up else None,
            logo_mime=logo_up.type if logo_up else None,
            clear_logo=clear_logo,
        )
        _admin_flash(f"Sponsor für {spot} gespeichert.")
    if ad and st.button("🗑️ Sponsor-Eintrag löschen", key=f"del_ad_{spot}"):
        delete_spot_ad(spot)
        _admin_flash(f"Sponsor für {spot} gelöscht.")

    st.markdown("### Produkte / Angebote")
    for prod in load_spot_products(spot):
        flag = "✅" if prod.get("active") else "⛔"
        price = f" · {prod['price']}" if prod.get("price") else ""
        with st.expander(f"{flag} {prod['title']}{price}"):
            _product_editor(spot, prod)
    st.markdown("#### ➕ Neues Produkt")
    _product_editor(spot, None)

    # --- Spot-Info (unten auf dem TV: Text links, Webcam/Bild rechts) ---
    info_head = st.columns([3, 1])
    info_head[0].markdown("### Spot-Info (unten auf dem TV)")
    if info_head[1].button("🔄 Aktualisieren", key=f"inforeload_{spot}",
                           help="Lädt frisch aus der DB (z.B. nach dem KI-Anreichern)."):
        _clear_ad_caches()
        st.rerun()
    info = load_spot_info(spot) or {}
    if info.get("auto_filled"):
        st.warning(
            "🤖 Dieser Text wurde automatisch per KI erstellt – bitte unten prüfen "
            "und mit dem Speichern-Button bestätigen (danach gilt er als geprüft)."
        )
    info_prefill_key = f"spotinfoprefill_{spot}"
    info_prefill = st.session_state.get(info_prefill_key)

    src_cols = st.columns([4, 1.5])
    src_url = src_cols[0].text_input(
        "Info-Quelle (URL, optional) – Text automatisch holen",
        key=f"infosrc_{spot}",
        placeholder="https://… (z.B. Wikipedia/Spot-Guide)",
    )
    if src_cols[1].button("🔗 Text holen", key=f"infofetch_{spot}", use_container_width=True):
        with st.spinner("Lade Text…"):
            res = fetch_page_description(src_url)
        if res.get("error"):
            st.error(res["error"])
        elif res.get("description"):
            st.session_state[info_prefill_key] = res["description"]
            st.rerun()
        else:
            st.warning("Keine Beschreibung gefunden – bitte selbst eintragen.")

    _cur_coords = _spot_coords(spot)
    with st.form(f"info_form_{spot}"):
        desc = st.text_area(
            "Beschreibung",
            value=(info_prefill if info_prefill is not None else (info.get("description") or "")),
            height=140,
        )
        country = st.text_input(
            "Land (für den Filter der Spots-Seite)", value=info.get("country") or "",
            help="Wird vom KI-Job automatisch gesetzt; hier überschreibbar.",
        )
        best_winds = st.text_input(
            "Beste Windrichtungen (für die Lohnt-sich-Ampel)",
            value=info.get("best_winds") or "",
            help="Kompass-Kürzel, z.B. 'SW, W, NW'. Wird vom KI-Job gesetzt; hier "
                 "überschreibbar. Leer = keine Richtungsbewertung.",
        )
        webcam = st.text_input(
            "Webcam- oder Bild-URL (optional, hat Vorrang vor Upload)",
            value=info.get("webcam_url") or "",
            help="Direktes Bild (…/snapshot.jpg) lädt sich auto. neu; eine "
                 "einbettbare Seite (YouTube-Live/Windy) wird als iFrame gezeigt.",
        )
        st.caption("📍 Koordinaten (für Wetter). Leer lassen = automatisches "
                   "Geocoding; hier setzen, um einen falschen Ort zu korrigieren.")
        gc1, gc2 = st.columns(2)
        lat_in = gc1.number_input(
            "Latitude (z.B. 52.30)", value=(_cur_coords[0] if _cur_coords else None),
            format="%.5f", step=0.001, key=f"lat_{spot}",
        )
        lon_in = gc2.number_input(
            "Longitude (z.B. 5.62)", value=(_cur_coords[1] if _cur_coords else None),
            format="%.5f", step=0.001, key=f"lon_{spot}",
        )
        if st.form_submit_button("💾 Spot-Info speichern"):
            save_spot_info(spot, desc, webcam, country=country, best_winds=best_winds)
            if lat_in is not None and lon_in is not None:
                update_spot_coords(spot, lat_in, lon_in)
            st.session_state.pop(info_prefill_key, None)
            _admin_flash("Spot-Info gespeichert.")

    # --- Bilder-Galerie (mehrere Bilder je Spot, ausserhalb des Formulars) ---
    st.markdown("#### 🖼️ Bilder (Galerie)")
    gallery = load_spot_images(spot)
    if info.get("image") and not gallery:
        st.caption("Es gibt noch ein altes Einzelbild – lade hier Galerie-Bilder "
                   "hoch, sie ersetzen es in der Anzeige.")
        st.image(bytes(info["image"]), width=150, caption="altes Einzelbild")
    if gallery:
        gcols = st.columns(4)
        for idx, gi in enumerate(gallery):
            with gcols[idx % 4]:
                if gi.get("image"):
                    st.image(bytes(gi["image"]), use_container_width=True)
                if st.button("🗑️", key=f"delimg_{gi['id']}", help="Dieses Bild löschen"):
                    delete_spot_image(gi["id"])
                    _admin_flash("Bild gelöscht.")
    new_imgs = st.file_uploader(
        "Bilder hinzufügen (mehrere möglich)", type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True, key=f"galup_{spot}",
    )
    if st.button("➕ Bilder hochladen", key=f"galadd_{spot}", disabled=not new_imgs):
        for up in new_imgs:
            add_spot_image(spot, up.getvalue(), up.type)
        _admin_flash(f"{len(new_imgs)} Bild(er) hinzugefügt.")


def _product_editor(spot, prod):
    pid = prod["id"] if prod else None
    suffix = pid if pid else "new"
    prefill_key = f"prodprefill_{spot}_{suffix}"
    prefill = st.session_state.get(prefill_key, {})

    # --- Auto-Befuellung aus einer Produkt-URL (ausserhalb des Formulars) ---
    auto_cols = st.columns([4, 1.5])
    auto_url = auto_cols[0].text_input(
        "Produkt-URL für Auto-Befüllung",
        value=prefill.get("url") or (prod or {}).get("url") or "",
        key=f"autourl_{spot}_{suffix}",
        placeholder="https://shop.example/produkt/…",
    )
    if auto_cols[1].button("🔗 Aus URL ziehen", key=f"fetch_{spot}_{suffix}",
                           use_container_width=True):
        with st.spinner("Lade Produktdaten…"):
            meta = fetch_product_meta(auto_url)
        if meta.get("error"):
            st.error(meta["error"])
        else:
            img_bytes, img_mime = (None, None)
            if meta.get("image_url"):
                img_bytes, img_mime = download_image_bytes(meta["image_url"])
            st.session_state[prefill_key] = {
                "url": auto_url,
                "title": meta.get("title") or "",
                "price": meta.get("price") or "",
                "image_bytes": img_bytes,
                "image_mime": img_mime,
            }
            note = "Daten übernommen."
            if not meta.get("price"):
                note += " Preis nicht gefunden – bitte eintragen."
            if not img_bytes:
                note += " Bild nicht gefunden – bitte hochladen."
            st.success(note)
            st.rerun()

    fetched_img = prefill.get("image_bytes")
    if fetched_img:
        st.image(bytes(fetched_img), width=160, caption="aus URL geladen")
    elif prod and prod.get("image"):
        st.image(bytes(prod["image"]), width=160)

    with st.form(f"prod_form_{spot}_{suffix}"):
        title = st.text_input(
            "Titel", value=prefill.get("title") or (prod or {}).get("title") or ""
        )
        c1, c2 = st.columns(2)
        price = c1.text_input(
            "Preis (z.B. €499)", value=prefill.get("price") or (prod or {}).get("price") or ""
        )
        sort_order = c2.number_input(
            "Reihenfolge", value=int((prod or {}).get("sort_order") or 0), step=1
        )
        url = st.text_input(
            "Link (Shop/Café)", value=prefill.get("url") or (prod or {}).get("url") or ""
        )
        active = st.checkbox("Aktiv", value=bool((prod or {}).get("active", True)))
        img_up = st.file_uploader(
            "Bild überschreiben (optional)", type=["png", "jpg", "jpeg", "webp"],
            key=f"img_{spot}_{suffix}",
        )
        clear_img = (
            st.checkbox("Aktuelles Bild entfernen", key=f"clr_{spot}_{suffix}")
            if (prod and prod.get("image")) else False
        )
        save = st.form_submit_button("💾 Produkt speichern")

    if save:
        if not title.strip():
            st.error("Titel ist erforderlich.")
        else:
            if img_up:
                img_bytes, img_mime = img_up.getvalue(), img_up.type
            elif fetched_img:
                img_bytes, img_mime = fetched_img, prefill.get("image_mime")
            else:
                img_bytes, img_mime = None, None
            save_spot_product(
                pid, spot, title, price, url, active, sort_order,
                image_bytes=img_bytes, image_mime=img_mime, clear_image=clear_img,
            )
            st.session_state.pop(prefill_key, None)
            _admin_flash("Produkt gespeichert.")
    if prod and st.button("🗑️ Produkt löschen", key=f"delprod_{pid}"):
        delete_spot_product(pid)
        _admin_flash("Produkt gelöscht.")


def render_admin_profiles():
    users = list_all_users()
    profiles = load_profiles()
    names = sorted({u["username"] for u in users} | set(profiles.keys()))
    if not names:
        st.info("Noch keine Profile vorhanden.")
        return

    name = st.selectbox("Profil / Fahrer", names, key="admin_prof_sel")
    user = next((u for u in users if u["username"] == name), None)
    rider = profiles.get(name, {})

    if user:
        st.caption(f"Konto-ID {user['id']} · {user.get('email') or 'keine E-Mail'}")
    else:
        st.caption("Profil ohne verknüpftes Konto (nur aus Sessions).")

    # ---- Equipment ----
    st.markdown("### Equipment")
    edit_keys = [("Spots (geteilt)", "spots")]
    for sp in SPORTS:
        meta = SPORT_META[sp]
        edit_keys.append((f"{meta['label']} – Boards", meta["boards_key"]))
        edit_keys.append((f"{meta['label']} – {meta['gear_key']}", meta["gear_key"]))

    with st.form(f"equip_{name}"):
        st.caption("Ein Eintrag pro Zeile. Leere Zeilen werden ignoriert.")
        new_vals = {}
        for label, key in edit_keys:
            cur = rider.get(key, []) or []
            txt = st.text_area(label, value="\n".join(cur), height=90, key=f"eq_{name}_{key}")
            new_vals[key] = [ln.strip() for ln in txt.splitlines() if ln.strip()]
        if st.form_submit_button("💾 Equipment speichern"):
            for key, vals in new_vals.items():
                set_profile_list(name, key, vals)
            _admin_flash("Equipment gespeichert.")

    # ---- Konto-Aktionen ----
    st.markdown("### Konto")

    if user:
        st.markdown("**Geräte-Token (Uhr-Upload)**")
        cur_tok = get_device_token(user["id"]) or ""
        with st.form(f"tok_{name}"):
            tok_in = st.text_input(
                "Token zuweisen (z.B. der in der Uhr fest hinterlegte)", value=cur_tok
            )
            tc1, tc2 = st.columns(2)
            set_tok = tc1.form_submit_button("Token zuweisen")
            gen_tok = tc2.form_submit_button("Neuen Token erzeugen")
        if set_tok:
            ok, msg = admin_set_device_token(user["id"], tok_in)
            if ok:
                _admin_flash(msg)
            else:
                st.error(msg)
        if gen_tok:
            regenerate_device_token(user["id"])
            _admin_flash("Neuer Geräte-Token erzeugt.")

        st.markdown("**Passwort zurücksetzen**")
        with st.form(f"pwd_{name}"):
            new_pw = st.text_input("Neues Passwort (min. 6 Zeichen)", type="password")
            if st.form_submit_button("Passwort setzen"):
                ok, msg = admin_set_password(user["id"], new_pw)
                if ok:
                    _admin_flash(msg)
                else:
                    st.error(msg)

    st.markdown("**Umbenennen**")
    with st.form(f"rename_{name}"):
        new_name = st.text_input("Neuer Name")
        if st.form_submit_button("Umbenennen"):
            ok, msg = admin_rename_profile(name, new_name)
            if ok:
                _admin_flash(msg)
            else:
                st.error(msg)

    others = [n for n in names if n != name]
    if others:
        st.markdown("**Zusammenführen** (verschiebt Sessions+Equipment, löscht dann die Quelle)")
        with st.form(f"merge_{name}"):
            target = st.selectbox("Zusammenführen in", others)
            if st.form_submit_button(f"'{name}' → Ziel zusammenführen"):
                _moved, msg = admin_merge_profiles(name, target)
                _admin_flash(msg)

    st.markdown("**Gefahrenzone**")
    confirm = st.checkbox(f"Ja, '{name}' und alle zugehörigen Daten löschen", key=f"cfm_{name}")
    if st.button("🗑️ Profil löschen", disabled=not confirm, key=f"delprof_{name}"):
        admin_delete_profile(name)
        _admin_flash(f"Profil '{name}' gelöscht.")


load_css(app_path("assets", "style.css"))

# Spot-TV-Vollbildmodus (Cafe/Shop/Club-Screen): wird per ?tv=... aufgerufen und
# rendert eine eigene, grossflaechige Ansicht statt der normalen Seite.
_tv_cfg = _spot_tv_config()
if _tv_cfg is not None:
    ensure_schema()
    ensure_watch_columns()
    render_spot_tv(_tv_cfg)
    st.stop()

# Admin-Backoffice (Werbung + Profile) – per ?admin=1, eigenes Login. Vor dem
# normalen App-Login, damit man unabhaengig vom Nutzerkonto hineinkommt.
if "admin" in st.query_params:
    ensure_schema()
    ensure_watch_columns()
    render_admin()
    st.stop()

logo_img = image_to_base64(app_path("assets", "windsurfer.png"))

# Aktiver Sport (aus ?sport=). Standard: Windsurf.
sport = active_sport()
_is_spots_view = st.query_params.get("view") == "spots"

# Header-Umschalter Sportarten + ganz rechts die reine Spots-Seite. Klick setzt
# ?sport= bzw. ?view=spots in der URL (bleibt über Reload/Link erhalten).
_sw_cols = st.columns([2] * len(SPORTS) + [2])
for _i, _key in enumerate(SPORTS):
    if _sw_cols[_i].button(
        SPORT_META[_key]["label"],
        key=f"switch_sport_{_key}",
        use_container_width=True,
        type="primary" if (_key == sport and not _is_spots_view) else "secondary",
    ):
        if "view" in st.query_params:
            del st.query_params["view"]
        st.query_params["sport"] = _key
        st.rerun()
if _sw_cols[len(SPORTS)].button(
    "🗺️ Spots", key="switch_view_spots", use_container_width=True,
    type="primary" if _is_spots_view else "secondary",
):
    if not _is_spots_view:
        st.query_params["view"] = "spots"
        st.rerun()

# Vollflächiges Hintergrundbild je Sport. Lege dein Wunschfoto als
# assets/background.webp (Windsurf) bzw. assets/background_kite.webp (Kite) ab
# (auch .jpg/.jpeg/.png). MIME-Typ wird automatisch passend gesetzt.
bg_uri = background_data_uri(sport)

# Hintergrund nur EINMAL pro Session/Sport setzen. Frueher wurde das ~200 KB
# grosse base64-Bild bei JEDEM Rerun via st.markdown erneut uebertragen -> traege.
# Jetzt schreiben wir es per JS einmalig in einen <style> im Eltern-Dokument;
# bei normalen Reruns wird nichts erneut gesendet (Bremse weg).
if bg_uri and st.session_state.get("_bg_sport") != sport:
    st.session_state["_bg_sport"] = sport
    components.html(
        """
        <script>
        (function () {
          try {
            var d = window.parent.document;
            var el = d.getElementById("ws-bg");
            if (!el) { el = d.createElement("style"); el.id = "ws-bg"; d.head.appendChild(el); }
            el.textContent = '.stApp { background-color:#02162b;'
              + ' background-image: linear-gradient(rgba(2,22,43,.45), rgba(2,22,43,.62)),'
              + ' url("__BG__");'
              + ' background-position:center center; background-size:cover;'
              + ' background-repeat:no-repeat; background-attachment:fixed; }';
          } catch (e) {}
        })();
        </script>
        """.replace("__BG__", bg_uri),
        height=0,
    )

if sport == "windsurf" and logo_img:
    logo_icon = (
        f'<img src="data:image/png;base64,{logo_img}" '
        'style="height:1em;vertical-align:-0.15em;margin-right:.15em;" alt="">'
    )
else:
    logo_icon = SPORT_META[sport]["emoji"]

st.markdown(f"""
<div class="hero">
    <div class="hero-content">
        <div class="logo">MyWaterSessions<span class="logo-dot">.</span>{BETA_BADGE}</div>
        <div class="logo-rule"></div>
        <div class="title">{SPORT_META[sport]["title"]}</div>
        <p>The home for everyone active on the water</p>
    </div>
</div>
""", unsafe_allow_html=True)


# Rechtsseiten (Impressum/Datenschutz) zuerst behandeln – ohne Login erreichbar.
if render_legal_page():
    st.stop()


# =====================================================================
#  Login / Registrierung (Gate vor dem Rest der App)
# =====================================================================

def render_login():
    st.markdown(f"## 🔐 Sign in {BETA_BADGE}", unsafe_allow_html=True)
    st.info(
        "Please log in or register. Afterwards you can join groups or "
        "create your own."
    )

    st.markdown(
        '<div style="text-align:center;margin:.5rem 0 1rem;opacity:.85;">'
        '<a href="?seite=impressum" target="_self" style="color:#2bd4d9;">Impressum</a>'
        ' &nbsp;·&nbsp; '
        '<a href="?seite=datenschutz" target="_self" style="color:#2bd4d9;">Datenschutzerklärung</a>'
        '</div>',
        unsafe_allow_html=True,
    )

    tab_login, tab_register = st.tabs(["Log in", "Register"])

    with tab_login:
        with st.form("login_form"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            remember = st.checkbox("Stay logged in", value=True)
            submitted = st.form_submit_button("Log in")

        if submitted:
            if verify_login(username, password):
                user = get_user(username.strip())
                login_session(user, remember)
                st.rerun()
            else:
                st.error("Wrong username or password.")

    with tab_register:
        with st.form("register_form"):
            new_username = st.text_input("Choose a username")
            new_email = st.text_input("Email")
            pwd1 = st.text_input("Password (at least 6 characters)", type="password")
            pwd2 = st.text_input("Repeat password", type="password")

            st.markdown("**Consents**")

            consent_save = st.checkbox(
                "I confirm that my uploaded sessions are stored in the ranking."
            )
            consent_visible = st.checkbox(
                "I agree that spot, date, speed and equipment are visible in "
                "the ranking."
            )
            consent_privacy = st.checkbox(
                "I have read the Datenschutzerklärung and accept it."
            )

            st.caption(
                "ℹ️ Your GPS track is not stored and never shown publicly – only "
                "the metrics appear in the ranking. Details in the "
                "Datenschutzerklärung (link above)."
            )

            submitted = st.form_submit_button("Register")

        if submitted:
            if not (consent_save and consent_visible and consent_privacy):
                st.error(
                    "Please confirm all three consents to register."
                )
            elif pwd1 != pwd2:
                st.error("The passwords do not match.")
            else:
                ok, message = register_user(new_username, pwd1, new_email)

                if ok:
                    st.success(message)
                else:
                    st.error(message)


def render_achievements(ach):
    """Hebt die Rekorde der gerade hochgeladenen Session hervor (für den Fahrer)."""
    personal = ach.get("personal", [])
    spot = ach.get("spot", [])
    year = ach.get("year", [])
    group_events = ach.get("group_events", [])

    if not (personal or spot or year or group_events):
        return

    st.balloons()
    st.markdown("### 🏆 Your achievements this session")

    def fmt(r):
        return f"{r['value']:.{r['decimals']}f} {r['unit']}"

    for r in spot:
        extra = "" if r["previous"] is None else f" – previous spot record: {r['previous']:.{r['decimals']}f}"
        st.success(f"🏆 **New spot record** at {r['spot']}: {r['label']} {fmt(r)}{extra}")

    for r in personal:
        if r["previous"] is None:
            st.success(f"🏆 **First personal best**: {r['label']} {fmt(r)}")
        else:
            st.success(
                f"🏆 **Personal record**: {r['label']} {fmt(r)} "
                f"(previously {r['previous']:.{r['decimals']}f} {r['unit']}) 📈"
            )

    for r in year:
        st.info(f"📅 **Best of {r['year']}**: {r['label']} {fmt(r)}")

    for e in [g for g in group_events if g.get("is_record")]:
        st.success(f"🏆 **Group record** in {e['group_name']}: {e['label']} {fmt(e)}")

    for e in [g for g in group_events if not g.get("is_record")]:
        st.info(f"🏅 Rank {e['rank']} in group {e['group_name']} for {e['label']} ({fmt(e)})")


def render_group_news_banner(user):
    """Banner mit ungelesenen Gruppen-Ereignissen (Rekorde/Top-3 anderer)."""
    events = unseen_group_events(user["id"], user["username"])

    if not events:
        return

    st.markdown("### 📣 News from your groups")

    for e in events:
        if e.get("rank") == 1:
            st.success(e["message"])
        else:
            st.info(e["message"])

    if st.button("✓ Mark all as read", key="mark_events_seen"):
        mark_group_events_seen(user["id"])
        st.rerun()

    st.markdown("---")


def _session_is_complete(spot, board, sail):
    def ok(v):
        v = str(v or "").strip().lower()
        return v != "" and v not in ("none", "nan", "null")
    return ok(spot) and ok(board) and ok(sail)


def render_session_editor(user):
    """Eigene Sessions nachpflegen: Spot/Board/Segel ergänzen, damit sie ins
    Ranking aufgenommen werden (Pflicht: Spot + Board + Segel/Kite). Zeigt auch
    unvollständige, von der Uhr hochgeladene Sessions an."""
    if not user:
        return

    sport = active_sport()
    gear_meta = SPORT_META[sport]
    gear_word = gear_meta["gear_label"]
    df = load_rider_sessions(user["username"], sport)

    with st.expander("✏️ Edit / complete my sessions", expanded=False):
        st.caption(
            "Sessions only appear in the ranking once spot, board and "
            f"{gear_word.lower()} are filled in. Watch uploads start incomplete (⚠️)."
        )

        if st.button("🔄 Refresh (show new uploads)", key="es_refresh",
                     use_container_width=True):
            clear_data_caches()
            st.rerun()

        if df is None or df.empty:
            st.info("No sessions yet for this sport.")
            return

        if "date" in df.columns:
            df = df.sort_values("date", ascending=False)

        # Auswahlliste (neueste zuerst); Status-Markierung vollständig/unvollständig.
        labels = []
        row_by_label = {}
        for _, row in df.iterrows():
            sid = int(row["id"])
            complete = _session_is_complete(
                row.get("surfspot"), row.get("board"), row.get("sail")
            )
            mark = "✅" if complete else "⚠️"
            spot = str(row.get("surfspot") or "").strip() or "no spot"
            label = f"{mark} {row.get('date')} · {spot}  [#{sid}]"
            labels.append(label)
            row_by_label[label] = row

        choice = st.selectbox("Session", labels, key=f"edit_sess_sel_{sport}")
        row = row_by_label[choice]
        sid = int(row["id"])

        rider = load_profiles().get(user["username"], {})
        spots = rider.get("spots", [])
        boards = rider.get(gear_meta["boards_key"], [])
        gears = rider.get(gear_meta["gear_key"], [])

        def _opts_index(opts, current):
            current = str(current or "").strip()
            lst = ["(empty)"] + list(opts)
            if current and current.lower() not in ("none", "nan", "null") and current not in lst:
                lst.insert(1, current)
            idx = lst.index(current) if current in lst else 0
            return lst, idx

        with st.form(f"edit_sess_form_{sid}"):
            spot_lst, spot_idx = _opts_index(spots, row.get("surfspot"))
            spot_sel = st.selectbox("📍 Spot", spot_lst, index=spot_idx, key=f"es_spot_{sid}")
            spot_new = st.text_input("…or new spot", key=f"es_spot_new_{sid}")

            board_lst, board_idx = _opts_index(boards, row.get("board"))
            board_sel = st.selectbox("🏄 Board", board_lst, index=board_idx, key=f"es_board_{sid}")

            gear_lst, gear_idx = _opts_index(gears, row.get("sail"))
            gear_sel = st.selectbox(f"🎽 {gear_word}", gear_lst, index=gear_idx, key=f"es_gear_{sid}")

            type_opts = ["(empty)", "Fin", "Foil", "Twintip"]
            cur_type = str(row.get("gear_type") or "").strip()
            type_idx = type_opts.index(cur_type) if cur_type in type_opts else 0
            type_sel = st.selectbox("Type", type_opts, index=type_idx, key=f"es_type_{sid}")

            if st.form_submit_button("Save session", use_container_width=True):
                def _pick(sel, new=None):
                    if new is not None and new.strip():
                        return new.strip()
                    return None if sel == "(empty)" else sel

                fields = {
                    "surfspot": _pick(spot_sel, spot_new),
                    "board": _pick(board_sel),
                    "sail": _pick(gear_sel),
                    "gear_type": _pick(type_sel),
                }
                update_session(sid, fields)

                if _session_is_complete(fields["surfspot"], fields["board"], fields["sail"]):
                    st.success("Saved – this session now counts in the ranking. ✅")
                else:
                    st.warning("Saved, but still incomplete – spot, board and "
                               f"{gear_word.lower()} are required for the ranking.")
                st.rerun()

        # Session löschen (z.B. fehlerhafte NaT-/Test-Sessions). Mit Bestätigung.
        confirm_del = st.checkbox("Confirm delete", key=f"es_delconf_{sid}")
        if st.button("🗑 Delete this session", key=f"es_del_{sid}",
                     use_container_width=True, disabled=not confirm_del):
            delete_session(sid)
            st.success("Session deleted.")
            st.rerun()


def render_user_profile(user):
    """Profilbereich: E-Mail/Gewicht/Größe bearbeiten, Passwort ändern und das
    Equipment (für den aktiven Sport) pflegen. Rendert im aktuellen Kontext."""
    sport = active_sport()
    gear_meta = SPORT_META[sport]

    with st.expander("👤 My profile", expanded=False):
        # Profildaten einmal pro Sitzung lesen (get_user ist NICHT gecacht; sonst
        # liefe pro Rerun eine zusätzliche DB-Abfrage – auch bei eingeklapptem
        # Expander). Nach dem Speichern wird der Cache aktualisiert.
        _profile_key = f"_profile_full_{user['id']}"
        if _profile_key not in st.session_state:
            st.session_state[_profile_key] = get_user(user["username"]) or {}
        full = st.session_state[_profile_key]

        # --- Account-Daten ---
        st.markdown("**Account details**")
        with st.form(f"profile_form_{user['id']}"):
            email = st.text_input("Email", value=full.get("email") or "")
            weight = st.number_input(
                "Weight (kg)", min_value=0.0, max_value=300.0, step=0.5,
                value=float(full.get("weight_kg") or 0.0),
                help="0 = not set",
            )
            height = st.number_input(
                "Height (cm)", min_value=0.0, max_value=260.0, step=0.5,
                value=float(full.get("height_cm") or 0.0),
                help="0 = not set",
            )

            if st.form_submit_button("Save profile", use_container_width=True):
                if email and not _valid_email(email):
                    st.error("Please enter a valid email address.")
                else:
                    update_user_account(
                        user["id"], email=email,
                        weight_kg=weight or None, height_cm=height or None,
                    )
                    st.session_state[_profile_key] = get_user(user["username"]) or {}
                    st.success("Profile saved.")

        # --- Passwort ändern ---
        st.markdown("**Change password**")
        with st.form(f"pw_form_{user['id']}"):
            old_pw = st.text_input("Current password", type="password")
            new_pw1 = st.text_input("New password (min. 6 characters)", type="password")
            new_pw2 = st.text_input("Repeat new password", type="password")

            if st.form_submit_button("Change password", use_container_width=True):
                if new_pw1 != new_pw2:
                    st.error("The new passwords do not match.")
                else:
                    ok, msg = change_password(user["id"], old_pw, new_pw1)
                    (st.success if ok else st.error)(msg)

        # --- Geräte-Token für den Upload von der WaterSession-Uhr ---
        st.markdown("**⌚ Watch upload (WaterSession)**")
        st.caption(
            "Enter this token in the watch app settings (Garmin Connect Mobile → "
            "WaterSession → Settings → Device Token), together with the server URL "
            "of your ingest service. Sessions then upload wirelessly after you "
            "stop recording – no USB needed."
        )
        _tok_key = f"_device_token_{user['id']}"
        if _tok_key not in st.session_state:
            st.session_state[_tok_key] = get_or_create_device_token(user["id"])
        st.code(st.session_state[_tok_key], language=None)
        if st.button("Regenerate token", key=f"regen_token_{user['id']}",
                     use_container_width=True):
            st.session_state[_tok_key] = regenerate_device_token(user["id"])
            st.success("New token generated – update it in the watch app.")
            st.rerun()

        # --- Equipment teilen / übernehmen (gleiches Material über Konten) ---
        st.markdown("**🔗 Share equipment (family / friends)**")
        st.caption(
            "Use the same gear across accounts without typing it three times. Give "
            "YOUR code to others so they can import your equipment (spots, boards, "
            "sails…) – or enter someone's code below to import theirs. Only works "
            "when both sides act (= mutual consent)."
        )
        _share_key = f"_share_code_{user['id']}"
        if _share_key not in st.session_state:
            st.session_state[_share_key] = get_or_create_share_code(user["username"])
        st.caption("Your share code:")
        st.code(st.session_state[_share_key], language=None)
        if st.button("Regenerate share code", key=f"regen_share_{user['id']}",
                     use_container_width=True):
            st.session_state[_share_key] = regenerate_share_code(user["username"])
            st.success("New code generated – the old one no longer works.")
            st.rerun()

        with st.form(f"import_equip_{user['id']}"):
            other_code = st.text_input("Import equipment with someone's code")
            if st.form_submit_button("📥 Import equipment", use_container_width=True):
                owner = owner_for_share_code(other_code)
                if not owner:
                    st.error("Unknown code.")
                elif owner == user["username"]:
                    st.warning("That's your own code.")
                else:
                    added = copy_equipment(owner, user["username"])
                    st.success(f"Imported {added} item(s) from {owner}. ✅")
                    st.rerun()

        # --- Equipment (aktiver Sport; Spots sind sportübergreifend) ---
        st.markdown(f"**Equipment · {gear_meta['label']}**")
        st.caption(
            "Add or remove gear per category (each section saves on its own) so "
            "the session upload is faster. Spots are shared across sports."
        )
        rider = load_profiles().get(user["username"], {})
        spots = rider.get("spots", [])
        boards = rider.get(gear_meta["boards_key"], [])
        gear = rider.get(gear_meta["gear_key"], [])

        gear_word = gear_meta["gear_label"]  # Sail/Kite/Wing/Paddle
        gear_unit = gear_meta.get("gear_size_unit", "m²")

        def _merge(kept, item):
            """Behaltene Auswahl + neuer Eintrag, ohne Duplikate/Leereinträge."""
            item = (item or "").strip()
            if item and item not in kept:
                return list(kept) + [item]
            return list(kept)

        def _equip_open(label, key):
            """Aufklappbare Sektion (Standard: eingeklappt). Ersatz für einen
            Expander – verschachtelte Expander sind in Streamlit nicht erlaubt
            (Equipment liegt bereits im „My profile"-Expander), daher ein
            Button-Toggle mit ▸/▾-Indikator. Liefert True, wenn ausgeklappt."""
            st.session_state.setdefault(key, False)
            count = label[1] if isinstance(label, tuple) else None
            text = label[0] if isinstance(label, tuple) else label
            arrow = "▾" if st.session_state[key] else "▸"
            suffix = f"  ({count})" if count is not None else ""
            if st.button(
                f"{arrow}  {text}{suffix}",
                key=f"toggle_{key}", use_container_width=True,
            ):
                st.session_state[key] = not st.session_state[key]
                st.rerun()
            return st.session_state[key]

        # Drei eigenständige, einzeln aufklappbare Sektionen -> Spot, Board und
        # {Gear} getrennt speicherbar (eigener Speichern-Button je Kategorie).

        # --- Spots (geteilt) ---
        if _equip_open(("📍 Spots (shared)", len(spots)), f"sec_spots_{user['id']}_{sport}"):
            with st.form(f"equip_spots_{user['id']}_{sport}"):
                keep_spots = st.multiselect(
                    "Current spots", spots, default=spots, label_visibility="collapsed"
                )
                new_spot = st.text_input("➕ Add spot")
                if st.form_submit_button("Save spots", use_container_width=True):
                    set_profile_list(user["username"], "spots", _merge(keep_spots, new_spot))
                    st.success("Spots updated.")
                    st.rerun()

        # --- Boards (je Sport) ---
        if _equip_open(("🛹 Boards", len(boards)), f"sec_boards_{user['id']}_{sport}"):
            with st.form(f"equip_boards_{user['id']}_{sport}"):
                keep_boards = st.multiselect(
                    "Current boards", boards, default=boards, label_visibility="collapsed"
                )
                nb_brand = st.text_input("Board brand", key=f"pf_board_brand_{sport}")
                nb_model = st.text_input("Board type / model", key=f"pf_board_model_{sport}")
                nb_vol = st.number_input(
                    "Volume in liters (optional)", min_value=0, step=1,
                    key=f"pf_board_vol_{sport}",
                )
                if st.form_submit_button("Save boards", use_container_width=True):
                    new_board = (
                        format_board(nb_brand, nb_model, nb_vol)
                        if nb_brand.strip() and nb_model.strip() else ""
                    )
                    set_profile_list(
                        user["username"], gear_meta["boards_key"], _merge(keep_boards, new_board)
                    )
                    st.success("Boards updated.")
                    st.rerun()

        # --- 2. Material: Sail/Kite/Wing/Paddle (je Sport) ---
        if _equip_open((f"🎽 {gear_word}s", len(gear)), f"sec_gear_{user['id']}_{sport}"):
            with st.form(f"equip_gear_{user['id']}_{sport}"):
                keep_gear = st.multiselect(
                    f"Current {gear_word.lower()}s", gear, default=gear,
                    label_visibility="collapsed",
                )
                ng_brand = st.text_input(f"{gear_word} brand", key=f"pf_gear_brand_{sport}")
                ng_model = st.text_input(f"{gear_word} name / model", key=f"pf_gear_model_{sport}")
                ng_size = 0.0
                if gear_unit:
                    ng_size = st.number_input(
                        f"{gear_word} size in {gear_unit} (optional)", min_value=0.0, step=0.1,
                        key=f"pf_gear_size_{sport}",
                    )
                if st.form_submit_button(f"Save {gear_word.lower()}s", use_container_width=True):
                    new_gear = (
                        format_gear(ng_brand, ng_model, ng_size, gear_unit)
                        if ng_brand.strip() and ng_model.strip() else ""
                    )
                    set_profile_list(
                        user["username"], gear_meta["gear_key"], _merge(keep_gear, new_gear)
                    )
                    st.success(f"{gear_word}s updated.")
                    st.rerun()


def render_account_sidebar(user):
    with st.sidebar:
        st.markdown(f"### 👤 {user['username']}")

        if st.button("Log out", use_container_width=True):
            logout_session()
            st.rerun()

        st.markdown("---")
        render_user_profile(user)

        st.markdown("---")
        st.markdown("### 👥 Groups")
        st.caption("You are always part of the group **All** (all results visible).")

        groups = list_groups()
        group_by_id = {g["id"]: g for g in groups}
        memberships = my_memberships(user["id"])

        member_groups = [
            group_by_id[gid] for gid, status in memberships.items()
            if status == "member" and gid in group_by_id
        ]
        pending_groups = [
            group_by_id[gid] for gid, status in memberships.items()
            if status == "pending" and gid in group_by_id
        ]

        if member_groups:
            st.markdown("**Your groups**")
            for g in member_groups:
                tag = "🔒" if g["is_private"] else "🌍"
                is_owner = g["owner_id"] == user["id"]
                cols = st.columns([4, 2])
                cols[0].write(f"{tag} {g['name']}" + (" · Owner" if is_owner else ""))
                if not is_owner:
                    if cols[1].button("Leave", key=f"leave_{g['id']}"):
                        leave_group(user["id"], g["id"])
                        st.rerun()

        if pending_groups:
            st.markdown("**Requested (waiting for approval)**")
            for g in pending_groups:
                st.write(f"⏳ {g['name']}")

        joinable = [g for g in groups if g["id"] not in memberships]

        if joinable:
            with st.expander("➕ Join a group"):
                option_map = {
                    f'{"🔒" if g["is_private"] else "🌍"} {g["name"]}': g
                    for g in joinable
                }
                choice = st.selectbox("Group", list(option_map.keys()), key="join_select")
                target = option_map[choice]
                btn_label = "Request to join" if target["is_private"] else "Join"

                if st.button(btn_label, key="join_btn", use_container_width=True):
                    ok, message = join_or_request_group(user["id"], target["id"])
                    (st.success if ok else st.warning)(message)
                    if ok:
                        st.rerun()

        with st.expander("🆕 Create a group"):
            grp_name = st.text_input("Group name", key="new_group_name")
            grp_type = st.radio(
                "Type",
                ["Open (anyone can join)", "Private (invite only)"],
                key="new_group_type",
            )

            if st.button("Create", key="new_group_btn", use_container_width=True):
                ok, message = create_group(grp_name, user["id"], grp_type.startswith("Private"))
                (st.success if ok else st.warning)(message)
                if ok:
                    st.rerun()

        owned_groups = [g for g in member_groups if g["owner_id"] == user["id"]]

        if owned_groups:
            with st.expander("🛠️ Manage my groups"):
                for g in owned_groups:
                    tag = "🔒" if g["is_private"] else "🌍"
                    st.markdown(f"**{tag} {g['name']}**")

                    requests = pending_requests(g["id"])

                    if requests:
                        st.caption("Pending join requests:")
                        for req in requests:
                            rc = st.columns([3, 1, 1])
                            rc[0].write(f"⏳ {req['username']}")
                            if rc[1].button("✓", key=f"approve_{g['id']}_{req['user_id']}"):
                                set_membership_status(req["user_id"], g["id"], "member")
                                st.rerun()
                            if rc[2].button("✗", key=f"reject_{g['id']}_{req['user_id']}"):
                                leave_group(req["user_id"], g["id"])
                                st.rerun()
                    else:
                        st.caption("No pending requests.")

                    invite_name = st.text_input(
                        "Invite / approve user (username)",
                        key=f"invite_{g['id']}",
                    )
                    if st.button("Add", key=f"invite_btn_{g['id']}"):
                        ok, message = invite_user(g["id"], invite_name)
                        (st.success if ok else st.warning)(message)
                        if ok:
                            st.rerun()

                    # Gruppe löschen – nur möglich, wenn man allein drin ist.
                    member_count = group_membership_count(g["id"])

                    if member_count and member_count > 1:
                        st.caption(
                            "🗑️ Deletion is only possible when you are the only "
                            "one in the group (remove members/requests first)."
                        )
                    else:
                        if st.button(
                            "🗑️ Delete group",
                            key=f"del_group_btn_{g['id']}",
                            use_container_width=True,
                        ):
                            ok, message = delete_group(g["id"], user["id"])
                            (st.success if ok else st.warning)(message)
                            if ok:
                                st.rerun()

                    st.markdown("---")

        st.markdown("---")

        with st.expander("⚠️ Account & delete data"):
            session_count = count_user_sessions(user["username"], active_sport())

            st.caption(
                f"You currently have {session_count} saved session(s). "
                "Only your own data is deleted – never the results of other "
                "riders."
            )

            # --- Einzelne Session löschen ---
            rider_sessions = load_rider_sessions(user["username"], active_sport())

            if not rider_sessions.empty and "id" in rider_sessions.columns:
                st.markdown("**Delete a single session**")

                del_opts = {"— please choose —": None}

                for i, (_, row) in enumerate(rider_sessions.iterrows()):
                    d = row.get("date")
                    date_str = d.strftime("%Y-%m-%d") if pd.notna(d) else "?"
                    spot = str(row.get("surfspot", "") or "")
                    sp = row.get("speed_1s_kmh")
                    sp_str = "" if sp is None or pd.isna(sp) else f" · {float(sp):.1f} km/h"
                    label = f"{date_str} · {spot}{sp_str}".strip(" ·") or f"Session {i + 1}"

                    if label in del_opts:
                        label = f"{label} ({i + 1})"

                    del_opts[label] = row.get("id")

                pick = st.selectbox(
                    "Choose session", list(del_opts.keys()), key="del_one_pick"
                )

                if pick != "— please choose —":
                    confirm_one = st.checkbox(
                        "Yes, permanently delete this session.",
                        key=f"confirm_del_one_{del_opts[pick]}",
                    )

                    if st.button(
                        f"🗑️ Delete session: {pick}",
                        key=f"del_one_btn_{del_opts[pick]}",
                        use_container_width=True,
                        disabled=not confirm_one,
                    ):
                        if delete_session(del_opts[pick], user["username"]):
                            st.success("Session deleted.")
                        else:
                            st.warning("Session could not be deleted.")
                        st.rerun()

                st.markdown("---")

            # --- Eigene Sessions löschen (Konto bleibt bestehen) ---
            st.markdown("**Delete all my sessions**")
            confirm_data = st.checkbox(
                "Yes, permanently delete all my sessions.",
                key="confirm_delete_data",
            )

            if st.button(
                "🗑️ Delete sessions",
                key="delete_data_btn",
                use_container_width=True,
                disabled=not confirm_data,
            ):
                removed = delete_user_sessions(user["username"], active_sport())
                st.success(f"{removed} session(s) deleted.")
                st.rerun()

            st.markdown("---")

            # --- Komplettes Konto löschen ---
            st.markdown("**Delete account**")
            st.caption(
                "Removes your account, all your sessions, your profile and the "
                "groups you created. This cannot be undone."
            )

            confirm_name = st.text_input(
                "Enter your username to confirm",
                key="confirm_delete_account_name",
                placeholder=user["username"],
            )

            if st.button(
                "❌ Permanently delete account",
                key="delete_account_btn",
                use_container_width=True,
                disabled=confirm_name != user["username"],
            ):
                delete_account(user["id"], user["username"])
                logout_session()
                st.success("Your account and all associated data have been deleted.")
                st.rerun()


# Fehlende Tabellen anlegen (z.B. nach Deploy mit neuen Tabellen) – vor jedem DB-Zugriff.
ensure_schema()
# Uhr-Spalten ungecacht nachziehen (ensure_schema ist @st.cache_resource-gegated).
ensure_watch_columns()

# Login aus „Angemeldet bleiben"-Cookie wiederherstellen.
# st.context.cookies liefert auf der Streamlit Community Cloud leider nichts,
# daher lesen wir das Cookie über die extra-streamlit-components-Komponente –
# ABER nur, wenn noch KEIN User in der Session ist (= frischer Seitenaufruf).
# Sobald eingeloggt, wird nichts gemountet → Reruns/Filter-Interaktionen bleiben
# schnell. Der Komponenten-Roundtrip fällt also nur einmal beim Laden an.
if "user" not in st.session_state:
    cm = _cookie_manager()
    saved_token = None

    if cm is not None:
        try:
            saved_token = cm.get("surf_auth")
        except Exception:
            saved_token = None

    if saved_token:
        restored = user_for_token(saved_token)
        if restored:
            st.session_state["user"] = {
                "id": restored["id"],
                "username": restored["username"],
            }

    # saved_token == None heißt: Cookie noch nicht geladen ODER nicht vorhanden.
    # Einmal auf das Cookie-Rerun der Komponente warten, damit der Login-Screen
    # nicht kurz aufblitzt, während das Cookie noch ankommt.
    if (
        "user" not in st.session_state
        and cm is not None
        and not st.session_state.get("_cookie_probed")
    ):
        st.session_state["_cookie_probed"] = True
        st.markdown(
            "<div style='text-align:center;margin-top:3rem;opacity:.7;'>⏳ one moment…</div>",
            unsafe_allow_html=True,
        )
        st.stop()

current_user = st.session_state.get("user")

if not current_user:
    render_login()
    st.stop()

render_account_sidebar(current_user)

# „Angemeldet bleiben"-Cookie setzen – nur direkt nach dem Login (wenn ein
# _pending_token vorliegt). Per JS aufs Eltern-Dokument geschrieben; die
# Komponente liest es beim nächsten frischen Laden wieder ein.
if st.session_state.get("_pending_token"):
    _persist_auth_cookie(st.session_state["_pending_token"])
    st.session_state.pop("_pending_token", None)


# Reine Spots-Seite (Revierführer) – ersetzt die Rankings; Header/Umschalter oben
# bleiben sichtbar. ensure_schema stellt sicher, dass spot_info.country existiert.
if _is_spots_view:
    ensure_schema()
    render_spots_page()
    st.stop()


# Bestleistungs-Filter direkt unter „Konto & Daten löschen" (Konto-Bereich der
# Sidebar, ÜBER den Tabs) – dort einfacher zu finden. Die zugehörige Tabelle
# erscheint im Hauptfenster (render_personal_best_table weiter unten).
with st.sidebar:
    st.markdown("---")
    pb_table, pb_table_caption, pb_total = render_personal_best_filter(current_user["username"])
    selected_history_record = render_session_history(current_user["username"])


# Linke Sidebar: Konto/Gruppen sind oben bereits gerendert. Darunter zwei
# einklappbare Bereiche (Expander) – konsistent mit Bestleistungen/Sessions
# darüber. Beide standardmäßig zu, damit die Sidebar aufgeräumt bleibt.
with st.sidebar:
    st.markdown("---")
    sidebar_tab_material = st.expander("🏄 Add session", expanded=False)
    sidebar_tab_filter = st.expander("🔎 Filter", expanded=False)


def autocollapse_sidebar():
    """Hält die Sidebar (Filter/Material) während der Start-Reruns eingeklappt.

    ``initial_sidebar_state="collapsed"`` greift nur beim allerersten Lauf einer
    Session. Bei den automatischen Start-Reruns (Cookie-Login über „Angemeldet
    bleiben", Wetterabruf) klappt Streamlit die Sidebar wieder auf. Wir klicken
    daher den vorhandenen Collapse-Button für die ersten Läufe einmalig zu –
    so stehen beim Start zuerst die Rankings im Vordergrund. Danach steuert der
    Nutzer die Sidebar wieder frei (das Budget ist dann aufgebraucht).
    """
    budget = st.session_state.get("_sidebar_collapse_budget", 3)
    if budget <= 0:
        return
    st.session_state["_sidebar_collapse_budget"] = budget - 1
    components.html(
        """
        <script>
        (function () {
            const doc = window.parent.document;
            let tries = 0;
            const timer = setInterval(function () {
                tries += 1;
                const btn = doc.querySelector('button[data-testid="stSidebarCollapseButton"]');
                if (btn) { btn.click(); clearInterval(timer); }
                else if (tries > 40) { clearInterval(timer); }
            }, 25);
        })();
        </script>
        """,
        height=0,
    )


autocollapse_sidebar()

# Reihenfolge für die Ladezeit: Rankings (Hauptinhalt) ZUERST rendern, das
# News-Banner und seine DB-Abfrage erst danach. Mit einem Platzhalter (news_slot)
# bleibt das Banner optisch oben, wird aber erst nach den Rankings befüllt – so
# erscheinen die Rankings als Erstes auf dem Bildschirm.
# Schlanker Profiler: misst die Dauer der Hauptbereiche. Anzeige nur mit ?perf=1.
_PERF = []


def _perf(label, fn, *a, **k):
    _t = time.perf_counter()
    try:
        return fn(*a, **k)
    finally:
        _PERF.append((label, (time.perf_counter() - _t) * 1000.0))


news_slot = st.empty()

ranking_results = st.empty()
with sidebar_tab_filter:
    _perf("rankings", render_rankings, ranking_results)

with news_slot.container():
    _perf("news", render_group_news_banner, current_user)

st.markdown("---")


# „left" = Material-Tab in der Sidebar, „right" = Ergebnis-Bereich in der Mitte.
# So bleiben alle bestehenden „with left:"/„with right:"-Blöcke unverändert gültig.
left = sidebar_tab_material
right = st.container()

# selected_history_record wird oben im Konto-Bereich gesetzt
# (render_session_history); hier NICHT erneut auf None setzen, sonst ginge die
# Auswahl verloren.


with left:
    st.markdown("### 👤 1. Session & equipment")

    name = current_user["username"]
    st.markdown(f"**Rider:** `{name}`")

    profiles = load_profiles()
    rider = profiles.get(name, {})

    # Verlauf („Meine Sessions ansehen") und Bestleistungen liegen jetzt im
    # Konto-Bereich der Sidebar (render_session_history / render_personal_best_*),
    # nicht mehr hier im Material-Tab.

    spot_options = rider.get("spots", [])
    spot_choice = st.selectbox("Surf spot", [NEW_ENTRY] + spot_options)

    if spot_choice == NEW_ENTRY:
        spot = st.text_input("New surf spot")
    else:
        spot = spot_choice

    gear_meta = SPORT_META[sport]
    gear_label = gear_meta["gear_label"]  # "Sail" (Windsurf) / "Kite" (Kite)

    st.markdown("**Board**")
    board_options = rider.get(gear_meta["boards_key"], [])
    board_choice = st.selectbox(
        "Select board", [NEW_ENTRY] + board_options, key=f"board_sel_{sport}"
    )

    if board_choice == NEW_ENTRY:
        board_brand = st.text_input("Board brand", key=f"board_brand_{sport}")
        board_model = st.text_input("Board type / model", key=f"board_model_{sport}")
        board_volume = st.number_input(
            "Volume in liters (optional)", min_value=0, step=1, key=f"board_vol_{sport}"
        )
        board_display = format_board(board_brand, board_model, board_volume)
        board_ok = bool(board_brand.strip() and board_model.strip())
    else:
        board_display = board_choice
        board_ok = True

    gear_type = st.radio(
        gear_meta["gear_type_label"],
        gear_meta["gear_type_options"],
        horizontal=True,
        key=f"gear_type_input_{sport}",
    )

    # Je nach gear_type ein optionales Größenfeld (für späteres Filtern).
    fin_size_cm = None
    foil_front_cm2 = None
    if gear_type == "Fin":
        _fin = st.number_input(
            "Fin size in cm (optional)", min_value=0.0, step=0.5,
            key=f"fin_size_{sport}",
        )
        fin_size_cm = _fin or None
    elif gear_type == "Foil":
        _front = st.number_input(
            "Front wing size in cm² (optional)", min_value=0, step=10,
            key=f"foil_front_{sport}",
        )
        foil_front_cm2 = _front or None

    st.markdown(f"**{gear_label}**")
    sail_options = rider.get(gear_meta["gear_key"], [])
    sail_choice = st.selectbox(
        f"Select {gear_label.lower()}", [NEW_ENTRY] + sail_options, key=f"gear_sel_{sport}"
    )

    gear_unit = gear_meta.get("gear_size_unit", "m²")

    if sail_choice == NEW_ENTRY:
        sail_brand = st.text_input(f"{gear_label} brand", key=f"gear_brand_{sport}")
        sail_model = st.text_input(f"{gear_label} name / model", key=f"gear_model_{sport}")

        if gear_unit:
            sail_size = st.number_input(
                f"{gear_label} size in {gear_unit}", min_value=0.0, step=0.1,
                key=f"gear_size_{sport}",
            )
            sail_display = format_gear(sail_brand, sail_model, sail_size, gear_unit)
            sail_ok = bool(sail_brand.strip() and sail_model.strip() and sail_size > 0)
        else:
            # z.B. SUP-Paddel: keine m²-Größe.
            sail_display = format_gear(sail_brand, sail_model, unit="")
            sail_ok = bool(sail_brand.strip() and sail_model.strip())
    else:
        sail_display = sail_choice
        sail_ok = True

    st.markdown("### ☁️ 2. Load activity")

    st.caption("ℹ️ Currently only **Garmin watches** (FIT files) are supported.")

    fit_source = None
    fit_name = None

    # Uhr-Auslesen (USB/MTP) klappt nur lokal unter Windows – auf einem
    # Server gibt es ausschließlich den Datei-Upload.
    if IS_WINDOWS:
        source = st.radio(
            "Source",
            ["📁 Upload file", "⌚ From watch (USB)"],
            horizontal=True,
        )
    else:
        source = "📁 Upload file"
        st.caption(
            "⌚ **From the watch:** connect the watch via USB, open the "
            "**GARMIN/Activity** folder and upload the desired **.fit file** "
            "below."
        )
        st.caption(
            "🍏 **Apple Watch & others:** export the workout as a **.fit** file "
            "(e.g. via the **HealthFit** app on Apple Watch) and upload it here – "
            "GPS track, distance and the 2 s / 30 s top speeds are computed "
            "automatically. Works with any watch that can export FIT."
        )

    if source == "📁 Upload file":
        uploaded_file = st.file_uploader("Upload FIT file", type=["fit"])

        if uploaded_file is not None:
            fit_source = uploaded_file
            fit_name = uploaded_file.name
    else:
        st.caption(
            "Connect the watch via USB. Whether it shows up as a drive (e.g. "
            "older Edge) or as a device without a drive letter (e.g. Fenix 6 "
            "Pro) – the activities appear automatically. Optionally enter a "
            "folder path."
        )

        manual_folder = st.text_input(
            "Optional: folder path (e.g. E:\\GARMIN\\Activity)"
        )

        activities = []

        # 1) Uhren mit Laufwerksbuchstaben (+ optionaler manueller Ordner)
        for path in find_watch_fit_files(manual_folder.strip() or None)[:50]:
            modified = datetime.fromtimestamp(os.path.getmtime(path))
            activities.append({
                "label": f"{modified:%Y-%m-%d %H:%M}  –  {os.path.basename(path)}",
                "sort": modified.strftime("%Y-%m-%d %H:%M:%S"),
                "kind": "local",
                "path": path,
                "name": os.path.basename(path),
            })

        # 2) MTP-Geräte (Fenix 6 Pro u.ä.)
        for entry in find_mtp_fit_files():
            activities.append({
                "label": (
                    f"{entry['modified'][:16]}  –  {entry['name']}.fit"
                    f"  (⌚ {entry['device']})"
                ),
                "sort": entry["modified"],
                "kind": "mtp",
                "device": entry["device"],
                "name": entry["name"],
            })

        activities.sort(key=lambda a: a["sort"], reverse=True)

        if not activities:
            st.warning(
                "No FIT files found. Connect the watch or enter a folder path."
            )
        else:
            labels = {}

            for i, activity in enumerate(activities):
                label = activity["label"]

                if label in labels:
                    label = f"{label} ({i + 1})"

                labels[label] = activity

            choice = st.selectbox(
                "Select activity (newest first)",
                list(labels.keys()),
            )

            selected = labels[choice]

            if selected["kind"] == "local":
                fit_source = selected["path"]
                fit_name = selected["name"]
            else:
                with st.spinner("Copying file from the watch …"):
                    copied = copy_mtp_file(selected["device"], selected["name"])

                if copied:
                    fit_source = copied
                    fit_name = f"{selected['name']}.fit"
                else:
                    st.error(
                        "The file could not be copied from the watch. "
                        "Reconnect the watch and try again."
                    )


# Bestleistungen als Fragment: Filter in der Sidebar (Filter-Tab, unter den
# Ranking-Filtern), Tabelle oben im Hauptfenster-Container `right`. Ein
# Filterklick rerunt nur dieses Fragment.
with right:
    _perf("pb_table", render_personal_best_table, pb_table, pb_table_caption, pb_total)
    # Session-Editor dezent unter den Personal Bests (statt prominent oben).
    if current_user:
        _perf("editor", render_session_editor, current_user)

# Profiler-Ausgabe (nur mit ?perf=1): zeigt die Dauer der Hauptbereiche.
if st.query_params.get("perf") and _PERF:
    st.caption("⏱ " + " · ".join(f"{lbl}: {ms:.0f} ms" for lbl, ms in _PERF)
               + f"  (Σ {sum(ms for _, ms in _PERF):.0f} ms)")


required_ok = all([
    spot.strip(),
    board_ok,
    sail_ok,
])


if fit_source is not None:
    if not required_ok:
        st.warning("Please fully enter surf spot, board and sail first.")

    df = read_fit_file(fit_source)

    if df.empty:
        st.error("The file contains no record data.")
        st.stop()

    max_speed = None
    avg_speed = None
    best_1s = None
    best_30s = None

    if "speed_kmh" in df.columns:
        max_speed = df["speed_kmh"].max()
        avg_speed = df["speed_kmh"].mean()
        best_1s = best_average_speed(df, 1)
        best_30s = best_average_speed(df, 30)

    distance_km = None

    if "distance" in df.columns:
        distance_km = df["distance"].max() / 1000

    session_date = None
    weather_iso = None

    if "timestamp" in df.columns:
        timestamps = df["timestamp"].dropna()

        if not timestamps.empty:
            session_date = timestamps.min().strftime("%Y-%m-%d")

    surf_time = surf_weather_time(df)

    if surf_time is not None:
        weather_iso = surf_time.strftime("%Y-%m-%dT%H:%M")

    session_lat = None
    session_lon = None

    if "lat" in df.columns and "lon" in df.columns:
        gps = df.dropna(subset=["lat", "lon"])

        if not gps.empty:
            session_lat = float(gps["lat"].iloc[0])
            session_lon = float(gps["lon"].iloc[0])

    weather = None

    if weather_iso and session_lat is not None and session_lon is not None:
        weather = get_weather(session_lat, session_lon, weather_iso)

    runs_df = detect_runs(df)

    longest_run_m = None
    longest_run_km = None

    if not runs_df.empty:
        longest_run_m = runs_df["Distance m"].max()
        longest_run_km = longest_run_m / 1000

    # Trust Score: Plausibilität der Aufzeichnung (inkl. Vergleich mit Spot-Bestwert).
    spot_top_kmh = None
    try:
        _spot = (spot or "").strip()
        _all = load_sessions(active_sport())
        if _spot and not _all.empty and {"surfspot", "speed_1s_kmh"}.issubset(_all.columns):
            _m = pd.to_numeric(
                _all.loc[_all["surfspot"].astype(str) == _spot, "speed_1s_kmh"],
                errors="coerce",
            ).max()
            if pd.notna(_m):
                spot_top_kmh = float(_m)
    except Exception:
        spot_top_kmh = None

    trust = compute_trust_score(df, spot_top_kmh)

    with right:
        st.markdown("## 🌊 Session overview")

        c1, c2, c3, c4 = st.columns(4)

        c1.metric(
            "Max Speed",
            "No data" if max_speed is None else f"{max_speed:.2f} km/h"
        )

        c2.metric(
            "Ø Speed",
            "No data" if avg_speed is None else f"{avg_speed:.2f} km/h"
        )

        c3.metric(
            "Total distance",
            "No data" if distance_km is None else f"{distance_km:.2f} km"
        )

        c4.metric(
            "Longest run",
            "No data" if longest_run_km is None else f"{longest_run_km:.2f} km"
        )

        st.markdown("## 🔍 Trust Score (plausibility)")
        render_trust_score(trust)

        st.markdown("## 🌦️ Weather during session")

        if weather is None:
            if session_lat is None:
                st.info("No GPS data – weather cannot be retrieved.")
            else:
                st.info("No weather data available (maybe no archive entry yet).")
        else:
            emoji, description = WEATHER_CODES.get(weather["code"], ("❓", "Unknown"))

            w1, w2, w3, w4 = st.columns(4)

            w1.metric(
                "Wind",
                "–" if weather["wind"] is None else f"{weather['wind']:.0f} km/h",
                None if weather["gust"] is None else f"Gusts {weather['gust']:.0f} km/h",
            )

            w2.metric(
                "Temperature",
                "–" if weather["temp"] is None else f"{weather['temp']:.1f} °C",
            )

            w3.metric(
                "Precipitation",
                "–" if weather["precip"] is None else f"{weather['precip']:.1f} mm",
            )

            w4.metric(
                "Condition",
                emoji,
                description,
                delta_color="off",
            )

            if weather["dir"] is not None:
                st.caption(
                    f"🧭 Wind from **{degrees_to_compass(weather['dir'])}** "
                    f"{wind_arrow(weather['dir'])} ({weather['dir']:.0f}°)"
                )

        st.markdown("## ⚡ Speed values")

        speed_table = pd.DataFrame([
            {
                "Category": "2 seconds",
                "Speed km/h": best_1s,
                "Speed kn": None if best_1s is None else best_1s / 1.852,
            },
            {
                "Category": "30 seconds",
                "Speed km/h": best_30s,
                "Speed kn": None if best_30s is None else best_30s / 1.852,
            },
        ])

        st.dataframe(
            speed_table.round(2),
            width="stretch",
            hide_index=True,
            height=df_height(len(speed_table)),
        )

        with st.expander("🌊 Detected runs / individual legs", expanded=False):
            if not runs_df.empty:
                show_runs = runs_df.copy()
                show_runs["Distance m"] = show_runs["Distance m"].round(2)
                show_runs["Distance km"] = show_runs["Distance km"].round(3)
                show_runs["Max Speed km/h"] = show_runs["Max Speed km/h"].round(2)
                show_runs["Ø Speed km/h"] = show_runs["Ø Speed km/h"].round(2)
                show_runs = show_runs.sort_values("Distance m", ascending=False).reset_index(drop=True)
                show_runs.insert(0, "Run", show_runs.index + 1)

                st.dataframe(
                    show_runs,
                    width="stretch",
                    hide_index=True,
                    height=df_height(len(show_runs)),
                )
            else:
                st.info("No runs detected. You may need to adjust the thresholds.")

        if st.session_state.get("just_added") == fit_name:
            st.success(
                f"✅ Session was added to the online ranking: **{fit_name}**."
            )

            achievements = st.session_state.get("last_achievements")
            if achievements:
                render_achievements(achievements)
        elif session_exists(fit_name, sport=active_sport()):
            st.info(
                f"⚠️ This file has already been uploaded: **{fit_name}**. "
                "It cannot be added to the ranking a second time."
            )
        elif required_ok and best_30s is not None:
            if st.button("🏆 Add session to the online ranking"):
                entry = {
                    "sport": active_sport(),
                    "date": session_date,
                    "name": name.strip(),
                    "surfspot": spot.strip(),
                    "board": board_display,
                    "sail": sail_display,
                    "gear_type": gear_type,
                    "fin_size_cm": fin_size_cm,
                    "foil_front_cm2": foil_front_cm2,
                    "filename": fit_name,
                    "total_distance_km": None if distance_km is None else round(distance_km, 2),
                    "longest_run_km": None if longest_run_km is None else round(longest_run_km, 3),
                    "longest_run_m": None if longest_run_m is None else round(longest_run_m, 2),
                    "speed_1s_kmh": None if best_1s is None else round(best_1s, 2),
                    "speed_1s_kn": None if best_1s is None else round(best_1s / 1.852, 2),
                    "speed_30s_kmh": None if best_30s is None else round(best_30s, 2),
                    "speed_30s_kn": None if best_30s is None else round(best_30s / 1.852, 2),
                    "wind_kmh": None if weather is None or weather["wind"] is None else round(weather["wind"], 1),
                    "gust_kmh": None if weather is None or weather["gust"] is None else round(weather["gust"], 1),
                    "wind_dir_deg": None if weather is None or weather["dir"] is None else round(weather["dir"]),
                    "temp_c": None if weather is None or weather["temp"] is None else round(weather["temp"], 1),
                    "precip_mm": None if weather is None or weather["precip"] is None else round(weather["precip"], 1),
                    "weather_code": None if weather is None or weather["code"] is None else int(weather["code"]),
                    "trust_score": trust.get("score"),
                }

                # Rekorde VOR dem Speichern erkennen (Vergleich mit dem Bestand).
                member_groups = my_member_groups(current_user["id"])
                achievements = detect_records(entry, load_sessions(active_sport()), member_groups)

                save_session(entry)
                update_profile(name.strip(), spot.strip(), board_display, sail_display, active_sport())
                update_spot_coords(spot.strip(), session_lat, session_lon)

                # Top-3/Rekord-Ereignisse für die Gruppen-Mitglieder hinterlegen.
                record_group_events(name.strip(), achievements["group_events"])

                st.session_state["last_achievements"] = achievements
                st.session_state["ranking_flash"] = (
                    "Session was saved in the online ranking."
                )
                st.session_state["just_added"] = fit_name
                st.rerun()

    st.markdown("---")

    with st.expander("📍 Track on map", expanded=False):
        show_map(df)

    with st.expander("📋 Raw data", expanded=False):
        st.dataframe(
            df.head(100),
            width="stretch",
            height=df_height(min(len(df), 100)),
        )

        csv = df.to_csv(index=False).encode("utf-8")

        st.download_button(
            "⬇️ Download raw data as CSV",
            data=csv,
            file_name="fit_export.csv",
            mime="text/csv",
        )

else:
    with right:
        if selected_history_record is not None:
            render_history_overview(selected_history_record)
        else:
            st.info(
                "Upload a FIT file in the sidebar on the left in the "
                "**🏄 Add session** tab (choose equipment & upload)."
            )
            st.info("After the upload the analysis appears here.")


st.markdown("---")

with st.expander("🌦️ Spot weather (current & forecast)", expanded=False):
    spot_list = all_known_spots()

    if not spot_list:
        st.info("No spots saved yet. Save a session with a surf spot first.")
    else:
        selected_spot = st.selectbox("Select spot", spot_list, key="weather_spot")

        spot_lat, spot_lon = resolve_spot_coords(selected_spot)

        if spot_lat is None:
            st.warning(
                f"No coordinates could be determined for \"{selected_spot}\". "
                "Save a session with GPS at this spot or give it a more specific name."
            )
        else:
            st.markdown(f"**Current & forecast** &nbsp; 📍 {spot_lat:.3f}, {spot_lon:.3f}")
            render_weather_browser(spot_lat, spot_lon)
            st.caption(
                "The weather is loaded directly in your browser from Open-Meteo "
                "(via your own IP, independent of the shared server limit)."
            )


st.markdown("---")

st.markdown(f"""
<div class="footer">
    <h3 style="color:white;font-weight:800;letter-spacing:-0.4px;">MyWaterSessions<span style="color:#2bd4d9;">.</span>{BETA_BADGE}</h3>
    <p style="text-transform:uppercase;letter-spacing:2px;font-size:12px;opacity:.8;">{SPORT_META[sport]["title"]} · The home for everyone active on the water</p>
    <p style="margin-top:.75rem;">
        <a href="?seite=impressum" target="_self" style="color:#2bd4d9;">Impressum</a>
        &nbsp;·&nbsp;
        <a href="?seite=datenschutz" target="_self" style="color:#2bd4d9;">Datenschutzerklärung</a>
    </p>
</div>
""", unsafe_allow_html=True)