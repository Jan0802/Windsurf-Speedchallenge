import base64
import glob
import hashlib
import json
import logging
import os
import secrets
import tempfile
import time
from datetime import datetime, timedelta
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

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
    page_title="Windsurf Speed Challenge",
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


def background_data_uri():
    """Findet das Hintergrundbild in mehreren Formaten und liefert die Data-URI.

    Reihenfolge der Kandidaten = Priorität. Du kannst das Bild als
    assets/background.webp ODER .jpg/.jpeg/.png ablegen (Fallback: header.*).
    Die Existenz-/mtime-Prüfung läuft ungecacht (billig); nur das Kodieren ist
    gecacht und invalidiert beim Bildwechsel automatisch.
    """
    candidates = [
        ("assets", "background.webp"),
        ("assets", "background.jpg"),
        ("assets", "background.jpeg"),
        ("assets", "background.png"),
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
    Column("name", String(80)),  # = Benutzername des Fahrers
    Column("date", String(20)),
    Column("surfspot", String(200)),
    Column("board", String(200)),
    Column("sail", String(200)),
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

        for table in DB_METADATA.tables.values():
            existing = {c["name"] for c in inspector.get_columns(table.name)}

            for column in table.columns:
                if column.name in existing:
                    continue

                col_type = column.type.compile(engine.dialect)

                with engine.begin() as conn:
                    conn.execute(text(
                        f'ALTER TABLE {table.name} ADD COLUMN "{column.name}" {col_type}'
                    ))
    except Exception:
        # Migration ist best-effort; ein Fehler darf die App nicht blockieren.
        pass

    return True


# ---- Sessions ----

def save_session(entry):
    # _py: numpy-Typen -> natives Python (Postgres kann numpy nicht binden)
    values = {k: _py(entry.get(k)) for k in SESSION_FIELDS if k in entry}

    with get_engine().begin() as conn:
        conn.execute(insert(sessions_table).values(**values))

    clear_data_caches()


@st.cache_data(ttl=3600, show_spinner=False)
def load_sessions():
    with get_engine().connect() as conn:
        rows = conn.execute(select(sessions_table)).mappings().all()

    if not rows:
        return pd.DataFrame()

    return pd.DataFrame([dict(r) for r in rows])


def current_username():
    user = st.session_state.get("user")
    return user["username"] if user else None


def session_exists(filename, name=None):
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

        return bool(conn.execute(query).scalar())


@st.cache_data(ttl=3600, show_spinner=False)
def load_rider_sessions(name):
    """Alle gespeicherten Sessions eines Fahrers, neueste zuerst."""
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(sessions_table).where(sessions_table.c.name == name)
        ).mappings().all()

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame([dict(r) for r in rows])

    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df = df.sort_values("date", ascending=False)

    return df.reset_index(drop=True)


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


def update_profile(name, spot, board, sail):
    if not name:
        return

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
            rider = {"spots": [], "boards": [], "sails": []}

        for key, value in (("spots", spot), ("boards", board), ("sails", sail)):
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


def register_user(username, password):
    username = (username or "").strip()

    if not username or not password:
        return False, "Username and password must not be empty."

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
        ))

    return True, "Registration successful – you can now log in."


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


# ---- Konto-/Datenlöschung (immer nur eigene Daten) ----

def count_user_sessions(name):
    if not name:
        return 0

    with get_engine().connect() as conn:
        return conn.execute(
            select(func.count())
            .select_from(sessions_table)
            .where(sessions_table.c.name == name)
        ).scalar()


def delete_user_sessions(name):
    """Löscht alle hochgeladenen Sessions des Fahrers. Gibt die Anzahl zurück."""
    if not name:
        return 0

    with get_engine().begin() as conn:
        result = conn.execute(
            delete(sessions_table).where(sessions_table.c.name == name)
        )

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
    {"key": "speed_1s_kmh", "label": "Top speed 1 s", "unit": "km/h", "decimals": 2},
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
            pd.to_datetime(df["date"], errors="coerce").dt.year
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


def unseen_group_events(user_id, username):
    """Noch ungelesene Gruppen-Ereignisse (nicht die eigenen) für das Banner."""
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


def personal_bests(name, spot=None, year=None):
    """Beste Werte je Metrik für einen Fahrer, optional nach Spot/Jahr gefiltert."""
    df = load_rider_sessions(name)

    if df.empty:
        return []

    if spot and spot != "All" and "surfspot" in df.columns:
        df = df[df["surfspot"].astype(str) == spot]

    if year and year != "All" and "date" in df.columns:
        df = df[pd.to_datetime(df["date"], errors="coerce").dt.year == int(year)]

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
    data = df.copy()

    if spot and spot != "All" and "surfspot" in data.columns:
        data = data[data["surfspot"].astype(str) == spot]

    if year and year != "All" and "date" in data.columns:
        data = data[pd.to_datetime(data["date"], errors="coerce").dt.year == int(year)]

    if board and board != "All" and "board" in data.columns:
        data = data[data["board"].astype(str) == board]

    if max_bft is not None and "wind_kmh" in data.columns:
        wind = pd.to_numeric(data["wind_kmh"], errors="coerce")
        bft = wind.apply(lambda v: kmh_to_beaufort(v) if pd.notna(v) else float("nan"))
        data = data[bft <= max_bft]  # NaN <= x ist False -> Sessions ohne Wind raus

    if "speed_1s_kmh" not in data.columns:
        return pd.DataFrame(), ""

    data = data.copy()
    data["_s1"] = pd.to_numeric(data["speed_1s_kmh"], errors="coerce")
    data = data.dropna(subset=["_s1"]).sort_values("_s1", ascending=False).head(limit)

    if data.empty:
        return pd.DataFrame(), ""

    out = pd.DataFrame()
    out["Rank"] = range(1, len(data) + 1)

    if "date" in data.columns:
        out["Date"] = [
            "" if pd.isna(d) else d.strftime("%Y-%m-%d")
            for d in pd.to_datetime(data["date"], errors="coerce")
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

    out["Top 1 s (km/h)"] = data["_s1"].round(2).values
    out["1 s (kn)"] = (data["_s1"] / 1.852).round(2).values

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

    return out, caption


def render_personal_best_filter(name):
    """Bestleistungs-Filter (Spot/Jahr/Board/Wind) – wird im Konto-Bereich der
    Sidebar gerendert (unter „Konto & Daten löschen", einfacher zu finden).

    Gibt (Anzeige-DataFrame, Caption) zurück; die eigentliche Tabelle wird separat
    im Hauptfenster über render_personal_best_table() angezeigt. Bewusst KEIN
    Fragment mehr (eine Filteränderung löst einen normalen Rerun aus) – die
    kleine Bestleistungs-Tabelle macht das unkritisch, und so lässt sich der
    Filter frei im Konto-Bereich platzieren.
    """
    pb_df = load_rider_sessions(name)

    with st.expander("🏅 Personal Bests", expanded=False):
        if pb_df.empty:
            st.info("No sessions yet – upload a FIT file.")
            return None, ""

        pb_spots = sorted(
            {str(s) for s in pb_df["surfspot"].dropna().astype(str) if str(s).strip()}
            if "surfspot" in pb_df.columns else set()
        )
        pb_years = sorted(
            {int(y) for y in pd.to_datetime(pb_df["date"], errors="coerce").dt.year.dropna()}
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

        pb_table, pb_table_caption = personal_best_table(
            pb_df, spot_pb, year_pb, board_pb, max_bft_pb
        )
        st.caption("➡️ Your top 10 speed table is shown in the main window.")

    return pb_table, pb_table_caption


def render_personal_best_table(pb_table, pb_table_caption):
    """Zeigt die Top-10-Bestleistungs-Tabelle im Hauptfenster (gefiltert über
    den Filter im Konto-Bereich)."""
    if pb_table is None:
        return

    st.markdown("## 🏅 Personal Bests")

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

    with st.expander("📅 View my sessions", expanded=False):
        history = load_rider_sessions(name)

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
                "sail": "Sail",
                "speed_30s_kmh": "30s km/h",
                "speed_1s_kmh": "1s km/h",
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

    ranking = load_sessions()

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
    group_choice = st.session_state.get("rank_group", preset.get("group") or ALL_GROUP)
    spot_filter = st.session_state.get("rank_spot", preset.get("spot") or "Overall")
    year_filter = st.session_state.get("rank_year", preset.get("year") or "All years")
    month_filter = st.session_state.get("rank_month", preset.get("month") or "Whole year")
    day_filter = st.session_state.get("rank_day", preset.get("day") or "Whole month")

    # ---- Tabellen ZUERST (Hauptinhalt) in den Haupt-Container ----
    # Reine Anzeige (keine Widgets) -> aus dem Fragment in externen Container ok;
    # .container() im st.empty()-Platzhalter ersetzt den Inhalt bei jedem Rerun.
    with results_container.container():
        _render_ranking_tables(
            ranking, group_choice, member_groups, months,
            spot_filter, year_filter, month_filter, day_filter,
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
                    "group": st.session_state.get("rank_group", ALL_GROUP),
                    "spot": st.session_state.get("rank_spot", "Overall"),
                    "year": st.session_state.get("rank_year", "All years"),
                    "month": st.session_state.get("rank_month", "Whole year"),
                    "day": st.session_state.get("rank_day", "Whole month"),
                })
                st.success("Saved – will be loaded on start from now on.")

            if preset and st.button("↺ Reset", use_container_width=True):
                delete_user_pref(username)
                for _k in ("rank_group", "rank_spot", "rank_year", "rank_month", "rank_day"):
                    st.session_state.pop(_k, None)
                st.rerun()

        group_options = [ALL_GROUP] + [g["name"] for g in member_groups]
        st.selectbox(
            "👥 Group",
            group_options,
            index=_preset_index(group_options, group_choice),
            key="rank_group",
            help="\"All\" shows every rider. You only see group results as a member.",
        )

        # Auswahl-Optionen für Lokation/Jahr abhängig von der gewählten Gruppe.
        if not ranking.empty:
            opt_df = ranking.copy()
            opt_df["_date"] = (
                pd.to_datetime(opt_df["date"], errors="coerce")
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
            index=_preset_index(spot_options, spot_filter), key="rank_spot",
        )
        st.selectbox(
            "📅 Year", year_options,
            index=_preset_index(year_options, year_filter), key="rank_year",
        )
        st.selectbox(
            "📆 Month", month_options,
            index=_preset_index(month_options, month_filter), key="rank_month",
        )
        st.selectbox(
            "🗓️ Day", day_options,
            index=_preset_index(day_options, day_filter), key="rank_day",
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
                           spot_filter, year_filter, month_filter, day_filter):
    """Rendert die vier Ranking-Tabellen – reine Anzeige (keine Widgets)."""
    st.markdown("## 🏆 Online rankings")

    flash = st.session_state.pop("ranking_flash", None)

    if flash:
        st.success(flash)

    if ranking.empty:
        st.info("No online ranking entries yet.")
        return

    if "date" not in ranking.columns:
        ranking["date"] = ""

    for column in ("wind_kmh", "wind_dir_deg", "temp_c", "weather_code", "trust_score"):
        if column not in ranking.columns:
            ranking[column] = None

    ranking["_date"] = pd.to_datetime(ranking["date"], errors="coerce")

    if group_choice != ALL_GROUP:
        group_id = next((g["id"] for g in member_groups if g["name"] == group_choice), None)

        if group_id is not None:
            member_names = set(group_member_names(group_id))
            ranking = ranking[ranking["name"].astype(str).isin(member_names)].copy()

            if ranking.empty:
                st.info(f"The group \"{group_choice}\" has no sessions yet.")
                return

    if spot_filter != "Overall":
        ranking = ranking[ranking["surfspot"].astype(str) == spot_filter]

    if year_filter != "All years":
        ranking = ranking[ranking["_date"].dt.year == int(year_filter)]

    if month_filter != "Whole year":
        ranking = ranking[ranking["_date"].dt.month == months.index(month_filter) + 1]

        if day_filter != "Whole month":
            ranking = ranking[ranking["_date"].dt.day == int(day_filter)]

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
            "sail": "Sail",
            "speed_30s_kmh": "30s km/h",
            "speed_30s_kn": "30s kn",
        })

        st.dataframe(r30, width="stretch", hide_index=True, height=df_height(len(r30)))

    with rcol2:
        st.markdown("### ⚡ Top speed 1 second")

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
            "sail": "Sail",
            "speed_1s_kmh": "1s km/h",
            "speed_1s_kn": "1s kn",
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
            "sail": "Sail",
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


def render_history_overview(record):
    """Session-Übersicht aus den gespeicherten Ranking-Werten (ohne Roh-FIT)."""

    def num(key):
        value = record.get(key)
        return None if value is None or pd.isna(value) else float(value)

    st.markdown("## 🌊 Session overview")

    meta_bits = []
    date_value = record.get("date")

    if date_value is not None and pd.notna(date_value):
        meta_bits.append(f"📅 {pd.to_datetime(date_value):%Y-%m-%d}")

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
    c1.metric("Top 1 s", "–" if best_1s is None else f"{best_1s:.2f} km/h")
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
            "Category": "1 second",
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

    st.caption(
        "ℹ️ The map, individual runs and max/avg speed are only available right "
        "after the upload – for saved sessions the stored metrics are shown."
    )


load_css(app_path("assets", "style.css"))

logo_img = image_to_base64(app_path("assets", "windsurfer.png"))

# Vollflächiges Hintergrundbild (Wasser/Surfer). Lege dein Wunschfoto als
# assets/background.webp ODER .jpg/.jpeg/.png ab (Fallback: header.*). Der
# MIME-Typ wird automatisch passend zur Datei gesetzt.
bg_uri = background_data_uri()

if bg_uri:
    st.markdown(
        f"""
<style>
.stApp {{
    /* cover = bildschirmfüllend; bei hochkantigem Bild auf Querformat-Schirm
       wird der mittlere Ausschnitt gezeigt (center center). Füllt komplett,
       keine Balken. */
    background-color: #02162b;
    background-image: linear-gradient(rgba(2,22,43,.45), rgba(2,22,43,.62)),
                      url("{bg_uri}");
    background-position: center center;
    background-size: cover;
    background-repeat: no-repeat;
    background-attachment: fixed;
}}
</style>
""",
        unsafe_allow_html=True,
    )

if logo_img:
    logo_icon = (
        f'<img src="data:image/png;base64,{logo_img}" '
        'style="height:1em;vertical-align:-0.15em;margin-right:.15em;" alt="">'
    )
else:
    logo_icon = "🏄"

st.markdown(f"""
<div class="hero">
    <div class="hero-content">
        <div class="logo">{logo_icon} WINDSURF</div>
        <div class="title">SPEED CHALLENGE</div>
        <p>Track your sessions. Compare your personal bests.<br>
        The community. Your spot. Your speed.</p>
        <div class="hero-nav">
            <span>⚡ Speed</span>
            <span>🏆 Ranking</span>
            <span>📍 Spots</span>
            <span>👥 Community</span>
        </div>
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
    st.markdown("## 🔐 Sign in")
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
                ok, message = register_user(new_username, pwd1)

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


def render_account_sidebar(user):
    with st.sidebar:
        st.markdown(f"### 👤 {user['username']}")

        if st.button("Log out", use_container_width=True):
            logout_session()
            st.rerun()

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
            session_count = count_user_sessions(user["username"])

            st.caption(
                f"You currently have {session_count} saved session(s). "
                "Only your own data is deleted – never the results of other "
                "riders."
            )

            # --- Einzelne Session löschen ---
            rider_sessions = load_rider_sessions(user["username"])

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
                removed = delete_user_sessions(user["username"])
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


# Bestleistungs-Filter direkt unter „Konto & Daten löschen" (Konto-Bereich der
# Sidebar, ÜBER den Tabs) – dort einfacher zu finden. Die zugehörige Tabelle
# erscheint im Hauptfenster (render_personal_best_table weiter unten).
with st.sidebar:
    st.markdown("---")
    pb_table, pb_table_caption = render_personal_best_filter(current_user["username"])
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
news_slot = st.empty()

ranking_results = st.empty()
with sidebar_tab_filter:
    render_rankings(ranking_results)

with news_slot.container():
    render_group_news_banner(current_user)

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

    st.markdown("**Board**")
    board_options = rider.get("boards", [])
    board_choice = st.selectbox("Select board", [NEW_ENTRY] + board_options)

    if board_choice == NEW_ENTRY:
        board_brand = st.text_input("Board brand")
        board_model = st.text_input("Board type / model")
        board_volume = st.number_input("Volume in liters", min_value=0, step=1)
        board_display = f"{board_brand.strip()} {board_model.strip()} {board_volume}L".strip()
        board_ok = bool(board_brand.strip() and board_model.strip() and board_volume > 0)
    else:
        board_display = board_choice
        board_ok = True

    st.markdown("**Sail**")
    sail_options = rider.get("sails", [])
    sail_choice = st.selectbox("Select sail", [NEW_ENTRY] + sail_options)

    if sail_choice == NEW_ENTRY:
        sail_brand = st.text_input("Sail brand")
        sail_model = st.text_input("Sail name / model")
        sail_size = st.number_input("Sail size in m²", min_value=0.0, step=0.1)
        sail_display = f"{sail_brand.strip()} {sail_model.strip()} {sail_size:.1f} m²".strip()
        sail_ok = bool(sail_brand.strip() and sail_model.strip() and sail_size > 0)
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
    render_personal_best_table(pb_table, pb_table_caption)


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
        _all = load_sessions()
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
                "Category": "1 second",
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
        elif session_exists(fit_name):
            st.info(
                f"⚠️ This file has already been uploaded: **{fit_name}**. "
                "It cannot be added to the ranking a second time."
            )
        elif required_ok and best_30s is not None:
            if st.button("🏆 Add session to the online ranking"):
                entry = {
                    "date": session_date,
                    "name": name.strip(),
                    "surfspot": spot.strip(),
                    "board": board_display,
                    "sail": sail_display,
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
                achievements = detect_records(entry, load_sessions(), member_groups)

                save_session(entry)
                update_profile(name.strip(), spot.strip(), board_display, sail_display)
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
    <h3 style="color:white;">{logo_icon} WINDSURF SPEED CHALLENGE</h3>
    <p>The community. Your spot. Your speed.</p>
    <p style="margin-top:.75rem;">
        <a href="?seite=impressum" target="_self" style="color:#2bd4d9;">Impressum</a>
        &nbsp;·&nbsp;
        <a href="?seite=datenschutz" target="_self" style="color:#2bd4d9;">Datenschutzerklärung</a>
    </p>
</div>
""", unsafe_allow_html=True)