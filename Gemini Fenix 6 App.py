import base64
import glob
import hashlib
import json
import os
import secrets
import tempfile
from datetime import datetime, timedelta
from urllib.parse import urlencode
from urllib.request import urlopen

import pandas as pd
import pydeck as pdk
import streamlit as st
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
    select,
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

NEW_ENTRY = "➕ Neu eingeben..."

# WMO-Wettercodes -> (Emoji, Beschreibung)
WEATHER_CODES = {
    0: ("☀️", "Klar"),
    1: ("🌤️", "Überwiegend klar"),
    2: ("⛅", "Teils bewölkt"),
    3: ("☁️", "Bedeckt"),
    45: ("🌫️", "Nebel"),
    48: ("🌫️", "Reifnebel"),
    51: ("🌦️", "Leichter Nieselregen"),
    53: ("🌦️", "Nieselregen"),
    55: ("🌧️", "Starker Nieselregen"),
    61: ("🌦️", "Leichter Regen"),
    63: ("🌧️", "Regen"),
    65: ("🌧️", "Starker Regen"),
    66: ("🌧️", "Gefrierender Regen"),
    67: ("🌧️", "Starker gefrierender Regen"),
    71: ("🌨️", "Leichter Schneefall"),
    73: ("🌨️", "Schneefall"),
    75: ("❄️", "Starker Schneefall"),
    77: ("🌨️", "Schneegriesel"),
    80: ("🌦️", "Leichte Schauer"),
    81: ("🌧️", "Schauer"),
    82: ("⛈️", "Heftige Schauer"),
    85: ("🌨️", "Schneeschauer"),
    86: ("❄️", "Starke Schneeschauer"),
    95: ("⛈️", "Gewitter"),
    96: ("⛈️", "Gewitter mit Hagel"),
    99: ("⛈️", "Schweres Gewitter mit Hagel"),
}

WEEKDAYS = [
    "Montag", "Dienstag", "Mittwoch", "Donnerstag",
    "Freitag", "Samstag", "Sonntag",
]

# 16-Punkte-Kompass (deutsche Abkürzungen, O = Ost)
COMPASS = [
    "N", "NNO", "NO", "ONO", "O", "OSO", "SO", "SSO",
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
    layout="wide"
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
    connect_args = {"check_same_thread": False} if url.startswith("sqlite") else {}
    engine = create_engine(url, connect_args=connect_args, pool_pre_ping=True)
    DB_METADATA.create_all(engine)
    _migrate_legacy(engine)
    return engine


# ---- Sessions ----

def save_session(entry):
    # _py: numpy-Typen -> natives Python (Postgres kann numpy nicht binden)
    values = {k: _py(entry.get(k)) for k in SESSION_FIELDS if k in entry}

    with get_engine().begin() as conn:
        conn.execute(insert(sessions_table).values(**values))


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

    profiles = load_profiles()
    rider = profiles.get(name) or {"spots": [], "boards": [], "sails": []}

    for key, value in (("spots", spot), ("boards", board), ("sails", sail)):
        items = rider.setdefault(key, [])

        if value and value not in items:
            items.append(value)

    data = json.dumps(rider, ensure_ascii=False)

    with get_engine().begin() as conn:
        exists = conn.execute(
            select(profiles_table.c.name).where(profiles_table.c.name == name)
        ).first()

        if exists:
            conn.execute(
                update(profiles_table).where(profiles_table.c.name == name).values(data=data)
            )
        else:
            conn.execute(insert(profiles_table).values(name=name, data=data))


# ---- Spots (Koordinaten-Cache) ----

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
        return False, "Benutzername und Passwort dürfen nicht leer sein."

    if len(password) < 6:
        return False, "Das Passwort muss mindestens 6 Zeichen haben."

    if get_user(username):
        return False, "Dieser Benutzername ist bereits vergeben."

    salt = secrets.token_hex(16)

    with get_engine().begin() as conn:
        conn.execute(insert(users_table).values(
            username=username,
            password_hash=_hash_password(password, salt),
            salt=salt,
        ))

    return True, "Registrierung erfolgreich – du kannst dich jetzt einloggen."


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


def _cookie_manager():
    """Eine Cookie-Manager-Instanz pro Skriptlauf (oder None ohne Paket)."""
    if not COOKIES_AVAILABLE:
        return None

    try:
        return stx.CookieManager(key="surf_cookies")
    except Exception:
        return None


def login_session(user, remember):
    st.session_state["user"] = {"id": user["id"], "username": user["username"]}

    if remember:
        # Das Cookie wird bewusst erst NACH dem Rerun gesetzt (siehe Gate
        # weiter unten). Ein set() direkt vor st.rerun() wird vom Browser
        # häufig verworfen, sodass das Cookie nie gespeichert würde.
        st.session_state["_pending_token"] = create_auth_token(user["id"])


def logout_session(cookie_manager):
    if cookie_manager is not None:
        try:
            token = cookie_manager.get("surf_auth")
            if token:
                delete_auth_token(token)
            cookie_manager.delete("surf_auth", key="auth_cookie_del")
        except Exception:
            pass

    st.session_state.pop("user", None)


# =====================================================================
#  Gruppen
# =====================================================================

ALL_GROUP = "Alle"


def list_groups():
    with get_engine().connect() as conn:
        rows = conn.execute(select(groups_table)).mappings().all()

    return [dict(r) for r in rows]


def create_group(name, owner_id, is_private):
    name = (name or "").strip()

    if not name:
        return False, "Bitte einen Gruppennamen eingeben."

    if name.lower() == ALL_GROUP.lower():
        return False, "Dieser Name ist reserviert."

    with get_engine().begin() as conn:
        exists = conn.execute(
            select(groups_table.c.id).where(func.lower(groups_table.c.name) == name.lower())
        ).first()

        if exists:
            return False, "Es gibt bereits eine Gruppe mit diesem Namen."

        result = conn.execute(insert(groups_table).values(
            name=name, is_private=bool(is_private), owner_id=owner_id
        ))
        group_id = result.inserted_primary_key[0]

        conn.execute(insert(memberships_table).values(
            user_id=owner_id, group_id=group_id, status="member"
        ))

    return True, f"Gruppe „{name}“ wurde erstellt."


def my_memberships(user_id):
    """group_id -> status ('member' | 'pending')."""
    with get_engine().connect() as conn:
        rows = conn.execute(
            select(memberships_table).where(memberships_table.c.user_id == user_id)
        ).mappings().all()

    return {r["group_id"]: r["status"] for r in rows}


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
            return False, "Gruppe nicht gefunden."

        existing = conn.execute(
            select(memberships_table).where(
                (memberships_table.c.user_id == user_id)
                & (memberships_table.c.group_id == group_id)
            )
        ).first()

        if existing:
            return False, "Du bist dieser Gruppe bereits zugeordnet."

        status = "pending" if group["is_private"] else "member"

        conn.execute(insert(memberships_table).values(
            user_id=user_id, group_id=group_id, status=status
        ))

    if status == "pending":
        return True, "Beitritt angefragt – der Ersteller muss dich freischalten."

    return True, "Du bist der Gruppe beigetreten."


def leave_group(user_id, group_id):
    with get_engine().begin() as conn:
        conn.execute(delete(memberships_table).where(
            (memberships_table.c.user_id == user_id)
            & (memberships_table.c.group_id == group_id)
        ))


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


def invite_user(group_id, username):
    user = get_user((username or "").strip())

    if not user:
        return False, "Kein Nutzer mit diesem Benutzernamen gefunden."

    with get_engine().begin() as conn:
        existing = conn.execute(
            select(memberships_table).where(
                (memberships_table.c.user_id == user["id"])
                & (memberships_table.c.group_id == group_id)
            )
        ).mappings().first()

        if existing:
            if existing["status"] == "member":
                return False, f"{user['username']} ist bereits Mitglied."

            conn.execute(
                update(memberships_table)
                .where(
                    (memberships_table.c.user_id == user["id"])
                    & (memberships_table.c.group_id == group_id)
                )
                .values(status="member")
            )
            return True, f"{user['username']} wurde freigeschaltet."

        conn.execute(insert(memberships_table).values(
            user_id=user["id"], group_id=group_id, status="member"
        ))

    return True, f"{user['username']} wurde zur Gruppe hinzugefügt."


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


@st.cache_data(show_spinner=False)
def geocode_spot(name):
    if not name:
        return None

    params = {"name": name, "count": 1, "language": "de", "format": "json"}

    try:
        url = f"https://geocoding-api.open-meteo.com/v1/search?{urlencode(params)}"
        with urlopen(url, timeout=10) as response:
            data = json.load(response)
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

    endpoints = (
        "https://archive-api.open-meteo.com/v1/archive",
        "https://api.open-meteo.com/v1/forecast",
    )

    any_response = False

    for base in endpoints:
        try:
            with urlopen(f"{base}?{urlencode(common)}", timeout=25) as response:
                data = json.load(response)
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
        return None


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

    url = f"https://api.open-meteo.com/v1/forecast?{urlencode(params)}"
    # Großzügiger Timeout: Open-Meteo ist zeitweise langsam (>20 s), liefert
    # aber. Erfolgreiche Antworten werden 30 Min gecacht.
    with urlopen(url, timeout=30) as response:
        return json.load(response)


def get_forecast(lat, lon):
    """Aktuelles Wetter + Stundenvorhersage von Open-Meteo (Forecast-API).

    Fehler werden NICHT gecacht – nur erfolgreiche Antworten landen im Cache,
    damit eine kurze Dienststörung nicht 30 Minuten hängen bleibt.
    """
    try:
        return _fetch_forecast(lat, lon)
    except Exception:
        return None


def render_rankings():
    st.markdown("## 🏆 Online-Rankings")

    flash = st.session_state.pop("ranking_flash", None)

    if flash:
        st.success(flash)

    ranking = load_sessions()

    if ranking.empty:
        st.info("Noch keine Online-Ranking-Einträge vorhanden.")
        return

    if "date" not in ranking.columns:
        ranking["date"] = ""

    for column in ("wind_kmh", "wind_dir_deg", "temp_c", "weather_code"):
        if column not in ranking.columns:
            ranking[column] = None

    # ---- Filter: Gruppe (nur eigene Gruppen + "Alle") ----
    user = st.session_state.get("user")
    member_groups = my_member_groups(user["id"]) if user else []
    group_choice = st.selectbox(
        "👥 Gruppe",
        [ALL_GROUP] + [g["name"] for g in member_groups],
        key="rank_group",
        help="„Alle“ zeigt alle Fahrer. Gruppen-Ergebnisse siehst du nur als Mitglied.",
    )

    if group_choice != ALL_GROUP:
        group_id = next((g["id"] for g in member_groups if g["name"] == group_choice), None)

        if group_id is not None:
            member_names = set(group_member_names(group_id))
            ranking = ranking[ranking["name"].astype(str).isin(member_names)].copy()

            if ranking.empty:
                st.info(f"In der Gruppe „{group_choice}“ gibt es noch keine Sessions.")
                return

    # ---- Filter: Lokation + Datum (Jahr / Monat / Tag) ----
    ranking["_date"] = pd.to_datetime(ranking["date"], errors="coerce")

    spot_values = (
        ranking["surfspot"].dropna().astype(str)
        if "surfspot" in ranking.columns
        else pd.Series(dtype=str)
    )
    spots = sorted(s for s in spot_values.unique() if s.strip())

    months = [
        "Januar", "Februar", "März", "April", "Mai", "Juni",
        "Juli", "August", "September", "Oktober", "November", "Dezember",
    ]
    years = sorted(ranking["_date"].dropna().dt.year.unique(), reverse=True)

    f1, f2, f3, f4 = st.columns(4)

    with f1:
        spot_filter = st.selectbox("📍 Lokation", ["Gesamt"] + spots, key="rank_spot")

    with f2:
        year_filter = st.selectbox(
            "📅 Jahr",
            ["Alle Jahre"] + [str(y) for y in years],
            key="rank_year",
        )

    with f3:
        month_filter = st.selectbox(
            "Monat",
            ["Ganzes Jahr"] + months,
            key="rank_month",
        )

    with f4:
        day_filter = st.selectbox(
            "Tag",
            ["Ganzer Monat"] + [str(d) for d in range(1, 32)],
            key="rank_day",
        )

    if spot_filter != "Gesamt":
        ranking = ranking[ranking["surfspot"].astype(str) == spot_filter]

    if year_filter != "Alle Jahre":
        ranking = ranking[ranking["_date"].dt.year == int(year_filter)]

    if month_filter != "Ganzes Jahr":
        ranking = ranking[ranking["_date"].dt.month == months.index(month_filter) + 1]

        if day_filter != "Ganzer Monat":
            ranking = ranking[ranking["_date"].dt.day == int(day_filter)]

    ranking = ranking.copy()

    if ranking.empty:
        st.info("Keine Einträge für die gewählte Filter-Auswahl.")
        return

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

    ranking["Wetter"] = ranking.apply(weather_summary, axis=1)

    rcol1, rcol2 = st.columns(2)

    with rcol1:
        st.markdown("### 🏆 Beste 30 Sekunden")

        r30 = ranking[[
            "date",
            "name",
            "surfspot",
            "board",
            "sail",
            "Wetter",
            "speed_30s_kmh",
            "speed_30s_kn",
        ]].copy()

        r30 = r30.sort_values("speed_30s_kmh", ascending=False).reset_index(drop=True)
        r30.insert(0, "Platz", r30.index + 1)

        r30 = r30.rename(columns={
            "date": "Datum",
            "name": "Name",
            "surfspot": "Surfspot",
            "board": "Board",
            "sail": "Segel",
            "speed_30s_kmh": "30s km/h",
            "speed_30s_kn": "30s kn",
        })

        st.dataframe(
            r30,
            use_container_width=True,
            hide_index=True,
        )

    with rcol2:
        st.markdown("### ⚡ Topgeschwindigkeit 1 Sekunde")

        r1 = ranking[[
            "date",
            "name",
            "surfspot",
            "board",
            "sail",
            "Wetter",
            "speed_1s_kmh",
            "speed_1s_kn",
        ]].copy()

        r1 = r1.sort_values("speed_1s_kmh", ascending=False).reset_index(drop=True)
        r1.insert(0, "Platz", r1.index + 1)

        r1 = r1.rename(columns={
            "date": "Datum",
            "name": "Name",
            "surfspot": "Surfspot",
            "board": "Board",
            "sail": "Segel",
            "speed_1s_kmh": "1s km/h",
            "speed_1s_kn": "1s kn",
        })

        st.dataframe(
            r1,
            use_container_width=True,
            hide_index=True,
        )

    rcol3, rcol4 = st.columns(2)

    with rcol3:
        st.markdown("### 🚩 Längster Run")

        rrun = ranking[[
            "date",
            "name",
            "surfspot",
            "board",
            "sail",
            "Wetter",
            "longest_run_km",
            "longest_run_m",
        ]].copy()

        rrun = rrun.sort_values("longest_run_m", ascending=False).reset_index(drop=True)
        rrun.insert(0, "Platz", rrun.index + 1)

        rrun = rrun.rename(columns={
            "date": "Datum",
            "name": "Name",
            "surfspot": "Surfspot",
            "board": "Board",
            "sail": "Segel",
            "longest_run_km": "Run km",
            "longest_run_m": "Run m",
        })

        st.dataframe(
            rrun,
            use_container_width=True,
            hide_index=True,
        )

    with rcol4:
        st.markdown("### 👥 Längste Gesamtstrecke je Fahrer")

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

        rtotal.insert(0, "Platz", rtotal.index + 1)

        rtotal = rtotal.rename(columns={
            "name": "Name",
            "total_distance_km": "Gesamtstrecke km",
            "last_date": "Letzte Session",
        })

        st.dataframe(
            rtotal,
            use_container_width=True,
            hide_index=True,
        )


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
                    "Ende": run["timestamp"].iloc[-1],
                    "Dauer": run["timestamp"].iloc[-1] - run["timestamp"].iloc[0],
                    "Distanz m": distance_m,
                    "Distanz km": distance_m / 1000,
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
                "Ende": run["timestamp"].iloc[-1],
                "Dauer": run["timestamp"].iloc[-1] - run["timestamp"].iloc[0],
                "Distanz m": distance_m,
                "Distanz km": distance_m / 1000,
                "Max Speed km/h": run["speed_kmh"].max(),
                "Ø Speed km/h": run["speed_kmh"].mean(),
            })

    return pd.DataFrame(runs)


def show_map(df):
    if "lat" not in df.columns or "lon" not in df.columns:
        st.warning("Keine GPS-Daten gefunden.")
        return

    gps_df = df.dropna(subset=["lat", "lon"]).copy()

    if gps_df.empty:
        st.warning("Keine GPS-Daten gefunden.")
        return

    layer = pdk.Layer(
        "PathLayer",
        data=[{"path": gps_df[["lon", "lat"]].values.tolist()}],
        get_path="path",
        get_width=4,
        width_min_pixels=2,
    )

    view_state = pdk.ViewState(
        latitude=gps_df["lat"].mean(),
        longitude=gps_df["lon"].mean(),
        zoom=13,
    )

    st.pydeck_chart(
        pdk.Deck(
            map_style=None,
            initial_view_state=view_state,
            layers=[layer],
        ),
        use_container_width=True,
    )


def render_history_overview(record):
    """Session-Übersicht aus den gespeicherten Ranking-Werten (ohne Roh-FIT)."""

    def num(key):
        value = record.get(key)
        return None if value is None or pd.isna(value) else float(value)

    st.markdown("## 🌊 Session Übersicht")

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
    c3.metric("Gesamtstrecke", "–" if distance_km is None else f"{distance_km:.2f} km")
    c4.metric("Längster Run", "–" if longest_run_km is None else f"{longest_run_km:.2f} km")

    st.markdown("## 🌦️ Wetter zur Session")

    wind = num("wind_kmh")
    gust = num("gust_kmh")
    temp = num("temp_c")
    precip = num("precip_mm")
    wdir = num("wind_dir_deg")
    code = record.get("weather_code")
    has_code = code is not None and pd.notna(code)

    if all(v is None for v in (wind, gust, temp, precip)) and not has_code:
        st.info("Keine Wetterdaten zu dieser Session gespeichert.")
    else:
        emoji, description = (
            WEATHER_CODES.get(int(code), ("❓", "Unbekannt")) if has_code else ("❓", "Unbekannt")
        )

        w1, w2, w3, w4 = st.columns(4)

        w1.metric(
            "Wind",
            "–" if wind is None else f"{wind:.0f} km/h",
            None if gust is None else f"Böen {gust:.0f} km/h",
        )
        w2.metric("Temperatur", "–" if temp is None else f"{temp:.1f} °C")
        w3.metric("Niederschlag", "–" if precip is None else f"{precip:.1f} mm")
        w4.metric("Bedingung", emoji, description, delta_color="off")

        if wdir is not None:
            st.caption(
                f"🧭 Wind aus **{degrees_to_compass(wdir)}** "
                f"{wind_arrow(wdir)} ({wdir:.0f}°)"
            )

    st.markdown("## ⚡ Speed-Werte")

    speed_table = pd.DataFrame([
        {
            "Wertung": "1 Sekunde",
            "Speed km/h": best_1s,
            "Speed kn": None if best_1s is None else best_1s / 1.852,
        },
        {
            "Wertung": "30 Sekunden",
            "Speed km/h": best_30s,
            "Speed kn": None if best_30s is None else best_30s / 1.852,
        },
    ])

    st.dataframe(
        speed_table.round(2),
        use_container_width=True,
        hide_index=True,
    )

    st.caption(
        "ℹ️ Karte, einzelne Runs sowie Max-/Ø-Speed sind nur direkt nach dem "
        "Upload verfügbar – für gespeicherte Sessions werden die abgelegten "
        "Kennzahlen angezeigt."
    )


load_css(app_path("assets", "style.css"))

logo_img = image_to_base64(app_path("assets", "windsurfer.png"))

# Vollflächiges Hintergrundbild (Wasser/Surfer). Tausche assets/background.jpg
# gegen dein Wunschfoto – Fallback ist header.jpg.
bg_img = (
    image_to_base64(app_path("assets", "background.jpg"))
    or image_to_base64(app_path("assets", "header.jpg"))
)

if bg_img:
    st.markdown(
        f"""
<style>
.stApp {{
    background: linear-gradient(rgba(2,22,43,.45), rgba(2,22,43,.62)),
                url("data:image/jpeg;base64,{bg_img}") center center / cover fixed no-repeat;
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
        <p>Tracke deine Sessions. Vergleiche deine Bestleistungen.<br>
        Die Community. Dein Spot. Dein Speed.</p>
        <div class="hero-nav">
            <span>⚡ Speed</span>
            <span>🏆 Ranking</span>
            <span>📍 Spots</span>
            <span>👥 Community</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)


# =====================================================================
#  Login / Registrierung (Gate vor dem Rest der App)
# =====================================================================

def render_login(cookie_manager):
    st.markdown("## 🔐 Anmeldung")
    st.info(
        "Bitte einloggen oder registrieren. Danach kannst du Gruppen "
        "beitreten oder eigene anlegen."
    )

    tab_login, tab_register = st.tabs(["Einloggen", "Registrieren"])

    with tab_login:
        with st.form("login_form"):
            username = st.text_input("Benutzername")
            password = st.text_input("Passwort", type="password")
            remember = st.checkbox("Angemeldet bleiben", value=True)
            submitted = st.form_submit_button("Einloggen")

        if submitted:
            if verify_login(username, password):
                user = get_user(username.strip())
                login_session(user, remember)
                st.rerun()
            else:
                st.error("Benutzername oder Passwort falsch.")

    with tab_register:
        with st.form("register_form"):
            new_username = st.text_input("Benutzername wählen")
            pwd1 = st.text_input("Passwort (mind. 6 Zeichen)", type="password")
            pwd2 = st.text_input("Passwort wiederholen", type="password")
            submitted = st.form_submit_button("Registrieren")

        if submitted:
            if pwd1 != pwd2:
                st.error("Die Passwörter stimmen nicht überein.")
            else:
                ok, message = register_user(new_username, pwd1)

                if ok:
                    st.success(message)
                else:
                    st.error(message)


def render_account_sidebar(user, cookie_manager):
    with st.sidebar:
        st.markdown(f"### 👤 {user['username']}")

        if st.button("Abmelden", use_container_width=True):
            logout_session(cookie_manager)
            st.rerun()

        st.markdown("---")
        st.markdown("### 👥 Gruppen")
        st.caption("Du bist immer Teil der Gruppe **Alle** (alle Ergebnisse sichtbar).")

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
            st.markdown("**Deine Gruppen**")
            for g in member_groups:
                tag = "🔒" if g["is_private"] else "🌍"
                is_owner = g["owner_id"] == user["id"]
                cols = st.columns([4, 2])
                cols[0].write(f"{tag} {g['name']}" + (" · Ersteller" if is_owner else ""))
                if not is_owner:
                    if cols[1].button("Verlassen", key=f"leave_{g['id']}"):
                        leave_group(user["id"], g["id"])
                        st.rerun()

        if pending_groups:
            st.markdown("**Angefragt (wartet auf Freigabe)**")
            for g in pending_groups:
                st.write(f"⏳ {g['name']}")

        joinable = [g for g in groups if g["id"] not in memberships]

        if joinable:
            with st.expander("➕ Gruppe beitreten"):
                option_map = {
                    f'{"🔒" if g["is_private"] else "🌍"} {g["name"]}': g
                    for g in joinable
                }
                choice = st.selectbox("Gruppe", list(option_map.keys()), key="join_select")
                target = option_map[choice]
                btn_label = "Beitritt anfragen" if target["is_private"] else "Beitreten"

                if st.button(btn_label, key="join_btn", use_container_width=True):
                    ok, message = join_or_request_group(user["id"], target["id"])
                    (st.success if ok else st.warning)(message)
                    if ok:
                        st.rerun()

        with st.expander("🆕 Gruppe erstellen"):
            grp_name = st.text_input("Gruppenname", key="new_group_name")
            grp_type = st.radio(
                "Typ",
                ["Offen (jeder darf beitreten)", "Privat (nur auf Einladung)"],
                key="new_group_type",
            )

            if st.button("Erstellen", key="new_group_btn", use_container_width=True):
                ok, message = create_group(grp_name, user["id"], grp_type.startswith("Privat"))
                (st.success if ok else st.warning)(message)
                if ok:
                    st.rerun()

        owned_groups = [g for g in member_groups if g["owner_id"] == user["id"]]

        if owned_groups:
            with st.expander("🛠️ Meine Gruppen verwalten"):
                for g in owned_groups:
                    tag = "🔒" if g["is_private"] else "🌍"
                    st.markdown(f"**{tag} {g['name']}**")

                    requests = pending_requests(g["id"])

                    if requests:
                        st.caption("Offene Beitrittsanfragen:")
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
                        st.caption("Keine offenen Anfragen.")

                    invite_name = st.text_input(
                        "Nutzer einladen / freischalten (Benutzername)",
                        key=f"invite_{g['id']}",
                    )
                    if st.button("Hinzufügen", key=f"invite_btn_{g['id']}"):
                        ok, message = invite_user(g["id"], invite_name)
                        (st.success if ok else st.warning)(message)
                        if ok:
                            st.rerun()
                    st.markdown("---")


cookie_manager = _cookie_manager()

if cookie_manager is not None:
    # Unsichtbares Cookie-Iframe ausblenden (verhindert eine Layout-Lücke)
    st.markdown(
        '<style>.element-container:has(iframe[height="0"]) { display:none; }</style>',
        unsafe_allow_html=True,
    )

# Login aus „Angemeldet bleiben"-Cookie wiederherstellen
if "user" not in st.session_state and cookie_manager is not None:
    try:
        saved_token = cookie_manager.get("surf_auth")
        if saved_token:
            restored = user_for_token(saved_token)
            if restored:
                st.session_state["user"] = {
                    "id": restored["id"],
                    "username": restored["username"],
                }
    except Exception:
        pass

current_user = st.session_state.get("user")

if not current_user:
    render_login(cookie_manager)
    st.stop()

render_account_sidebar(current_user, cookie_manager)

# „Angemeldet bleiben"-Cookie erst hier setzen – in einem Lauf, der ganz
# durchläuft, damit der Browser es zuverlässig speichert.
if cookie_manager is not None and st.session_state.get("_pending_token"):
    try:
        cookie_manager.set(
            "surf_auth", st.session_state["_pending_token"],
            expires_at=datetime.now() + timedelta(days=30),
            key="auth_cookie_set",
        )
    except Exception:
        pass
    st.session_state.pop("_pending_token", None)


render_rankings()

st.markdown("---")


left, right = st.columns([1, 2])

selected_history_record = None


with left:
    st.markdown("### 👤 1. Session & Material")

    name = current_user["username"]
    st.markdown(f"**Fahrer:** `{name}`")

    profiles = load_profiles()
    rider = profiles.get(name, {})

    if name:
        with st.expander("📅 Meine letzten Sessions", expanded=False):
            history = load_rider_sessions(name)

            if history.empty:
                st.info("Noch keine gespeicherten Sessions für diesen Fahrer.")
            else:
                valid_dates = history["date"].dropna() if "date" in history.columns else pd.Series(dtype="datetime64[ns]")

                if not valid_dates.empty:
                    min_date = valid_dates.min().date()
                    max_date = valid_dates.max().date()

                    date_range = st.date_input(
                        "Zeitraum filtern",
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

                history = history.head(10)

                if history.empty:
                    st.info("Keine Sessions im gewählten Zeitraum.")
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
                        "date": "Datum",
                        "surfspot": "Surfspot",
                        "board": "Board",
                        "sail": "Segel",
                        "speed_30s_kmh": "30s km/h",
                        "speed_1s_kmh": "1s km/h",
                        "longest_run_km": "Run km",
                        "total_distance_km": "Strecke km",
                    })

                    st.dataframe(
                        show_history,
                        use_container_width=True,
                        hide_index=True,
                    )

                    record_by_label = {}

                    for i, (_, row) in enumerate(history.iterrows()):
                        if "date" in history.columns and pd.notna(row["date"]):
                            date_str = row["date"].strftime("%Y-%m-%d")
                        else:
                            date_str = "?"

                        spot_label = str(row.get("surfspot", "") or "")
                        label = f"{date_str} · {spot_label}".strip(" ·") or f"Session {i + 1}"

                        if label in record_by_label:
                            label = f"{label} ({i + 1})"

                        record_by_label[label] = row

                    chosen_label = st.selectbox(
                        "Session-Details rechts anzeigen",
                        ["—"] + list(record_by_label.keys()),
                        key=f"history_pick_{name}",
                    )

                    if chosen_label != "—":
                        selected_history_record = record_by_label[chosen_label]

    spot_options = rider.get("spots", [])
    spot_choice = st.selectbox("Surfspot", [NEW_ENTRY] + spot_options)

    if spot_choice == NEW_ENTRY:
        spot = st.text_input("Neuer Surfspot")
    else:
        spot = spot_choice

    st.markdown("**Board**")
    board_options = rider.get("boards", [])
    board_choice = st.selectbox("Board auswählen", [NEW_ENTRY] + board_options)

    if board_choice == NEW_ENTRY:
        board_brand = st.text_input("Board-Marke")
        board_model = st.text_input("Board-Typ / Modell")
        board_volume = st.number_input("Volumen in Liter", min_value=0, step=1)
        board_display = f"{board_brand.strip()} {board_model.strip()} {board_volume}L".strip()
        board_ok = bool(board_brand.strip() and board_model.strip() and board_volume > 0)
    else:
        board_display = board_choice
        board_ok = True

    st.markdown("**Segel**")
    sail_options = rider.get("sails", [])
    sail_choice = st.selectbox("Segel auswählen", [NEW_ENTRY] + sail_options)

    if sail_choice == NEW_ENTRY:
        sail_brand = st.text_input("Segelhersteller")
        sail_model = st.text_input("Segelname / Modell")
        sail_size = st.number_input("Segelgröße in m²", min_value=0.0, step=0.1)
        sail_display = f"{sail_brand.strip()} {sail_model.strip()} {sail_size:.1f} m²".strip()
        sail_ok = bool(sail_brand.strip() and sail_model.strip() and sail_size > 0)
    else:
        sail_display = sail_choice
        sail_ok = True

    st.markdown("### ☁️ 2. Aktivität laden")

    fit_source = None
    fit_name = None

    # Uhr-Auslesen (USB/MTP) klappt nur lokal unter Windows – auf einem
    # Server gibt es ausschließlich den Datei-Upload.
    if IS_WINDOWS:
        source = st.radio(
            "Quelle",
            ["📁 Datei hochladen", "⌚ Von Uhr (USB)"],
            horizontal=True,
        )
    else:
        source = "📁 Datei hochladen"

    if source == "📁 Datei hochladen":
        uploaded_file = st.file_uploader("FIT-Datei hochladen", type=["fit"])

        if uploaded_file is not None:
            fit_source = uploaded_file
            fit_name = uploaded_file.name
    else:
        st.caption(
            "Uhr per USB anschließen. Egal ob als Laufwerk (z.B. ältere Edge) "
            "oder als Gerät ohne Laufwerksbuchstaben (z.B. Fenix 6 Pro) – die "
            "Aktivitäten erscheinen automatisch. Optional Ordnerpfad angeben."
        )

        manual_folder = st.text_input(
            "Optional: Ordnerpfad (z.B. E:\\GARMIN\\Activity)"
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
                "Keine FIT-Dateien gefunden. Uhr anschließen oder Ordnerpfad angeben."
            )
        else:
            labels = {}

            for i, activity in enumerate(activities):
                label = activity["label"]

                if label in labels:
                    label = f"{label} ({i + 1})"

                labels[label] = activity

            choice = st.selectbox(
                "Aktivität auswählen (neueste zuerst)",
                list(labels.keys()),
            )

            selected = labels[choice]

            if selected["kind"] == "local":
                fit_source = selected["path"]
                fit_name = selected["name"]
            else:
                with st.spinner("Datei wird von der Uhr kopiert …"):
                    copied = copy_mtp_file(selected["device"], selected["name"])

                if copied:
                    fit_source = copied
                    fit_name = f"{selected['name']}.fit"
                else:
                    st.error(
                        "Die Datei konnte nicht von der Uhr kopiert werden. "
                        "Uhr neu verbinden und erneut versuchen."
                    )


required_ok = all([
    spot.strip(),
    board_ok,
    sail_ok,
])


if fit_source is not None:
    if not required_ok:
        st.warning("Bitte zuerst Surfspot, Board und Segel vollständig eingeben.")

    df = read_fit_file(fit_source)

    if df.empty:
        st.error("Die Datei enthält keine Record-Daten.")
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
        longest_run_m = runs_df["Distanz m"].max()
        longest_run_km = longest_run_m / 1000

    with right:
        st.markdown("## 🌊 Session Übersicht")

        c1, c2, c3, c4 = st.columns(4)

        c1.metric(
            "Max Speed",
            "Keine Daten" if max_speed is None else f"{max_speed:.2f} km/h"
        )

        c2.metric(
            "Ø Speed",
            "Keine Daten" if avg_speed is None else f"{avg_speed:.2f} km/h"
        )

        c3.metric(
            "Gesamtstrecke",
            "Keine Daten" if distance_km is None else f"{distance_km:.2f} km"
        )

        c4.metric(
            "Längster Run",
            "Keine Daten" if longest_run_km is None else f"{longest_run_km:.2f} km"
        )

        st.markdown("## 🌦️ Wetter zur Session")

        if weather is None:
            if session_lat is None:
                st.info("Keine GPS-Daten – Wetter kann nicht abgerufen werden.")
            else:
                st.info("Keine Wetterdaten verfügbar (evtl. noch kein Archiv-Eintrag).")
        else:
            emoji, description = WEATHER_CODES.get(weather["code"], ("❓", "Unbekannt"))

            w1, w2, w3, w4 = st.columns(4)

            w1.metric(
                "Wind",
                "–" if weather["wind"] is None else f"{weather['wind']:.0f} km/h",
                None if weather["gust"] is None else f"Böen {weather['gust']:.0f} km/h",
            )

            w2.metric(
                "Temperatur",
                "–" if weather["temp"] is None else f"{weather['temp']:.1f} °C",
            )

            w3.metric(
                "Niederschlag",
                "–" if weather["precip"] is None else f"{weather['precip']:.1f} mm",
            )

            w4.metric(
                "Bedingung",
                emoji,
                description,
                delta_color="off",
            )

            if weather["dir"] is not None:
                st.caption(
                    f"🧭 Wind aus **{degrees_to_compass(weather['dir'])}** "
                    f"{wind_arrow(weather['dir'])} ({weather['dir']:.0f}°)"
                )

        st.markdown("## ⚡ Speed-Werte")

        speed_table = pd.DataFrame([
            {
                "Wertung": "1 Sekunde",
                "Speed km/h": best_1s,
                "Speed kn": None if best_1s is None else best_1s / 1.852,
            },
            {
                "Wertung": "30 Sekunden",
                "Speed km/h": best_30s,
                "Speed kn": None if best_30s is None else best_30s / 1.852,
            },
        ])

        st.dataframe(
            speed_table.round(2),
            use_container_width=True,
            hide_index=True,
        )

        with st.expander("🌊 Erkannte Runs / Einzelstrecken", expanded=False):
            if not runs_df.empty:
                show_runs = runs_df.copy()
                show_runs["Distanz m"] = show_runs["Distanz m"].round(2)
                show_runs["Distanz km"] = show_runs["Distanz km"].round(3)
                show_runs["Max Speed km/h"] = show_runs["Max Speed km/h"].round(2)
                show_runs["Ø Speed km/h"] = show_runs["Ø Speed km/h"].round(2)
                show_runs = show_runs.sort_values("Distanz m", ascending=False).reset_index(drop=True)
                show_runs.insert(0, "Run", show_runs.index + 1)

                st.dataframe(
                    show_runs,
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("Keine Runs erkannt. Eventuell Grenzwerte anpassen.")

        if st.session_state.get("just_added") == fit_name:
            st.success(
                f"✅ Session wurde zum Online-Ranking hinzugefügt: **{fit_name}**."
            )
        elif session_exists(fit_name):
            st.info(
                f"⚠️ Diese Datei wurde bereits hochgeladen: **{fit_name}**. "
                "Sie kann kein zweites Mal zum Ranking hinzugefügt werden."
            )
        elif required_ok and best_30s is not None:
            if st.button("🏆 Session zum Online-Ranking hinzufügen"):
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
                }

                save_session(entry)
                update_profile(name.strip(), spot.strip(), board_display, sail_display)
                update_spot_coords(spot.strip(), session_lat, session_lon)
                st.session_state["ranking_flash"] = (
                    "Session wurde im Online-Ranking gespeichert."
                )
                st.session_state["just_added"] = fit_name
                st.rerun()

    st.markdown("---")

    with st.expander("📍 Track auf Karte", expanded=False):
        show_map(df)

    with st.expander("📋 Rohdaten", expanded=False):
        st.dataframe(
            df.head(100),
            use_container_width=True,
        )

        csv = df.to_csv(index=False).encode("utf-8")

        st.download_button(
            "⬇️ Rohdaten als CSV herunterladen",
            data=csv,
            file_name="fit_export.csv",
            mime="text/csv",
        )

else:
    with right:
        if selected_history_record is not None:
            render_history_overview(selected_history_record)
        else:
            st.info("Bitte links Teilnehmerdaten eingeben und eine FIT-Datei hochladen.")
            st.info("Nach dem Upload erscheint hier die Auswertung.")


st.markdown("---")

with st.expander("🌦️ Spot-Wetter (aktuell & Vorhersage)", expanded=False):
    spot_list = all_known_spots()

    if not spot_list:
        st.info("Noch keine Spots hinterlegt. Speichere zuerst eine Session mit Surfspot.")
    else:
        selected_spot = st.selectbox("Spot auswählen", spot_list, key="weather_spot")

        spot_lat, spot_lon = resolve_spot_coords(selected_spot)

        if spot_lat is None:
            st.warning(
                f"Für „{selected_spot}\" konnten keine Koordinaten ermittelt werden. "
                "Speichere eine Session mit GPS an diesem Spot oder benenne ihn eindeutiger."
            )
        else:
            forecast = get_forecast(spot_lat, spot_lon)

            if not forecast:
                st.warning(
                    "Wetterdienst (Open-Meteo) gerade nicht erreichbar – bitte "
                    "später erneut versuchen. Das liegt am Dienst, nicht an "
                    "deinen Daten."
                )
            else:
                current = forecast.get("current", {})
                code = current.get("weather_code")
                emoji, description = WEATHER_CODES.get(code, ("❓", "Unbekannt"))

                st.markdown(f"**Aktuell** &nbsp; 📍 {spot_lat:.3f}, {spot_lon:.3f}")

                m1, m2, m3, m4 = st.columns(4)

                wind_now = current.get("wind_speed_10m")
                gust_now = current.get("wind_gusts_10m")
                temp_now = current.get("temperature_2m")
                rain_now = current.get("precipitation") or 0

                m1.metric(
                    "Wind",
                    "–" if wind_now is None else f"{wind_now:.0f} km/h",
                    None if gust_now is None else f"Böen {gust_now:.0f} km/h",
                )

                m2.metric(
                    "Temperatur",
                    "–" if temp_now is None else f"{temp_now:.1f} °C",
                )

                m3.metric(
                    "Regen",
                    "Ja" if rain_now > 0 else "Nein",
                    None if rain_now <= 0 else f"{rain_now:.1f} mm",
                    delta_color="off",
                )

                m4.metric(
                    "Bedingung",
                    emoji,
                    description,
                    delta_color="off",
                )

                wind_dir_now = current.get("wind_direction_10m")

                if wind_dir_now is not None:
                    st.caption(
                        f"🧭 Wind aus **{degrees_to_compass(wind_dir_now)}** "
                        f"{wind_arrow(wind_dir_now)} ({wind_dir_now:.0f}°)"
                    )

                st.markdown("**Vorhersage (3 Tage, tagsüber 09–21 Uhr)**")

                hourly = forecast.get("hourly", {})
                times = hourly.get("time", [])
                now = current.get("time", "")

                start = 0

                for i, t in enumerate(times):
                    if t >= now:
                        start = i
                        break

                daytime_hours = {9, 12, 15, 18, 21}
                rows = []
                seen_dates = []

                for i in range(start, len(times)):
                    if int(times[i][11:13]) not in daytime_hours:
                        continue

                    day = times[i][:10]

                    if day not in seen_dates:
                        if len(seen_dates) >= 3:
                            break

                        seen_dates.append(day)

                    code_i = hourly["weather_code"][i]
                    emoji_i = WEATHER_CODES.get(code_i, ("", ""))[0]
                    precip_i = hourly["precipitation"][i] or 0
                    dir_i = hourly["wind_direction_10m"][i]
                    wind_i = hourly["wind_speed_10m"][i]
                    gust_i = hourly["wind_gusts_10m"][i]
                    weekday = WEEKDAYS[datetime.strptime(times[i][:10], "%Y-%m-%d").weekday()]

                    rows.append({
                        "Tag": weekday,
                        "Datum": times[i][5:10],
                        "Uhrzeit": times[i][11:16],
                        "Wetter": emoji_i,
                        "Temp °C": round(hourly["temperature_2m"][i], 1),
                        "Wind km/h": round(wind_i),
                        "Bft": kmh_to_beaufort(wind_i),
                        "Böen km/h": round(gust_i),
                        "Böen Bft": kmh_to_beaufort(gust_i),
                        "Richtung": degrees_to_compass(dir_i),
                        "Regen": "Ja" if precip_i > 0 else "Nein",
                    })

                st.dataframe(
                    pd.DataFrame(rows),
                    use_container_width=True,
                    hide_index=True,
                )


st.markdown("---")

st.markdown(f"""
<div class="footer">
    <h3 style="color:white;">{logo_icon} WINDSURF SPEED CHALLENGE</h3>
    <p>Die Community. Dein Spot. Dein Speed.</p>
</div>
""", unsafe_allow_html=True)