# Windsurf Speed Challenge

Streamlit-App zum Auswerten von Garmin-FIT-Dateien (Speed, Runs, Wetter) mit
Login, Gruppen und Online-Ranking.

## Lokal starten (Windows)

```powershell
pip install -r requirements.txt
streamlit run "Gemini Fenix 6 App.py"
```

- Ohne weitere Konfiguration nutzt die App eine lokale SQLite-Datei
  (`surfapp.db`), die automatisch angelegt wird.
- Vorhandene Altdaten (`ranking.csv`, `profiles.json`, `spots.json`) werden beim
  ersten Start einmalig in die Datenbank übernommen.
- „⌚ Von Uhr (USB)" (inkl. Fenix 6 Pro per MTP) funktioniert nur lokal unter
  Windows.

## Deployment (Streamlit Community Cloud)

> ⚠️ **Wichtig:** Das Dateisystem der kostenlosen Cloud ist *flüchtig*. Eine
> lokale SQLite-Datei wäre bei jedem Neustart weg. Für den Server **muss** eine
> dauerhafte Datenbank per `DATABASE_URL` gesetzt werden.

### 1. Kostenlose Postgres anlegen
Z. B. bei [Neon](https://neon.tech) oder [Supabase](https://supabase.com).
Dort den Verbindungs-String („Connection string", Format `postgresql://...`)
kopieren.

### 2. Repo zu GitHub pushen
Diese Dateien müssen mit hochgeladen werden:
- `Gemini Fenix 6 App.py`
- `requirements.txt`
- der Ordner `assets/` (Bilder + `style.css`)

**Nicht** committen: `surfapp.db`, `*.json`, `*.csv`, `secrets.toml`
(siehe `.gitignore`).

### 3. App auf share.streamlit.io anlegen
- Repository auswählen, Main file: `Gemini Fenix 6 App.py`
- Unter **Settings → Secrets** eintragen (psycopg2-Treiber davorsetzen):

```toml
DATABASE_URL = "postgresql+psycopg2://USER:PASSWORT@HOST:5432/DBNAME"
```

Fertig – Streamlit Cloud liefert automatisch HTTPS. Beim ersten Start legt die
App alle Tabellen selbst an.

## Eigener Server / VPS
Mit dauerhafter Festplatte reicht die SQLite-Standardeinstellung – kein
`DATABASE_URL` nötig. **Wichtig:** unbedingt einen Reverse-Proxy mit TLS/HTTPS
davorschalten, damit Passwörter nicht im Klartext übertragen werden.

## Bekannte Beta-Einschränkungen
- Kein Passwort-Reset (vergessenes Passwort muss in der DB zurückgesetzt werden).
- Gruppe „Alle" zeigt allen registrierten Nutzern alle Ergebnisse (so gewollt).
