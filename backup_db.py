"""
Einfaches, abhängigkeitsarmes Backup der MyWaterSessions-Datenbank.

Sichert ALLE Tabellen der per DATABASE_URL angegebenen (Neon-)Datenbank als
CSV-Dateien in einen Zeitstempel-Ordner. Binärspalten (Logos/Spot-Bilder) werden
base64-kodiert mitgesichert, sind also vollständig wiederherstellbar.

Nutzt nur SQLAlchemy + pandas (sind durch die App ohnehin installiert) – kein
pg_dump / keine Postgres-Client-Tools nötig.

AUSFÜHREN (Windows PowerShell), im Ordner dieses Skripts:
    $env:DATABASE_URL = "postgresql+psycopg2://USER:PASS@HOST/DB?sslmode=require"
    python backup_db.py

(Die DATABASE_URL ist genau der String aus den Render-Env-Variablen des
App-Service. NICHT ins Git/öffentlich – nur lokal als Umgebungsvariable setzen.)

Ergebnis: Ordner  db-backup-YYYYMMDD-HHMMSS/  mit je einer CSV pro Tabelle
+ _manifest.txt (Tabellen + Zeilenzahl). Den Ordner danach an einen sicheren
Ort kopieren (Cloud-Speicher / externe Platte).
"""

import base64
import datetime
import os
import sys

import pandas as pd
from sqlalchemy import create_engine, inspect


def _b64_if_bytes(v):
    if isinstance(v, (bytes, bytearray, memoryview)):
        return base64.b64encode(bytes(v)).decode("ascii")
    return v


def main():
    url = os.environ.get("DATABASE_URL")
    if not url:
        print("FEHLER: Umgebungsvariable DATABASE_URL ist nicht gesetzt.")
        print('  PowerShell:  $env:DATABASE_URL = "postgresql+psycopg2://…"')
        sys.exit(1)

    # SQLAlchemy versteht 'postgresql://' und 'postgresql+psycopg2://'.
    engine = create_engine(url)
    stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    out_dir = f"db-backup-{stamp}"
    os.makedirs(out_dir, exist_ok=True)

    tables = inspect(engine).get_table_names()
    if not tables:
        print("Keine Tabellen gefunden – stimmt die DATABASE_URL?")
        sys.exit(1)

    manifest = [f"MyWaterSessions DB-Backup · {stamp}", f"{len(tables)} Tabellen", ""]
    total_rows = 0
    for t in sorted(tables):
        try:
            df = pd.read_sql_table(t, engine)
        except Exception as exc:  # noqa: BLE001
            print(f"  ! {t}: übersprungen ({exc})")
            manifest.append(f"{t}: FEHLER ({exc})")
            continue
        # Binärspalten base64-kodieren, damit sie sauber in CSV passen.
        for col in df.columns:
            if df[col].map(lambda v: isinstance(v, (bytes, bytearray, memoryview))).any():
                df[col] = df[col].map(_b64_if_bytes)
        df.to_csv(os.path.join(out_dir, f"{t}.csv"), index=False, encoding="utf-8")
        total_rows += len(df)
        manifest.append(f"{t}: {len(df)} Zeilen")
        print(f"  ✓ {t}: {len(df)} Zeilen")

    with open(os.path.join(out_dir, "_manifest.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(manifest) + "\n")

    print(f"\nFertig: {len(tables)} Tabellen, {total_rows} Zeilen -> {out_dir}\\")
    print("Diesen Ordner an einen sicheren Ort kopieren (Cloud / externe Platte).")


if __name__ == "__main__":
    main()
