from pathlib import Path
import sqlite3

# =========================================================
# CONFIG
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "echaff.db"
SCHEMA_PATH = BASE_DIR / "echaff_sqlite_schema.sql"

# =========================================================
# CONNEXION
# =========================================================

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

# =========================================================
# INITIALISATION BASE
# =========================================================

def init_db():
    """
    Initialise la base SQLite à partir du fichier SQL.
    """
    # Créer le dossier si besoin
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)

    if not SCHEMA_PATH.exists():
        raise FileNotFoundError(f"Fichier SQL introuvable : {SCHEMA_PATH}")

    with get_db() as conn:
        with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
            sql_script = f.read()

        conn.executescript(sql_script)
        conn.commit()

    print(f"[OK] Base SQLite initialisée : {DB_PATH}")

# =========================================================
# HELPERS GENERIQUES
# =========================================================

def fetch_all(query, params=()):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute(query, params)
        return [dict(row) for row in cur.fetchall()]

def fetch_one(query, params=()):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute(query, params)
        row = cur.fetchone()
        return dict(row) if row else None

def execute(query, params=()):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute(query, params)
        conn.commit()
        return cur.lastrowid
