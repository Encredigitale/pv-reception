from pathlib import Path
import sqlite3

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "echaff.db"
SCHEMA_PATH = BASE_DIR / "echaff_sqlite_schema.sql"


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)

    if not SCHEMA_PATH.exists():
        raise FileNotFoundError(f"Fichier SQL introuvable : {SCHEMA_PATH}")

    with get_db() as conn:
        with open(SCHEMA_PATH, "r", encoding="utf-8") as f:
            conn.executescript(f.read())
        conn.commit()

    print(f"[OK] Base SQLite initialisée : {DB_PATH}")


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


# =========================================================
# VERIFICATEURS
# Compatibilité avec main.py
# =========================================================

def insert_verificateur(
    nom,
    prenom,
    email,
    telephone,
    numero_diplome,
    date_obtention_diplome,
    date_echeance_diplome,
    fichier_carte_recto,
    fichier_carte_verso,
    fichier_diplome
):
    query = """
        INSERT INTO verificateurs (
            nom,
            prenom,
            email,
            telephone,
            numero_diplome,
            date_obtention_diplome,
            date_echeance_diplome,
            fichier_carte_recto,
            fichier_carte_verso,
            fichier_diplome,
            actif
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
    """

    return execute(query, (
        nom,
        prenom,
        email,
        telephone,
        numero_diplome,
        date_obtention_diplome,
        date_echeance_diplome,
        fichier_carte_recto,
        fichier_carte_verso,
        fichier_diplome
    ))


def get_all_verificateurs():
    query = """
        SELECT *
        FROM verificateurs
        ORDER BY nom ASC, prenom ASC
    """
    return fetch_all(query)


def search_verificateurs(q=""):
    q = f"%{q}%"

    query = """
        SELECT *
        FROM verificateurs
        WHERE actif = 1
        AND (
            nom LIKE ?
            OR prenom LIKE ?
            OR email LIKE ?
            OR numero_diplome LIKE ?
        )
        ORDER BY nom ASC, prenom ASC
        LIMIT 20
    """

    return fetch_all(query, (q, q, q, q))
