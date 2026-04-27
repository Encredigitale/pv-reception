from pathlib import Path
from datetime import datetime
from uuid import uuid4
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

# =========================================================
# VERIFICATEURS
# =========================================================

def insert_verificateur(
    nom: str,
    prenom: str,
    email: str,
    telephone: str = "",
    numero_diplome: str = "",
    date_obtention_diplome: str = "",
    date_echeance_diplome: str = "",
    fichier_carte_recto: str = "",
    fichier_carte_verso: str = "",
    fichier_diplome: str = "",
) -> str:
    """
    Insère un nouveau vérificateur dans la base et retourne son id.
    """
    now = datetime.now().isoformat()
    verificateur_id = uuid4().hex

    execute(
        """
        INSERT INTO verificateurs (
            id, nom, prenom, email, telephone,
            numero_diplome, date_obtention_diplome, date_echeance_diplome,
            fichier_carte_recto, fichier_carte_verso, fichier_diplome,
            created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            verificateur_id, nom, prenom, email, telephone,
            numero_diplome, date_obtention_diplome, date_echeance_diplome,
            fichier_carte_recto, fichier_carte_verso, fichier_diplome,
            now, now,
        ),
    )

    return verificateur_id


def get_all_verificateurs() -> list[dict]:
    """
    Retourne tous les vérificateurs triés par nom puis prénom.
    """
    return fetch_all(
        "SELECT * FROM verificateurs ORDER BY nom ASC, prenom ASC"
    )


def search_verificateurs(query: str = "") -> list[dict]:
    """
    Recherche des vérificateurs par nom, prénom, email ou numéro de diplôme.
    Retourne tous les vérificateurs si la requête est vide.
    """
    if not query or not query.strip():
        return get_all_verificateurs()

    pattern = f"%{query.strip()}%"
    return fetch_all(
        """
        SELECT * FROM verificateurs
        WHERE nom LIKE ?
           OR prenom LIKE ?
           OR email LIKE ?
           OR numero_diplome LIKE ?
        ORDER BY nom ASC, prenom ASC
        """,
        (pattern, pattern, pattern, pattern),
    )
