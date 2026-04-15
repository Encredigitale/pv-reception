import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "pv_reception.db"


def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS verificateurs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nom TEXT NOT NULL,
            prenom TEXT NOT NULL,
            email TEXT NOT NULL,
            telephone TEXT,
            numero_diplome TEXT NOT NULL,
            date_obtention_diplome TEXT,
            date_echeance_diplome TEXT,
            fichier_carte_recto TEXT,
            fichier_carte_verso TEXT,
            fichier_diplome TEXT,
            actif INTEGER NOT NULL DEFAULT 1,
            date_creation TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("PRAGMA table_info(verificateurs)")
    existing_columns = [row["name"] for row in cursor.fetchall()]

    if "date_obtention_diplome" not in existing_columns:
        cursor.execute("ALTER TABLE verificateurs ADD COLUMN date_obtention_diplome TEXT")

    if "date_echeance_diplome" not in existing_columns:
        cursor.execute("ALTER TABLE verificateurs ADD COLUMN date_echeance_diplome TEXT")

    conn.commit()
    conn.close()


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
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
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
            fichier_diplome
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
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

    conn.commit()
    verificateur_id = cursor.lastrowid
    conn.close()
    return verificateur_id


def get_all_verificateurs():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT *
        FROM verificateurs
        ORDER BY nom ASC, prenom ASC
    """)

    rows = cursor.fetchall()
    conn.close()
    return rows


def search_verificateurs(query=""):
    conn = get_connection()
    cursor = conn.cursor()

    sql = """
        SELECT *
        FROM verificateurs
        WHERE actif = 1
    """
    params = []

    if query.strip():
        sql += """
            AND (
                nom LIKE ?
                OR prenom LIKE ?
                OR email LIKE ?
                OR numero_diplome LIKE ?
            )
        """
        q = f"%{query.strip()}%"
        params.extend([q, q, q, q])

    sql += " ORDER BY nom ASC, prenom ASC LIMIT 20"

    cursor.execute(sql, params)
    rows = cursor.fetchall()
    conn.close()
    return rows


if __name__ == "__main__":
    init_db()
    print(f"Base de données initialisée : {DB_PATH}")