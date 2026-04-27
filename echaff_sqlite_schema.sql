CREATE TABLE IF NOT EXISTS societes (
    id TEXT PRIMARY KEY,
    nom TEXT NOT NULL DEFAULT '',
    siret TEXT DEFAULT '',
    adresse TEXT DEFAULT '',
    code_postal TEXT DEFAULT '',
    ville TEXT DEFAULT '',
    pays TEXT DEFAULT 'France',
    telephone TEXT DEFAULT '',
    email TEXT DEFAULT '',
    representant_nom TEXT DEFAULT '',
    representant_prenom TEXT DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS profils (
    id TEXT PRIMARY KEY,
    societe_id TEXT,
    nom TEXT NOT NULL,
    prenom TEXT NOT NULL,
    email TEXT NOT NULL,
    telephone TEXT DEFAULT '',
    role TEXT NOT NULL,
    actif INTEGER NOT NULL DEFAULT 1,
    signature_electronique TEXT DEFAULT '',
    certification TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS chantiers (
    id TEXT PRIMARY KEY,
    societe_id TEXT,
    nom TEXT NOT NULL,
    reference_interne TEXT UNIQUE NOT NULL,
    adresse_complete TEXT DEFAULT '',
    batiment_zone_etage_secteur TEXT DEFAULT '',
    client_maitre_ouvrage TEXT DEFAULT '',
    date_debut TEXT DEFAULT '',
    date_fin_estimee TEXT DEFAULT '',
    date_fin_reelle TEXT DEFAULT '',
    statut TEXT NOT NULL DEFAULT 'brouillon',
    societe_echafaudage_responsable TEXT DEFAULT '',
    societes_utilisatrices_autorisees TEXT NOT NULL DEFAULT '[]',
    documents_associes TEXT NOT NULL DEFAULT '[]',
    historique TEXT NOT NULL DEFAULT '[]',
    qr_token TEXT UNIQUE,
    qr_code_url TEXT DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS pv_reception (
    id TEXT PRIMARY KEY,
    dossier_id TEXT UNIQUE NOT NULL,
    numero_pv TEXT NOT NULL,
    chantier_id TEXT,
    chantier_nom TEXT DEFAULT '',
    statut_document TEXT NOT NULL DEFAULT 'pv_reception',
    excel_file TEXT DEFAULT '',
    pdf_file TEXT DEFAULT '',
    json_file TEXT DEFAULT '',
    client_signature_url TEXT DEFAULT '',
    data TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS historique_actions (
    id TEXT PRIMARY KEY,
    societe_id TEXT,
    chantier_id TEXT,
    pv_id TEXT,
    type_action TEXT NOT NULL,
    description TEXT DEFAULT '',
    auteur TEXT DEFAULT 'system',
    metadata TEXT NOT NULL DEFAULT '{}',
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS verificateurs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nom TEXT NOT NULL,
    prenom TEXT NOT NULL,
    email TEXT NOT NULL,
    telephone TEXT DEFAULT '',
    numero_diplome TEXT NOT NULL,
    date_obtention_diplome TEXT NOT NULL,
    date_echeance_diplome TEXT NOT NULL,
    fichier_carte_recto TEXT DEFAULT '',
    fichier_carte_verso TEXT DEFAULT '',
    fichier_diplome TEXT DEFAULT '',
    actif INTEGER NOT NULL DEFAULT 1,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
