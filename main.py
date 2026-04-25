from pathlib import Path
from datetime import datetime, date
from uuid import uuid4
from copy import copy
import json
import base64
import shutil
import io
import traceback
import subprocess
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from fastapi import FastAPI, Request, Form, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU

from PIL import Image as PILImage

from database import init_db, insert_verificateur, get_all_verificateurs, search_verificateurs


# =========================================================
# CONFIG
# =========================================================

ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "Omnilux2026")
APP_SECRET_KEY = os.getenv("APP_SECRET_KEY", "SUPER_SECRET_KEY_CHANGE_MOI")

APP_PUBLIC_URL = os.getenv("APP_PUBLIC_URL", "http://127.0.0.1:8000").rstrip("/")

SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.office365.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
SMTP_FROM = os.getenv("SMTP_FROM", SMTP_USER)

MAX_SOCIETES_UTILISATRICES = 10

ROW_START_SOCIETES = 33
ROW_END_SOCIETES = 42

COL_SOCIETE = "A"
COL_REPRESENTANT = "H"
COL_DATE = "O"
COL_SIGNATURE = "T"

# ---- Zone vérificateur
VERIF_SIGNATURE_CELL = "AC38:AR38"
VERIF_NAME_CELL = "AC35"
VERIF_DATE_CELL = "AC37"
VERIF_HOUR_CELL = "AP37"
VERIF_DIPLOME_CELL = "AP35"

# ---- Zone entreprise utilisatrice / maître d'œuvre
CLIENT_NAME_CELL = "AI41"
CLIENT_DATE_CELL = "AC42"
CLIENT_HOUR_CELL = "AP42"
CLIENT_SIGNATURE_CELL = "AC43"

# ---- Type d'entreprise
TYPE_ENTREPRISE_TEXT_CELL = "AB33"

# ---- Observations
OBSERVATIONS_CELL = "A48"
PRINT_AREA = "$A$1:$AR$124"

ATTENTE_FILL = PatternFill(fill_type="solid", fgColor="FF4D4D")
ATTENTE_FONT = Font(name="Arial", size=9, bold=True, color="FFFFFF")
ATTENTE_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
SIGNATURES_DIR = BASE_DIR / "signatures"

UPLOADS_DIR = BASE_DIR / "uploads"
CARTES_DIR = UPLOADS_DIR / "cartes_identite"
DIPLOMES_DIR = UPLOADS_DIR / "diplomes"


# =========================================================
# INITIALISATION DOSSIERS
# =========================================================

for directory in [
    TEMPLATES_DIR,
    DATA_DIR,
    OUTPUT_DIR,
    SIGNATURES_DIR,
    UPLOADS_DIR,
    CARTES_DIR,
    DIPLOMES_DIR,
]:
    directory.mkdir(parents=True, exist_ok=True)


# =========================================================
# APP
# =========================================================

app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key=APP_SECRET_KEY)

templates = Jinja2Templates(directory=str(TEMPLATES_DIR))

app.mount("/uploads", StaticFiles(directory=str(UPLOADS_DIR)), name="uploads")
app.mount("/output", StaticFiles(directory=str(OUTPUT_DIR)), name="output")
app.mount("/data", StaticFiles(directory=str(DATA_DIR)), name="data")


# =========================================================
# STARTUP
# =========================================================

@app.on_event("startup")
def startup_event():
    init_db()


@app.get("/test")
def test():
    return {"status": "ok"}


# =========================================================
# DETECTION DU MODELE EXCEL
# =========================================================

def find_excel_template() -> Path:
    candidates = [
        TEMPLATES_DIR / "PV_MODELE.xlsx",
        TEMPLATES_DIR / "PV_MODELE.xlsm",
        TEMPLATES_DIR / "PV_MODELE.xltx",
        TEMPLATES_DIR / "PV_MODELE",
    ]

    for path in candidates:
        if path.exists():
            return path

    for path in TEMPLATES_DIR.iterdir():
        if path.is_file() and path.stem.upper() == "PV_MODELE":
            return path

    raise FileNotFoundError(
        "Aucun modèle Excel trouvé dans templates/. Nom attendu : PV_MODELE.xlsx."
    )


# =========================================================
# HELPERS - FILES / JSON
# =========================================================

def save_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def build_paths_for_dossier(dossier_id: str) -> dict:
    dossier_dir = DATA_DIR / dossier_id
    dossier_dir.mkdir(parents=True, exist_ok=True)

    temp_dir = dossier_dir / "tmp"
    temp_dir.mkdir(parents=True, exist_ok=True)

    return {
        "dossier_dir": dossier_dir,
        "json_path": dossier_dir / "state.json",
        "xlsx_path": OUTPUT_DIR / f"{dossier_id}.xlsx",
        "pdf_path": OUTPUT_DIR / f"{dossier_id}.pdf",
        "temp_dir": temp_dir,
    }


def save_upload_file(upload_file: UploadFile, destination_dir: Path) -> str:
    extension = Path(upload_file.filename).suffix.lower()
    unique_filename = f"{uuid4().hex}{extension}"
    destination = destination_dir / unique_filename

    with destination.open("wb") as buffer:
        shutil.copyfileobj(upload_file.file, buffer)

    relative_path = destination.relative_to(BASE_DIR)
    return str(relative_path).replace("\\", "/")


# =========================================================
# HELPERS - SIGNATURES / IMAGES
# =========================================================

def decode_base64_image(signature_data_url: str) -> bytes | None:
    if not signature_data_url or not signature_data_url.startswith("data:image"):
        return None

    try:
        _, encoded = signature_data_url.split(",", 1)
        return base64.b64decode(encoded)
    except Exception:
        return None


def save_signature_from_base64(signature_data_url: str) -> Path | None:
    image_data = decode_base64_image(signature_data_url)
    if not image_data:
        return None

    try:
        filename = f"signature_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.png"
        filepath = SIGNATURES_DIR / filename

        with filepath.open("wb") as f:
            f.write(image_data)

        return filepath
    except Exception as e:
        print("Erreur sauvegarde signature :", e)
        return None


def save_base64_signature_to_temp_png(signature_b64: str, output_path: Path) -> Path:
    image_data = decode_base64_image(signature_b64)
    if not image_data:
        raise ValueError("Signature vide ou invalide")

    image = PILImage.open(io.BytesIO(image_data)).convert("RGBA")
    background = PILImage.new("RGBA", image.size, (255, 255, 255, 255))
    image = PILImage.alpha_composite(background, image).convert("RGB")
    image.save(output_path, format="PNG")

    return output_path


def excel_col_width_to_pixels(width):
    if width is None:
        width = 8.43
    return int(width * 7 + 5)


def excel_row_height_to_pixels(height):
    if height is None:
        height = 15
    return int(height * 96 / 72)


def get_cell_or_range_bounds(ws, cell_or_range: str):
    """
    Accepte :
    - une cellule simple : AC38
    - une plage : AC38:AR38
    - une cellule appartenant à une zone fusionnée
    """
    if ":" in cell_or_range:
        min_col, min_row, max_col, max_row = range_boundaries(cell_or_range)
        return min_col, min_row, max_col, max_row

    for merged_range in ws.merged_cells.ranges:
        if cell_or_range in merged_range:
            return (
                merged_range.min_col,
                merged_range.min_row,
                merged_range.max_col,
                merged_range.max_row,
            )

    col_letters = "".join(filter(str.isalpha, cell_or_range))
    row_digits = "".join(filter(str.isdigit, cell_or_range))

    col = column_index_from_string(col_letters)
    row = int(row_digits)

    return col, row, col, row


def get_range_size_pixels(ws, min_col, min_row, max_col, max_row):
    total_width = 0
    total_height = 0

    for col_idx in range(min_col, max_col + 1):
        col_letter = get_column_letter(col_idx)
        width = ws.column_dimensions[col_letter].width
        total_width += excel_col_width_to_pixels(width)

    for row_idx in range(min_row, max_row + 1):
        height = ws.row_dimensions[row_idx].height
        total_height += excel_row_height_to_pixels(height)

    return total_width, total_height


def insert_signature_fit_area(ws, cell_or_range: str, image_path: Path, padding_px: int = 4):
    min_col, min_row, max_col, max_row = get_cell_or_range_bounds(ws, cell_or_range)
    box_width, box_height = get_range_size_pixels(ws, min_col, min_row, max_col, max_row)

    max_width = max(20, box_width - (padding_px * 2))
    max_height = max(20, box_height - (padding_px * 2))

    with PILImage.open(image_path) as pil_img:
        img_width, img_height = pil_img.size

    ratio = min(max_width / img_width, max_height / img_height)
    final_width = max(1, int(img_width * ratio))
    final_height = max(1, int(img_height * ratio))

    offset_x = int((box_width - final_width) / 2)
    offset_y = int((box_height - final_height) / 2)

    xl_img = XLImage(str(image_path))
    xl_img.width = final_width
    xl_img.height = final_height

    marker = AnchorMarker(
        col=min_col - 1,
        row=min_row - 1,
        colOff=pixels_to_EMU(offset_x),
        rowOff=pixels_to_EMU(offset_y),
    )

    size = XDRPositiveSize2D(
        cx=pixels_to_EMU(final_width),
        cy=pixels_to_EMU(final_height),
    )

    xl_img.anchor = OneCellAnchor(_from=marker, ext=size)
    ws.add_image(xl_img)


# =========================================================
# HELPERS - EXCEL / PDF
# =========================================================

def export_excel_to_pdf(excel_path: Path, pdf_path: Path):
    """
    Conversion XLSX -> PDF via LibreOffice headless.
    Nécessite LibreOffice installé sur Railway.
    """
    output_dir = pdf_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)

    cmd = [
        "soffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(excel_path),
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        raise RuntimeError(
            f"Conversion PDF échouée : {result.stderr or result.stdout}"
        )

    generated_pdf = output_dir / f"{excel_path.stem}.pdf"

    if not generated_pdf.exists():
        raise RuntimeError("Le PDF n'a pas été généré par LibreOffice")

    if generated_pdf.resolve() != pdf_path.resolve():
        generated_pdf.replace(pdf_path)


def write_merged_cell(ws, cell_ref: str, value, font_size: int | None = None, bold: bool | None = None):
    target_cell = ws[cell_ref]

    for merged_range in ws.merged_cells.ranges:
        if cell_ref in merged_range:
            target_cell = ws[merged_range.start_cell.coordinate]
            break

    target_cell.value = value if value is not None else ""

    if font_size is not None or bold is not None:
        new_font = copy(target_cell.font)

        if font_size is not None:
            new_font.sz = font_size

        if bold is not None:
            new_font.b = bold

        target_cell.font = new_font

    return target_cell


def clear_societes_table(ws):
    for row in range(ROW_START_SOCIETES, ROW_END_SOCIETES + 1):
        for col in [COL_SOCIETE, COL_REPRESENTANT, COL_DATE, COL_SIGNATURE]:
            cell = ws[f"{col}{row}"]
            cell.value = None

            if col == COL_SIGNATURE:
                cell.fill = PatternFill(fill_type=None)
                cell.font = Font(name="Arial", size=10, bold=False, color="000000")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def validate_societes_limit(societes_utilisatrices: list):
    if len(societes_utilisatrices) > MAX_SOCIETES_UTILISATRICES:
        raise ValueError("Maximum 10 sociétés utilisatrices sur ce PV")


def fill_societes_utilisatrices_table(ws, societes_utilisatrices: list, temp_dir: Path):
    validate_societes_limit(societes_utilisatrices)
    clear_societes_table(ws)
    temp_dir.mkdir(parents=True, exist_ok=True)

    for index, soc in enumerate(societes_utilisatrices):
        row = ROW_START_SOCIETES + index

        cell_societe = f"{COL_SOCIETE}{row}"
        cell_representant = f"{COL_REPRESENTANT}{row}"
        cell_date = f"{COL_DATE}{row}"
        cell_signature = f"{COL_SIGNATURE}{row}"

        ws[cell_societe] = soc.get("societe", "")
        ws[cell_representant] = soc.get("representant", "")

        signed = bool(soc.get("signed", False))
        signature_b64 = soc.get("signature_b64", "")

        if signed and signature_b64:
            date_signature = soc.get("date_signature", "")
            heure_signature = soc.get("heure_signature", "")
            ws[cell_date] = f"{date_signature} {heure_signature}".strip()

            temp_signature_path = temp_dir / f"signature_societe_{row}.png"
            save_base64_signature_to_temp_png(signature_b64, temp_signature_path)
            insert_signature_fit_area(ws, cell_signature, temp_signature_path)
        else:
            ws[cell_date] = ""
            ws[cell_signature].fill = ATTENTE_FILL
            ws[cell_signature].font = ATTENTE_FONT
            ws[cell_signature].alignment = ATTENTE_ALIGNMENT
            ws[cell_signature] = "EN ATTENTE"


def apply_page_setup(ws):
    ws.print_area = PRINT_AREA


# =========================================================
# EMAIL
# =========================================================

def send_email(destinataires: str, sujet: str, contenu: str):
    if not destinataires:
        return

    if not SMTP_USER or not SMTP_PASSWORD or not SMTP_FROM:
        print("Email non envoyé : variables SMTP manquantes")
        return

    try:
        msg = MIMEMultipart()
        msg["From"] = SMTP_FROM
        msg["To"] = destinataires
        msg["Subject"] = sujet
        msg.attach(MIMEText(contenu, "plain", "utf-8"))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)

    except Exception as e:
        print("Erreur envoi mail :", e)


# =========================================================
# MAPPING EXCEL
# =========================================================

TEXT_CELL_MAP = {
    "chantier": "B4",
    "adresse": "B5",
    "date_montage": "B6",
    "maitre_ouvrage": "B8",
    "contact_mo": "B9",
    "tel_mo": "T9",
    "entreprise_montage": "B11",
    "contact_montage": "B12",
    "tel_montage": "T12",
    "entreprise_utilisatrice": "B13",
    "contact_utilisatrice": "B14",
    "tel_utilisatrice": "T14",
    "echafaudages_speciaux": "B19",
    "restriction_utilisation": "B24",
}

TYPE_ECHAFAUDAGE_MAP = {
    "type_facade": "A16",
    "type_recueil": "A17",
    "type_filet": "A18",
    "type_bache": "H16",
    "type_plateforme": "H17",
    "type_escaliers": "H18",
    "type_toit": "O16",
    "type_toiture": "O17",
}

CLASSE_CHARGE_MAP = {
    "150": "D21",
    "200": "I21",
    "300": "N21",
    "450": "T21",
    "600": "T22",
}

CLASSE_LARGEUR_MAP = {
    "W06": "H23",
    "W09": "N23",
    "W": "T23",
}

CHECKLIST_MAP = {
    "q_apparentement_intacts": {"oui": "AP4", "non": "AQ4", "na": "AR4"},
    "q_resistance_support": {"oui": "AP6", "non": "AQ6", "na": "AR6"},
    "q_verins_reglage": {"oui": "AP7", "non": "AQ7", "na": "AR7"},
    "q_contreventements": {"oui": "AP8", "non": "AQ8", "na": "AR8"},
    "q_traverses_longitudinales": {"oui": "AP9", "non": "AQ9", "na": "AR9"},
    "q_poutres_treillis": {"oui": "AP10", "non": "AQ10", "na": "AR10"},
    "q_ancrages_nombre": {"value": "AP11"},
    "q_niveaux_recouverts": {"oui": "AP12", "non": "AQ12", "na": "AR12"},
    "q_planchers_compris": {"oui": "AP13", "non": "AQ13", "na": "AR13"},
    "q_au_niveau_des_angles": {"oui": "AP14", "non": "AQ14", "na": "AR14"},
    "q_madriers": {"oui": "AP15", "non": "AQ15", "na": "AR15"},
    "q_ouvertures": {"oui": "AP16", "non": "AQ16", "na": "AR16"},
    "q_dispositifs_securite": {"oui": "AP17", "non": "AQ17", "na": "AR17"},
    "q_distance_mur": {"oui": "AP18", "non": "AQ18", "na": "AR18"},
    "q_garde_corps_interieur": {"oui": "AP19", "non": "AQ19", "na": "AR19"},
    "q_montees_acces": {"oui": "AP20", "non": "AQ20", "na": "AR20"},
    "q_tour_escaliers": {"oui": "AP21", "non": "AQ21", "na": "AR21"},
    "q_echelle_appui": {"oui": "AP22", "non": "AQ22", "na": "AR22"},
    "q_exigences_recueil": {"oui": "AP23", "non": "AQ23", "na": "AR23"},
    "q_conduites_tension": {"oui": "AP24", "non": "AQ24", "na": "AR24"},
    "q_ecran_protection": {"oui": "AP25", "non": "AQ25", "na": "AR25"},
    "q_toit_protection_ctrl": {"oui": "AP26", "non": "AQ26", "na": "AR26"},
    "q_securite_circulation": {"oui": "AP27", "non": "AQ27", "na": "AR27"},
    "q_aux_acces": {"oui": "AP28", "non": "AQ28", "na": "AR28"},
    "q_clotures": {"oui": "AP29", "non": "AQ29", "na": "AR29"},
}


def mark_x(ws, cell_ref: str):
    cell = write_merged_cell(ws, cell_ref, "X", font_size=12, bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    return cell


def fill_simple_text_fields(ws, dossier_data: dict):
    for payload_key, cell_ref in TEXT_CELL_MAP.items():
        value = dossier_data.get(payload_key, "")
        write_merged_cell(ws, cell_ref, value)


def fill_type_echafaudage_fields(ws, dossier_data: dict):
    for payload_key, cell_ref in TYPE_ECHAFAUDAGE_MAP.items():
        if dossier_data.get(payload_key, False):
            mark_x(ws, cell_ref)


def fill_type_entreprise_field(ws, dossier_data: dict):
    value = str(dossier_data.get("type_entreprise", "")).strip().lower()

    if value == "montage":
        texte = "Entreprise de montage"
    elif value == "propre":
        texte = "Entreprise de montage pour usage propre"
    else:
        texte = ""

    if texte:
        cell = write_merged_cell(ws, TYPE_ENTREPRISE_TEXT_CELL, texte)
        cell.font = Font(name="Calibri", size=18, bold=False)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def fill_classe_charge(ws, dossier_data: dict):
    value = str(dossier_data.get("classe_charge", "")).strip()
    cell_ref = CLASSE_CHARGE_MAP.get(value)
    if cell_ref:
        mark_x(ws, cell_ref)


def fill_classe_largeur(ws, dossier_data: dict):
    value = str(dossier_data.get("classe_largeur", "")).strip().upper()
    cell_ref = CLASSE_LARGEUR_MAP.get(value)
    if cell_ref:
        mark_x(ws, cell_ref)

    if value == "W":
        largeur_libre = str(dossier_data.get("largeur_libre", "")).strip()
        if largeur_libre:
            current = ws["T23"].value or ""
            ws["T23"] = f"X ({largeur_libre})" if current == "X" else largeur_libre


def fill_checklist_fields(ws, dossier_data: dict):
    for payload_key, mapping in CHECKLIST_MAP.items():
        value = dossier_data.get(payload_key, "")

        if "value" in mapping:
            if value not in ("", None):
                write_merged_cell(ws, mapping["value"], value)
            continue

        value = str(value).strip().lower()
        cell_ref = mapping.get(value)

        if cell_ref:
            mark_x(ws, cell_ref)


def fill_observations_block(ws, dossier_data: dict):
    observations = dossier_data.get("observations", "")
    cell_obs = write_merged_cell(ws, OBSERVATIONS_CELL, observations, font_size=18)
    cell_obs.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)


def fill_verificateur_block(ws, dossier_data: dict):
    verificateur_nom = dossier_data.get("verificateur_nom", "").strip()
    verificateur_numero_diplome = dossier_data.get("verificateur_numero_diplome", "").strip()
    verificateur_lien_diplome = dossier_data.get("verificateur_lien_diplome", "").strip()

    if verificateur_nom:
        write_merged_cell(ws, VERIF_NAME_CELL, verificateur_nom)

    if verificateur_numero_diplome:
        cell_diplome = write_merged_cell(ws, VERIF_DIPLOME_CELL, verificateur_numero_diplome)

        if verificateur_lien_diplome:
            full_path = (BASE_DIR / verificateur_lien_diplome.strip("/")).resolve()
            if full_path.exists():
                cell_diplome.hyperlink = str(full_path)
                cell_diplome.value = verificateur_numero_diplome
                cell_diplome.style = "Hyperlink"

    verification_datetime = dossier_data.get("verification_datetime")
    if verification_datetime:
        try:
            dt = datetime.fromisoformat(verification_datetime)
        except ValueError:
            dt = datetime.now()
    else:
        dt = datetime.now()
        dossier_data["verification_datetime"] = dt.isoformat()

    write_merged_cell(ws, VERIF_DATE_CELL, dt.strftime("%d/%m/%Y"))
    write_merged_cell(ws, VERIF_HOUR_CELL, dt.strftime("%H:%M"))

    signature_data = dossier_data.get("signature", "")
    signature_path = save_signature_from_base64(signature_data)

    if signature_path and signature_path.exists():
        try:
            insert_signature_fit_area(ws, VERIF_SIGNATURE_CELL, signature_path, padding_px=4)
        except Exception as e:
            print("Erreur insertion signature vérificateur Excel :", e)


def fill_client_block(ws, dossier_data: dict):
    client = dossier_data.get("client_signature", {}) or {}

    client_nom = client.get("nom_signataire", "").strip()
    client_signature = client.get("signature_b64", "")
    client_datetime = client.get("signature_datetime")

    if client_nom:
        write_merged_cell(ws, CLIENT_NAME_CELL, client_nom)

    if client_datetime:
        try:
            dt = datetime.fromisoformat(client_datetime)
        except ValueError:
            dt = None
    else:
        dt = None

    if dt:
        write_merged_cell(ws, CLIENT_DATE_CELL, dt.strftime("%d/%m/%Y"))
        write_merged_cell(ws, CLIENT_HOUR_CELL, dt.strftime("%H:%M"))

    if client_signature:
        signature_path = save_signature_from_base64(client_signature)
        if signature_path and signature_path.exists():
            try:
                insert_signature_fit_area(ws, CLIENT_SIGNATURE_CELL, signature_path, padding_px=4)
            except Exception as e:
                print("Erreur insertion signature client Excel :", e)


def regenerate_excel_from_data(output_xlsx_path: Path, dossier_data: dict, temp_dir: Path):
    template_path = find_excel_template()
    shutil.copy2(template_path, output_xlsx_path)

    wb = load_workbook(output_xlsx_path)
    ws = wb["Formulaire"] if "Formulaire" in wb.sheetnames else wb.active

    numero_pv = dossier_data.get("numero_pv") or datetime.now().strftime("%Y%m%d%H%M%S")
    dossier_data["numero_pv"] = numero_pv

    ws["A1"] = f"PROCÈS-VERBAL DE CONTRÔLE N°{numero_pv}"
    ws["A1"].font = Font(name="Calibri", size=16, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    fill_simple_text_fields(ws, dossier_data)
    fill_type_echafaudage_fields(ws, dossier_data)
    fill_classe_charge(ws, dossier_data)
    fill_classe_largeur(ws, dossier_data)
    fill_checklist_fields(ws, dossier_data)
    fill_type_entreprise_field(ws, dossier_data)
    fill_observations_block(ws, dossier_data)
    fill_verificateur_block(ws, dossier_data)
    fill_client_block(ws, dossier_data)
    fill_societes_utilisatrices_table(ws, dossier_data.get("societes_utilisatrices", []), temp_dir)
    apply_page_setup(ws)

    wb.save(output_xlsx_path)


# =========================================================
# HELPERS - BUSINESS
# =========================================================

def get_diplome_status(date_echeance_str: str | None):
    if not date_echeance_str:
        return {"label": "Date manquante", "color": "grey"}

    try:
        echeance = datetime.strptime(date_echeance_str, "%Y-%m-%d").date()
    except ValueError:
        return {"label": "Date invalide", "color": "grey"}

    today = date.today()
    delta_days = (echeance - today).days

    if delta_days < 0:
        return {"label": "Expiré", "color": "red"}
    if delta_days <= 183:
        return {"label": "Renouvellement < 6 mois", "color": "orange"}
    return {"label": "Valide", "color": "green"}


def prepare_pv_payload(data: dict) -> dict:
    dossier_id = data.get("dossier_id") or f"pv_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    numero_pv = data.get("numero_pv") or datetime.now().strftime("%Y%m%d%H%M%S")

    payload = {
        "dossier_id": dossier_id,
        "numero_pv": numero_pv,
        "chantier": data.get("chantier", ""),
        "adresse": data.get("adresse", ""),
        "date_montage": data.get("date_montage", ""),
        "observations": data.get("observations", ""),
        "maitre_ouvrage": data.get("maitre_ouvrage", ""),
        "contact_mo": data.get("contact_mo", ""),
        "tel_mo": data.get("tel_mo", ""),
        "entreprise_montage": data.get("entreprise_montage", ""),
        "contact_montage": data.get("contact_montage", ""),
        "tel_montage": data.get("tel_montage", ""),
        "entreprise_utilisatrice": data.get("entreprise_utilisatrice", ""),
        "contact_utilisatrice": data.get("contact_utilisatrice", ""),
        "tel_utilisatrice": data.get("tel_utilisatrice", ""),
        "type_facade": data.get("type_facade", False),
        "type_recueil": data.get("type_recueil", False),
        "type_filet": data.get("type_filet", False),
        "type_bache": data.get("type_bache", False),
        "type_plateforme": data.get("type_plateforme", False),
        "type_escaliers": data.get("type_escaliers", False),
        "type_toit": data.get("type_toit", False),
        "type_toiture": data.get("type_toiture", False),
        "type_entreprise": data.get("type_entreprise", ""),
        "echafaudages_speciaux": data.get("echafaudages_speciaux", ""),
        "classe_charge": data.get("classe_charge", ""),
        "classe_largeur": data.get("classe_largeur", ""),
        "largeur_libre": data.get("largeur_libre", ""),
        "restriction_utilisation": data.get("restriction_utilisation", ""),
        "verificateur_nom": data.get("verificateur_nom", "").strip(),
        "verificateur_prenom": data.get("verificateur_prenom", "").strip(),
        "verificateur_email": data.get("verificateur_email", "").strip(),
        "verificateur_telephone": data.get("verificateur_telephone", "").strip(),
        "verificateur_numero_diplome": data.get("verificateur_numero_diplome", "").strip(),
        "verificateur_lien_diplome": data.get("verificateur_lien_diplome", "").strip(),
        "verificateur_statut_color": data.get("verificateur_statut_color", "").strip(),
        "verificateur_statut_label": data.get("verificateur_statut_label", "").strip(),
        "verificateur_date_echeance": data.get("verificateur_date_echeance", "").strip(),
        "signature": data.get("signature", ""),
        "verification_datetime": data.get("verification_datetime") or datetime.now().isoformat(),
        "client_signature": data.get("client_signature", {}),
        "societes_utilisatrices": data.get("societes_utilisatrices", []),
        "email_utilisatrice": data.get("email_utilisatrice", "").strip(),
        "email_mo": data.get("email_mo", "").strip(),
    }

    for key in CHECKLIST_MAP.keys():
        payload[key] = data.get(key, "")

    validate_societes_limit(payload["societes_utilisatrices"])
    return payload


def apply_client_signature_payload(dossier_data: dict, payload: dict) -> dict:
    societes_utilisatrices = payload.get("societes_utilisatrices", [])
    validate_societes_limit(societes_utilisatrices)

    dossier_data["client_signature"] = {
        "nom_signataire": payload.get("client_nom_signataire", "").strip(),
        "email": payload.get("client_email", "").strip(),
        "telephone": payload.get("client_telephone", "").strip(),
        "signature_b64": payload.get("client_signature", ""),
        "signature_datetime": datetime.now().isoformat(),
    }

    dossier_data["societes_utilisatrices"] = societes_utilisatrices
    return dossier_data


def regenerate_pv_files(dossier_data: dict) -> dict:
    dossier_id = dossier_data["dossier_id"]
    paths = build_paths_for_dossier(dossier_id)

    save_json(paths["json_path"], dossier_data)

    regenerate_excel_from_data(
        output_xlsx_path=paths["xlsx_path"],
        dossier_data=dossier_data,
        temp_dir=paths["temp_dir"],
    )

    pdf_generated = True
    try:
        export_excel_to_pdf(paths["xlsx_path"], paths["pdf_path"])
    except Exception as e:
        traceback.print_exc()
        print("Erreur export PDF :", e)
        pdf_generated = False

    return {
        "json_path": paths["json_path"],
        "xlsx_path": paths["xlsx_path"],
        "pdf_path": paths["pdf_path"] if pdf_generated else None,
    }


# =========================================================
# ROUTES - HOME / PV VERIFICATEUR
# =========================================================

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={"request": request}
    )


@app.post("/api/pv")
async def create_or_regenerate_pv(request: Request):
    try:
        raw_data = await request.json()
        data = prepare_pv_payload(raw_data)
        generated = regenerate_pv_files(data)

        emails = raw_data.get("emails_destinataires", "").strip()

        if emails and generated["pdf_path"]:
            pdf_url = f"{APP_PUBLIC_URL}/output/{generated['pdf_path'].name}"
            signature_url = f"{APP_PUBLIC_URL}/client-signature/{data['dossier_id']}"

            contenu = f"""Bonjour,

Veuillez trouver ci-dessous le PV de réception :

PDF :
{pdf_url}

Lien de vérification et signature :
{signature_url}

Cordialement,
OMNILUX
"""

            send_email(
                destinataires=emails,
                sujet=f"PV réception échafaudage n°{data['numero_pv']}",
                contenu=contenu
            )

        return {
            "success": True,
            "message": "PV généré avec succès",
            "dossier_id": data["dossier_id"],
            "numero_pv": data["numero_pv"],
            "excel_file": f"/output/{generated['xlsx_path'].name}",
            "pdf_file": f"/output/{generated['pdf_path'].name}" if generated["pdf_path"] else None,
            "json_file": f"/data/{data['dossier_id']}/state.json",
            "client_signature_url": f"/client-signature/{data['dossier_id']}",
        }

    except Exception as e:
        traceback.print_exc()
        return {"success": False, "error": str(e)}


# =========================================================
# ROUTES - CLIENT SIGNATURE
# =========================================================

@app.get("/client-signature/{dossier_id}", response_class=HTMLResponse)
def client_signature_form(request: Request, dossier_id: str):
    paths = build_paths_for_dossier(dossier_id)

    if not paths["json_path"].exists():
        return templates.TemplateResponse(
            request=request,
            name="client_signature.html",
            context={
                "request": request,
                "dossier_id": dossier_id,
                "numero_pv": "",
                "chantier": "",
            }
        )

    dossier_data = load_json(paths["json_path"])

    return templates.TemplateResponse(
        request=request,
        name="client_signature.html",
        context={
            "request": request,
            "dossier_id": dossier_id,
            "numero_pv": dossier_data.get("numero_pv", ""),
            "chantier": dossier_data.get("chantier", ""),
        }
    )


@app.post("/client-signature/{dossier_id}")
async def client_signature_submit(dossier_id: str, request: Request):
    try:
        paths = build_paths_for_dossier(dossier_id)

        if not paths["json_path"].exists():
            raise HTTPException(status_code=404, detail="Dossier introuvable")

        payload = await request.json()
        dossier_data = load_json(paths["json_path"])
        dossier_data = apply_client_signature_payload(dossier_data, payload)
        generated = regenerate_pv_files(dossier_data)

        return {
            "success": True,
            "message": "Signature client enregistrée, PV régénéré",
            "dossier_id": dossier_id,
            "excel_file": f"/output/{generated['xlsx_path'].name}",
            "pdf_file": f"/output/{generated['pdf_path'].name}" if generated["pdf_path"] else None,
            "json_file": f"/data/{dossier_id}/state.json",
        }

    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        return {"success": False, "error": str(e)}


# =========================================================
# ROUTES - SIGNATURE D'UNE SOCIETE
# =========================================================

@app.post("/api/pv/{dossier_id}/societes/{societe_index}/sign")
async def sign_societe(dossier_id: str, societe_index: int, request: Request):
    try:
        payload = await request.json()
        signature_b64 = payload.get("signature_b64", "")

        paths = build_paths_for_dossier(dossier_id)

        if not paths["json_path"].exists():
            raise HTTPException(status_code=404, detail="Dossier introuvable")

        dossier_data = load_json(paths["json_path"])
        societes = dossier_data.get("societes_utilisatrices", [])

        if societe_index < 0 or societe_index >= len(societes):
            raise HTTPException(status_code=400, detail="Index société invalide")

        if not signature_b64:
            raise HTTPException(status_code=400, detail="Signature manquante")

        now = datetime.now()
        societes[societe_index]["signed"] = True
        societes[societe_index]["date_signature"] = now.strftime("%d/%m/%Y")
        societes[societe_index]["heure_signature"] = now.strftime("%H:%M")
        societes[societe_index]["signature_b64"] = signature_b64
        dossier_data["societes_utilisatrices"] = societes

        generated = regenerate_pv_files(dossier_data)

        return {
            "success": True,
            "message": "Signature société enregistrée, PV régénéré",
            "dossier_id": dossier_id,
            "excel_file": f"/output/{generated['xlsx_path'].name}",
            "pdf_file": f"/output/{generated['pdf_path'].name}" if generated["pdf_path"] else None,
            "json_file": f"/data/{dossier_id}/state.json",
        }

    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        return {"success": False, "error": str(e)}


# =========================================================
# ROUTES - VERIFICATEURS
# =========================================================

@app.get("/verificateurs/nouveau", response_class=HTMLResponse)
def form_verificateur(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="verificateur_form.html",
        context={"request": request}
    )


@app.post("/verificateurs/nouveau", response_class=HTMLResponse)
async def create_verificateur(
    request: Request,
    nom: str = Form(...),
    prenom: str = Form(...),
    email: str = Form(...),
    telephone: str = Form(""),
    numero_diplome: str = Form(...),
    date_obtention_diplome: str = Form(...),
    date_echeance_diplome: str = Form(...),
    carte_recto: UploadFile = File(...),
    carte_verso: UploadFile = File(...),
    diplome: UploadFile = File(...)
):
    fichier_carte_recto = save_upload_file(carte_recto, CARTES_DIR)
    fichier_carte_verso = save_upload_file(carte_verso, CARTES_DIR)
    fichier_diplome = save_upload_file(diplome, DIPLOMES_DIR)

    insert_verificateur(
        nom=nom.strip(),
        prenom=prenom.strip(),
        email=email.strip(),
        telephone=telephone.strip(),
        numero_diplome=numero_diplome.strip(),
        date_obtention_diplome=date_obtention_diplome.strip(),
        date_echeance_diplome=date_echeance_diplome.strip(),
        fichier_carte_recto=fichier_carte_recto,
        fichier_carte_verso=fichier_carte_verso,
        fichier_diplome=fichier_diplome
    )

    statut = get_diplome_status(date_echeance_diplome)

    return templates.TemplateResponse(
        request=request,
        name="verificateur_success.html",
        context={
            "request": request,
            "nom": nom,
            "prenom": prenom,
            "email": email,
            "telephone": telephone,
            "numero_diplome": numero_diplome,
            "date_obtention_diplome": date_obtention_diplome,
            "date_echeance_diplome": date_echeance_diplome,
            "statut_label": statut["label"],
            "statut_color": statut["color"]
        }
    )


# =========================================================
# ROUTES - ADMIN
# =========================================================

@app.get("/admin/login", response_class=HTMLResponse)
def admin_login_form(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="admin_login.html",
        context={"request": request}
    )


@app.post("/admin/login", response_class=HTMLResponse)
async def admin_login(request: Request, password: str = Form(...)):
    if password == ADMIN_PASSWORD:
        request.session["is_admin"] = True
        return RedirectResponse("/admin/verificateurs", status_code=303)

    return templates.TemplateResponse(
        request=request,
        name="admin_login.html",
        context={"request": request, "error": "Mot de passe incorrect"}
    )


@app.get("/admin/logout")
def admin_logout(request: Request):
    request.session.clear()
    return RedirectResponse("/admin/login", status_code=303)


@app.get("/admin/verificateurs", response_class=HTMLResponse)
def liste_verificateurs(request: Request):
    if not request.session.get("is_admin"):
        return RedirectResponse("/admin/login", status_code=303)

    verificateurs_raw = get_all_verificateurs()
    verificateurs = []

    for v in verificateurs_raw:
        statut = get_diplome_status(v["date_echeance_diplome"])
        verificateurs.append({
            "id": v["id"],
            "nom": v["nom"],
            "prenom": v["prenom"],
            "email": v["email"],
            "telephone": v["telephone"],
            "numero_diplome": v["numero_diplome"],
            "date_obtention_diplome": v["date_obtention_diplome"],
            "date_echeance_diplome": v["date_echeance_diplome"],
            "fichier_carte_recto": v["fichier_carte_recto"],
            "fichier_carte_verso": v["fichier_carte_verso"],
            "fichier_diplome": v["fichier_diplome"],
            "actif": v["actif"],
            "statut_label": statut["label"],
            "statut_color": statut["color"]
        })

    return templates.TemplateResponse(
        request=request,
        name="verificateurs_liste.html",
        context={"request": request, "verificateurs": verificateurs}
    )


# =========================================================
# ROUTES - API VERIFICATEURS
# =========================================================

@app.get("/api/verificateurs")
def api_verificateurs(q: str = ""):
    verificateurs_raw = search_verificateurs(q)
    results = []

    for v in verificateurs_raw:
        statut = get_diplome_status(v["date_echeance_diplome"])

        results.append({
            "id": v["id"],
            "nom": v["nom"],
            "prenom": v["prenom"],
            "nom_complet": f'{v["nom"]} {v["prenom"]}',
            "email": v["email"],
            "telephone": v["telephone"],
            "numero_diplome": v["numero_diplome"],
            "date_obtention_diplome": v["date_obtention_diplome"],
            "date_echeance_diplome": v["date_echeance_diplome"],
            "fichier_diplome": f'/{v["fichier_diplome"]}' if v["fichier_diplome"] else "",
            "statut_label": statut["label"],
            "statut_color": statut["color"]
        })

    return JSONResponse(content=results)
