from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment
import json
import base64

app = FastAPI()

templates = Jinja2Templates(directory="templates")

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

SIGNATURES_DIR = Path("signatures")
SIGNATURES_DIR.mkdir(exist_ok=True)

OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={"request": request}
    )


def set_value_safe(ws, cell, value):
    target = cell
    for merged_range in ws.merged_cells.ranges:
        if cell in merged_range:
            target = merged_range.start_cell.coordinate
            print(f"{cell} appartient à la fusion {merged_range} -> écriture dans {target}")
            break
    ws[target] = value
    return target


def put_x(ws, cell):
    target = set_value_safe(ws, cell, "X")
    ws[target].font = Font(bold=True)


def set_check_status(ws, row, value):
    if value == "oui":
        put_x(ws, f"AP{row}")
    elif value == "non":
        put_x(ws, f"AQ{row}")
    elif value == "na":
        put_x(ws, f"AR{row}")


@app.post("/api/pv")
async def receive_pv(request: Request):
    data = await request.json()

    print("===== DONNÉES REÇUES =====")
    print(data)
    print("==========================")
    print("type_facade =", data.get("type_facade"))
    print("type_bache =", data.get("type_bache"))
    print("type_escaliers =", data.get("type_escaliers"))
    print("echafaudages_speciaux =", data.get("echafaudages_speciaux"))
    print("classe_charge =", data.get("classe_charge"))
    print("classe_largeur =", data.get("classe_largeur"))
    print("largeur_libre =", data.get("largeur_libre"))
    print("restriction_utilisation =", data.get("restriction_utilisation"))
    print("q_apparentement_intacts =", data.get("q_apparentement_intacts"))
    print("q_resistance_support =", data.get("q_resistance_support"))
    print("q_verins_reglage =", data.get("q_verins_reglage"))
    print("q_ancrages_nombre =", data.get("q_ancrages_nombre"))
    print("q_niveaux_recouverts =", data.get("q_niveaux_recouverts"))
    print("q_dispositifs_securite =", data.get("q_dispositifs_securite"))
    print("q_aux_acces =", data.get("q_aux_acces"))
    print("q_clotures =", data.get("q_clotures"))
    print("observations =", data.get("observations"))
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    numero_document = timestamp.split("_")[1]
    now = datetime.now()
    date_validation = now.strftime("%d/%m/%Y")
    heure_validation = now.strftime("%H:%M")
    # 1) Sauvegarde JSON
    json_path = DATA_DIR / f"pv_{timestamp}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # 2) Sauvegarde signature PNG
    signature_data = data.get("signature", "")
    png_path = None

    if signature_data.startswith("data:image/png;base64,"):
        base64_data = signature_data.split(",", 1)[1]
        image_bytes = base64.b64decode(base64_data)

        png_path = SIGNATURES_DIR / f"signature_{timestamp}.png"
        with open(png_path, "wb") as f:
            f.write(image_bytes)

    # 3) Génération Excel
    template_path = Path("templates") / "PV_MODELE.xlsx"
    output_path = OUTPUT_DIR / f"pv_{timestamp}.xlsx"

    wb = load_workbook(template_path)
    ws = wb.active

    print("Feuilles du classeur :", wb.sheetnames)
    print("Feuille active :", ws.title)

    # 4) Titre document
    ws["A1"] = f"PV RÉCEPTION D’ÉCHAFAUDAGE N°{numero_document}"
    ws["A1"].font = Font(name="Calibri", size=20, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # 5) Bloc haut du document
    ws["B4"] = data.get("chantier", "")
    ws["B5"] = data.get("adresse", "")
    ws["B6"] = data.get("date_montage", "")

    ws["B8"] = data.get("maitre_ouvrage", "")
    ws["B9"] = data.get("contact_mo", "")
    ws["T9"] = data.get("tel_mo", "")

    ws["B11"] = data.get("entreprise_montage", "")
    ws["B12"] = data.get("contact_montage", "")
    ws["T12"] = data.get("tel_montage", "")

    ws["B13"] = data.get("entreprise_utilisatrice", "")
    ws["B14"] = data.get("contact_utilisatrice", "")
    ws["T14"] = data.get("tel_utilisatrice", "")

    # 6) Types d’échafaudage
    if data.get("type_facade"):
        put_x(ws, "A16")

    if data.get("type_recueil"):
        put_x(ws, "A17")

    if data.get("type_filet"):
        put_x(ws, "A18")

    if data.get("type_bache"):
        put_x(ws, "H16")

    if data.get("type_plateforme"):
        put_x(ws, "H17")

    if data.get("type_escaliers"):
        put_x(ws, "H18")

    if data.get("type_toit"):
        put_x(ws, "O16")

    if data.get("type_toiture"):
        put_x(ws, "O17")

    # 7) Echafaudages spéciaux
    set_value_safe(ws, "B19", data.get("echafaudages_speciaux", ""))

    # 8) Classe de charge
    classe_charge = data.get("classe_charge", "")

    if classe_charge == "150":
        put_x(ws, "D21")
    elif classe_charge == "200":
        put_x(ws, "I21")
    elif classe_charge == "300":
        put_x(ws, "N21")
    elif classe_charge == "450":
        put_x(ws, "T21")
    elif classe_charge == "600":
        put_x(ws, "T22")

    # 9) Classe de largeur
    classe_largeur = data.get("classe_largeur", "")

    if classe_largeur == "W06":
        put_x(ws, "H23")
    elif classe_largeur == "W09":
        put_x(ws, "N23")
    elif classe_largeur == "W":
        put_x(ws, "T23")
        set_value_safe(ws, "V23", data.get("largeur_libre", ""))

    # 10) Restriction d’utilisation
    set_value_safe(ws, "B24", data.get("restriction_utilisation", ""))

    # 11) Liste de contrôle
    set_check_status(ws, 5,  data.get("q_apparentement_intacts", ""))
    set_check_status(ws, 6,  data.get("q_resistance_support", ""))
    set_check_status(ws, 7,  data.get("q_verins_reglage", ""))
    set_check_status(ws, 8,  data.get("q_contreventements", ""))
    set_check_status(ws, 9,  data.get("q_traverses_longitudinales", ""))
    set_check_status(ws, 10, data.get("q_poutres_treillis", ""))

    # Ancrages - NOMBRE
    set_value_safe(ws, "AP11", data.get("q_ancrages_nombre", ""))

    set_check_status(ws, 12, data.get("q_niveaux_recouverts", ""))
    set_check_status(ws, 13, data.get("q_planchers_compris", ""))
    set_check_status(ws, 14, data.get("q_au_niveau_des_angles", ""))
    set_check_status(ws, 15, data.get("q_madriers", ""))
    set_check_status(ws, 16, data.get("q_ouvertures", ""))
    set_check_status(ws, 17, data.get("q_dispositifs_securite", ""))
    set_check_status(ws, 18, data.get("q_distance_mur", ""))
    set_check_status(ws, 19, data.get("q_garde_corps_interieur", ""))
    set_check_status(ws, 20, data.get("q_montees_acces", ""))
    set_check_status(ws, 21, data.get("q_tour_escaliers", ""))
    set_check_status(ws, 22, data.get("q_echelle_appui", ""))
    set_check_status(ws, 23, data.get("q_exigences_recueil", ""))
    set_check_status(ws, 24, data.get("q_conduites_tension", ""))
    set_check_status(ws, 25, data.get("q_ecran_protection", ""))
    set_check_status(ws, 26, data.get("q_toit_protection_ctrl", ""))
    set_check_status(ws, 27, data.get("q_securite_circulation", ""))
    set_check_status(ws, 28, data.get("q_aux_acces", ""))
    set_check_status(ws, 29, data.get("q_clotures", ""))
    # 12) Type d’entreprise (AB33:AB34 fusionnées)
    type_entreprise = data.get("type_entreprise", "")
    if type_entreprise == "montage":
        set_value_safe(ws, "AB33", "Entreprise de montage")
    elif type_entreprise == "propre":
        set_value_safe(ws, "AB33", "Entreprise de montage pour usage propre")
        # 13) Observations (A61:AR137 fusionnées)
    ws["A61"] = data.get("observations", "")
    ws["A61"].font = Font(name="Calibri", size=18)
    ws["A61"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
       
        # Vérificateur + date/heure automatiques
    ws["AC35"] = data.get("verificateur_nom", "")
    ws["AC37"] = date_validation
    ws["AP37"] = heure_validation
       
        # 14) Signature
    if png_path and png_path.exists():
        img = Image(str(png_path))
        img.width = 200
        img.height = 40
        ws.add_image(img, "AC38")

    print("Chemin de sauvegarde Excel =", output_path)
    wb.save(output_path)
    print("Fichier Excel sauvegardé OK")

    return {
        "status": "ok",
        "message": "PV reçu",
        "saved_json": str(json_path),
        "saved_signature": str(png_path) if png_path else None,
        "saved_excel": str(output_path)
    }