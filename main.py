from pathlib import Path
from datetime import datetime, date
import json
import base64
import shutil
from uuid import uuid4
from copy import copy
import win32com.client

from fastapi import FastAPI, Request, Form, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment

from database import init_db, insert_verificateur, get_all_verificateurs, search_verificateurs


ADMIN_PASSWORD = "Omnilux2026"

app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key="SUPER_SECRET_KEY_CHANGE_MOI")

BASE_DIR = Path(__file__).resolve().parent

TEMPLATES_DIR = BASE_DIR / "templates"
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
SIGNATURES_DIR = BASE_DIR / "signatures"

UPLOADS_DIR = BASE_DIR / "uploads"
CARTES_DIR = UPLOADS_DIR / "cartes_identite"
DIPLOMES_DIR = UPLOADS_DIR / "diplomes"

TEMPLATES_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
SIGNATURES_DIR.mkdir(exist_ok=True)
UPLOADS_DIR.mkdir(exist_ok=True)
CARTES_DIR.mkdir(exist_ok=True)
DIPLOMES_DIR.mkdir(exist_ok=True)

templates = Jinja2Templates(directory=str(TEMPLATES_DIR))

app.mount("/uploads", StaticFiles(directory=str(UPLOADS_DIR)), name="uploads")
app.mount("/output", StaticFiles(directory=str(OUTPUT_DIR)), name="output")

@app.on_event("startup")
def startup_event():
    init_db()


@app.get("/test")
def test():
    return {"status": "ok"}


def save_upload_file(upload_file: UploadFile, destination_dir: Path) -> str:
    extension = Path(upload_file.filename).suffix.lower()
    unique_filename = f"{uuid4().hex}{extension}"
    destination = destination_dir / unique_filename

    with destination.open("wb") as buffer:
        shutil.copyfileobj(upload_file.file, buffer)

    relative_path = destination.relative_to(BASE_DIR)
    return str(relative_path).replace("\\", "/")


def save_signature_from_base64(signature_data_url: str) -> Path | None:
    if not signature_data_url:
        return None

    if not signature_data_url.startswith("data:image"):
        return None

    try:
        _, encoded = signature_data_url.split(",", 1)
        image_data = base64.b64decode(encoded)
        filename = f"signature_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        filepath = SIGNATURES_DIR / filename

        with open(filepath, "wb") as f:
            f.write(image_data)

        return filepath
    except Exception as e:
        print("Erreur sauvegarde signature :", e)
        return None


def export_excel_to_pdf(excel_path: Path, pdf_path: Path):
    excel = None
    workbook = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(str(excel_path.resolve()))

        # 1er onglet uniquement
        ws = workbook.Worksheets(1)

        # Zone d'impression : ajuste si besoin
        ws.PageSetup.PrintArea = "$A$1:$AR$133"

        # Mise en page
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 2

        # Centrage
        ws.PageSetup.CenterHorizontally = True
        ws.PageSetup.CenterVertically = True

        # Marges réduites
        ws.PageSetup.LeftMargin = excel.CentimetersToPoints(0.7)
        ws.PageSetup.RightMargin = excel.CentimetersToPoints(0.7)
        ws.PageSetup.TopMargin = excel.CentimetersToPoints(0.7)
        ws.PageSetup.BottomMargin = excel.CentimetersToPoints(0.7)

        # Orientation paysage
        ws.PageSetup.Orientation = 2  # xlLandscape

        # Export SEULEMENT de la feuille 1
        ws.ExportAsFixedFormat(0, str(pdf_path.resolve()))

    finally:
        if workbook is not None:
            workbook.Close(False)
        if excel is not None:
            excel.Quit()


def write_merged_cell(ws, cell_ref: str, value, font_size: int | None = None, bold: bool | None = None):
    target_cell = ws[cell_ref]

    for merged_range in ws.merged_cells.ranges:
        if cell_ref in merged_range:
            real_ref = merged_range.start_cell.coordinate
            print(f"{cell_ref} appartient à la fusion {merged_range} -> écriture dans {real_ref}")
            target_cell = ws[real_ref]
            break

    target_cell.value = value

    if font_size is not None or bold is not None:
        new_font = copy(target_cell.font)

        if font_size is not None:
            new_font.sz = font_size

        if bold is not None:
            new_font.b = bold

        target_cell.font = new_font

    return target_cell


def get_diplome_status(date_echeance_str: str | None):
    if not date_echeance_str:
        return {
            "label": "Date manquante",
            "color": "grey"
        }

    try:
        echeance = datetime.strptime(date_echeance_str, "%Y-%m-%d").date()
    except ValueError:
        return {
            "label": "Date invalide",
            "color": "grey"
        }

    today = date.today()
    delta_days = (echeance - today).days

    if delta_days < 0:
        return {
            "label": "Expiré",
            "color": "red"
        }
    if delta_days <= 183:
        return {
            "label": "Renouvellement < 6 mois",
            "color": "orange"
        }
    return {
        "label": "Valide",
        "color": "green"
    }


@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={"request": request}
    )


@app.post("/api/pv")
async def create_pv(request: Request):
    try:
        data = await request.json()

        print("===== DONNÉES REÇUES =====")
        print(data)
        print("==========================")

        json_filename = f"pv_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        json_path = DATA_DIR / json_filename
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        excel_template = TEMPLATES_DIR / "PV_MODELE.xlsx"
        if not excel_template.exists():
            return {
                "success": False,
                "error": f"Modèle Excel introuvable : {excel_template}"
            }

        wb = load_workbook(excel_template)
        ws = wb.active
        # Génération numéro PV (simple et propre)
        numero_pv = datetime.now().strftime("%Y%m%d%H%M%S")

        titre_pv = f"PV RÉCEPTION D’ÉCHAFAUDAGE N°{numero_pv}"

        write_merged_cell(ws, "A1", titre_pv, font_size=16, bold=True)
        print("Feuilles du classeur :", wb.sheetnames)
        print("Feuille active :", ws.title)

        # Champs principaux
        write_merged_cell(ws, "B19", data.get("chantier", ""))
        write_merged_cell(ws, "B24", data.get("adresse", ""))
        write_merged_cell(ws, "AP11", data.get("date_montage", ""))

        # Observations
        observations = data.get("observations", "")
        if observations:
            cell_obs = write_merged_cell(ws, "A61", observations, font_size=18)
            cell_obs.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        # Vérificateur
        verificateur_nom_complet = data.get("verificateur_nom", "").strip()
        verificateur_numero_diplome = data.get("verificateur_numero_diplome", "").strip()
        verificateur_lien_diplome = data.get("verificateur_lien_diplome", "").strip()

        print("verificateur_nom_complet =", verificateur_nom_complet)
        print("verificateur_numero_diplome =", verificateur_numero_diplome)
        print("verificateur_lien_diplome =", verificateur_lien_diplome)

        if verificateur_nom_complet:
            write_merged_cell(ws, "AC35", verificateur_nom_complet)

        if verificateur_numero_diplome:
            cell_diplome = write_merged_cell(ws, "AP35", verificateur_numero_diplome)

            if verificateur_lien_diplome:
                full_path = (BASE_DIR / verificateur_lien_diplome.strip("/")).resolve()

                print("Chemin absolu final :", full_path)

                if full_path.exists():
                    excel_path = str(full_path)
                    print("Lien envoyé à Excel :", excel_path)
                    cell_diplome.hyperlink = excel_path
                    cell_diplome.value = verificateur_numero_diplome
                    cell_diplome.style = "Hyperlink"
                else:
                    print("❌ FICHIER INTROUVABLE :", full_path)

        # Date / heure auto
        now = datetime.now()
        write_merged_cell(ws, "AC37", now.strftime("%d/%m/%Y"))
        write_merged_cell(ws, "AP37", now.strftime("%H:%M"))

        # Signature
        signature_data = data.get("signature", "")
        signature_path = save_signature_from_base64(signature_data)
        print("signature_path =", signature_path)

        if signature_path and signature_path.exists():
            try:
                img = Image(str(signature_path))
                img.width = 220
                img.height = 90
                ws.add_image(img, "AC39")
            except Exception as e:
                print("Erreur insertion signature Excel :", e)

        # Sauvegarde Excel
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_output_name = f"pv_{timestamp}.xlsx"
        excel_output_path = OUTPUT_DIR / excel_output_name
        # Mise en page du 1er onglet
        ws.print_area = "A1:AR61"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 2
        ws.sheet_properties.pageSetUpPr.fitToPage = True

        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.25
        ws.page_margins.bottom = 0.25

        ws.print_options.horizontalCentered = True
        ws.print_options.verticalCentered = True
        wb.save(excel_output_path)

        print(f"Chemin de sauvegarde Excel = {excel_output_path}")
        print("Fichier Excel sauvegardé OK")

        # Export PDF
        pdf_output_name = f"pv_{timestamp}.pdf"
        pdf_output_path = OUTPUT_DIR / pdf_output_name

        try:
            export_excel_to_pdf(excel_output_path, pdf_output_path)
            print(f"PDF sauvegardé OK : {pdf_output_path}")
        except Exception as e:
            print("Erreur export PDF :", e)
            pdf_output_path = None

        return {
            "success": True,
            "message": "PV généré avec succès",
            "excel_file": str(excel_output_path).replace("\\", "/"),
            "pdf_file": str(pdf_output_path).replace("\\", "/") if pdf_output_path else None,
            "json_file": str(json_path).replace("\\", "/")
        }

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {
            "success": False,
            "error": str(e)
        }


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
        context={
            "request": request,
            "error": "Mot de passe incorrect"
        }
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
        context={
            "request": request,
            "verificateurs": verificateurs
        }
    )


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