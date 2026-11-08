# -*- coding: utf-8 -*-
"""
Génère automatiquement un PowerPoint (une diapo par ligne) à partir du planning Excel.
Aucun modèle PowerPoint n'est nécessaire.

Installation (une seule fois) :  python -m pip install openpyxl python-pptx
Utilisation :
    python generer_pptx.py                         -> prend le fichier .xlsx du dossier
    python generer_pptx.py "mon planning.xlsx"     -> ou un fichier précis
    python generer_pptx.py planning.xlsx -o rapport.pptx   -> avec le nom du PowerPoint voulu
    python generer_pptx.py --help                  -> aide

Logos clients (facultatif) : placez image_1.jpeg, image_2.jpeg, image_3.jpeg ...
dans le dossier du script (ou dans un sous-dossier "logos"). Ils sont affichés
en haut à droite de chaque diapo (première et dernière pages comprises),
dans l'ordre des numéros.
"""
import argparse
import glob
import io
import json
import os
import re
import ssl
import sys
import urllib.error
import urllib.parse
import urllib.request
from html.parser import HTMLParser

import openpyxl
from PIL import Image, ImageOps
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

# ======================= PARAMÈTRES À ADAPTER =======================
# Titre de la page de garde : une ligne par élément de la liste
TITRE_LIGNES    = ["PRESENTATION AFFICHAGE URBAIN"]
# ex : ["CITROEN C4X", "OPEL MOKKA", "PEUGEOT LANDTREK"]

LOGOS_SUR_LIGNES = True         # False = logos clients seulement sur la première et la dernière page
PREFIXE_PERIODE = ""            # ex : "Q1 " -> "Q1 OCTOBRE 2026" (sinon "OCTOBRE 2026")
TEXTE_PHOTO     = "PHOTO DE JOUR"
TEXTE_FIN       = "Merci pour votre confiance."

# Utilisées seulement si l'en-tête de l'Excel ne contient pas "Du jj/mm/aaaa au jj/mm/aaaa"
PERIODE_DU_DEFAUT = "01/12/2026"
PERIODE_AU_DEFAUT = "31/12/2026"

SEULEMENT_DISPONIBLES = False   # True = ignorer les lignes dont "Dispo" n'est pas DISPONIBLE

# --- Photos des emplacements ---
TELECHARGER_IMAGES = True       # télécharge la photo depuis le lien de la colonne Code
DOSSIER_IMAGES  = "images"      # photos enregistrées en images/CODE.jpg (vous pouvez aussi les déposer à la main)
MODE_PHOTO      = "remplir"     # "remplir" = toutes les photos à la même taille (rognées au centre) ; "ajuster" = photo entière
FORCER_TELECHARGEMENT = False   # True = retélécharger même si l'image existe déjà
VERIFIER_SSL    = True          # mettre False seulement en cas d'erreur de certificat SSL
FICHIER_LOGO    = "logo.png"    # facultatif : logo en haut à gauche (sinon pris dans l'Excel)
# ====================================================================

ROUGE = RGBColor(0xFF, 0x00, 0x00)
NOIR = RGBColor(0, 0, 0)
POLICE = "Times New Roman"
MOIS = ["JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN", "JUILLET",
        "AOUT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE"]
DOSSIER_SCRIPT = os.path.dirname(os.path.abspath(__file__))


# ----------------------------- Excel -----------------------------
def lire_arguments():
    p = argparse.ArgumentParser(
        description="Génère le PowerPoint d'affichage à partir d'un planning Excel.")
    p.add_argument("excel", nargs="?",
                   help="fichier Excel (.xlsx) ; par défaut : le .xlsx du dossier du script")
    p.add_argument("-o", "--sortie",
                   help="nom du PowerPoint à créer (défaut : PRESENTATION_<nom de l'Excel>.pptx)")
    return p.parse_args()


def trouver_excel(argument):
    if argument:
        essais = [argument, argument + ".xlsx", os.path.join(DOSSIER_SCRIPT, argument),
                  os.path.join(DOSSIER_SCRIPT, argument + ".xlsx")]
        for chemin in essais:
            if os.path.isfile(chemin):
                return chemin
        sys.exit(f"Fichier Excel introuvable : {argument}")
    fichiers = [f for f in glob.glob(os.path.join(DOSSIER_SCRIPT, "*.xlsx"))
                if not os.path.basename(f).startswith("~$")]
    if not fichiers:
        sys.exit("Aucun fichier .xlsx trouvé dans le dossier du script.\n"
                 'Utilisez : python generer_pptx.py "mon_fichier.xlsx"')
    if len(fichiers) > 1:
        print("Plusieurs fichiers Excel trouvés, le plus récent est utilisé.")
    return max(fichiers, key=os.path.getmtime)


def num(v):
    """Accepte 15, 15.0, '15.00', '2,5' ..."""
    return float(str(v).strip().replace(",", "."))


def nb(x):
    """15.0 -> '15' ; 2.5 -> '2.5'"""
    x = float(x)
    return str(int(x)) if x.is_integer() else f"{x:g}"


def lien_cellule(cellule):
    """Lien contenu dans la cellule Code (hyperlien, formule HYPERLINK ou URL écrite en clair)."""
    if cellule.hyperlink and cellule.hyperlink.target:
        return cellule.hyperlink.target.strip()
    v = str(cellule.value or "")
    m = re.match(r'\s*=\s*HYPERLINK\(\s*"([^"]+)"', v, flags=re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return v.strip() if v.strip().lower().startswith(("http://", "https://")) else None


def lire_excel(chemin):
    try:
        ws = openpyxl.load_workbook(chemin).active
    except PermissionError:
        sys.exit("Impossible de lire l'Excel (accès refusé).\n"
                 "  1) Fermez le fichier dans Excel puis relancez.\n"
                 "  2) Si le dossier est dans OneDrive : clic droit sur le fichier > "
                 "'Toujours conserver sur cet appareil', ou copiez le projet hors OneDrive.")
    except FileNotFoundError:
        sys.exit(f"Fichier Excel introuvable : {chemin}")
    ligne_entete = next((r for r in range(1, ws.max_row + 1)
                         if str(ws.cell(r, 2).value).strip().lower() == "nombre"), None)
    if ligne_entete is None:
        sys.exit('Colonne "Nombre" introuvable dans l\'Excel (en-têtes attendus en colonne B).')

    entetes = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(ligne_entete, c).value
        if v is not None:
            entetes[str(v).strip().lower()] = c
    manquantes = [n for n in ("code", "emplacement", "largeur / m", "hauteur / m") if n not in entetes]
    if manquantes:
        sys.exit(f"Colonnes introuvables dans l'Excel : {manquantes}")

    col_dispo = max(entetes.values())                       # dernière colonne = période + dispo
    texte_periode = str(ws.cell(ligne_entete, col_dispo).value or "")

    lignes = []
    r = ligne_entete + 1
    while ws.cell(r, entetes["code"]).value:
        largeur = num(ws.cell(r, entetes["largeur / m"]).value)
        hauteur = num(ws.cell(r, entetes["hauteur / m"]).value)
        cellule_code = ws.cell(r, entetes["code"])
        lignes.append({
            "numero": int(ws.cell(r, entetes["nombre"]).value or len(lignes) + 1),
            "code": str(cellule_code.value).strip(),
            "lien": lien_cellule(cellule_code),
            "emplacement": " ".join(str(ws.cell(r, entetes["emplacement"]).value or "").split()),
            "largeur": largeur,
            "hauteur": hauteur,
            "surface": largeur * hauteur,       # recalculée (la formule Excel n'est pas nécessaire)
            "dispo": str(ws.cell(r, col_dispo).value or "").strip(),
        })
        r += 1

    logo = None
    chemin_logo = os.path.join(DOSSIER_SCRIPT, FICHIER_LOGO)
    if os.path.exists(chemin_logo):
        with open(chemin_logo, "rb") as f:
            logo = f.read()
    elif ws._images:
        logo = ws._images[0]._data()
    return lignes, texte_periode, logo


def lire_periode(texte):
    """'Du 01/12/2026 au 31/12/2026' (ou 'Du 01 au 31/12/2026') -> (début, fin, 'DECEMBRE 2026').
    Le mois et l'année affichés sont ceux de la date de fin."""
    m = re.search(r"(\d{1,2})(?:/(\d{1,2}))?(?:/(\d{4}))?\s+(?:au|a|à)\s+(\d{1,2})/(\d{1,2})/(\d{4})",
                  texte, flags=re.IGNORECASE)
    if m:
        j1, m1, a1, j2, m2, a2 = m.groups()
        m1, a1 = m1 or m2, a1 or a2
        du = f"{int(j1):02d}/{int(m1):02d}/{a1}"
        au = f"{int(j2):02d}/{int(m2):02d}/{a2}"
    else:
        print('Attention : période non trouvée dans l\'en-tête ("Du jj/mm/aaaa au jj/mm/aaaa"), '
              "valeurs par défaut utilisées.")
        du, au = PERIODE_DU_DEFAUT, PERIODE_AU_DEFAUT
        _, m2, a2 = au.split("/")
    mois = int(m2)
    if not 1 <= mois <= 12:
        sys.exit(f"Mois invalide dans la période : {texte!r}")
    return du, au, f"{PREFIXE_PERIODE}{MOIS[mois - 1]} {a2}"


def trouver_logos_clients():
    """Cherche image_1, image_2, ... (.jpeg / .jpg / .png) dans le dossier du script ou dans logos/."""
    trouves = []
    for n in range(1, 11):
        for dossier in (DOSSIER_SCRIPT, os.path.join(DOSSIER_SCRIPT, "logos")):
            fichier = next((os.path.join(dossier, f"image_{n}{ext}") for ext in (".jpeg", ".jpg", ".png", ".JPEG", ".JPG", ".PNG")
                            if os.path.exists(os.path.join(dossier, f"image_{n}{ext}"))), None)
            if fichier:
                trouves.append(fichier)
                break
    return trouves


# ----------------------------- Téléchargement des photos -----------------------------
NAVIGATEUR = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
MOTS_A_EVITER = ("logo", "icon", "favicon", "sprite", "avatar", "pixel", "blank", "placeholder",
                 "loader", "spinner", "facebook", "twitter", "instagram", "whatsapp", "linkedin")
EXT_IMAGES = r"(?:jpe?g|png|webp|gif|avif|bmp)"


def nom_fichier(code):
    return re.sub(r"[^\w\-]", "_", code)


def telecharger(url, referer=None):
    """Télécharge une URL -> (octets, type_contenu, url_finale)."""
    url = urllib.parse.quote(url, safe=":/?&=%#+@,;~!$'()*[]")
    entetes = {"User-Agent": NAVIGATEUR, "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
               "Accept": "image/avif,image/webp,image/*,text/html,*/*;q=0.8"}
    if referer:
        entetes["Referer"] = referer
    contexte = None if VERIFIER_SSL else ssl._create_unverified_context()
    req = urllib.request.Request(url, headers=entetes)
    with urllib.request.urlopen(req, timeout=30, context=contexte) as r:
        return r.read(40 * 1024 * 1024), r.headers.get("Content-Type", ""), r.geturl()


def ouvrir_image(donnees):
    try:
        im = Image.open(io.BytesIO(donnees))
        im.load()
        return im
    except Exception:
        return None


class ExtracteurImages(HTMLParser):
    """Repère les images d'une page web : (url, score de base, texte associé)."""
    def __init__(self):
        super().__init__()
        self.candidats = []

    def handle_starttag(self, tag, attrs):
        a = {k.lower(): (v or "") for k, v in attrs}
        if tag == "meta":
            cle = (a.get("property") or a.get("name") or "").lower()
            if cle.startswith(("og:image", "twitter:image")) and a.get("content"):
                self.candidats.append((a["content"], 5, ""))
        elif tag == "link" and "image_src" in a.get("rel", "").lower() and a.get("href"):
            self.candidats.append((a["href"], 5, ""))
        elif tag == "img":
            src = a.get("data-src") or a.get("data-original") or a.get("src") or ""
            if not src and a.get("srcset"):
                src = a["srcset"].split(",")[0].split()[0]
            if src:
                self.candidats.append((src, 0, a.get("alt", "") + " " + a.get("title", "")))
        elif tag == "a" and re.search(r"\." + EXT_IMAGES + r"(\?|$)", a.get("href", ""), re.IGNORECASE):
            self.candidats.append((a["href"], 1, ""))


def liste_candidats(html, url_page, code):
    """Toutes les images possibles de la page, triées de la plus probable à la moins probable."""
    parseur = ExtracteurImages()
    try:
        parseur.feed(html)
    except Exception:
        pass
    brut = list(parseur.candidats)
    brut += [(m, -1, "") for m in re.findall(r"[\w\-./:%@+]+\." + EXT_IMAGES + r"(?:\?[\w=&%.\-]*)?", html, re.IGNORECASE)]
    m = re.search(r"zupimages\.net/viewer\.php\?id=([\w/.\-]+)", url_page)
    if m:                                                    # visionneuse zupimages -> image réelle
        brut.insert(0, (f"https://zupimages.net/up/{m.group(1)}", 12, ""))

    code_norm = re.sub(r"[^a-z0-9]", "", code.lower())
    vus, resultat = set(), []
    for url, base, texte in brut:
        url = url.strip()
        if not url or url.startswith(("data:", "javascript:", "#")) or url.lower().split("?")[0].endswith(".svg"):
            continue
        absolue = urllib.parse.urljoin(url_page, url)
        if absolue in vus:
            continue
        vus.add(absolue)
        bas = (absolue + " " + texte).lower()
        score = base
        if code_norm and code_norm in re.sub(r"[^a-z0-9]", "", bas):
            score += 10
        if any(mot in bas for mot in MOTS_A_EVITER):
            score -= 10
        resultat.append((score, absolue))
    resultat.sort(key=lambda t: -t[0])                       # tri stable : l'ordre de la page est gardé
    return resultat


def chercher_image(lien, code):
    """Retourne l'image PIL trouvée à ce lien (image directe ou page web contenant l'image)."""
    donnees, type_contenu, url_finale = telecharger(lien)
    im = ouvrir_image(donnees)
    if im:                                                   # le lien est directement une image
        return im
    charset = re.search(r"charset=([\w\-]+)", type_contenu)
    html = donnees.decode(charset.group(1) if charset else "utf-8", errors="replace")

    valides = []
    for score, url in liste_candidats(html, url_finale, code)[:15]:
        try:
            d, _, _ = telecharger(url, referer=url_finale)
        except Exception:
            continue
        img = ouvrir_image(d)
        if not img:
            continue
        if score >= 10 and min(img.size) >= 200:             # image portant le code de l'emplacement
            return img
        valides.append((score, img.size[0] * img.size[1], img))
    grandes = [v for v in valides if v[2].size[0] >= 300 and v[0] > -5]
    if grandes:
        return max(grandes, key=lambda v: (v[0] >= 0, v[1]))[2]

    dossier_debug = os.path.join(DOSSIER_SCRIPT, "debug_html")
    os.makedirs(dossier_debug, exist_ok=True)
    with open(os.path.join(dossier_debug, nom_fichier(code) + ".html"), "w", encoding="utf-8") as f:
        f.write(html)
    raise RuntimeError(f"aucune photo trouvée dans la page (page enregistrée dans debug_html/{nom_fichier(code)}.html)")


def enregistrer_jpeg(im, chemin):
    """Convertit en JPEG (RGB, fond blanc pour la transparence, 2400 px max)."""
    im = ImageOps.exif_transpose(im)
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        im = im.convert("RGBA")
        fond = Image.new("RGB", im.size, (255, 255, 255))
        fond.paste(im, mask=im.split()[-1])
        im = fond
    else:
        im = im.convert("RGB")
    if max(im.size) > 2400:
        im.thumbnail((2400, 2400), Image.LANCZOS)
    im.save(chemin, "JPEG", quality=90, optimize=True)
    return im.size


def telecharger_images(lignes):
    dossier = os.path.join(DOSSIER_SCRIPT, DOSSIER_IMAGES)
    os.makedirs(dossier, exist_ok=True)
    fichier_sources = os.path.join(dossier, "sources.json")
    try:
        with open(fichier_sources, encoding="utf-8") as f:
            sources = json.load(f)
    except Exception:
        sources = {}

    echecs = []
    print("Photos :")
    for d in lignes:
        code, lien = d["code"], d.get("lien")
        chemin = os.path.join(dossier, nom_fichier(code) + ".jpg")
        deja = os.path.exists(chemin)
        if deja and not FORCER_TELECHARGEMENT and (not lien or sources.get(code, lien) == lien):
            print(f"  {code:8} déjà présente ({os.path.basename(chemin)})")
            continue
        if not lien:
            echecs.append((code, "aucun lien dans la cellule Code"))
            continue
        try:
            taille = enregistrer_jpeg(chercher_image(lien, code), chemin)
            sources[code] = lien
            print(f"  {code:8} OK  {taille[0]}x{taille[1]}")
        except urllib.error.HTTPError as e:
            echecs.append((code, f"erreur HTTP {e.code} sur {lien}"))
        except urllib.error.URLError as e:
            echecs.append((code, f"connexion impossible ({e.reason})"))
        except Exception as e:
            echecs.append((code, str(e)))
    with open(fichier_sources, "w", encoding="utf-8") as f:
        json.dump(sources, f, ensure_ascii=False, indent=1)
    for code, raison in echecs:
        print(f"  {code:8} ECHEC : {raison}")
    if echecs:
        print("  -> Vous pouvez déposer la photo à la main : "
              f"{DOSSIER_IMAGES}/CODE.jpg (elle sera utilisée telle quelle).")


# ----------------------------- PowerPoint -----------------------------
def zone_texte(slide, x, y, w, h, paragraphes, taille, gras=True, couleur=NOIR,
               police=POLICE, ancrage=MSO_ANCHOR.MIDDLE):
    """paragraphes : liste de textes, ou de listes de (texte, couleur) pour mélanger les couleurs."""
    box = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = ancrage
    for i, para in enumerate(paragraphes):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        morceaux = para if isinstance(para, list) else [(para, couleur)]
        for texte, coul in morceaux:
            run = p.add_run()
            run.text = texte
            run.font.size = Pt(taille)
            run.font.bold = gras
            run.font.color.rgb = coul
            if police:
                run.font.name = police
    return box


def numero_diapo(slide):
    """Numéro de diapo automatique, en gris, en bas à droite."""
    box = slide.shapes.add_textbox(Emu(6553200), Emu(6356349), Emu(2133600), Emu(365125))
    box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    fld = etree.SubElement(p._p, f"{{{ns}}}fld", id="{C0C85C69-8E57-4F0C-AFB6-D18482E6D58F}", type="slidenum")
    rpr = etree.SubElement(fld, f"{{{ns}}}rPr", lang="fr-FR", sz="1200")
    fill = etree.SubElement(rpr, f"{{{ns}}}solidFill")
    etree.SubElement(fill, f"{{{ns}}}srgbClr", val="8B8B8B")
    etree.SubElement(fld, f"{{{ns}}}t").text = "‹#›"


def image_ajustee(slide, source, x, y, w, h, remplir=False):
    """Insère une image dans le cadre (x, y, w, h). remplir=False : image entière ; True : cadre rempli (rogne)."""
    pic = slide.shapes.add_picture(source, Emu(x), Emu(y))
    if not remplir:
        ech = min(w / pic.width, h / pic.height)
        pic.width, pic.height = int(pic.width * ech), int(pic.height * ech)
        pic.left = x + (w - pic.width) // 2
        pic.top = y + (h - pic.height) // 2
    else:
        ech = max(w / pic.width, h / pic.height)
        sw, sh = pic.width * ech, pic.height * ech
        pic.crop_left = pic.crop_right = max(0.0, (sw - w) / sw / 2)
        pic.crop_top = pic.crop_bottom = max(0.0, (sh - h) / sh / 2)
        pic.left, pic.top, pic.width, pic.height = x, y, w, h
    return pic


def en_tete(slide, logo, logos_clients, compact=False):
    """Logo (haut gauche) + logos clients alignés en haut à droite.
    compact=True : zone réduite, à droite du titre "Ligne 0N" (diapos de ligne)."""
    if logo:
        slide.shapes.add_picture(io.BytesIO(logo), Emu(285720), Emu(142852), Emu(2362200), Emu(1219200))
    n = len(logos_clients)
    if n:
        zone_x, haut = (5300000, 1350000) if compact else (3571868, 1450000)
        cellule = min(1857388, (9144000 - zone_x) // n)
        debut = 9144000 - cellule * n                     # groupe aligné à droite
        for i, fichier in enumerate(logos_clients):
            image_ajustee(slide, fichier, debut + i * cellule, 0, cellule, haut)


def diapo_titre(prs, logo, logos_clients, lignes_texte):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    en_tete(s, logo, logos_clients)
    n = len(lignes_texte)
    taille = 40 if n <= 4 else 34 if n == 5 else 30 if n == 6 else 26
    zone_texte(s, 685800, 2130425, 7772400, 3013087, lignes_texte, taille, couleur=ROUGE)
    numero_diapo(s)


def diapo_ligne(prs, logo, logos_clients, d, du, au):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    en_tete(s, logo, logos_clients if LOGOS_SUR_LIGNES else [], compact=True)
    zone_texte(s, 3643306, 500042, 2143140, 714380, [f"Ligne {d['numero']:02d}"], 16, couleur=ROUGE)
    zone_texte(s, 214282, 1500174, 1357290, 500066, [d["code"]], 16, couleur=ROUGE)
    zone_texte(s, 214282, 2357430, 1357290, 714380,
               ["Dimensions",
                f"{nb(d['largeur'])}m X {nb(d['hauteur'])}m",
                [(f"{nb(d['surface'])} m²", ROUGE)]], 14)
    zone_texte(s, 214282, 3429000, 1357290, 700000, [TEXTE_PHOTO], 14, couleur=ROUGE)
    zone_texte(s, 214282, 5214950, 1357290, 900000, ["Du", f"{du} au", au], 14)

    desc = f"{d['code']}, {d['emplacement']}"
    taille = 12 if len(desc) <= 150 else 11 if len(desc) <= 190 else 10
    zone_texte(s, 2285984, 5929330, 5143536, 600000, [desc], taille, gras=False, police=None)

    # photo : insérée si images/CODE.jpg existe (téléchargée ou déposée à la main)
    if DOSSIER_IMAGES:
        for ext in (".jpg", ".jpeg", ".png"):
            chemin = os.path.join(DOSSIER_SCRIPT, DOSSIER_IMAGES, nom_fichier(d["code"]) + ext)
            if os.path.exists(chemin):
                image_ajustee(s, chemin, 1785918, 1500174, 7020000, 4320000, remplir=(MODE_PHOTO == "remplir"))
                break
    numero_diapo(s)


def main():
    args = lire_arguments()
    chemin_excel = trouver_excel(args.excel)
    print("Excel utilisé :", os.path.basename(chemin_excel))
    lignes, texte_periode, logo = lire_excel(chemin_excel)
    if SEULEMENT_DISPONIBLES:
        lignes = [l for l in lignes if l["dispo"].upper().startswith("DISPONIBLE")]
    if not lignes:
        sys.exit("Aucune ligne de données trouvée.")
    du, au, libelle = lire_periode(texte_periode)
    if TELECHARGER_IMAGES:
        telecharger_images(lignes)
    logos_clients = trouver_logos_clients()
    print(f"{len(lignes)} ligne(s) lue(s) - période : {du} au {au} ({libelle})")
    print(f"{len(logos_clients)} logo(s) client trouvé(s)" +
          (" : " + ", ".join(os.path.basename(f) for f in logos_clients) if logos_clients else ""))
    if not logo:
        print("Attention : aucun logo trouvé (ni dans l'Excel, ni dans logo.png).")

    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(9144000), Emu(6858000)    # format 4:3

    diapo_titre(prs, logo, logos_clients, list(TITRE_LIGNES) + ["", libelle])
    for d in lignes:
        diapo_ligne(prs, logo, logos_clients, d, du, au)
    diapo_titre(prs, logo, logos_clients, [TEXTE_FIN])

    if args.sortie:
        sortie = args.sortie if args.sortie.lower().endswith(".pptx") else args.sortie + ".pptx"
    else:
        base = os.path.splitext(os.path.basename(chemin_excel))[0].replace(" ", "_")
        sortie = os.path.join(os.path.dirname(os.path.abspath(chemin_excel)), f"PRESENTATION_{base}.pptx")
    try:
        prs.save(sortie)
    except PermissionError:
        sys.exit(f"Impossible d'enregistrer : fermez {os.path.basename(sortie)} dans PowerPoint puis relancez.")
    print("Fichier créé :", sortie)


if __name__ == "__main__":
    main()