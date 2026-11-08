# Generateur de presentation d'affichage

Ce script cree un PowerPoint a partir d'un planning Excel. Il genere une diapositive de couverture, une diapositive par emplacement, puis une diapositive de fin.

## Prerequis

- Un fichier Excel `.xlsx` contenant le planning
- Une connexion Internet si le telechargement automatique des photos est active

### Installer et verifier Python (Windows)

1. Telechargez l'installateur Windows depuis la [page officielle des versions de Python](https://www.python.org/downloads/windows/).
2. Lancez le fichier `.exe` telecharge.
3. Dans l'installateur, cochez **Add Python to PATH**, puis cliquez sur **Install Now**. L'installation pour tous les utilisateurs est facultative.
4. Ouvrez PowerShell et verifiez l'installation :

```powershell
python --version
python -m pip --version
```

Les deux commandes doivent afficher une version. Si `python` n'est pas reconnu, relancez l'installateur, cochez **Add Python to PATH**, puis ouvrez une nouvelle fenetre PowerShell.

Depuis PowerShell, ouvrez le dossier du projet puis installez les dependances :

```powershell
python -m pip install openpyxl pillow lxml python-pptx
```

## Utilisation

Placez le fichier Excel dans le dossier du projet, puis lancez le script avec le nom du fichier :

```powershell
python .\generer_pptx.py "PLANNING AFFICHAGE DECEMBRE 2026.xlsx"
```

Le PowerPoint est cree a cote du fichier Excel. Son nom commence par `PRESENTATION_`.

On peut aussi choisir le nom du PowerPoint :

```powershell
python .\generer_pptx.py "mon_planning.xlsx" -o "ma_presentation.pptx"
```

Si aucun fichier Excel n'est indique, le script prend le fichier `.xlsx` le plus recent dans le dossier du script :

```powershell
python .\generer_pptx.py
```

Pour afficher les options :

```powershell
python .\generer_pptx.py --help
```

## Format du fichier Excel

Le script recherche la ligne d'en-tetes ou la cellule de la colonne B contient `Nombre`. Cette ligne doit contenir les en-tetes suivants :

- `Nombre`
- `Code`
- `Emplacement`
- `Largeur / m`
- `Hauteur / m`

La derniere colonne de l'en-tete doit contenir la periode, par exemple `Du 01/12/2026 au 31/12/2026`. Le script lit les lignes d'emplacements jusqu'a la premiere cellule `Code` vide. Pour le telechargement automatique des photos, ajoutez un lien vers la photo ou sa page Web dans la cellule `Code`.

## Photos et logos

Le telechargement des photos est active par defaut. Les photos trouvees sont converties en JPEG et enregistrees dans `images/` sous la forme `CODE.jpg`. Le fichier `images/sources.json` est cree ou mis a jour par le script. Les images peuvent aussi etre ajoutees manuellement dans `images/` sous le nom du code, par exemple `images/UV09.jpg`.

Les logos sont facultatifs :

- `logo.png` dans le dossier du script pour le logo principal; a defaut, le premier logo integre au fichier Excel est utilise.
- `image_1.jpeg`, `image_2.jpeg`, etc. dans le dossier du script ou dans `logos/` pour les logos clients. Les extensions `.jpg` et `.png` sont aussi acceptees.

Les options de mise en page, les textes, la periode par defaut et le telechargement des images se reglent au debut de `generer_pptx.py`.

## Fichiers a publier sur GitHub

Pour partager le programme, publiez au minimum `generer_pptx.py`. Comme ce guide explique l'installation et l'utilisation, ajoutez aussi `README.md`. Les dependances peuvent etre documentees ici; un fichier `requirements.txt` est facultatif.

Ne publiez pas le planning Excel, le PowerPoint genere, les logos ou les photos s'ils contiennent des informations confidentielles ou si vous n'avez pas les droits de les partager. Ces fichiers restent locaux et sont fournis au script au moment de son execution.
