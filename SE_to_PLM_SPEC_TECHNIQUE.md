# SE_to_PLM — Spécification technique complète

> Document destiné à une IA. Contient la totalité des fonctionnalités, leur implémentation précise, le but du projet et le résultat attendu.

---

## 1. But du projet

Extraire automatiquement la structure complète d'un projet Solid Edge (assemblage `.asm`, sous-assemblages, pièces `.par`/`.psm`) et générer un fichier Excel (`.xlsx`) prêt à être importé dans un système PLM (Product Lifecycle Management).

L'outil fait le lien entre :
- les **fichiers 3D** (`.asm`, `.par`, `.psm`) et leur hiérarchie d'assemblage
- les **plans 2D** associés (`.dft`) retrouvés par scan du système de fichiers
- les **métadonnées** stockées dans les propriétés Solid Edge (désignation, révision, version, auteur, dates, matière, densité)

---

## 2. Architecture du projet

```
SE_to_PLM/
├── se_to_plm.py         # Moteur métier pur (pas d'UI)
├── gui_se_to_plm.py     # Interface graphique PyQt5 — appelle le moteur
├── style.qss            # Feuille de style Qt (thème visuel)
├── RUN.vbs              # Lanceur Windows (venv auto + démarrage silencieux)
└── requirements.txt     # pywin32, openpyxl, PyQt5
```

**Séparation stricte** : `se_to_plm.py` est le moteur sans UI. `gui_se_to_plm.py` l'importe et l'appelle via la fonction `generer_export_excel()`. Le mode CLI (`se_to_plm.py` exécuté directement) utilise le même moteur.

---

## 3. Fichier `RUN.vbs` — Lanceur automatique

- Détecte le dossier courant depuis `WScript.ScriptFullName`
- Si le dossier `venv` n'existe **pas** :
  - Affiche une notification Windows (bulle système tray via PowerShell + `System.Windows.Forms.NotifyIcon`) : "Création de l'environnement virtuel..."
  - Exécute `python -m venv venv` (bloquant)
  - Exécute `venv\Scripts\pip.exe install -r requirements.txt` (bloquant)
  - Affiche une notification "Installation terminée !"
- Lance `venv\Scripts\pythonw.exe gui_se_to_plm.py` (non bloquant, sans fenêtre console)

---

## 4. Fichier `style.qss` — Thème visuel

Feuille de style Qt appliquée globalement à la fenêtre principale :

| Widget | Style appliqué |
|--------|---------------|
| `QMainWindow` | Fond `#f0f2f5` (gris clair) |
| `QWidget` | Police Segoe UI 10pt, couleur `#333333` |
| `QLabel` | Gras 500, couleur `#333333` |
| `QLineEdit` | Fond blanc, bordure `#cccccc`, rayon 4px, focus en bleu `#0078d4` |
| `QLineEdit` read-only | Fond `#f5f5f5`, texte `#666666` |
| `QPushButton` | Fond bleu `#0078d4`, blanc, états hover/pressed/disabled |
| `QPushButton` disabled | Fond `#cccccc`, texte `#666666` |
| `QTextEdit` (console) | Thème dark `#1e1e1e`, texte `#d4d4d4`, police Consolas 9pt |

---

## 5. Fichier `se_to_plm.py` — Moteur métier

### 5.1 Imports et dépendances

```python
win32com.client      # COM Solid Edge (pywin32)
openpyxl             # Génération Excel
pythoncom            # Gestion COM multi-thread
tkinter.filedialog   # Sélecteur fichier (mode CLI uniquement)
concurrent.futures   # ThreadPoolExecutor pour scan parallèle
threading            # Lock pour cache thread-safe
```

---

### 5.2 `demander_fichier_asm()`

- Mode CLI uniquement
- Ouvre une boîte `filedialog.askopenfilename` filtrée sur `*.asm`
- Retourne le chemin sélectionné ou chaîne vide si annulé

---

### 5.3 Scan parallèle des plans `.dft`

#### `_scanner_dossier_thread_safe(args)`

Worker unitaire pour `ThreadPoolExecutor`. Reçoit `(chemin_dossier, depth, max_depth)`.

- Utilise `os.scandir()` pour lister le contenu du dossier
- Pour chaque entrée :
  - Si fichier `.dft` → ajoute `(nom_base_lowercase, chemin_absolu)` à `plans_trouves`
  - Si dossier ET `depth < max_depth` → ajoute à `sous_dossiers` pour traitement ultérieur
- Ignore les erreurs `PermissionError` et `OSError`
- Retourne `(plans_trouves, sous_dossiers, dossiers_count=1)`

#### `indexer_les_plans_projet_entier(chemin_asm_initial, dossier_dft, mode_recherche, max_depth, callback_progress)`

**Rôle** : Construire un dictionnaire `{nom_base_lowercase: chemin_absolu}` de tous les `.dft` trouvés.

**Modes de recherche** (`mode_recherche`) :
- `"arborescence"` : remonte 3 niveaux depuis le dossier du `.asm` pour trouver la racine projet, puis scan récursif
- `"dossier_specifique"` : scan uniquement le dossier `dossier_dft` passé en paramètre
- `"les_deux"` (défaut) : combine les deux stratégies ci-dessus

**Algorithme de scan parallèle** :
1. Initialise une pile (`pile`) avec les dossiers de départ
2. Boucle principale : extrait des batches de 16 dossiers depuis la pile
3. Soumet chaque batch en parallèle via `ThreadPoolExecutor(max_workers=8)`
4. Collecte les résultats via `as_completed()` :
   - Ajoute les plans trouvés au dict `index`
   - Repousse les sous-dossiers dans la pile
5. Appelle `callback_progress(scanned, found, total)` tous les 50 dossiers
6. Affiche un log tous les 100 dossiers

**Profondeur max** : `max_depth=3` par défaut (configurable)

---

### 5.4 `MetadataCache` — Cache LRU thread-safe

Cache en mémoire pour éviter les lectures COM répétées sur les mêmes fichiers.

| Méthode | Comportement |
|---------|-------------|
| `get(key)` | Retourne la valeur si présente, met à jour l'ordre LRU, retourne `None` sinon |
| `set(key, value)` | Ajoute/met à jour. Si `len >= max_size` (1000), évince le plus ancien (LRU) |
| `clear()` | Vide le cache et la liste d'ordre |

Toutes les opérations sont protégées par `threading.Lock()`.

Instance globale : `g_metadata_cache = MetadataCache()`

---

### 5.5 `lister_proprietes(doc_obj)` — Débogage COM

Fonction de debug uniquement. Liste :
- Le type de l'objet COM
- Les 20 premiers attributs Python de l'objet
- Les `PropertySets` disponibles avec leurs propriétés Custom
- Le `SummaryInformation.Title` si disponible

---

### 5.6 `normaliser_date(valeur)`

Convertit n'importe quelle représentation de date textuelle vers le format PLM cible : `JJ/MM/AAAA 12:00:00 AM`.

Formats supportés (du plus au moins spécifique) :
```
%d/%m/%Y %H:%M:%S    %d/%m/%Y %H:%M    %d-%m-%Y %H:%M:%S    %d-%m-%Y %H:%M
%Y-%m-%d %H:%M:%S    %Y-%m-%d %H:%M    %Y/%m/%d %H:%M:%S    %Y/%m/%d %H:%M
%d/%m/%Y    %d-%m-%Y    %d.%m.%Y    %Y-%m-%d    %Y/%m/%d
%d/%m/%y    %d-%m-%y    %d.%m.%y
```

Fallback : si aucun format ne correspond, tente de normaliser les séparateurs mixtes (`/`, `-`, `.`) puis réessaie. Si toujours impossible, retourne la valeur d'origine inchangée.

---

### 5.7 `extraire_metadonnees_rapide(chemin_fichier, debug, use_cache)`

**Méthode d'extraction SANS ouvrir Solid Edge** (ultra rapide).

Utilise `win32com.client.Dispatch("SolidEdge.FileProperties")` qui lit les métadonnées directement dans le fichier binaire, sans lancer l'interface graphique.

**Séquence d'extraction** :

1. Vérifie le cache (`g_metadata_cache.get(chemin_fichier)`)
2. Ouvre via `prop_reader.Open(chemin_fichier)`
3. Extrait depuis `ExtendedSummaryInformation` → cherche la propriété `"Username"` → `meta["auteur"]`
4. Extrait depuis `Custom` :
   - **Accès direct par nom** (sensible à la casse) pour : `"auteur modif"`, `"date modif"`, `"Matière"`, `"Densité"`
   - **Itération par index** pour les autres champs (insensible à la casse) :
     - `"désignation"` / `"designation"` / `"desig"` → `meta["designation"]`
     - `"date de création"` / `"date de creation"` → `meta["date_creation"]` (normalisée)
     - `"indice de modification"` / `"revision index"` / `"index"` / `"revision"` / `"rev"` → `meta["version"]`
     - `"matière"` / `"matiere"` → `meta["matiere"]` (si pas déjà rempli par accès direct)
     - `"densité"` / `"densite"` → `meta["densite"]` (si pas déjà rempli)
5. `prop_reader.Close()`
6. Met en cache et retourne

**Valeurs par défaut** si champ absent :
```python
{"designation": "", "revision": "1", "version": "-",
 "auteur": "", "date_creation": "", "auteur_modif": "",
 "date_modif": "", "matiere": "", "densite": ""}
```

---

### 5.8 `extraire_metadonnees(doc_obj, debug, use_cache)`

Variante pour extraction depuis un **objet document COM déjà ouvert** dans Solid Edge (utilisé pour le document racine uniquement).

Même logique que `extraire_metadonnees_rapide` mais itère via `doc_obj.Properties` (PropertySets COM) au lieu de `SolidEdge.FileProperties`.

La clé de cache est `doc_obj.FullName`.

---

### 5.9 `calculer_indices_precedents(version)`

Calcule les deux indices précédents (n-1 et n-2) depuis une version alphabétique courante.

**Logique** :
- `"-"` ou vide → retourne `("-", "-")`
- Lettre simple : `"A"` → `"-"`, `"B"` → `"A"`, `"Z"` → `"Y"`
- Double lettre :
  - `"AA"` → `"Z"` (le préfixe "A" se réduit à "-", donc "Z")
  - `"AB"` → `"AA"`
  - `"AZ"` → `"AY"`
  - `"BA"` → `"AZ"`

Exemples d'appel : `"C"` → `("B", "A")`, `"A"` → `("-", "-")`, `"AA"` → `("Z", "-")`

---

### 5.10 `determiner_classe(nom_fichier, chemin_complet, est_projet)`

Détermine la classe PLM d'un fichier :

| Condition | Classe retournée |
|-----------|-----------------|
| `est_projet=True` | `"SUB_ASSY_A"` |
| Extension `.asm` | `"SUB_ASSY_A"` |
| Extension `.par` ou `.psm` ET chemin contient `"bibliothèque"`, `"bibliotheque"` ou `"library"` (à n'importe quel niveau, insensible à la casse) | `"PART_PURCH_A"` |
| Extension `.par` ou `.psm` (hors bibliothèque) | `"PART_A"` |
| Extension `.dft` | `"CAD_DRAWING_A"` |
| Autre | `"Folder"` |

---

### 5.11 `generer_export_excel(...)` — Fonction moteur principale

**Signature complète** :
```python
generer_export_excel(
    chemin_fichier,       # str: chemin du fichier principal (.asm, .par, .psm)
    dossier_sortie,       # str: dossier de destination du .xlsx
    nom_sortie,           # str: nom du fichier Excel (avec ou sans .xlsx)
    dossier_dft=None,     # str|None: dossier spécifique pour plans
    mode_recherche="les_deux",  # str: mode scan DFT
    type_fichier="asm",   # str: "asm", "pieces", ou "les_deux"
    callback_log=None,    # callable(message, type): type in ['info','success','error','warning']
    callback_progress=None, # callable(valeur, maximum, message)
    check_cancelled=None  # callable() -> bool: True pour annuler
)
```

**Retourne** : `{'chemin_fichier': str, 'stats': {'3d': int, '2d': int}}`

---

#### Séquence d'exécution détaillée :

**Étape 1 — Indexation des plans (0 → 10%)**
- Appelle `indexer_les_plans_projet_entier()` avec callback de progression
- Construit `index_plans = {nom_base_lowercase: chemin_absolu}`

**Étape 2 — Connexion Solid Edge (10%)**
- Tente `win32com.client.GetActiveObject("SolidEdge.Application")` → réutilise une instance déjà ouverte
- Si échec `com_error` → `win32com.client.dynamic.Dispatch("SolidEdge.Application")` + `app.Visible = False`
- `app.DisplayAlerts = False`

**Étape 3 — Ouverture du document racine**
- Vérifie si le document est déjà ouvert parmi `app.Documents` (comparaison `os.path.normcase` sur `FullName`)
- Si déjà ouvert → réutilise l'objet, sinon → `app.Documents.Open(chemin_fichier)`

**Étape 4 — Fonctions internes**

##### `get_suffixe_fichier(nom_fichier)`
Retourne le suffixe textuel de l'attachement : `"(ASM)"`, `"(PRT)"`, `"(DRW)"` ou `""`

##### `normaliser_chemin_reseau(chemin)`
- Si chemin UNC (`\\`) avec adresse IP comme serveur (test `.replace('.','').isdigit()`) → ne normalise pas (évite résolution DNS), corrige juste les slashs
- Sinon → `os.path.normpath(chemin)`

##### `ajouter_ligne(niveau, relation, nom_fichier, chemin_complet, classe, qte, rev, desig, ver, auteur, date_crea, auteur_modif, date_modif, matiere, densite)`
- Calcule `ref_util = os.path.splitext(nom_fichier)[0]`
- Calcule `attachement = chemin_normalise + suffixe`
- Appelle `calculer_indices_precedents(ver)` → `indice_1`, `indice_2`
- Appende dans `lignes_excel` la liste de 21 valeurs dans l'ordre des colonnes Excel
- Incrémente `compteur_ordre`

##### `explorer_occurrences(occurrences, niveau)`

Explore récursivement les occurrences d'un assemblage Solid Edge.

1. **Déduplique** les occurrences par nom de fichier dans `dict_occ = {nom: {qte, obj, chemin}}` :
   - Tente `occ.OccurrenceDocument.FullName` (méthode principale)
   - Fallback 1 : `occ.OccurrenceFileName`
   - Fallback 2 : `occ.FileName`
   - Fallback 3 : `occ.Name.split(':')[0]` (sans chemin)
   - Si même fichier déjà présent → incrémente la quantité (`qte += 1`)

2. Pour chaque entrée unique :
   - Détermine la classe via `determiner_classe()`
   - Extrait les métadonnées via `extraire_metadonnees_rapide()`
   - Appelle `ajouter_ligne()` avec `relation="ComposedOf"`
   - Cherche un plan associé dans `index_plans` (clé = `nom_sans_ext.lower()`)
   - Si plan trouvé → l'ajoute dans `liste_plans_a_rajouter`
   - Si c'est un sous-assemblage (`occ.Subassembly == True`) → appel récursif `explorer_occurrences(occ.OccurrenceDocument.Occurrences, niveau + 1)`

**Étape 5 — Ajout du nœud racine (niveau 0)**
- Extrait les métadonnées du document racine via `extraire_metadonnees(doc_racine)`
- Détermine la classe selon `type_fichier` :
  - `"asm"` → force `"SUB_ASSY_A"`
  - `"pieces"` → appelle `determiner_classe()`
  - `"les_deux"` → logique hybride selon l'extension
- Appelle `ajouter_ligne()` avec `relation=""`
- Cherche un plan `.dft` pour la racine dans `index_plans`
- Si `.asm` → appelle `explorer_occurrences(doc_racine.Occurrences, 1)`

**Étape 6 — Extraction métadonnées des plans (50 → 70%)**

Pour chaque plan dans `liste_plans_a_rajouter` (dédupliqués par chemin dans `plans_uniques`) :

1. `extraire_metadonnees_rapide(dft_path)` → `meta_dft`
2. `extraire_metadonnees_rapide(src_path)` → `meta_piece`
3. **Héritage auteur/date** : si les champs `auteur` ou `date_creation` du DFT sont vides, copie depuis la pièce source (avec log)
4. **Ligne DFT** : niveau=0, relation="", classe=`"CAD_DRAWING_A"`, désignation = celle de la pièce source, matière/densité = depuis la pièce source
5. **Ligne pièce liée** : niveau=1, relation="Drawing", classe=celle de la pièce, données de la pièce

Chaque paire DFT/pièce n'est ajoutée **qu'une seule fois** (set `plans_deja_traites`).

**Étape 7 — Génération Excel (75 → 100%)**

- Crée un `openpyxl.Workbook()`, onglet renommé `"Structure"`
- Écrit la ligne d'en-têtes :
```
Level | Relationship | ordre | quantite | repere | SpecialCAD | Class |
ref_utilisat | version | indice_1 | indice_2 | revision | designation |
cus_createur | cus_date_crea | user_version_1 | date_version_1 |
matiere | densite | dia_se | Attachments
```
- Formatage en-têtes :
  - Colonnes 1–12 : fond vert clair `#CCFFCC`, gras, alignement gauche
  - Colonnes 13+ : fond orange `#FFC000`, gras, alignement gauche
- Écrit toutes les lignes de données
- **Calcul largeur colonnes** en parallèle via `ThreadPoolExecutor(max_workers=4)` : pour chaque colonne, calcule `max(len(str(cell.value))) + 2`
- Ajoute `.xlsx` si absent dans le nom
- Sauvegarde via `wb.save(chemin_complet)`

---

### 5.12 `lancer_extraction_plm()` — Point d'entrée CLI

- Appelle `demander_fichier_asm()` pour obtenir le chemin
- Dossier de sortie = dossier du fichier ASM
- Nom de sortie = `Export_PLM_{timestamp}.xlsx`
- Appelle `generer_export_excel()` sans callbacks (logs sur stdout)
- Affiche les stats finales (`3D | Plans`)

---

## 6. Fichier `gui_se_to_plm.py` — Interface graphique

### 6.1 `ExtractionThread(QThread)`

Thread de travail qui exécute `generer_export_excel()` sans bloquer l'UI.

**Signaux PyQt5** :
```python
log_signal      = pyqtSignal(str, str)        # (message, type)
finished_signal = pyqtSignal()                # fin (succès ou erreur)
progress_signal = pyqtSignal(int, int, str)   # (valeur, maximum, message)
```

**Méthode `cancel()`** : positionne `self._cancelled = True` → transmis au moteur via `check_cancelled()`

**Méthode `run()`** :
- Définit les 3 callbacks (log, progress, cancelled) qui émettent les signaux
- Appelle `generer_export_excel()` avec ces callbacks
- Émet `finished_signal` dans le bloc `finally` (même en cas d'erreur)

---

### 6.2 `PLMExtractorGUI(QMainWindow)`

#### Construction de l'interface (`creer_interface()`)

Widgets de haut en bas :

| Widget | Rôle |
|--------|------|
| `QComboBox` type_fichier_combo | 3 options : "Assemblage (.asm)", "Pièces (.par/.psm)", "Les deux" |
| `QLabel` + `QLineEdit` + `QPushButton` | Sélection du fichier principal (read-only, bouton Parcourir) |
| `QComboBox` mode_dft_combo | 3 options : "Arborescence uniquement", "Dossier spécifique uniquement", "Les deux" |
| `QLabel` + `QLineEdit` + `QPushButton` | Sélection dossier DFT (activé/désactivé selon mode) |
| `QLineEdit` nom_sortie_edit | Nom du fichier Excel de sortie (pré-rempli avec timestamp) |
| `QLabel` | Affichage du dossier de sortie fixe (`~/Documents/Exports_PLM/`) |
| `QPushButton` btn_extraire | Lance l'extraction (hauteur min 40px) |
| `QProgressBar` | 0–100%, texte `"%p% - %v"` |
| `QLabel` lbl_progress | Message textuel de progression |
| `QLabel` + `QTextEdit` console | Console dark de logs colorés |

#### Style appliqué

Charge `style.qss` depuis le dossier du script. S'il n'existe pas, style natif Qt.

#### Comportements dynamiques

**`on_type_fichier_changed(index)`** :
- Change le label "Fichier ASM :" / "Fichier pièce :" / "Fichier principal :"
- Vide le champ de fichier

**`choisir_fichier()`** :
- Adapte le filtre du `QFileDialog` selon le type :
  - Assemblage : `*.asm`
  - Pièces : `*.par *.psm`
  - Les deux : `*.asm *.par *.psm`
- Si fichier sélectionné → génère un nouveau nom de sortie : `Export_PLM_{nom_sans_ext}_{timestamp}.xlsx`

**`on_mode_dft_changed(index)`** :
- Mode 0 (arborescence) → désactive le champ dossier DFT et le bouton
- Mode 1 (dossier spécifique) → active les deux
- Mode 2 (les deux) → active les deux

**`choisir_dossier_dft()`** : `QFileDialog.getExistingDirectory()`

**`log(message, msg_type)`** :
- Déplace le curseur à la fin, insère le texte coloré :
  - `'info'` → noir `#000000`
  - `'success'` → vert
  - `'error'` → rouge
  - `'warning'` → orange
- `ensureCursorVisible()` pour auto-scroll

**`lancer_extraction()`** :
- Vérifie qu'une extraction n'est pas déjà en cours
- Vérifie que le fichier sélectionné existe
- Crée `~/Documents/Exports_PLM/` si absent
- Convertit les index des combos en chaînes métier :
  - `0→"arborescence"`, `1→"dossier_specifique"`, `2→"les_deux"`
  - `0→"asm"`, `1→"pieces"`, `2→"les_deux"`
- Crée et démarre `ExtractionThread`
- Connecte les signaux
- Désactive le bouton, le renomme "Extraction en cours..."

**`update_progress(value, maximum, message)`** : met à jour `QProgressBar` et `lbl_progress`

**`extraction_terminee()`** : réactive le bouton, remet la barre à 0 et le label à "Prêt"

#### Fermeture propre (`closeEvent`)

1. Tente `GetActiveObject("SolidEdge.Application")` — si erreur (SE non lancé) → ferme l'app normalement
2. Affiche un `QMessageBox.question` à 3 boutons (Oui / Non / Annuler) :
   - Annuler → `event.ignore()` (ne ferme pas)
   - Non → `event.accept()` (ferme l'app, laisse SE ouvert)
   - Oui → appelle `_fermer_solid_edge(app_se)` puis `event.accept()`

**`_fermer_solid_edge(app_se)`** — 3 méthodes en cascade :
1. `_essayer_quitter(app_se, delai=3)` : `app_se.Quit()` + attente 3s + vérification si encore actif
2. `_fermer_documents_puis_quitter(app_se)` : ferme tous les `app_se.Documents` un par un, puis `Quit()`, attente 2s + vérification
3. `os.system("taskkill /f /im SolidEdge.exe")` : force kill

Pendant la fermeture, affiche `_creer_fenetre_chargement()` : `QDialog` fixe 300×100 avec label centré "Fermeture de Solid Edge..."

---

## 7. Structure du fichier Excel généré

### 7.1 Colonnes (21 au total)

| N° | Nom colonne | Source |
|----|-------------|--------|
| 1 | `Level` | Profondeur dans l'arborescence (0 = racine) |
| 2 | `Relationship` | `""`, `"ComposedOf"` ou `"Drawing"` |
| 3 | `ordre` | Compteur séquentiel global |
| 4 | `quantite` | Nombre d'occurrences de la pièce dans l'assemblage parent |
| 5 | `repere` | Vide (prévu pour usage PLM) |
| 6 | `SpecialCAD` | Nom sans extension du fichier |
| 7 | `Class` | Classe PLM (`SUB_ASSY_A`, `PART_A`, `PART_PURCH_A`, `CAD_DRAWING_A`, `Folder`) |
| 8 | `ref_utilisat` | Nom sans extension du fichier |
| 9 | `version` | "Indice de modification" depuis Custom Solid Edge, `-` si vide |
| 10 | `indice_1` | Version n-1 calculée |
| 11 | `indice_2` | Version n-2 calculée |
| 12 | `revision` | Propriété "Révision" depuis Solid Edge, `1` par défaut |
| 13 | `designation` | Propriété "Désignation" depuis Custom Solid Edge |
| 14 | `cus_createur` | Propriété "Username" depuis ExtendedSummaryInformation |
| 15 | `cus_date_crea` | Propriété "Date de création" normalisée |
| 16 | `user_version_1` | Propriété "auteur modif" |
| 17 | `date_version_1` | Propriété "date modif" normalisée |
| 18 | `matiere` | Propriété "Matière" depuis Custom Solid Edge |
| 19 | `densite` | Propriété "Densité" depuis Custom Solid Edge |
| 20 | `dia_se` | Vide (prévu) |
| 21 | `Attachments` | Chemin absolu + suffixe `(ASM)`/`(PRT)`/`(DRW)` |

### 7.2 Formatage Excel

- Ligne 1 : en-têtes en gras
- Colonnes 1–12 : fond vert clair `CCFFCC`
- Colonnes 13–21 : fond orange `FFC000`
- Largeur de chaque colonne = `max(len(valeur))` + 2 caractères (calculé en parallèle)
- Onglet nommé `"Structure"`

### 7.3 Structure logique des lignes

```
Level 0 | ""         | MonAssemblage.asm    | SUB_ASSY_A    ← racine
Level 1 | ComposedOf | SousAsm1.asm         | SUB_ASSY_A    ← sous-assemblage
Level 2 | ComposedOf | Piece1.par           | PART_A        ← pièce
Level 0 | ""         | Piece1.dft           | CAD_DRAWING_A ← plan DFT
Level 1 | Drawing    | Piece1.par           | PART_A        ← lien pièce→plan
```

---

## 8. Résultat attendu

- Un fichier `.xlsx` nommé `Export_PLM_{nom_asm}_{timestamp}.xlsx`
- Déposé dans `~/Documents/Exports_PLM/` (créé automatiquement si absent)
- Contenant une ligne par composant 3D et une paire de lignes par plan 2D associé
- Toutes les métadonnées extraites directement depuis Solid Edge sans manipulation manuelle
- Prêt à être importé dans un système PLM (structure et nommage de colonnes conformes)

---

## 9. Points techniques notables pour une IA

- **COM réentrant** : les objets COM Solid Edge ne sont pas thread-safe. `extraire_metadonnees_rapide()` utilise `SolidEdge.FileProperties` (lecture de fichier, pas d'objet UI) → peut être appelé sans `pythoncom.CoInitialize()`. En revanche, `app.Documents.Open()` doit rester sur le thread principal.
- **Cache LRU** : le cache `g_metadata_cache` est partagé entre tous les appels de la session. Il survit à plusieurs extractions consécutives dans la même exécution GUI.
- **Déduplications** :
  - Occurrences : une même pièce présente N fois dans un assemblage → une seule ligne avec `quantite=N`
  - Plans : un même `.dft` lié à plusieurs pièces → une seule paire de lignes DFT/Drawing
- **Fallbacks COM** : 3 niveaux pour récupérer le chemin d'une occurrence (robustesse face aux assemblages avec liens brisés)
- **Héritage métadonnées DFT** : si un plan n'a pas d'auteur/date, les valeurs sont copiées depuis la pièce 3D correspondante
- **Chemins réseau UNC avec IP** : la normalisation `os.path.normpath` est court-circuitée pour les adresses IP afin d'éviter une résolution DNS involontaire
