# 📘 Guide d'Utilisation — SE to PLM

**SE to PLM** est une application industrielle haute performance permettant d'extraire la structure complète d'un assemblage Solid Edge (**ASM**) ou de pièces simples (**PAR**, **PSM**) vers un format Excel optimisé pour l'import PLM.

---

## 🎯 1. Flux de travail global

```mermaid
flowchart LR
    A[Choix du Mode] --> B[Sélection Fichier / Dossier]
    B --> C[Indexation des Plans .DFT]
    C --> D[Options d'Abréviation]
    D --> E[Génération Excel PLM]
```

---

## 📂 2. Modes d'Extraction

L'application propose 3 modes d'extraction adaptés à vos besoins :

### 🔹 2.1. Onglet "Export Assemblage" (ASM principal)
Utilisez ce mode pour extraire l'arborescence complète d'un assemblage principal et de l'ensemble de ses sous-assemblages et pièces.

1. Cliquez sur **Parcourir...** pour sélectionner votre fichier `.asm` principal.
2. Choisissez le **Mode de recherche des mises en plan (.dft)** :
   - **Arborescence du projet** : Recherche automatiquement les plans `.dft` associés dans tous les sous-dossiers du projet.
   - **Dossier spécifique** : Recherche les plans dans un dossier dédié spécifié par l'utilisateur.
   - **Arborescence + Dossier spécifique** : Recherche d'abord dans l'arborescence, puis complète dans le dossier dédié.

---

### 🔹 2.2. Onglet "Export par Lot" (Dossier PAR / PSM)
Utilisez ce mode pour extraire une liste de pièces indépendantes situées dans un dossier (sans fichier d'assemblage parent).

1. Sélectionnez le dossier contenant vos pièces.
2. Cochez **Inclure les sous-dossiers** pour effectuer une recherche récursive.
3. Configurez les options d'indexation des fichiers `.dft` associés.

---

### 🔹 2.3. Onglet "Multi-ASM" (Traitement par lots d'assemblages)
Utilisez ce mode pour analyser automatiquement plusieurs assemblages `.asm` contenus dans un même dossier.

1. Sélectionnez le dossier racine contenant les assemblages.
2. Choisissez le mode de sortie :
   - **Fichiers séparés** : Génère un fichier Excel indépendant par assemblage.
   - **Fichier unifié** : Consolide tous les assemblages dans un seul fichier Excel par blocs.

---

## ⚙️ 3. Options d'Export & Lancement

Situées en bas de la fenêtre principale :

* **Nom du fichier Excel** : Saisissez le nom souhaité pour l'export (ex: `Export_Structure_Machine1.xlsx`).
* **Optimiser les désignations (max 32 car.)** :
  - Cochez cette option pour appliquer le dictionnaire d'abréviations PLM.
  - Seules les désignations dépassant **32 caractères** seront abrégées et converties en MAJUSCULES sans accents.
* **Mettre en rouge les lignes modifiées** :
  - Surligne en rouge dans la feuille Excel toutes les lignes dont la désignation a été modifiée par l'abréviateur.
* **LANCER L'EXTRACTION** : Cliquez pour démarrer le traitement. La barre de progression et le **Journal d'exécution** affichent les étapes en temps réel.

---

## ⚙️ 4. Onglet Paramètres

L'onglet **Paramètres** regroupe la configuration personnalisable du système :

### 📚 A. Dictionnaire d'Abréviations
Permet de définir les règles de conversion des termes longs vers leurs abréviations normées.

* **Ajouter** : Insère un nouveau terme, son abréviation et sa priorité (1 à 3).
* **Supprimer** : Supprime les entrées sélectionnées dans le tableau.
* **Enregistrer** : Sauvegarde le dictionnaire dans votre profil utilisateur.
* **Restaurer défauts** : Réinitialise le dictionnaire aux valeurs d'origine.
* **Importer / Exporter** : Permet de partager le dictionnaire sous forme de fichier Excel (`.xlsx`) ou ODS.

### 📊 B. Configuration des Colonnes
Permet de personnaliser les en-têtes et le choix des propriétés exportées dans le fichier Excel final.

* **Modifier l'en-tête** : Double-cliquez sur une cellule de la colonne *En-tête Excel* pour renommer une colonne.
* **Importer / Exporter la configuration** : Sauvegarder ou réutiliser une structure de colonnes spécifique.
* **Restaurer la configuration par défaut** : Rétablit le modèle de colonnes d'origine.

---

## 💡 5. ASTUCES ET BONNES PRATIQUES

> [!TIP]
> **Recherche rapide dans les abréviations**  
> Dans l'onglet Dictionnaire, utilisez la barre de recherche supérieure pour filtrer instantanément par terme ou par abréviation.

> [!NOTE]
> **Sauvegarde automatique des préférences**  
> L'application retient automatiquement vos derniers choix (dossiers sélectionnés, modes de recherche, options cochées) à la fermeture. Vous n'avez pas besoin de les resaisir à chaque ouverture.
