# SE to PLM - Extracteur de Structure Solid Edge

Application industrielle permettant d'extraire la structure complète d'un assemblage Solid Edge (.asm) ou de pièces simples (.par, .psm) vers un format Excel compatible avec les systèmes PLM.

## 🚀 Fonctionnalités Clés

- **Extraction Récursive** : Parcourt toute l'arborescence des assemblages Solid Edge.
- **Indexation des Plans** : Recherche automatique des fichiers de mise en plan (.dft) associés aux pièces 3D.
- **Métadonnées Complètes** : Récupère automatiquement :
  - Désignation, Créateur, Date de création.
  - Historique de modification (Auteur modif, Date modif).
  - Propriétés physiques (Matière, Densité).
  - Propriétés spécifiques (dia_se, version, indices de révision).
- **Ultra-Rapide** : Utilise `SolidEdge.FileProperties` pour lire les métadonnées sans ouvrir l'interface graphique de Solid Edge.
- **Export Excel Stylisé** : Génère un fichier Excel prêt à l'emploi avec mise en forme automatique.

## 📁 Structure du Projet

```text
SE_to_PLM_V2/
├── SE_to_PLM/              # Code source principal
│   ├── app/                # Initialisation et bootstrapping
│   ├── core/               # Logique métier (Cœur de l'application)
│   │   ├── models/         # Définition des données (Metadata, ExportRow, etc.)
│   │   ├── services/       # Services spécialisés (Solid Edge, Excel, Indexing)
│   │   └── use_cases/      # Orchestration des flux de travail
│   ├── infrastructure/     # Adaptateurs techniques (Connexion COM, Système de fichiers)
│   ├── ui/                 # Interface graphique PyQt5
│   └── scripts/            # Scripts utilitaires (RUN.vbs)
├── tests/                  # Suite de tests unitaires et intégration
├── requirements.txt        # Dépendances Python
└── RUN.vbs                 # Lanceur Windows (crée le venv et lance l'app)
```

## 🛠️ Installation et Utilisation

### Prérequis
- Windows avec **Solid Edge** installé.
- Python 3.8+ (si vous ne passez pas par le launcher).

### Lancement Rapide (Windows)
Double-cliquez sur le fichier `RUN.vbs` à la racine ou dans le dossier `scripts/`. 
*Ce script créera automatiquement l'environnement virtuel et installera les dépendances si nécessaire.*

### Lancement Manuel
1. Créez un environnement virtuel : `python -m venv venv`
2. Activez-le : `venv\Scripts\activate`
3. Installez les dépendances : `pip install -r requirements.txt`
4. Lancez l'application : `python -m SE_to_PLM.app.main`

## 🧪 Développement et Tests

Le projet suit les principes de la **Clean Architecture** pour garantir la testabilité et la maintenance.

- **Tests** : Exécutez `pytest` pour lancer l'ensemble des tests.
- **Logs** : Les logs sont affichés dans la console de l'interface graphique et peuvent être étendus via `logger_service`.

## 📄 Licence
Propriété exclusive. Usage industriel réservé.
