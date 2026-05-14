# SE to PLM - Extracteur de Structure Solid Edge

Application industrielle de haute performance permettant d'extraire la structure complète d'un assemblage Solid Edge (.asm) ou de pièces simples (.par, .psm) vers un format Excel optimisé pour l'import dans les systèmes PLM.

## 🚀 Fonctionnalités Clés

- **Extraction Récursive Multi-Niveau** : Parcourt toute l'arborescence des assemblages Solid Edge, gérant les sous-assemblages de manière intelligente.
- **Mode Multi-ASM (Unifié)** : 
  - Scan automatique de dossiers entiers pour extraire tous les assemblages détectés.
  - Option de sortie en fichiers séparés ou en un seul fichier Excel consolidé par blocs.
- **Indexation Intelligente des Plans** : Recherche parallèle ultra-rapide des fichiers de mise en plan (.dft) associés aux pièces 3D (recherche dans l'arborescence projet et/ou dossiers spécifiques).
- **Métadonnées Complètes** : Récupère automatiquement les propriétés standard et personnalisées :
  - Désignation, Créateur, Date de création.
  - Historique de modification (Auteur modif, Date modif).
  - Propriétés physiques (Matière, Densité).
  - Propriétés spécifiques (dia_se, version, indices de révision calculés).
- **Optimisation de Performance** : 
  - Utilise `SolidEdge.FileProperties` pour une lecture ultra-rapide sans ouvrir l'interface CAO.
  - Système de cache LRU pour les métadonnées afin d'accélérer les extractions répétitives.
- **Interface Graphique Moderne** : GUI intuitive développée en PyQt5, entièrement traduite en français, avec journal d'exécution en temps réel.
- **Gestion Robuste de Solid Edge** : Gestion intelligente des connexions COM, support multi-thread pour éviter le gel de l'interface, et nettoyage sécurisé des processus à la fermeture.

## 📁 Structure du Projet

```text
SE_to_PLM_V2/
├── SE_to_PLM/              # Code source principal
│   ├── app/                # Initialisation, constantes et bootstrapping
│   ├── core/               # Logique métier (Clean Architecture)
│   │   ├── models/         # Objets de données (Metadata, CADFile, ExportRow)
│   │   ├── services/       # Services (Solid Edge, Excel, Cache, Indexing)
│   │   └── use_cases/      # Orchestration des flux (Assembly, Batch, Multi-ASM)
│   ├── infrastructure/     # Adaptateurs techniques (Connexion COM Solid Edge)
│   ├── ui/                 # Interface graphique PyQt5 (GUI, Threads, Styles)
│   └── scripts/            # Scripts utilitaires
├── tests/                  # Suite de tests (Unitaires, Intégration, Performance)
├── requirements.txt        # Dépendances Python (PyQt5, pywin32, openpyxl)
└── RUN.vbs                 # Lanceur Windows automatisé
```

## 🛠️ Installation et Utilisation

### Prérequis
- Windows avec **Solid Edge** installé.
- Python 3.8+ (recommandé pour une exécution manuelle).

### Lancement Rapide (Windows)
Double-cliquez sur le fichier `RUN.vbs` à la racine. 
*Ce script gère tout : création de l'environnement virtuel (`venv`), installation des dépendances et lancement de l'application.*

### Lancement Manuel
1. Créez un environnement virtuel : `python -m venv venv`
2. Activez-le : `venv\Scripts\activate`
3. Installez les dépendances : `pip install -r requirements.txt`
4. Lancez l'application : `python -m SE_to_PLM.app.main`

## 🧪 Développement et Tests

Le projet est conçu selon les principes de la **Clean Architecture** et du **Domain Driven Design** pour garantir une maintenance aisée.

- **Tests** : Exécutez `pytest` pour valider les changements.
- **Robustesse** : Le système de gestion des connexions COM a été spécifiquement optimisé pour éviter les fuites de mémoire et les conflits de threads lors d'extractions successives.

## 📄 Licence
Propriété exclusive. Usage industriel réservé.
