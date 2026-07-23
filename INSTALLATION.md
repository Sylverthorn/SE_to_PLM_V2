# 🛠️ Guide d'Installation — SE to PLM

Ce document détaille les méthodes d'installation et de déploiement de **SE to PLM** sur un poste de travail Windows.

---

## 📋 1. Prérequis Système

| Élément | Exigence minimale |
| :--- | :--- |
| **Système d'exploitation** | Windows 10 / 11 (64 bits) |
| **Logiciel CAO** | Siemens Solid Edge (installé et configuré sur le poste) |
| **Espace disque** | ~100 Mo d'espace disponible |
| **Droits d'accès** | Droits de lecture sur les fichiers 3D/2D et écriture dans le dossier de sortie |

> [!IMPORTANT]
> **Solid Edge** doit être correctement installé sur le poste de travail afin que les bibliothèques COM (`SolidEdge.FileProperties`) soient enregistrées dans le registre Windows.

---

## 🚀 2. Méthode 1 : Déploiement Rapide (Exécutable Sans Python)

C'est la méthode recommandée pour les utilisateurs finaux et les postes de production. Aucune installation de Python n'est requise.

### Étapes d'utilisation :
1. Récupérez le fichier exécutable `SE_to_PLM.exe` situé dans le dossier `dist/`.
2. Placez `SE_to_PLM.exe` dans le dossier de votre choix (ex: Bureau ou `C:\Outils\SE_to_PLM\`).
3. Double-cliquez sur `SE_to_PLM.exe` pour lancer l'application.

> [!NOTE]
> **Gestion des configurations utilisateur :**  
> L'application sauvegarde automatiquement vos préférences (dictionnaire d'abréviations, colonnes Excel, chemins récents) dans `%APPDATA%\SE_to_PLM\`. Vos données restent conservées lors du remplacement ou du déplacement du fichier `.exe`.

---

## 💻 3. Méthode 2 : Lancement depuis les Sources (Mode Développeur)

Si vous souhaitez exécuter ou modifier le code source Python :

### Option A : Lancement automatisé via script VBS (Recommandé)
Double-cliquez sur le fichier `RUN.vbs` situé à la racine du projet.  
Ce script se charge de :
- Créer l'environnement virtuel Python (`venv`) s'il n'existe pas.
- Installer les dépendances depuis `requirements.txt`.
- Lancer l'interface graphique.

### Option B : Lancement manuel via Terminal PowerShell
```powershell
# 1. Ouvrir PowerShell dans le dossier du projet
cd c:\Users\ykorichi\Desktop\DEV\SE_to_PLM_V2

# 2. Autoriser temporairement l'exécution des scripts PowerShell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned

# 3. Créer l'environnement virtuel
python -m venv venv

# 4. Activer l'environnement virtuel
.\venv\Scripts\Activate.ps1

# 5. Installer les dépendances requises
pip install -r requirements.txt

# 6. Lancer l'application
python -m SE_to_PLM.app.main
```

---

## 📦 4. Compilation d'un nouvel exécutable (.exe)

Si vous apportez des modifications au code source et souhaitez générer un nouvel exécutable :

1. Ouvrez un terminal dans le dossier du projet.
2. Lancez le script de compilation automatique :
   ```cmd
   build.bat
   ```
3. L'exécutable mis à jour sera généré dans le dossier `dist/SE_to_PLM.exe`.

---

## 🔧 5. Dépannage à l'installation

> [!WARNING]
> **Erreur "Solid Edge non détecté" ou erreur COM**  
> Assurez-vous d'avoir exécuté Solid Edge au moins une fois sur la session utilisateur pour que ses composants COM soient enregistrés dans le registre Windows.

> [!TIP]
> **Antivirus / Windows Defender Flag**  
> Lors du premier lancement d'un exécutable recompilé, SmartScreen peut afficher un avertissement. Cliquez sur *"Informations complémentaires"* puis *"Exécuter quand même"*.
