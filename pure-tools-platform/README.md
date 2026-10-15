# Pure Tools Platform

## Présentation

Pure Tools Platform est une application web qui centralise plusieurs outils d'automatisation liés au paramétrage réglementaire PURE.

Le projet comprend une interface web permettant d'accéder aux outils et un backend chargé de traiter les données et d'exécuter les opérations.

## Fonctionnalités principales

- Initialisation de projets avec intégration Jira.
- Génération de formulaires réglementaires PURE à partir de fichiers DPM et dictionnaires.
- Conversion de fichiers ESPF vers RDS.
- Comparaison de données.
- Suivi de l'exécution des traitements et téléchargement des fichiers générés.

## Technologies utilisées

### Backend
- Python
- FastAPI
- SQLAlchemy
- SQL Server

### Frontend
- React
- TypeScript
- Vite

## Prérequis

- Python 3.12
- Node.js et npm
- SQL Server accessible
- Les fichiers et dossiers nécessaires aux outils PURE

## Installation et lancement

### 1. Cloner le projet

```bash
git clone <URL_DU_DEPOT>
cd pure-tools-platform
```

### 2. Configurer le backend

```bash
cd backend
py -3.12 -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

Configurer ensuite les variables d'environnement nécessaires dans le fichier `.env`, selon la configuration décrite dans le projet.

Démarrer le serveur :

```bash
python -m uvicorn app.main:app --reload
```

L'API sera accessible à l'adresse :

http://127.0.0.1:8000

Documentation interactive de l'API :

http://127.0.0.1:8000/docs

### 3. Configurer le frontend

Dans un autre terminal, depuis la racine du projet :

```bash
cd frontend
npm install
npm run dev
```

Ouvrir l'adresse indiquée dans le terminal, généralement :

http://localhost:5173

## Configuration requise

L'application peut nécessiter une configuration de la connexion SQL Server et du chemin vers les outils PURE externes.

Les variables d'environnement requises, ainsi que les éventuels fichiers modèles nécessaires, devront être configurés localement avant l'utilisation des fonctionnalités concernées.

## Sécurité

Ne pas publier de mots de passe, de clés secrètes, de fichiers `.env` contenant des informations confidentielles ou de données sensibles.

## Auteur

Projet développé dans le cadre de travaux d'ingénierie informatique et d'automatisation d'outils PURE.
