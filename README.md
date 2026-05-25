# Budget Construction Excel

<div align="center">

### Système avancé de suivi budgétaire chantier sous Excel

Prototype fonctionnel d'une future application web full-stack dédiée au suivi financier de projets de construction et rénovation.

<br>

![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-867DB1?style=for-the-badge)
![Power Query](https://img.shields.io/badge/Power_Query-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![Google Apps Script](https://img.shields.io/badge/Google_Apps_Script-4285F4?style=for-the-badge&logo=google&logoColor=white)

<br>

![Status](https://img.shields.io/badge/Status-Prototype_Functionnel-success?style=flat-square)
![Architecture](https://img.shields.io/badge/Architecture-ETL_Excel-blue?style=flat-square)
![Migration](https://img.shields.io/badge/Migration-Web_App-orange?style=flat-square)
![Platform](https://img.shields.io/badge/Platform-Mac_&_Windows-lightgrey?style=flat-square)

</div>

---

# Aperçu du projet

`budget_construction_excel` est un système Excel avancé conçu pour piloter le budget complet d’un chantier de construction ou de rénovation.

Le projet centralise :

- les devis
- les factures
- les estimations DIY
- les fournisseurs
- les produits et sous-produits
- les coûts réels
- les documents associés
- les recherches analytiques

L’objectif initial était de remplacer un processus manuel complexe par un workflow structuré, automatisé et maintenable pour un particulier gérant lui-même la construction de sa maison.

Ce prototype constitue aujourd’hui la base fonctionnelle du projet web :

[budget_construction](https://github.com/luneroka/budget_construction)

actuellement en cours de migration vers :

- FastAPI
- PostgreSQL
- React
- Docker

---

# Fonctionnalités principales

## Gestion budgétaire chantier

- suivi du budget prévisionnel
- comparaison budget vs dépenses réelles
- consolidation automatique des coûts
- suivi des acomptes et paiements

---

## Gestion des transactions

Gestion centralisée des :

- devis
- factures
- estimations DIY

Chaque transaction peut être :

- catégorisée
- enrichie
- associée à un fournisseur
- reliée à un document PDF

---

## Gestion fournisseurs

- annuaire fournisseurs
- gestion des contacts
- suivi des achats
- intégration directe dans les workflows

---

## Organisation produits

Hiérarchie complète :

```text
Catégorie
  ↓
Sous-catégorie
  ↓
Produit
  ↓
Sous-produit
```

---

## Recherche et filtrage

- recherche multi-critères
- filtres dynamiques
- recherche fournisseur
- recherche produit
- recherche documents

---

## Intégration Google Drive

Upload automatisé des documents :

- devis
- factures
- pièces justificatives

Via :

- VBA
- Google Apps Script
- Google Drive

---

## Pipeline Power Query

Le projet utilise un pipeline ETL complet dans Excel :

- staging des données
- enrichissement
- normalisation
- consolidation
- tables analytiques
- reporting

---

## Automatisation VBA

Les macros VBA gèrent :

- les workflows utilisateur
- les validations
- la synchronisation des données
- les uploads
- les recherches
- la gestion des référentiels

---

# Architecture du projet

## Pipeline de données

```text
Saisie utilisateur
        ↓
Tables de staging
        ↓
Power Query (ETL)
        ↓
Enrichissement des données
        ↓
Tables analytiques
        ↓
Recherche / Reporting / Dashboard
```

---

## Architecture logique

```text
Excel UI
    ↓
VBA Workflows
    ↓
Tables Structurées
    ↓
Power Query
    ↓
Data Model
    ↓
Reporting & Search
```

---

# Modèle de données

## Tables principales

| Table                    | Rôle                              |
| ------------------------ | --------------------------------- |
| `dim_produits`           | Catalogue principal produits      |
| `tbl_sous_produits`      | Variantes et détails produits     |
| `tbl_fournisseurs`       | Référentiel fournisseurs          |
| `input_staging`          | Zone de staging des transactions  |
| `input_staging_enriched` | Données enrichies via Power Query |
| `fact_transactions`      | Consolidation des transactions    |
| `fact_couts`             | Table analytique des coûts        |
| `r_search_filters`       | Filtres de recherche              |
| `r_search_results`       | Résultats dynamiques              |

---

# Stack technique

<div align="center">

| Technologie        | Usage                                    |
| ------------------ | ---------------------------------------- |
| Excel              | Interface utilisateur & moteur principal |
| VBA                | Automatisation & logique métier          |
| Power Query        | ETL & transformation de données          |
| Google Apps Script | Upload & intégration Drive               |
| Google Drive       | Stockage documentaire                    |

</div>

---

# Fonctionnement Power Query

Le projet utilise Power Query comme véritable pipeline ETL embarqué dans Excel.

## Capacités utilisées

- nettoyage de données
- enrichissement
- jointures
- consolidation
- normalisation
- tables analytiques
- préparation reporting

---

# Structure du repository

```text
budget_construction_excel/
│
├── workbook/
│   └── budget_construction.xlsm
│
├── src/
│   ├── vba/
│   ├── power-query/
│   └── google-apps-script/
│
├── docs/
│   ├── screenshots/
│   ├── architecture.md
│   ├── power-query.md
│   ├── vba.md
│   └── data-model.md
│
└── README.md
```

---

# Captures d’écran

## Dashboard principal

![Dashboard](docs/screenshots/dashboard.png)

Fonctionnalités visibles :

- suivi budgétaire
- pipeline Power Query
- tables analytiques
- recherche
- pilotage chantier

---

# Installation

## Prérequis

- Microsoft Excel
- Macros VBA activées

---

## Ouverture du projet

1. Télécharger le fichier `.xlsm`
2. Activer les macros VBA
3. Rafraîchir les requêtes Power Query si nécessaire

---

# Configuration Google Apps Script

Le projet utilise Google Apps Script pour :

- l’upload de documents
- la génération des URLs Google Drive
- la synchronisation des métadonnées

Pour des raisons de sécurité :

- les IDs réels ne sont pas inclus
- des placeholders sont utilisés dans le repository

Exemple :

```javascript
const CONFIG = {
  DRIVE_FOLDER_ID: '<GOOGLE_DRIVE_FOLDER_ID>',
  SPREADSHEET_ID: '<GOOGLE_SPREADSHEET_ID>',
};
```

---

# Migration vers la web app

Ce projet Excel constitue le prototype fonctionnel du projet :

## Repo associé

➡️ [budget_construction — Version web full-stack](https://github.com/luneroka/budget_construction)

---

## Stack cible

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white)
![FastAPI](https://img.shields.io/badge/FastAPI-009688?style=flat-square&logo=fastapi&logoColor=white)
![PostgreSQL](https://img.shields.io/badge/PostgreSQL-316192?style=flat-square&logo=postgresql&logoColor=white)
![React](https://img.shields.io/badge/React-20232A?style=flat-square&logo=react&logoColor=61DAFB)
![Docker](https://img.shields.io/badge/Docker-2496ED?style=flat-square&logo=docker&logoColor=white)

---

## Objectifs de la migration

- architecture scalable
- multi-utilisateur
- API REST
- persistance PostgreSQL
- UX moderne
- reporting avancé
- gestion documentaire native

---

# Compétences démontrées

## Data

- modélisation de données
- architecture ETL
- structuration analytique
- transformation Power Query
- consolidation de données

---

## Développement

- VBA
- logique métier
- automatisation
- architecture applicative
- intégration Google Apps Script

---

## Produit / métier

- analyse des besoins
- conception d’outil métier
- optimisation de workflow
- suivi budgétaire chantier
- pilotage financier

---

# Notes

Ce repository représente :

- la version historique
- la version fonctionnelle
- le prototype métier original

Le développement actif continue désormais sur la version web full-stack.

---

# Auteur

## Yoann Robert

Développeur full-stack & passionné de data analytics.

Projet développé pour répondre à un besoin réel de pilotage financier chantier sans passer par un constructeur immobilier traditionnel.

Compétences mobilisées :

- Data Engineering
- ETL
- VBA
- Automation
- Architecture applicative
- Développement full-stack
