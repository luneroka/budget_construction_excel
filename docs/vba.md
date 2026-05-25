# Documentation VBA — Budget Construction Excel

Ce document décrit les principaux modules VBA du projet `budget_construction_excel`.

L’objectif de ces macros est d’automatiser :
- la saisie de transactions ;
- la gestion des fournisseurs ;
- la gestion des sous-produits ;
- la recherche et le filtrage ;
- l’intégration avec Google Apps Script ;
- certaines fonctions utilitaires du classeur.

---

# Architecture VBA

```text
INPUT
  ↓
Macros VBA
  ↓
Tables de staging Excel
  ↓
Power Query
  ↓
Tables enrichies / reporting / recherche
```

---

# Structure des modules

| Module | Rôle |
|---|---|
| `modDocuments` | Gestion de l’upload de documents via Google Apps Script |
| `modInputStaging` | Ajout des transactions dans `input_staging` |
| `modSearch` | Recherche, filtrage et suppression logique |
| `modSousProduits` | Ajout dynamique de sous-produits |
| `modSuppliers` | Ajout dynamique de fournisseurs |
| `modUtils` | Fonctions utilitaires partagées |

---

# Module `modDocuments`

## Rôle

Ce module ouvre une page Google Apps Script permettant d’associer un document (PDF, facture, devis, etc.) à une transaction.

La macro :
- génère un `input_id` unique ;
- verrouille temporairement la transaction ;
- construit une URL d’upload ;
- ouvre la page d’upload dans le navigateur.

---

## Configuration requise

Le module nécessite un déploiement Google Apps Script.

Dans le code VBA :

```vba
baseUrl = "https://script.google.com/macros/s/<GOOGLE_APPS_SCRIPT_DEPLOYMENT_ID>/dev"
```

Remplacer :

```text
<GOOGLE_APPS_SCRIPT_DEPLOYMENT_ID>
```

par l’ID réel du déploiement Apps Script.

---

## Paramètres transmis à Google Apps Script

| Paramètre | Description |
|---|---|
| `input_id` | Identifiant unique de transaction |
| `docType` | Type de document |
| `fournisseur` | Fournisseur associé |
| `ref` | Référence transaction |
| `sous_produit` | Sous-produit associé |

---

## Script Google Apps Script

Le script doit être placé dans :

```text
src/google-apps-script/
```

Exemple :

```text
src/google-apps-script/upload-page.gs
```

---

# Module `modInputStaging`

## Rôle

Ce module ajoute une transaction dans la table `input_staging`.

Il gère :
- la validation des champs obligatoires ;
- la génération d’identifiants ;
- l’écriture dans les tables structurées ;
- la réinitialisation du formulaire.

---

## Flux de données

```text
INPUT
  ↓
Validation
  ↓
input_staging
  ↓
Power Query
  ↓
input_staging_enriched
```

---

## Champs utilisés dans `INPUT`

| Cellule | Champ |
|---|---|
| `D4` | Type |
| `D5` | Catégorie |
| `D6` | Sous-catégorie |
| `D7` | Produit |
| `D8` | Sous-produit |
| `D9` | Fournisseur |
| `D10` | Référence |
| `D11` | Quantité |
| `D12` | Prix unitaire |
| `D14` | Commentaire |

---

## Colonnes générées automatiquement

| Champ | Description |
|---|---|
| `input_id` | Identifiant unique |
| `date` | Date de création |
| `categorie_id` | Lookup depuis `dim_produits` |
| `produit_id` | Lookup depuis `dim_produits` |

---

## Notes techniques

- Si `sous_produit` est vide, le produit principal est utilisé.
- La suppression des transactions fonctionne via `is_deleted`.
- La table est alimentée directement sans passer par une feuille intermédiaire.

---

# Module `modSearch`

## Rôle

Ce module pilote la recherche dans les transactions enrichies.

Il :
- lit les filtres depuis `r_search_filters` ;
- filtre les données de `input_staging_enriched` ;
- écrit les résultats dans `r_search_results` ;
- reconstruit les liens vers les documents source ;
- gère les suppressions logiques.

---

## Fonctionnement des filtres

Les filtres utilisent :
- une recherche partielle ;
- insensible à la casse ;
- compatible Mac.

---

## Filtres de date

Filtres spéciaux :
- `date_min`
- `date_max`

---

## Suppression logique

Les transactions ne sont pas supprimées physiquement.

La macro :

```text
DeleteSelectedTransaction()
```

passe :

```text
is_deleted = TRUE
```

dans `input_staging`.

---

## Liens Google Drive

Les liens présents dans `source_drive` sont reconstruits automatiquement sous forme d’hyperliens cliquables.

---

# Module `modSousProduits`

## Rôle

Ce module permet d’ajouter dynamiquement un sous-produit au référentiel :

```text
tbl_sous_produits
```

---

## Fonctionnement

La macro :
1. lit le produit sélectionné dans `INPUT` ;
2. récupère `categorie_id` et `produit_id` ;
3. vérifie les doublons ;
4. génère un `sous_produit_id` ;
5. ajoute la ligne ;
6. met à jour `INPUT!D8`.

---

## Structure attendue

### INPUT

| Cellule | Champ |
|---|---|
| `D5` | Catégorie |
| `D6` | Sous-catégorie |
| `D7` | Produit |
| `D8` | Sous-produit |

---

## Tables utilisées

| Table | Usage |
|---|---|
| `dim_produits` | Lookup IDs |
| `tbl_sous_produits` | Référentiel sous-produits |

---

# Module `modSuppliers`

## Rôle

Ce module ajoute dynamiquement un fournisseur dans :

```text
tbl_fournisseurs
```

---

## Informations collectées

| Champ | Obligatoire |
|---|---|
| Fournisseur | Oui |
| Adresse | Non |
| Contact principal | Non |
| Téléphone principal | Non |
| Email | Non |

---

## Fonctionnalités

- création de liens `mailto:` ;
- mise à jour automatique du champ fournisseur dans `INPUT!D9` ;
- retour automatique à la feuille précédente.

---

## Champs non gérés automatiquement

Les champs suivants restent modifiables manuellement dans la table :

- `contact_secondaire`
- `tel_contact_secondaire`

---

# Module `modUtils`

## Rôle

Ce module contient les fonctions utilitaires partagées.

---

## Fonctions principales

| Fonction | Rôle |
|---|---|
| `FindHeaderCol()` | Recherche une colonne par son en-tête |
| `GetNextInputID()` | Génère le prochain ID disponible |
| `GenerateInputID()` | Génération temporaire d’ID |
| `ShowConfiguration()` | Affiche la configuration locale |

---

## Configuration Google Drive

Le chemin Drive ne doit pas être publié en clair.

Utiliser un placeholder :

```vba
Public Const DRIVE_FOLDER_PATH As String = "<GOOGLE_DRIVE_FOLDER_PATH>"
```

Chaque utilisateur doit remplacer :

```text
<GOOGLE_DRIVE_FOLDER_PATH>
```

par son propre chemin local synchronisé Google Drive.

---

# Dépendances structurelles importantes

Le projet dépend fortement des :
- noms de feuilles ;
- noms de tables ;
- plages nommées ;
- structures Power Query.

Modifier ces éléments nécessite également une mise à jour des macros VBA.

---

# Compatibilité

## Supporté
- Excel Mac
- Excel Windows

## Évité volontairement
- `Scripting.Dictionary`
- références COM spécifiques Windows

---

# Sécurité et publication GitHub

Avant publication :
- supprimer chemins locaux ;
- remplacer IDs Google Apps Script ;
- retirer données personnelles ;
- anonymiser fournisseurs et contacts si nécessaire.

---

# Structure recommandée du repository

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
│   ├── data-model.md
│   └── vba.md
│
└── README.md
```

---

# Migration vers la web app

Ce projet Excel constitue le prototype fonctionnel du projet :

```text
budget_construction
```

Migration en cours vers :
- FastAPI
- PostgreSQL
- React
- Docker

Objectif :
transformer le workflow Excel en application web multi-utilisateur scalable.
