# Documentation Modèle de Données — Budget Construction Excel

Ce document décrit le modèle de données complet du projet `budget_construction_excel`, incluant les tables Excel, le pipeline Power Query et les relations entre les entités.

---

## Table des matières

1. [Vue générale](#vue-générale)
2. [Modèle conceptuel](#modèle-conceptuel)
3. [Tables de référence](#tables-de-référence)
4. [Tables de saisie](#tables-de-saisie)
5. [Tables enrichies](#tables-enrichies)
6. [Tables analytiques](#tables-analytiques)
7. [Flux de données](#flux-de-données)
8. [Types de données](#types-de-données)
9. [Clés et relations](#clés-et-relations)
10. [Règles métier](#règles-métier)

---

# Vue générale

Le modèle de données suit une architecture **mini-BI / Data Warehouse** structurée en couches :

```
┌─────────────────────────────────────────────────────┐
│             COUCHE PRÉSENTATION                      │
│  (Dashboards, Recherche, Reporting)                 │
├─────────────────────────────────────────────────────┤
│             TABLES ANALYTIQUES                       │
│  (fact_couts, r_search_results)                      │
├─────────────────────────────────────────────────────┤
│           TRANSFORMATION POWER QUERY                 │
│  (fact_transactions, devis_only, facture_only, etc.) │
├─────────────────────────────────────────────────────┤
│            TABLES D'ENRICHISSEMENT                   │
│  (input_staging_enriched, tbl_input_staging_sync)    │
├─────────────────────────────────────────────────────┤
│            TABLES DE SAISIE ET RÉFÉRENCE             │
│  (input_staging, dim_produits, tbl_fournisseurs)    │
├─────────────────────────────────────────────────────┤
│                 SOURCE UTILISATEUR                   │
│  (Saisie Excel, Google Apps Script, Google Drive)   │
└─────────────────────────────────────────────────────┘
```

---

# Modèle conceptuel

## Entités principales

```
┌──────────────────┐         ┌──────────────────┐
│   PRODUITS       │         │   FOURNISSEURS   │
│                  │         │                  │
│ • categorie_id   │         │ • fournisseur_id │
│ • categorie      │         │ • fournisseur    │
│ • produit_id     │         │ • adresse        │
│ • produit        │         │ • contact        │
│ • sous_produit   │         │ • telephone      │
└──────────────────┘         │ • email          │
         │                   └──────────────────┘
         │                            ▲
         │                            │
         ├────────────────────────────┘
         │
    ┌────▼────────────────┐
    │   TRANSACTIONS      │
    │                     │
    │ • input_id          │
    │ • date              │
    │ • type (Devis/DIY/  │
    │   Facture)          │
    │ • quantite          │
    │ • prix_unitaire     │
    │ • cout_total        │
    │ • fournisseur       │
    │ • ref               │
    │ • cout_diy          │
    │ • cout_reel         │
    │ • estimation_basse  │
    │ • estimation_haute  │
    └─────────────────────┘
```

---

# Tables de référence

## Table `dim_produits`

### Rôle

Référentiel maître des produits et catégories. Table structurée Excel gérée manuellement.

### Structure

| Colonne          | Type      | Obligatoire | Description                          |
| ---------------- | --------- | ----------- | ------------------------------------ |
| `categorie_id`   | Entier    | Oui         | Identifiant unique catégorie         |
| `categorie`      | Texte     | Oui         | Nom de la catégorie                  |
| `sous_categorie` | Texte     | Oui         | Classement fin au sein de catégorie  |
| `produit`        | Texte     | Oui         | Nom du produit principal             |
| `produit_id`     | Entier    | Oui         | Identifiant unique produit           |

### Exemple de données

| categorie_id | categorie   | sous_categorie    | produit        | produit_id |
| ------------ | ----------- | ----------------- | -------------- | ---------- |
| 1            | Structure   | Fondations        | Béton          | 101        |
| 1            | Structure   | Murs              | Brique         | 102        |
| 2            | Finitions   | Revêtements murs  | Carrelage      | 201        |
| 2            | Finitions   | Sols              | Parquet        | 202        |

### Dépendances

- Utilisée par `dim_produits` Power Query ;
- Lookup par `input_staging` (VBA) ;
- Référence pour `dim_produits` (Power Query).

---

## Table `tbl_fournisseurs`

### Rôle

Annuaire des fournisseurs. Gérée via macro VBA `modSuppliers`.

### Structure

| Colonne               | Type  | Obligatoire | Description                    |
| --------------------- | ----- | ----------- | ------------------------------ |
| `fournisseur_id`      | Entier | Oui        | Identifiant unique             |
| `fournisseur`         | Texte | Oui        | Nom du fournisseur / entreprise |
| `adresse`             | Texte | Non        | Adresse postale                |
| `contact_principal`   | Texte | Non        | Nom du contact principal       |
| `tel_contact`         | Texte | Non        | Téléphone contact principal    |
| `email_contact`       | Texte | Non        | Email contact principal        |
| `contact_secondaire`  | Texte | Non        | Nom contact secondaire         |
| `tel_contact_secondaire` | Texte | Non     | Téléphone contact secondaire    |

### Exemple de données

| fournisseur_id | fournisseur     | adresse                     | contact_principal |
| -------------- | --------------- | --------------------------- | ----------------- |
| 1              | MaCo Matériaux  | 123 Rue de Paris, 75001    | Jean Dupont       |
| 2              | Plomberie Pro   | 456 Avenue Marcel, 69000   | Marie Martin      |
| 3              | Électricité +   | 789 Boulevard Paris, 13000 | Luc Lefevre       |

---

## Table `tbl_sous_produits`

### Rôle

Variantes et déclinaisons des produits principaux. Gérée via macro VBA `modSousProduits`.

### Structure

| Colonne        | Type      | Obligatoire | Description                |
| -------------- | --------- | ----------- | -------------------------- |
| `sous_produit_id` | Entier | Oui        | Identifiant unique         |
| `produit_id`   | Entier    | Oui        | Lien vers produit principal |
| `sous_produit` | Texte     | Oui        | Nom de la variante         |
| `description`  | Texte     | Non        | Description technique      |

### Exemple de données

| sous_produit_id | produit_id | sous_produit       | description             |
| --------------- | ---------- | ------------------ | ----------------------- |
| 1001            | 201        | Carrelage 30x30    | Format standard         |
| 1002            | 201        | Carrelage 60x60    | Grand format            |
| 1003            | 201        | Carrelage sanitaire| Pour salle de bain      |

---

# Tables de saisie

## Table `input_staging`

### Rôle

Zone de saisie des transactions utilisateur. Les données sont ajoutées par :
- Macro VBA `modInputStaging` (formulaire Excel) ;
- Import manuel direct.

Les données sont ensuite enrichies et transformées par Power Query.

### Structure complète

| Colonne          | Type      | Obligatoire | Description                                    |
| ---------------- | --------- | ----------- | ---------------------------------------------- |
| `input_id`       | Entier    | Oui         | Identifiant unique transaction                 |
| `date`           | Date      | Oui         | Date de la transaction                         |
| `categorie_id`   | Entier    | Oui         | Lien vers `dim_produits`                       |
| `categorie`      | Texte     | Oui         | Libellé catégorie (lookup)                     |
| `sous_categorie` | Texte     | Oui         | Libellé sous-catégorie (lookup)                |
| `produit`        | Texte     | Oui         | Libellé produit principal (lookup)             |
| `produit_id`     | Entier    | Oui         | Lien vers produit dans `dim_produits`          |
| `sous_produit`   | Texte     | Non         | Variante ou sous-produit ; remplacé par produit si vide |
| `type`           | Texte     | Oui         | Valeur exacte : `Devis` \| `Facture` \| `DIY` |
| `fournisseur`    | Texte     | Non         | Nom du fournisseur                             |
| `ref`            | Texte     | Non         | Référence ou numéro de commande                |
| `quantite`       | Nombre    | Oui         | Quantité commandée / estimée                   |
| `prix_unitaire`  | Nombre    | Oui         | Prix unitaire HT                               |
| `Commentaire`    | Texte     | Non         | Notes libres (max 255 caractères)              |
| `is_deleted`     | Booléen   | Non         | Marque suppression logique (TRUE/FALSE)        |

### Exemple de données

| input_id | date       | categorie_id | categorie | sous_categorie | produit    | produit_id | sous_produit | type    | fournisseur   | ref      | quantite | prix_unitaire | is_deleted |
| -------- | ---------- | ------------ | --------- | -------------- | ---------- | ---------- | ------------ | ------- | ------------- | -------- | -------- | ------------- | ---------- |
| 1        | 2024-01-15 | 2            | Finitions | Revêt. murs    | Carrelage  | 201        | 30x30        | Devis   | MaCo Mat.     | DEV-2024 | 15       | 12.50         | FALSE      |
| 2        | 2024-01-20 | 2            | Finitions | Revêt. murs    | Carrelage  | 201        | 30x30        | Facture | MaCo Mat.     | FAC-0042 | 15       | 12.00         | FALSE      |
| 3        | 2024-02-01 | 2            | Finitions | Revêt. murs    | Carrelage  | 201        | 30x30        | DIY     | NULL          | NULL     | 20       | 0.00          | FALSE      |

### Normalisation automatique

Power Query applique les transformations suivantes :

```
Si sous_produit est NULL ou vide
  Alors sous_produit = produit
```

Cela normalise les trois scénarios :
- produits sans variante (carrelage générique) ;
- produits avec sous-produit explicite (carrelage 30x30) ;
- saisies manuelles incomplètes.

---

# Tables enrichies

## Table `tbl_input_staging_sync`

### Rôle

Métadonnées des documents uploadés via Google Apps Script. Importées depuis Google Sheets via CSV.

### Structure

| Colonne        | Type      | Obligatoire | Description                  |
| -------------- | --------- | ----------- | ---------------------------- |
| `input_id`     | Entier    | Oui         | Lien vers transaction Excel  |
| `fichier`      | Texte     | Oui         | Nom du fichier uploadé       |
| `source_drive` | Texte     | Oui         | URL Google Drive du fichier  |
| `fichier_id`   | Entier    | Oui         | Identifiant technique fichier |

### Exemple de données

| input_id | fichier              | source_drive                      | fichier_id |
| -------- | -------------------- | --------------------------------- | ---------- |
| 1        | devis_maco_2024.pdf  | https://drive.google.com/file/... | 98765      |
| 2        | facture_maco_042.pdf | https://drive.google.com/file/... | 98766      |

### Configuration

Le lien CSV vers le Google Sheet doit être remplacé dans Power Query :

```powerquery
Web.Contents("<GOOGLE_SHEET_CSV_EXTRACT_LINK>")
```

Ne pas publier le lien réel en clair.

---

## Table `input_staging_enriched`

### Rôle

Résultat de la jointure entre `input_staging` et `tbl_input_staging_sync` via Power Query.

### Structure

Toutes les colonnes de `input_staging` +

| Colonne        | Type   | Description                |
| -------------- | ------ | -------------------------- |
| `fichier`      | Texte  | Nom du fichier uploadé     |
| `source_drive` | Texte  | URL Google Drive           |
| `fichier_id`   | Entier | ID technique fichier       |

### Jointure

```
Jointure : Left Outer
  input_staging.input_id = tbl_input_staging_sync.input_id
```

**Résultat :** Toutes les transactions sont conservées, même sans document uploadé.

### Usage

- Utilisée par `fact_transactions` pour produire la table de faits ;
- Utilisée par `modSearch` pour afficher les résultats avec liens cliquables.

---

# Tables analytiques

## Table `fact_transactions`

### Rôle

Table de faits consolidée. Résultat de la transformation de `input_staging_enriched` par Power Query.

### Transformations appliquées

1. **Typages** : application stricte des types de données ;
2. **Suppression logique** : exclusion des lignes où `is_deleted = TRUE` ;
3. **Calcul coût** : `cout_total = quantite × prix_unitaire` ;
4. **Clé de comparaison** : `comparison_key = produit_id | sous_produit_normalisé` ;
5. **Suppression colonnes** : `fichier`, `source_drive`, `fichier_id` sont retirées.

### Structure

Toutes les colonnes de `input_staging` +

| Colonne         | Type   | Calcul / Description              |
| --------------- | ------ | --------------------------------- |
| `cout_total`    | Nombre | `quantite * prix_unitaire`        |
| `comparison_key` | Texte | `produit_id \| sous_produit_norm` |

*Moins* : `fichier`, `source_drive`, `fichier_id`

### Exemple de données

| input_id | date       | categorie_id | produit | produit_id | sous_produit | type    | quantite | prix_unitaire | cout_total | comparison_key |
| -------- | ---------- | ------------ | ------- | ---------- | ------------ | ------- | -------- | ------------- | ---------- | -------------- |
| 1        | 2024-01-15 | 2            | Carrel. | 201        | 30x30        | Devis   | 15       | 12.50         | 187.50     | 201\|30x30     |
| 2        | 2024-01-20 | 2            | Carrel. | 201        | 30x30        | Facture | 15       | 12.00         | 180.00     | 201\|30x30     |

---

## Table `devis_only`

### Rôle

Agrégation des transactions de type `Devis` par produit / sous-produit.

### Groupement

```
GROUP BY
  categorie, sous_categorie, produit, produit_id,
  sous_produit, comparison_key
```

### Colonnes calculées

| Colonne            | Calcul                                  |
| ------------------ | --------------------------------------- |
| `estimation_basse` | MIN(cout_total)                         |
| `estimation_haute` | MAX(cout_total)                         |
| `cout_moyen`       | (estimation_basse + estimation_haute)/2 |

### Colonnes enrichies (rejointes depuis `fact_transactions`)

| Colonne        | Description                        |
| -------------- | ---------------------------------- |
| `vendeur_bas`  | Fournisseur du devis le moins cher |
| `ref_bas`      | Référence du devis le moins cher   |
| `vendeur_haut` | Fournisseur du devis le plus cher  |
| `ref_haut`     | Référence du devis le plus cher    |

### Exemple de données

| produit   | sous_produit | estimation_basse | estimation_haute | cout_moyen | vendeur_bas    | vendeur_haut |
| --------- | ------------ | ---------------- | ---------------- | ---------- | -------------- | ------------ |
| Carrelage | 30x30        | 150.00           | 200.00           | 175.00     | MaCo Matériaux | Carrell'Pro  |

---

## Table `diy_only`

### Rôle

Agrégation des transactions de type `DIY` (Do It Yourself) par produit / sous-produit.

### Groupement

```
GROUP BY
  categorie, sous_categorie, produit, produit_id,
  sous_produit, comparison_key
```

### Colonne calculée

| Colonne    | Calcul            |
| ---------- | ----------------- |
| `cout_diy` | SUM(cout_total)   |

### Exemple de données

| produit   | sous_produit | cout_diy |
| --------- | ------------ | -------- |
| Carrelage | 30x30        | 0.00     |

---

## Table `facture_only`

### Rôle

Agrégation des transactions de type `Facture` (coûts réels) par produit / sous-produit.

### Groupement

```
GROUP BY
  categorie, sous_categorie, produit, produit_id,
  sous_produit, comparison_key
```

### Colonne calculée

| Colonne     | Calcul          |
| ----------- | --------------- |
| `cout_reel` | SUM(cout_total) |

### Exemple de données

| produit   | sous_produit | cout_reel |
| --------- | ------------ | --------- |
| Carrelage | 30x30        | 180.00    |

---

## Table `master_keys`

### Rôle

Index unique des clés produit / sous-produit présentes dans au moins une des trois tables analytiques intermédiaires.

### Construction

```
Combine :
  - comparison_key depuis devis_only
  - comparison_key depuis diy_only
  - comparison_key depuis facture_only

Déduplique sur comparison_key
```

### Colonnes conservées

| Colonne          |
| ---------------- |
| `categorie`      |
| `sous_categorie` |
| `produit`        |
| `produit_id`     |
| `sous_produit`   |
| `comparison_key` |

---

## Table `fact_couts`

### Rôle

**Table analytique principale** : consolidation complète des coûts par produit.

Combine :
- clés uniques de `master_keys` ;
- estimations de `devis_only` ;
- estimations DIY de `diy_only` ;
- coûts réels de `facture_only`.

### Jointures

```
master_keys [LEFT OUTER] devis_only ON comparison_key
   ↓
résultat [LEFT OUTER] diy_only ON comparison_key
   ↓
résultat [LEFT OUTER] facture_only ON comparison_key
```

### Structure complète

| Colonne             | Source | Type   | Description                     |
| ------------------- | ------ | ------ | ------------------------------- |
| `categorie`         | keys   | Texte  | Catégorie produit               |
| `sous_categorie`    | keys   | Texte  | Sous-catégorie                  |
| `produit`           | keys   | Texte  | Produit principal               |
| `produit_id`        | keys   | Entier | ID produit                      |
| `sous_produit`      | keys   | Texte  | Variante                        |
| `comparison_key`    | keys   | Texte  | Clé de rapprochement            |
| `vendeur_bas`       | devis  | Texte  | Fournisseur devis minimum       |
| `ref_bas`           | devis  | Texte  | Référence devis minimum         |
| `estimation_basse`  | devis  | Nombre | Montant minimum devis           |
| `vendeur_haut`      | devis  | Texte  | Fournisseur devis maximum       |
| `ref_haut`          | devis  | Texte  | Référence devis maximum         |
| `estimation_haute`  | devis  | Nombre | Montant maximum devis           |
| `cout_moyen`        | devis  | Nombre | Moyenne estimations             |
| `cout_diy`          | diy    | Nombre | Total estimations DIY           |
| `cout_reel`         | fact   | Nombre | Total factures                  |

### Utilisation

Cette table alimente :
- Dashboard budget vs réalité ;
- Analyse écarts estimation / réalité ;
- Reporting financier ;
- Suivi par produit / catégorie.

---

## Table `r_search_filters`

### Rôle

Critères de recherche saisis par l'utilisateur dans l'interface `SEARCH`.

### Structure

| Colonne      | Type   | Description              |
| ------------ | ------ | ------------------------ |
| `filter_name` | Texte | Nom du critère           |
| `filter_value` | Texte | Valeur recherchée        |

### Filtres disponibles

| Critère         | Type de valeur | Comportement            |
| --------------- | -------------- | ----------------------- |
| `categorie`     | Texte          | Recherche partielle     |
| `produit`       | Texte          | Recherche partielle     |
| `fournisseur`   | Texte          | Recherche partielle     |
| `type`          | Texte          | Correspondance exacte   |
| `date_min`      | Date           | Borne inférieure        |
| `date_max`      | Date           | Borne supérieure        |

---

## Table `r_search_results`

### Rôle

Résultats de la recherche filtrée, produits par la macro VBA `modSearch`.

### Structure

Toutes les colonnes de `input_staging_enriched`, filtrées selon les critères de `r_search_filters`.

### Colonnes documentaires

| Colonne        | Type   | Description                              |
| -------------- | ------ | ---------------------------------------- |
| `fichier`      | Texte  | Nom du document uploadé                  |
| `source_drive` | Texte  | **URL cliquable** vers Google Drive      |
| `fichier_id`   | Entier | Identifiant technique                    |

Les URLs Google Drive sont converties en hyperliens Excel cliquables par VBA.

---

# Flux de données

## Flux global

```
┌─────────────────────────────────┐
│  SAISIE UTILISATEUR             │
│  - Formulaire Excel INPUT       │
│  - Google Apps Script (uploads) │
└────────────┬────────────────────┘
             │
             ├─→ VALIDATION VBA
             │
             ▼
┌─────────────────────────────────┐
│  STAGING ZONE (Excel)           │
│  - input_staging                │
│  - tbl_input_staging_sync       │
└────────────┬────────────────────┘
             │
             ▼
┌─────────────────────────────────┐
│  POWER QUERY TRANSFORMATION      │
│  - input_staging_enriched       │
│  - fact_transactions            │
│  - Agrégations (devis/DIY/fact) │
└────────────┬────────────────────┘
             │
             ├─→ fact_couts (analytique)
             ├─→ r_search_results (recherche)
             └─→ Dashboards & Reporting
```

## Flux détaillé par module

### Insertion de transaction

```
INPUT formulaire VBA
  ↓
Validation (modInputStaging)
  ↓
Lookup IDs (modUtils)
  ↓
Insertion dans input_staging
  ↓
Power Query refresh
  ↓
input_staging_enriched (jointure documents)
  ↓
fact_transactions (calcul coûts)
```

### Recherche de transaction

```
Utilisateur saisit filtres dans r_search_filters
  ↓
Macro modSearch lit les filtres
  ↓
Filtre input_staging_enriched
  ↓
Écrit résultats dans r_search_results
  ↓
Formatage hyperliens Google Drive
  ↓
Affichage à l'utilisateur
```

### Suppression logique

```
Utilisateur demande suppression
  ↓
Macro modSearch marque is_deleted = TRUE
  ↓
Power Query re-filtre fact_transactions
  ↓
Coûts recalculés automatiquement dans fact_couts
```

---

# Types de données

## Typage par domaine

### Identifiants

| Colonne          | Type Power Query | Format              | Plage               |
| ---------------- | ---------------- | ------------------- | ------------------- |
| `categorie_id`   | Int64            | Entier              | 1+                  |
| `produit_id`     | Int64            | Entier              | 1+                  |
| `input_id`       | Int64            | Entier              | 1+                  |
| `fournisseur_id` | Int64            | Entier              | 1+                  |
| `fichier_id`     | Int64            | Entier              | 1+                  |

### Dates

| Colonne | Type Power Query | Format        | Contrainte      |
| ------- | ---------------- | ------------- | --------------- |
| `date`  | Date             | YYYY-MM-DD    | Obligatoire     |

### Texte

| Colonne          | Type Power Query | Longueur max | Notes                  |
| ---------------- | ---------------- | ------------ | ---------------------- |
| `categorie`      | Text             | 50           | Vocabulaire fixe       |
| `produit`        | Text             | 100          | -                      |
| `sous_produit`   | Text             | 100          | Normalisé (trim, lower) |
| `fournisseur`    | Text             | 100          | Avec doublons possible |
| `ref`            | Text             | 50           | Libre                  |
| `type`           | Text             | 20           | Valeurs : Devis, DIY, Facture |
| `Commentaire`    | Text             | 255          | Libre                  |

### Nombres

| Colonne           | Type Power Query | Décimales | Contrainte         |
| ----------------- | ---------------- | --------- | ------------------ |
| `quantite`        | Number           | 2         | ≥ 0                |
| `prix_unitaire`   | Number           | 2         | ≥ 0                |
| `cout_total`      | Number           | 2         | Calculé            |
| `estimation_basse` | Number          | 2         | Calculé            |
| `estimation_haute` | Number          | 2         | Calculé            |
| `cout_moyen`      | Number           | 2         | Calculé            |
| `cout_diy`        | Number           | 2         | Calculé            |
| `cout_reel`       | Number           | 2         | Calculé            |

### Booléens

| Colonne     | Type Power Query | Valeurs possibles |
| ----------- | ---------------- | ----------------- |
| `is_deleted` | Logical         | TRUE / FALSE      |

---

# Clés et relations

## Relations d'intégrité

### Modèle relationnel

```
dim_produits ─┐
              ├─ input_staging ─ tbl_input_staging_sync
              │
tbl_fournisseurs
```

### Clés primaires

| Table                 | Clé primaire |
| --------------------- | ------------ |
| `dim_produits`        | `produit_id` |
| `tbl_fournisseurs`    | `fournisseur_id` |
| `input_staging`       | `input_id`   |
| `tbl_input_staging_sync` | `input_id` |

### Clés étrangères

| Table          | Colonne        | Référence        | Notes              |
| -------------- | -------------- | ---------------- | ------------------ |
| `input_staging` | `produit_id`   | dim_produits     | Via lookup VBA     |
| `input_staging` | `categorie_id` | dim_produits     | Via lookup VBA     |
| `input_staging_enriched` | `input_id` | tbl_input_staging_sync | Jointure PQ left outer |

### Clés de comparaison

```
comparison_key = Text.From(produit_id) & "|" & Text.Lower(Text.Trim(sous_produit))

Exemple : "201|30x30"

Usage :
  - Groupement dans agrégations
  - Jointures devis_only, diy_only, facture_only
  - Déduplication master_keys
```

---

# Règles métier

## Validation de saisie

### Obligatoires

```
date, type, quantite, prix_unitaire sont obligatoires

Contrôles :
  - date doit être valide
  - type doit être l'une des valeurs : Devis, DIY, Facture
  - quantite et prix_unitaire > 0
```

### Métiers

```
Déduction automatique :
  - Si sous_produit vide → sous_produit = produit
  - Récupération categorie_id et produit_id depuis dim_produits
  - Génération input_id unique via modUtils.GetNextInputID()
```

## Suppression

```
Règle : Soft delete (suppression logique)

Implémentation :
  1. Marquage is_deleted = TRUE dans input_staging
  2. Power Query filtre ces lignes dans fact_transactions
  3. Coûts recalculés automatiquement
  4. Aucune suppression physique

Avantage :
  - Traçabilité complète
  - Récupération facile
  - Audit possible
```

## Types de transaction

```
Types supportés (case-sensitive) :

Devis    → Estimations fournisseur (min, max, moyen)
DIY      → Estimations faites soi-même
Facture  → Coûts réels payés/facturés

Chaque type alimente une requête PQ dédiée.
Modifier ces libellés nécessite une mise à jour de :
  - Power Query
  - Validation Excel
  - VBA si applicable
```

## Consolidation des coûts

```
Structure : Fact table

input_staging
  (quantite, prix_unitaire)
    ↓
fact_transactions
  (cout_total = quantite × prix_unitaire)
    ↓
Agrégations :
  - devis_only   → min, max, moyen
  - diy_only     → somme
  - facture_only → somme
    ↓
fact_couts
  (table analytique complète)
```

## Enrichissement documents

```
Flux :
  1. Utilisateur upload via Google Apps Script
  2. Google Apps Script écrit métadonnées dans Google Sheets
  3. Google Sheets expose CSV
  4. Power Query importe via tbl_input_staging_sync
  5. input_staging_enriched joint les métadonnées
  6. modSearch crée hyperliens cliquables

Lien effectué sur : input_id
```

---

## Résumé des dépendances

```
VBA
  ├─ dim_produits (lookup IDs)
  ├─ input_staging (write)
  ├─ tbl_fournisseurs (write)
  └─ tbl_sous_produits (write)
         ↓
Power Query
  ├─ input_staging (source)
  ├─ tbl_input_staging_sync (web import)
  ├─ input_staging_enriched (transformation)
  ├─ fact_transactions (calcul)
  ├─ devis_only (agrégation)
  ├─ diy_only (agrégation)
  ├─ facture_only (agrégation)
  ├─ master_keys (déduplication)
  └─ fact_couts (consolidation)
         ↓
Reporting & Dashboard
  ├─ Dashboard (fact_couts)
  ├─ Recherche (input_staging_enriched → r_search_results)
  └─ Tracking budget
```

---

## Compatibilité

### Supporté

- Excel Windows & Mac
- Power Query (M language)
- Tables structurées Excel
- VBA
- Web imports (`Web.Contents`)

### Points d'attention

- URLs Google ne doivent pas être en clair en production
- Placeholders doivent être remplacés localement
- Noms de tables et requêtes doivent rester stables
- Colonnes dépendent les unes des autres dans Power Query

---

## Feuille de route de migration

### Vers architecture web

```
Excel input_staging
  ↓
API FastAPI (CREATE /transactions)
  ↓
PostgreSQL transactions table
  ↓
Python ETL (remplace Power Query)
  ↓
PostgreSQL fact_couts
  ↓
React Dashboard
```
