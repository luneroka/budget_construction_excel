# Documentation Power Query — Budget Construction Excel

Ce document décrit les requêtes Power Query du projet `budget_construction_excel`.

L’objectif de ces requêtes est d’automatiser :

- la normalisation des transactions saisies dans Excel ;
- l’enrichissement avec les documents synchronisés depuis Google Drive ;
- la suppression logique des lignes marquées comme supprimées ;
- le calcul des coûts par transaction ;
- la consolidation des devis, factures et estimations DIY ;
- la production d’une table analytique exploitable pour le reporting.

---

# Architecture Power Query

```text
Tables Excel / Google Sheets
  ↓
Requêtes de staging
  ↓
Transactions enrichies
  ↓
Table de faits
  ↓
Agrégations par type
  ↓
Table analytique des coûts
```

---

# Flux de données

```text
dim_produits

input_staging
  ↓
input_staging_enriched ← tbl_input_staging_sync
  ↓
fact_transactions
  ├─ devis_only ┐
  ├─ diy_only   ├─ master_keys
  └─ facture_only ┘
                  ↓
              fact_couts
```

---

# Structure des requêtes

| Requête                  | Rôle                                                                      |
| ------------------------ | ------------------------------------------------------------------------- |
| `dim_produits`           | Charge et type le catalogue principal des produits                        |
| `input_staging`          | Normalise les transactions saisies dans Excel                             |
| `tbl_input_staging_sync` | Importe les métadonnées de documents depuis un export CSV Google Sheets   |
| `input_staging_enriched` | Joint les transactions aux documents uploadés                             |
| `fact_transactions`      | Filtre les transactions actives et calcule les coûts unitaires consolidés |
| `devis_only`             | Agrège les devis par produit / sous-produit                               |
| `diy_only`               | Agrège les estimations DIY                                                |
| `facture_only`           | Agrège les factures réelles                                               |
| `master_keys`            | Construit la liste unique des clés de comparaison                         |
| `fact_couts`             | Combine les devis, DIY et factures dans une table analytique              |

---

# Fichiers source

| Fichier                                     | Requête                  |
| ------------------------------------------- | ------------------------ |
| `src/power-query/dim_produits.pq`           | `dim_produits`           |
| `src/power-query/input_staging.pq`          | `input_staging`          |
| `src/power-query/tbl_input_staging_sync.pq` | `tbl_input_staging_sync` |
| `src/power-query/input_staging_enriched.pq` | `input_staging_enriched` |
| `src/power-query/fact_transactions.pq`      | `fact_transactions`      |
| `src/power-query/devis_only.pq`             | `devis_only`             |
| `src/power-query/diy_only.pq`               | `diy_only`               |
| `src/power-query/facture_only`              | `facture_only`           |
| `src/power-query/master_keys.pq`            | `master_keys`            |
| `src/power-query/fact_couts.pq`             | `fact_couts`             |

---

# Requête `dim_produits`

## Rôle

Cette requête charge la table structurée Excel :

```text
dim_produits
```

Elle sert de référentiel produit pour les macros VBA et pour les transformations Power Query.

---

## Colonnes attendues

| Colonne          | Type Power Query |
| ---------------- | ---------------- |
| `categorie_id`   | Entier           |
| `categorie`      | Texte            |
| `sous_categorie` | Texte            |
| `produit`        | Texte            |
| `produit_id`     | Entier           |

---

## Notes techniques

- La requête ne crée pas d’identifiants.
- Les identifiants sont supposés déjà présents dans la table Excel.
- Les macros VBA utilisent aussi cette table pour retrouver `categorie_id` et `produit_id`.

---

# Requête `input_staging`

## Rôle

Cette requête charge la table Excel :

```text
input_staging
```

Elle normalise les transactions saisies par l’utilisateur avant enrichissement.

---

## Colonnes attendues

| Colonne          | Type Power Query |
| ---------------- | ---------------- |
| `input_id`       | Entier           |
| `date`           | Date             |
| `categorie_id`   | Entier           |
| `categorie`      | Texte            |
| `sous_categorie` | Texte            |
| `produit`        | Texte            |
| `produit_id`     | Entier           |
| `sous_produit`   | Texte            |
| `type`           | Texte            |
| `fournisseur`    | Texte            |
| `ref`            | Texte            |
| `quantite`       | Nombre           |
| `prix_unitaire`  | Nombre           |
| `Commentaire`    | Texte            |
| `is_deleted`     | Booléen          |

---

## Normalisation du sous-produit

Si `sous_produit` est vide, la requête utilise le nom du produit principal.

```text
si sous_produit est vide
  alors sous_produit = produit
```

Cette règle permet de comparer correctement :

- les produits sans variante ;
- les produits avec sous-produit explicite ;
- les transactions saisies depuis le formulaire VBA.

---

## Ordre de sortie

La requête réordonne les colonnes pour conserver une structure stable :

```text
input_id
date
categorie_id
categorie
sous_categorie
produit
produit_id
sous_produit
type
fournisseur
ref
quantite
prix_unitaire
Commentaire
is_deleted
```

---

# Requête `tbl_input_staging_sync`

## Rôle

Cette requête importe les métadonnées des documents uploadés via Google Apps Script.

La source est un export CSV publié depuis Google Sheets :

```powerquery
Web.Contents("<GOOGLE_SHEET_CSV_EXTRACT_LINK>")
```

---

## Configuration requise

Remplacer le placeholder :

```text
<GOOGLE_SHEET_CSV_EXTRACT_LINK>
```

par l’URL CSV réelle du Google Sheet utilisé comme table de synchronisation.

---

## Colonnes attendues

| Colonne        | Type Power Query | Description                             |
| -------------- | ---------------- | --------------------------------------- |
| `input_id`     | Entier           | Identifiant de transaction Excel        |
| `fichier`      | Texte            | Nom du fichier stocké dans Google Drive |
| `source_drive` | Texte            | URL du fichier Google Drive             |
| `fichier_id`   | Entier           | Identifiant technique du fichier        |

---

## Notes techniques

- La requête attend un CSV avec en-têtes.
- L’encodage utilisé est `1252`.
- Le séparateur attendu est la virgule.
- Le lien CSV ne doit pas être publié en clair dans un dépôt public.

---

# Requête `input_staging_enriched`

## Rôle

Cette requête enrichit les transactions avec les métadonnées de documents.

Elle joint :

```text
input_staging
```

avec :

```text
tbl_input_staging_sync
```

sur :

```text
input_id
```

---

## Type de jointure

La jointure utilisée est :

```text
Left Outer
```

Cela signifie que toutes les transactions Excel sont conservées, même si aucun document n’a encore été uploadé.

---

## Colonnes ajoutées

| Colonne        | Description                 |
| -------------- | --------------------------- |
| `fichier`      | Nom du fichier Google Drive |
| `source_drive` | URL du fichier              |
| `fichier_id`   | Identifiant fichier         |

---

## Usage

Cette requête est utilisée par :

- `fact_transactions` pour produire la table de faits ;
- `modSearch` pour afficher les résultats de recherche avec liens vers les documents.

---

# Requête `fact_transactions`

## Rôle

Cette requête transforme les transactions enrichies en table de faits exploitable pour les calculs budgétaires.

Elle :

- applique les types de données ;
- exclut les lignes supprimées logiquement ;
- calcule le coût total ;
- génère une clé de comparaison ;
- retire les colonnes documentaires non nécessaires au reporting coût.

---

## Suppression logique

Les lignes ne sont pas supprimées physiquement de `input_staging`.

La requête conserve uniquement les lignes où :

```text
is_deleted <> true
```

Cette règle permet aux macros VBA de masquer une transaction sans la perdre.

---

## Coût total

La colonne calculée :

```text
cout_total
```

est définie par :

```text
quantite * prix_unitaire
```

---

## Clé de comparaison

La colonne :

```text
comparison_key
```

est construite avec :

```text
produit_id|sous_produit normalisé
```

Exemple :

```text
42|carrelage mural
```

La normalisation applique :

- `Text.Trim` ;
- `Text.Lower` ;
- `Text.From` sur `produit_id`.

---

## Colonnes supprimées

Les colonnes suivantes sont retirées de la table de faits :

| Colonne        |
| -------------- |
| `fichier`      |
| `source_drive` |
| `fichier_id`   |

Elles restent disponibles dans `input_staging_enriched` pour la recherche et l’accès aux documents.

---

# Requête `devis_only`

## Rôle

Cette requête extrait uniquement les transactions de type :

```text
Devis
```

Elle calcule, pour chaque produit / sous-produit :

- l’estimation basse ;
- l’estimation haute ;
- les vendeurs associés ;
- les références associées ;
- le coût moyen.

---

## Groupement

Les devis sont groupés par :

| Colonne          |
| ---------------- |
| `categorie`      |
| `sous_categorie` |
| `produit`        |
| `produit_id`     |
| `sous_produit`   |
| `comparison_key` |

---

## Colonnes calculées

| Colonne            | Calcul                                  |
| ------------------ | --------------------------------------- |
| `estimation_basse` | Minimum de `cout_total`                 |
| `estimation_haute` | Maximum de `cout_total`                 |
| `cout_moyen`       | Moyenne entre estimation basse et haute |

---

## Colonnes enrichies

La requête rejoint ensuite `fact_transactions` pour récupérer les fournisseurs et références correspondant aux montants minimum et maximum.

| Colonne        | Description                        |
| -------------- | ---------------------------------- |
| `vendeur_bas`  | Fournisseur du devis le moins cher |
| `ref_bas`      | Référence du devis le moins cher   |
| `vendeur_haut` | Fournisseur du devis le plus cher  |
| `ref_haut`     | Référence du devis le plus cher    |

---

## Note technique

Si plusieurs devis ont exactement le même montant minimum ou maximum pour la même clé, la jointure peut produire plusieurs lignes.

---

# Requête `diy_only`

## Rôle

Cette requête extrait uniquement les transactions de type :

```text
DIY
```

Elle consolide les estimations faites soi-même par produit / sous-produit.

---

## Groupement

Les données sont groupées par :

| Colonne          |
| ---------------- |
| `categorie`      |
| `sous_categorie` |
| `produit`        |
| `produit_id`     |
| `sous_produit`   |
| `comparison_key` |

---

## Colonne calculée

| Colonne    | Calcul                |
| ---------- | --------------------- |
| `cout_diy` | Somme de `cout_total` |

---

# Requête `facture_only`

## Rôle

Cette requête extrait uniquement les transactions de type :

```text
Facture
```

Elle consolide les coûts réels payés ou facturés.

---

## Groupement

Les données sont groupées par :

| Colonne          |
| ---------------- |
| `categorie`      |
| `sous_categorie` |
| `produit`        |
| `produit_id`     |
| `sous_produit`   |
| `comparison_key` |

---

## Colonne calculée

| Colonne     | Calcul                |
| ----------- | --------------------- |
| `cout_reel` | Somme de `cout_total` |

---

# Requête `master_keys`

## Rôle

Cette requête construit la liste unique des produits / sous-produits présents dans au moins une des tables analytiques intermédiaires.

Elle combine :

- `devis_only` ;
- `diy_only` ;
- `facture_only`.

---

## Colonnes conservées

| Colonne          |
| ---------------- |
| `categorie`      |
| `sous_categorie` |
| `produit`        |
| `produit_id`     |
| `sous_produit`   |
| `comparison_key` |

---

## Déduplication

La déduplication est faite sur :

```text
comparison_key
```

Cela garantit qu’un produit / sous-produit apparaît une seule fois dans la table finale, même s’il existe à la fois en devis, facture et DIY.

---

# Requête `fact_couts`

## Rôle

Cette requête est la table analytique principale des coûts.

Elle combine :

- les clés uniques de `master_keys` ;
- les estimations issues des devis ;
- les estimations DIY ;
- les coûts réels issus des factures.

---

## Jointures

La requête applique trois jointures `Left Outer` successives.

| Étape   | Source jointe  | Clé              | Colonnes ajoutées                                                                                          |
| ------- | -------------- | ---------------- | ---------------------------------------------------------------------------------------------------------- |
| Devis   | `devis_only`   | `comparison_key` | `vendeur_bas`, `ref_bas`, `estimation_basse`, `vendeur_haut`, `ref_haut`, `estimation_haute`, `cout_moyen` |
| DIY     | `diy_only`     | `comparison_key` | `cout_diy`                                                                                                 |
| Facture | `facture_only` | `comparison_key` | `cout_reel`                                                                                                |

---

## Colonnes finales principales

| Colonne            | Description                            |
| ------------------ | -------------------------------------- |
| `categorie`        | Catégorie produit                      |
| `sous_categorie`   | Sous-catégorie produit                 |
| `produit`          | Produit principal                      |
| `produit_id`       | Identifiant produit                    |
| `sous_produit`     | Variante ou sous-produit               |
| `comparison_key`   | Clé technique de rapprochement         |
| `vendeur_bas`      | Fournisseur du devis minimum           |
| `ref_bas`          | Référence du devis minimum             |
| `estimation_basse` | Montant minimum des devis              |
| `vendeur_haut`     | Fournisseur du devis maximum           |
| `ref_haut`         | Référence du devis maximum             |
| `estimation_haute` | Montant maximum des devis              |
| `cout_moyen`       | Moyenne des estimations basse et haute |
| `cout_diy`         | Total des estimations DIY              |
| `cout_reel`        | Total des factures                     |

---

# Ordre de rafraîchissement recommandé

Power Query sait résoudre les dépendances automatiquement, mais l’ordre logique est :

1. `dim_produits`
2. `input_staging`
3. `tbl_input_staging_sync`
4. `input_staging_enriched`
5. `fact_transactions`
6. `devis_only`
7. `diy_only`
8. `facture_only`
9. `master_keys`
10. `fact_couts`

---

# Types de transactions supportés

Les requêtes analytiques s’appuient sur les valeurs exactes suivantes dans la colonne `type` :

| Valeur    | Requête associée | Usage                       |
| --------- | ---------------- | --------------------------- |
| `Devis`   | `devis_only`     | Estimations fournisseur     |
| `DIY`     | `diy_only`       | Estimations faites soi-même |
| `Facture` | `facture_only`   | Coûts réels                 |

Modifier ces libellés nécessite une mise à jour des requêtes Power Query et des validations de saisie Excel.

---

# Dépendances structurelles importantes

Le pipeline Power Query dépend fortement des :

- noms de requêtes ;
- noms de tables Excel ;
- noms de colonnes ;
- valeurs exactes de `type` ;
- placeholders de configuration Google ;
- liens entre `input_id` et les documents synchronisés.

Modifier un de ces éléments nécessite aussi une mise à jour des macros VBA et, selon le cas, du script Google Apps Script.

---

# Intégration avec VBA

## `modInputStaging`

Le module VBA écrit les transactions dans :

```text
input_staging
```

Power Query reprend ensuite ces lignes pour alimenter :

```text
input_staging_enriched
fact_transactions
fact_couts
```

---

## `modSearch`

Le module de recherche lit :

```text
input_staging_enriched
```

Cette table doit conserver les colonnes documentaires :

- `fichier` ;
- `source_drive` ;
- `fichier_id`.

---

## Suppression logique

La macro de suppression ne supprime pas physiquement la ligne source.

Elle passe :

```text
is_deleted = TRUE
```

dans `input_staging`.

Power Query exclut ensuite cette ligne de `fact_transactions`, donc des agrégations de coûts.

---

# Intégration avec Google Apps Script

Le script Google Apps Script :

- reçoit les documents uploadés ;
- les stocke dans Google Drive ;
- écrit les métadonnées dans Google Sheets ;
- expose ces métadonnées via un CSV lu par `tbl_input_staging_sync`.

Le rapprochement avec Excel se fait via :

```text
input_id
```

---

# Sécurité et publication GitHub

Avant publication :

- remplacer l’URL CSV Google Sheets par `<GOOGLE_SHEET_CSV_EXTRACT_LINK>` ;
- vérifier qu’aucun lien Google Drive privé n’est présent dans les fichiers exportés ;
- anonymiser les fournisseurs, références et commentaires si nécessaire ;
- ne pas publier de données réelles de chantier ;
- ne pas publier d’ID Google Apps Script, Google Drive ou Google Sheets.

---

# Compatibilité

## Supporté

- Excel Mac
- Excel Windows
- Power Query avec langage M
- Tables structurées Excel
- Source CSV externe via `Web.Contents`

## Points d’attention

- `Web.Contents` peut nécessiter une configuration d’autorisations dans Excel.
- Les chemins et URLs Google doivent être propres à chaque utilisateur.
- Les types de colonnes doivent rester cohérents avec les tables Excel.
- Les noms de requêtes doivent rester stables car les requêtes se référencent entre elles.
