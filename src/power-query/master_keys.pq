let
    DevisKeys =
        Table.SelectColumns(
            devis_only,
            {
                "categorie",
                "sous_categorie",
                "produit",
                "produit_id",
                "sous_produit",
                "comparison_key"
            }
        ),

    DiyKeys =
        Table.SelectColumns(
            diy_only,
            {
                "categorie",
                "sous_categorie",
                "produit",
                "produit_id",
                "sous_produit",
                "comparison_key"
            }
        ),

    FactureKeys =
        Table.SelectColumns(
            facture_only,
            {
                "categorie",
                "sous_categorie",
                "produit",
                "produit_id",
                "sous_produit",
                "comparison_key"
            }
        ),

    #"Appended Keys" =
        Table.Combine({DevisKeys, DiyKeys, FactureKeys}),

    #"Removed Duplicates" =
        Table.Distinct(#"Appended Keys", {"comparison_key"})
in
    #"Removed Duplicates"
