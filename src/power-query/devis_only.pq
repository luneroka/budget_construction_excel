let
    Source = fact_transactions,

    #"Filtered Rows" =
        Table.SelectRows(Source, each [type] = "Devis"),

    #"Grouped Rows" =
        Table.Group(
            #"Filtered Rows",
            {
                "categorie",
                "sous_categorie",
                "produit",
                "produit_id",
                "sous_produit",
                "comparison_key"
            },
            {
                {"estimation_basse", each List.Min([cout_total]), type number},
                {"estimation_haute", each List.Max([cout_total]), type number}
            }
        ),

    #"Merged Low" =
        Table.NestedJoin(
            #"Grouped Rows",
            {"comparison_key", "estimation_basse"},
            fact_transactions,
            {"comparison_key", "cout_total"},
            "low_transaction",
            JoinKind.LeftOuter
        ),

    #"Expanded Low" =
        Table.ExpandTableColumn(
            #"Merged Low",
            "low_transaction",
            {"fournisseur", "ref"},
            {"vendeur_bas", "ref_bas"}
        ),

    #"Merged High" =
        Table.NestedJoin(
            #"Expanded Low",
            {"comparison_key", "estimation_haute"},
            fact_transactions,
            {"comparison_key", "cout_total"},
            "high_transaction",
            JoinKind.LeftOuter
        ),

    #"Expanded High" =
        Table.ExpandTableColumn(
            #"Merged High",
            "high_transaction",
            {"fournisseur", "ref"},
            {"vendeur_haut", "ref_haut"}
        ),

    #"Added Average" =
        Table.AddColumn(
            #"Expanded High",
            "cout_moyen",
            each ([estimation_basse] + [estimation_haute]) / 2,
            type number
        )
in
    #"Added Average"
