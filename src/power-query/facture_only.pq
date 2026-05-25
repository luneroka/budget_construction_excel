let
    Source = fact_transactions,

    #"Filtered Rows" =
        Table.SelectRows(Source, each [type] = "Facture"),

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
                {"cout_reel", each List.Sum([cout_total]), type number}
            }
        )
in
    #"Grouped Rows"
