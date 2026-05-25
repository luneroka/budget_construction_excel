let
    Source = input_staging_enriched,

    #"Changed Type" =
        Table.TransformColumnTypes(
            Source,
            {
                {"input_id", Int64.Type},
                {"date", type date},
                {"categorie_id", Int64.Type},
                {"categorie", type text},
                {"sous_categorie", type text},
                {"produit", type text},
                {"produit_id", Int64.Type},
                {"sous_produit", type text},
                {"type", type text},
                {"fournisseur", type text},
                {"ref", type text},
                {"quantite", type number},
                {"prix_unitaire", type number},
                {"Commentaire", type text},
                {"is_deleted", type logical}
            }
        ),

    #"Filtered Rows" =
        Table.SelectRows(
            #"Changed Type",
            each [is_deleted] <> true
        ),

    #"Added Cout Total" =
        Table.AddColumn(
            #"Filtered Rows",
            "cout_total",
            each [quantite] * [prix_unitaire],
            type number
        ),

    #"Added Comparison Key" =
        Table.AddColumn(
            #"Added Cout Total",
            "comparison_key",
            each Text.From([produit_id]) & "|" & Text.Lower(Text.Trim([sous_produit])),
            type text
        ),

    #"Removed File Columns" =
        Table.RemoveColumns(
            #"Added Comparison Key",
            {"fichier", "source_drive", "fichier_id"}
        )
in
    #"Removed File Columns"
