let
    Source = Excel.CurrentWorkbook(){[Name="input_staging"]}[Content],

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

    #"Filled Sous Produit" =
        Table.AddColumn(
            #"Changed Type",
            "sous_produit_clean",
            each
                if [sous_produit] = null or Text.Trim([sous_produit]) = ""
                then [produit]
                else Text.Trim([sous_produit]),
            type text
        ),

    #"Removed Old Sous Produit" =
        Table.RemoveColumns(#"Filled Sous Produit", {"sous_produit"}),

    #"Renamed Sous Produit" =
        Table.RenameColumns(
            #"Removed Old Sous Produit",
            {{"sous_produit_clean", "sous_produit"}}
        ),
    #"Reordered Columns" = Table.ReorderColumns(#"Renamed Sous Produit",{"input_id", "date", "categorie_id", "categorie", "sous_categorie", "produit", "produit_id", "sous_produit", "type", "fournisseur", "ref", "quantite", "prix_unitaire", "Commentaire", "is_deleted"})
in
    #"Reordered Columns"
