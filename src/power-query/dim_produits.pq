let
    Source = Excel.CurrentWorkbook(){[Name="dim_produits"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"categorie_id", Int64.Type}, {"categorie", type text}, {"sous_categorie", type text}, {"produit", type text}, {"produit_id", Int64.Type}})
in
    #"Changed Type"
