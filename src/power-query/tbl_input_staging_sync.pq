let
    Source = Csv.Document(Web.Contents("<GOOGLE_SHEET_CSV_EXTRACT_LINK>"),[Delimiter=",", Columns=4, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
  #"Changed column type" = Table.TransformColumnTypes(#"Promoted Headers", {{"input_id", Int64.Type}, {"fichier", type text}, {"source_drive", type text}, {"fichier_id", Int64.Type}})
in
    #"Changed column type"
