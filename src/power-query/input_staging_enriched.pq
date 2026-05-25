let
    Source =
        Table.NestedJoin(
            input_staging,
            {"input_id"},
            tbl_input_staging_sync,
            {"input_id"},
            "tbl_input_staging_sync",
            JoinKind.LeftOuter
        ),

    #"Expanded tbl_input_staging_sync" =
        Table.ExpandTableColumn(
            Source,
            "tbl_input_staging_sync",
            {"fichier", "source_drive", "fichier_id"},
            {"fichier", "source_drive", "fichier_id"}
        )
in
    #"Expanded tbl_input_staging_sync"
