let
    Source = master_keys,

    #"Merged Devis" =
        Table.NestedJoin(
            Source,
            {"comparison_key"},
            devis_only,
            {"comparison_key"},
            "devis_only",
            JoinKind.LeftOuter
        ),

    #"Expanded Devis" =
        Table.ExpandTableColumn(
            #"Merged Devis",
            "devis_only",
            {
                "vendeur_bas",
                "ref_bas",
                "estimation_basse",
                "vendeur_haut",
                "ref_haut",
                "estimation_haute",
                "cout_moyen"
            },
            {
                "vendeur_bas",
                "ref_bas",
                "estimation_basse",
                "vendeur_haut",
                "ref_haut",
                "estimation_haute",
                "cout_moyen"
            }
        ),

    #"Merged DIY" =
        Table.NestedJoin(
            #"Expanded Devis",
            {"comparison_key"},
            diy_only,
            {"comparison_key"},
            "diy_only",
            JoinKind.LeftOuter
        ),

    #"Expanded DIY" =
        Table.ExpandTableColumn(
            #"Merged DIY",
            "diy_only",
            {"cout_diy"},
            {"cout_diy"}
        ),

    #"Merged Facture" =
        Table.NestedJoin(
            #"Expanded DIY",
            {"comparison_key"},
            facture_only,
            {"comparison_key"},
            "facture_only",
            JoinKind.LeftOuter
        ),

    #"Expanded Facture" =
        Table.ExpandTableColumn(
            #"Merged Facture",
            "facture_only",
            {"cout_reel"},
            {"cout_reel"}
        )
in
    #"Expanded Facture"
