Attribute VB_Name = "modSearch"
Option Explicit

'===============================================================================
' MODULE : modSearch
'
' Rôle :
' - Rafraîchir les résultats de recherche depuis la table input_staging_enriched.
' - Appliquer les filtres saisis dans la feuille RECHERCHE.
' - Reconstruire les liens vers les documents source.
' - Marquer une transaction comme supprimée via un soft delete.
'
' Compatibilité :
' - Version compatible Mac.
' - N'utilise pas Scripting.Dictionary.
'
' Pré-requis :
' - Une feuille nommée "RECHERCHE"
' - Une feuille nommée "input_staging"
' - Une feuille nommée "input_staging_enriched"
' - Une table nommée "r_search_filters"
' - Une table nommée "r_search_results"
' - Une table nommée "input_staging"
' - Une table nommée "input_staging_enriched"
'
' Note GitHub :
' - Les noms de feuilles, tables et colonnes ne sont pas masqués ici.
' - Ils font partie de la structure attendue du classeur.
' - Adapter ces noms si la structure du fichier Excel est modifiée.
'===============================================================================

Public Sub RefreshSearchResults()

    Dim wsSearch As Worksheet
    Dim wsSource As Worksheet
    Dim loFilters As ListObject
    Dim loResults As ListObject
    Dim loSource As ListObject

    Dim filterNames() As String
    Dim filterValues() As Variant
    Dim filterCount As Long
    Dim filterCell As Range
    Dim sourceData As Variant
    Dim resultData() As Variant
    Dim resultCount As Long
    Dim i As Long, j As Long, k As Long
    Dim passesFilter As Boolean
    Dim dateMin As Variant, dateMax As Variant
    Dim srcHeaders As Variant
    Dim srcColCount As Long
    Dim sourceRowCount As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrorHandler

    '===========================================================================
    ' 1. Initialisation des feuilles et des tables
    '===========================================================================

    Set wsSearch = ThisWorkbook.Worksheets("RECHERCHE")
    Set wsSource = ThisWorkbook.Worksheets("input_staging_enriched")

    Set loFilters = wsSearch.ListObjects("r_search_filters")
    Set loResults = wsSearch.ListObjects("r_search_results")
    Set loSource = wsSource.ListObjects("input_staging_enriched")

    '===========================================================================
    ' 2. Lecture des filtres actifs
    '
    ' La table r_search_filters doit contenir au minimum :
    ' - une colonne "filtre"
    ' - une colonne de valeur placée juste à droite
    '
    ' Les filtres date_min et date_max sont traités séparément.
    ' Les autres filtres utilisent une correspondance partielle insensible à la casse.
    '===========================================================================

    filterCount = 0
    ReDim filterNames(1 To 10)
    ReDim filterValues(1 To 10)

    dateMin = Empty
    dateMax = Empty

    If Not loFilters.DataBodyRange Is Nothing Then

        For Each filterCell In loFilters.ListColumns("filtre").DataBodyRange

            Dim fName As String
            Dim fValue As Variant

            fName = LCase(Trim(CStr(filterCell.Value)))
            fValue = filterCell.Offset(0, 1).Value

            If Len(Trim(CStr(fValue))) > 0 Then

                If fName = "date_min" Then
                    dateMin = fValue

                ElseIf fName = "date_max" Then
                    dateMax = fValue

                Else
                    filterCount = filterCount + 1

                    If filterCount > UBound(filterNames) Then
                        ReDim Preserve filterNames(1 To filterCount + 10)
                        ReDim Preserve filterValues(1 To filterCount + 10)
                    End If

                    filterNames(filterCount) = fName
                    filterValues(filterCount) = fValue
                End If

            End If

        Next filterCell

    End If

    '===========================================================================
    ' 3. Lecture des en-têtes de la table source
    '===========================================================================

    srcHeaders = loSource.HeaderRowRange.Value
    srcColCount = UBound(srcHeaders, 2)

    '===========================================================================
    ' 4. Lecture et filtrage des données source
    '===========================================================================

    If loSource.DataBodyRange Is Nothing Then
        resultCount = 0
        GoTo WriteResults
    End If

    sourceData = loSource.DataBodyRange.Value
    sourceRowCount = UBound(sourceData, 1)

    ReDim resultData(1 To sourceRowCount, 1 To 12)
    resultCount = 0

    ' Colonnes attendues dans la table de résultats
    Dim resColNames(1 To 12) As String

    resColNames(1) = "input_id"
    resColNames(2) = "date"
    resColNames(3) = "categorie"
    resColNames(4) = "sous_categorie"
    resColNames(5) = "produit"
    resColNames(6) = "sous_produit"
    resColNames(7) = "type"
    resColNames(8) = "fournisseur"
    resColNames(9) = "ref"
    resColNames(10) = "prix_total"
    resColNames(11) = "commentaire"
    resColNames(12) = "source_drive"

    ' Index des colonnes source correspondant aux colonnes de résultat
    Dim srcIdx(1 To 12) As Long
    Dim qtyIdx As Long, priceIdx As Long, dateIdx As Long
    Dim isDeletedIdx As Long
    Dim headerName As String

    qtyIdx = 0
    priceIdx = 0
    dateIdx = 0
    isDeletedIdx = 0

    For j = 1 To 12

        srcIdx(j) = 0

        For k = 1 To srcColCount

            headerName = LCase(Trim(CStr(srcHeaders(1, k))))

            If headerName = LCase(Trim(resColNames(j))) Then
                srcIdx(j) = k
                Exit For
            End If

        Next k

    Next j

    ' Recherche des colonnes techniques utilisées pour le calcul et le filtrage
    For k = 1 To srcColCount

        headerName = LCase(Trim(CStr(srcHeaders(1, k))))

        Select Case headerName
            Case "quantite"
                qtyIdx = k

            Case "prix_unitaire"
                priceIdx = k

            Case "date"
                dateIdx = k

            Case "is_deleted"
                isDeletedIdx = k
        End Select

    Next k

    If srcIdx(12) = 0 Then
        MsgBox "Colonne 'source_drive' introuvable dans input_staging_enriched." & vbCrLf & _
               "Vérifiez que le nom de colonne est exactement 'source_drive'.", _
               vbExclamation, "Colonne manquante"
    End If

    ' Parcourt chaque ligne source et applique les filtres
    For i = 1 To sourceRowCount

        passesFilter = True

        '-----------------------------------------------------------------------
        ' Application des filtres texte actifs
        '-----------------------------------------------------------------------

        For j = 1 To filterCount

            Dim targetCol As Long
            targetCol = 0

            ' Recherche la colonne correspondant au filtre actif
            For k = 1 To srcColCount

                headerName = LCase(Trim(CStr(srcHeaders(1, k))))

                If headerName = filterNames(j) Then
                    targetCol = k
                    Exit For
                End If

            Next k

            If targetCol > 0 Then

                Dim cellVal As String
                cellVal = CStr(sourceData(i, targetCol))

                ' Correspondance partielle, sans tenir compte de la casse
                If InStr(1, cellVal, CStr(filterValues(j)), vbTextCompare) = 0 Then
                    passesFilter = False
                    Exit For
                End If

            End If

        Next j

        '-----------------------------------------------------------------------
        ' Application des filtres de date
        '-----------------------------------------------------------------------

        If passesFilter And dateIdx > 0 Then

            Dim rowDate As Variant
            rowDate = sourceData(i, dateIdx)

            If Not IsEmpty(dateMin) Then
                If IsDate(rowDate) And IsDate(dateMin) Then
                    If CDate(rowDate) < CDate(dateMin) Then
                        passesFilter = False
                    End If
                End If
            End If

            If passesFilter And Not IsEmpty(dateMax) Then
                If IsDate(rowDate) And IsDate(dateMax) Then
                    If CDate(rowDate) > CDate(dateMax) Then
                        passesFilter = False
                    End If
                End If
            End If

        End If

        '-----------------------------------------------------------------------
        ' Exclusion des transactions supprimées
        '
        ' La suppression est logique : la ligne reste dans input_staging,
        ' mais is_deleted = TRUE l'exclut des résultats.
        '-----------------------------------------------------------------------

        If passesFilter And isDeletedIdx > 0 Then

            Dim delVal As Variant
            delVal = sourceData(i, isDeletedIdx)

            If Not IsEmpty(delVal) Then
                If LCase(Trim(CStr(delVal))) = "true" Or delVal = True Then
                    passesFilter = False
                End If
            End If

        End If

        '-----------------------------------------------------------------------
        ' Ajout de la ligne aux résultats si tous les filtres sont validés
        '-----------------------------------------------------------------------

        If passesFilter Then

            resultCount = resultCount + 1

            For j = 1 To 12

                If j = 10 Then

                    ' Calcul du prix total : quantite * prix_unitaire
                    Dim qty As Double
                    Dim prc As Double

                    qty = 1
                    prc = 0

                    If qtyIdx > 0 Then
                        If IsNumeric(sourceData(i, qtyIdx)) Then qty = CDbl(sourceData(i, qtyIdx))
                    End If

                    If priceIdx > 0 Then
                        If IsNumeric(sourceData(i, priceIdx)) Then prc = CDbl(sourceData(i, priceIdx))
                    End If

                    resultData(resultCount, j) = qty * prc

                Else

                    If srcIdx(j) > 0 Then
                        resultData(resultCount, j) = sourceData(i, srcIdx(j))
                    Else
                        resultData(resultCount, j) = ""
                    End If

                End If

            Next j

        End If

    Next i

WriteResults:

    '===========================================================================
    ' 5. Nettoyage des anciens résultats et écriture des nouvelles données
    '===========================================================================

    If Not loResults.DataBodyRange Is Nothing Then
        loResults.DataBodyRange.Delete
    End If

    If resultCount > 0 Then

        Dim r As Long

        For r = 1 To resultCount
            loResults.ListRows.Add
        Next r

        Dim writeData() As Variant
        ReDim writeData(1 To resultCount, 1 To 12)

        For i = 1 To resultCount
            For j = 1 To 12
                writeData(i, j) = resultData(i, j)
            Next j
        Next i

        loResults.DataBodyRange.Value = writeData

        ' Formatage de la colonne prix_total
        Dim prixTotalCol As ListColumn

        Set prixTotalCol = loResults.ListColumns("prix_total")

        If Not prixTotalCol.DataBodyRange Is Nothing Then
            prixTotalCol.DataBodyRange.NumberFormat = "# ##0 €"
        End If

    End If

    '===========================================================================
    ' 6. Reconstruction des liens vers les documents source
    '===========================================================================

    If resultCount > 0 Then

        Dim srcDriveCol As ListColumn
        Set srcDriveCol = Nothing

        On Error Resume Next
        Set srcDriveCol = loResults.ListColumns("source_drive")
        On Error GoTo ErrorHandler

        If Not srcDriveCol Is Nothing Then

            Dim cell As Range

            On Error Resume Next
            srcDriveCol.DataBodyRange.Hyperlinks.Delete
            On Error GoTo ErrorHandler

            For Each cell In srcDriveCol.DataBodyRange

                If Len(Trim(CStr(cell.Value))) > 0 Then

                    Dim driveUrl As String
                    driveUrl = Trim(CStr(cell.Value))

                    If Len(driveUrl) > 0 Then

                        If LCase(Left(driveUrl, 4)) = "http" Then
                            wsSearch.Hyperlinks.Add _
                                Anchor:=cell, _
                                address:=driveUrl, _
                                TextToDisplay:="Ouvrir fichier"
                        Else
                            cell.Value = driveUrl
                        End If

                    End If

                End If

            Next cell

        End If

    End If

SafeExit:

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Exit Sub

ErrorHandler:

    MsgBox "Erreur dans RefreshSearchResults : " & Err.Description, vbCritical, "Erreur"
    Resume SafeExit

End Sub

Public Sub DeleteSelectedTransaction()

    Dim wsSearch As Worksheet
    Dim wsStaging As Worksheet
    Dim loResults As ListObject
    Dim loStaging As ListObject

    Dim selectedCell As Range
    Dim inputId As Variant
    Dim i As Long
    Dim idColIdx As Long
    Dim deletedColIdx As Long
    Dim found As Boolean

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '===========================================================================
    ' 1. Initialisation des feuilles et des tables
    '===========================================================================

    Set wsSearch = ThisWorkbook.Worksheets("RECHERCHE")
    Set wsStaging = ThisWorkbook.Worksheets("input_staging")

    Set loResults = wsSearch.ListObjects("r_search_results")
    Set loStaging = wsStaging.ListObjects("input_staging")

    '===========================================================================
    ' 2. Vérification de la sélection utilisateur
    '===========================================================================

    Set selectedCell = Selection

    If loResults.DataBodyRange Is Nothing Then
        MsgBox "Le tableau de résultats est vide.", vbExclamation, "Aucune donnée"
        GoTo SafeExit
    End If

    If Intersect(selectedCell, loResults.DataBodyRange) Is Nothing Then
        MsgBox "Veuillez sélectionner une ligne dans le tableau de résultats (r_search_results).", _
               vbExclamation, "Sélection invalide"
        GoTo SafeExit
    End If

    '===========================================================================
    ' 3. Récupération de l'input_id depuis la ligne sélectionnée
    '===========================================================================

    Dim inputIdCol As ListColumn
    Set inputIdCol = loResults.ListColumns("input_id")

    Dim rowOffset As Long
    rowOffset = selectedCell.Row - loResults.DataBodyRange.Row + 1

    inputId = loResults.DataBodyRange.Cells(rowOffset, inputIdCol.Index).Value

    If IsEmpty(inputId) Or Len(Trim(CStr(inputId))) = 0 Then
        MsgBox "Impossible de récupérer l'input_id de la ligne sélectionnée.", _
               vbExclamation, "Erreur"
        GoTo SafeExit
    End If

    '===========================================================================
    ' 4. Confirmation utilisateur
    '===========================================================================

    Dim confirmMsg As String

    confirmMsg = "Voulez-vous vraiment supprimer la transaction #" & inputId & " ?" & vbCrLf & vbCrLf & _
                 "Cette action marquera la transaction comme supprimée."

    If MsgBox(confirmMsg, vbYesNo + vbQuestion, "Confirmer la suppression") <> vbYes Then
        GoTo SafeExit
    End If

    '===========================================================================
    ' 5. Soft delete dans la table input_staging
    '
    ' La ligne n'est pas supprimée physiquement.
    ' La colonne is_deleted est passée à TRUE.
    '===========================================================================

    idColIdx = 0
    deletedColIdx = 0

    For i = 1 To loStaging.ListColumns.Count

        Select Case LCase(Trim(CStr(loStaging.ListColumns(i).Name)))

            Case "input_id"
                idColIdx = i

            Case "is_deleted"
                deletedColIdx = i

        End Select

    Next i

    If idColIdx = 0 Or deletedColIdx = 0 Then
        MsgBox "Colonnes 'input_id' ou 'is_deleted' introuvables dans input_staging.", _
               vbCritical, "Erreur de structure"
        GoTo SafeExit
    End If

    found = False

    If Not loStaging.DataBodyRange Is Nothing Then

        For i = 1 To loStaging.DataBodyRange.Rows.Count

            If CStr(loStaging.DataBodyRange.Cells(i, idColIdx).Value) = CStr(inputId) Then
                loStaging.DataBodyRange.Cells(i, deletedColIdx).Value = True
                found = True
                Exit For
            End If

        Next i

    End If

    If Not found Then
        MsgBox "Transaction #" & inputId & " introuvable dans input_staging.", _
               vbExclamation, "Non trouvé"
        GoTo SafeExit
    End If

    '===========================================================================
    ' 6. Rafraîchissement des résultats
    '===========================================================================

    Call RefreshSearchResults

    MsgBox "Transaction #" & inputId & " supprimée avec succès.", vbInformation, "Suppression effectuée"

SafeExit:

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ThisWorkbook.RefreshAll

    Exit Sub

ErrorHandler:

    MsgBox "Erreur dans DeleteSelectedTransaction : " & Err.Description, vbCritical, "Erreur"
    Resume SafeExit

End Sub

