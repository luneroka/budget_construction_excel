Attribute VB_Name = "modInputStaging"
Option Explicit

'===============================================================================
' MODULE : modInputStaging
'
' Rôle :
' - Ajouter une transaction saisie dans la feuille INPUT vers la table input_staging.
' - Générer ou réutiliser un input_id.
' - Compléter automatiquement certains champs techniques.
' - Réinitialiser le formulaire après ajout.
'
' Pré-requis :
' - Une feuille nommée "INPUT"
' - Une feuille nommée "input_staging"
' - Une table structurée nommée "input_staging"
' - Une plage nommée "InputId"
' - Une plage nommée "upload_status"
' - Une fonction GetNextInputID disponible dans un autre module
'
' Note GitHub :
' - Les noms de feuilles, tables et plages nommées ne sont pas masqués ici,
'   car ils font partie de la structure attendue du classeur.
' - Adapter ces noms si la structure du fichier Excel est modifiée.
'===============================================================================

Public Sub AddInputToStaging()

    Dim wsInput As Worksheet
    Dim wsStaging As Worksheet
    Dim wsReturn As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim input_id As Long

    ' Index des colonnes dans la table input_staging
    Dim colInputID As Long, colDate As Long, colType As Long
    Dim colCategorieId As Long, colCategorie As Long, colSousCat As Long
    Dim colProduit As Long, colProduitId As Long, colSousProduit As Long
    Dim colFournisseur As Long, colRef As Long, colQte As Long
    Dim colPrix As Long, colCommentaire As Long, colIsDeleted As Long

    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsInput = ThisWorkbook.Worksheets("INPUT")
    Set wsStaging = ThisWorkbook.Worksheets("input_staging")
    Set wsReturn = ActiveSheet
    Set tbl = wsStaging.ListObjects("input_staging")

    If tbl Is Nothing Then
        MsgBox "Erreur : table 'input_staging' introuvable.", vbCritical, "Erreur"
        GoTo CleanExit
    End If

    '===========================================================================
    ' Lecture des valeurs saisies dans la feuille INPUT
    '
    ' Structure attendue du formulaire :
    ' D4  = Type
    ' D5  = Catégorie
    ' D6  = Sous-catégorie
    ' D7  = Produit
    ' D8  = Sous-produit
    ' D9  = Fournisseur
    ' D10 = Référence
    ' D11 = Quantité
    ' D12 = Prix unitaire
    ' D13 = Fichier/document, non traité ici
    ' D14 = Commentaire
    '===========================================================================

    Dim typ As String, categorie As String, sousCategorie As String
    Dim produit As String, sousProduit As String, fournisseur As String
    Dim refValue As String, quantite As Variant, prixUnitaire As Variant
    Dim commentaire As String

    typ = Trim(CStr(Nz(wsInput.Range("D4").Value, "")))
    categorie = Trim(CStr(Nz(wsInput.Range("D5").Value, "")))
    sousCategorie = Trim(CStr(Nz(wsInput.Range("D6").Value, "")))
    produit = Trim(CStr(Nz(wsInput.Range("D7").Value, "")))
    sousProduit = Trim(CStr(Nz(wsInput.Range("D8").Value, "")))
    fournisseur = Trim(CStr(Nz(wsInput.Range("D9").Value, "")))
    refValue = Trim(CStr(Nz(wsInput.Range("D10").Value, "")))
    quantite = wsInput.Range("D11").Value
    prixUnitaire = wsInput.Range("D12").Value
    commentaire = Trim(CStr(Nz(wsInput.Range("D14").Value, "")))

    ' Si aucun sous-produit n'est renseigné, utiliser le produit principal
    If sousProduit = "" Then
        sousProduit = produit
    End If

    '===========================================================================
    ' Validation des champs obligatoires
    '===========================================================================

    If typ = "" Then
        MsgBox "Erreur : le TYPE (D4) est obligatoire.", vbExclamation, "Champ manquant"
        GoTo CleanExit
    End If

    If categorie = "" Then
        MsgBox "Erreur : la CATEGORIE (D5) est obligatoire.", vbExclamation, "Champ manquant"
        GoTo CleanExit
    End If

    If sousCategorie = "" Then
        MsgBox "Erreur : la SOUS-CATEGORIE (D6) est obligatoire.", vbExclamation, "Champ manquant"
        GoTo CleanExit
    End If

    If produit = "" Then
        MsgBox "Erreur : le PRODUIT (D7) est obligatoire.", vbExclamation, "Champ manquant"
        GoTo CleanExit
    End If

    '===========================================================================
    ' Récupération ou génération de l'identifiant de transaction
    '
    ' Si un InputId existe déjà, il est réutilisé.
    ' Sinon, un nouvel identifiant est généré.
    '===========================================================================

    If IsNumeric(wsInput.Range("InputId").Value) And wsInput.Range("InputId").Value <> "" Then
        input_id = CLng(wsInput.Range("InputId").Value)
    Else
        input_id = GetNextInputID()
        wsInput.Range("InputId").Value = input_id
    End If

    '===========================================================================
    ' Recherche des colonnes dans la table input_staging
    '
    ' La recherche se fait par nom d'en-tête afin d'éviter une dépendance
    ' trop forte à l'ordre exact des colonnes.
    '===========================================================================

    colInputID = FindHeaderCol(tbl, "input_id")
    colDate = FindHeaderCol(tbl, "date")
    colType = FindHeaderCol(tbl, "type")
    colCategorieId = FindHeaderCol(tbl, "categorie_id")
    colCategorie = FindHeaderCol(tbl, "categorie")
    colSousCat = FindHeaderCol(tbl, "sous_categorie")
    colProduit = FindHeaderCol(tbl, "produit")
    colProduitId = FindHeaderCol(tbl, "produit_id")
    colSousProduit = FindHeaderCol(tbl, "sous_produit")
    colFournisseur = FindHeaderCol(tbl, "fournisseur")
    colRef = FindHeaderCol(tbl, "ref")
    colQte = FindHeaderCol(tbl, "quantite")
    colPrix = FindHeaderCol(tbl, "prix_unitaire")
    colCommentaire = FindHeaderCol(tbl, "Commentaire")
    colIsDeleted = FindHeaderCol(tbl, "is_deleted")

    '===========================================================================
    ' Ajout d'une nouvelle ligne dans la table de staging
    '===========================================================================

    Set newRow = tbl.ListRows.Add

    Dim r As Long
    r = newRow.Range.Row

    '===========================================================================
    ' Écriture des données saisies
    '===========================================================================

    If colInputID > 0 Then wsStaging.Cells(r, colInputID).Value = input_id
    If colDate > 0 Then wsStaging.Cells(r, colDate).Value = Date
    If colType > 0 Then wsStaging.Cells(r, colType).Value = typ
    If colCategorie > 0 Then wsStaging.Cells(r, colCategorie).Value = categorie
    If colSousCat > 0 Then wsStaging.Cells(r, colSousCat).Value = sousCategorie
    If colProduit > 0 Then wsStaging.Cells(r, colProduit).Value = produit
    If colSousProduit > 0 Then wsStaging.Cells(r, colSousProduit).Value = sousProduit
    If colFournisseur > 0 Then wsStaging.Cells(r, colFournisseur).Value = fournisseur
    If colRef > 0 Then wsStaging.Cells(r, colRef).Value = refValue
    If colQte > 0 Then wsStaging.Cells(r, colQte).Value = quantite
    If colPrix > 0 Then wsStaging.Cells(r, colPrix).Value = prixUnitaire
    If colCommentaire > 0 Then wsStaging.Cells(r, colCommentaire).Value = commentaire

    ' La colonne is_deleted est volontairement laissée vide.
    ' Une valeur vide correspond à une transaction active par défaut.

    '===========================================================================
    ' Formules de recherche pour les identifiants techniques
    '===========================================================================

    If colCategorieId > 0 Then
        wsStaging.Cells(r, colCategorieId).Formula = _
            "=IFERROR(XLOOKUP(" & _
            wsStaging.Cells(r, colCategorie).address(False, False) & _
            ",dim_produits[categorie],dim_produits[categorie_id]),"""")"
    End If

    If colProduitId > 0 Then
        wsStaging.Cells(r, colProduitId).Formula = _
            "=IFERROR(XLOOKUP(" & _
            wsStaging.Cells(r, colCategorie).address(False, False) & "&""|""&" & _
            wsStaging.Cells(r, colSousCat).address(False, False) & "&""|""&" & _
            wsStaging.Cells(r, colProduit).address(False, False) & "," & _
            "dim_produits[categorie]&""|""&dim_produits[sous_categorie]&""|""&dim_produits[produit]," & _
            "dim_produits[produit_id]),"""")"
    End If

    '===========================================================================
    ' Réinitialisation du formulaire après ajout
    '===========================================================================

    On Error Resume Next
    wsInput.Range("upload_status").Value = ""
    wsInput.Range("InputId").ClearContents
    On Error GoTo ErrHandler

    wsInput.Range("D4:D14").ClearContents

    MsgBox "Transaction #" & input_id & " ajoutée avec succès." & vbCrLf & _
           "Sous-produit : " & sousProduit, vbInformation, "Succès"

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    On Error Resume Next

    If Not wsReturn Is Nothing Then wsReturn.Activate

    ThisWorkbook.RefreshAll

    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "Erreur : " & Err.Description, vbCritical, "Erreur"

    On Error Resume Next
    wsInput.Range("upload_status").Value = ""
    On Error GoTo 0

    Resume CleanExit

End Sub

'===============================================================================
' Fonction utilitaire : trouver l'index réel d'une colonne par son nom d'en-tête
'
' Retourne :
' - le numéro de colonne Excel si l'en-tête existe
' - 0 si l'en-tête est introuvable
'===============================================================================

Private Function FindHeaderCol(tbl As ListObject, headerName As String) As Long

    Dim col As ListColumn

    On Error Resume Next
    Set col = tbl.ListColumns(headerName)
    On Error GoTo 0

    If col Is Nothing Then
        FindHeaderCol = 0
    Else
        FindHeaderCol = col.Range.Column
    End If

End Function

'===============================================================================
' Fonction utilitaire : remplacer les valeurs nulles ou vides
'
' Cette fonction reproduit un comportement similaire à Nz dans Access.
' Elle permet d'éviter les erreurs lors de la lecture de cellules vides.
'===============================================================================

Private Function Nz(val As Variant, defaultVal As Variant) As Variant

    If IsNull(val) Or IsEmpty(val) Then
        Nz = defaultVal
    Else
        Nz = val
    End If

End Function

'===============================================================================
' Réinitialiser le formulaire de saisie
'
' Cette macro vide les champs de saisie de la feuille INPUT ainsi que le statut
' d'upload éventuel.
'===============================================================================

Public Sub ResetInputForm()

    Dim wsInput As Worksheet

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Set wsInput = ThisWorkbook.Worksheets("INPUT")

    ' Vide les champs du formulaire de saisie
    wsInput.Range("D4:D14").ClearContents

    ' Réinitialise le statut d'upload si la plage nommée existe
    On Error Resume Next
    wsInput.Range("upload_status").Value = ""
    On Error GoTo 0

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur : " & Err.Description, vbCritical, "Erreur"

End Sub

