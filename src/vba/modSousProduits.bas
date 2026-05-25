Attribute VB_Name = "modSousProduits"
Option Explicit

'===============================================================================
' MODULE : modSousProduits
'
' Rôle :
' - Ajouter un nouveau sous-produit dans la table tbl_sous_produits.
' - Associer ce sous-produit à la catégorie, sous-catégorie et produit sélectionnés.
' - Générer automatiquement un sous_produit_id.
' - Renseigner automatiquement categorie_id et produit_id depuis dim_produits.
' - Mettre à jour le champ Sous-produit dans la feuille INPUT après création.
'
' Pré-requis :
' - Une feuille nommée "SOUS-PRODUITS"
' - Une feuille nommée "INPUT"
' - Une feuille nommée "PRODUITS"
' - Une table nommée "tbl_sous_produits"
' - Une table nommée "dim_produits"
'
' Structure attendue dans la feuille INPUT :
' - D5 = Catégorie
' - D6 = Sous-catégorie
' - D7 = Produit
' - D8 = Sous-produit
'
' Structure attendue de la table dim_produits :
' - colonne 1 = categorie_id
' - colonne 2 = categorie
' - colonne 3 = sous_categorie
' - colonne 4 = produit
' - colonne 5 = produit_id
' - colonne 6 = unite
'
' Structure attendue de la table tbl_sous_produits :
' - colonne 1 = sous_produit_id
' - colonne 2 = categorie_id
' - colonne 3 = categorie
' - colonne 4 = sous_categorie
' - colonne 5 = produit_id
' - colonne 6 = produit
' - colonne 7 = sous_produit
' - colonne 8 = is_active
'
' Note GitHub :
' - Les noms de feuilles, tables, colonnes et cellules ne sont pas masqués.
' - Ils représentent la structure attendue du fichier Excel.
' - Adapter ces références si la structure du classeur change.
'===============================================================================

Public Sub AddSousProduit()

    Dim wsSousProduits As Worksheet
    Dim wsInput As Worksheet
    Dim wsDimProduits As Worksheet
    Dim wsReturn As Worksheet

    Dim tblSousProduits As ListObject
    Dim tblDimProduits As ListObject
    Dim newRow As ListRow

    Dim sousProduitName As String
    Dim categorie As String
    Dim sousCategorie As String
    Dim produit As String
    Dim categorieId As Variant
    Dim produitId As Variant
    Dim newSousProduitId As Long
    Dim recapText As String

    Dim i As Long
    Dim found As Boolean
    Dim alreadyExists As Boolean

    ' Mémorise la feuille active afin d'y revenir en fin d'exécution
    Set wsReturn = ActiveSheet

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '===========================================================================
    ' Initialisation des feuilles et des tables
    '===========================================================================

    Set wsSousProduits = ThisWorkbook.Worksheets("SOUS-PRODUITS")
    Set wsInput = ThisWorkbook.Worksheets("INPUT")
    Set wsDimProduits = ThisWorkbook.Worksheets("PRODUITS")

    Set tblSousProduits = wsSousProduits.ListObjects("tbl_sous_produits")
    Set tblDimProduits = wsDimProduits.ListObjects("dim_produits")

    If tblSousProduits Is Nothing Then
        MsgBox "Erreur : table 'tbl_sous_produits' introuvable.", vbCritical, "Erreur"
        GoTo CleanExit
    End If

    If tblDimProduits Is Nothing Then
        MsgBox "Erreur : table 'dim_produits' introuvable.", vbCritical, "Erreur"
        GoTo CleanExit
    End If

    '===========================================================================
    ' Lecture des champs obligatoires depuis la feuille INPUT
    '
    ' D5 = Catégorie
    ' D6 = Sous-catégorie
    ' D7 = Produit
    '===========================================================================

    categorie = Trim(CStr(wsInput.Range("D5").Value))

    If categorie = "" Then
        MsgBox "Erreur : le champ CATEGORIE (INPUT!D5) est obligatoire." & vbCrLf & _
               "Veuillez remplir ce champ avant de continuer.", vbExclamation, "Champ manquant"
        GoTo CleanExit
    End If

    sousCategorie = Trim(CStr(wsInput.Range("D6").Value))

    If sousCategorie = "" Then
        MsgBox "Erreur : le champ SOUS-CATEGORIE (INPUT!D6) est obligatoire." & vbCrLf & _
               "Veuillez remplir ce champ avant de continuer.", vbExclamation, "Champ manquant"
        GoTo CleanExit
    End If

    produit = Trim(CStr(wsInput.Range("D7").Value))

    If produit = "" Then
        MsgBox "Erreur : le champ PRODUIT (INPUT!D7) est obligatoire." & vbCrLf & _
               "Veuillez remplir ce champ avant de continuer.", vbExclamation, "Champ manquant"
        GoTo CleanExit
    End If

    '===========================================================================
    ' Recherche des identifiants categorie_id et produit_id depuis dim_produits
    '
    ' Le sous-produit doit obligatoirement être rattaché à un produit existant.
    '===========================================================================

    categorieId = ""
    produitId = ""
    found = False

    If tblDimProduits.ListRows.Count > 0 Then

        For i = 1 To tblDimProduits.ListRows.Count

            ' Correspondance sur le produit sélectionné
            If Trim(CStr(tblDimProduits.ListRows(i).Range.Cells(1, 4).Value)) = produit Then
                categorieId = tblDimProduits.ListRows(i).Range.Cells(1, 1).Value
                produitId = tblDimProduits.ListRows(i).Range.Cells(1, 5).Value
                found = True
                Exit For
            End If

        Next i

    End If

    If Not found Then
        MsgBox "Erreur : produit '" & produit & "' introuvable dans dim_produits." & vbCrLf & _
               "Impossible de créer un sous-produit sans produit valide.", vbCritical, "Produit introuvable"
        GoTo CleanExit
    End If

    '===========================================================================
    ' Saisie du nom du nouveau sous-produit
    '===========================================================================

    sousProduitName = InputBox( _
        "Saisir le nom du sous-produit :" & vbCrLf & vbCrLf & _
        "Catégorie : " & categorie & vbCrLf & _
        "Sous-catégorie : " & sousCategorie & vbCrLf & _
        "Produit : " & produit & vbCrLf & vbCrLf & _
        "(Annuler pour quitter)", _
        "Nouveau sous-produit" _
    )

    If sousProduitName = "" Then
        MsgBox "Opération annulée.", vbInformation, "Annulé"
        GoTo CleanExit
    End If

    sousProduitName = Trim(sousProduitName)

    If sousProduitName = "" Then
        MsgBox "Erreur : le nom du sous-produit ne peut pas être vide.", vbExclamation, "Champ manquant"
        GoTo CleanExit
    End If

    '===========================================================================
    ' Vérification des doublons
    '
    ' Un même sous-produit ne doit pas exister deux fois pour le même produit_id.
    ' La comparaison est insensible à la casse.
    '===========================================================================

    alreadyExists = False

    If tblSousProduits.ListRows.Count > 0 Then

        For i = 1 To tblSousProduits.ListRows.Count

            If CStr(tblSousProduits.ListRows(i).Range.Cells(1, 5).Value) = CStr(produitId) And _
               LCase(Trim(CStr(tblSousProduits.ListRows(i).Range.Cells(1, 7).Value))) = LCase(sousProduitName) Then

                alreadyExists = True
                Exit For

            End If

        Next i

    End If

    If alreadyExists Then
        MsgBox "Erreur : le sous-produit '" & sousProduitName & "' existe déjà pour ce produit." & vbCrLf & _
               "Produit : " & produit & " (ID : " & produitId & ")" & vbCrLf & vbCrLf & _
               "Veuillez choisir un autre nom.", vbExclamation, "Doublon détecté"
        GoTo CleanExit
    End If

    '===========================================================================
    ' Génération du nouvel identifiant sous_produit_id
    '
    ' L'identifiant est calculé à partir du plus grand ID existant + 1.
    '===========================================================================

    If tblSousProduits.ListRows.Count = 0 Then
        newSousProduitId = 1
    Else

        Dim maxId As Long
        maxId = 0

        For i = 1 To tblSousProduits.ListRows.Count

            If IsNumeric(tblSousProduits.ListRows(i).Range.Cells(1, 1).Value) Then

                If CLng(tblSousProduits.ListRows(i).Range.Cells(1, 1).Value) > maxId Then
                    maxId = CLng(tblSousProduits.ListRows(i).Range.Cells(1, 1).Value)
                End If

            End If

        Next i

        newSousProduitId = maxId + 1

    End If

    '===========================================================================
    ' Récapitulatif avant ajout
    '===========================================================================

    recapText = "Vérifiez les informations du nouveau sous-produit :" & vbCrLf & vbCrLf & _
        "Sous-produit ID : " & newSousProduitId & vbCrLf & _
        "Catégorie ID : " & categorieId & vbCrLf & _
        "Catégorie : " & categorie & vbCrLf & _
        "Sous-catégorie : " & sousCategorie & vbCrLf & _
        "Produit ID : " & produitId & vbCrLf & _
        "Produit : " & produit & vbCrLf & _
        "Sous-produit : " & sousProduitName & vbCrLf & _
        "Actif : Oui" & vbCrLf & vbCrLf & _
        "Confirmer l'ajout ?"

    If MsgBox(recapText, vbYesNo + vbQuestion, "Confirmation") = vbNo Then
        MsgBox "Opération annulée. Aucune donnée ajoutée.", vbInformation, "Annulé"
        GoTo CleanExit
    End If

    '===========================================================================
    ' Insertion du nouveau sous-produit dans tbl_sous_produits
    '===========================================================================

    Set newRow = tblSousProduits.ListRows.Add

    With newRow.Range
        .Cells(1, 1).Value = newSousProduitId
        .Cells(1, 2).Value = categorieId
        .Cells(1, 3).Value = categorie
        .Cells(1, 4).Value = sousCategorie
        .Cells(1, 5).Value = produitId
        .Cells(1, 6).Value = produit
        .Cells(1, 7).Value = sousProduitName
        .Cells(1, 8).Value = True
    End With

    '===========================================================================
    ' Rafraîchissement des calculs et listes dépendantes
    '===========================================================================

    Application.Calculate

    ' Met à jour la cellule Sous-produit du formulaire INPUT
    wsInput.Range("D8").Value = sousProduitName

    MsgBox "Sous-produit '" & sousProduitName & "' (ID : " & newSousProduitId & ") ajouté avec succès." & vbCrLf & _
           "Le champ Sous-produit dans INPUT a été mis à jour.", vbInformation, "Succès"

CleanExit:

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ' Retourne l'utilisateur sur la feuille d'origine
    On Error Resume Next

    If Not wsReturn Is Nothing Then
        wsReturn.Activate
    Else
        ThisWorkbook.Worksheets("INPUT").Activate
    End If

    On Error GoTo 0
    Exit Sub

CleanFail:

    MsgBox "Erreur : " & Err.Description, vbCritical, "Erreur"
    Resume CleanExit

End Sub

