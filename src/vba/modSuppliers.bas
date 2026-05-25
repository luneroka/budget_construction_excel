Attribute VB_Name = "modSuppliers"
Option Explicit

'===============================================================================
' MODULE : modSuppliers
'
' Rôle :
' - Ajouter un nouveau fournisseur dans la table tbl_fournisseurs.
' - Collecter les informations principales via des boîtes de dialogue.
' - Créer automatiquement un lien cliquable mailto: si un email est renseigné.
' - Mettre à jour le champ Fournisseur dans la feuille INPUT après création.
'
' Compatibilité :
' - Compatible Mac.
'
' Pré-requis :
' - Une feuille nommée "FOURNISSEURS"
' - Une feuille nommée "INPUT"
' - Une table nommée "tbl_fournisseurs"
'
' Structure attendue de la table tbl_fournisseurs :
' - colonne 1 = fournisseur
' - colonne 2 = adresse
' - colonne 3 = contact_principal
' - colonne 4 = tel_contact_principal
' - colonne 5 = mail
' - colonne 6 = contact_secondaire
' - colonne 7 = tel_contact_secondaire
'
' Note GitHub :
' - Les noms de feuilles, tables et colonnes ne sont pas masqués.
' - Ils représentent la structure attendue du fichier Excel.
' - Adapter ces références si la structure du classeur change.
'===============================================================================

Public Sub AddSupplier()

    Dim wsSuppliers As Worksheet
    Dim wsReturn As Worksheet
    Dim tblSuppliers As ListObject
    Dim newRow As ListRow

    Dim supplierName As String
    Dim address As String
    Dim primaryContact As String
    Dim primaryPhone As String
    Dim email As String
    Dim recapText As String

    ' Mémorise la feuille active afin d'y revenir en fin d'exécution
    Set wsReturn = ActiveSheet

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '===========================================================================
    ' Initialisation de la feuille et de la table fournisseurs
    '===========================================================================

    Set wsSuppliers = ThisWorkbook.Worksheets("FOURNISSEURS")
    Set tblSuppliers = wsSuppliers.ListObjects("tbl_fournisseurs")

    If tblSuppliers Is Nothing Then
        MsgBox "Erreur : table 'tbl_fournisseurs' introuvable.", vbCritical, "Erreur"
        GoTo CleanExit
    End If

    '===========================================================================
    ' Saisie des informations fournisseur
    '
    ' Les informations de contact secondaire ne sont pas collectées ici.
    ' Elles peuvent être complétées manuellement dans la table fournisseurs.
    '===========================================================================

    ' 1. Nom du fournisseur, obligatoire
    supplierName = InputBox( _
        "Saisir le nom du fournisseur :" & vbCrLf & vbCrLf & _
        "(Annuler pour quitter)", _
        "Nouveau fournisseur - Étape 1/5" _
    )

    If supplierName = "" Then
        MsgBox "Opération annulée.", vbInformation, "Annulé"
        GoTo CleanExit
    End If

    supplierName = Trim(supplierName)

    If supplierName = "" Then
        MsgBox "Erreur : le nom du fournisseur est obligatoire.", vbExclamation, "Champ manquant"
        GoTo CleanExit
    End If

    ' 2. Adresse, optionnelle
    address = InputBox( _
        "Saisir l'adresse :" & vbCrLf & vbCrLf & _
        "(Optionnel - Annuler pour quitter)", _
        "Nouveau fournisseur - Étape 2/5" _
    )

    If StrPtr(address) = 0 Then
        MsgBox "Opération annulée.", vbInformation, "Annulé"
        GoTo CleanExit
    End If

    address = Trim(address)

    ' 3. Contact principal, optionnel
    primaryContact = InputBox( _
        "Saisir le contact principal :" & vbCrLf & vbCrLf & _
        "(Optionnel - Annuler pour quitter)", _
        "Nouveau fournisseur - Étape 3/5" _
    )

    If StrPtr(primaryContact) = 0 Then
        MsgBox "Opération annulée.", vbInformation, "Annulé"
        GoTo CleanExit
    End If

    primaryContact = Trim(primaryContact)

    ' 4. Téléphone principal, optionnel
    primaryPhone = InputBox( _
        "Saisir le téléphone principal :" & vbCrLf & vbCrLf & _
        "(Optionnel - Annuler pour quitter)", _
        "Nouveau fournisseur - Étape 4/5" _
    )

    If StrPtr(primaryPhone) = 0 Then
        MsgBox "Opération annulée.", vbInformation, "Annulé"
        GoTo CleanExit
    End If

    primaryPhone = Trim(primaryPhone)

    ' 5. Email, optionnel
    email = InputBox( _
        "Saisir l'email :" & vbCrLf & vbCrLf & _
        "(Optionnel - Annuler pour quitter)", _
        "Nouveau fournisseur - Étape 5/5" _
    )

    If StrPtr(email) = 0 Then
        MsgBox "Opération annulée.", vbInformation, "Annulé"
        GoTo CleanExit
    End If

    email = Trim(email)

    '===========================================================================
    ' Récapitulatif avant ajout
    '===========================================================================

    recapText = "Vérifiez les informations :" & vbCrLf & vbCrLf & _
        "Fournisseur : " & supplierName & vbCrLf & _
        "Adresse : " & IIf(address = "", "(vide)", address) & vbCrLf & _
        "Contact principal : " & IIf(primaryContact = "", "(vide)", primaryContact) & vbCrLf & _
        "Téléphone principal : " & IIf(primaryPhone = "", "(vide)", primaryPhone) & vbCrLf & _
        "Email : " & IIf(email = "", "(vide)", email) & vbCrLf & vbCrLf & _
        "Confirmer l'ajout ?"

    If MsgBox(recapText, vbYesNo + vbQuestion, "Confirmation") = vbNo Then
        MsgBox "Opération annulée. Aucune donnée ajoutée.", vbInformation, "Annulé"
        GoTo CleanExit
    End If

    '===========================================================================
    ' Insertion du nouveau fournisseur dans tbl_fournisseurs
    '===========================================================================

    Set newRow = tblSuppliers.ListRows.Add

    With newRow.Range

        .Cells(1, 1).Value = supplierName
        .Cells(1, 2).Value = address
        .Cells(1, 3).Value = primaryContact
        .Cells(1, 4).Value = primaryPhone

        ' Ajoute l'email sous forme de lien cliquable mailto:
        If email <> "" Then
            wsSuppliers.Hyperlinks.Add _
                Anchor:=.Cells(1, 5), _
                address:="mailto:" & email, _
                TextToDisplay:=email
        End If

        ' Les colonnes suivantes sont laissées vides pour saisie manuelle :
        ' .Cells(1, 6) = contact_secondaire
        ' .Cells(1, 7) = tel_contact_secondaire

    End With

    ' Met à jour le champ Fournisseur du formulaire INPUT
    ThisWorkbook.Worksheets("INPUT").Range("D9").Value = supplierName

    MsgBox "Fournisseur '" & supplierName & "' ajouté avec succès.", vbInformation, "Succès"

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

