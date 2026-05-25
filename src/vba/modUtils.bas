Attribute VB_Name = "modUtils"
Option Explicit

'===============================================================================
' MODULE : modUtils
'
' Rôle :
' - Regrouper les fonctions utilitaires utilisées par plusieurs macros.
' - Afficher la configuration du classeur.
' - Rechercher une colonne par son en-tête dans une table structurée.
' - Générer ou récupérer le prochain identifiant de transaction.
'
' Pré-requis :
' - Une feuille nommée "input_staging"
' - Une table nommée "input_staging"
' - Une colonne nommée "input_id"
'
' Note GitHub :
' - DRIVE_FOLDER_PATH doit être défini dans un module de configuration séparé.
' - Pour publier le projet, remplacer le chemin réel par un placeholder.
'===============================================================================

'===============================================================================
' Macro : afficher la configuration actuelle
'
' Cette macro affiche le chemin du dossier Google Drive utilisé par le classeur.
'
' Remplacer <GOOGLE_DRIVE_FOLDER_PATH> dans le module de configuration par
' le chemin réel du dossier local synchronisé avec Google Drive.
'
' Exemple :
' Public Const DRIVE_FOLDER_PATH As String = "<GOOGLE_DRIVE_FOLDER_PATH>"
'===============================================================================

Public Sub ShowConfiguration()

    MsgBox _
        "Chemin Google Drive :" & vbCrLf & vbCrLf & _
        DRIVE_FOLDER_PATH, _
        vbInformation, _
        "Configuration"

End Sub

'===============================================================================
' Fonction utilitaire : rechercher une colonne par son nom d'en-tête
'
' Paramètres :
' - tbl : table structurée Excel dans laquelle rechercher
' - headerName : nom de l'en-tête à trouver
'
' Retourne :
' - l'index de la colonne dans la table si trouvée
' - 0 si aucune colonne ne correspond
'
' La comparaison est insensible à la casse et ignore les espaces avant/après.
'===============================================================================

Public Function FindHeaderCol(tbl As ListObject, headerName As String) As Long

    Dim i As Long

    For i = 1 To tbl.ListColumns.Count

        If LCase(Trim(tbl.ListColumns(i).Name)) = LCase(Trim(headerName)) Then
            FindHeaderCol = tbl.ListColumns(i).Index
            Exit Function
        End If

    Next i

    FindHeaderCol = 0

End Function

'===============================================================================
' Fonction utilitaire : générer un identifiant temporaire en mémoire
'
' Attention :
' - Cette fonction utilise une variable Static.
' - L'identifiant est réinitialisé lorsque le projet VBA est réinitialisé.
' - Pour les transactions réelles, préférer GetNextInputID.
'
' Usage recommandé :
' - tests rapides
' - génération temporaire
' - cas où la table input_staging n'est pas encore disponible
'===============================================================================

Public Function GenerateInputID() As Long

    Static lastID As Long

    If lastID = 0 Then
        lastID = 1
    Else
        lastID = lastID + 1
    End If

    GenerateInputID = lastID

End Function

'===============================================================================
' Fonction utilitaire : récupérer le prochain input_id disponible
'
' Cette fonction lit la colonne input_id de la table input_staging,
' trouve l'identifiant le plus élevé, puis retourne max + 1.
'
' En cas d'erreur ou de table vide, la fonction retourne 1.
'===============================================================================

Public Function GetNextInputID() As Long

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim idRange As Range
    Dim idCell As Range
    Dim maxInputId As Long

    On Error GoTo SafeExit

    Set ws = ThisWorkbook.Worksheets("input_staging")
    Set tbl = ws.ListObjects("input_staging")

    maxInputId = 0

    On Error Resume Next
    Set idRange = tbl.ListColumns("input_id").DataBodyRange
    On Error GoTo SafeExit

    If Not idRange Is Nothing Then

        For Each idCell In idRange

            If IsNumeric(idCell.Value) Then

                If CLng(idCell.Value) > maxInputId Then
                    maxInputId = CLng(idCell.Value)
                End If

            End If

        Next idCell

    End If

SafeExit:

    GetNextInputID = maxInputId + 1

End Function

