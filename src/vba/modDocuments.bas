Attribute VB_Name = "modDocuments"
Option Explicit

Sub OpenUploadPage()

    Dim baseUrl As String
    Dim url As String
    Dim input_id As Long
    Dim sousProduit As String

    ' ------------------------------------------------------------
    ' CONFIGURATION UTILISATEUR
    ' ------------------------------------------------------------
    ' Remplacer <GOOGLE_APPS_SCRIPT_DEPLOYMENT_ID> par l'ID réel
    ' du déploiement Google Apps Script.
    '
    ' Exemple d'URL générée par Google Apps Script :
    ' https://script.google.com/macros/s/AKfycbxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx/dev
    '
    ' Dans ce cas, ne remplacer que la partie :
    ' <GOOGLE_APPS_SCRIPT_DEPLOYMENT_ID>
    ' ------------------------------------------------------------
    baseUrl = "https://script.google.com/macros/s/<GOOGLE_APPS_SCRIPT_DEPLOYMENT_ID>/dev"

    ' Génère un nouvel identifiant unique pour la transaction en cours
    input_id = GetNextInputID()

    ' Stocke l'identifiant généré dans la feuille INPUT
    ThisWorkbook.Sheets("INPUT").Range("InputId").Value = input_id

    With ThisWorkbook.Sheets("INPUT")

        ' Empêche l'ajout de plusieurs PDF pour une même transaction
        If .Range("upload_status").Value = "UPLOADING" Then

            MsgBox _
                "Un fichier PDF a déjà été ajouté pour cette transaction." & vbCrLf & vbCrLf & _
                "Veuillez ajouter la transaction avant d'importer un nouveau fichier.", _
                vbExclamation, _
                "Upload déjà effectué"

            Exit Sub

        End If

        ' Vérifie que les informations obligatoires sont bien renseignées
        If Trim(.Range("D4").Value) = "" _
            Or Trim(.Range("D9").Value) = "" _
            Or Trim(.Range("D10").Value) = "" Then

            MsgBox _
                "Impossible d'ajouter un fichier : certaines informations sont manquantes." & vbCrLf & vbCrLf & _
                "Veuillez vérifier : type, fournisseur et référence.", _
                vbExclamation, _
                "Informations incomplètes"

            Exit Sub

        End If

        ' Si aucun sous-produit n'est renseigné, utilise le produit principal
        sousProduit = Trim(CStr(.Range("D8").Value))

        If sousProduit = "" Then
            sousProduit = Trim(CStr(.Range("D7").Value))
        End If

        ' Verrouille temporairement la transaction pour éviter les doublons
        .Range("upload_status").Value = "UPLOADING"

        ' Construit l'URL d'ouverture de la page Google Apps Script
        url = baseUrl & _
            "?input_id=" & URLEncode(CStr(input_id)) & _
            "&docType=" & URLEncode(CStr(.Range("D4").Value)) & _
            "&fournisseur=" & URLEncode(CStr(.Range("D9").Value)) & _
            "&ref=" & URLEncode(CStr(.Range("D10").Value)) & _
            "&sous_produit=" & URLEncode(sousProduit)

    End With

    ' Affiche l'URL générée dans la fenêtre d'exécution VBA pour debug
    Debug.Print url

    ' Ouvre la page d'upload dans le navigateur
    ThisWorkbook.FollowHyperlink url

End Sub

Public Function URLEncode(ByVal Text As String) As String

    Dim i As Long
    Dim CharCode As Long
    Dim Char As String
    Dim Result As String

    For i = 1 To Len(Text)

        Char = Mid$(Text, i, 1)
        CharCode = AscW(Char)

        Select Case CharCode

            ' Lettres et chiffres : a-z, A-Z, 0-9
            Case 48 To 57, 65 To 90, 97 To 122
                Result = Result & Char

            ' Caractères autorisés dans une URL
            Case 45, 46, 95, 126
                Result = Result & Char

            ' Espaces
            Case 32
                Result = Result & "%20"

            ' Encodage UTF-8 pour les autres caractères
            Case Else

                If CharCode < 128 Then

                    Result = Result & "%" & _
                        Right$("0" & Hex(CharCode), 2)

                ElseIf CharCode < 2048 Then

                    Result = Result & "%" & _
                        Right$("0" & Hex((CharCode \ 64) Or 192), 2)

                    Result = Result & "%" & _
                        Right$("0" & Hex((CharCode And 63) Or 128), 2)

                Else

                    Result = Result & "%" & _
                        Right$("0" & Hex((CharCode \ 4096) Or 224), 2)

                    Result = Result & "%" & _
                        Right$("0" & Hex(((CharCode \ 64) And 63) Or 128), 2)

                    Result = Result & "%" & _
                        Right$("0" & Hex((CharCode And 63) Or 128), 2)

                End If

        End Select

    Next i

    URLEncode = Result

End Function

