VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Settings_3_Add_Typo 
   Caption         =   "Paramètres | Ajouter une Typologie"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4125
   OleObjectBlob   =   "Settings_3_Add_Typo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Settings_3_Add_Typo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
    
    'Déclaration des Variables
    Dim wsTypo As Worksheet
    Dim lastRow As Byte
    Dim pole As String
    Dim checkRange As Range
    
    'Initialisation de la fuille
    Set wsTypo = ThisWorkbook.Sheets("Set_Typo")
    
    'Vérifier si l'affiliation à été sélectionnée
    If checkOpttionButton = True Then
    
        'Définir le paramètre d'affiliation séléctionné
        If OptionButton1 = True Then
            pole = "HG1"
        Else
            pole = "HG2"
        End If
        
        'Définir la dernière ligne et Ajouter la typologie
        If pole = "HG1" Then
            lastRow = wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row
            Set checkRange = wsTypo.Range("A2:A" & lastRow).Find(What:=TextBox1.Value, LookAt:=xlWhole)
            If Not checkRange Is Nothing Then
                MsgBox "Attention cette typologie existe déjà veuillez saisir un nom différent", vbExclamation, "Attention"
                Exit Sub
            End If
            wsTypo.Cells(lastRow + 1, "A").Value = TextBox1.Value
        Else
            lastRow = wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row
            Set checkRange = wsTypo.Range("B2:B" & lastRow).Find(What:=TextBox1.Value, LookAt:=xlWhole)
            If Not checkRange Is Nothing Then
                MsgBox "Attention cette typologie existe déjà veuillez saisir un nom différent", vbExclamation, "Attention"
                Exit Sub
            End If
            wsTypo.Cells(lastRow + 1, "B").Value = TextBox1.Value
        End If
        
        MsgBox "La typologie " & TextBox1.Value & "A bien été enregistré", vbInformation, "Succès"
        
    End If
    
End Sub

Private Function checkOpttionButton() As Boolean
    
    'Définir le paramètre de fonction sur True
    checkOpttionButton = True
    
    'Vérifier si une affiliation a été sélectionnée
    If OptionButton1 = False And OptionButton2 = False Then
        MsgBox "Veuillez sélectionner une affiliation de pôle", vbExclamation, "Attention"
        checkOpttionButton = False
    End If
End Function

Private Sub UserForm_Click()

End Sub
