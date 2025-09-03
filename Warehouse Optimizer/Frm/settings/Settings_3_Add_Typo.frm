VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Settings_3_Add_Typo 
   Caption         =   "Param�tres | Ajouter une Typologie"
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
    
    'D�claration des Variables
    Dim wsTypo As Worksheet
    Dim lastRow As Byte
    Dim pole As String
    Dim checkRange As Range
    
    'Initialisation de la fuille
    Set wsTypo = ThisWorkbook.Sheets("Set_Typo")
    
    'V�rifier si l'affiliation � �t� s�lectionn�e
    If checkOpttionButton = True Then
    
        'D�finir le param�tre d'affiliation s�l�ctionn�
        If OptionButton1 = True Then
            pole = "HG1"
        Else
            pole = "HG2"
        End If
        
        'D�finir la derni�re ligne et Ajouter la typologie
        If pole = "HG1" Then
            lastRow = wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row
            Set checkRange = wsTypo.Range("A2:A" & lastRow).Find(What:=TextBox1.Value, LookAt:=xlWhole)
            If Not checkRange Is Nothing Then
                MsgBox "Attention cette typologie existe d�j� veuillez saisir un nom diff�rent", vbExclamation, "Attention"
                Exit Sub
            End If
            wsTypo.Cells(lastRow + 1, "A").Value = TextBox1.Value
        Else
            lastRow = wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row
            Set checkRange = wsTypo.Range("B2:B" & lastRow).Find(What:=TextBox1.Value, LookAt:=xlWhole)
            If Not checkRange Is Nothing Then
                MsgBox "Attention cette typologie existe d�j� veuillez saisir un nom diff�rent", vbExclamation, "Attention"
                Exit Sub
            End If
            wsTypo.Cells(lastRow + 1, "B").Value = TextBox1.Value
        End If
        
        MsgBox "La typologie " & TextBox1.Value & "A bien �t� enregistr�", vbInformation, "Succ�s"
        
    End If
    
End Sub

Private Function checkOpttionButton() As Boolean
    
    'D�finir le param�tre de fonction sur True
    checkOpttionButton = True
    
    'V�rifier si une affiliation a �t� s�lectionn�e
    If OptionButton1 = False And OptionButton2 = False Then
        MsgBox "Veuillez s�lectionner une affiliation de p�le", vbExclamation, "Attention"
        checkOpttionButton = False
    End If
End Function

Private Sub UserForm_Click()

End Sub
