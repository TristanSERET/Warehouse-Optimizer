VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Settings_4_Remove_Typo 
   Caption         =   "Param�tres | Supprimer une Typologie"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4110
   OleObjectBlob   =   "Settings_4_Remove_Typo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Settings_4_Remove_Typo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

    'D�claration des Variables
    Dim wsTypo As Worksheet
    Dim lastRow As Byte
    Dim rowIndex As Byte
    
    'Initialisation de la feuille
    Set wsTypo = ThisWorkbook.Sheets("Set_Typo")
    
    'D�finir la derni�re ligne
    lastRow = wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row
    
    'M�canique de d�activation des CheckBox et Remplissage de la ComboBox
    If CheckBox1 = True Then
        CheckBox2 = False
        ComboBox1.Clear
        For rowIndex = 2 To lastRow
            ComboBox1.AddItem wsTypo.Cells(rowIndex, "A").Value
        Next rowIndex
    End If
    
End Sub

Private Sub CheckBox2_Click()

    'D�claration des Variables
    Dim wsTypo As Worksheet
    Dim lastRow As Byte
    Dim rowIndex As Byte
    
        'Initialisation de la feuille
        Set wsTypo = ThisWorkbook.Sheets("Set_Typo")
        
        'D�finir la derni�re ligne
        lastRow = wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row
        
        'M�canique de d�activation des CheckBox
        If CheckBox2 = True Then
            CheckBox1 = False
            ComboBox1.Clear
            For rowIndex = 2 To lastRow
                ComboBox1.AddItem wsTypo.Cells(rowIndex, "B").Value
            Next rowIndex
        End If
        
End Sub

Private Sub Image1_Click()

    'D�claration des Variables
    Dim wsTypo As Worksheet
    Dim lastRow As Byte
    Dim rng As Range

    'V�rifier si au moin une case et coch�
    If checkPole = True Then
    
        'Initialisation de la feuille
        Set wsTypo = ThisWorkbook.Sheets("Set_Typo")
        
        'D�finir la derni�re ligne
        If CheckBox1 = True Then
            lastRow = wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row
            'Supprimer la typologie d�sign�
            Set rng = wsTypo.Range("A1:A" & lastRow).Find(What:=ComboBox1.Value, LookAt:=xlWhole)
            rng.ClearContents
        Else
            lastRow = wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row
            'Supprimer la typologie d�sign�
            Set rng = wsTypo.Range("B1:B" & lastRow).Find(What:=ComboBox1.Value, LookAt:=xlWhole)
            rng.ClearContents
        End If
        
        MsgBox "La typologie " & ComboBox1.Value & " � bien �t� supprim�", vbInformation, "Succ�s"
    End If
    
End Sub

Private Function checkPole() As Boolean

    'Affecer le param�tre de la fonction sur True par d�faut
    checkPole = True
    
    'Identifier si aucune case n'est coch�
    If CheckBox1 = False And CheckBox2 = False Then
        MsgBox "Veuillez s�lectionner un p�le", vbExclamation, "Attention"
        checkPole = False
    End If
End Function

Private Sub UserForm_Click()

End Sub
