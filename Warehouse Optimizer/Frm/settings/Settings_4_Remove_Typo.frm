VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Settings_4_Remove_Typo 
   Caption         =   "Paramètres | Supprimer une Typologie"
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

    'Déclaration des Variables
    Dim wsTypo As Worksheet
    Dim lastRow As Byte
    Dim rowIndex As Byte
    
    'Initialisation de la feuille
    Set wsTypo = ThisWorkbook.Sheets("Set_Typo")
    
    'Définir la dernière ligne
    lastRow = wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row
    
    'Mécanique de déactivation des CheckBox et Remplissage de la ComboBox
    If CheckBox1 = True Then
        CheckBox2 = False
        ComboBox1.Clear
        For rowIndex = 2 To lastRow
            ComboBox1.AddItem wsTypo.Cells(rowIndex, "A").Value
        Next rowIndex
    End If
    
End Sub

Private Sub CheckBox2_Click()

    'Déclaration des Variables
    Dim wsTypo As Worksheet
    Dim lastRow As Byte
    Dim rowIndex As Byte
    
        'Initialisation de la feuille
        Set wsTypo = ThisWorkbook.Sheets("Set_Typo")
        
        'Définir la dernière ligne
        lastRow = wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row
        
        'Mécanique de déactivation des CheckBox
        If CheckBox2 = True Then
            CheckBox1 = False
            ComboBox1.Clear
            For rowIndex = 2 To lastRow
                ComboBox1.AddItem wsTypo.Cells(rowIndex, "B").Value
            Next rowIndex
        End If
        
End Sub

Private Sub Image1_Click()

    'Déclaration des Variables
    Dim wsTypo As Worksheet
    Dim lastRow As Byte
    Dim rng As Range

    'Vérifier si au moin une case et coché
    If checkPole = True Then
    
        'Initialisation de la feuille
        Set wsTypo = ThisWorkbook.Sheets("Set_Typo")
        
        'Définir la dernière ligne
        If CheckBox1 = True Then
            lastRow = wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row
            'Supprimer la typologie désigné
            Set rng = wsTypo.Range("A1:A" & lastRow).Find(What:=ComboBox1.Value, LookAt:=xlWhole)
            rng.ClearContents
        Else
            lastRow = wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row
            'Supprimer la typologie désigné
            Set rng = wsTypo.Range("B1:B" & lastRow).Find(What:=ComboBox1.Value, LookAt:=xlWhole)
            rng.ClearContents
        End If
        
        MsgBox "La typologie " & ComboBox1.Value & " à bien été supprimé", vbInformation, "Succès"
    End If
    
End Sub

Private Function checkPole() As Boolean

    'Affecer le paramètre de la fonction sur True par défaut
    checkPole = True
    
    'Identifier si aucune case n'est coché
    If CheckBox1 = False And CheckBox2 = False Then
        MsgBox "Veuillez sélectionner un pôle", vbExclamation, "Attention"
        checkPole = False
    End If
End Function

Private Sub UserForm_Click()

End Sub
