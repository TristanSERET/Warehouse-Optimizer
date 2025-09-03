VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Clear_2_Select_Cell 
   Caption         =   "Sélection de la cellule"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9420.001
   OleObjectBlob   =   "Clear_2_Select_Cell.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Clear_2_Select_Cell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    'Initialisation de la combobox
    With ComboBox1
        .AddItem "Cellule_A"
        .AddItem "Cellule_B"
        .AddItem "Cellule_E"
        .AddItem "Cellule_F"
        .AddItem "Cellule_G"
    End With

End Sub

Private Sub Image1_Click()

    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim cellSelected As String
    Dim clearRange As Range
    Dim cell As Range
    
    'Déactiver l'affichage pendant le traitement
    Application.ScreenUpdating = False
    
    'Initialisation de la feuille
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    
    'Définir la cellule sélectionnée
    cellSelected = ComboBox1.Value
    
    'Initialisation de la plage de suppression en fonction de la cellule
    Select Case cellSelected
        Case "Cellule_A": Set clearRange = wsimplant.Range("ES3:FX90")
        Case "Cellule_B": Set clearRange = wsimplant.Range("DJ3:EO98")
        Case "Cellule_E": Set clearRange = wsimplant.Range("CA3:DF90")
        Case "Cellule_F": Set clearRange = wsimplant.Range("AQ3:BV98")
        Case "Cellule_G": Set clearRange = wsimplant.Range("E3:AJ92")
        Case Else: MsgBox "Aucune cellule n'a été sélectionnée", vbExclamation, "Error": Exit Sub
    End Select
    
    'Parcourrir la plage de données et supprimer les couleurs + data si les conditions sont remplis
    For Each cell In clearRange
        If cell.Interior.Color <> RGB(217, 217, 217) Then
            cell.Interior.Color = vbWhite
            cell.ClearContents
        End If
        If cell.Interior.Pattern = xlLightDown Then
            cell.Interior.Pattern = xlNone
        End If
    Next cell
    
    'Activer l'affichage
    Application.ScreenUpdating = True
    
    'Message de confirmation
    MsgBox "Le cash de la " & cellSelected & " à bien été vidé", vbInformation, "Succès"
    
    'Fermer le formulaire
    Unload Me
    
End Sub
