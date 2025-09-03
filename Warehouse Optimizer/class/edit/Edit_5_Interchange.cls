VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Edit_5_Interchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe d'échange sur les emplacements

'Implémentation de la classe mère
Implements Edit_1_Main

Private Sub Edit_1_Main_SelectEdit()
    
    'Déclaration des Variables
    Dim selectRange As Range
    Dim startArea As Range
    Dim endArea As Range
    Dim Index As Variant
    Dim temp As Variant
    Dim tempColor As Variant
    
    'Initialisation de la plage sélectionnée
    Set selectRange = Selection
    
    'Définir les plages à interchanger
    Set startArea = selectRange.Areas(1)
    Set endArea = selectRange.Areas(2)
    
    'Vérifier que les plages ont la même taille
    If startArea.Cells.Count <> endArea.Cells.Count Then
        MsgBox "Attention les pages sélectionnées n'ont pas la même taille !", vbExclamation, "Error"
        Exit Sub
    End If
    
    'Echanger les données
    For Index = 1 To startArea.Cells.Count
        temp = startArea.Cells(Index).Value
        tempColor = startArea.Cells(Index).Interior.Color
        startArea.Cells(Index).Value = endArea.Cells(Index).Value
        startArea.Cells(Index).Interior.Color = endArea.Cells(Index).Interior.Color
        endArea.Cells(Index).Value = temp
        endArea.Cells(Index).Interior.Color = tempColor
    Next Index
    
End Sub
