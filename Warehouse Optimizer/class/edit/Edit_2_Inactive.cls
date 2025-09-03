VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Edit_2_Inactive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe d'édition des emplacements inactifs

'Implémentation de la classe mère
Implements Edit_1_Main

Private Sub Edit_1_Main_SelectEdit()

    'Déclaration des Variables
    Dim cell As Range
    Dim selectRange As Range
    Dim isOk As Boolean
    
    'Initialisation de la plage sélectionnée
    Set selectRange = Selection
    
    'Vérifier si une plage est sélectionnée
    If TypeName(selectRange) = "Range" Then
        isOk = True
        For Each cell In selectRange
            If cell.Interior.Color = RGB(217, 217, 217) Then
                isOk = False
            End If
        Next cell
        
        If isOk = True Then
            selectRange.Borders(xlDiagonalDown).LineStyle = xlContinuous
            selectRange.Borders(xlDiagonalUp).LineStyle = xlContinuous
        Else
            MsgBox "La sélection n'est pas valide.", vbExclamation, "Sélection incorrect"
        End If
    End If
            
End Sub
