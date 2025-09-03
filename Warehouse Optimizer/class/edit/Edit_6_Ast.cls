VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Edit_6_Ast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe d'édition des AST

'Implémentation de la classe mère
Implements Edit_1_Main

Private Sub Edit_1_Main_SelectEdit()

    'Déclaration des Variables
    Dim selectRange As Range
    
    'Initialisation de la plage sélectionnée
    Set selectRange = Selection
    
    'Vérifier si une plage est sélectionnée
    If TypeName(selectRange) = "Range" Then
        'Supprimer les diagonales
        selectRange.Borders(xlDiagonalDown).LineStyle = xlNone
        selectRange.Borders(xlDiagonalUp).LineStyle = xlNone
        
        ' Supprimer les bordures intérieures
        selectRange.Borders(xlInsideVertical).LineStyle = xlNone
        selectRange.Borders(xlInsideHorizontal).LineStyle = xlNone
        
        'Supprimer le motif intérieur
        selectRange.Interior.Pattern = xlNone
        
        'Appliquer la couleur de l'AST
        selectRange.Interior.Color = RGB(217, 217, 217)
    End If
        
End Sub


