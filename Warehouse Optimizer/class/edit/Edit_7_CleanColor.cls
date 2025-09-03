VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Edit_7_CleanColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe d'édition pour nettoyer la couleur

'Implémentation de la classe mère
Implements Edit_1_Main

Private Sub Edit_1_Main_SelectEdit()

    'Déclaration des Variables
    Dim selectRange As Range
    Dim cell As Range
    Dim reponse As VbMsgBoxResult
    'Initialisation de la plage sélectionnée
    Set selectRange = Selection
    
    'Vérifier si une plage est sélectionnée
    If TypeName(selectRange) = "Range" Then
        For Each cell In selectRange
            If cell.Interior.Color = RGB(217, 217, 217) Then
                reponse = MsgBox("Votre sélection inclue une AST. Voulez-vous continuer ?", vbYesNo + vbQuestion, "Confirmation")
                If reponse = vbNo Then
                    Exit Sub
                End If
            End If
        Next cell
               
        'Nettoyer les couleurs
        selectRange.Interior.ColorIndex = xlNone
    End If

End Sub
