Attribute VB_Name = "Edit_1_Button"
Sub Edit_Button()

    'Déclaration des Variables
    Dim selectButton As String

    'Quel bouton sélectionné
    selectButton = Application.Caller
    
    'Déclencher l'action en fonction du bouton
    Select Case selectButton
        Case "EmpInactive": Set Editor = New Edit_2_Inactive: Editor.SelectEdit
        Case "EmpActive": Set Editor = New Edit_3_Active: Editor.SelectEdit
        Case "EmpUnknown": Set Editor = New Edit_4_Unknown: Editor.SelectEdit
        Case "Interchange": Set Editor = New Edit_5_Interchange: Editor.SelectEdit
        Case "AST": Set Editor = New Edit_6_Ast: Editor.SelectEdit
        Case "CleanColor": Set Editor = New Edit_7_CleanColor: Editor.SelectEdit
        Case "Dynamic": Set Editor = New Edit_8_Dynamic: Editor.SelectEdit
        Case Else
            MsgBox "Erreur Trigger inconnus !", vbCritical, "Error"
    End Select
    
End Sub
