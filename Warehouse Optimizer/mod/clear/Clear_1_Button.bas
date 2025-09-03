Attribute VB_Name = "Clear_1_Button"
Sub Clear_Session()

    'Déclaration des Variables
    Dim response As VbMsgBoxResult
    
    reponse = MsgBox("Les données présentes dans les feuilles [Calccul Besoin] [ABC] [ABC Code modèle] seront supprimées. Voulez vous continuer ?", vbYesNo + vbQuestion, "Confirmation")
    
    If reponse = vbYes Then
        Set Clear = New Clear_2_Analysis
        Clear.ClearCalcul
        Clear.ClearABC
        Clear.ClearABCCM
        
        'Confirmation
        MsgBox "Les données ont été supprimées."
    End If
    
End Sub


