Attribute VB_Name = "Clear_1_Button"
Sub Clear_Session()

    'D�claration des Variables
    Dim response As VbMsgBoxResult
    
    reponse = MsgBox("Les donn�es pr�sentes dans les feuilles [Calccul Besoin] [ABC] [ABC Code mod�le] seront supprim�es. Voulez vous continuer ?", vbYesNo + vbQuestion, "Confirmation")
    
    If reponse = vbYes Then
        Set Clear = New Clear_2_Analysis
        Clear.ClearCalcul
        Clear.ClearABC
        Clear.ClearABCCM
        
        'Confirmation
        MsgBox "Les donn�es ont �t� supprim�es."
    End If
    
End Sub


