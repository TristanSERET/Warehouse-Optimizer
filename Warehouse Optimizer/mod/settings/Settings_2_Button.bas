Attribute VB_Name = "Settings_2_Button"
Sub Settings_Button()
    
    'D�claration des Variables
    Dim nameButton As String
    
    'S�lectionner le multipage pertinent en fonction du bouton
    nameButton = Application.Caller
    
    Select Case nameButton
        Case "set1": Settings_1_Main_Menu.MultiPage1.Value = 0
        Case "set2": Settings_1_Main_Menu.MultiPage1.Value = 2
    End Select
    
    'Executer le Formulaire
    Settings_1_Main_Menu.Show
    
End Sub
