Attribute VB_Name = "Export_1_Button"
Sub Export_Button()

    Set Export = New ExportIT_2_Global
    
    Export.clearIT
    Export.fixeArea
    Export.dynamicArea
    
    'Message de Confirmation
    MsgBox "Les donn�es ont bien �t� export�es vers la feuille IT", vbInformation, "Succ�s"

End Sub
