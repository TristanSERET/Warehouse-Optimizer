Attribute VB_Name = "Settings_1_Selector"
Public Function GetSettings(nameSettings As String) As Variant

    'Déclaration des Variables
    Dim wsSettings As Worksheet
    Dim rangeSettings As Range
    
    'Initialisation de la feuille
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    
    'Initialisation de la plage
    Set rangeSettings = wsSettings.Columns(2).Find(What:=nameSettings, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Recherche du pramaètre en colonne A
    If Not rangeSettings Is Nothing Then
        GetSettings = rangeSettings.Offset(0, 1).Value
    Else
        GetSettings = CVErr(xlErrNA)
        MsgBox "Le paramètre " & nameSettings & " est introuvable !", vbCritical, "Error"
    End If

End Function
