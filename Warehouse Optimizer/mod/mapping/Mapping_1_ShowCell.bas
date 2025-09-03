Attribute VB_Name = "Mapping_1_ShowCell"
Sub Show_Cellule()

    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim buttonName As String
    Dim targetColumn As Range
    
    'Initialisation de la feuille
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    
    'Définir le bouton cliqué
    buttonName = Application.Caller
    
    'Déterminer les colonnes correspondantes
    Select Case buttonName
        Case "Cellule_A": Set targetColumn = wsimplant.Columns("ER:FY")
        Case "Cellule_B": Set targetColumn = wsimplant.Columns("DI:EP")
        Case "Cellule_E": Set targetColumn = wsimplant.Columns("BZ:DG")
        Case "Cellule_F": Set targetColumn = wsimplant.Columns("AO:BX")
        Case "Cellule_G": Set targetColumn = wsimplant.Columns("B:AM")
    End Select
    
    'Inverser l'état actuel
    targetColumn.Hidden = Not targetColumn.Hidden
    
    'Message de confirmation
    If targetColumn.Hidden = True Then
        MsgBox "La " & buttonName & " a été masquée", vbInformation, "Confirmation"
    Else
        MsgBox "La " & buttonName & " a été démasquée", vbInformation, "Confirmation"
    End If
    
End Sub
