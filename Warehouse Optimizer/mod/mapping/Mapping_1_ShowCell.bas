Attribute VB_Name = "Mapping_1_ShowCell"
Sub Show_Cellule()

    'D�claration des Variables
    Dim wsimplant As Worksheet
    Dim buttonName As String
    Dim targetColumn As Range
    
    'Initialisation de la feuille
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    
    'D�finir le bouton cliqu�
    buttonName = Application.Caller
    
    'D�terminer les colonnes correspondantes
    Select Case buttonName
        Case "Cellule_A": Set targetColumn = wsimplant.Columns("ER:FY")
        Case "Cellule_B": Set targetColumn = wsimplant.Columns("DI:EP")
        Case "Cellule_E": Set targetColumn = wsimplant.Columns("BZ:DG")
        Case "Cellule_F": Set targetColumn = wsimplant.Columns("AO:BX")
        Case "Cellule_G": Set targetColumn = wsimplant.Columns("B:AM")
    End Select
    
    'Inverser l'�tat actuel
    targetColumn.Hidden = Not targetColumn.Hidden
    
    'Message de confirmation
    If targetColumn.Hidden = True Then
        MsgBox "La " & buttonName & " a �t� masqu�e", vbInformation, "Confirmation"
    Else
        MsgBox "La " & buttonName & " a �t� d�masqu�e", vbInformation, "Confirmation"
    End If
    
End Sub
