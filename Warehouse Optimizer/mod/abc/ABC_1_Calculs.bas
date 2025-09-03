Attribute VB_Name = "ABC_1_Calculs"
Public Sub ABC_Calculs()

    'Déclaration des Variables
    Dim wsABC As Worksheet
    Dim wsSettings As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    
    Dim sumOfSales As Double
    Dim cumulateRate As Variant
    
    Dim sensiClass_A As Double
    Dim sensiClass_B As Double
    Dim sensiClass_C As Double
    
    'Initialisation de la feuille
    Set wsABC = ThisWorkbook.Sheets("ABC")
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    
    'Définir la dernière ligne
    lastRow = wsABC.Cells(wsABC.Rows.Count, "B").End(xlUp).Row
    
    'Définir les paramètres
    sensiClass_A = GetSettings("Sensibilité de la Classe A")
    sensiClass_B = GetSettings("Sensibilité de la Classe B")
    sensiClass_C = GetSettings("Sensibilité de la Classe C")
    
    'Calculer la somme des ventes
    sumOfSales = Application.WorksheetFunction.sum(wsABC.Range("E3:E" & lastRow))
    wsABC.Cells(lastRow + 2, "D").Value = "Total"
    wsABC.Cells(lastRow + 2, "E").Value = sumOfSales
    
    'Activer les filtres et trier du plus grand au plus petit
    wsABC.Range("B2:J" & lastRow).Sort Key1:=wsABC.Range("E1"), Order1:=xlDescending, Header:=xlYes
        
    'Parcourrir les lignes
    For rowIndex = 3 To lastRow
    
        'Calculer les % des Ventes
        wsABC.Cells(rowIndex, "F").Value = wsABC.Cells(rowIndex, "E").Value / sumOfSales
        wsABC.Cells(rowIndex, "F").Style = "Percent"
        
        'Calculer le % cumulé des ventes
        If rowIndex <> 3 Then
            wsABC.Cells(rowIndex, "G").Value = wsABC.Cells(rowIndex, "F").Value + wsABC.Cells(rowIndex - 1, "G").Value
            wsABC.Cells(rowIndex, "G").Style = "Percent"
        Else
            wsABC.Cells(rowIndex, "G").Value = wsABC.Cells(rowIndex, "F").Value
            wsABC.Cells(rowIndex, "G").Style = "Percent"
        End If
        
        'Définir la classe
        cumulateRate = wsABC.Cells(rowIndex, "G").Value
        
        If cumulateRate <= sensiClass_A Then
            wsABC.Cells(rowIndex, "J").Value = "A"
            
        ElseIf cumulateRate <= sensiClass_B Then
            wsABC.Cells(rowIndex, "J").Value = "B"
            
        ElseIf cumulateRate <= sensiClass_C Then
            wsABC.Cells(rowIndex, "J").Value = "C"
        Else
            wsABC.Cells(rowIndex, "J").Value = "C"
        End If
        
    Next rowIndex
    
    'Affecter les couleurs en fonction des classes
    Color_Assignement

End Sub

Private Sub Color_Assignement()

    'Déclaration des Variables
    Dim wsABC As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim Class As String
    
    'Initialisation de la feuille
    Set wsABC = ThisWorkbook.Sheets("ABC")
    
    'Définir la dernière ligne
    lastRow = wsABC.Cells(wsABC.Rows.Count, "B").End(xlUp).Row
    
    'Parcourir les lignes et les colorier en fonction de la classe
    For rowIndex = 3 To lastRow
    
        'Définir la class actuelle
        Class = wsABC.Cells(rowIndex, "J").Value
        
        Select Case Class
            Case "A": wsABC.Range("B" & rowIndex & ":J" & rowIndex).Interior.Color = RGB(198, 224, 180)
            Case "B": wsABC.Range("B" & rowIndex & ":J" & rowIndex).Interior.Color = RGB(248, 203, 173)
            Case "C": wsABC.Range("B" & rowIndex & ":J" & rowIndex).Interior.Color = RGB(174, 170, 170)
        End Select
    Next rowIndex

End Sub
