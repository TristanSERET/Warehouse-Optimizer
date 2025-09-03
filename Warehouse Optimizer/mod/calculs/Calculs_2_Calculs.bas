Attribute VB_Name = "Calculs_2_Calculs"
Public Sub Check_NbSupport()

    'D�clration des Variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim ref As Variant
    
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'D�finir la derni�re ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
    
    'Prcourrir la colonne H et v�rifier que la liste est numeric
    For rowIndex = 3 To lastRow
        If Not IsNumeric(wsCalcul.Cells(rowIndex, "H").Value) Then
            ref = wsCalcul.Cells(rowIndex, "B").Value
            MsgBox "L'execution de l'analyse est impossible, le nombre de supports par al�vole n'est pas correctement d�fini pour certaines r�f�rences. R�f�rence identifi� : " & ref, vbCritical, "Error"
            
            'Supprimer le tableau inccorect
            Set Clear = New Clear_2_Analysis
            Clear.ClearCalcul
            End
        End If
    Next rowIndex

End Sub

Public Sub Ecart_Calcul()

   'D�claration des Variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    
    Dim actualTotal As Double
    Dim beforeTotal As Double
    
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'D�finir la derni�re ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
    
    'Calculer l'�cart entre N et N-1
    For rowIndex = 3 To lastRow
        actualTotal = wsCalcul.Cells(rowIndex, "BJ").Value
        beforeTotal = wsCalcul.Cells(rowIndex, "BK").Value
        wsCalcul.Cells(rowIndex, "BL").Value = actualTotal - beforeTotal
    Next rowIndex

End Sub

Public Sub Solution_Max()

    'D�claration des Variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim maxSales As Double
    Dim qt�Support As Long
    Dim dispositionRate As Variant
    
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'D�finir la derni�re ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
            
    'D�finir le % de mise � disposition
    dispositionRate = GetSettings("% Mise � disposition")
    
    'Parcourir les lignes
    For rowIndex = 3 To lastRow
    
        'Calculer la consommation maximale � l'ann�e
        maxSales = Application.WorksheetFunction.Max(wsCalcul.Range("J" & rowIndex & ":BI" & rowIndex))
        wsCalcul.Cells(rowIndex, "BN").Value = maxSales
        
        'Calculer le Besoin Picking des pcs par rapport aux rate mise � dispo
        wsCalcul.Cells(rowIndex, "BO").Value = maxSales * dispositionRate
        
        'Calculer le nombre de supports n�cessaire par r�f�rence
        qt�Support = wsCalcul.Cells(rowIndex, "I").Value
        wsCalcul.Cells(rowIndex, "BP").Value = wsCalcul.Cells(rowIndex, "BO").Value / qt�Support
        
        'Calculer le Besoin en Alv�ole par r�f�rence
        wsCalcul.Cells(rowIndex, "BQ").Value = Application.WorksheetFunction.RoundUp(wsCalcul.Cells(rowIndex, "BP").Value / wsCalcul.Cells(rowIndex, "H").Value, 0)
        
    Next rowIndex
    
End Sub

Public Sub Solution_Average()
    
    'D�claration des Variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim averageSales As Double
    Dim qt�Support As Long
    Dim dispositionRate As Variant
        
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'D�finir la derni�re ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
    
    'Parcourir les lignes
    For rowIndex = 3 To lastRow
          
        'Calculer la consommation moyenne � l'ann�e
        averageSales = Application.WorksheetFunction.Average(wsCalcul.Range("J" & rowIndex & ":BI" & rowIndex))
        wsCalcul.Cells(rowIndex, "BS").Value = averageSales
        
        'D�finir le % de mise � disposition
        dispositionRate = GetSettings("% Mise � disposition")
        
        'Calculer le Besoin Picking des pcs par rapport aux rate mise � dispo
        wsCalcul.Cells(rowIndex, "BT").Value = averageSales * dispositionRate
        
        'Calculer le nombre de supports n�cessaire par r�f�rence
        qt�Support = wsCalcul.Cells(rowIndex, "I").Value
        wsCalcul.Cells(rowIndex, "BU").Value = wsCalcul.Cells(rowIndex, "BS").Value / qt�Support
        
        'Calculer le Besoin en Alv�ole par r�f�rence
        wsCalcul.Cells(rowIndex, "BV").Value = Application.WorksheetFunction.RoundUp(wsCalcul.Cells(rowIndex, "BU").Value / wsCalcul.Cells(rowIndex, "H").Value, 0)
        
    Next rowIndex

End Sub

Public Sub Solution_BestAverage()

    'D�claration des variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim limitsAverage As Byte
    Dim dispositionRate As Variant
    Dim bestWeek As Byte
    Dim sum As Double
    Dim averageSales As Double
    Dim qt�Support As Long
    
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'D�finir la derni�re ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
    
    'D�finir les param�tres du nombre de meilleure semaine � calculer et mise � disposition
    limitsAverage = GetSettings("Limite de semaine (Meilleure Moyenne)")
    dispositionRate = GetSettings("% Mise � disposition")
    
    'Parcourir les lignes
    For rowIndex = 3 To lastRow
    
        'Calculer la moyenne des meilleures semaines
        For bestWeek = 1 To limitsAverage
            sum = sum + Application.WorksheetFunction.Large(wsCalcul.Range(wsCalcul.Cells(rowIndex, "J"), wsCalcul.Cells(rowIndex, "BI")), bestWeek)
        Next bestWeek
        
        averageSales = sum / limitsAverage
        wsCalcul.Cells(rowIndex, "BX").Value = averageSales
        
        'Calculer le Besoin Picking des pcs par rapport aux rate mise � dispo
        wsCalcul.Cells(rowIndex, "BY").Value = averageSales * dispositionRate
        
        'Calculer le nombre de supports n�cessaire par r�f�rence
        qt�Support = wsCalcul.Cells(rowIndex, "I").Value
        wsCalcul.Cells(rowIndex, "BZ").Value = wsCalcul.Cells(rowIndex, "BS").Value / qt�Support
        
        'Calculer le Besoin en Alv�ole par r�f�rence
        wsCalcul.Cells(rowIndex, "CA").Value = Application.WorksheetFunction.RoundUp(wsCalcul.Cells(rowIndex, "BZ").Value / wsCalcul.Cells(rowIndex, "H").Value, 0)
        
        'R�initialiser la variable pour la prochaine it�ration
        sum = 0
        
    Next rowIndex

End Sub

Public Sub Solution_Epiphenomene()

    'D�claration des Variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim sensitivity As Variant
    Dim rowIndex As Long
    Dim selectRange As Range
    Dim cell As Range
    Dim selectMax As Byte
    Dim epiph�nom�ne As Double
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'D�finir la derni�re ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
    
    'D�finir la sensibilit� du calcul
    sensitivity = GetSettings("Sensibilit� des �piph�nom�nes")
    
    'Parcourrir les lignes et neutraliser les �piph�nom�nes
    If Not sensitivity = 0 Then
        For rowIndex = 3 To lastRow
            Set selectRange = wsCalcul.Range(wsCalcul.Cells(rowIndex, "J"), wsCalcul.Cells(rowIndex, "BI"))
            For selectMax = 1 To sensitivity
                epiph�nom�ne = Application.WorksheetFunction.Max(selectRange)
                For Each cell In selectRange
                    If cell.Value = epiph�nom�ne Then
                        cell.ClearContents
                        Exit For
                    End If
                Next cell
            Next selectMax
        Next rowIndex
    End If

End Sub
