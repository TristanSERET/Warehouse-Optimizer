Attribute VB_Name = "Calculs_2_Calculs"
Public Sub Check_NbSupport()

    'Déclration des Variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim ref As Variant
    
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'Définir la dernière ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
    
    'Prcourrir la colonne H et vérifier que la liste est numeric
    For rowIndex = 3 To lastRow
        If Not IsNumeric(wsCalcul.Cells(rowIndex, "H").Value) Then
            ref = wsCalcul.Cells(rowIndex, "B").Value
            MsgBox "L'execution de l'analyse est impossible, le nombre de supports par alévole n'est pas correctement défini pour certaines références. Référence identifié : " & ref, vbCritical, "Error"
            
            'Supprimer le tableau inccorect
            Set Clear = New Clear_2_Analysis
            Clear.ClearCalcul
            End
        End If
    Next rowIndex

End Sub

Public Sub Ecart_Calcul()

   'Déclaration des Variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    
    Dim actualTotal As Double
    Dim beforeTotal As Double
    
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'Définir la dernière ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
    
    'Calculer l'écart entre N et N-1
    For rowIndex = 3 To lastRow
        actualTotal = wsCalcul.Cells(rowIndex, "BJ").Value
        beforeTotal = wsCalcul.Cells(rowIndex, "BK").Value
        wsCalcul.Cells(rowIndex, "BL").Value = actualTotal - beforeTotal
    Next rowIndex

End Sub

Public Sub Solution_Max()

    'Déclaration des Variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim maxSales As Double
    Dim qtéSupport As Long
    Dim dispositionRate As Variant
    
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'Définir la dernière ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
            
    'Définir le % de mise à disposition
    dispositionRate = GetSettings("% Mise à disposition")
    
    'Parcourir les lignes
    For rowIndex = 3 To lastRow
    
        'Calculer la consommation maximale à l'année
        maxSales = Application.WorksheetFunction.Max(wsCalcul.Range("J" & rowIndex & ":BI" & rowIndex))
        wsCalcul.Cells(rowIndex, "BN").Value = maxSales
        
        'Calculer le Besoin Picking des pcs par rapport aux rate mise à dispo
        wsCalcul.Cells(rowIndex, "BO").Value = maxSales * dispositionRate
        
        'Calculer le nombre de supports nécessaire par référence
        qtéSupport = wsCalcul.Cells(rowIndex, "I").Value
        wsCalcul.Cells(rowIndex, "BP").Value = wsCalcul.Cells(rowIndex, "BO").Value / qtéSupport
        
        'Calculer le Besoin en Alvéole par référence
        wsCalcul.Cells(rowIndex, "BQ").Value = Application.WorksheetFunction.RoundUp(wsCalcul.Cells(rowIndex, "BP").Value / wsCalcul.Cells(rowIndex, "H").Value, 0)
        
    Next rowIndex
    
End Sub

Public Sub Solution_Average()
    
    'Déclaration des Variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim averageSales As Double
    Dim qtéSupport As Long
    Dim dispositionRate As Variant
        
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'Définir la dernière ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
    
    'Parcourir les lignes
    For rowIndex = 3 To lastRow
          
        'Calculer la consommation moyenne à l'année
        averageSales = Application.WorksheetFunction.Average(wsCalcul.Range("J" & rowIndex & ":BI" & rowIndex))
        wsCalcul.Cells(rowIndex, "BS").Value = averageSales
        
        'Définir le % de mise à disposition
        dispositionRate = GetSettings("% Mise à disposition")
        
        'Calculer le Besoin Picking des pcs par rapport aux rate mise à dispo
        wsCalcul.Cells(rowIndex, "BT").Value = averageSales * dispositionRate
        
        'Calculer le nombre de supports nécessaire par référence
        qtéSupport = wsCalcul.Cells(rowIndex, "I").Value
        wsCalcul.Cells(rowIndex, "BU").Value = wsCalcul.Cells(rowIndex, "BS").Value / qtéSupport
        
        'Calculer le Besoin en Alvéole par référence
        wsCalcul.Cells(rowIndex, "BV").Value = Application.WorksheetFunction.RoundUp(wsCalcul.Cells(rowIndex, "BU").Value / wsCalcul.Cells(rowIndex, "H").Value, 0)
        
    Next rowIndex

End Sub

Public Sub Solution_BestAverage()

    'Déclaration des variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim limitsAverage As Byte
    Dim dispositionRate As Variant
    Dim bestWeek As Byte
    Dim sum As Double
    Dim averageSales As Double
    Dim qtéSupport As Long
    
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'Définir la dernière ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
    
    'Définir les paramètres du nombre de meilleure semaine à calculer et mise à disposition
    limitsAverage = GetSettings("Limite de semaine (Meilleure Moyenne)")
    dispositionRate = GetSettings("% Mise à disposition")
    
    'Parcourir les lignes
    For rowIndex = 3 To lastRow
    
        'Calculer la moyenne des meilleures semaines
        For bestWeek = 1 To limitsAverage
            sum = sum + Application.WorksheetFunction.Large(wsCalcul.Range(wsCalcul.Cells(rowIndex, "J"), wsCalcul.Cells(rowIndex, "BI")), bestWeek)
        Next bestWeek
        
        averageSales = sum / limitsAverage
        wsCalcul.Cells(rowIndex, "BX").Value = averageSales
        
        'Calculer le Besoin Picking des pcs par rapport aux rate mise à dispo
        wsCalcul.Cells(rowIndex, "BY").Value = averageSales * dispositionRate
        
        'Calculer le nombre de supports nécessaire par référence
        qtéSupport = wsCalcul.Cells(rowIndex, "I").Value
        wsCalcul.Cells(rowIndex, "BZ").Value = wsCalcul.Cells(rowIndex, "BS").Value / qtéSupport
        
        'Calculer le Besoin en Alvéole par référence
        wsCalcul.Cells(rowIndex, "CA").Value = Application.WorksheetFunction.RoundUp(wsCalcul.Cells(rowIndex, "BZ").Value / wsCalcul.Cells(rowIndex, "H").Value, 0)
        
        'Réinitialiser la variable pour la prochaine itération
        sum = 0
        
    Next rowIndex

End Sub

Public Sub Solution_Epiphenomene()

    'Déclaration des Variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim sensitivity As Variant
    Dim rowIndex As Long
    Dim selectRange As Range
    Dim cell As Range
    Dim selectMax As Byte
    Dim epiphénomène As Double
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'Définir la dernière ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
    
    'Définir la sensibilité du calcul
    sensitivity = GetSettings("Sensibilité des épiphénomènes")
    
    'Parcourrir les lignes et neutraliser les épiphénomènes
    If Not sensitivity = 0 Then
        For rowIndex = 3 To lastRow
            Set selectRange = wsCalcul.Range(wsCalcul.Cells(rowIndex, "J"), wsCalcul.Cells(rowIndex, "BI"))
            For selectMax = 1 To sensitivity
                epiphénomène = Application.WorksheetFunction.Max(selectRange)
                For Each cell In selectRange
                    If cell.Value = epiphénomène Then
                        cell.ClearContents
                        Exit For
                    End If
                Next cell
            Next selectMax
        Next rowIndex
    End If

End Sub
