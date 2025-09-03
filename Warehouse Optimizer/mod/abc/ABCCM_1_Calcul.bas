Attribute VB_Name = "ABCCM_1_Calcul"
Public Sub Calcul_ABCCM()

    'Déclaration des Variables
    Dim wsABCCM As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim cumulateRate As Double
    Dim sensiClassA As Double
    Dim sensiClassB As Double
    Dim sensiClassC As Double
    
    'Initialisation de la feuille
    Set wsABCCM = ThisWorkbook.Sheets("ABC Code Modèle")
    
    'Définir les parmamètres de sensibilité des % de classe
    sensiClassA = GetSettings("Sensibilité de la Classe A")
    sensiClassB = GetSettings("Sensibilité de la Classe B")
    sensiClassC = GetSettings("Sensibilité de la Classe C")
    
    'Définir la dernière ligne
    lastRow = wsABCCM.Cells(wsABCCM.Rows.Count, "B").End(xlUp).Row
    
    'Parcourir l'ABC et calculer les éléments
    For rowIndex = 4 To lastRow - 1
        If InStr(1, wsABCCM.Cells(rowIndex, "B").Value, "Total", vbTextCompare) Then
        
            'Calculer le % des alvéoles
            If wsABCCM.Cells(lastRow, "F").Value <> 0 Then
                wsABCCM.Cells(rowIndex, "G").Value = wsABCCM.Cells(rowIndex, "F").Value / wsABCCM.Cells(lastRow, "F").Value
                wsABCCM.Cells(rowIndex, "G").Style = "Percent"
            Else
                wsABCCM.Cells(rowIndex, "G").Value = 0
            End If
            
            'Calculer le % cumulé
            cumulateRate = cumulateRate + wsABCCM.Cells(rowIndex, "G").Value
            wsABCCM.Cells(rowIndex, "H").Value = cumulateRate
            wsABCCM.Cells(rowIndex, "H").Style = "Percent"
            
            'Affecter la classe
            If cumulateRate <= sensiClassA Then
                wsABCCM.Cells(rowIndex, "I").Value = "A"
            ElseIf cumulateRate <= sensiClassB Then
                wsABCCM.Cells(rowIndex, "I").Value = "B"
            ElseIf cumulateRate <= sensiClassC Then
                wsABCCM.Cells(rowIndex, "I").Value = "C"
            End If
        End If
    Next rowIndex
    
    'Affecter une couleur
    Color_Assignement

End Sub

Private Sub Color_Assignement()

    'Déclaration des Variables
    Dim wsABCCM As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim Class As String
    
    'Initialisation de la feuille
    Set wsABCCM = ThisWorkbook.Sheets("ABC Code Modèle")
    
    'Définir la dernière ligne
    lastRow = wsABCCM.Cells(wsABCCM.Rows.Count, "B").End(xlUp).Row
    
    'Parcourir les lignes et les colorier en fonction de la classe
    For rowIndex = 3 To lastRow - 1
    
        'Définir la class actuelle
        Class = wsABCCM.Cells(rowIndex, "I").Value
        
        Select Case Class
            Case "A": wsABCCM.Range("B" & rowIndex & ":I" & rowIndex).Interior.Color = RGB(198, 224, 180)
            Case "B": wsABCCM.Range("B" & rowIndex & ":I" & rowIndex).Interior.Color = RGB(248, 203, 173)
            Case "C": wsABCCM.Range("B" & rowIndex & ":I" & rowIndex).Interior.Color = RGB(174, 170, 170)
        End Select
    Next rowIndex

End Sub
