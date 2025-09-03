VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Clear_2_Analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe d'effacement du cache partie Analyse

'Implémentation de la classe mère
Implements Clear_1_Main

Private Sub Clear_1_Main_ClearCalcul()

    'Déclaration des Variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim rangeToClear As Range
    
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'Définir la dernière ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
    
    'Initialisation de la plage à supprimer
    Set rangeToClear = wsCalcul.Range("B3:CA" & lastRow)
    
    'Clear le Cache
    If lastRow > 2 Then
        rangeToClear.ClearContents
    End If
    
End Sub

Private Sub Clear_1_Main_ClearABC()

    'Déclaration des Variables
    Dim wsABC As Worksheet
    Dim lastRow As Long
    Dim rangeToClear As Range
    
    'Initialisation de la feuille
    Set wsABC = ThisWorkbook.Sheets("ABC")
    
    'Définir la dernière ligne
    lastRow = wsABC.Cells(wsABC.Rows.Count, "D").End(xlUp).Row
    
    'Initialisation de la plage à supprimer
    Set rangeToClear = wsABC.Range("B3:J" & lastRow)
    
    'Clear le cache
    If lastRow > 2 Then
        rangeToClear.ClearContents
        rangeToClear.Interior.ColorIndex = xlColorIndexNone
    End If
    
End Sub

Private Sub Clear_1_Main_ClearABCCM()

    'Déclaration des Variables
    Dim wsABCCM As Worksheet
    Dim lastRow As Integer
    Dim tcd As PivotTable
    Dim rangeToDelete As Range
    
    'Initialisation de la feuille
    Set wsABCCM = ThisWorkbook.Sheets("ABC Code Modèle")
    
    'Définir la dernière ligne
    lastRow = wsABCCM.Cells(wsABCCM.Rows.Count, "B").End(xlUp).Row
    
    'Initialisation de la plage à supprimer
    Set rangeToDelete = wsABCCM.Range("B1:I" & lastRow)
    
    'Initialisation du tcd
    Set tcd = wsABCCM.PivotTables("TCD_CM")
    
    'Supprimer le TCD
    tcd.TableRange2.Delete
    Set tcd = Nothing
    
    'Supprimer les couleurs et les données de la plage
    With rangeToDelete
        .ClearContents
        .Interior.ColorIndex = xlNone
    End With
    
End Sub
