VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ABCCM_3_Weight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe d'ABC au Code Modèle par rapport au poids

'Implémentation de la classe mère
Implements ABCCM_1_Main

Public Sub ABCCM_1_Main_AddTCD()
 
    'Déclaration des Variables
    Dim wsABC As Worksheet
    Dim wsABCCM As Worksheet
    Dim lastRow As Long
    Dim destinationCell As Range
    Dim tcdSource As Range
    Dim tcdCache As PivotCache
    Dim tcd As PivotTable
    Dim st As PivotField
    Dim settings As Variant
    Dim preference As Variant
    
    'Initialisation de la feuille
    Set wsABC = ThisWorkbook.Sheets("ABC")
    Set wsABCCM = ThisWorkbook.Sheets("ABC Code Modèle")
    
    'Définir la dernière ligne
    lastRow = wsABC.Cells(wsABC.Rows.Count, "B").End(xlUp).Row
    
    'Définir les paramètres de préférence du trie
    settings = GetSettings("Préférence du trie ABC au Code Modèle")
    
    Select Case settings
        Case "Somme des Alvéoles": preference = "Somme de Besoin Alvéole"
        Case "Somme des Poids": preference = "Somme de Poids"
    End Select
    
    'Définir la plage de destination
    Set destinationCell = wsABCCM.Range("B2")
    
    'Définir la plage source du TCD
    Set tcdSource = wsABC.Range("B2:J" & lastRow)
    
    'Initialisation du cache du TCD
    Set tcdCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=tcdSource)
    
    'Initialisation du TCD
    Set tcd = tcdCache.CreatePivotTable(TableDestination:=destinationCell, TableName:="TCD_CM")
    
    'Ajouter les champs en ligne
    With tcd
        .PivotFields("Code Modele").Orientation = xlRowField
        .PivotFields("Référence").Orientation = xlRowField
        .PivotFields("Designation").Orientation = xlRowField
        
        'Ajouter les champs en valeur
        .PivotFields("Poids").Orientation = xlDataField
        .PivotFields("Besoin Alvéole").Orientation = xlDataField
        
        'Trier par "Somme des ventes" décroissant sur le premier champ ("Code Modèle")
        .PivotFields("Code Modele").AutoSort Order:=xlDescending, Field:=preference
        
        'Gestion du format
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        
        For Each st In .RowFields
            st.Subtotals(1) = False
            Exit For
        Next st
    End With

End Sub

Private Sub ABCCM_1_MAIN_CalculABCCM()

Call Calcul_ABCCM

End Sub
