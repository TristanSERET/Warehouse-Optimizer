VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ABC_2_Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe ABC pour les Ventes

'Implémentation de la classe mère
Implements ABC_1_Main

Private Sub ABC_1_Main_ImportABC()

    'Déclaration des Variables
    Dim wsABC As Worksheet
    Dim wsCalcul As Worksheet
    Dim lastRowCalcul As Long
    Dim lastRowABC As Long
    Dim rowIndex As Long
    
    Dim setCalcul As Variant
    
    Dim ref As Variant
    Dim codeMod As Variant
    Dim description As String
    Dim sales As Integer
    Dim needSupport As Double
    Dim needAlveole As Byte
    
    'Initialisation des feuilles
    Set wsABC = ThisWorkbook.Sheets("ABC")
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'Définir la dernière ligne de la feuille calcul
    lastRowCalcul = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
    
    'Définir le settings
    setCalcul = GetSettings("Calcul retenu en sortie")
    
    'Parcourir la feuille calcul et Importer les données dans l'ABC
    For rowIndex = 3 To lastRowCalcul
        
        'Affecter les variables
        ref = wsCalcul.Cells(rowIndex, "B").Value
        codeMod = wsCalcul.Cells(rowIndex, "D").Value
        description = wsCalcul.Cells(rowIndex, "C").Value
        sales = wsCalcul.Cells(rowIndex, "BJ").Value
        
        Select Case setCalcul
            Case "Meilleure Moyenne"
                needSupport = wsCalcul.Cells(rowIndex, "BZ").Value
                needAlveole = wsCalcul.Cells(rowIndex, "CA").Value
            Case "Max"
                needSupport = wsCalcul.Cells(rowIndex, "BP").Value
                needAlveole = wsCalcul.Cells(rowIndex, "BQ").Value
            Case "Moyenne"
                needSupport = wsCalcul.Cells(rowIndex, "BU").Value
                needAlveole = wsCalcul.Cells(rowIndex, "BV").Value
        End Select
            
        'Importer les variables dans l'ABC
        lastRowABC = wsABC.Cells(wsABC.Rows.Count, "B").End(xlUp).Row + 1
        
        wsABC.Cells(lastRowABC, "B").Value = ref
        wsABC.Cells(lastRowABC, "C").Value = codeMod
        wsABC.Cells(lastRowABC, "D").Value = description
        wsABC.Cells(lastRowABC, "E").Value = sales
        wsABC.Cells(lastRowABC, "H").Value = needSupport
        wsABC.Cells(lastRowABC, "I").Value = needAlveole
    Next rowIndex

End Sub

Private Sub ABC_1_Main_CalculABC()

    'Faire les Calculs ABC
    Call ABC_Calculs

End Sub

