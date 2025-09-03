Attribute VB_Name = "Calculs_1_Import"
Public Sub Forecast_Import()

    'D�claration des Variables
    Dim wsPV As Worksheet
    Dim wsCalcul As Worksheet
    Dim lastRowPV As Long
    Dim lastRowCalcul As Long
    Dim rowIndex As Long
    Dim typologie As String
    Dim importRange As Range
    
    'Initialisation des feuilles
    Set wsPV = ThisWorkbook.Sheets("D_PV")
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'D�finir les derni�res lignes
    lastRowPV = wsPV.Cells(wsPV.Rows.Count, "A").End(xlUp).Row
    lastRowCalcul = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row + 1
    
    'D�finir la typologie � importer
    typologie = GetSettings("Typologie")
    
    'Importer les pr�visions en fonction de la typologie
    For rowIndex = 2 To lastRowPV
        If Not IsError(wsPV.Cells(rowIndex, "D").Value) Then
            If wsPV.Cells(rowIndex, "D").Value = typologie Then
                Set importRange = wsPV.Range("A" & rowIndex & ":BI" & rowIndex)
                importRange.Copy
                    wsCalcul.Range("B" & lastRowCalcul).PasteSpecial Paste:=xlPasteValues
            End If
        End If
        'R�initialiser la variable de derni�re ligne pour la prochaine it�ration
        lastRowCalcul = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row + 1
    Next rowIndex
                
End Sub

Public Sub History_Import()

    'D�claration des Variables
    Dim wsCalcul As Worksheet
    Dim wsHV As Worksheet
    Dim lastRowCalcul As Long
    Dim lastRowHV As Long
    Dim rowIndex As Long
    Dim refToSearch As Variant
    Dim searchRange As Range
    Dim foundref As Range
    
    'Initilisation des feuilles
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    Set wsHV = ThisWorkbook.Sheets("D_HV")
    
    'D�finir les derni�res lignes
    lastRowCalcul = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row + 1
    lastRowHV = wsHV.Cells(wsHV.Rows.Count, "A").End(xlUp).Row + 1
    
    'D�finir la plage de recherche
    Set searchRange = wsHV.Range("A2:A" & lastRowHV)
    
    'Rechercher la valeur � importer selon la r�f�rence
    For rowIndex = 3 To lastRowCalcul
        refToSearch = wsCalcul.Cells(rowIndex, "B").Value
        Set foundref = searchRange.Find(What:=refToSearch, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not foundref Is Nothing Then
            wsCalcul.Cells(rowIndex, "BK").Value = wsHV.Cells(foundref.Row, "BE").Value
        Else
            wsCalcul.Cells(rowIndex, "BK").Value = ""
        End If
    Next rowIndex
        
End Sub
