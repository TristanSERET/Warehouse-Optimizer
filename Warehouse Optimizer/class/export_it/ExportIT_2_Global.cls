VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportIT_2_Global"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe d'export des données vers IT

'Implémentation de la classe mère
Implements ExportIT_1_Main

Private Sub ExportIT_1_Main_fixeArea()

    'Déclration des Variables
    Dim celluleArray(1 To 5) As String
    Dim celluleSelected As Variant
    
    Dim startColumn As Integer
    Dim endColumn As Integer
    
    Dim startRow As Integer
    Dim endRow As Integer
    
    'Initialisation du Tableau des Cellules
    celluleArray(1) = "Cellule_A"
    celluleArray(2) = "Cellule_B"
    celluleArray(3) = "Cellule_E"
    celluleArray(4) = "Cellule_F"
    celluleArray(5) = "Cellule_G"
    
    'Définir les coordonnées à traiter en fonction de la cellule en cours
    For Each celluleSelected In celluleArray
        Call Find_Column(celluleSelected, startColumn, endColumn)
        Call Find_Row(celluleSelected, startRow, endRow)
        
        Select Case celluleSelected
            Case "Cellule_A": Call Cellule_A(startColumn, endColumn, startRow, endRow)
            Case "Cellule_B": Call Cellule_B(startColumn, endColumn, startRow, endRow)
            Case "Cellule_E": Call Cellule_E(startColumn, endColumn, startRow, endRow)
            Case "Cellule_F": Call Cellule_F(startColumn, endColumn, startRow, endRow)
            Case "Cellule_G": Call Cellule_G(startColumn, endColumn, startRow, endRow)
        End Select
    Next celluleSelected
    
    'Affecter les Gen-Codes en fonction des emplacements
    Call Affect_Gencode
        
End Sub

Private Sub Find_Column(celluleSelected As Variant, ByRef startColumn As Integer, ByRef endColumn As Integer)
    
    'Définir les colonnes de références
    Select Case celluleSelected
        Case "Cellule_A": startColumn = 149: endColumn = 180
        Case "Cellule_B": startColumn = 114: endColumn = 145
        Case "Cellule_E": startColumn = 79: endColumn = 110
        Case "Cellule_F": startColumn = 43: endColumn = 74
        Case "Cellule_G": startColumn = 5: endColumn = 36
    End Select
    
End Sub

Private Sub Find_Row(celluleSelected As Variant, ByRef startRow As Integer, ByRef endRow As Integer)
    
    'En Fonction de la cellule appliquer des valeurs différentes
    Select Case celluleSelected
        Case "Cellule_A": startRow = 90: endRow = 3
        Case "Cellule_B": startRow = 98: endRow = 3
        Case "Cellule_E": startRow = 90: endRow = 3
        Case "Cellule_F": startRow = 98: endRow = 3
        Case "Cellule_G": startRow = 90: endRow = 3
    End Select
        
End Sub

Private Sub Cellule_A(startColumn As Integer, endColumn As Integer, startRow As Integer, endRow As Integer)

    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim wsIT As Worksheet
    
    Dim lastRowIT As Long
    
    Dim colIndex As Integer
    Dim rowIndex As Integer
    Dim reference As String
    
    Dim iRow As String
    Dim iCol As String
    Dim empTag As String
    
    'Initialisation des feuilles
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    Set wsIT = ThisWorkbook.Sheets("IT")

    'Définir la dernière ligne de la feuille IT
    lastRowIT = wsIT.Cells(wsIT.Rows.Count, "A").End(xlUp).Row
    
    'Parcourrir la cellule et prendre les coordonnées
    For colIndex = startColumn To endColumn
        For rowIndex = startRow To endRow Step -1
            reference = wsimplant.Cells(rowIndex, colIndex).Value
            iRow = wsimplant.Cells(rowIndex, "EF").Value
            iCol = wsimplant.Cells(91, colIndex).Value
            
            If wsimplant.Cells(rowIndex, colIndex).Interior.Color <> RGB(217, 217, 217) Then
                If iCol <> "" Then
                    empTag = "AP2 R" & iCol & " C" & iRow & " T1"
                    
                    'Saisir les données dans la feuille IT
                    wsIT.Cells(lastRowIT + 1, "A").Value = empTag
                    wsIT.Cells(lastRowIT + 1, "B").Value = reference
                    
                    'Réinitialiser la variable lastRow pour la prochaine itération
                    lastRowIT = wsIT.Cells(wsIT.Rows.Count, "A").End(xlUp).Row
                End If
            End If
        Next rowIndex
    Next colIndex
    
End Sub

Private Sub Cellule_B(startColumn As Integer, endColumn As Integer, startRow As Integer, endRow As Integer)
    
    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim wsIT As Worksheet
    
    Dim lastRowIT As Long
    
    Dim colIndex As Integer
    Dim rowIndex As Integer
    Dim reference As String
    

    'Initialisation des feuilles
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    Set wsIT = ThisWorkbook.Sheets("IT")

    'Définir la dernière ligne de la feuille IT
    lastRowIT = wsIT.Cells(wsIT.Rows.Count, "A").End(xlUp).Row
    
    For colIndex = startColumn To endColumn
        For rowIndex = startRow To endRow Step -1
            reference = wsimplant.Cells(rowIndex, colIndex).Value
            iRow = wsimplant.Cells(rowIndex, "DI").Value
            iCol = wsimplant.Cells(99, colIndex).Value
            
            If wsimplant.Cells(rowIndex, colIndex).Interior.Color <> RGB(217, 217, 217) Then
                If iCol <> "" Then
                    empTag = "AP2 R" & iCol & " C" & iRow & " T1"
                    
                    'Saisir les données dans la feuille IT
                    wsIT.Cells(lastRowIT + 1, "A").Value = empTag
                    wsIT.Cells(lastRowIT + 1, "B").Value = reference
                    
                    'Réinitialiser la variable lastRow pour la prochaine itération
                    lastRowIT = wsIT.Cells(wsIT.Rows.Count, "A").End(xlUp).Row
                End If
            End If
        Next rowIndex
    Next colIndex

End Sub


Private Sub Cellule_E(startColumn As Integer, endColumn As Integer, startRow As Integer, endRow As Integer)

    'Déclration des Variables
    Dim wsimplant As Worksheet
    Dim wsIT As Worksheet
    
    Dim lastRowIT As Long
    
    Dim colIndex As Integer
    Dim rowIndex As Integer
    Dim reference As Variant
    
    Dim iRow As String
    Dim iCol As String
    Dim empTag As String
    
    'Initialisation des feuilles
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    Set wsIT = ThisWorkbook.Sheets("IT")
    
    'Définir la dernière ligne de la feuille IT
    lastRowIT = wsIT.Cells(wsIT.Rows.Count, "A").End(xlUp).Row
    
    'Parcourir la cellule et prendre les coordonnées
    For colIndex = startColumn To endColumn
        For rowIndex = startRow To endRow Step -1
            reference = wsimplant.Cells(rowIndex, colIndex).Value
            iRow = wsimplant.Cells(rowIndex, "BZ").Value
            iCol = wsimplant.Cells(91, colIndex).Value
            
            If wsimplant.Cells(rowIndex, colIndex).Interior.Color <> RGB(217, 217, 217) Then
                If iCol <> "" Then
                    empTag = "FP5 R" & iCol & " C" & iRow & " T1"
                    
                    'Saisir les données dans la feuille IT
                    wsIT.Cells(lastRowIT + 1, "A").Value = empTag
                    wsIT.Cells(lastRowIT + 1, "B").Value = reference
                    
                    'Réinitialiser la variable lastRow pour la prochaine itération
                    lastRowIT = wsIT.Cells(wsIT.Rows.Count, "A").End(xlUp).Row
                End If
            End If
        Next rowIndex
    Next colIndex

End Sub

Private Sub Cellule_F(startColumn As Integer, endColumn As Integer, startRow As Integer, endRow As Integer)

    'Déclration des Variables
    Dim wsimplant As Worksheet
    Dim wsIT As Worksheet
    
    Dim lastRowIT As Long
    
    Dim colIndex As Integer
    Dim rowIndex As Integer
    Dim reference As Variant
    
    Dim iRow_1 As String
    Dim iRow_2 As String
    Dim iCol As String
    Dim empTag As String
    
    'Initialisation des feuilles
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    Set wsIT = ThisWorkbook.Sheets("IT")
    
    'Définir la dernière ligne de la feuille IT
    lastRowIT = wsIT.Cells(wsIT.Rows.Count, "A").End(xlUp).Row
    
    'Parcourir la cellule et prendre les coordonnées
    For colIndex = startColumn To endColumn
        For rowIndex = startRow To endRow Step -1
            reference = wsimplant.Cells(rowIndex, colIndex).Value
            iRow_1 = wsimplant.Cells(rowIndex, "BX").Value
            iRow_2 = wsimplant.Cells(rowIndex, "BW").Value
            iCol = wsimplant.Cells(99, colIndex).Value

            If wsimplant.Cells(rowIndex, colIndex).Interior.Color <> RGB(217, 217, 217) Then
                If iCol <> "" Then
                    If iCol = "01" Or iCol = "16" Then
                        If iRow_1 <> "" Then
                            empTag = "FP5 R" & iCol & " C" & iRow_1 & " T1"
                        Else
                            Exit For
                        End If
                    Else
                        empTag = "FP5 R" & iCol & " C" & iRow_2 & " T1"
                    End If
                    
                    'Saisir les données dans la feuille IT
                    wsIT.Cells(lastRowIT + 1, "A").Value = empTag
                    wsIT.Cells(lastRowIT + 1, "B").Value = reference
                    
                    'Réinitialiser la variable lastRow pour la prochaine itération
                    lastRowIT = wsIT.Cells(wsIT.Rows.Count, "A").End(xlUp).Row
                End If
            End If
        Next rowIndex
    Next colIndex

End Sub

Private Sub Cellule_G(startColumn As Integer, endColumn As Integer, startRow As Integer, endRow As Integer)

    'Déclration des Variables
    Dim wsimplant As Worksheet
    Dim wsIT As Worksheet
    
    Dim lastRowIT As Long
    
    Dim colIndex As Integer
    Dim rowIndex As Integer
    Dim reference As Variant
    
    Dim iRow_1 As String
    Dim iRow_2 As String
    Dim iCol As String
    Dim empTag As String
    
    'Initialisation des feuilles
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    Set wsIT = ThisWorkbook.Sheets("IT")
    
    'Définir la dernière ligne de la feuille IT
    lastRowIT = wsIT.Cells(wsIT.Rows.Count, "A").End(xlUp).Row
    
    'Parcourir la cellule et prendre les coordonnées
    For colIndex = startColumn To endColumn
        For rowIndex = startRow To endRow Step -1
            reference = wsimplant.Cells(rowIndex, colIndex).Value
            iRow_1 = wsimplant.Cells(rowIndex, "AL").Value
            iRow_2 = wsimplant.Cells(rowIndex, "AK").Value
            iRow_3 = wsimplant.Cells(rowIndex, "AM").Value
            iCol = wsimplant.Cells(93, colIndex).Value
            
            If wsimplant.Cells(rowIndex, colIndex).Interior.Color <> RGB(217, 217, 217) Then
                If iCol <> "" Then
                    If iCol = "17" Then
                        empTag = "FP5 R" & iCol & " C" & iRow_1 & " T1"
                    ElseIf iCol = "18" Or iCol = "19" Then
                        empTag = "FP5 R" & iCol & " C" & iRow_2 & " T1"
                    Else
                        empTag = "FP5 R" & iCol & " C" & iRow_3 & " T1"
                    End If
                    
                    'Saisir les données dans la feuille IT
                    wsIT.Cells(lastRowIT + 1, "A").Value = empTag
                    wsIT.Cells(lastRowIT + 1, "B").Value = reference
                    
                    'Réinitialiser la variable lastRow pour la prochaine itération
                    lastRowIT = wsIT.Cells(wsIT.Rows.Count, "A").End(xlUp).Row
                End If
            End If
        Next rowIndex
    Next colIndex

End Sub

Private Sub Affect_Gencode()

    'Déclration des Variables
    Dim wsEP As Worksheet
    Dim wsIT As Worksheet
    
    Dim lastRowEP As Long
    Dim lastRowIT As Integer
    
    Dim rowIndex As Integer
    
    Dim refSelected As Variant
    Dim searchRange As Range
    Dim genCode As Variant
    
    'Initialisation des feuilles
    Set wsEP = ThisWorkbook.Sheets("D_EP")
    Set wsIT = ThisWorkbook.Sheets("IT")

    'Définir les dernières lignes
    lastRowEP = wsEP.Cells(wsEP.Rows.Count, "B").End(xlUp).Row
    lastRowIT = wsIT.Cells(wsIT.Rows.Count, "A").End(xlUp).Row
    
    'Parcourir les feuilles et affecter à une référence à un GenCode
    For rowIndex = 34 To lastRowIT
        refSelected = wsIT.Cells(rowIndex, "B").Value
        If refSelected = "" Then
            GoTo NextEMP
        End If
        
        'Rechercher la référence dans l'export produit
        Set searchRange = wsEP.Columns(2).Find(What:=refSelected, LookIn:=xlValues, LookAt:=xlWhole)
        
        genCode = searchRange.Offset(0, 3).Value
            
        'Saisir le Gen-Code dans la feuille IT
        wsIT.Cells(rowIndex, "D").Value = genCode
        
        'Réinitialiser la varaible refSelected pour la prochaine itération
        refSelected = ""
NextEMP:
    Next rowIndex
        
End Sub

Private Sub ExportIT_1_Main_dynamicArea()

    'Déclration des Variables
    Dim celluleArray(1 To 5) As String
    Dim celluleSelected As Variant
    
    Dim startColumn As Integer
    Dim endColumn As Integer
    
    Dim startRow As Integer
    Dim endRow As Integer
    
    'Initialisation du Tableau des cellules
    celluleArray(1) = "Cellule_A"
    celluleArray(2) = "Cellule_B"
    celluleArray(3) = "Cellule_E"
    celluleArray(4) = "Cellule_F"
    celluleArray(5) = "Cellule_G"
    
    'Définir les coordonnées à traiter en fonction de la cellule en cours
    For Each celluleSelected In celluleArray
        Call Find_Column(celluleSelected, startColumn, endColumn)
        Call Find_Row(celluleSelected, startRow, endRow)
        
        Select Case celluleSelected
            Case "Cellule_A": Call Dynamic_A(startColumn, endColumn, startRow, endRow)
            Case "Cellule_B": Call Dynamic_B(startColumn, endColumn, startRow, endRow)
            Case "Cellule_E": Call Dynamic_E(startColumn, endColumn, startRow, endRow)
            Case "Cellule_F": Call Dynamic_F(startColumn, endColumn, startRow, endRow)
            Case "Cellule_G": Call Dynamic_G(startColumn, endColumn, startRow, endRow)
        End Select
    Next celluleSelected
        
End Sub

Private Sub Dynamic_A(startColumn As Integer, endColumn As Integer, startRow As Integer, endRow As Integer)

    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim wsIT As Worksheet
    
    Dim colIndex As Integer
    Dim rowIndex As Integer
    
    Dim iRow As String
    Dim iCol As String
    
    Dim startEmp As String
    Dim endEmp As String
    Dim stIsActif As Boolean
    
    'Initialisation des feuilles
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    Set wsIT = ThisWorkbook.Sheets("IT")
    
    'Parcourir la cellule et prendre les coordonnées
    For colIndex = startColumn To endColumn
        '(Ré)initialiser le Boolean pour la prochaine itération
        stIsActif = False
        startEmp = ""
        endEmp = ""
        
        For rowIndex = startRow To endRow Step -1
            If wsimplant.Cells(rowIndex, colIndex).Interior.Pattern = xlLightDown Then
                iRow = wsimplant.Cells(rowIndex, "EF").Value
                iCol = wsimplant.Cells(91, colIndex).Value
                
                If stIsActif = False Then
                    startEmp = "AP2 R" & iCol & " C" & iRow & " T1"
                    stIsActif = True
                End If
                
                endEmp = "AP2 R" & iCol & " C" & iRow & " T1"
            End If
        Next rowIndex
        
        'Saisir les plage d'emplacement dans la feuille IT
        Call Write_Dynamic_EMP(startEmp, endEmp)
        
    Next colIndex

End Sub

Private Sub Dynamic_B(startColumn As Integer, endColumn As Integer, startRow As Integer, endRow As Integer)

    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim wsIT As Worksheet
    
    Dim colIndex As Integer
    Dim rowIndex As Integer
    
    Dim iRow As String
    Dim iCol As String
    
    Dim startEmp As String
    Dim endEmp As String
    Dim stIsActif As Boolean
    
    'Initialisation des feuilles
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    Set wsIT = ThisWorkbook.Sheets("IT")
    
    'Parcourir la cellule et prendre les coordonnées
    For colIndex = startColumn To endColumn
        '(Ré)initialiser le Boolean pour la prochaine itération
        stIsActif = False
        startEmp = ""
        endEmp = ""
        
        For rowIndex = startRow To endRow Step -1
            If wsimplant.Cells(rowIndex, colIndex).Interior.Pattern = xlLightDown Then
                iRow = wsimplant.Cells(rowIndex, "DI").Value
                iCol = wsimplant.Cells(99, colIndex).Value
                
                If stIsActif = False Then
                    startEmp = "AP2 R" & iCol & " C" & iRow & " T1"
                    stIsActif = True
                End If
                
                endEmp = "AP2 R" & iCol & " C" & iRow & " T1"
            End If
        Next rowIndex
        
        'Saisir les plage d'emplacement dans la feuille IT
        Call Write_Dynamic_EMP(startEmp, endEmp)
        
    Next colIndex

End Sub

Private Sub Dynamic_E(startColumn As Integer, endColumn As Integer, startRow As Integer, endRow As Integer)

    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim wsIT As Worksheet
    
    Dim colIndex As Integer
    Dim rowIndex As Integer
    
    Dim iRow As String
    Dim iCol As String
    
    Dim startEmp As String
    Dim endEmp As String
    Dim stIsActif As Boolean
    
    'Initialisation des feuilles
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    Set wsIT = ThisWorkbook.Sheets("IT")
    
    'Parcourir la cellule et prendre les coordonnées
    For colIndex = startColumn To endColumn
        '(Ré)initialiser le Boolean pour la prochaine itération
        stIsActif = False
        startEmp = ""
        endEmp = ""
        
        For rowIndex = startRow To endRow Step -1
            If wsimplant.Cells(rowIndex, colIndex).Interior.Pattern = xlLightDown Then
                iRow = wsimplant.Cells(rowIndex, "BZ").Value
                iCol = wsimplant.Cells(91, colIndex).Value
                
                If stIsActif = False Then
                    startEmp = "FP5 R" & iCol & " C" & iRow & " T1"
                    stIsActif = True
                End If
                
                endEmp = "FP5 R" & iCol & " C" & iRow & " T1"
            End If
        Next rowIndex
        
        'Saisir les plage d'emplacement dans la feuille IT
        Call Write_Dynamic_EMP(startEmp, endEmp)
        
    Next colIndex
    
End Sub

Private Sub Dynamic_F(startColumn As Integer, endColumn As Integer, startRow As Integer, endRow As Integer)

    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim wsIT As Worksheet
    
    Dim colIndex As Integer
    Dim rowIndex As Integer
    
    Dim iRow_1 As String
    Dim iRow_2 As String
    Dim iCol As String
    
    Dim startEmp As String
    Dim endEmp As String
    Dim stIsActif As Boolean
    
    'Initialisation des feuilles
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    Set wsIT = ThisWorkbook.Sheets("IT")
    
    'Parcourir la cellule et prendre les coordonnées
    For colIndex = startColumn To endColumn
        '(Ré)initialiser les Varaibles pour la prochaine itération
        stIsActif = False
        startEmp = ""
        endEmp = ""
        
        For rowIndex = startRow To endRow Step -1
            If wsimplant.Cells(rowIndex, colIndex).Interior.Pattern = xlLightDown Then
                iRow_1 = wsimplant.Cells(rowIndex, "BX").Value
                iRow_2 = wsimplant.Cells(rowIndex, "BW").Value
                iCol = wsimplant.Cells(99, colIndex).Value
                
                If stIsActif = False Then
                    If iCol = "01" Or iCol = "16" Then
                        startEmp = "FP5 R" & iCol & " C" & iRow_1 & " T1"
                        stIsActif = True
                    Else
                        startEmp = "FP5 R" & iCol & " C" & iRow_2 & " T1"
                        stIsActif = True
                    End If
                End If
                
                If iCol = "01" Or iCol = "16" Then
                    endEmp = "FP5 R" & iCol & " C" & iRow_1 & " T1"
                Else
                    endEmp = "FP5 R" & iCol & " C" & iRow_2 & " T1"
                End If
            End If
        Next rowIndex
        
        'Saisir les plage d'emplacement dans la feuille IT
        Call Write_Dynamic_EMP(startEmp, endEmp)
        
    Next colIndex
    
End Sub

Private Sub Dynamic_G(startColumn As Integer, endColumn As Integer, startRow As Integer, endRow As Integer)

    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim wsIT As Worksheet
    
    Dim colIndex As Integer
    Dim rowIndex As Integer
    
    Dim iRow_1 As String
    Dim iRow_2 As String
    Dim iRow_3 As String
    Dim iCol As String
    
    Dim startEmp As String
    Dim endEmp As String
    Dim stIsActif As Boolean
    
    'Initialisation des feuilles
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    Set wsIT = ThisWorkbook.Sheets("IT")
    
    'Parcourir la cellule et prendre les coordonnées
    For colIndex = startColumn To endColumn
        '(Ré)initialiser le Boolean pour la prochaine itération
        stIsActif = False
        startEmp = ""
        endEmp = ""
        
        For rowIndex = startRow To endRow Step -1
            If wsimplant.Cells(rowIndex, colIndex).Interior.Pattern = xlLightDown Then
            iRow_1 = wsimplant.Cells(rowIndex, "AL").Value
            iRow_2 = wsimplant.Cells(rowIndex, "AK").Value
            iRow_3 = wsimplant.Cells(rowIndex, "AM").Value
            iCol = wsimplant.Cells(93, colIndex).Value
                
                If stIsActif = False Then
                    If iCol = "17" Then
                        startEmp = "FP5 R" & iCol & " C" & iRow_1 & " T1"
                        stIsActif = True
                    ElseIf iCol = "18" Or iCol = "19" Then
                        startEmp = "FP5 R" & iCol & " C" & iRow_2 & " T1"
                        stIsActif = True
                    Else
                        startEmp = "FP5 R" & iCol & " C" & iRow_3 & " T1"
                        stIsActif = True
                    End If
                End If
                
                If iCol = "17" Then
                    endEmp = "FP5 R" & iCol & " C" & iRow_1 & " T1"
                ElseIf iCol = "18" Or iCol = "19" Then
                    endEmp = "FP5 R" & iCol & " C" & iRow_2 & " T1"
                Else
                    endEmp = "FP5 R" & iCol & " C" & iRow_3 & " T1"
                End If
            End If
        Next rowIndex
        
        'Saisir les plage d'emplacement dans la feuille IT
        Call Write_Dynamic_EMP(startEmp, endEmp)
        
    Next colIndex
    
End Sub

Private Sub Write_Dynamic_EMP(empStart As String, empEnd As String)

    'Déclration des Variables
    Dim wsIT As Worksheet
    Dim lastRowIT As Byte
    Dim rowIndex As Byte
    
    'Initialisation de la feuille
    Set wsIT = ThisWorkbook.Sheets("IT")
    
    'Définir la dernière ligne
    lastRowIT = 30
    
    'Si aucun emplacement dynamique identifié alors next
    If empStart = "" Or empEnd = "" Then
        Exit Sub
    End If
    
    'Saisir les données
    For rowIndex = 3 To 30
        If wsIT.Cells(rowIndex, "A").Value = "" Then
            wsIT.Cells(rowIndex, "A").Value = empStart
            wsIT.Cells(rowIndex, "B").Value = empEnd
            
            'Eclater la saisie du startEmp
            wsIT.Cells(rowIndex, "D").Value = Split(empStart, " ")(0)
            wsIT.Cells(rowIndex, "E").Value = Split(empStart, " ")(1)
            wsIT.Cells(rowIndex, "F").Value = Split(empStart, " ")(2)
            wsIT.Cells(rowIndex, "G").Value = Split(empStart, " ")(3)
    
            'Eclater la saisie du endEmp
            wsIT.Cells(rowIndex, "I").Value = Split(empEnd, " ")(0)
            wsIT.Cells(rowIndex, "J").Value = Split(empEnd, " ")(1)
            wsIT.Cells(rowIndex, "K").Value = Split(empEnd, " ")(2)
            wsIT.Cells(rowIndex, "L").Value = Split(empEnd, " ")(3)
            
            Exit For
        End If
    Next rowIndex
    
End Sub

Private Sub ExportIT_1_Main_clearIT()

    'Déclration des Variables
    Dim wsIT As Worksheet
    Dim lastRow As Long
    Dim fixRange As Range
    Dim dynamicRange As Range
    
    'Initialisation de la feuille
    Set wsIT = ThisWorkbook.Sheets("IT")
    
    'Définir la dernière ligne de la feuille
    lastRow = wsIT.Cells(wsIT.Rows.Count, "A").End(xlUp).Row
    
    'Initialisation des plages de données à supprimer
    Set dynamicRange = wsIT.Range("A3:L30")
    Set fixRange = wsIT.Range("A34:D" & lastRow + 1)
    
    'Clear les plages
    dynamicRange.ClearContents
    fixRange.ClearContents
    
End Sub
