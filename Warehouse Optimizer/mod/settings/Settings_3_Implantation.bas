Attribute VB_Name = "Settings_3_Implantation"
Public Sub SetSensOfimplant(ByRef startColumn As Integer, ByRef endColumn As Integer)

    'Déclaration des Variables
    Dim sensSelected As Variant
    Dim globalStart As Integer
    Dim globalEnd As Integer
    Dim localStart As Integer
    Dim localEnd As Integer
    Dim endCol As Integer
    
    'Définir les colonnes Global en fonction des settings de cellule
    Call GlobalRange(globalStart, globalEnd)
    
    'Rechercher le sens sélectionné
    sensSelected = GetSettings("Sens d'implantation")
    
    Select Case sensSelected
        Case "Gauche à Droite": startColumn = Local_Start_Range: endColumn = globalEnd
        Case "Droite à Gauche": startColumn = Local_Start_Range: endColumn = globalStart
    End Select

End Sub

Private Sub GlobalRange(ByRef startColumn As Integer, ByRef endColumn As Integer)

    'Déclaration des Variables
    Dim celluleSelected As Variant
    
    'Rechercher la cellule sélectionnée
    celluleSelected = GetSettings("Cellule d'implantation")
    
    'Définir les colonnes de références
    Select Case celluleSelected
        Case "Cellule_A": startColumn = 149: endColumn = 180
        Case "Cellule_B": startColumn = 114: endColumn = 145
        Case "Cellule_E": startColumn = 79: endColumn = 110
        Case "Cellule_F": startColumn = 43: endColumn = 74
        Case "Cellule_G": startColumn = 5: endColumn = 36
    End Select
End Sub

Private Function Local_Start_Range() As Integer

    'Déclaration des Variables
    Dim startRangee As Byte
    Dim secteur As String
    
    'Rechercher le secteur de Départ AP5 ou FP5
    secteur = GetSettings("Cellule d'implantation")
    
    'Rechercher la rangée de départ en fonction des Settings
    startRangee = GetSettings("Rangée de départ")
    
    If secteur = "Cellule_E" Or secteur = "Cellule_F" Or secteur = "Cellule_G" Then
    
        Select Case startRangee
            'Cellule F
            Case 1: Local_Start_Range = 74
            Case 2: Local_Start_Range = 71
            Case 3: Local_Start_Range = 70
            Case 4: Local_Start_Range = 67
            Case 5: Local_Start_Range = 66
            Case 6: Local_Start_Range = 63
            Case 7: Local_Start_Range = 62
            Case 8: Local_Start_Range = 59
            Case 9: Local_Start_Range = 58
            Case 10: Local_Start_Range = 55
            Case 11: Local_Start_Range = 54
            Case 12: Local_Start_Range = 51
            Case 13: Local_Start_Range = 50
            Case 14: Local_Start_Range = 47
            Case 15: Local_Start_Range = 46
            Case 16: Local_Start_Range = 43
            
            'Cellule G
            Case 17: Local_Start_Range = 36
            Case 18: Local_Start_Range = 33
            Case 19: Local_Start_Range = 32
            Case 20: Local_Start_Range = 29
            Case 21: Local_Start_Range = 28
            Case 22: Local_Start_Range = 25
            Case 23: Local_Start_Range = 24
            Case 24: Local_Start_Range = 21
            Case 25: Local_Start_Range = 20
            Case 26: Local_Start_Range = 17
            Case 27: Local_Start_Range = 16
            Case 28: Local_Start_Range = 13
            Case 29: Local_Start_Range = 12
            Case 30: Local_Start_Range = 9
            Case 31: Local_Start_Range = 8
            Case 32: Local_Start_Range = 5
            
            'Cellule E
            Case 35: Local_Start_Range = 79
            Case 36: Local_Start_Range = 82
            Case 37: Local_Start_Range = 83
            Case 38: Local_Start_Range = 86
            Case 39: Local_Start_Range = 87
            Case 40: Local_Start_Range = 90
            Case 41: Local_Start_Range = 91
            Case 42: Local_Start_Range = 94
            Case 43: Local_Start_Range = 95
            Case 44: Local_Start_Range = 98
            Case 45: Local_Start_Range = 99
            Case 46: Local_Start_Range = 102
            Case 47: Local_Start_Range = 103
            Case 48: Local_Start_Range = 106
            Case 49: Local_Start_Range = 107
            Case 50: Local_Start_Range = 110
            
        End Select
    ElseIf secteur = "Cellule_A" Or secteur = "Cellule_B" Then
        Select Case startRangee
            'Cellule A
            Case 1: Local_Start_Range = 180
            Case 2: Local_Start_Range = 177
            Case 3: Local_Start_Range = 176
            Case 4: Local_Start_Range = 173
            Case 5: Local_Start_Range = 172
            Case 6: Local_Start_Range = 169
            Case 7: Local_Start_Range = 168
            Case 8: Local_Start_Range = 165
            Case 9: Local_Start_Range = 164
            Case 10: Local_Start_Range = 161
            Case 11: Local_Start_Range = 160
            Case 12: Local_Start_Range = 157
            Case 13: Local_Start_Range = 156
            Case 14: Local_Start_Range = 153
            Case 15: Local_Start_Range = 152
            Case 16: Local_Start_Range = 149
            
            'Cellule B
            Case 17: Local_Start_Range = 145
            Case 18: Local_Start_Range = 142
            Case 19: Local_Start_Range = 141
            Case 20: Local_Start_Range = 138
            Case 21: Local_Start_Range = 137
            Case 22: Local_Start_Range = 134
            Case 23: Local_Start_Range = 133
            Case 24: Local_Start_Range = 130
            Case 25: Local_Start_Range = 129
            Case 26: Local_Start_Range = 126
            Case 27: Local_Start_Range = 125
            Case 28: Local_Start_Range = 122
            Case 29: Local_Start_Range = 121
            Case 30: Local_Start_Range = 118
            Case 31: Local_Start_Range = 117
            Case 32: Local_Start_Range = 114
        End Select
    End If
        
End Function


Public Function SetPermissions(Class As String, ByRef permStart As Integer, ByRef permEnd As Integer) As Variant

    'Déclaration des Variables
    Dim permissionSelected As Variant
    Dim setCellule As String
    
    'Définir le paramètre de cellule
    setCellule = GetSettings("Cellule d'implantation")
    
    'Définir les permissions d'implantation vis à vis de la class
        Select Case Class
            Case "A": permissionSelected = GetSettings("Autorisation d'implantation Classe A")
            Case "B": permissionSelected = GetSettings("Autorisation d'implantation Classe B")
            Case "C": permissionSelected = GetSettings("Autorisation d'implantation Classe C")
        End Select
        
        'En Fonction de la cellule appliquer des valeurs différentes
        Select Case setCellule
            Case "Cellule_A"
                Select Case permissionSelected
                    Case "Avant passage chariot uniquement": permStart = 90: permEnd = 30
                    Case "Après passage chariot uniquement": permStart = 29: permEnd = 3
                    Case "Tout": permStart = 90: permEnd = 3
                End Select
            Case "Cellule_B"
                Select Case permissionSelected
                    Case "Avant passage chariot uniquement": permStart = 98: permEnd = 39
                    Case "Après passage chariot uniquement": permStart = 38: permEnd = 3
                    Case "Tout": permStart = 98: permEnd = 3
                End Select
            Case "Cellule_E"
                Select Case permissionSelected
                    Case "Avant passage chariot uniquement": permStart = 90: permEnd = 30
                    Case "Après passage chariot uniquement": permStart = 29: permEnd = 3
                    Case "Tout": permStart = 90: permEnd = 3
                End Select
            Case "Cellule_F"
                Select Case permissionSelected
                    Case "Avant passage chariot uniquement": permStart = 98: permEnd = 35
                    Case "Après passage chariot uniquement": permStart = 30: permEnd = 3
                    Case "Tout": permStart = 98: permEnd = 3
                End Select
            Case "Cellule_G"
                Select Case permissionSelected
                    Case "Avant passage chariot uniquement": permStart = 90: permEnd = 30
                    Case "Après passage chariot uniquement": permStart = 30: permEnd = 3
                    Case "Tout": permStart = 90: permEnd = 3
                End Select
        End Select
        
End Function

Public Function Need_EMP(ref As Variant) As Byte

    'Déclaration des Variables
    Dim wsCalcul As Worksheet
    Dim lastRow As Long
    Dim rangeRef As Range
    
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'Définir la dernière ligne
    lastRow = wsCalcul.Cells(wsCalcul.Rows.Count, "B").End(xlUp).Row
    
    'Initialisation de la plage de recherche en colonne B
    Set rangeRef = wsCalcul.Columns(2).Find(What:=ref, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rangeRef Is Nothing Then
        Need_EMP = rangeRef.Offset(0, Select_Calcul).Value * NbreEMP_Alveole
        
    Else
        Need_EMP = CVErr(xlErrNA)
        MsgBox "La référence est introuvable !", vbExclamation, "Error"
    End If
    
End Function

Private Function NbreEMP_Alveole() As Byte

    'Déclaration des Variables
    Dim selectedCellule As String
    Dim nbreByAlvéole As Byte
    
    'Définir la cellule en fonction des options
    selectedCellule = GetSettings("Cellule d'implantation")
    
    'Déterminer le nombre d'emplacement / Alvéole en fonction des cellules
    Select Case selectedCellule
        Case "Cellule_A": NbreEMP_Alveole = 3
        Case "Cellule_B": NbreEMP_Alveole = 4
        Case "Cellule_E": NbreEMP_Alveole = 3
        Case "Cellule_F": NbreEMP_Alveole = 4
        Case "Cellule_G": NbreEMP_Alveole = 3
    End Select

End Function

Private Function Select_Calcul() As Byte

    'Déclaration des Variables
    Dim selectedCalcul As String
    
    'Définir le calcul en fonction des settings
    selectedCalcul = GetSettings("Calcul retenu en sortie")
    
    'Déterminer le calcul sélectionné en Sortie
    Select Case selectedCalcul
        Case "Meilleure Moyenne": Select_Calcul = 77
        Case "Max": Select_Calcul = 67
        Case "Moyenne": Select_Calcul = 72
    End Select
    
End Function

Public Function Search_Class(ref As Variant) As String

    'Déclaration des Variables
    Dim wsCM As Worksheet
    Dim lastRow As Integer
    Dim searchRef As Range
    Dim cm As Variant
    Dim searchTotalCm As Range
    
    'Initialisation de la feuille
    Set wsCM = ThisWorkbook.Sheets("ABC Code Modèle")
    
    'Définir la dernière ligne
    lastRow = wsCM.Cells(wsCM.Rows.Count, "C").End(xlUp).Row
    
    'Rechercher la référence
    Set searchRef = wsCM.Range("C4:C" & lastRow).Find(What:=ref, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Définir le code modèle
    cm = searchRef.Offset(0, -1).Value
    
    'Trouver la ligne Total du Codemodèle en question
    Set searchTotalCm = wsCM.Range("B4:B" & lastRow + 1).Find(What:="Total " & cm, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Trouver la class du Codemodèle
    Search_Class = searchTotalCm.Offset(0, 7).Value

End Function

Public Function Search_CM(ref As Variant) As Variant

    'Déclaration des Variables
    Dim wsCM As Worksheet
    Dim lastRow As Integer
    Dim searchRef As Range

    'Initialisation de la feuille
    Set wsCM = ThisWorkbook.Sheets("ABC Code Modèle")
    
    'Définir la dernière ligne
    lastRow = wsCM.Cells(wsCM.Rows.Count, "C").End(xlUp).Row
    
    'Rechercher la référence
    Set searchRef = wsCM.Range("C4:C" & lastRow).Find(What:=ref, LookIn:=xlValues, LookAt:=xlWhole)
    
    'Définir le code modèle
    Search_CM = searchRef.Offset(0, -1).Value

End Function

Public Sub Allocate_Picking_Dynamic()

    'Déclaration des Variables
    Dim affectation As String

    'Initialisation des Settings Automatique or Manuel
    affectation = GetSettings("Affectation du Picking Dynamique")
    
    Select Case affectation
        Case "Automatique": Call Allocate_Auto
        Case "Manuelle": Call Allocate_Manuel
        Case Else: MsgBox "Le type d'affectation du Picking Dynamqiue n'a pas été sélectionné dans les paramètres le calcul n'a pas pu être effectué !", vbExclamation, "Error": Exit Sub
    End Select
    
End Sub

Private Sub Allocate_Auto()

    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim celluleSelected As Variant
    Dim nombreAlveole As Double
    Dim needEMP As Byte
    Dim Index As Byte
    
    Dim startCol As Integer
    Dim endCol As Integer
    Dim colIndex As Integer
    
    Dim permStart As Integer
    Dim permEnd As Integer
    Dim rowIndex As Integer

    'Initialisation de la feuille
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    
    'Définition des Settings
    celluleSelected = GetSettings("Cellule d'implantation")
    nombreAlveole = Calcul_Alveole
    Call SetSensOfimplant(startCol, endCol)
    Call DynamicPermission(permStart, permEnd)
    
    'Définir le besoin en emplacement
    Select Case celluleSelected
        Case "Cellule_A": needEMP = 3 * nombreAlveole
        Case "Cellule_B": needEMP = 4 * nombreAlveole
        Case "Cellule_E": needEMP = 3 * nombreAlveole
        Case "Cellule_F": needEMP = 4 * nombreAlveole
        Case "Cellule_G": needEMP = 3 * nombreAlveole
    End Select
    
    For Index = 1 To needEMP
        If endCol > startCol Then
            For colIndex = startCol To endCol
                For rowIndex = permStart To permEnd Step -1
                    If wsimplant.Cells(rowIndex, colIndex).Value = "" And _
                        wsimplant.Cells(rowIndex, colIndex).Interior.Color <> RGB(217, 217, 217) And _
                        wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalDown).LineStyle <> xlContinuous And wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalUp).LineStyle <> xlContinuous And _
                        wsimplant.Cells(rowIndex, colIndex).Interior.Pattern <> xlGrid And _
                        wsimplant.Cells(rowIndex, colIndex).Interior.Pattern <> xlLightDown Then
                        
                        'Appliquer la mise en forme
                        wsimplant.Cells(rowIndex, colIndex).Interior.Pattern = xlLightDown
                        GoTo Next_Emp
                    End If
                Next rowIndex
            Next colIndex
        Else
            For colIndex = startCol To endCol Step -1
                For rowIndex = permStart To permEnd Step -1
                    If wsimplant.Cells(rowIndex, colIndex).Value = "" And _
                        wsimplant.Cells(rowIndex, colIndex).Interior.Color <> RGB(217, 217, 217) And _
                        wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalDown).LineStyle <> xlContinuous And wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalUp).LineStyle <> xlContinuous And _
                        wsimplant.Cells(rowIndex, colIndex).Interior.Pattern <> xlGrid And _
                        wsimplant.Cells(rowIndex, colIndex).Interior.Pattern <> xlLightDown Then
                        
                        'Appliquer la mise en forme
                        wsimplant.Cells(rowIndex, colIndex).Interior.Pattern = xlLightDown
                        GoTo Next_Emp
                    End If
                Next rowIndex
            Next colIndex
        End If
Next_Emp:
    Next Index

End Sub

Private Function Calcul_Alveole() As Double

    'Déclartion des Variables
    Dim wsABC As Worksheet
    Dim lastRow As Integer
    Dim sumOfAlveole As Double
    Dim needOfAlveole As Double
    Dim calculSelected As Variant
    
    'Initialisation de la feuille
    Set wsABC = ThisWorkbook.Sheets("ABC")
    
    'Définir la dernière ligne
    lastRow = wsABC.Cells(wsABC.Rows.Count, "B").End(xlUp).Row
    
    'Faire la somme des alvéoles nécessaires
    sumOfAlveole = Application.WorksheetFunction.sum(wsABC.Range("I3:I" & lastRow))
    
    'En fonction des settings définir un % du besoin
    calculSelected = GetSettings("Calcul retenu en sortie")
    
    Select Case calculSelected
        Case "Max": Calcul_Alveole = sumOfAlveole * 0.1
        Case " Meilleure Moyenne": Calcul_Alveole = sumOfAlveole * 0.15
        Case "Moyenne": Calcul_Alveole = sumOfAlveole * 0.2
    End Select
 
End Function

Private Sub Allocate_Manuel()

    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim celluleSelected As Variant
    Dim nombreAlveole As Variant
    Dim startColumn As Integer
    Dim endColumn As Integer
    Dim colIndex As Integer
    Dim permStart As Integer
    Dim permEnd As Integer
    Dim rowIndex As Integer
    Dim Index As Integer
    Dim needEMP As Byte
    
    'Initialisation de la feuille
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    
    'Définition des Settings
    celluleSelected = GetSettings("Cellule d'implantation")
    nombreAlveole = GetSettings("Nombre d'alvéoles à allouer")
    Call SetSensOfimplant(startColumn, endColumn)
    Call DynamicPermission(permStart, permEnd)
    
    'Définir le besoin en emplacement
    Select Case celluleSelected
        Case "Cellule_A": needEMP = 3 * nombreAlveole
        Case "Cellule_B": needEMP = 4 * nombreAlveole
        Case "Cellule_E": needEMP = 3 * nombreAlveole
        Case "Cellule_F": needEMP = 4 * nombreAlveole
        Case "Cellule_G": needEMP = 3 * nombreAlveole
    End Select
    
    'Parcourrir les colonnes et les lignes
    For Index = 1 To needEMP
        If endColumn > startColumn Then
            For colIndex = startColumn To endColumn
                For rowIndex = permStart To permEnd Step -1
                    If wsimplant.Cells(rowIndex, colIndex).Value = "" And _
                        wsimplant.Cells(rowIndex, colIndex).Interior.Color <> RGB(217, 217, 217) And _
                        wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalDown).LineStyle <> xlContinuous And wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalUp).LineStyle <> xlContinuous And _
                        wsimplant.Cells(rowIndex, colIndex).Interior.Pattern <> xlGrid And _
                        wsimplant.Cells(rowIndex, colIndex).Interior.Pattern <> xlLightDown Then
                        
                        'Appliquer la mise en forme
                        wsimplant.Cells(rowIndex, colIndex).Interior.Pattern = xlLightDown
                        GoTo Next_Emp
                    End If
                Next rowIndex
            Next colIndex
        Else
            For colIndex = startColumn To endColumn Step -1
                For rowIndex = permStart To permEnd Step -1
                    If wsimplant.Cells(rowIndex, colIndex).Value = "" And _
                        wsimplant.Cells(rowIndex, colIndex).Interior.Color <> RGB(217, 217, 217) And _
                        wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalDown).LineStyle <> xlContinuous And wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalUp).LineStyle <> xlContinuous And _
                        wsimplant.Cells(rowIndex, colIndex).Interior.Pattern <> xlGrid And _
                        wsimplant.Cells(rowIndex, colIndex).Interior.Pattern <> xlLightDown Then
                        
                        'Appliquer la mise en forme
                        wsimplant.Cells(rowIndex, colIndex).Interior.Pattern = xlLightDown
                        GoTo Next_Emp
                    End If
                Next rowIndex
            Next colIndex
        End If
            
Next_Emp:
    Next Index

End Sub

Private Sub DynamicPermission(ByRef permStart As Integer, ByRef permEnd As Integer)

    'Déclaration des Variables
    Dim permissionSelected As Variant
    Dim setCellule As Variant
    
    'Définir le paramètre de cellule
    setCellule = GetSettings("Cellule d'implantation")
    permissionSelected = GetSettings("Positionnement du Picking Dynamique")
        
    'En Fonction de la cellule appliquer des valeurs différentes
    Select Case setCellule
            Case "Cellule_A"
            Select Case permissionSelected
                Case "Avant passage chariot uniquement": permStart = 90: permEnd = 30
                Case "Après passage chariot uniquement": permStart = 29: permEnd = 3
                Case "Tout": permStart = 90: permEnd = 3
            End Select
            Case "Cellule_B"
            Select Case permissionSelected
                Case "Avant passage chariot uniquement": permStart = 98: permEnd = 39
                Case "Après passage chariot uniquement": permStart = 38: permEnd = 3
                Case "Tout": permStart = 98: permEnd = 3
            End Select
        Case "Cellule_E"
            Select Case permissionSelected
                Case "Avant passage chariot uniquement": permStart = 90: permEnd = 30
                Case "Après passage chariot uniquement": permStart = 29: permEnd = 3
                Case "Tout": permStart = 90: permEnd = 3
            End Select
        Case "Cellule_F"
            Select Case permissionSelected
                Case "Avant passage chariot uniquement": permStart = 98: permEnd = 35
                Case "Après passage chariot uniquement": permStart = 30: permEnd = 3
                Case "Tout": permStart = 98: permEnd = 3
            End Select
        Case "Cellule_G"
            Select Case permissionSelected
                Case "Avant passage chariot uniquement": permStart = 92: permEnd = 30
                Case "Après passage chariot uniquement": permStart = 30: permEnd = 3
                Case "Tout": permStart = 92: permEnd = 3
            End Select
    End Select
        
End Sub
