VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Implant_2_REF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe d'implantation des références

'Implementation de la classe mère
Implements Implant_1_Main

Private Sub Implant_1_Main_ImplantRef()

    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim wsABC As Worksheet
    Dim lastRowABC As Integer
    Dim abcIndex As Integer
    
    Dim selectedRef As Variant
    Dim needEMP As Byte
    Dim Class As String
    
    Dim colIndex As Integer
    Dim startColumn As Integer
    Dim endColumn As Integer
    
    Dim rowIndex As Integer
    Dim permStart As Integer
    Dim permEnd As Integer
    
    Dim nbresOfSupport As Byte
    
    Dim nbreSecondaire As Variant
    Dim secondIndex As Byte
    
    'Initialisation des feuilles
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    Set wsABC = ThisWorkbook.Sheets("ABC")
    
    'Définir la dernière ligne dela feuille ABC
    lastRowABC = wsABC.Cells(wsABC.Rows.Count, "B").End(xlUp).Row
    
    'Définir les settings de colonnes d'implantation
    Call SetSensOfimplant(startColumn, endColumn)
    
    'Parcourir l'ABC et saisir les données dans l'implantation
    For abcIndex = 3 To lastRowABC
        selectedRef = wsABC.Cells(abcIndex, "B").Value
        needEMP = Need_EMP(selectedRef)
        Class = wsABC.Cells(abcIndex, "J").Value
        
        'Définir les lignes à sélectionner en Fonction des settings
        Call SetPermissions(Class, permStart, permEnd)

        'Parcourir le nombre d'emplacement nécessaire
        For nbresOfSupport = 1 To needEMP
            'Parcourrir les colonnes
            If endColumn > startColumn Then
                For colIndex = startColumn To endColumn
                    'Parcourrir les lignes
                    For rowIndex = permStart To permEnd Step -1
    
                        'Ajouter des Vérifications de positionnement
                        If wsimplant.Cells(rowIndex, colIndex).Value = "" And _
                           wsimplant.Cells(rowIndex, colIndex).Interior.Color <> RGB(217, 217, 217) And _
                           wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalDown).LineStyle <> xlContinuous And wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalUp).LineStyle <> xlContinuous And _
                           wsimplant.Cells(rowIndex, colIndex).Interior.Pattern <> xlGrid And _
                           wsimplant.Cells(rowIndex, colIndex).Interior.Pattern <> xlLightDown Then
                            'Saisir la référence
                            wsimplant.Cells(rowIndex, colIndex).Value = selectedRef
                            
                            'Saisir l'emplacement secondaire si nécessaire
                            If Second_EMP(selectedRef, nbreSecondaire) = True And rowIndex > permEnd And _
                               wsimplant.Cells(rowIndex - 1, colIndex).Value = "" And _
                               wsimplant.Cells(rowIndex - 1, colIndex).Interior.Color <> RGB(217, 217, 217) And _
                               wsimplant.Cells(rowIndex - 1, colIndex).Borders(xlDiagonalDown).LineStyle <> xlContinuous And wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalUp).LineStyle <> xlContinuous And _
                               wsimplant.Cells(rowIndex - 1, colIndex).Interior.Pattern <> xlGrid And _
                               wsimplant.Cells(rowIndex - 1, colIndex).Interior.Pattern <> xlLightDown Then
                                
                                'Saisir les emplacements secondaires
                                For secondIndex = 1 To nbreSecondaire
                                    wsimplant.Cells(rowIndex - secondIndex, colIndex).Value = "s"
                                Next secondIndex
                            End If
                            GoTo NextEMP
                        End If
                    Next rowIndex
                Next colIndex
            Else
                For colIndex = startColumn To endColumn Step -1
                    'Parcourrir les lignes
                    For rowIndex = permStart To permEnd Step -1
                        'Ajouter des Vérifications de positionnement
                        If wsimplant.Cells(rowIndex, colIndex).Value = "" And _
                           wsimplant.Cells(rowIndex, colIndex).Interior.Color <> RGB(217, 217, 217) And _
                           wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalDown).LineStyle <> xlContinuous And wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalUp).LineStyle <> xlContinuous And _
                           wsimplant.Cells(rowIndex, colIndex).Interior.Pattern <> xlGrid And _
                           wsimplant.Cells(rowIndex, colIndex).Interior.Pattern <> xlLightDown Then
                            'Saisir la référence
                            wsimplant.Cells(rowIndex, colIndex).Value = selectedRef
                            
                            'Saisir l'emplacement secondaire si nécessaire
                            If Second_EMP(selectedRef, nbreSecondaire) = True And rowIndex > permEnd And _
                               wsimplant.Cells(rowIndex - 1, colIndex).Value = "" And _
                               wsimplant.Cells(rowIndex - 1, colIndex).Interior.Color <> RGB(217, 217, 217) And _
                               wsimplant.Cells(rowIndex - 1, colIndex).Borders(xlDiagonalDown).LineStyle <> xlContinuous And wsimplant.Cells(rowIndex, colIndex).Borders(xlDiagonalUp).LineStyle <> xlContinuous And _
                               wsimplant.Cells(rowIndex - 1, colIndex).Interior.Pattern <> xlGrid And _
                               wsimplant.Cells(rowIndex - 1, colIndex).Interior.Pattern <> xlLightDown Then

                                'Saisir les emplacements secondaires
                                For secondIndex = 1 To nbreSecondaire
                                    wsimplant.Cells(rowIndex - secondIndex, colIndex).Value = "s"
                                Next secondIndex
                            End If
                            GoTo NextEMP
                        End If
                    Next rowIndex
                Next colIndex
            End If
NextEMP:
        Next nbresOfSupport
    Next abcIndex
    
    'Allouer les emplacements du picking dynamique
    Call Allocate_Picking_Dynamic

End Sub

Private Sub Implant_1_Main_GenerateColor()

    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim celluleRange As Range
    Dim cell As Range
    Dim celluleSelected As String
    Dim dictCouleur As Object
    Dim ref As Variant
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
    
    'Initialisation de la feuille
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    
    'Définir le paramètre
    celluleSelected = GetSettings("Cellule d'implantation")
    
    'Initialisation de la plage à colorier en fonction de la cellule choisi
    Select Case celluleSelected
        Case "Cellule_E": Set celluleRange = wsimplant.Range("CA3:DF90")
        Case "Cellule_F": Set celluleRange = wsimplant.Range("AQ3:BV98")
        Case "Cellule_G": Set celluleRange = wsimplant.Range("E3:AJ92")
    End Select
    
    'Initialisation d'un dictionnaire de Couleur
    Set dictCouleur = CreateObject("Scripting.Dictionary")
    
    'Parcourrir les cellules et affecter une couleur si les conditions sont remplis
    For Each cell In celluleRange
        If cell.Value <> "" And _
           cell.Interior.Color <> RGB(217, 217, 217) And _
           cell.Borders(xlDiagonalDown).LineStyle <> xlContinuous And cell.Borders(xlDiagonalUp).LineStyle <> xlContinuous And _
           cell.Interior.Pattern <> xlGrid Then
           
            ref = cell.Value
            If Not dictCouleur.exists(ref) Then
                'Générer une couleur aléatoire
                r = 50 + Int((205 + 1) * Rnd)
                g = 50 + Int((205 + 1) * Rnd)
                b = 50 + Int((205 + 1) * Rnd)
                dictCouleur.Add ref, RGB(r, g, b)
            End If
            'Appliquer la couleur
            cell.Interior.Color = dictCouleur(ref)
        End If
    Next cell
End Sub

Private Function Second_EMP(selectedRef As Variant, ByRef needEMP As Variant) As Boolean

    'Déclaration des Variables
    Dim wsCalcul As Worksheet
    Dim refRange As Range
    Dim typeOfSupport As Variant
    Dim alveoleSelected As Variant
    
    'Initialisation de la feuille
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    
    'Définir les settings
    typeOfSupport = GetSettings("Type de support logistique")
    alveoleSelected = GetSettings("Cellule d'implantation")
    
    'Trouver la référence sélectionnée dans la base de calcul et le nombre d'emplacement nécessaire
    Set refRange = wsCalcul.Columns(2).Find(What:=selectedRef, LookIn:=xlValues, LookAt:=xlWhole)
    needEMP = refRange.Offset(0, 6).Value - 1
    
    'Modifier le résultat de la fonction selon les conditions
    Second_EMP = False
    
    If typeOfSupport = "Palette 80x120" And needEMP > 1 Then
        Second_EMP = True
    End If
    
End Function
