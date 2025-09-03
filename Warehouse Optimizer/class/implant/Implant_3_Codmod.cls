VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Implant_3_Codmod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe d'implantation des CodMod

'Implémentation de la classe mère
Implements Implant_1_Main

Private Sub Implant_1_Main_ImplantRef()

    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim wsABCCM As Worksheet
    Dim lastRow As Integer
    Dim abcCMIndex As Integer
    
    Dim selectedRef As Variant
    Dim needEMP As Byte
    Dim Class As String
    
    Dim colIndex As Integer
    Dim startColumn As Integer
    Dim endColumn As Integer
    
    Dim rowIndex As Integer
    Dim permStart As Integer
    Dim permEnd As Integer
    
    Dim nbreOfSupport As Byte
    
    Dim nbreSecondaire As Variant
    Dim secondIndex As Byte
    
    'Initialisation des feuilles
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    Set wsABCCM = ThisWorkbook.Sheets("ABC Code Modèle")
    
    'Définir la dernière ligne de la feuille ABCCM
    lastRow = wsABCCM.Cells(wsABCCM.Rows.Count, "C").End(xlUp).Row
    
    'Définir les settings des colonnes d'implantations
    Call SetSensOfimplant(startColumn, endColumn)
    
    'Parcourir l'ABC au CodMod et saisir les données dans l'implantation
    For abcCMIndex = 4 To lastRow
        selectedRef = wsABCCM.Cells(abcCMIndex, "C").Value
        'Sauter les lignes Vides
        If selectedRef <> "" Then
            needEMP = Need_EMP(selectedRef)
            Class = Search_Class(selectedRef)
        
            'Définir les lignes à sélectionner en fonction des Settings
            Call SetPermissions(Class, permStart, permEnd)
            
            'Parcourrir le nombre d'emplacements nécessaires
            For nbreOfSupport = 1 To needEMP
                'Parcourir les colonnes
                If endColumn > startColumn Then
                    For colIndex = startColumn To endColumn
                        'Parcourrir les lignes
                        For rowIndex = permStart To permEnd Step -1
                            
                            'Vérifications avant implantation
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
                    'Parcourir les colonnes
                    For colIndex = startColumn To endColumn Step -1
                        'Parcourrir les lignes
                        For rowIndex = permStart To permEnd Step -1
                            
                            'Vérifications avant implantation
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
                                        wsimplant.Cells(rowIndex - 1, colIndex).Value = "s"
                                    Next secondIndex
                                End If
                                GoTo NextEMP
                            End If
                        Next rowIndex
                    Next colIndex
                End If
NextEMP:
            Next nbreOfSupport
        End If
    Next abcCMIndex
    
    'Allouer les emplacements du picking dynamique
    Call Allocate_Picking_Dynamic
        
End Sub

Private Sub Implant_1_Main_GenerateColor()
    
    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim celluleSelected As String
    Dim dictCouleur As Object
    Dim celluleRange As Range
    Dim cell As Range
    Dim ref As Variant
    Dim cm As Variant
    
    'Initialisation de la feuille
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    
    'Définir le paramètre de cellule d'implantation
    celluleSelected = GetSettings("Cellule d'implantation")
    
    'Initialisation de la plage à colorier en fonction de la cellule choisi
    Select Case celluleSelected
        Case "Cellule_A": Set celluleRange = wsimplant.Range("ES3:FX90")
        Case "Cellule_B": Set celluleRange = wsimplant.Range("DJ3:EO98")
        Case "Cellule_E": Set celluleRange = wsimplant.Range("CA3:DF90")
        Case "Cellule_F": Set celluleRange = wsimplant.Range("AQ3:BV98")
        Case "Cellule_G": Set celluleRange = wsimplant.Range("E3:AJ92")
    End Select
    
    'Initialisation d'un dictionnaire de couleur
    Set dictCouleur = CreateObject("Scripting.Dictionary")
    
    'Parcourrir les celulles et affecter une couleur si les conditions sont remplis et par CM
    For Each cell In celluleRange
        If cell.Value <> "" And _
           cell.Interior.Color <> RGB(217, 217, 217) And _
           cell.Borders(xlDiagonalDown).LineStyle <> xlContinuous And cell.Borders(xlDiagonalUp).LineStyle <> xlContinuous And _
           cell.Interior.Pattern <> xlGrid Then
           
            ref = cell.Value
            cm = Search_CM(ref)
           
            If Not dictCouleur.exists(cm) Then
                'Générer une couleur aléatoire
                r = 50 + Int((205 + 1) * Rnd)
                g = 50 + Int((205 + 1) * Rnd)
                b = 50 + Int((205 + 1) * Rnd)
                dictCouleur.Add cm, RGB(r, g, b)
            End If
            
            'Appliquer la couleur
            cell.Interior.Color = dictCouleur(cm)
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
