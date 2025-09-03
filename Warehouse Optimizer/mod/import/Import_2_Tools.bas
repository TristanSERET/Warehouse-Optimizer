Attribute VB_Name = "Import_2_Tools"
Public Sub Check_Format(wbToImport As Workbook, wbWo As Workbook, wsWo As Worksheet, wsToImport As Worksheet)

    'Déclaration des Variables
    Dim Index As Long
    Dim lastCol As Byte
    Dim lastColT As Byte
    Dim referenceCollection As New Collection
    Dim testerCollection As New Collection
    
    'Définir les dernières colonnes
    lastCol = wsWo.Cells(1, wsWo.Columns.Count).End(xlToLeft).Column
    lastColT = wsToImport.Cells(1, wsToImport.Columns.Count).End(xlToLeft).Column
    
    'Ajouter les données de références sur le claseur de référence
    For colIndex = 1 To lastCol
        referenceCollection.Add wsWo.Cells(1, colIndex).Value
    Next colIndex
    
    'Ajouter les données à tester sur le classeur de test
    For colIndex = 1 To lastColT
        testerCollection.Add wsToImport.Cells(1, colIndex).Value
    Next colIndex
    
    'Vérifier si les deux classeur sont identique
    For Index = 1 To lastCol
        If testerCollection(Index) <> referenceCollection(Index) Then
            MsgBox "Votre importation à échouer, la structure du classeur d'importation est non valide. Vérifiez également les en-têtes.", vbCritical, "Error"
            End
        End If
    Next Index

End Sub

Public Function Path() As Variant

    Path = Application.GetOpenFilename(FileFilter:="Classeurs Excel (*.xlsx; *.csv), *.xlsx; *.csv", Title:="Sélectionnez le classeur à importer")
    
    If Path = False Then
        MsgBox "Aucun fichier sélectionné.", vbExclamation
    End If

End Function
