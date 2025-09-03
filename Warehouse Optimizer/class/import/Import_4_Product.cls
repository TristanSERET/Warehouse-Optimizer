VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Import_4_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class d'Importation des Articles

'Implémenter la class primaire
Implements Import_1_Main

Private Sub Import_1_Main_folderPath()

    'Déclaration des Variables
    Dim wbWo As Workbook
    Dim wbProduct As Workbook
    Dim wsEP As Worksheet
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim importRange As Range
    Dim response As VbMsgBoxResult
    
    'Message de prévention
    reponse = MsgBox("Assurez-vous que le fichier Export_Produits respecte le format du modèle disponible dans le dossier Template, en suivant la structure définie sur la plage de colonnes [A:BI].", vbYesNo + vbQuestion, "Attention")
    If reponse = vbNo Then
        Exit Sub
    End If
    
    'Clear les données obsolètes
    Call Import_1_Main_Clear
    
    'Initialisation des classeurs
    Set wbWo = Workbooks("WareHouse_Optimizer.xlsm")
    Set wbProduct = Workbooks.Open(Path)
    
    'Initialisation des feuilles
    Set wsEP = wbWo.Sheets("D_EP")
    Set wsData = wbProduct.Sheets(1)
    
    'Vérifier la structure du classeur à importer
    Call Check_Format(wbProduct, wbWo, wsEP, wsData)
    
    'Définir les dernières lignes
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    'Définir la plage de donnée à importer
    Set importRange = wsData.Range("A2:AD" & lastRow)
    
    'Copier les données du classeur Export_Produits puis les coller dans le WO
    importRange.Copy
    wsEP.Range("A2").PasteSpecial Paste:=xlPasteValues

    'Fermer le classeur
    wbProduct.Close SaveChanges:=False
    
    'Convertir les données en FR
    Convert_FR
    
    'Message de Confirmation
    MsgBox "L'export produits à été importé avec succès !", vbInformation, "Succès"
    
End Sub

Private Sub Convert_FR()

    'Déclaration des Variables
    Dim wsEP As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim convertArray As Variant
    
    'Initialisation de la feuille
    Set wsEP = ThisWorkbook.Sheets("D_EP")
    
    'Définir la dernière ligne
    lastRow = wsEP.Cells(wsEP.Rows.Count, "A").End(xlUp).Row
    
    'Initialisation du tableau de données
    convertArray = wsEP.Range("K2:K" & lastRow).Value
    
    'Parcourir le tableau et le convertir en FR
    For rowIndex = 2 To UBound(convertArray, 1)
        convertArray(rowIndex, 1) = Replace(convertArray(rowIndex, 1), ".", ",")
    Next rowIndex
    
    'Saisir les données modifiées
    wsEP.Range("K2:K" & lastRow).Value = convertArray

End Sub

Private Sub Import_1_Main_Clear()

    'Déclaration des Variables
    Dim wsEP As Worksheet
    Dim lastRow As Long
    Dim clearRange As Range

    'Initialisation de la feuille
    Set wsEP = ThisWorkbook.Sheets("D_EP")
    
    'Définir la dernière ligne de la feuille Produit
    lastRow = wsEP.Cells(wsEP.Rows.Count, "A").End(xlUp).Row
    
    'Initialiser la plage à clear
    Set clearRange = wsEP.Range("A2:AD" & lastRow)
    
    'Supprimer les données de la plage
    If lastRow >= 2 Then
        clearRange.ClearContents
    End If
End Sub
