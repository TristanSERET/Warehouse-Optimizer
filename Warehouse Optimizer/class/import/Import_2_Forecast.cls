VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Import_2_Forecast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class d'impportation des ForeCast

'Implémenter la classe primaire
Implements Import_1_Main

Private Sub Import_1_Main_folderPath()

    'Déclaration des Variables
    Dim wbWo As Workbook
    Dim wbForecast As Workbook
    Dim wsPV As Worksheet
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim importRange As Range
    Dim reponse As VbMsgBoxResult
    
    'Message de prévention
    reponse = MsgBox("Assurez-vous que le fichier des prévisions de ventes respecte le format du modèle disponible dans le dossier Template, en suivant la structure définie sur la plage de colonnes [A:BI].", vbYesNo + vbQuestion, "Attention")
    If reponse = vbNo Then
        Exit Sub
    End If
    
    'Clear les données obsolètes
    Call Import_1_Main_Clear
     
    'Initialisation des calsseurs
    Set wbWo = Workbooks("Warehouse_Optimizer.xlsm")
    Set wbForecast = Workbooks.Open(Path)
    
    'Initialisation des feuilles
    Set wsPV = wbWo.Sheets("D_PV")
    Set wsData = wbForecast.Sheets(1)
    
    'Vérifier la structure du classeur à importer
    Call Check_Format(wbForecast, wbWo, wsPV, wsData)
    
    'Définir la dernière ligne du classeur des ForeCast
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    'Définir la plage de données à importer
    Set importRange = wsData.Range("A2:BL" & lastRow)
    
    'Copier les données du classeur Forecast puis les coller dans le WO
    importRange.Copy
    wsPV.Range("A2").PasteSpecial Paste:=xlPasteValues
    
    'Fermer le classeur
    wbForecast.Close SaveChanges:=False
    
    'Message de Confirmation
    MsgBox "Les prévisions de ventes ont été importées avec succès !", vbInformation, "Succès"

End Sub

Private Sub Import_1_Main_Clear()

    'Déclaration des Variables
    Dim wsPV As Worksheet
    Dim lastRow As Long
    Dim clearRange As Range

    'Initialisation de la feuille
    Set wsPV = ThisWorkbook.Sheets("D_PV")
    
    'Définir la dernière ligne
    lastRow = wsPV.Cells(wsPV.Rows.Count, "A").End(xlUp).Row
    
    'Initialiser la plage à clear
    Set clearRange = wsPV.Range("A2:BI" & lastRow)
    
    'Supprimer les données de la plage
    If lastRow >= 2 Then
        clearRange.ClearContents
    End If
    
End Sub
