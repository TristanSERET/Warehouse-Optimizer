VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Import_3_History"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class d'Importation de l'historique des ventes

'Implémenter la classe primaire
Implements Import_1_Main

Private Sub Import_1_Main_folderPath()

    'Déclaration des Variables
    Dim wbWo As Workbook
    Dim wbHistory As Workbook
    Dim wsHV As Worksheet
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim importRange As Range
    Dim response As VbMsgBoxResult
    
    'Message de prévention
    reponse = MsgBox("Assurez-vous que le fichier de l'historique des ventes respecte le format du modèle disponible dans le dossier Template, en suivant la structure définie sur la plage de colonnes [A:BE].", vbYesNo + vbQuestion, "Attention")
    If reponse = vbNo Then
        Exit Sub
    End If
    
    'Clear les données obsolètes
    Call Import_1_Main_Clear
    
    'Initialisation des classeurs
    Set wbWo = Workbooks("Warehouse_Optimizer.xlsm")
    Set wbHistory = Workbooks.Open(Path)
    
    'Initialisation des feuilles
    Set wsHV = wbWo.Sheets("D_HV")
    Set wsData = wbHistory.Sheets("TCD")
    
    'Vérifier la structure du classeur à importer
    Call Check_Format(wbHistory, wbWo, wsHV, wsData)
    
    'Définir la dernière ligne du classeur de l'historique
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    'Définir la plage de données à importer
    Set importRange = wsData.Range("A10:BE" & lastRow)
    
    'Copier les données du classeur History puis les coller dans le WO
    importRange.Copy
    wsHV.Range("A2").PasteSpecial Paste:=xlPasteValues
    
    'Fermer le classeur
    wbHistory.Close SaveChanges:=False
    
    'Message de Confirmation
    MsgBox "L'historique des ventes à été importé avec succès !", vbInformation, "Succès"
    
End Sub

Private Sub Import_1_Main_Clear()
    
    'Déclaration des Variables
    Dim wsHV As Worksheet
    Dim lastRow As Long
    Dim clearRange As Range
    
    'Initialisation de la feuille
    Set wsHV = ThisWorkbook.Sheets("D_HV")
    
    'Définir la dernière ligne
    lastRow = wsHV.Cells(wsHV.Rows.Count, "A").End(xlUp).Row
    
    'Initialiser la plage à clear
    Set clearRange = wsHV.Range("A2:BE" & lastRow)
    
    'Supprimer les données de la plage
    If lastRow >= 2 Then
        clearRange.ClearContents
    End If
End Sub
