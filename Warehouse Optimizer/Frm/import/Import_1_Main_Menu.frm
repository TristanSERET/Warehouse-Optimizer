VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Import_1_Main_Menu 
   Caption         =   "Importer des données"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9240.001
   OleObjectBlob   =   "Import_1_Main_Menu.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Import_1_Main_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

    'Gestion du choix et du formulaire
    Unload Me
    Set Importeur = New Import_2_Forecast
    Importeur.folderPath

End Sub

Private Sub Image2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Image2_Click()

    'Gestion du choix et du formulaire
    Unload Me
    Set Importeur = New Import_3_History
    Importeur.folderPath

End Sub

Private Sub Image3_Click()

    'Gestion du choix et du formulaire
    Unload Me
    Set Importeur = New Import_4_Product
    Importeur.folderPath

End Sub

Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image1.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image2.SpecialEffect = fmSpecialEffectBump
End Sub
Private Sub Image3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image3.SpecialEffect = fmSpecialEffectBump
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image1.SpecialEffect = fmSpecialEffectFlat
    Image2.SpecialEffect = fmSpecialEffectFlat
    Image3.SpecialEffect = fmSpecialEffectFlat
End Sub
