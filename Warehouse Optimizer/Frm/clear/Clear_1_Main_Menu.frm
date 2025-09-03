VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Clear_1_Main_Menu 
   Caption         =   "Supprimer le Cash"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   OleObjectBlob   =   "Clear_1_Main_Menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Clear_1_Main_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image1.SpecialEffect = fmSpecialEffectBump
End Sub

Private Sub Image2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image2.SpecialEffect = fmSpecialEffectBump
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Image1.SpecialEffect = fmSpecialEffectFlat
    Image2.SpecialEffect = fmSpecialEffectFlat
End Sub

Private Sub Image1_Click()

    Unload Me
    Clear_2_Select_Cell.Show
    
End Sub

Private Sub Image2_Click()
    
    'Déclaration des Variables
    Dim wsimplant As Worksheet
    Dim clearRange As Range
    Dim cell As Range
    Dim reponse As VbMsgBoxResult
    
    'Déactiver l'affichage lors du traitement
    Application.ScreenUpdating = False
        
    'Initialisation de la feuille
    Set wsimplant = ThisWorkbook.Sheets("Implantation")
    
    'Initialisation de la plage du user
    Set clearRange = Selection
    
    'Message de validation
    reponse = MsgBox("Votre sélection actuelle est " & clearRange.Address & " Souhaitez vous continuer ?", vbYesNo + vbQuestion, "Confirmation")
    
    'Si la réponse de l'user et positive poursuivre
    If reponse = vbYes Then
        'Parcourir la plage de données et supprimer les couleurs + data si les conditions sont remplis
        For Each cell In clearRange
            If cell.Interior.Color <> RGB(217, 217, 217) Then
                cell.Interior.Color = vbWhite
                cell.ClearContents
            End If
        Next cell
        
        'Message de confirmation
        MsgBox "Le cash de votre sélection " & clearRange.Address & " à été supprimé", vbInformation, "Succès"
    End If
    
    'Activer l'affichage
    Application.ScreenUpdating = True
    
    'Fermer le formulaire
    Unload Me
    
End Sub
