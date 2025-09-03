VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Settings_2_Select_Modify 
   Caption         =   "UserForm1"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10125
   OleObjectBlob   =   "Settings_2_Select_Modify.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Settings_2_Select_Modify"
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
    Settings_3_Add_Typo.Show
End Sub

Private Sub Image2_Click()
    Unload Me
    Settings_4_Remove_Typo.Show
End Sub

