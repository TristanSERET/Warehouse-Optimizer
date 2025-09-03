VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Help_1_Main 
   Caption         =   "Cellule d'aide"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13635
   OleObjectBlob   =   "Help_1_Main.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Help_1_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    'Déclaration des Variables
    Dim trigger As String
    
    'Quel est le bouton déclanché
    trigger = Application.Caller
    
    'Vérifier quel bouton est déclanché
    Select Case trigger
        Case "Help_Implant": Help_1_Main.MultiPage1.Value = 2
        Case "Help": Help_1_Main.MultiPage1.Value = 0
    End Select

End Sub
