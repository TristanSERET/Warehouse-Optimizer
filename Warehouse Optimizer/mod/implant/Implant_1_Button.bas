Attribute VB_Name = "Implant_1_Button"
Sub Execute_Implantation()

    'Déactiver l'affichage lors du traitement
    Application.ScreenUpdating = False
    
    'Déclaration des Variables
    Dim typeSelected As String
    Dim rangee As Variant
    
    'Vérifier qu'une rangée de départ et sélectionnée
    rangee = GetSettings("Rangée de départ")
    If rangee = "" Then
        MsgBox "Veuillez sélectionner dans les parmètres une rangée de départ !", vbExclamation, "Attention"
        Exit Sub
    End If
    
    'Définir l'option sélectionné
    typeSelected = GetSettings("Type d'implantation")
    
    'Executer l'algorithme en fonction des options
    
    Select Case typeSelected
        Case "Suivant l'ABC par référence"
            Set Implant = New Implant_2_REF
            Implant.ImplantRef
            On Error Resume Next
            Implant.GenerateColor
            On Error GoTo 0
        Case "Suivant l'ABC par CodeModele"
            Set Implant = New Implant_3_Codmod
            Implant.ImplantRef
            On Error Resume Next
            Implant.GenerateColor
            On Error GoTo 0
    End Select
    
    'Activer l'affichage
    Application.ScreenUpdating = True
    
End Sub
