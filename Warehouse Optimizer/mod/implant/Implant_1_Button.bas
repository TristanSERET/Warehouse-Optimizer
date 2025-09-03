Attribute VB_Name = "Implant_1_Button"
Sub Execute_Implantation()

    'D�activer l'affichage lors du traitement
    Application.ScreenUpdating = False
    
    'D�claration des Variables
    Dim typeSelected As String
    Dim rangee As Variant
    
    'V�rifier qu'une rang�e de d�part et s�lectionn�e
    rangee = GetSettings("Rang�e de d�part")
    If rangee = "" Then
        MsgBox "Veuillez s�lectionner dans les parm�tres une rang�e de d�part !", vbExclamation, "Attention"
        Exit Sub
    End If
    
    'D�finir l'option s�lectionn�
    typeSelected = GetSettings("Type d'implantation")
    
    'Executer l'algorithme en fonction des options
    
    Select Case typeSelected
        Case "Suivant l'ABC par r�f�rence"
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
