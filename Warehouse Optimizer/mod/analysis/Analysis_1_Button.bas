Attribute VB_Name = "Analysis_1_Button"
Sub Start_Analysis()

    'D�claration des Variables
    Dim typeOfSupport As Variant
    Dim ABCPriority As Variant
    Dim reponse As VbMsgBoxResult
    
    'Message de confirmation
    reponse = MsgBox("Attention, cette action �crasera l'analyse pr�c�dente(hors implantation). Souhaitez vous poursuivre ?", vbQuestion + vbYesNo, "Confirmation")
    
    If reponse = vbNo Then
        Exit Sub
    End If
    
    'D�activer l'affichage
    Application.ScreenUpdating = False
    
    'Clear l'analyse pr�c�dente
    Set Clear = New Clear_2_Analysis
    With Clear
        On Error Resume Next
        .ClearABCCM
        On Error GoTo 0
        .ClearABC
        .ClearCalcul
    End With
    
    'D�finir les settings
    typeOfSupport = GetSettings("Type de support logistique")
    ABCPriority = GetSettings("Priorit�")
    
    'S�lection des proc�dures de calculs en fonction des Settings
    Select Case typeOfSupport
        Case "Rolls"
            Set Calcul = New Calculs_2_Rolls
            Calcul.ImportData
            Calcul.Calculs
            Calcul.Epiphenomenes
        Case "Palette 80x120"
            Set Calcul = New Calculs_3_PAL
            Calcul.ImportData
            Calcul.Calculs
            Calcul.Epiphenomenes
        End Select
        
    'S�lection des proc�dures ABC en fonction des Settings
    Select Case ABCPriority
        Case "Ventes"
            Set ABC = New ABC_2_Sales
            ABC.ImportABC
            ABC.CalculABC
        Case "Poids"
            Set ABC = New ABC_3_Weight
            ABC.ImportABC
            ABC.CalculABC
    End Select
    
    'S�lection des proc�dures ABCCM en fonction des Settings
    Select Case ABCPriority
        Case "Ventes"
            Set ABCCM = New ABCCM_2_Sales
            ABCCM.AddTCD
            ABCCM.CalculABCCM
        Case "Poids"
            Set ABCCM = New ABCCM_3_Weight
            ABCCM.AddTCD
            ABCCM.CalculABCCM
    End Select
    
    'Activer l'affichage
    Application.ScreenUpdating = False
    
    'Message de Confirmation
    MsgBox "L'analyse est termin�, vous pouvez vous rendre dans les onglets [Calcul Besoin], [ABC] , [ABC Code Mod�le] pour la consulter.", vbInformation, "Succ�s"

End Sub

