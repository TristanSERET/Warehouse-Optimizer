Attribute VB_Name = "Analysis_1_Button"
Sub Start_Analysis()

    'Déclaration des Variables
    Dim typeOfSupport As Variant
    Dim ABCPriority As Variant
    Dim reponse As VbMsgBoxResult
    
    'Message de confirmation
    reponse = MsgBox("Attention, cette action écrasera l'analyse précédente(hors implantation). Souhaitez vous poursuivre ?", vbQuestion + vbYesNo, "Confirmation")
    
    If reponse = vbNo Then
        Exit Sub
    End If
    
    'Déactiver l'affichage
    Application.ScreenUpdating = False
    
    'Clear l'analyse précédente
    Set Clear = New Clear_2_Analysis
    With Clear
        On Error Resume Next
        .ClearABCCM
        On Error GoTo 0
        .ClearABC
        .ClearCalcul
    End With
    
    'Définir les settings
    typeOfSupport = GetSettings("Type de support logistique")
    ABCPriority = GetSettings("Priorité")
    
    'Sélection des procèdures de calculs en fonction des Settings
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
        
    'Sélection des procèdures ABC en fonction des Settings
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
    
    'Sélection des procèdures ABCCM en fonction des Settings
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
    MsgBox "L'analyse est terminé, vous pouvez vous rendre dans les onglets [Calcul Besoin], [ABC] , [ABC Code Modèle] pour la consulter.", vbInformation, "Succès"

End Sub

