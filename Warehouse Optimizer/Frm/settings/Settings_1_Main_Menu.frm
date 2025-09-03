VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Settings_1_Main_Menu 
   Caption         =   "Paramètres"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11355
   OleObjectBlob   =   "Settings_1_Main_Menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Settings_1_Main_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image1_Click()

    'Saisir les données sélectionnées
    Write_Data
    
    'Appliquer les modifications
    Apply_Settings
    
    'Message de Confirmation
    MsgBox "Paramètres enregistré", vbInformation, "Succès"
    
    'Fermer le formulaire
    Unload Me

End Sub

Private Sub Apply_Settings()
    
    'Déclaration des Variables
    Dim wsCalcul As Worksheet
    Dim wsABC As Worksheet
    
    'Initialisation des feuilles
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    Set wsABC = ThisWorkbook.Sheets("ABC")
    
    'Modifier les lib en fonction du setting
    Select Case ComboBox2.Value
        Case "Rolls"
            wsCalcul.Range("I2").Value = "qté/Rolls"
            wsCalcul.Range("H2").Value = "nbRolls_Alvéole"
            wsCalcul.Range("BP2").Value = "Besoin Pick Rolls"
            wsCalcul.Range("BT2").Value = "Besoin Pick Rolls"
            wsCalcul.Range("BZ2").Value = "Besoin Pick Rolls"
            wsABC.Range("H2").Value = "Besoin Rolls"
        Case "Palette 80x120"
            wsCalcul.Range("I2").Value = "qté/Pal"
            wsCalcul.Range("H2").Value = "EMP_Requis"
            wsCalcul.Range("BP2").Value = "Besoin Pick PAL"
            wsCalcul.Range("BT2").Value = "Besoin Pick PAL"
            wsCalcul.Range("BZ2").Value = "Besoin Pick PAL"
            wsABC.Range("H2").Value = "Besoin Palette"
    End Select
    
    'Modifier la priorité de l'ABC
    Select Case ComboBox3.Value
        Case "Poids"
            wsABC.Range("E2").Value = "Poids"
            wsABC.Range("F2").Value = "% du Poids"
        Case "Ventes"
            wsABC.Range("E2").Value = "Ventes"
            wsABC.Range("F2").Value = "% des Ventes"
    End Select
    
End Sub

Private Sub Write_Data()

    'Déclaration des Variables
    Dim wsSettings As Worksheet
    Dim lastRow As Byte
    Dim rowIndex As Byte
    
    'Initialiser la feuille Settings
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    
    'Définir la dernière ligne
    lastRow = wsSettings.Cells(wsSettings.Rows.Count, "B").End(xlUp).Row
    
    'Parcourrir les paramètres et saisir les modifications
    For rowIndex = 2 To lastRow
        Select Case wsSettings.Cells(rowIndex, "B").Value
            Case "Type de support logistique": wsSettings.Cells(rowIndex, "C").Value = ComboBox2.Value
            Case "% Mise à disposition": wsSettings.Cells(rowIndex, "C").Value = Format(TextBox1.Value, "0%")
            Case "Typologie": wsSettings.Cells(rowIndex, "C").Value = ComboBox1.Value
            Case "Limite de semaine (Meilleure Moyenne)": wsSettings.Cells(rowIndex, "C").Value = TextBox5.Value
            Case "Sensibilité des épiphénomènes": wsSettings.Cells(rowIndex, "C").Value = TextBox6.Value
            Case "Priorité":  wsSettings.Cells(rowIndex, "C").Value = ComboBox3.Value
            Case "Calcul retenu en sortie":  wsSettings.Cells(rowIndex, "C").Value = ComboBox4.Value
            Case "Sensibilité de la Classe A":  wsSettings.Cells(rowIndex, "C").Value = TextBox2.Value
            Case "Sensibilité de la Classe B":  wsSettings.Cells(rowIndex, "C").Value = TextBox3.Value
            Case "Sensibilité de la Classe C":  wsSettings.Cells(rowIndex, "C").Value = TextBox4.Value
            Case "Préférence du trie ABC au Code Modèle": wsSettings.Cells(rowIndex, "C").Value = ComboBox5.Value
            Case "Cellule d'implantation": wsSettings.Cells(rowIndex, "C").Value = ComboBox6.Value
            Case "Sens d'implantation": wsSettings.Cells(rowIndex, "C").Value = ComboBox7.Value
            Case "Type d'implantation": wsSettings.Cells(rowIndex, "C").Value = ComboBox8.Value
            Case "Autorisation d'implantation Classe A": wsSettings.Cells(rowIndex, "C").Value = ComboBox9.Value
            Case "Autorisation d'implantation Classe B": wsSettings.Cells(rowIndex, "C").Value = ComboBox10.Value
            Case "Autorisation d'implantation Classe C": wsSettings.Cells(rowIndex, "C").Value = ComboBox11.Value
            Case "Rangée de départ": wsSettings.Cells(rowIndex, "C").Value = ComboBox12.Value
            Case "Affectation du Picking Dynamique": wsSettings.Cells(rowIndex, "C").Value = ComboBox13.Value
            Case "Positionnement du Picking Dynamique": wsSettings.Cells(rowIndex, "C").Value = ComboBox14.Value
            Case "Nombre d'alvéoles à allouer": wsSettings.Cells(rowIndex, "C").Value = TextBox7.Value
        End Select
    Next rowIndex
End Sub


Private Sub UserForm_Initialize()

    'Déclaration des Variables
    Dim wsSettings As Worksheet
    Dim wsTypo As Worksheet
    Dim lastRow As Byte
    Dim lastRowTypo As Byte
    Dim rowIndex As Byte
    Dim minRangee As Byte
    Dim maxRangee As Byte
    Dim rangee As Byte
    
    'Initialisation de la feuille Settings
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    Set wsTypo = ThisWorkbook.Sheets("Set_Typo")
    
    'Définir la dernière ligne
    lastRow = wsSettings.Cells(wsSettings.Rows.Count, "B").End(xlUp).Row
    lastRowTypo = wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row
    
    'Initialisation des Paramètres
    For rowIndex = 3 To lastRow
        Select Case wsSettings.Cells(rowIndex, "B").Value
            Case "Type de support logistique"
                With ComboBox2
                    .AddItem "Rolls"
                    .AddItem "Palette 80x120"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "% Mise à disposition"
                TextBox1.Value = Format(wsSettings.Cells(rowIndex, "C").Value, "0%")
            Case "Typologie"
                ComboBox1.Value = wsSettings.Cells(rowIndex, "C").Value
            Case "Limite de semaine (Meilleure Moyenne)"
                TextBox5.Value = wsSettings.Cells(rowIndex, "C").Value
            Case "Sensibilité des épiphénomènes"
                TextBox6.Value = wsSettings.Cells(rowIndex, "C").Value
            Case "Priorité"
                With ComboBox3
                    .AddItem "Poids"
                    .AddItem "Ventes"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Calcul retenu en sortie"
                With ComboBox4
                    .AddItem "Meilleure Moyenne"
                    .AddItem "Max"
                    .AddItem "Moyenne"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Préférence du trie ABC au Code Modèle"
                With ComboBox5
                    .AddItem "Somme des Alvéoles"
                    .AddItem "Somme des Ventes"
                    .AddItem "Somme des Poids"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Cellule d'implantation"
                With ComboBox6
                    .AddItem "Cellule_A"
                    .AddItem "Cellule_B"
                    .AddItem "Cellule_E"
                    .AddItem "Cellule_F"
                    .AddItem "Cellule_G"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Sens d'implantation"
                With ComboBox7
                    .AddItem "Gauche à Droite"
                    .AddItem "Droite à Gauche"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Type d'implantation"
                With ComboBox8
                    .AddItem "Suivant l'ABC par référence"
                    .AddItem "Suivant l'ABC par CodeModele"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Autorisation d'implantation Classe A"
                With ComboBox9
                    .AddItem "Avant passage chariot uniquement"
                    .AddItem "Après passage chariot uniquement"
                    .AddItem "Tout"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Autorisation d'implantation Classe B"
                With ComboBox10
                    .AddItem "Avant passage chariot uniquement"
                    .AddItem "Après passage chariot uniquement"
                    .AddItem "Tout"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Autorisation d'implantation Classe C"
                With ComboBox11
                    .AddItem "Avant passage chariot uniquement"
                    .AddItem "Après passage chariot uniquement"
                    .AddItem "Tout"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Rangée de Départ"
                Call dependanceRangee(minRangee, maxRangee)
                For rangee = minRangee To maxRangee
                    ComboBox12.AddItem rangee
                Next rangee
                    ComboBox12.Value = wsSettings.Cells(rowIndex, "C").Value
            Case "Affectation du Picking Dynamique"
                With ComboBox13
                    .AddItem "Automatique"
                    .AddItem "Manuelle"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Positionnement du Picking Dynamique"
                With ComboBox14
                    .AddItem "Avant passage chariot uniquement"
                    .AddItem "Après passage chariot uniquement"
                    .AddItem "Tout"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Sensibilité de la Classe A": TextBox2.Value = Format(wsSettings.Cells(rowIndex, "C").Value, "0%")
            Case "Sensibilité de la Classe B": TextBox3.Value = Format(wsSettings.Cells(rowIndex, "C").Value, "0%")
            Case "Sensibilité de la Classe C": TextBox4.Value = Format(wsSettings.Cells(rowIndex, "C").Value, "0%")
        End Select
    Next rowIndex
    
End Sub


Private Sub dependanceRangee(ByRef minRangee As Byte, ByRef maxRangee As Byte)

    'Déclaration des Variables
    Dim Cellule As String
    
    'Vérifier la valeur dans la combobox6
    Cellule = ComboBox6.Value

    'Sélectionner la plage de rangée correspondante
    Select Case Cellule
        Case "Cellule_A": minRangee = 1: maxRangee = 16
        Case "Cellule_B": minRangee = 17: maxRangee = 32
        Case "Cellule_E": minRangee = 35: maxRangee = 50
        Case "Cellule_F": minRangee = 1: maxRangee = 16
        Case "Cellule_G": minRangee = 17: maxRangee = 32
    End Select
    
End Sub

Private Sub ComboBox6_Change()
    
    'Déclaration des Variables
    Dim minRangee As Byte
    Dim maxRangee As Byte
    Dim rangee As Byte
    
    'Clear la liste de rangée de départ
    ComboBox12.Clear
    
    'Réajuster l'option de rangée en fonction de la cellule sélectionnée
    Call dependanceRangee(minRangee, maxRangee)
    
    For rangee = minRangee To maxRangee
        ComboBox12.AddItem rangee
    Next rangee

End Sub

Private Sub ComboBox13_Change()

    'En fonction de l'état de la ComboBox changer le statut d'accès au paramètre manuelle
    If ComboBox13.Value = "Automatique" Then
        TextBox7.Enabled = False
        TextBox7.BackColor = RGB(160, 160, 160)
    Else
        TextBox7.Enabled = True
        TextBox7.BackColor = RGB(255, 255, 255)
    End If

End Sub

Private Sub CheckBox1_Change()

    'Déclaration des Variables
    Dim wsTypo As Worksheet
    Dim lastRow As Byte
    Dim rowIndex As Byte
    
    'Initialisation de la feuille Typo
    Set wsTypo = ThisWorkbook.Sheets("Set_Typo")
    
    'Définir la dernière ligne
    If wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row > wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row Then
        lastRow = wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row
    Else
        lastRow = wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row
    End If
    
    'Si cette CheckBox est active déactiver la deuxième (HG2)
    If CheckBox1 = True Then
        CheckBox2 = False
        
        'Charger les élements HG1
        ComboBox1.Clear
        For rowIndex = 2 To lastRow
            ComboBox1.AddItem wsTypo.Cells(rowIndex, "A").Value
        Next rowIndex
    End If
    
End Sub

Private Sub CheckBox2_Change()
    
    'Déclaration des Variables
    Dim wsTypo As Worksheet
    Dim lastRow As Byte
    Dim rowIndex As Byte
    
    'Initialisation de la feuille Typo
    Set wsTypo = ThisWorkbook.Sheets("Set_Typo")
    
    'Définir la dernière ligne
    If wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row > wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row Then
        lastRow = wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row
    Else
        lastRow = wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row
    End If
    
    'Si cette CheckBox est active déactiver la deuxième (HG2)
    If CheckBox2 = True Then
        CheckBox1 = False
        
        'Charger les élements HG2
        ComboBox1.Clear
        For rowIndex = 2 To lastRow
            ComboBox1.AddItem wsTypo.Cells(rowIndex, "B").Value
        Next rowIndex
    End If
    
End Sub

Private Sub CommandButton1_Click()

    Settings_2_Select_Modify.Show

End Sub

