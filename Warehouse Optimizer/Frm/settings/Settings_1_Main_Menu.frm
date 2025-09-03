VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Settings_1_Main_Menu 
   Caption         =   "Param�tres"
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

    'Saisir les donn�es s�lectionn�es
    Write_Data
    
    'Appliquer les modifications
    Apply_Settings
    
    'Message de Confirmation
    MsgBox "Param�tres enregistr�", vbInformation, "Succ�s"
    
    'Fermer le formulaire
    Unload Me

End Sub

Private Sub Apply_Settings()
    
    'D�claration des Variables
    Dim wsCalcul As Worksheet
    Dim wsABC As Worksheet
    
    'Initialisation des feuilles
    Set wsCalcul = ThisWorkbook.Sheets("Calcul Besoin")
    Set wsABC = ThisWorkbook.Sheets("ABC")
    
    'Modifier les lib en fonction du setting
    Select Case ComboBox2.Value
        Case "Rolls"
            wsCalcul.Range("I2").Value = "qt�/Rolls"
            wsCalcul.Range("H2").Value = "nbRolls_Alv�ole"
            wsCalcul.Range("BP2").Value = "Besoin Pick Rolls"
            wsCalcul.Range("BT2").Value = "Besoin Pick Rolls"
            wsCalcul.Range("BZ2").Value = "Besoin Pick Rolls"
            wsABC.Range("H2").Value = "Besoin Rolls"
        Case "Palette 80x120"
            wsCalcul.Range("I2").Value = "qt�/Pal"
            wsCalcul.Range("H2").Value = "EMP_Requis"
            wsCalcul.Range("BP2").Value = "Besoin Pick PAL"
            wsCalcul.Range("BT2").Value = "Besoin Pick PAL"
            wsCalcul.Range("BZ2").Value = "Besoin Pick PAL"
            wsABC.Range("H2").Value = "Besoin Palette"
    End Select
    
    'Modifier la priorit� de l'ABC
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

    'D�claration des Variables
    Dim wsSettings As Worksheet
    Dim lastRow As Byte
    Dim rowIndex As Byte
    
    'Initialiser la feuille Settings
    Set wsSettings = ThisWorkbook.Sheets("Settings")
    
    'D�finir la derni�re ligne
    lastRow = wsSettings.Cells(wsSettings.Rows.Count, "B").End(xlUp).Row
    
    'Parcourrir les param�tres et saisir les modifications
    For rowIndex = 2 To lastRow
        Select Case wsSettings.Cells(rowIndex, "B").Value
            Case "Type de support logistique": wsSettings.Cells(rowIndex, "C").Value = ComboBox2.Value
            Case "% Mise � disposition": wsSettings.Cells(rowIndex, "C").Value = Format(TextBox1.Value, "0%")
            Case "Typologie": wsSettings.Cells(rowIndex, "C").Value = ComboBox1.Value
            Case "Limite de semaine (Meilleure Moyenne)": wsSettings.Cells(rowIndex, "C").Value = TextBox5.Value
            Case "Sensibilit� des �piph�nom�nes": wsSettings.Cells(rowIndex, "C").Value = TextBox6.Value
            Case "Priorit�":  wsSettings.Cells(rowIndex, "C").Value = ComboBox3.Value
            Case "Calcul retenu en sortie":  wsSettings.Cells(rowIndex, "C").Value = ComboBox4.Value
            Case "Sensibilit� de la Classe A":  wsSettings.Cells(rowIndex, "C").Value = TextBox2.Value
            Case "Sensibilit� de la Classe B":  wsSettings.Cells(rowIndex, "C").Value = TextBox3.Value
            Case "Sensibilit� de la Classe C":  wsSettings.Cells(rowIndex, "C").Value = TextBox4.Value
            Case "Pr�f�rence du trie ABC au Code Mod�le": wsSettings.Cells(rowIndex, "C").Value = ComboBox5.Value
            Case "Cellule d'implantation": wsSettings.Cells(rowIndex, "C").Value = ComboBox6.Value
            Case "Sens d'implantation": wsSettings.Cells(rowIndex, "C").Value = ComboBox7.Value
            Case "Type d'implantation": wsSettings.Cells(rowIndex, "C").Value = ComboBox8.Value
            Case "Autorisation d'implantation Classe A": wsSettings.Cells(rowIndex, "C").Value = ComboBox9.Value
            Case "Autorisation d'implantation Classe B": wsSettings.Cells(rowIndex, "C").Value = ComboBox10.Value
            Case "Autorisation d'implantation Classe C": wsSettings.Cells(rowIndex, "C").Value = ComboBox11.Value
            Case "Rang�e de d�part": wsSettings.Cells(rowIndex, "C").Value = ComboBox12.Value
            Case "Affectation du Picking Dynamique": wsSettings.Cells(rowIndex, "C").Value = ComboBox13.Value
            Case "Positionnement du Picking Dynamique": wsSettings.Cells(rowIndex, "C").Value = ComboBox14.Value
            Case "Nombre d'alv�oles � allouer": wsSettings.Cells(rowIndex, "C").Value = TextBox7.Value
        End Select
    Next rowIndex
End Sub


Private Sub UserForm_Initialize()

    'D�claration des Variables
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
    
    'D�finir la derni�re ligne
    lastRow = wsSettings.Cells(wsSettings.Rows.Count, "B").End(xlUp).Row
    lastRowTypo = wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row
    
    'Initialisation des Param�tres
    For rowIndex = 3 To lastRow
        Select Case wsSettings.Cells(rowIndex, "B").Value
            Case "Type de support logistique"
                With ComboBox2
                    .AddItem "Rolls"
                    .AddItem "Palette 80x120"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "% Mise � disposition"
                TextBox1.Value = Format(wsSettings.Cells(rowIndex, "C").Value, "0%")
            Case "Typologie"
                ComboBox1.Value = wsSettings.Cells(rowIndex, "C").Value
            Case "Limite de semaine (Meilleure Moyenne)"
                TextBox5.Value = wsSettings.Cells(rowIndex, "C").Value
            Case "Sensibilit� des �piph�nom�nes"
                TextBox6.Value = wsSettings.Cells(rowIndex, "C").Value
            Case "Priorit�"
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
            Case "Pr�f�rence du trie ABC au Code Mod�le"
                With ComboBox5
                    .AddItem "Somme des Alv�oles"
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
                    .AddItem "Gauche � Droite"
                    .AddItem "Droite � Gauche"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Type d'implantation"
                With ComboBox8
                    .AddItem "Suivant l'ABC par r�f�rence"
                    .AddItem "Suivant l'ABC par CodeModele"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Autorisation d'implantation Classe A"
                With ComboBox9
                    .AddItem "Avant passage chariot uniquement"
                    .AddItem "Apr�s passage chariot uniquement"
                    .AddItem "Tout"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Autorisation d'implantation Classe B"
                With ComboBox10
                    .AddItem "Avant passage chariot uniquement"
                    .AddItem "Apr�s passage chariot uniquement"
                    .AddItem "Tout"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Autorisation d'implantation Classe C"
                With ComboBox11
                    .AddItem "Avant passage chariot uniquement"
                    .AddItem "Apr�s passage chariot uniquement"
                    .AddItem "Tout"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Rang�e de D�part"
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
                    .AddItem "Apr�s passage chariot uniquement"
                    .AddItem "Tout"
                    .Value = wsSettings.Cells(rowIndex, "C").Value
                End With
            Case "Sensibilit� de la Classe A": TextBox2.Value = Format(wsSettings.Cells(rowIndex, "C").Value, "0%")
            Case "Sensibilit� de la Classe B": TextBox3.Value = Format(wsSettings.Cells(rowIndex, "C").Value, "0%")
            Case "Sensibilit� de la Classe C": TextBox4.Value = Format(wsSettings.Cells(rowIndex, "C").Value, "0%")
        End Select
    Next rowIndex
    
End Sub


Private Sub dependanceRangee(ByRef minRangee As Byte, ByRef maxRangee As Byte)

    'D�claration des Variables
    Dim Cellule As String
    
    'V�rifier la valeur dans la combobox6
    Cellule = ComboBox6.Value

    'S�lectionner la plage de rang�e correspondante
    Select Case Cellule
        Case "Cellule_A": minRangee = 1: maxRangee = 16
        Case "Cellule_B": minRangee = 17: maxRangee = 32
        Case "Cellule_E": minRangee = 35: maxRangee = 50
        Case "Cellule_F": minRangee = 1: maxRangee = 16
        Case "Cellule_G": minRangee = 17: maxRangee = 32
    End Select
    
End Sub

Private Sub ComboBox6_Change()
    
    'D�claration des Variables
    Dim minRangee As Byte
    Dim maxRangee As Byte
    Dim rangee As Byte
    
    'Clear la liste de rang�e de d�part
    ComboBox12.Clear
    
    'R�ajuster l'option de rang�e en fonction de la cellule s�lectionn�e
    Call dependanceRangee(minRangee, maxRangee)
    
    For rangee = minRangee To maxRangee
        ComboBox12.AddItem rangee
    Next rangee

End Sub

Private Sub ComboBox13_Change()

    'En fonction de l'�tat de la ComboBox changer le statut d'acc�s au param�tre manuelle
    If ComboBox13.Value = "Automatique" Then
        TextBox7.Enabled = False
        TextBox7.BackColor = RGB(160, 160, 160)
    Else
        TextBox7.Enabled = True
        TextBox7.BackColor = RGB(255, 255, 255)
    End If

End Sub

Private Sub CheckBox1_Change()

    'D�claration des Variables
    Dim wsTypo As Worksheet
    Dim lastRow As Byte
    Dim rowIndex As Byte
    
    'Initialisation de la feuille Typo
    Set wsTypo = ThisWorkbook.Sheets("Set_Typo")
    
    'D�finir la derni�re ligne
    If wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row > wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row Then
        lastRow = wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row
    Else
        lastRow = wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row
    End If
    
    'Si cette CheckBox est active d�activer la deuxi�me (HG2)
    If CheckBox1 = True Then
        CheckBox2 = False
        
        'Charger les �lements HG1
        ComboBox1.Clear
        For rowIndex = 2 To lastRow
            ComboBox1.AddItem wsTypo.Cells(rowIndex, "A").Value
        Next rowIndex
    End If
    
End Sub

Private Sub CheckBox2_Change()
    
    'D�claration des Variables
    Dim wsTypo As Worksheet
    Dim lastRow As Byte
    Dim rowIndex As Byte
    
    'Initialisation de la feuille Typo
    Set wsTypo = ThisWorkbook.Sheets("Set_Typo")
    
    'D�finir la derni�re ligne
    If wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row > wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row Then
        lastRow = wsTypo.Cells(wsTypo.Rows.Count, "A").End(xlUp).Row
    Else
        lastRow = wsTypo.Cells(wsTypo.Rows.Count, "B").End(xlUp).Row
    End If
    
    'Si cette CheckBox est active d�activer la deuxi�me (HG2)
    If CheckBox2 = True Then
        CheckBox1 = False
        
        'Charger les �lements HG2
        ComboBox1.Clear
        For rowIndex = 2 To lastRow
            ComboBox1.AddItem wsTypo.Cells(rowIndex, "B").Value
        Next rowIndex
    End If
    
End Sub

Private Sub CommandButton1_Click()

    Settings_2_Select_Modify.Show

End Sub

