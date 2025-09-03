VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calculs_2_Rolls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class de Calul pour les Rolls

'Implémentation de la classe primaire
Implements Calculs_1_Main

Private Sub Calculs_1_Main_ImportData()
 
    'Importer les foreCast dans la feuille Calcul Besoin
    Call Forecast_Import
    
    'Importer l'historique des ventes dans la feuille Calcul Besoin
    Call History_Import
    
    'Vérifier si aucune erreur existe dans la colonne "H" "nombre de support par alvéole
    Call Check_NbSupport
    
End Sub

Private Sub Calculs_1_Main_Calculs()
    
    'Calculer l'écart des ventes
    Call Ecart_Calcul
    
    'Calculs de la solution Max
    Call Solution_Max
    
    'Calcul de la solution Average
    Call Solution_Average
    
    'Calcul de la solution Meilleure Moyenne
    Call Solution_BestAverage
 
End Sub

Private Sub Calculs_1_Main_Epiphenomenes()
    
    'Neutraliser les Epiphenomènes
    Call Solution_Epiphenomene
    
End Sub
    

