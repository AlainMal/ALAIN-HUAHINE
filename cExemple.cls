VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExemple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' ========== Procédure Affiche(numéro de la ligne) ==========
'
'Ce programme affiche les contenus des différentes valeurs de la feuille "Exemple"
'c'est une sous routine de la fonction PGN,
'en fonction de la case à cocher "Affichage seul"
'La procédure RAZ sert à effacer l'affichage des cellules
'
'============================================================

'RAZ de l'Exemple
Public Sub RAZ()
Application.Calculation = xlCalculationManual
Worksheets(FEUIL_EXEMPLE).Unprotect
           
            Worksheets(FEUIL_EXEMPLE).Range(VITESSE_SURFACE).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(VITESSE_FOND).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(VENT_VRAI_ORIENTATION).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(VITESSE_VENT_VRAI).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(CALCUL_ATTITUDE_VRAI).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(VENT_APPARENT_ORIENTATION).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(VITESSE_VENT_APPARENT).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(CALCUL_ATTITUDE_APPARENT).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(TAUX_CHARGE_SERVICE).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(TEMPARATURE_EAU_GPS2).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(TAUX_CHARGE_MOTEUR).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(VOLT_SERVICE).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(VOLT_MOTEUR).Value = ""
            'Worksheets(FEUIL_EXEMPLE).Range("F15").Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(DIRECTION_BATEAU).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(DONNEE_SOG).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(COG).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(SOG).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(POSITION_LATITUDE).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(POSITION_LONGITUDE).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(PROFONDEUR).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(DONNEE_HEADING).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(PRESSION).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(ORIENTATION_BARRE).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(NIVEAU_EAU).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(QUANTITE_TOTAL_EAU).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(CALCUL_RESTANT_EAU).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(NIVEAU_GASOIL).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(NIVEAU_GASOIL).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(QUANTITE_GASOIL).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(CALCUL_RESTANT_GASOIL).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(TEMPERATURE_CABINE).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(TEMPERATURE_MOTEUR).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(TEMPERATURE_EAU_GPS).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(TEMPARATURE_EAU_GPS2).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(TEMPERATURE_EAU_SPEED).Value = ""
            Worksheets(FEUIL_EXEMPLE).Range(TEMPERATURE__BATTERIE_SERVICE).Value = ""

Application.Calculation = xlCalculationAutomatic
Worksheets(FEUIL_EXEMPLE).Protect

End Sub

'============== Fonction Affiche ===============
'Affiche les valeurs sur la feuille "Exemple"
' En fonction de la ligne
'===============================================
Public Sub Affiche(ligne As Long)

'En fonction de la valeur du PGN issu du bus CAN
Select Case cID.Pgn(ligne)
    Case 128259 'Vitesse du bateau
        ii = ii + 1
        If ValeurChoisie <> "" Then
            Worksheets(FEUIL_EXEMPLE).Range(VITESSE_SURFACE).Value = Format(ValeurChoisie, "0.00")
        End If
        If ValeurChoisie2 <> "" Then
            Worksheets(FEUIL_EXEMPLE).Range(VITESSE_FOND).Value = Format(ValeurChoisie2, "0.00")
        End If

    Case 130578 'Vitesse du bateau donne la vitesse surface et fond sur la même trame (numéro 0)
        ii = ii + 1
        If ValeurChoisie2 <> "" And ValeurChoisie = 0 Then
            Worksheets(FEUIL_EXEMPLE).Range(VITESSE_SURFACE).Value = Format(ValeurChoisie2, "0.00")
        End If
        If ValeurChoisie3 <> "" And ValeurChoisie = 0 Then
            Worksheets(FEUIL_EXEMPLE).Range(VITESSE_FOND).Value = Format(ValeurChoisie3, "0.00")
        End If
    
    Case 130306 'Vent vrai et apparent
        ii = ii + 1
        If ValeurChoisie <> "" Then
            If ValeurChoisieTab = 2 Then
                Worksheets(FEUIL_EXEMPLE).Range(VENT_APPARENT_ORIENTATION).Value = Format(ValeurChoisie, "0.00")
            Else
                Worksheets(FEUIL_EXEMPLE).Range(VENT_VRAI_ORIENTATION).Value = Format(ValeurChoisie, "0.00")
            End If
        End If
        If ValeurChoisie2 <> "" Then
            If ValeurChoisieTab <> 2 Then
                Worksheets(FEUIL_EXEMPLE).Range(VITESSE_VENT_VRAI).Value = Format(ValeurChoisie2, "0.00")
            Else
                Worksheets(FEUIL_EXEMPLE).Range(VITESSE_VENT_APPARENT).Value = Format(ValeurChoisie2, "0.00")
            End If
            'CALCUL de l'orientation Tb et Bd
            If Format(ValeurChoisie2, "0.00") > 180 Then
                If ValeurChoisieTab <> 2 Then
                    Worksheets(FEUIL_EXEMPLE).Range(CALCUL_ATTITUDE_VRAI).Value = Format(360 - Format(Abs(ValeurChoisie2), "0.00"), "0") & "° Bd"
                Else
                    Worksheets(FEUIL_EXEMPLE).Range(CALCUL_ATTITUDE_APPARENT).Value = Format(360 - Format(Abs(ValeurChoisie2), "0.00"), "0") & "° Bd"
                End If
            Else
                If ValeurChoisieTab <> 2 Then
                    Worksheets(FEUIL_EXEMPLE).Range(CALCUL_ATTITUDE_VRAI).Value = Format(Abs(ValeurChoisie2), "0") & "° Tb"
                Else
                    Worksheets(FEUIL_EXEMPLE).Range(CALCUL_ATTITUDE_APPARENT).Value = Abs(Format(Abs(ValeurChoisie2), "0")) & "° Tb"
                End If
            End If
        End If
    
    Case 127506 'Etat détaillé de charge des Batterie Service et Moteur
        ii = ii + 1
        If ValeurChoisie2 <> "" And ValeurChoisie = "0" And ValeurChoisieTab = "0" Then
            Worksheets(FEUIL_EXEMPLE).Range(TAUX_CHARGE_SERVICE).Value = ValeurChoisie2
        ElseIf ValeurChoisie2 <> "" And ValeurChoisie = "0" And ValeurChoisieTab = "1" Then
            Worksheets(FEUIL_EXEMPLE).Range(TAUX_CHARGE_MOTEUR).Value = ValeurChoisie2
        End If
    
    Case 127508 'Etat Batteries Services et Moteurs
        ii = ii + 1
        If ValeurChoisie <> "" And ValeurChoisieTab = "255" Then
            Worksheets(FEUIL_EXEMPLE).Range(VOLT_SERVICE).Value = Format(ValeurChoisie, "0.00")
        Else
            Worksheets(FEUIL_EXEMPLE).Range(VOLT_MOTEUR).Value = Format(ValeurChoisie, "0.00")
        End If
        If ValeurChoisie3 <> "" And ValeurChoisieTab = "255" Then
            Worksheets(FEUIL_EXEMPLE).Range(TEMPERATURE__BATTERIE_SERVICE).Value = Format(ValeurChoisie3, "0.00")
        Else
            Worksheets(FEUIL_EXEMPLE).Range("F15").Value = Format(ValeurChoisie3, "0.00")
        End If
     
     Case 130577 'Données de direction du navire
        ii = ii + 1
        If ValeurChoisie = 1 Then 'Sur la trame 1
            If ValeurChoisie2 <> "" Then
                Worksheets(FEUIL_EXEMPLE).Range(DIRECTION_BATEAU).Value = Format(ValeurChoisie2, "0.00")
            End If
            If ValeurChoisie3 <> "" Then
                Worksheets(FEUIL_EXEMPLE).Range(DONNEE_SOG).Value = Format(ValeurChoisie3, "0.00")
                Worksheets(FEUIL_EXEMPLE).Range(VITESSE_FOND).Value = Format(ValeurChoisie3, "0.00")
            End If
        End If
     
     Case 129026 'COG & SOG
        ii = ii + 1
        If ValeurChoisie <> "" Then
            Worksheets(FEUIL_EXEMPLE).Range(COG).Value = Format(ValeurChoisie, "0.00")
        End If
        If ValeurChoisie2 <> "" Then
            Worksheets(FEUIL_EXEMPLE).Range(SOG).Value = Format(ValeurChoisie2, "0.00")
            Worksheets(FEUIL_EXEMPLE).Range(VITESSE_FOND).Value = Format(ValeurChoisie2, "0.00")
        End If
    
    Case 129025 'Position du bateau
        ii = ii + 1
        If ValeurChoisie <> "" Then
            Worksheets(FEUIL_EXEMPLE).Range(POSITION_LATITUDE).Value = Format(Fix(ValeurChoisie), "0") & "°" & Format((ValeurChoisie - Fix(ValeurChoisie)) * 60, "0.0000") & "'"
        End If
        If ValeurChoisie2 <> "" Then
            Worksheets(FEUIL_EXEMPLE).Range(POSITION_LONGITUDE).Value = Format(Fix(ValeurChoisie2), "0") & "°" & Format((ValeurChoisie2 - Fix(ValeurChoisie2)) * 60, "0.0000") & "'"
        End If
    
    Case 128267 'Profondeur
        ii = ii + 1
        If ValeurChoisie <> "" Then
            Worksheets(FEUIL_EXEMPLE).Range(PROFONDEUR).Value = Format(ValeurChoisie, "0.00")
       End If
    
    Case 127250 'Heading
        ii = ii + 1
        If ValeurChoisie <> "" Then
            Worksheets(FEUIL_EXEMPLE).Range(DONNEE_HEADING).Value = Format(ValeurChoisie, "0.00")
       End If
    
    Case 130310 'Données climatiques
        ii = ii + 1
        If ValeurChoisie <> "" Then
            Worksheets(FEUIL_EXEMPLE).Range(PRESSION).Value = Format(ValeurChoisie, "0")
        End If
        If ValeurChoisie2 <> "" Then
            Worksheets(FEUIL_EXEMPLE).Range(TEMPARATURE_EAU_GPS2).Value = Format(ValeurChoisie2, "0.00")
        End If
     
     Case 127245 'Pilote auto orientation de la barre
        ii = ii + 1
        If ValeurChoisie <> PAS_INFO And ValeurChoisie <> "" Then
            Worksheets(FEUIL_EXEMPLE).Range(ORIENTATION_BARRE).Value = Format(ValeurChoisie, "0.00")
        End If
     
     Case 127505 'Niveaux réservoirs
        ii = ii + 1
        If ValeurChoisieTab = "1" Then 'Sur la trame 1
            If ValeurChoisie <> "" Then
                Worksheets(FEUIL_EXEMPLE).Range(NIVEAU_EAU).Value = Format(ValeurChoisie, "0.00")
            End If
            If ValeurChoisie2 <> "" Then
                Worksheets(FEUIL_EXEMPLE).Range(QUANTITE_TOTAL_EAU).Value = Format(ValeurChoisie2, "0.00")
            End If
            Worksheets(FEUIL_EXEMPLE).Range(CALCUL_RESTANT_EAU).Value = Format(ValeurChoisie2, "0.00") * Format(ValeurChoisie, "0.00") / 100
            
        End If
        If ValeurChoisieTab = "0" Then 'Sur la trame 0
            If ValeurChoisie <> "" Then
                Worksheets(FEUIL_EXEMPLE).Range(NIVEAU_GASOIL).Value = Format(ValeurChoisie, "0.00")
            End If
            If ValeurChoisie2 <> "" Then
                Worksheets(FEUIL_EXEMPLE).Range(QUANTITE_GASOIL).Value = Format(ValeurChoisie2, "0.00")
            End If
            Worksheets(FEUIL_EXEMPLE).Range(CALCUL_RESTANT_GASOIL).Value = Format(ValeurChoisie2, "0.00") * Format(ValeurChoisie, "0.00") / 100
        End If
     
     Case 130312 'Températures
        ii = ii + 1
        If ValeurChoisieTab = "2" Then
            If ValeurChoisie <> "" Then
                Worksheets(FEUIL_EXEMPLE).Range(TEMPERATURE_CABINE).Value = Format(ValeurChoisie, "0.00")
            End If
        End If
        If ValeurChoisieTab = "3" Then
            If ValeurChoisie <> "" Then
                Worksheets(FEUIL_EXEMPLE).Range(TEMPERATURE_MOTEUR).Value = Format(ValeurChoisie, "0.00")
            End If
        End If
        If ValeurChoisieTab = "0" Then
            If ValeurChoisie <> "" Then
                Worksheets(FEUIL_EXEMPLE).Range(TEMPERATURE_EAU_GPS).Value = Format(ValeurChoisie, "0.00")
            End If
        End If
    
    Case 130316 'Température étendue
        ii = ii + 1
        If ValeurChoisie <> "" Then
            Worksheets(FEUIL_EXEMPLE).Range(TEMPERATURE_EAU_SPEED).Value = Format(ValeurChoisie, "0.00")
        End If


End Select

'Ecrit le nombre d'informations utilisées
Worksheets(FEUIL_EXEMPLE).Range("E25").Value = ii

End Sub

