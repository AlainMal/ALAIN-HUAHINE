Attribute VB_Name = "Fonction_PGN"
Option Explicit

'============== Procédure PGN_Decode ===============
'
'Ce programme réalise l'analyse des octets recçus par PGN du bus CAN
'pour les traduire en NMEA 2000
'
'=================================================


'================================================
'Procédure de génération des lignes NMEA 2000
'================================================
Sub PGN_Decode()

pi = Application.WorksheetFunction.pi() 'Valeur de pi


'Defini les octets du bus CAN
Dim a1 As Variant, a2 As Variant, a3 As Variant, a4 As Variant, a5 As Variant, a6 As Variant, a7 As Variant, a8 As Variant

Dim z As String                                                                             'Valeur mesuurée
Dim ligne As Long                                                                        'Numéro de ligne
Dim PGN1 As Variant, PGN2 As Variant, PGN3 As Variant                                        'Le texte du PGN
Dim aa As String, bb As String, cc As Long                                               'Nombre de ligne à générer
Dim Debut As Long, Fin As Long                                                        'Valeur indiquant le début et la fin
Dim az As String                                                                            'Contient le "." pour afficher "Patienter..."
Dim Memoire(NB_OCTETS, Nb_PGN) As String                                                     'Tableau pour les chevauchements
Dim aFin As Boolean                                                                         'Indique la fin de la génération
Dim i As Long
Dim Tps, Tps2                                                                              'Mesure le temps de la génération

Application.OnKey "{ESC}"  'au cas ou le curseur est en cours de modification

'Vérifie que les feuilles existent
If FeuilleExiste(FEUIL_NMEA) = False Then
    If MsgBox("La feuille " & FEUIL_NMEA & " porte le mauvais nom" & vbCr & "Pour re-nonnmer automatiquement cliquez sur ""Recommencer""" & vbCr & vbCr & vbCr & "Sinon aller voir dans le module ""Constantes""", vbCritical + vbRetryCancel) = vbRetry Then
            Call Renommer
    Else
        GoTo FINFINFIN
    End If
End If
If FeuilleExiste(FEUIL_EXEMPLE) = False Then
    If MsgBox("La feuille " & FEUIL_EXEMPLE & " porte le mauvais nom" & vbCr & "Pour re-nonnmer automatiquement cliquez sur ""Recommencer""" & vbCr & vbCr & vbCr & "Sinon aller voir dans le module ""Constantes""", vbCritical + vbRetryCancel) = vbRetry Then
            Call Renommer
    Else
        GoTo FINFINFIN
    End If
End If
If FeuilleExiste(FEUIL_TEMPS) = False Then
    If MsgBox("La feuille " & FEUIL_TEMPS & " porte le mauvais nom" & vbCr & "Pour re-nonnmer automatiquement cliquez sur ""Recommencer""" & vbCr & vbCr & vbCr & "Sinon aller voir dans le module ""Constantes""", vbCritical + vbRetryCancel) = vbRetry Then
            Call Renommer
    Else
        GoTo FINFINFIN
    End If
End If
If FeuilleExiste(FEUIL_MMSI) = False Then
    If MsgBox("La feuille " & FEUIL_MMSI & " porte le mauvais nom" & vbCr & "Pour re-nonnmer automatiquement cliquez sur ""Recommencer""" & vbCr & vbCr & vbCr & "Sinon aller voir dans le module ""Constantes""", vbCritical + vbRetryCancel) = vbRetry Then
            Call Renommer
    Else
        GoTo FINFINFIN
    End If
End If


'Vient récupérer les adresses écritent par l'itilisateur sur les colonnes AD et AE (32 @ maxi) dans un tableau
For i = 0 To 32
    Participants(i, ADRESSE_PARTICIPANT) = Worksheets(FEUIL_NMEA).Cells(i + 2, COLONNE_ADRESSE)
    Participants(i, NOM_PARTICIPANT) = Worksheets(FEUIL_NMEA).Cells(i + 2, COLONNE_NOM)
Next i

StopA = False                                                                               'La valeur peut être mise à True par le bouton "Arrêtrer

Derniereligne = Worksheets(FEUIL_NMEA).Cells(Rows.Count, 1).End(xlUp).Row                    'Compte le nombre de lignes maximum

'Initialise la valeur choisie par Input
If Choix = "" Then
    Choix = "2:" & Derniereligne
End If
 

'Qd il y a une erreur de saisie on rcommence
RECOMMENCE:

'--------------------------------------------------------- Pose la question pour rafraichir -------------------------------------------------------------------------
'Ne tient pas compte des erreurs
On Error Resume Next

'Récupère les numéros de lignes à rafraichir
Choix = InputBox("Entrez le groupement des trames à rafraichir" & vbCr & vbCr & "Sous la forme -> Début : Fin" & vbCr & vbCr & "Nombres de trames minimum =" & 2 & " et maximum =" & Derniereligne, "Génération", Choix)     'Envoie la question "nombre de trames"

'Si on annule le choix on va à la fin
If Choix = "" Then GoTo FinFin

'défini le début et la fin du choix
cc = InStr(1, Choix, ":")                   'Emplacment du ":"
aa = Left(Choix, cc - 1)                    'Récupère la partie gauche
bb = Right(Choix, Len(Choix) - Len(aa) - 1) 'Récupère la partie droite
Debut = CDbl(aa)                            'Converti la partie gauche en integer
Fin = CDbl(bb)                              'Converti la partie droite en integer

'Réinitialise les erreurs
On Error GoTo 0

'Si le choix est > à la longueur on redéfini le Choix
If Fin > Derniereligne Then Choix = Debut & ":" & Fin

'Envoi le message s'il y a une erreur de dimension
'Va à la fin si le nombre de ligne est en dehors des limite
If Debut < 2 Or Fin > Derniereligne Or Debut > Fin Then
    If MsgBox("Entrez les valeurs entre 2 et " & Derniereligne & vbCr & "Sous la forme -> Début : Fin", vbQuestion + vbOKCancel, "Génération") = vbOK Then GoTo RECOMMENCE Else GoTo FinFin
    GoTo FinFin
End If
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------


'-------------------- Initialisation -----------------------------
EnCours = True  'La génération est en cours
User_Form_PGN.Patienter.Caption = "< En cours >" & vbCr

iii = 1  'Compte la totalité des ligne avec le buffer tournat

'Déprotège les feuilles
Worksheets(FEUIL_NMEA).Unprotect
Worksheets(FEUIL_EXEMPLE).Unprotect
'Worksheets(FEUIL_TEMPS).Unprotect 'On ne déprotège la feuille à la fin lors ont écrit dessus

ligne = 2                                                                                  'Commence par la ligne 2
az = "."                                                                                   'Affiche le nombre de point sur "Patientez..."

'Prépare l'etat de l'application
Tps = Timer                         'Mesure le temps au départ
ii = 0                              'Initialise le nombre d'informations faites sur Exemple
Application.Cursor = xlWait         'Affiche le curseur en attente

User_Form_PGN.Debut.Caption = Debut     'Ecrit la valeur du début

'Si l'Apperçu seul n'est pas mis alors on rend visible le bouton RAZ sinon inaccessible
If User_Form_PGN.Appercu = False Then
    User_Form_PGN.CommandButton3.Enabled = True 'Bouton "Arrêter" est visible lors de la génération
Else
    User_Form_PGN.CommandButton3.Enabled = False 'Bouton "Arrêter" est inaccessible lors de la génération
End If


User_Form_PGN.CommandButton1.Enabled = False                    'Bouton "Générer" est éteint pendant la génération
User_Form_PGN.CommandButton4.Enabled = False                    'Bouton "RAZ" est éteint pendant la génération
User_Form_PGN.Appercu.Enabled = False                                 'Case à cocher est inaccessible
User_Form_PGN.FeuxStop.Visible = False                                  'LED stop éteinte

DoEvents 'Laisse la main au système
 
'N'autorise pas l'actualisation automatique
Application.Calculation = xlCalculationManual
  '------------- FIN d'initialisation -----------------------
 

'------- Début de la conversion des PGN ---------
For ligne = Debut To Fin
   
iii = iii + 1 'Compte les lignes pour le buffer tournant
  
Debut:
    
    'Arrête si demnadé par le bouton "Arrêter"
     If StopA = True Then GoTo FinA
   
    'Affiche le ... "en cours" et clignotement de la led si ont est pas en "Exemple seul"
    If Not User_Form_PGN.Appercu Then
        If az = "........" Then az = "."
        If ligne Mod 10 = 0 Then
            'If User_Form_PGN.Buffer = False Then
                User_Form_PGN.Patienter.Caption = "Patientez " & az
            'End If
            az = az & "."
        End If
    
        'Fait clignoter pendant la génération
        If User_Form_PGN.Feux.Visible = True Then
            User_Form_PGN.Feux.Visible = False
        Else
            User_Form_PGN.Feux.Visible = True
        End If
        
        'If User_Form_PGN.Buffer Then
        '        User_Form_PGN.Fin.Caption = iii   'Affiche le numéro du buffer tournant
        'Else
            User_Form_PGN.Fin.Caption = ligne   'Affiche le numéro de la ligne
        'End If
        
        If ligne Mod 500 Then DoEvents                'Laisse la main au système toutes les 500 lignes
    End If
    

' Initialise les 8 cellules reçue en hexadecimal
a1 = Worksheets(FEUIL_NMEA).Cells(ligne, PREMIERE_COLONNE_OCTET)           '--- Poids faible
a2 = Worksheets(FEUIL_NMEA).Cells(ligne, PREMIERE_COLONNE_OCTET + 1)         '         v
a3 = Worksheets(FEUIL_NMEA).Cells(ligne, PREMIERE_COLONNE_OCTET + 2)         '         v
a4 = Worksheets(FEUIL_NMEA).Cells(ligne, PREMIERE_COLONNE_OCTET + 3)         '         v
a5 = Worksheets(FEUIL_NMEA).Cells(ligne, PREMIERE_COLONNE_OCTET + 4)         '         v
a6 = Worksheets(FEUIL_NMEA).Cells(ligne, PREMIERE_COLONNE_OCTET + 5)         '         v
a7 = Worksheets(FEUIL_NMEA).Cells(ligne, PREMIERE_COLONNE_OCTET + 6)         '         v
a8 = Worksheets(FEUIL_NMEA).Cells(ligne, PREMIERE_COLONNE_OCTET + 7)         '--- Poids fort

'Initialise les difféentes valeurs mesurées
ValeurChoisie = PAS_INFO
ValeurChoisie2 = PAS_INFO
ValeurChoisie3 = PAS_INFO
ValeurChoisieTab = PAS_INFO

'Récupère la valeur du PGN dans l'ID et fait le choix
Select Case cID.Pgn(ligne)

    Case 126993
        z = CDbl("&H" & a3 & a2) * 0.001
        ValeurChoisie = z
        PGN1 = "Heart beat"
        Echelle = "Sec"
        FormatAffichage = "0.00"
        
        ValeurChoisie2 = ""
        PGN2 = "Présent"
        Echelle2 = ""

    Case 130306
        z = CDbl("&H" & a3 & a2) * 0.01 * 1.944
        ValeurChoisie = z
        PGN1 = "Noeuds du Vent"
        Echelle = "Nds"
        FormatAffichage = "0.00"
        
        z = CDbl("&H" & a5 & a4) * 0.0001 * 180 / pi
        ValeurChoisie2 = z
        PGN2 = "Direction du vent"
        Echelle2 = "Deg"
        FormatAffichage2 = "0.00"
        
        z = CDbl("&H" & a6)
        ValeurChoisieTab = z And CDbl("&H07")
        PGN_Tab = "Table(" & ValeurChoisieTab & ")-> " & cTable.Wind(ValeurChoisieTab) 'Ecrit le texte correspondant au numéro dans la table

    Case 127245
        z = CDbl("&H" & a6 & a5) * 0.0001 * 180 / pi
        If CDbl("&H" & a6) <> 255 Then
            ValeurChoisie = z
            PGN1 = "la Barre"
            Echelle = "Deg"
            FormatAffichage = "0.00"
        Else
            ValeurChoisie = ""
            PGN1 = "la Barre"
            Echelle = ""
        End If
   
    Case 129026
        z = CDbl("&H" & a4 & a3) * 0.0001 * 180 / pi
        ValeurChoisie = z
        PGN1 = "COG & SOG:      COG"
        Echelle = "Deg"
        FormatAffichage = "0.00"
        
        z = CDbl("&H" & a6 & a5) * 0.0001 * 180 / pi
        ValeurChoisie2 = z
        PGN2 = "SOG"
        Echelle2 = "Nds"
        FormatAffichage2 = "0.00"
    
    Case 127250
        z = CDbl("&H" & a3 & a2) * 0.0001 * 180 / pi
        ValeurChoisie = z
        PGN1 = "Heading"
        Echelle = "Deg"
        FormatAffichage = "0.00"
     
     Case 127251
        z = CDbl("&H" & a3 & a2) * 0.00000003125 * 180 / pi
        ValeurChoisie = z
        PGN1 = "Taux de virement"
        Echelle = "Deg"
        FormatAffichage = "0.00"
    
    Case 128267
        z = CDbl("&H" & a5 & a4 & a3 & a2) * 0.01
        ValeurChoisie = z
        PGN1 = "Profondeur"
        Echelle = "m"
        FormatAffichage = "0.00"
    
    Case 130312
        z = CDbl("&H" & a5 & a4) * 0.01 - 273.15
        ValeurChoisie = z
        PGN1 = "Température"
        Echelle = "°C"
        FormatAffichage = "0.00"
        
        'Affiche la valeur du tableau
        z = CDbl("&H" & a3)
        ValeurChoisieTab = z
        PGN_Tab = "Table(" & ValeurChoisieTab & ")-> " & cTable.Temperature(ValeurChoisieTab)
        EchellTab = ""
        FormatAffichageTab = ""
     
    Case 130316
        z = CDbl("&H" & a6 & a5 & a4) * 0.001 - 273.15
        ValeurChoisie = z
        PGN1 = "Température étendue"
        Echelle = "°C"
        FormatAffichage = "0.00"
                
        'Affiche la valeur du tableau
        z = CDbl("&H" & a3)
        ValeurChoisieTab = z
        PGN_Tab = "Table(" & ValeurChoisieTab & ")-> " & cTable.Temperature(ValeurChoisieTab)
        EchellTab = ""
        FormatAffichageTab = ""

     Case 130310
        z = CDbl("&H" & a7 & a6)
        ValeurChoisie = z
        If z > 32000 Then
            ValeurChoisie = ""
        End If
        PGN1 = "Pression atmosphérique"
        Echelle = "mBar"
        FormatAffichage = "0.00"
        
        z = CDbl("&H" & a3 & a2) * 0.01 - 273.15
        ValeurChoisie2 = z
        If CDbl("&H" & a3 & a2) > 32000 Then
            ValeurChoisie2 = ""
        End If
        PGN2 = "Température de l'eau"
        Echelle2 = "°C"
        FormatAffichage2 = "0.00"
        
        z = CDbl("&H" & a5 & a4) * 0.01 - 273.15
        ValeurChoisie3 = z
        If CDbl("&H" & a5 & a4) > 32000 Then
            ValeurChoisie3 = ""
        End If
        PGN3 = "Température de l'air"
        Echelle3 = "C°"
        FormatAffichage3 = "0.00"
          
    Case 128259
        z = CDbl("&H" & a3 & a2) * 0.01 * 0.01944
        ValeurChoisie = z
        If CDbl("&H" & a3 & a2) > 255 Then
            ValeurChoisie = ""
        End If
        PGN1 = "Vitesse surface"
        Echelle = "Nds"
        FormatAffichage = "0.00"
        
        z = CDbl("&H" & a5 & a4) * 0.01 * 0.01944
        ValeurChoisie2 = z
        If CDbl("&H" & a5 & a4) > 255# Then
            ValeurChoisie2 = ""
        End If
        PGN2 = "Vitesse fond SOG"
        Echelle2 = "Nds"
        FormatAffichage2 = "0.00"
        
        z = CDbl("&H" & a6) And CDbl("&H07")
        If z < 7 Then
            ValeurChoisieTab = z
            PGN_Tab = "Table(" & ValeurChoisieTab & ")-> " & cTable.Speed(ValeurChoisieTab)
        End If
        
    Case 127508
        z = CDbl("&H" & a3 & a2) * 0.01
        If z < 100 Then
            ValeurChoisie = z
        End If
        PGN1 = "Volts Batterie"
        Echelle = "V"
        FormatAffichage = "0.00"
        
        z = CDbl("&H" & a5)
        ValeurChoisieTab = z
        If z = CDbl("&HF7") Then
            PGN_Tab = "Batterie Moteur"
        Else
            PGN_Tab = "Batterie Service"
        End If
        
        z = CDbl("&H" & a7 & a6) * 0.01 - 273.15
        If z < 100 Then
            ValeurChoisie3 = z
        Else
            ValeurChoisie3 = ""
        End If
        PGN3 = "Température Batterie"
        Echelle3 = "°C"
        FormatAffichage3 = "0.00"
   
    Case 129025
        z = CDbl("&H" & a4 & a3 & a2 & a1) * (10) ^ -7
        ValeurChoisie = z
        PGN1 = "Position rapide:        Lattitude"
        Echelle = "Deg"
        FormatAffichage = "0.0000"
        Ma_position_latitude = ValeurChoisie
        
        z = CDbl("&H" & a8 & a7 & a6 & a5) * (10) ^ -7
        ValeurChoisie2 = z
        PGN2 = "Longitude"
        Echelle2 = "Deg"
        FormatAffichage2 = "0.0000"
        Ma_position_longitude = ValeurChoisie2

    Case 129038
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "AIS Posittion Class A"
        Echelle = "N°"
        FormatAffichage = "0"
        
        If ValeurChoisie = 0 Then
            z = CLng("&H" & a7 & a6 & a5 & a4)
            ValeurChoisie2 = z
            PGN2 = "MMSI"
            Echelle2 = "N°"
            FormatAffichage2 = "###\ ###\ ###"
            
            Memoire(MEMOIRE_PGN_a7, PGN_129038) = CDbl("&H" & a7)
            Memoire(MEMOIRE_PGN_a8, PGN_129038) = CDbl("&H" & a8)
           
        ElseIf ValeurChoisie = 1 Then
            z = CLng("&H" & a8 & a7 & a6 & a5) * 10 ^ -7
            ValeurChoisie2 = z
            PGN2 = "Lattitude"
            Echelle2 = "°3"
            FormatAffichage2 = "0.0000"

            z = CLng("&H" & a4 & a3 & a3 & Memoire(MEMOIRE_PGN_a7, PGN_129038)) * 10 ^ -7
            ValeurChoisie3 = z
            PGN3 = "Longitude"
            Echelle3 = "Deg"
            FormatAffichage3 = "0.0000"
            
            
        ElseIf ValeurChoisie = 2 Then
            z = CLng("&H" & a4 & a3) * 0.0001 * 180 / pi
            ValeurChoisie2 = z
            PGN2 = "COG"
            Echelle2 = "Deg"
            FormatAffichage2 = "0.00"
            
            z = CLng("&H" & a6 & a5) * 0.01 * 0.01944
            ValeurChoisie3 = z
            PGN3 = "SOG"
            Echelle3 = "Nds"
            FormatAffichage3 = "0.00"

        ElseIf ValeurChoisie = 3 Then
            z = CLng("&H" & a2 & a3) * 0.0001 * 180 / pi
            ValeurChoisie2 = z
            PGN2 = "??True Heading"
            Echelle2 = "Deg"
            FormatAffichage2 = "0.00"

        End If

    Case 129794
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "AIS A Données de voyage"
        Echelle = "N°"
        FormatAffichage = "0"
        
        If ValeurChoisie = 0 Then
            z = CLng("&H" & a7 & a6 & a5 & a4)
            ValeurChoisie2 = z
            PGN2 = "MMSI"
            Echelle2 = "N°"
            FormatAffichage2 = "###\ ###\ ###"
        ElseIf ValeurChoisie = 2 Then
            z = Chr(CInt("&H" & a5)) & Chr(CInt("&H" & a6)) & Chr(CInt("&H" & a7)) & Chr(CInt("&H" & a8))
            ValeurChoisie2 = z
            PGN2 = "Nom du navire"
            Echelle2 = "Texte"
            FormatAffichage2 = "#"
        
        ElseIf ValeurChoisie = 3 Or ValeurChoisie = 4 Then
            z = Chr(CInt("&H" & a2)) & Chr(CInt("&H" & a3)) & Chr(CInt("&H" & a4)) & Chr(CInt("&H" & a5)) & Chr(CInt("&H" & a6)) & Chr(CInt("&H" & a7)) & Chr(CInt("&H" & a8))
            'z = CLng("&H" & a8 & a7 & a6 & a5)
            ValeurChoisie2 = z
            PGN2 = "Nom du navire"
            Echelle2 = "Texte"
            FormatAffichage2 = "#"
        
        ElseIf ValeurChoisie = 5 Then
            z = CLng("&H" & a6 & a5)
            ValeurChoisie2 = z * 0.1
            PGN2 = "Longueur"
            Echelle2 = "m"
            FormatAffichage2 = "0.0"
            
            z = CLng("&H" & a8 & a7)
            ValeurChoisie3 = z * 0.1
            PGN3 = "Largeur"
            Echelle3 = "m"
            FormatAffichage3 = "0.0"
           
         ElseIf ValeurChoisie = 6 Then
            z = CLng("&H" & a4 & a3)
            ValeurChoisie2 = z * 0.0001 / 1.85
            PGN2 = "Distance"
            Echelle2 = "MN"
            FormatAffichage2 = "0.00"

            z = CLng("&H" & a6 & a5)
            ValeurChoisie3 = z * 0.0001 / 1.85
            PGN3 = "Distance proue"
            Echelle3 = "MN"
            FormatAffichage3 = "0.00"
        
         ElseIf ValeurChoisie = 7 Then
            z = Chr(CInt("&H" & a7)) & Chr(CInt("&H" & a8))
            ValeurChoisie2 = z
            PGN2 = "Destination"
            Echelle2 = "Texte"
            FormatAffichage2 = "#"
            
         ElseIf ValeurChoisie = 8 Then
            z = Chr(CInt("&H" & a2)) & Chr(CInt("&H" & a3)) & Chr(CInt("&H" & a4)) & Chr(CInt("&H" & a5)) & Chr(CInt("&H" & a6)) & Chr(CInt("&H" & a7)) & Chr(CInt("&H" & a8))
            ValeurChoisie2 = z
            PGN2 = "Destination"
            Echelle2 = "Texte"
            FormatAffichage2 = "#"
         
         ElseIf ValeurChoisie = 9 Then
            z = Chr(CInt("&H" & a2)) & Chr(CInt("&H" & a3)) & Chr(CInt("&H" & a4)) & Chr(CInt("&H" & a5)) & Chr(CInt("&H" & a6)) & Chr(CInt("&H" & a7)) & Chr(CInt("&H" & a8))
            ValeurChoisie2 = z
            PGN2 = "Destination"
            Echelle2 = "Texte"
            FormatAffichage2 = "#"
            
        End If
     Case 129809
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "AIS classe B données"
        Echelle = "N°"
        FormatAffichage = "0"
    
         If ValeurChoisie = 0 Then
            z = CLng("&H" & a7 & a6 & a5 & a4)
            ValeurChoisie2 = z
            PGN2 = "MMSI"
            Echelle2 = "N°"
            FormatAffichage2 = "###\ ###\ ###"
            
            z = Chr(CInt("&H" & a8))
            ValeurChoisie3 = z
            PGN3 = "Nom du navire"
            Echelle3 = "Texte"
            FormatAffichage3 = "#"
        End If
         
        If ValeurChoisie = 1 Then
            z = Chr(CInt("&H" & a2)) & Chr(CInt("&H" & a3)) & Chr(CInt("&H" & a4)) & Chr(CInt("&H" & a5)) & Chr(CInt("&H" & a6)) & Chr(CInt("&H" & a7)) & Chr(CInt("&H" & a8))
            ValeurChoisie3 = z
            PGN3 = "Nom du navire"
            Echelle3 = "Texte"
            FormatAffichage3 = "#"
        End If
         
         If ValeurChoisie = 2 Then
            z = Chr(CInt("&H" & a2)) & Chr(CInt("&H" & a3)) & Chr(CInt("&H" & a4)) & Chr(CInt("&H" & a5)) & Chr(CInt("&H" & a6))
            ValeurChoisie3 = z
            PGN3 = "Nom du navire"
            Echelle3 = "Texte"
            FormatAffichage3 = "#"
        End If
         
     Case 129039
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "AIS Posittion Class B"
        Echelle = "N°"
        FormatAffichage = "0"
    
         If ValeurChoisie = 0 Then
            z = CLng("&H" & a7 & a6 & a5 & a4)
            ValeurChoisie2 = z
            PGN2 = "MMSI"
            Echelle2 = "N°"
            FormatAffichage2 = "###\ ###\ ###"
            Memoire(MEMOIRE_PGN_a8, PGN_129039) = CDbl("&H" & a8)
            
        ElseIf ValeurChoisie = 1 Then
            'Ne donne pas les bonnes longitude
            z = CLng("&H" & a4 & a3 & a2 & Memoire(MEMOIRE_PGN_a8, PGN_129039)) * 10 ^ -7
            ValeurChoisie2 = z '
            PGN2 = "Longitude" '
            Echelle2 = "Deg"
            FormatAffichage2 = "0.0000"
            
            z = CLng("&H" & a8 & a7 & a6 & a5) * 10 ^ -7
            ValeurChoisie3 = z
            PGN3 = "Lattitude"
            Echelle3 = "Deg"
            FormatAffichage3 = "0.0000"
            
        ElseIf ValeurChoisie = 2 Then
            z = CLng("&H" & a4 & a3) * 0.0001 * 180 / pi
            ValeurChoisie2 = z
            PGN2 = "COG"    'Cap suivi par le bateau
            Echelle2 = "Deg"
            FormatAffichage2 = "0.00"
            
            z = CLng("&H" & a6 & a5) * 0.01 * 0.01944
            ValeurChoisie3 = z
            PGN3 = "SOG"
            Echelle3 = "Nds"
            FormatAffichage3 = "0.00"
            
            z = CLng("&H" & a8 & a7) * 10 ^ -7
            ValeurChoisieTab = z
            PGN_Tab = "Heading " & Format(ValeurChoisieTab, "0.0000")
            'EchelleTab = "Deg"
            FormatAffichageTab = "0.00"
        End If
     
     Case 129049
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "AIS Posittion Class B"
        Echelle = "N°"
        FormatAffichage = "0"
    
         If ValeurChoisie = 0 Then
            z = CLng("&H" & a7 & a6 & a5 & a4)
            ValeurChoisie2 = z
            PGN2 = "MMSI"
            Echelle2 = "N°"
            FormatAffichage2 = "###\ ###\ ###"
            Memoire(MEMOIRE_PGN_a8, PGN_129049) = CDbl("&H" & a8)
            
        ElseIf ValeurChoisie = 1 Then
            z = CLng("&H" & a4 & a3 & a2 & Memoire(MEMOIRE_PGN_a8, PGN_129049)) * 10 ^ -7
            ValeurChoisie2 = z
            PGN2 = "Longitude"
            Echelle2 = "Deg"
            FormatAffichage2 = "0.0000"
            
            z = CLng("&H" & a8 & a7 & a6 & a5) * 10 ^ -7
            ValeurChoisie3 = z
            PGN3 = "Lattitude"
            Echelle3 = "Deg"
            FormatAffichage3 = "0.0000"
            
        ElseIf ValeurChoisie = 2 Then
            z = CLng("&H" & a4 & a3) * 0.0001 * 180 / pi
            ValeurChoisie3 = z
            PGN3 = "COG"
            Echelle3 = "Deg"
            FormatAffichage3 = "0.00"
            
            z = CLng("&H" & a6 & a5) * 0.01 * 0.01944
            ValeurChoisie2 = z
            PGN2 = "SOG"
            Echelle2 = "Nds"
            FormatAffichage2 = "0.00"
        End If
     
    Case 129810
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "AIS Données statique Class B"
        Echelle = "N°"
        FormatAffichage = "0"
    
        If ValeurChoisie = 0 Then
            z = CLng("&H" & a7 & a6 & a5 & a4)
            ValeurChoisie2 = z
            PGN2 = "MMSI"
            Echelle2 = "N°"
            FormatAffichage2 = "###\ ###\ ###"
        
        ElseIf ValeurChoisie = 1 Then
            z = Chr(CInt("&H" & a2)) & Chr(CInt("&H" & a3)) & Chr(CInt("&H" & a4)) & Chr(CInt("&H" & a5)) & Chr(CInt("&H" & a6)) & Chr(CInt("&H" & a7)) & Chr(CInt("&H" & a8))
            ValeurChoisie2 = z
            PGN2 = "Identifiant"
            Echelle2 = ""
       
        ElseIf ValeurChoisie = 2 Then
            z = Chr(CInt("&H" & a2)) & Chr(CInt("&H" & a3)) & Chr(CInt("&H" & a4)) & Chr(CInt("&H" & a5)) & Chr(CInt("&H" & a6)) & Chr(CInt("&H" & a7)) & Chr(CInt("&H" & a8))
            ValeurChoisie2 = z
            PGN2 = "Indicatif"
            Echelle2 = ""
       
        Else
            ValeurChoisie2 = ""
            PGN2 = ""
            Echelle2 = ""
        End If

    Case 129283
        z = CDbl("&H" & a6 & a5 & a4 & a3) * 0.01
        ValeurChoisie = z
        If Abs(CDbl("&H" & a6 & a5 & a4 & a3)) > 2 ^ 30 Then
            ValeurChoisie = ""
        End If
        PGN1 = "XTE"
        Echelle = "m"
        FormatAffichage = "0.00"
      
      Case 126996
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "Information produit (" & ValeurChoisie & ")"
        Echelle = "N°"
        FormatAffichage = "0"
        
        If ValeurChoisie = 0 Then
            z = IIf(a7 = "ff", "", Chr(CInt("&H" & a7))) & IIf(a8 = "ff", "", Chr(CInt("&H" & a8)))
            ValeurChoisie2 = z
            Echelle2 = ""
            PGN2 = "Configuration"
            aFin = False
            
        ElseIf aFin = False And (ValeurChoisie = 1 Or ValeurChoisie = 2 Or ValeurChoisie = 2 Or ValeurChoisie = 4) Then
            z = IIf(a2 = "ff" Or a2 = "00", "", Chr(CInt("&H" & a2))) & Chr(CInt("&H" & a3)) & Chr(CInt("&H" & a4)) & Chr(CInt("&H" & a5)) & Chr(CInt("&H" & a6)) & Chr(CInt("&H" & a7)) & Chr(CInt("&H" & a8))
            If a2 <> "ff" Or a3 <> "ff" Or a4 <> "ff" Or a5 <> "ff" Or a6 <> "ff" Or a7 <> "ff" Or a8 <> "ff" Then
                ValeurChoisie2 = z
                Echelle2 = "Texte"
                PGN2 = "Configuration"
            Else
                ValeurChoisie2 = ""
                Echelle2 = ""
                PGN2 = ""
                aFin = True
            End If
        
         ElseIf ValeurChoisie = 5 Then
            z = IIf(a4 = "ff", "", Chr(CInt("&H" & a4))) & IIf(a5 = "ff", "", Chr(CInt("&H" & a5))) & IIf(a6 = "ff", "", Chr(CInt("&H" & a6))) & IIf(a7 = "ff", "", Chr(CInt("&H" & a7))) & IIf(a8 = "ff", "", Chr(CInt("&H" & a8)))
            ValeurChoisie2 = z
            Echelle2 = ""
            PGN2 = "Version"
            aFin = False

         ElseIf ValeurChoisie = 10 Then
            z = IIf(a4 = "ff", "", Chr(CInt("&H" & a4))) & IIf(a5 = "ff", "", Chr(CInt("&H" & a5))) & IIf(a6 = "ff", "", Chr(CInt("&H" & a6))) & IIf(a7 = "ff", "", Chr(CInt("&H" & a7))) & IIf(a8 = "ff", "", Chr(CInt("&H" & a8)))
            ValeurChoisie2 = z
            Echelle2 = ""
            PGN2 = "Compléments"
            aFin = False
         
         ElseIf ValeurChoisie = 11 Then
            z = IIf(a4 = "ff", "", Chr(CInt("&H" & a4))) & IIf(a5 = "ff", "", Chr(CInt("&H" & a5))) & IIf(a6 = "ff", "", Chr(CInt("&H" & a6))) & IIf(a7 = "ff", "", Chr(CInt("&H" & a7))) & IIf(a8 = "ff", "", Chr(CInt("&H" & a8)))
            ValeurChoisie2 = z
            Echelle2 = ""
            PGN2 = "Indiquation"
            aFin = False

         ElseIf ValeurChoisie = 12 Then
            z = IIf(a4 = "ff", "", Chr(CInt("&H" & a4))) & IIf(a5 = "ff", "", Chr(CInt("&H" & a5))) & IIf(a6 = "ff", "", Chr(CInt("&H" & a6))) & IIf(a7 = "ff", "", Chr(CInt("&H" & a7))) & IIf(a8 = "ff", "", Chr(CInt("&H" & a8)))
            ValeurChoisie2 = z
            Echelle2 = ""
            PGN2 = "Indiquation"
            aFin = False

         ElseIf aFin = False And (ValeurChoisie = 6 Or ValeurChoisie = 7 Or ValeurChoisie = 8 Or ValeurChoisie = 9) Then
            If a2 <> "ff" Or a3 <> "ff" Or a4 <> "ff" Or a5 <> "ff" Or a6 <> "ff" Or a7 <> "ff" Or a8 <> "ff" Then
                z = IIf(a2 = "ff" Or a2 = "00", "", Chr(CInt("&H" & a2)))
                z = z & IIf(a3 = "ff" Or a3 = "00", "", Chr(CInt("&H" & a3)))
                z = z & IIf(a4 = "ff" Or a4 = "00", "", Chr(CInt("&H" & a4)))
                z = z & IIf(a5 = "ff" Or a5 = "00", "", Chr(CInt("&H" & a5)))
                z = z & IIf(a6 = "ff" Or a6 = "00", "", Chr(CInt("&H" & a8)))
                z = z & IIf(a7 = "ff" Or a7 = "00", "", Chr(CInt("&H" & a7)))
                z = z & IIf(a8 = "ff" Or a8 = "00", "", Chr(CInt("&H" & a8)))
                ValeurChoisie2 = z
                If ValeurChoisie2 = "" Then
                    PGN2 = ""
                Else
                    PGN2 = "Version"
                End If
                Echelle2 = ""
            Else
                aFin = True
            End If

         End If

        FormatAffichage2 = "#"
        
      Case 126998
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "Info. configuration (" & ValeurChoisie & ")"
        Echelle = "N°"
        FormatAffichage = "0"
         
        If ValeurChoisie = 0 Then
            z = "" 'IIf(a7 = "ff", "", Chr(CInt("&H" & a7))) & IIf(a8 = "ff", "", Chr(CInt("&H" & a8)))
            ValeurChoisie2 = z
            Echelle2 = ""
        Else
            z = IIf(a2 = "ff", "", Chr(CInt("&H" & a2))) & IIf(a3 = "ff", "", Chr(CInt("&H" & a3))) & IIf(a4 = "ff", "", Chr(CInt("&H" & a4))) & IIf(a5 = "ff", "", Chr(CInt("&H" & a5))) & IIf(a6 = "ff", "", Chr(CInt("&H" & a6))) & IIf(a7 = "ff", "", Chr(CInt("&H" & a7))) ' & a3 & a4 & a5 & a6 & a7))
            ValeurChoisie2 = z
            Echelle2 = "Texte"
        End If
        FormatAffichage2 = "#"
        PGN2 = "Configuration"
      
      Case 129029
        z = CDbl("&H" & a1) And CDbl("&h1F")
        'z = ""
        ValeurChoisie = z
        PGN1 = "Infor. de positioon GNSS"
        Echelle = "N°"
        FormatAffichage = "0"
        
        If ValeurChoisie = 4 Then
            z = CLng("&H" & a8)
            ValeurChoisie2 = z
            PGN2 = "Nb de satélites"
            Echelle2 = "Nb"
            FormatAffichage2 = "0"
       End If
       
       Case 129029
        z = CDbl("&H" & a1) And CDbl("&h1F")
        'z = ""
        ValeurChoisie = z
        PGN1 = "Infor? de positioon GNSS"
        Echelle = "N°"
        FormatAffichage = "0"
      
      Case 129539
        z = "" 'CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "Présision GNSS "
        Echelle = ""
        FormatAffichage = "0"
     
     Case 130577
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "Données de direction (" & ValeurChoisie & ")"
        Echelle = "N°"
        FormatAffichage = "0"
        
        If ValeurChoisie = 1 Then
            z = CDbl("&H" & a3 & a2) * 0.0001 * 180 / pi
            ValeurChoisie2 = z
            PGN2 = "COG"
            Echelle2 = "Deg"
            FormatAffichage2 = "0.0"
        
            z = CDbl("&H" & a5 & a4) * 0.001 * 180 / pi
            ValeurChoisie3 = z
            PGN3 = "SOG"
            Echelle3 = "Nds"
            FormatAffichage3 = "0.0"
        End If
     
     Case 127506
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "Etat batteries détaillées"
        Echelle = "N°"
        FormatAffichage = "0"
        
        If ValeurChoisie = 0 Then
            If CDbl("&H" & a8) <> 255 Then
                z = CDbl("&H" & a6)
                ValeurChoisie2 = z
                PGN2 = "Etat de charge"
                Echelle2 = "%"
                FormatAffichage2 = "0"
                ValeurChoisieTab = 0
            Else
                z = CDbl("&H" & a6)
                ValeurChoisie2 = z
                PGN2 = "Etat de charge"
                Echelle2 = "%"
                FormatAffichage2 = "0"
                ValeurChoisieTab = 1
            End If
            
            z = CDbl("&H" & a7)
            ValeurChoisie3 = z
            PGN3 = "Etat de santé"
            Echelle3 = "%"
            FormatAffichage3 = "0"
            
            z = CDbl("&H" & a8)
            If a8 = "ff" Then
                PGN_Tab = ""
            Else
                PGN_Tab = "Batterie Service"
            End If
            Memoire(MEMOIRE_PGN_a8, PGN_127506) = a8
        End If
        
        If ValeurChoisie = 1 Then
            z = CDbl("&H" & a2 & Memoire(MEMOIRE_PGN_a8, PGN_127506))
            If Memoire(MEMOIRE_PGN_a8, PGN_127506) = "ff" Then
                ValeurChoisie2 = ""
                FormatAffichage2 = "#"
                PGN2 = ""
            Else
                ValeurChoisie2 = z
                FormatAffichage2 = "0.00"
                Echelle2 = "j"
                PGN2 = "Temps restant"
            End If
        End If
     
    Case 59392
        z = CDbl("&H" & a8 & a7 & a6)
        ValeurChoisie = z
        PGN1 = "Aquittement"
        Echelle = "PGN"
        FormatAffichage = "0"

    Case 59904
        z = CDbl("&H" & a3 & a2 & a1)
        ValeurChoisie = z
        PGN1 = "Reclame PGN:" & ValeurChoisie & " à l'@ " & cID.DestinationID(ligne)
        Echelle = "PGN"
        FormatAffichage = "0"

    Case 60928
        z = "" 'CDbl("&H" & a3 & a2 & a1)
        ValeurChoisie = z
        PGN1 = "Adresse revendiquée"
        Echelle = ""
        FormatAffichage = "0"

        z = CDbl("&H" & a6)
        ValeurChoisie2 = z
        PGN2 = "Fonction du device"
        Echelle2 = "Tab"
        FormatAffichage2 = "0"

        z = CDbl("&H" & a7)
        ValeurChoisie3 = z
        PGN3 = "Classe"
        Echelle3 = "Tab"
        FormatAffichage3 = "0"

     Case 129029
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "Info. de positioon GNSS (" & ValeurChoisie & ")"
        Echelle = "N°"
        FormatAffichage = "0"
           
     Case 129540
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "Satélites en vues (" & ValeurChoisie & ")"
        Echelle = "N°"
        FormatAffichage = "0"
           
        If ValeurChoisie = 0 Then
            z = CDbl("&H" & a6)
            ValeurChoisie2 = z
            PGN2 = "Nombre de satélites"
            Echelle2 = "Nb"
            FormatAffichage2 = "0"
        End If
           
     Case 126720
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "Info. propriétaire (" & ValeurChoisie & ")"
        Echelle = "N°"
        FormatAffichage = "0"
        
        If ValeurChoisie = 0 Then
            z = CDbl("&H" & a3)
            ValeurChoisie2 = z
            PGN2 = "Code"
            Echelle2 = "N°"
            FormatAffichage2 = "0"
            
            ValeurChoisieTab = 229
            PGN_Tab = "Ex. Code:229 -> Garmin"
        End If
                   
     Case 127258
        z = CDbl("&H" & a6 & a5) * 0.0001 * 180 / pi
        ValeurChoisie = z
        PGN1 = "Variation magnétique"
        Echelle = "Deg"
        FormatAffichage = "0.0000"
           
      Case 130578
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "Détail Vitesse du bateau (" & ValeurChoisie & ")"
        Echelle = "N°"
        FormatAffichage = "0"
   
        If ValeurChoisie = 0 Then
            z = CDbl("&H" & a4 & a3) * 0.001 * 180 / pi
            If z < 255 Then ValeurChoisie2 = z Else ValeurChoisie2 = ""
            PGN2 = "Vitesse Surface"
            Echelle2 = "Nds"
            FormatAffichage2 = "0.00"
            
            z = CDbl("&H" & a8 & a7) * 0.001 * 180 / pi
            If z < 255 Then ValeurChoisie3 = z Else ValeurChoisie3 = ""
            PGN3 = "Vitesse Fond SOG"
            Echelle3 = "Nds"
            FormatAffichage3 = "0.00"
        Else
            ValeurChoisie3 = ""
            z = ""
            ValeurChoisie2 = z
            PGN2 = "Autres Vitesses"
            Echelle2 = ""
        End If
        
          
     Case 126464            '======= Tableau(N°octet, 1-> Signifie le PGN 126464)
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "Liste des PGN: (" & ValeurChoisie & ")"
        Echelle = "N°"
        FormatAffichage = "0"
         
        'Affiche les données des lignes modulo 3, contenant les PGN émissent
        If (ValeurChoisie + 3) Mod 3 = 0 Then 'ValeurChoisie = 0 Then
            z = CDbl("&H" & a6 & a5 & a4)
            If Not (z < 59392 Or z > 130944) Then
                ValeurChoisie3 = z
                PGN3 = "N° du PGN"
                Echelle3 = "PGN"
                FormatAffichage3 = "0"
                Memoire(MEMOIRE_PGN_a7, PGN_126464) = a7
                Memoire(MEMOIRE_PGN_a8, PGN_126464) = a8
            End If
         ElseIf (ValeurChoisie + 2) Mod 3 = 0 Then
            z = CDbl("&H" & a2 & Memoire(MEMOIRE_PGN_a8, PGN_126464) & Memoire(MEMOIRE_PGN_a7, PGN_126464))
            If Not (z < 59392 Or z > 130944) Then
                ValeurChoisie2 = z
                PGN2 = "N° du PGN"
                Echelle2 = "PGN"
                FormatAffichage2 = "0"
            End If
            
            z = CDbl("&H" & a5 & a4 & a3)
            If Not (z < 59392 Or z > 130944) Then
                ValeurChoisie3 = z
                PGN3 = "N° du PGN"
                Echelle3 = "PGN"
                FormatAffichage3 = "0"
            End If
            
            z = CDbl("&H" & a8 & a7 & a6)
            If Not (z < 59392 Or z > 130944) Then
                ValeurChoisieTab = z
                PGN_Tab = "Signification 4 PGN: " & ValeurChoisieTab
                'EchelleTab = "PGN"
                FormatAffichageTab = "0"
            End If
            
         ElseIf (z + 1) Mod 3 = 0 Then
            z = CDbl("&H" & a4 & a3 & a2)
            If Not (z < 59392 Or z > 130944) Then
                ValeurChoisie2 = z
                PGN2 = "N° du PGN"
                Echelle2 = "PGN"
                FormatAffichage2 = "0"
            End If
            
            z = CDbl("&H" & a7 & a6 & a5)
            If Not (z < 59392 Or z > 130944) Then
                ValeurChoisie3 = z
                PGN3 = "N° du PGN"
                Echelle3 = "PGN"
                FormatAffichage3 = "0"
            End If
            
        End If
        
         
      Case 129284
        z = CDbl("&H" & a1) And CDbl("&h1F")
        ValeurChoisie = z
        PGN1 = "Données de navigation "
        Echelle = "N°"
        FormatAffichage = "0"
        
        If ValeurChoisie = 4 Then
            'z = CDbl("&H" & a1) And CDbl("&h1F")
            ValeurChoisie2 = ""
            PGN2 = "Lattitude Waypoint"
            Echelle2 = "Deg"
            FormatAffichage2 = "0"
           
           'z = CDbl("&H" & a1) And CDbl("&h1F")
            ValeurChoisie3 = ""
            PGN3 = "Longitute Waypoint"
            Echelle3 = "Deg"
            FormatAffichage3 = "0"
        End If
       
    Case 130314
        z = CDbl("&H" & a7 & a6 & a5 & a4) * 0.001
        ValeurChoisie = z
        PGN1 = "Préssion Réelle"
        Echelle = "mBar"
        FormatAffichage = "0.00"
        
        z = CDbl("&H" & a3)
        ValeurChoisieTab = z
        PGN_Tab = "Table(" & ValeurChoisieTab & ")-> " & cTable.PRESSION(ValeurChoisieTab) 'Ecrit le texte correspondant au numéro dans la colonne "Issu des tables"

    Case 127505
        z = CDbl("&H" & a3 & a2) * 0.004
        ValeurChoisie = z
        If CDbl("&H" & a3 & a2) > 32000 Then
            ValeurChoisie2 = ""
            PGN1 = "Niveau"
        Else
            PGN1 = "Niveau du réservoir"
        End If
        Echelle = "%"
        FormatAffichage = "0.00"
    
        z = CDbl("&H" & a5 & a4) * 0.1
        ValeurChoisie2 = z
        If CDbl("&H" & a5 & a4) > 32000 Then
            ValeurChoisie2 = ""
            PGN2 = "Capacité du réservoir"
        Else
            PGN2 = "Capacité du réservoir"
        End If
        Echelle2 = "L"
        FormatAffichage2 = "0.00"
        
        z = CDbl("&H" & a1) And CDbl("&hF0")
        ValeurChoisieTab = WorksheetFunction.Bitrshift(Arg1:=z, Arg2:=4)
        PGN_Tab = "Table(" & ValeurChoisieTab & ")-> " & cTable.Tank(ValeurChoisieTab)   'Ecrit le texte correspondant au numéro dans la colonne "Issu des tables"

    Case Else
        
        PGN1 = "Inconnu"
        
End Select
    
    
'------------------------------- Ecrit les résultats ----------------------------------------
'Ne pas afficher les résultats sur "Huahine" si on est en "Exemple"
If User_Form_PGN.Appercu = True Then GoTo PasAfficher
    'Efface la ligne
    Worksheets(FEUIL_NMEA).Range(COLONNE_DEBUT_RESULTAT & ligne & ":" & COLONNE_FIN_RESULTAT & ligne).ClearContents
    
    'Affiche les adresses sources et destinationsvrnant de l'ID du bus CAN dans les colonnes AY et AZ
    Worksheets(FEUIL_NMEA).Range(COLONNE_SOURCE & ligne).Value = cID.SourceID(ligne)
    Worksheets(FEUIL_NMEA).Range(COLONNE_DESTINATION & ligne).Value = cID.DestinationID(ligne)

    'Affcihe la valeur du PGN sur la colonne AX
    Worksheets(FEUIL_NMEA).Range(COLONNE_PGN_ID & ligne).Value = cID.Pgn(ligne)

    'Affiche les définitions et isuses de la table
    If ValeurChoisie <> PAS_INFO Then                                      '1ere définition
        Worksheets(FEUIL_NMEA).Range(COLONNE_DEF_1 & ligne).Value = PGN1                                         'Signification
        Worksheets(FEUIL_NMEA).Range(COLONNE_VAL_1 & ligne).Value = Format(ValeurChoisie, FormatAffichage)       'Valeur
        Worksheets(FEUIL_NMEA).Range(COLONNE_ECH_1 & ligne).Value = Echelle                                     'Echelle
    End If
    If ValeurChoisie2 <> PAS_INFO Then                                     '2ème définition
        Worksheets(FEUIL_NMEA).Range(COLONNE_DEF_2 & ligne).Value = PGN2                                         'Signification
        Worksheets(FEUIL_NMEA).Range(COLONNE_VAL_2 & ligne).Value = Format(ValeurChoisie2, FormatAffichage2)      'Valeur
        Worksheets(FEUIL_NMEA).Range(COLONNE_ECH_2 & ligne).Value = Echelle2                                     'Echelle
    End If
    If ValeurChoisie3 <> PAS_INFO Then                                     '3ème définition
        Worksheets(FEUIL_NMEA).Range(COLONNE_DEF_3 & ligne).Value = PGN3                                        'Signification
        Worksheets(FEUIL_NMEA).Range(COLONNE_VAL_3 & ligne).Value = Format(ValeurChoisie3, FormatAffichage3)     'Valeur
        Worksheets(FEUIL_NMEA).Range(COLONNE_ECH_3 & ligne).Value = Echelle3                                    'Echelle
    End If
    If ValeurChoisieTab <> PAS_INFO Then                                   'Définition dans une table
        Worksheets(FEUIL_NMEA).Range(COLONNE_TAB & ligne).Value = PGN_Tab                                     'Signification
    End If
   
    'Affiche les émetteurs et récepteurs défini par l'utilisateur dans les colonnes AD et AE, détaillés  dans les colonnes BL et BM
    Worksheets(FEUIL_NMEA).Range(COLONNE_SOURCE_DETAIL & ligne).Value = "(" & cID.SourceID(ligne) & ")  " & IIf(cTable.Adresses(cID.SourceID(ligne)) = "", "Inconnu", cTable.Adresses(cID.SourceID(ligne)))
    Worksheets(FEUIL_NMEA).Range(COLONNE_DESTINATION_DETAIL & ligne).Value = IIf(cID.DestinationID(ligne) < 240, " (" & cID.DestinationID(ligne) & ") ", "") & IIf(cTable.Adresses(cID.DestinationID(ligne)) = "", "Inconnu", cTable.Adresses(cID.DestinationID(ligne)))
'---------------------------------- Fin de l'écriture ----------------------------------------------
   

PasAfficher:

'--------- Valable pour l'exemple --------
'limite la valeur de manière à ne pas avoir de dépassement
If ii > 100000 Then ii = 100000

'Affiche les valeurs en fonction des numéros de lignes dans l'exemple
If ActiveSheet.Name = FEUIL_EXEMPLE Then
    cExemple.Affiche (ligne)
End If

'Affiche les MMSI
If Not User_Form_PGN.Appercu Then
    cMMSI.Affiche_MMSI (ligne)
End If


'Quanq on arrive à la fin on recommence sur buffer tournant
'If ligne = Fin And User_Form_PGN.Buffer = True Then
 '   ligne = Debut
'End If
'-------------- Fin de l'exemple ------------


Next ligne
'---------------------------------------------------Retour pour la ligne suivante ------------------------------------------------------


'--------------------------------------------- Programme terminé ------------------------------------
Fin:
Tps2 = Timer                            'Mesure le 2ème temps


'N'affiche pas le voyant en cours et affiche le voyant en stop
User_Form_PGN.Feux.Visible = False
User_Form_PGN.FeuxStop.Visible = True

'Envoi le message terminé
If User_Form_PGN.Appercu = False Then
    MsgBox "Vos données sont rafraichies sur les trames " & aa & " à " & Format(bb, " #,##0 ") & vbCr & "Sur la feuille """ & FEUIL_NMEA & """" & vbCr & vbCr & "ATTENTION: les autres trames restent inchangées." & vbCr & "Les feuilles Exemple et MMSI sont aussi rafraichies", vbOKOnly + vbInformation, "Générer"
Else
    MsgBox "Vous avez vue sur """ & FEUIL_EXEMPLE & """ l'affichage des trames " & aa & " à " & Format(bb, " #,##0 ") & vbCr & "En " & Format(Tps2 - Tps, "0.00") & " secondes" & vbCr & vbCr & "ATTENTION: Les trames ne sont pas mises à jours sur """ & FEUIL_NMEA & """" & vbCr & "Ni sur l'ecran MMSI", vbOKOnly + vbInformation, "Générer"
End If

'Affiche le temps passé pour la génération
If ActiveSheet.Name <> FEUIL_TEMPS Then
    Worksheets(FEUIL_EXEMPLE).Range("D1").Value = Format(Tps2 - Tps, " 0.00 ")
Else
    Worksheets(FEUIL_TEMPS).Unprotect
    If User_Form_PGN.Appercu Then
        Worksheets(FEUIL_TEMPS).Range("E9").Value = Format(Tps2 - Tps, "0.00") & " Secondes sur " & Format(ligne - 1, " #,##0 ") & " lignes"
    Else
        ActiveSheet.Range("E10") = Format(Tps2 - Tps, " 0.00 ") & " Secondes sur " & Format(ligne - 1, " #,##0 ") & " lignes"
    End If
End If

'Ecrit le nombre de ligne généré
If User_Form_PGN.Appercu Then User_Form_PGN.Fin.Caption = iii



GoTo FinFin

'Fin sur arrêt
FinA:

'Affiche le message si arret demandé ou si on est pas en "Apperçu"
If Not StopA Or User_Form_PGN.Appercu = False Then
    If MsgBox("Voulez vous arrêter ?" & vbCr & "Vos données sont rafraichies sur " & Format(Debut + iii - 1, " #,##0 ") & " trames uniquement" & vbCr & "Sur la feuille """ & FEUIL_NMEA & """" & vbCr & vbCr & "Les autres trames restent inchangées", vbYesNo + vbQuestion + vbDefaultButton2, "Générer") <> vbYes Then
        StopA = False
        GoTo Debut
    Else
        GoTo FinFin
    End If
Else    'On est en "Apperçu"
    If MsgBox("Voulez vous arrêter l'" & FEUIL_EXEMPLE & " ?" & vbCr & "Vos données sont rafraichies sur " & Format(Debut + iii - 1, " #,##0 ") & " trames uniquement", vbYesNo + vbQuestion + vbDefaultButton2, "Générer") <> vbYes Then
        StopA = False
        GoTo Debut
    Else
        GoTo FinFin
    End If
End If

'Fin pout tous
FinFin:

'Remet en place les informations
Application.Cursor = xlDefault

User_Form_PGN.FeuxStop.Visible = True                   'LED
User_Form_PGN.CommandButton1.Enabled = True     'Bouton Rafraichir
User_Form_PGN.Appercu.Enabled = True                  'Case à coché
User_Form_PGN.CommandButton3.Enabled = False    'Bouton Arrêté non accéssible
User_Form_PGN.CommandButton4.Enabled = True     'Bouton RAZ accéssible


'Protége les feuilles
Worksheets(FEUIL_NMEA).Protect
Worksheets(FEUIL_EXEMPLE).Protect
Worksheets(FEUIL_TEMPS).Protect

'Remet en automatique les calculs
Application.Calculation = xlCalculationAutomatic

EnCours = False

User_Form_PGN.Patienter.Caption = "Terminé ..." & vbCr

'Met en Normal si la fenêtre est minimisé
If Application.WindowState = xlMinimized Then
    Application.WindowState = xlNormal
End If

'Fin globale
FINFINFIN:

Exit Sub

Erreur:
    MsgBox "Erreur à la ligne " & ligne & " sur " & Worksheets(1).Name & """" & vbCr & "Vérifier que vous êtes sur la bonne feuille", vbCritical + vbOKOnly, "ERREUR"
GoTo Fin

End Sub



'Fonction trouvé sur Internet
Public Function FeuilleExiste(FeuilleAVerifier As String) As Boolean
'fonction qui vérifie si la "FeuilleAVerifier" existe dans le Classeur actif
 
    On Error Resume Next
    Sheets(FeuilleAVerifier).Name = Sheets(FeuilleAVerifier).Name
    FeuilleExiste = (Err.Number = 0)
End Function
