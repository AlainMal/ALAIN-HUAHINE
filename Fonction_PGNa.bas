Attribute VB_Name = "Fonction_PGNa"
Option Explicit

'--------------------- Fonction de décodage des trames reçus en temps réel ------------------------
' La structrue des trames reçus en temps réel: ID;Long:Data,data,data,data ... CR-LF.
' On vérifie sa structure et on met à jour l'ID, nombre d'octets et les octets.
' On initialise les valeurs ensuite on fait un Select Case pour le PGN comme la procédure PGN_Decode() (copie-coller).
' Si on veut on enregistre sur un fichier .txt qui pourra être lu plus tard.
' On met à jour l'Exemple et les MMSI, on duplique les classes cID, cMMSI et cExemple car il s'agit d'une infomation
' en temps réel qui ont en entrée la valeur du PGN et non la ligne.
' les fonctions portent le même nom et leurs classe se nomment avec "a" pour dernier caractère.
'-------------------------------------------------------------------------------------------------
Sub PGN_DecodeA(Trame As String)
Dim Pgn As String

Dim z As String                                                                             'Valeur mesuurée
Dim PGN1 As Variant, PGN2 As Variant, PGN3 As Variant                                        'Le texte du PGN
Dim Memoire(NB_OCTETS, Nb_PGN) As String                                                     'Tableau pour les chevauchements
Dim aFin As Boolean                                                                         'Indique la fin de la génération
Dim ID As String
Dim Nbr_Str As String  'Nombre d'octet
Dim N_Trouve As Integer 'Emplacement du trouvé ";   " début de trame
Dim N_Trouve1 As Integer 'Emplacement du trouvé ";  "
Dim N_Trouve2 As Integer 'Emplacement du trouvé ";  "
Dim N_TrouveA1 As Integer 'Emplacement du trouvé ", " octet 1
Dim N_TrouveA2 As Integer 'Emplacement du trouvé ", " octet 2
Dim N_TrouveA3 As Integer 'Emplacement du trouvé ", " octet 3
Dim N_TrouveA4 As Integer 'Emplacement du trouvé ", " octet 4
Dim N_TrouveA5 As Integer 'Emplacement du trouvé ", " octet 5
Dim N_TrouveA6 As Integer 'Emplacement du trouvé ", " octet 6
Dim N_TrouveA7 As Integer 'Emplacement du trouvé ", " octet 7
Dim N_TrouveA8 As Integer 'Emplacement du trouvé ","" octet 8
Dim Debut_Trame As String

Static i As Integer

'Defini les octets du bus CAN
Dim a1 As Variant, a2 As Variant, a3 As Variant, a4 As Variant, a5 As Variant, a6 As Variant, a7 As Variant, a8 As Variant

'n'est pas utile au cas ou ?
N_Trouve = InStr(1, Trame, ".")
If N_Trouve > 0 Then
    Debut_Trame = Mid(Trame, 1, N_Trouve - 1)
Else
   ' GoTo Fin
End If
'N_Trouve = 1

'Met à jour l'ID issu de la trame, commence au début et fini sur un ";", génère le PGN
N_Trouve1 = InStr(1, Trame, ";")
If N_Trouve1 > 0 Then
    ID = Mid(Trame, N_Trouve + 1, N_Trouve1 - N_Trouve - 1)
    If ID = "" Then GoTo Fin
    If ID <> "" Then
        Pgn = cIDa.Pgn(ID)
        If Pgn = "" Then GoTo Fin
    Else
        GoTo Fin
    End If
Else
    GoTo Fin
End If

'Met à jour le nombre d'octet, se termine pas un ":"
N_Trouve2 = InStr(N_Trouve1, Trame, ":")
If N_Trouve2 > 0 Then
    Nbr_Str = Mid(Trame, N_Trouve1 + 1, N_Trouve2 - N_Trouve1 - 1)
    If Nbr_Str = "" Then GoTo Fin
    i = 1
Else
    GoTo Fin
End If

'Met à jour les octets, commence par un ":" et se termine par un "," ou fin de trame "?"
If i = 1 And CInt(Nbr_Str) > 0 Then
    'Si le nombre d'octet est = 1 alors on cherhe la fin de trame (?)
    If Nbr_Str = 1 Then
        N_TrouveA1 = InStr(N_Trouve2 + 1, Trame, "?")
    Else
        N_TrouveA1 = InStr(N_Trouve2 + 1, Trame, ",")
    End If
    If N_TrouveA1 > 0 Then
        a1 = Mid(Trame, N_Trouve2 + 1, N_TrouveA1 - N_Trouve2 - 1)
    Else
        a1 = Mid(Trame, N_Trouve2 + 1, Len(Trame))
    End If
    If Len(a1) > 2 Then GoTo Fin
    i = i + 1  'Incrément le nombre d'octets reçu
End If
If i = 2 And CInt(Nbr_Str) > 1 Then
    If Nbr_Str = 2 Then
        N_TrouveA2 = InStr(N_TrouveA1 + 1, Trame, "?")
    Else
        N_TrouveA2 = InStr(N_TrouveA1 + 1, Trame, ",")
    End If
    If N_TrouveA2 > 0 Then
        a2 = Mid(Trame, N_TrouveA1 + 1, N_TrouveA2 - N_TrouveA1 - 1)
    Else
        a2 = Mid(Trame, N_TrouveA1 + 1, Len(Trame))
    End If
    If Len(a2) > 2 Then GoTo Fin
    i = i + 1
End If
If i = 3 And CInt(Nbr_Str) > 2 Then
    If Nbr_Str = 3 Then
        N_TrouveA3 = InStr(N_TrouveA2 + 1, Trame, "?")
    Else
        N_TrouveA3 = InStr(N_TrouveA2 + 1, Trame, ",")
    End If
    If N_TrouveA3 > 0 Then
        a3 = Mid(Trame, N_TrouveA2 + 1, N_TrouveA3 - N_TrouveA2 - 1)
    Else
        a3 = Mid(Trame, N_TrouveA2 + 1, Len(Trame))
    End If
    If Len(a3) > 2 Then GoTo Fin
    i = i + 1
End If
If i = 4 And CInt(Nbr_Str) > 3 Then
    If Nbr_Str = 4 Then
        N_TrouveA4 = InStr(N_TrouveA3 + 1, Trame, "?")
    Else
        N_TrouveA4 = InStr(N_TrouveA3 + 1, Trame, ",")
    End If
    If N_TrouveA4 > 0 Then
        a4 = Mid(Trame, N_TrouveA3 + 1, N_TrouveA4 - N_TrouveA3 - 1)
    Else
        a4 = Mid(Trame, N_TrouveA3 + 1, Len(Trame))
    End If
    If Len(a4) > 2 Then GoTo Fin
    i = i + 1
End If
If i = 5 And CInt(Nbr_Str) > 4 Then
    If Nbr_Str = 5 Then
        N_TrouveA5 = InStr(N_TrouveA4 + 1, Trame, "?")
    Else
        N_TrouveA5 = InStr(N_TrouveA4 + 1, Trame, ",")
    End If
    If N_TrouveA5 > 0 Then
        a5 = Mid(Trame, N_TrouveA4 + 1, N_TrouveA5 - N_TrouveA4 - 1)
    Else
        a5 = Mid(Trame, N_TrouveA4 + 1, Len(Trame))
    End If
    If Len(a5) > 2 Then GoTo Fin
    i = i + 1
End If
If i = 6 And CInt(Nbr_Str) > 5 Then
    If Nbr_Str = 6 Then
        N_TrouveA6 = InStr(N_TrouveA5 + 1, Trame, "?")
    Else
        N_TrouveA6 = InStr(N_TrouveA5 + 1, Trame, ",")
    End If
    If N_TrouveA6 > 0 Then
        a6 = Mid(Trame, N_TrouveA5 + 1, N_TrouveA6 - N_TrouveA5 - 1)
    Else
        a6 = Mid(Trame, N_TrouveA5 + 1, Len(Trame))
    End If
    If Len(a6) > 2 Then GoTo Fin
    i = i + 1
End If
If i = 7 And CInt(Nbr_Str) > 6 Then
    If Nbr_Str = 7 Then
        N_TrouveA7 = InStr(N_TrouveA6 + 1, Trame, "?")
    Else
        N_TrouveA7 = InStr(N_TrouveA6 + 1, Trame, ",")
    End If
    If N_TrouveA7 > 0 Then
        a7 = Mid(Trame, N_TrouveA6 + 1, N_TrouveA7 - N_TrouveA6 - 1)
    Else
        a7 = Mid(Trame, N_TrouveA6 + 1, Len(Trame))
    End If
    If Len(a7) > 2 Then GoTo Fin
    i = i + 1
End If
If i = 8 And CInt(Nbr_Str) > 7 Then
    If Nbr_Str = 8 Then
        N_TrouveA8 = InStr(N_TrouveA7 + 1, Trame, "?")
    Else
        N_TrouveA8 = InStr(N_TrouveA7 + 1, Trame, ",")
    End If
    If N_TrouveA8 > 0 Then
        a8 = Mid(Trame, N_TrouveA7 + 1, N_TrouveA8 - N_TrouveA7 - 1)
    Else
        a8 = Mid(Trame, N_TrouveA7 + 1, Len(Trame))
    End If
    If Len(a8) > 2 Then GoTo Fin
    'i = i + 1
End If


Select Case Pgn 'cID.PGN(ligne)

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
        z = CStr(CDbl("&H" & a5 & a4) * 0.01 - 273.15)
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
        PGN1 = "Reclame PGN:" & ValeurChoisie   ' & " à l'@ " & cID.DestinationID(ligne)
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

'Rafraichi l'Exemple
If User_Form_COM.Check_Exemple Then
    cExemplea.Affiche (Pgn)
End If

'Rafraichi le MMSI
If User_Form_COM.Check_MMSI Then
    cMMSIa.Affiche_MMSI (Pgn)
End If

Fin:
End Sub

