Attribute VB_Name = "Communication"
Option Explicit

'***************************************************************
'
' Programme de réception des données arrivant du bus CAN
' Il fonctionne bien c'est un port venant de l'USB qui pour l'instant fonctionne
'
'***************************************************************
Sub Temps_Reel()

Dim o As Long, c As Long, r As Long  'Variables pour Ouvrir, Fermer et Lire
Dim i As Long                        'Compte le de réception
Dim recu As String  'Caractères reçus

'Ouvre le port com
o = CommOpen(A_PORT_ID, A_COM, "baud=" & A_VITESSE & " parity=N data=8 stop=1")
   
If o = 0 Then 'Port ouvert
            
    Worksheets(FEUIL_EXEMPLE).Unprotect 'Déprotège la feuille
    
    'Tant que la case à cocher est TRUE
    Do While User_Form_PGN.BufferA
        DoEvents
                
        r = CommRead(A_PORT_ID, recu, A_NOMBRE_CARACTERE)
        'Attend la trame ????
                
        'S'il y a au moins des caractères
        If r > 0 Then
            Analyse (recu) 'Analyse ce qui est reçu
            i = i + 1
            If i Mod 500 Then DoEvents
        End If
    Loop
    
    'Ferme le fichier
    Close #1

    Worksheets(FEUIL_EXEMPLE).Protect 'Protège la feuille

Else
    Debug.Print "Port COM n'est pas ouvert"
    MsgBox "Le port COM n'est pas ouvert" & vbCr & "Veuillez vérifier que vous êtes bien raccordé", vbInformation, "TEMPS REEL"
    User_Form_PGN.BufferA = False
End If

'Ferme le port com
c = CommClose(A_PORT_ID)

'Ferme le fichier
Close #1

Debug.Print "o-> " & o
Debug.Print "c-> " & c

End Sub

'----------------- Analyse les trammes reçus ---------------
'Analyse Total_Recu qui peut comporter plusieurs octets et plusieurs trames
'La structrue des trames CAN émisent du l'Arduino
'           .ID;Long:Data,data,data,data ... fini par "?"
'           sur le dernier octet il n'y a pas de ","
'-----------------------------------------------------------
Sub Analyse(Total_Recu As String)

Dim c As Long                   'pour fermer le port
Dim Pointeur As Long            'Pointeur du buffer sortant vers PGN_DecodeA(Trame As String)
Dim N_Buffer As Long            'Compte le nombre chaines misent dans le buffer en entrée
Dim Buffer(20000) As String     'Chaines contenants tous les carcatères reçu
Dim Debut_Trame As Integer      'Debut de la trame
Dim Partie_Gauche As String     'Récupère la partie gauche de la trame, on n'en a pas besoin
Dim Partie_Droite As String     'Récupère la partie droite de la trame
Dim Fin_Trame As Integer        'Longueur de la trame complète
Dim Longueur_Trame As Integer   'Mémorise la valeur de Longueur_Trame
Dim Chaine_Total_Tous As String 'Récupère le concaténation du reste de l'ancienne trame avec la nouvelle
Dim Chaine_Attente As String    'Cumule les octets reçu jusqu'a il y est un début de trame
Dim Pointeur_Trame As Long      'Ce point de déplace sur la trame
Dim i As Long                   'incrément le nombre de tour, passé une certaine valeur, arrête le programme

' a faire en Static car elle contient la fin d'une trame, à cumulé avec la nouvelle
Static Buffer_Restant As String 'on ne l'initialse pas car elle est écrite plus loin

'Initialise les valeurs
Pointeur = 0        'Initialyse le pointeur pour les sorties des trames
Fin_Trame = 0       'Initialyse la fin de la trame
Longueur_Trame = 0  'Initialyse la logueur de la trame
Pointeur_Trame = 0  'Initialyse le pointeur de la trame reçu
i = 0               'Initialise le compteur de boucles

    '------------------------------------------
    'Attend qu'il y ait un "." (Début de trame)
    '------------------------------------------
        'Cumule Chaine_Attente avec Total_Recu, attend les nouveaux octets.
        Chaine_Attente = Chaine_Attente & Total_Recu
        
        'On commence à cherche le début par un "."
        Debut_Trame = InStr(Longueur_Trame + 1, Chaine_Attente, ".")
        
        'Tant qu'on a pas de point on attends les autres octets.
        If Debut_Trame = 0 Then
            Exit Sub
        End If
        
        'On a trouvé un début de trame et on analyse tous.
        Chaine_Total_Tous = Buffer_Restant & Chaine_Attente
    
    'La partie droite commence par un ".", ce qu'on vient de trouvé.
    'La partie gauche n'est pas utile car c'est le début de la transmission.
    If Len(Chaine_Total_Tous) > Len(Partie_Droite) Then
        Partie_Droite = Chaine_Total_Tous
    Else
        Partie_Droite = Right(Chaine_Total_Tous, Len(Chaine_Total_Tous) - Debut_Trame + 1)
    End If
    
    'On tourne en permance, jusqu'a avoir trouvé une fin de trame, on sort de la boucle si on n'a pas trouvé de fin de trame.
    Do While True
    
        ' Cherche la fin de la trame et trouve le nombre d'octets jusqu'a la fin de la trame ("?").
        Fin_Trame = InStr(Fin_Trame + 1, Partie_Droite, "?")
            
        'Tant qu'on n'a pas trouvé on attends les octets suivants.
        'il s'agit de la fin des octets reçus qui vont se trouver dans Buffer_Restant
        If Fin_Trame = 0 Then
            Exit Do
        End If
            
        'Longueur de la nouvelle trame
        Longueur_Trame = Fin_Trame - Pointeur_Trame
            
        'On sort tant qu'on a pas de fin de trame "?", on attend la autres octets
        If Fin_Trame - Pointeur_Trame = 0 Then Exit Do
            
        'Si  on trouve la fin de trame, on la met dans le buffer
        If Fin_Trame - Pointeur_Trame > 0 Then
            'La trame bus CAN compléte se trouve dans la partie droite
            Buffer(N_Buffer) = Mid(Partie_Droite, Pointeur_Trame + 1, Longueur_Trame)
            N_Buffer = N_Buffer + 1
                
            'Le restant sur trouve dans ce qu'il reste des octets reçus sur la partie droite
            Buffer_Restant = Right(Partie_Droite, Len(Partie_Droite) - Fin_Trame)
                               
            'On trouve le nouveau point d'origine
            Pointeur_Trame = Pointeur_Trame + Longueur_Trame

            If Buffer_Restant = "" Then Exit Do
        Else
            'On laisse tomber car c'est un début, on laisse tomber la partie gauche
            'Si on ne trouve pas la fin de la trame on va à la suite.. Voir ci dessus "if Fin_Trame = 0 Then Exit Do"
        End If
            
        'Contrôle s'il y a un problème (Ex. pas de ".") 'Cest falcutatif
        i = i + 1
        If i > 30000 Then
            'Ferme le port com
            c = CommClose(A_PORT_ID)
            User_Form_PGN.BufferA = False
            MsgBox "Il y a trop de trames" & vbCr & "Ou les trames sont incorrects" & vbCr & vbCr & "Nous avons fermé le port" & vbCr & Buffer(N_Buffer - 1), vbCritical, "COMMUNICATION"
            Exit Sub
        End If
        
    Loop
             
'On a fini la trame et on mémorise dans Buffer_Restant ce qu'il reste

'Envoi les trames recu vers PGN_Decode(Trame As String) dans l'ordre ou elles sont arrivées
Do While Pointeur < N_Buffer

    'Verifie la coherence de la trame, normalement ce n'est pas obligatoire mais j'ai eu quelques problèmes.
    If InStr(1, Buffer(Pointeur), ".") = 1 And InStr(2, Buffer(Pointeur), ".") = 0 Then
        'on enregistre
        'If User_Form_COM.CheckBox1 Then
            Print #1, Buffer(Pointeur)
        'End If
    
        'On décode
        PGN_DecodeA (Buffer(Pointeur))
    End If
    
'On incrémente le pointeur
Pointeur = Pointeur + 1
    
Loop

End Sub 'On attend les prochains octets
'--------------------------------------------------------------
'                    Fin analyse
'--------------------------------------------------------------

'TEST ------------>>>
Sub t()
Dim i As Integer

'Analyse ("")


Open "d:\alain\tito.txt" For Append As #1

Worksheets(FEUIL_EXEMPLE).Unprotect
        'initiale la valeur de Pi
        pi = Application.WorksheetFunction.pi() 'Valeur de pi
        
'dure environ 1 secondes pour 8 trames envoyé 500 fois
For i = 1 To 500
Analyse (",3,43,71,FF,FF,FF?;8:81,FF,7F,FF,7F,FF,7F,FF?.95FD0810;8:0,0,2,C1,70,FF,FF,FF?.89FD0203;8:CC,3E,1,E8,19,FA,FF,FF?.95FD0810;8:0,0,3,43,71,FF,FF,FF?.89FD0203;8:9,8,7,6,5,4,3,1?")
'Esai avec  130306 et 130012
Analyse ("fe?8.AD9")  'Ca fonctionne
Analyse ("ppiu.")  'N'importe quoi, ça fonctionne bien, il a fallu controlé sur erreur le PGN dans cIDa.Pgn(ID as String)

Analyse (".15fd0810;8:0,6,7,8,9,FF,FF,FF?.15fd0810;8:0,0,2,7D,1,2,3,4?.15fd0810;")
Analyse ("7:10,7D,6,7,8,9?.15fd0819;")
Analyse ("8:0,0,2,7D,73,FF,FF,FF?.09fd0203;8:a7,7c,01,fc,a4,fa,F0,F1?.15fd")

Analyse ("fe8AD9")  'Ca fonctionne
Analyse ("ppiu")  'N'importe quoi, ça fonctionne
Next i
Worksheets(FEUIL_EXEMPLE).Protect
Close #1
End Sub

