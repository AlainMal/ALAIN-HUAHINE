Attribute VB_Name = "Communication"
Option Explicit

'***************************************************************
'
' Programme de réception des données arrivant du bus CAN
' C'est un port venant de l'USB qui pour l'instant fonctionne
'
'***************************************************************
Sub Temps_Reel()

Dim o As Long, C As Long, r As Long  'Variables pour Ouvrir, Fermer et Lire
Dim i As Long                        'Compte pour faire un DoEvents
Dim Recu As String  'Caractères reçus

'Ouvre le port com
o = CommOpen(A_PORT_ID, A_COM, "baud=" & A_VITESSE & " parity=N data=8 stop=1")
   
If o = 0 Then 'Port ouvert
            
    Worksheets(FEUIL_EXEMPLE).Unprotect 'Déprotège la feuille
    
    'Tant que la case à cocher est TRUE
    Do While User_Form_PGN.BufferA
        DoEvents
                
        r = CommRead(A_PORT_ID, Recu, A_NOMBRE_CARACTERE)
        'Attend la trame ????
                
        'S'il y a au moins des caractères
        If r > 0 Then
            Analyse (Recu) 'Analyse ce qui est reçu
            i = i + 1
            If i Mod 500 Then DoEvents
        End If
    Loop
    
    'On est sortie de la boucle
    Worksheets(FEUIL_EXEMPLE).Protect 'Protège la feuille

Else
    Debug.Print "Port COM n'est pas ouvert"
    MsgBox "Le port COM n'est pas ouvert" & vbCr & "Veuillez vérifier que vous êtes bien raccordé", vbInformation, "TEMPS REEL"
    User_Form_PGN.BufferA = False
    User_Form_COM.Show
End If

'Ferme le port com
C = CommClose(A_PORT_ID)

'Ferme le fichier
Close #1

Debug.Print "o-> " & o
Debug.Print "c-> " & C

End Sub

'----------------- Analyse les trammes reçus ---------------
'Analyse Total_Recu qui peut comporter plusieurs octets et plusieurs trames
'La structrue des trames CAN émisent du l'Arduino (Sans espace)
'           . ID ; Long : Data , data , data , data ... fini par "?"
'           sur le dernier octet il n'y a pas de ","
'-----------------------------------------------------------
Sub Analyse(Total_Recu As String)

Dim C As Long                   'pour fermer le port
Dim Pointeur As Long            'Pointeur du buffer sortant vers PGN_DecodeA(Trame As String)
Dim N_Buffer As Long            'Compte le nombre chaines misent dans le buffer en entrée
Dim Buffer(2000) As String      'Chaines contenants tous les carcatères reçu
Dim Debut_Trame As Integer      'Pointe sur début de la trame
Dim Partie_Gauche As String     'Récupère la partie gauche de la trame, on n'en a pas besoin
Dim Partie_Droite As String     'Récupère la partie droite de la trame
Dim Pointeur_Fin_Trame As Integer        'Pointe sur la fin de la trame
Dim Longueur_Trame As Integer   'Mémorise la longueur de la trame
Dim Chaine_Total As String      'Récupère le concaténation du reste de l'ancienne trame avec la nouvelle
Dim Chaine_Attente As String    'Cumule les octets reçu jusqu'a il y est un début de trame
Dim Pointeur_Debut_Trame As Long      'Ce point se déplace sur la trame
Dim i As Long                   'incrément le nombre de tour, passé une certaine valeur, arrête le programme
Dim Buffer_Analyse As String    'Analyse la trame avant de l'envoyer au décodeur
Dim ID As String                'ID pour analyse
Dim Nbr_Octet   As Integer      'Nombre d'octet pour analyse

' a faire en Static car elle contient la fin d'une trame, à cumulé avec la nouvelle
Static Buffer_Restant As String 'on ne l'initialse pas car elle est écrite plus loin

'Initialise les valeurs
N_Buffer = 0        'Initialyse le pointeur d'entrée des trames
Pointeur = 0        'Initialyse le pointeur pour les sorties des trames
Pointeur_Fin_Trame = 0       'Initialyse la fin de la trame
Longueur_Trame = 0  'Initialyse la logueur de la trame
Pointeur_Debut_Trame = 0  'Initialyse le pointeur de la trame reçu
i = 0               'Initialise le compteur de boucles

    '------------------------------------------
    'Attend qu'il y ait un DEBUT_CAN (Début de trame)
    '------------------------------------------
        'Cumule Chaine_Attente avec Total_Recu, attend les nouveaux octets.
        Chaine_Attente = Chaine_Attente & Total_Recu
        
        'On commence à cherche le début par un DEBUT_CAN
        Debut_Trame = InStr(Longueur_Trame + 1, Buffer_Restant & Chaine_Attente, DEBUT_CAN)
        
        'Tant qu'on a pas de DEBUT_CAN on attends les autres octets.
        If Debut_Trame = 0 Then
            Exit Sub
        End If
        
        'On a trouvé un début de trame et on le  tous dans Chaine_Total
        Chaine_Total = Buffer_Restant & Chaine_Attente
    
    'La partie droite commence par un DEBUT_CAN, ce qu'on vient de trouvé.
    'La partie gauche n'est pas utile car c'est le début de la transmission (la première trame) aprés toutes les trames sont collées.
        Partie_Droite = Right(Chaine_Total, Len(Chaine_Total) - Debut_Trame + 1)
    
    'On a trouvé le début de la trame, maintenant on cherche la fin de trame, on sort de la boucle tant on n'a pas trouvé de fin de trame.
    Do While True
    
        ' Cherche la fin de la trame et trouve le nombre d'octets jusqu'a la fin de la trame ("?").
        Pointeur_Fin_Trame = InStr(Pointeur_Fin_Trame + 1, Partie_Droite, FIN_CAN)
            
        'Tant qu'on n'a pas trouvé on attends les octets suivants.
        If Pointeur_Fin_Trame = 0 Then
            Exit Do
        End If
            
        'Longueur de la nouvelle trame, Pointeur trame est égal à 0 pour commencer mais il se met à jour en fonction de la longueur de la trame.
        Longueur_Trame = Pointeur_Fin_Trame - Pointeur_Debut_Trame
            
        'La trame bus CAN compléte se trouve dans la partie droite
        Buffer_Analyse = Mid(Partie_Droite, Pointeur_Debut_Trame + 1, Longueur_Trame)
            
        'Analyse la trame CAN, ce n'est pas obligatoire car L'USB fait déjà un contrôle CRC. Mais c'est utile avec nos essais car on met n'importte quoi dans les trames.
        'On le fait pour le bus CAN, uniquement étendu.
        'On contrôle que l'ID et le nombre d'octets et on sait que ça commence par une DEBUT_CAN et fini pas un FIN_CAN, quand aux octets ils seront contrôlés dans le Decode_PGNa.
        ID = Mid(Buffer_Analyse, 2, 8)
        On Error Resume Next
        Err.Clear
        ID = CDbl("&H" & ID)
        Nbr_Octet = CDbl(Mid(Buffer_Analyse, 11, 1))
        If Err = 0 And Nbr_Octet <= 8 Then
            'On le met dans le buffer
            Buffer(N_Buffer) = Buffer_Analyse
            'On incrément le Numéro du buffer.
            N_Buffer = N_Buffer + 1
        End If
        On Error GoTo 0
            
        'Le restant sur trouve dans ce qu'il reste sur la partie droite.
        Buffer_Restant = Right(Partie_Droite, Len(Partie_Droite) - Pointeur_Fin_Trame)
                               
        'On trouve le nouveau point d'origine.
        Pointeur_Debut_Trame = Pointeur_Debut_Trame + Longueur_Trame
            
        'S'il n'y a rien dans buffer restant on sort de la boucle
        If Buffer_Restant = "" Then Exit Do
            
        'Contrôle s'il y a un problème (Ex. pas de DEBUT_CAN) 'Cest falcutatif
        i = i + 1
        If i > 3000 Then
            'Ferme le port com
            C = CommClose(A_PORT_ID)
            User_Form_PGN.BufferA = False
            MsgBox "Il y a trop de trames" & vbCr & "Ou les trames sont incorrects" & vbCr & vbCr & "Nous avons fermé le port" & vbCr & Buffer(N_Buffer - 1), vbCritical, "COMMUNICATION"
            Exit Sub
        End If
        
    Loop
             
'On a fini la trame et on mémorise dans Buffer_Restant ce qu'il reste
Pointeur = 0
'Envoi les trames recu vers PGN_Decode(Trame As String) dans l'ordre ou elles sont arrivées
Do While Pointeur < N_Buffer

        'on enregistre
        If User_Form_COM.CheckBox1 Then
            Print #1, Buffer(Pointeur)
        End If
    
        'On décode
        PGN_DecodeA (Buffer(Pointeur))
        
        'Ecrit le nombre de ligne
        User_Form_PGN.NLigne = N_Ligne_Recupere

        
        'On incrément le nombre de ligne
        N_Ligne_Recupere = N_Ligne_Recupere + 1
        
            
        'On incrémente le pointeur
        Pointeur = Pointeur + 1
Loop

End Sub 'On attend les prochains octets
'--------------------------------------------------------------
'                    Fin analyse
'--------------------------------------------------------------

'TEST ------------>>>
Sub Test()
Dim i As Integer

N_Ligne_Recupere = 0

        On Error Resume Next
        Open A_StrFichier For Append As #1
        On Error GoTo 0

Worksheets(FEUIL_EXEMPLE).Unprotect
        'initiale la valeur de Pi
        pi = Application.WorksheetFunction.pi() 'Valeur de pi
        
'Dure environ 35 secondes pour 9 trames complètes et quelques trames "n'importe quoi" envoyé 500 fois
For i = 1 To 500
DoEvents
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
