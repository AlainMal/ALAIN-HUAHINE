Attribute VB_Name = "Constantes_NMEA2000"
Option Explicit

'====== D�nommer les feuilles =========
'Noms des feuilles, peut �tre modifi�s
Public Const FEUIL_NMEA = "NMEA 2000"       'Feuille d'affichage des r�sultats du bus NMEA 2000
Public Const FEUIL_EXEMPLE = "Exemple"      'Feuille d'Exemple d'un exemple de donn�es NMEA 2000
Public Const FEUIL_TEMPS = "Temps"      'Feuille d'�x�cution des temps pass�s pour les diff�rentes actions
Public Const FEUIL_MMSI = "MMSI"      'Feuille des MMSIs d�tect�s

'======================================


'-------------------------- MEMOIRES -------------------------
'D�finition des PGN et des octets pour les stocker
'Utilis� pour la m�moire des octets d'un PGN
'Quand on change de trames, Ex.  La trame num�ro 0 on attend la trame suivante donc num�ro 1,
'il faut m�moriser pour pouvoir l'utiliser sur la trame suivante
'Le programme fonctionne jusqu'� pr�sent mais il peut �tre am�lior� (m�moris� le nuim�ro de la trame)
'
'Tableau des m�moires
Public Const NB_OCTETS = 8        'D�fini le nombre des octets
Public Const Nb_PGN = 255         'D�fini le nombre de PGN

'Les PGN utilis�s
Public Const PGN_129038 = 0
Public Const PGN_129039 = 1
Public Const PGN_129049 = 2
Public Const PGN_127506 = 3
Public Const PGN_126464 = 4
'A compl�ter....

'8 octets par trame
Public Const MEMOIRE_PGN_a1 = 0
Public Const MEMOIRE_PGN_a2 = 1
Public Const MEMOIRE_PGN_a3 = 2
Public Const MEMOIRE_PGN_a4 = 3
Public Const MEMOIRE_PGN_a5 = 4
Public Const MEMOIRE_PGN_a6 = 5
Public Const MEMOIRE_PGN_a7 = 6
Public Const MEMOIRE_PGN_a8 = 7
'-------------------- Fin des m�moires -----------------------


'----------------------------- Participants ---------------------------------------
'Liste des participants dans le tableau "Participants" � deux dimensions
'1�re dimension donne le num�ro du participant, la 2�me donne son Adresse  = 0 et son nom d'objet = 1
Public Const ADRESSE_PARTICIPANT = 0
Public Const NOM_PARTICIPANT = 1
Public Const NBR_PARTICIPANT = 240      'Nombres de participants
Public Const NBR_COLONNES_PARTICIPANTS = 2  'Nombres de colonnes (Adresse, Nom)
'-------------------------- Fin des participants --------------------------------


'========== LES COLONNES NMEA 2000 =====================
'      Liste des colonnes sur "NMEA 2000" Excel
'
'
'------------------------  Colonnes Bus CAN -----------------------------
'Ces 3  colonnes peuvent �tre modifi�es
'si vous avez un autre logiciel que USB CAN
'
'1 - Colonne de lD du bus CAN (Num�ro de  Colonnes 'H')
Public Const COLONNE_ID = 8 ' "H"
'
'2 - Premi�re Colonne des octets du bus CAN les autres sont � la suite (+1 +2 +3 etc.)
Public Const PREMIERE_COLONNE_OCTET = 10  '  "J"
'
'3 - Colonne du nombre d'octets re�us du bus CAN (0 � 7)
Public Const COLONNE_LONG = 8 ' "I"  n''est pas utilis�
'-------------------- Fin colonnes Bus CAN ------------------------------


'------------- Colonnes du NMEA 2000 -----------------------
'D�fini les colonnes de d�part et de fin des �critures NMEA 2000
Public Const COLONNE_DEBUT_RESULTAT = "AX"       'Colonne de d�part
Public Const COLONNE_FIN_RESULTAT = "BM"        'Colonne de fin

'D�fini la colonne du PGN issu de l'ID
Public Const COLONNE_PGN_ID = "AX"       'Valeur du PGN

'Colonnes des sources et destinations venant de l'ID du bus CAN
Public Const COLONNE_SOURCE = "AY"
Public Const COLONNE_DESTINATION = "AZ"


'Colonnes des affichages sources et destinations d�taill�es BL et BM
Public Const COLONNE_SOURCE_DETAIL = "BL"
Public Const COLONNE_DESTINATION_DETAIL = "BM"
'------------- Fin des colonnes NMEA 2000 -------------


'--------------------------------------- Colonnes r�sultats -----------------------------------------------------------
'Colonnes d 'affichage des r�sultats de la fonction PGN (D�finition, Valeur, �chelle, issues de la table)
Public Const COLONNE_DEF_1 = "BA"       'D�finition 1
Public Const COLONNE_VAL_1 = "BB"       'Valeur 1
Public Const COLONNE_ECH_1 = "BC"      'Echelle 1

Public Const COLONNE_DEF_2 = "BD"       'D�finition 2
Public Const COLONNE_VAL_2 = "BE"        'Valeur 2
Public Const COLONNE_ECH_2 = "BF"       'Echelle 2

Public Const COLONNE_DEF_3 = "BG"       'D�finition 3
Public Const COLONNE_VAL_3 = "BH"       'Valeur 3
Public Const COLONNE_ECH_3 = "BI"       'Echelle 3

Public Const COLONNE_TAB = "BJ"         'Donn�es Issu des tables ou 4�me difinition
'-------------------------------------- Fin des colonnes r�sultats ---------------------------------------------------


'----- Colonnes des participants �critent pas l'utilisateur ----------
Public Const COLONNE_ADRESSE = 30    ' "AD"
Public Const COLONNE_NOM = 31      ' "AE"


'----- Ces colonnes sont utils�es pour l'import -----
Public Const COLONNE_DEBUT_IMPORT = "A"
Public Const COLONNE_FIN_IMPORT = "Q"
Public Const CELL_DEBUT_IMPORT = 1 ' "A"
Public Const CELL_FIN_IMPORT = 17  ' "Q"
'


'------------- Colonnes pour les MMSI ------------------
Public Const PREMIERE_COLONNE_MMSI = "A"
Public Const COLONNE_NUM_MMSI = "A"
Public Const COLONNE_HEURE_MMSI = "B"
Public Const COLONNE_LATITUDE_MMSI = "C"
Public Const COLONNE_LONGITUDE_MMSI = "D"
Public Const COLONNE_NOM_MMSI = "E"
Public Const COLONNE_COG_MMSI = "F"
Public Const COLONNE_SOG_MMSI = "G"
Public Const COLONNE_DISTANCE_MMSI = "H"
Public Const COLONNE_CLASSE_MMSI = "I"
'
'=============== Fin des colonnes =======================

'Si pas d'info � �crire dans les variables issu du bus CAN
Public Const PAS_INFO = "z"
Public Const NumFic = 1
'



'Ex�cutez cette proc�dure pour faire la mise � jour des nom des Feuilles
Public Sub Renommer()
    'Renomme les feuilles
    Sheets(1).Name = FEUIL_NMEA
    Sheets(3).Name = FEUIL_EXEMPLE
    Sheets(4).Name = FEUIL_TEMPS
    Sheets(2).Name = FEUIL_MMSI
End Sub



