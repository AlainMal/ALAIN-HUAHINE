Attribute VB_Name = "Constantes_NMEA2000"
Option Explicit

'====== Dénommer les feuilles =========
'Noms des feuilles, peut être modifiés
Public Const FEUIL_NMEA = "NMEA 2000"       'Feuille d'affichage des résultats du bus NMEA 2000
Public Const FEUIL_EXEMPLE = "Exemple"      'Feuille d'Exemple d'un exemple de données NMEA 2000
Public Const FEUIL_TEMPS = "Temps"      'Feuille d'éxécution des temps passés pour les différentes actions
Public Const FEUIL_MMSI = "MMSI"      'Feuille des MMSIs détectés

'======================================


'-------------------------- MEMOIRES -------------------------
'Définition des PGN et des octets pour les stocker
'Utilisé pour la mémoire des octets d'un PGN
'Quand on change de trames, Ex.  La trame numéro 0 on attend la trame suivante donc numéro 1,
'il faut mémoriser pour pouvoir l'utiliser sur la trame suivante
'Le programme fonctionne jusqu'à présent mais il peut être amélioré (mémorisé le nuiméro de la trame)
'
'Tableau des mémoires
Public Const NB_OCTETS = 8        'Défini le nombre des octets
Public Const Nb_PGN = 255         'Défini le nombre de PGN

'Les PGN utilisés
Public Const PGN_129038 = 0
Public Const PGN_129039 = 1
Public Const PGN_129049 = 2
Public Const PGN_127506 = 3
Public Const PGN_126464 = 4
'A compléter....

'8 octets par trame
Public Const MEMOIRE_PGN_a1 = 0
Public Const MEMOIRE_PGN_a2 = 1
Public Const MEMOIRE_PGN_a3 = 2
Public Const MEMOIRE_PGN_a4 = 3
Public Const MEMOIRE_PGN_a5 = 4
Public Const MEMOIRE_PGN_a6 = 5
Public Const MEMOIRE_PGN_a7 = 6
Public Const MEMOIRE_PGN_a8 = 7
'-------------------- Fin des mémoires -----------------------


'----------------------------- Participants ---------------------------------------
'Liste des participants dans le tableau "Participants" à deux dimensions
'1ère dimension donne le numéro du participant, la 2ème donne son Adresse  = 0 et son nom d'objet = 1
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
'Ces 3  colonnes peuvent être modifiées
'si vous avez un autre logiciel que USB CAN
'
'1 - Colonne de lD du bus CAN (Numéro de  Colonnes 'H')
Public Const COLONNE_ID = 8 ' "H"
'
'2 - Première Colonne des octets du bus CAN les autres sont à la suite (+1 +2 +3 etc.)
Public Const PREMIERE_COLONNE_OCTET = 10  '  "J"
'
'3 - Colonne du nombre d'octets reçus du bus CAN (0 à 7)
Public Const COLONNE_LONG = 8 ' "I"  n''est pas utilisé
'-------------------- Fin colonnes Bus CAN ------------------------------


'------------- Colonnes du NMEA 2000 -----------------------
'Défini les colonnes de départ et de fin des écritures NMEA 2000
Public Const COLONNE_DEBUT_RESULTAT = "AX"       'Colonne de départ
Public Const COLONNE_FIN_RESULTAT = "BM"        'Colonne de fin

'Défini la colonne du PGN issu de l'ID
Public Const COLONNE_PGN_ID = "AX"       'Valeur du PGN

'Colonnes des sources et destinations venant de l'ID du bus CAN
Public Const COLONNE_SOURCE = "AY"
Public Const COLONNE_DESTINATION = "AZ"


'Colonnes des affichages sources et destinations détaillées BL et BM
Public Const COLONNE_SOURCE_DETAIL = "BL"
Public Const COLONNE_DESTINATION_DETAIL = "BM"
'------------- Fin des colonnes NMEA 2000 -------------


'--------------------------------------- Colonnes résultats -----------------------------------------------------------
'Colonnes d 'affichage des résultats de la fonction PGN (Définition, Valeur, échelle, issues de la table)
Public Const COLONNE_DEF_1 = "BA"       'Définition 1
Public Const COLONNE_VAL_1 = "BB"       'Valeur 1
Public Const COLONNE_ECH_1 = "BC"      'Echelle 1

Public Const COLONNE_DEF_2 = "BD"       'Définition 2
Public Const COLONNE_VAL_2 = "BE"        'Valeur 2
Public Const COLONNE_ECH_2 = "BF"       'Echelle 2

Public Const COLONNE_DEF_3 = "BG"       'Définition 3
Public Const COLONNE_VAL_3 = "BH"       'Valeur 3
Public Const COLONNE_ECH_3 = "BI"       'Echelle 3

Public Const COLONNE_TAB = "BJ"         'Données Issu des tables ou 4ème difinition
'-------------------------------------- Fin des colonnes résultats ---------------------------------------------------


'----- Colonnes des participants écritent pas l'utilisateur ----------
Public Const COLONNE_ADRESSE = 30    ' "AD"
Public Const COLONNE_NOM = 31      ' "AE"


'----- Ces colonnes sont utilsées pour l'import -----
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

'Si pas d'info à écrire dans les variables issu du bus CAN
Public Const PAS_INFO = "z"
Public Const NumFic = 1
'



'Exécutez cette procédure pour faire la mise à jour des nom des Feuilles
Public Sub Renommer()
    'Renomme les feuilles
    Sheets(1).Name = FEUIL_NMEA
    Sheets(3).Name = FEUIL_EXEMPLE
    Sheets(4).Name = FEUIL_TEMPS
    Sheets(2).Name = FEUIL_MMSI
End Sub



