Attribute VB_Name = "Variables_Globales"
Option Explicit

'----------------------------------
'Définie les variables globales
'----------------------------------
Public pi
'Défini les valeurs publics
Public Derniereligne As Long                                                             'contenant la dernière ligne
Public Participants(NBR_PARTICIPANT, NBR_COLONNES_PARTICIPANTS) As String                   'Table contenant les participants sur le réseau
Public StopA As Boolean                                                                     'Demande d'arrêt
Public EnCours As Boolean                                                                   'Génération en cours
Public ValeurChoisie As String, ValeurChoisie2 As String, ValeurChoisie3 As String          'Valeur a afficher
Public ValeurChoisieTab As String                                                           'Valeur à afficher venant du tableau
Public FormatAffichage As String, FormatAffichage2 As String, FormatAffichage3 As String    'Format d'affichage
Public Echelle As Variant, Echelle2 As Variant, Echelle3 As Variant, EchellTab As Variant   'Echelle d'affichage
Public PGN_Tab As String, FormatAffichageTab As String                                      'Valeur venant des octet à mémorisés par PGN
Public Choix As String                                                                      'Valeur entrée pour le choix de la génération
Public ii As Long, iii As Long                                                              'Compte le nombre de lignes
Public Ma_position_latitude As Double                                                       'Récupère ma postion latitude
Public Ma_position_longitude As Double                                                      'Récupère ma position longitude

'Modules de Classes
Public cTable As New cTable         'Récuipère le contenue des tables dans les PGN
Public cID As New cID               'Gère le contenu des ID du bus CAN en mode normal
Public cExemple As New cExemple     'RAZ et Affiche le contenu de "Exemple" en mode normal
Public cExemplea As New cExemplea   'RAZ et Affiche le contenu de "Exemple" en temps réel
Public cImport As New cImport       'Importe les fichiers TXT de USB CAN venant du USB_CAN
Public cImportA As New cImportA     'Importe les fichiers TXT de USB CAN venant de l'enregistrement
Public cMMSIa As New cMMSIa         'Gère l'affiche des MMSI en temps réel
Public cMMSI As New cMMSI           'Gère l'affichage des MMSI en mode normal
Public cIDa As New cIDa             'Gère le contenu des ID du bus CAN en temps réel

'Non du fichier à écrire pour le bus CAN
Public NomFichier As String

'Mémoire du port COM
Public A_StrFichier As String
Public A_PORT_ID As Integer
Public A_COM As String
Public A_VITESSE As String
Public A_NOMBRE_CARACTERE As Long
Public A_Check_Exemple As Boolean
Public A_Check_MMSI As Boolean
Public A_CheckBox1 As Boolean

'Variable qui surveille le temps trop long
Public temps_trop_long As Long


