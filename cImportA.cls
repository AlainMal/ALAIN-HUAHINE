VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cImportA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========================== Import ========================
'
'Cette procédure  Txt(Chemin) réalise l'import du fichier TXT contenant les informations du bus CAN
'
'=======================================================


Const LIGNE_DEPART = 2  'Indice le numéro de la ligne dans lesquelles on écrit
'
'----------------------------------------------------------------------


'La Procédure Lit le fichier TXT correspondant au fichier importé
Public Sub Txt(chemin As String)

Dim i As Long, j As Long
Dim Texte As String
Dim resultat() As String
Dim ligne As String
Dim Tps, Tps2
Dim ID As String
Dim Longueur As String
Dim a1 As String, a2 As String, a3 As String, a4 As String, a5 As String, a6 As String, a7 As String, a8 As String
Dim l1 As Integer
Dim l2 As Integer

'DESACTIVE le bouton de lecture des txt
User_Form_PGN.CommandButton9.Enabled = False
User_Form_PGN.CommandButton5.Enabled = False

'Ouvre le fichier
On Error Resume Next        'Ne traite pas les erreurs
Open chemin For Input As #1
On Error GoTo 0             'Traite à nouveau les erreurs

  
j = LIGNE_DEPART                    'Ecrit la valeur de la deuxième ligne

User_Form_PGN.MAJ.ForeColor = 500   'Défini la couleur de l'affichage des ligne en cours

On Error GoTo Erreur        'S'il y a une erreur va vers Erreur:
Tps = Timer                 'Mesure le temps

'N'autorise pas l'actualisation automatique
Application.Calculation = xlCalculationManual

Do Until EOF(1)             'Temps que le fichier n'est pas terminé"

User_Form_PGN.MAJ = "MAJ " & j - 1 & " lignes"   'Ecrit le nombre de ligne en cours

If j Mod 200 Then DoEvents  'Donne la main au système pour qu'il affiche le nombre de ligne, peut être supprimé

    Line Input #1, ligne    'la phrase suivante
    Texte = ligne           'Duplique la ligne
    If Len(Texte) > 40 Or Len(Texte) < 10 Then
        MsgBox "ATTENTION: Ce fichier ne correpond pas à l'import", vbOKOnly + vbExclamation, "IMPORT TEXTE"
        GoTo FinFin
    End If
           
   ' resultat = Split(Texte, " ")    'Met la ligne dans un tableau avec le délimiteur "espaces"
 
    ID = Left(Texte, InStr(1, Texte, ";") - 1)
    ID = Right(ID, Len(ID) - 1)
    Longueur = "0" & Mid(Texte, InStr(1, Texte, ";") + 1, InStr(1, Texte, ":") - 1 - InStr(1, Texte, ";"))
    a1 = Mid(Texte, InStr(1, Texte, ":") + 1, InStr(1, Texte, ",") - 1 - InStr(1, Texte, ":"))
    If Len(a1) = 1 Then a1 = "0" & a1
    l1 = InStr(InStr(InStr(1, Texte, ","), Texte, ","), Texte, ",")
    l2 = InStr(l1 + 1, Texte, ",")
    If l2 = 0 Then
        a2 = Mid(Texte, l1 + 1, InStr(l1 + 1, Texte, "?") - l1 - 1)
        GoTo Debut
    Else
        a2 = Mid(Texte, l1 + 1, l2 - l1 - 1)
    End If
    If Len(a2) = 1 Then a2 = "0" & a2
    l1 = InStr(InStr(InStr(l1 + 1, Texte, ","), Texte, ","), Texte, ",")
    l2 = InStr(l1 + 1, Texte, ",")
    If l2 = 0 Then
        a3 = Mid(Texte, l1 + 1, InStr(l1 + 1, Texte, "?") - l1 - 1)
        GoTo Debut
    Else
        a3 = Mid(Texte, l1 + 1, l2 - l1 - 1)
    End If
    If Len(a3) = 1 Then a3 = "0" & a3
    l1 = InStr(InStr(InStr(l1 + 1, Texte, ","), Texte, ","), Texte, ",")
    l2 = InStr(l1 + 1, Texte, ",")
    If l2 = 0 Then
        a4 = Mid(Texte, l1 + 1, InStr(l1 + 1, Texte, "?") - l1 - 1)
        GoTo Debut
    Else
        a4 = Mid(Texte, l1 + 1, l2 - l1 - 1)
    End If
    If Len(a4) = 1 Then a4 = "0" & a4
    l1 = InStr(InStr(InStr(l1 + 1, Texte, ","), Texte, ","), Texte, ",")
    l2 = InStr(l1 + 1, Texte, ",")
    If l2 = 0 Then
        a5 = Mid(Texte, l1 + 1, InStr(l1 + 1, Texte, "?") - l1 - 1)
        GoTo Debut
    Else
        a5 = Mid(Texte, l1 + 1, l2 - l1 - 1)
    End If
    If Len(a5) = 1 Then a5 = "0" & a5
    l1 = InStr(InStr(InStr(l1 + 1, Texte, ","), Texte, ","), Texte, ",")
    l2 = InStr(l1 + 1, Texte, ",")
    If l2 = 0 Then
        a6 = Mid(Texte, l1 + 1, InStr(l1 + 1, Texte, "?") - l1 - 1)
        GoTo Debut
    Else
        a6 = Mid(Texte, l1 + 1, l2 - l1 - 1)
    End If
    If Len(a6) = 1 Then a6 = "0" & a6
    l1 = InStr(InStr(InStr(l1 + 1, Texte, ","), Texte, ","), Texte, ",")
    l2 = InStr(l1 + 1, Texte, ",")
    If l2 = 0 Then
        a7 = Mid(Texte, l1 + 1, InStr(l1 + 1, Texte, "?") - l1 - 1)
        GoTo Debut
    Else
        a7 = Mid(Texte, l1 + 1, l2 - l1 - 1)
    End If
    If Len(a7) = 1 Then a7 = "0" & a7
    l1 = InStr(InStr(InStr(l1 + 1, Texte, ","), Texte, ","), Texte, ",")
    l2 = InStr(l1 + 1, Texte, ",")
    If l2 = 0 Then
        a8 = Mid(Texte, l1 + 1, InStr(l1 + 1, Texte, "?") - l1 - 1)
    Else
        a8 = Mid(Texte, l1 + 1, l2 - l1 - 1)
    End If
    If Len(a8) = 1 Then a8 = "0" & a8
   
Debut:
    'Ecrit les 17 colonnes
    For i = CELL_DEBUT_IMPORT To CELL_FIN_IMPORT
        
        'On saute les lignes non ecritent
        If i = 2 Then i = 8
        
        'En fonction de la colonne
        Select Case i
            Case 1: Worksheets(FEUIL_NMEA).Cells(j, i) = j - 1 'Numéro de ligne
            'Défini l'ID
            Case 8: Worksheets(FEUIL_NMEA).Cells(j, i) = ID
            'Défini le nombre d'octet
            Case 9: Worksheets(FEUIL_NMEA).Cells(j, i) = Longueur
            'En fonction du nombre d'octet
            Case 10: If CInt(Longueur) > 0 Then Worksheets(FEUIL_NMEA).Cells(j, i) = a1 Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 11: If CInt(Longueur) > 1 Then Worksheets(FEUIL_NMEA).Cells(j, i) = a2 Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 12: If CInt(Longueur) > 2 Then Worksheets(FEUIL_NMEA).Cells(j, i) = a3 Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 13: If CInt(Longueur) > 3 Then Worksheets(FEUIL_NMEA).Cells(j, i) = a4 Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 14: If CInt(Longueur) > 5 Then Worksheets(FEUIL_NMEA).Cells(j, i) = a5 Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 15: If CInt(Longueur) > 4 Then Worksheets(FEUIL_NMEA).Cells(j, i) = a6 Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 16: If CInt(Longueur) > 6 Then Worksheets(FEUIL_NMEA).Cells(j, i) = a7 Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 17: If CInt(Longueur) > 7 Then Worksheets(FEUIL_NMEA).Cells(j, i) = a8 Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
        End Select
    Next i
        
     'Chronomètre
    Tps2 = Timer
    
    'Va à la ligne suivante
    j = j + 1

    
Loop

'Si 2ème erreur envois le 2ème message
On Error GoTo Erreur2

j = j - 1   'Enlève la ligne suplémentaire

Fin:



'Exit Sub '=================>>>


'Indique terminé si l'enssemble de fichier est lu
User_Form_PGN.MAJ.ForeColor = 32768   'Défini la couleur de l'affichage en vert
User_Form_PGN.MAJ = "Importés " & j - 1 & " lignes"
Beep   'Signale par un beep

'Ecrit les durées dans "Exemple"
Worksheets(FEUIL_TEMPS).Unprotect
Worksheets(FEUIL_TEMPS).Range("E11") = Format(Tps2 - Tps, "0.00") & " Secondes sur " & Format(j - 1, " #,##0 ") & " lignes"
Worksheets(FEUIL_TEMPS).Protect

User_Form_PGN.CommandButton9.Enabled = True     'Donne accèes au bouton RAZ

'Efface les lignes suivantes
Worksheets(FEUIL_NMEA).Unprotect
Worksheets(FEUIL_NMEA).Range(COLONNE_DEBUT_IMPORT & j & ":" & COLONNE_FIN_IMPORT & Worksheets(FEUIL_NMEA).Cells(Rows.Count, 1).End(xlUp).Row).ClearContents        'Les colonnes importées
Worksheets(FEUIL_NMEA).Range(COLONNE_DEBUT_RESULTAT & j & ":" & COLONNE_FIN_RESULTAT & Worksheets(FEUIL_NMEA).Cells(Rows.Count, 1).End(xlUp).Row + 100000).ClearContents  'Les colonnes d'affichage des résultats
Worksheets(FEUIL_NMEA).Protect

'Autorise l'actualisation d'excel
Application.Calculation = xlCalculationAutomatic

'Met à jour le choix pour la question "Rafaichir"
Choix = "2:" & j - 1

FinFin:
Close #1                                'Ferme le fichier
'SACTIVE le bouton de lecture des txt
User_Form_PGN.CommandButton9.Enabled = True
User_Form_PGN.CommandButton5.Enabled = True

Exit Sub

'S'il y a un erreur envoi le message
Erreur:
    MsgBox "Il y a une erreur dans le fichier à la ligne " & j, vbCritical + vbOKOnly, "Téléchargement"
GoTo Fin

Erreur2:
    MsgBox "Il y a une erreur dans la seconde partie" & vbCr & "mais l'import c'est passé sans encombre", vbCritical + vbOKOnly, "Téléchargement"
GoTo Fin

End Sub




