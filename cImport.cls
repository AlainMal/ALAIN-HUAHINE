VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========================== Import ========================
'
'Cette proc�dure  Txt(Chemin) r�alise l'import du fichier TXT contenant les informations du bus CAN
'
'Doit �tre modifi� sur vous avez d'autre sources de donn�es que USB
'ou le num�ro de la ligne (j) et num�ro de colonne (i),
'La longueur de la premi�re ligne (LONG_TRAME) dans les constantes
'
'=======================================================


Const LONG_TRAME = 103  'Indique la longueur d'une trame lue
Const LIGNE_DEPART = 2  'Indice le num�ro de la ligne dans lesquelles on �crit

'La Proc�dure Lit le fichier TXT correspondant au USB CAN
Public Sub Txt(chemin As String)

Dim i As Integer, j As Integer
Dim Texte As String
Dim resultat() As String
Dim Lg As Integer
Dim ligne As String
Dim Tps, Tps2

'DESACTIVE le bouton de lecture des txt
User_Form_PGN.CommandButton5.Enabled = False
User_Form_PGN.CommandButton9.Enabled = False

'Ouvre lr fichier
On Error Resume Next        'Ne traite pas les erreurs
Open chemin For Input As #1

'lit la premiere ligne du fichier
Line Input #1, ligne
On Error GoTo 0             'Traite � nouveau les erreurs

If ligne = "" Then GoTo Fin
  
'Si ne n'est pas le bon fichier envoi le message, si la 1�re ligne n'a pas la bonne longueur
If Len(ligne) <> LONG_TRAME Then
        MsgBox "ATTENTION: Ce fichier ne correpond pas � l'USB CAN", vbOKOnly + vbExclamation, "IMPORT TEXTE"
    GoTo Fin
End If

j = LIGNE_DEPART                    'Ecrit la valeur de la deuxi�me ligne

User_Form_PGN.MAJ.ForeColor = 500   'D�fini la couleur de l'affichage des ligne en cours

On Error GoTo Erreur        'S'il y a une erreur va vers Erreur:
Tps = Timer                 'Mesure le temps

'N'autorise pas l'actualisation automatique
Application.Calculation = xlCalculationManual

Do Until EOF(1)             'Temps que le fichier n'est pas termin�"

User_Form_PGN.MAJ = "MAJ " & j - 1 & " lignes"   'Ecrit le nombre de ligne en cours

If j Mod 200 Then DoEvents  'Donne la main au syst�me pour qu'il affiche le nombre de ligne, peut �tre supprim�

    Line Input #1, ligne    'la phrase suivante
    Texte = ligne           'Duplique la ligne
    
    'Supprime les espaces en trop dans la ligne, � controler sur de nouveaux fichiers
    For i = 0 To 5
        Texte = Replace(Texte, "  ", " ")
    Next i
       
    resultat = Split(Texte, " ")    'Met la ligne dans un tableau avec le d�limiteur "espaces"
 
    'Ecrit les 17 colonnes
    For i = CELL_DEBUT_IMPORT To CELL_FIN_IMPORT
        'En fonction de la colonne
        Select Case i
            Case 1: Worksheets(FEUIL_NMEA).Cells(j, i) = resultat(i - 1)
            Case 8: Worksheets(FEUIL_NMEA).Cells(j, i) = resultat(i - 1)
            'D�fini le nombre d'octet
            Case 9: Worksheets(FEUIL_NMEA).Cells(j, i) = resultat(i - 1): Lg = resultat(i - 1)
            'En fonction du nombre d'octet
            Case 10: If Lg > 0 Then Worksheets(FEUIL_NMEA).Cells(j, i) = resultat(i - 1) Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 11: If Lg > 1 Then Worksheets(FEUIL_NMEA).Cells(j, i) = resultat(i - 1) Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 12: If Lg > 2 Then Worksheets(FEUIL_NMEA).Cells(j, i) = resultat(i - 1) Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 13: If Lg > 3 Then Worksheets(FEUIL_NMEA).Cells(j, i) = resultat(i - 1) Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 14: If Lg > 5 Then Worksheets(FEUIL_NMEA).Cells(j, i) = resultat(i - 1) Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 15: If Lg > 4 Then Worksheets(FEUIL_NMEA).Cells(j, i) = resultat(i - 1) Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 16: If Lg > 6 Then Worksheets(FEUIL_NMEA).Cells(j, i) = resultat(i - 1) Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
            Case 17: If Lg > 7 Then Worksheets(FEUIL_NMEA).Cells(j, i) = resultat(i - 1) Else Worksheets(FEUIL_NMEA).Cells(j, i) = ""
        End Select
    Next i
        
     'Chronom�tre
    Tps2 = Timer

    'Si le nombre de ligne est trop long (car le num�ero est coll� � la valeur suivante)
   If j > 10000 Then
        User_Form_PGN.MAJ = "Import�s " & Format(j - 1, " #,##0 ") & " lignes"
        MsgBox "Vous n'avez droit qu'� 10 000 lignes maximum" & vbCr & "Vous avez import� les " & Format(j - 1, " #,##0 ") & " lignes", vbInformation + vbOKOnly
        GoTo Fin
    End If
    
    'Va � la ligne suivante
    j = j + 1

    
Loop

'Si 2�me erreur envois le 2�me message
On Error GoTo Erreur2

j = j - 1   'Enl�ve la ligne supl�mentaire

Fin:

Close #1                                'Ferme le fichier

'Indique termin� si l'enssemble de fichier est lu
User_Form_PGN.MAJ.ForeColor = 32768   'D�fini la couleur de l'affichage en vert
User_Form_PGN.MAJ = "Import�s " & j - 1 & " lignes"
Beep   'Signale par un beep

'Ecrit les dur�es dans "Exemple"
Worksheets(FEUIL_TEMPS).Unprotect
Worksheets(FEUIL_TEMPS).Range("E11") = Format(Tps2 - Tps, "0.00") & " Secondes sur " & Format(j - 1, " #,##0 ") & " lignes"
Worksheets(FEUIL_TEMPS).Protect

User_Form_PGN.CommandButton5.Enabled = True     'Donne acc�es au bouton RAZ
User_Form_PGN.CommandButton9.Enabled = True     'Donne acc�es au bouton RAZ

'Efface les lignes suivantes
Worksheets(FEUIL_NMEA).Unprotect
Worksheets(FEUIL_NMEA).Range(COLONNE_DEBUT_IMPORT & j & ":" & COLONNE_FIN_IMPORT & Worksheets(FEUIL_NMEA).Cells(Rows.Count, 1).End(xlUp).Row).ClearContents        'Les colonnes import�es
Worksheets(FEUIL_NMEA).Range(COLONNE_DEBUT_RESULTAT & j & ":" & COLONNE_FIN_RESULTAT & Worksheets(FEUIL_NMEA).Cells(Rows.Count, 1).End(xlUp).Row + 100000).ClearContents  'Les colonnes d'affichage des r�sultats
Worksheets(FEUIL_NMEA).Protect

'Autorise l'actualisation d'excel
Application.Calculation = xlCalculationAutomatic

'Met � jour le choix pour la question "Rafaichir"
Choix = "2:" & j - 1

Exit Sub

'S'il y a un erreur envoi le message
Erreur:
    MsgBox "Il y a une erreur dans le fichier � la ligne " & j, vbCritical + vbOKOnly, "T�l�chargement"
GoTo Fin

Erreur2:
    MsgBox "Il y a une erreur dans la seconde partie" & vbCr & "mais l'import c'est pass� sans encombre", vbCritical + vbOKOnly, "T�l�chargement"
GoTo Fin

End Sub


'Module de test
Sub titi()
cImport.Txt (UseFileDialogOpen)

End Sub



