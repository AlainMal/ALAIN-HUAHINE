VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} User_Form_PGN 
   Caption         =   "G�n�rez les trames NMEA 2000"
   ClientHeight    =   1944
   ClientLeft      =   14712
   ClientTop       =   3456
   ClientWidth     =   4392
   OleObjectBlob   =   "User_Form_PGN.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "User_Form_PGN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NomFichier As String



Private Sub BufferA_Change()
Dim NomFichier As String
   
    If BufferA = True Then
        User_Form_PGN.CommandButton1.Enabled = False
        User_Form_PGN.CommandButton4.Enabled = False
        CommandButton8.Enabled = False
        CommandButton1.Enabled = False
        CommandButton4.Enabled = False
        CommandButton5.Enabled = False
        CommandButton8.Enabled = False
        If EnCours Then
            MsgBox "le programme de rafraichement est en cours" & vbCr & "Veuillez attendre ou cliquer sur le bouton ""Arr�ter""", vbInformation, "TEMPS REEL"
            BufferA = False
            Exit Sub
        End If
        'initiale la valeur de Pi
        pi = Application.WorksheetFunction.pi() 'Valeur de pi

        'Initialise les diff�entes valeurs mesur�es
        ValeurChoisie = PAS_INFO
        ValeurChoisie2 = PAS_INFO
        ValeurChoisie3 = PAS_INFO
        ValeurChoisieTab = PAS_INFO
    
        Call Temps_Reel  'Appele la fonction temps r�el
    Else
        User_Form_PGN.CommandButton1.Enabled = True
        User_Form_PGN.CommandButton4.Enabled = True
        CommandButton8.Enabled = True
        CommandButton1.Enabled = True
        CommandButton4.Enabled = True
        CommandButton5.Enabled = True
        CommandButton8.Enabled = True
        BufferA.Enabled = False
    End If


End Sub

Private Sub CommandButton1_Click()
    
    PGN_Decode                                 'Appelle la fonction de d�codage du PGN

End Sub

Private Sub CommandButton3_Click()

    If EnCours Then StopA = True Else MsgBox "Il n'y a pas de g�n�ration en cours", vbInformation, "Arr�ter"

End Sub

Private Sub CommandButton4_Click()
  
If ActiveSheet.Name = FEUIL_NMEA Then
    If MsgBox("Voulez vous effacer les valeurs import�es du bus CAN ?" & vbCr & vbCr & "ATTENTION: Si vous effacer les valeurs, vous ne pourrez plus faire un ""Rafraichir"", vous devriez importer les nouvelles valeurs", vbYesNo + vbQuestion, "RAZ du " & FEUIL_NMEA) = vbYes Then
        Worksheets(FEUIL_NMEA).Unprotect
'        Worksheets(FEUIL_NMEA).Range("AX:BM10000").ClearContents
        Worksheets(FEUIL_NMEA).Range("A2:Q10000").ClearContents
        Worksheets(FEUIL_NMEA).Protect
    End If
End If
  
If ActiveSheet.Name = FEUIL_EXEMPLE Then
    If MsgBox("Voulez vous effacer les valeurs de l'" & FEUIL_EXEMPLE & " ?" & vbCr & vbCr & "Faites un rafraichir pour r�actualiser" & vbCr & "Ou branchez le c�ble sur le bus CAN et lancez la communication", vbYesNo + vbQuestion, "RAZ de l'" & FEUIL_EXEMPLE) = vbYes Then
        cExemple.RAZ
    End If
End If
If ActiveSheet.Name = FEUIL_MMSI Then
    If MsgBox("Voulez vous effacer les valeurs du " & FEUIL_MMSI & " ?" & vbCr & vbCr & "Faites un rafraichir pour r�actualiser" & vbCr & "Ou branchez le c�ble sur le bus CAN et lancez la communication", vbYesNo + vbQuestion, "RAZ du " & FEUIL_MMSI) = vbYes Then
        Worksheets(FEUIL_MMSI).Range("A2:K100").ClearContents
    End If
End If
If Not ActiveSheet.Name = FEUIL_EXEMPLE And Not ActiveSheet.Name = FEUIL_MMSI And Not ActiveSheet.Name = FEUIL_NMEA Then
    MsgBox "Il n'y a rien � effac� sur cette feuille", vbInformation, "RAZ"
End If
End Sub
'Bouton importer les donn�es du bus USB_CAN
Private Sub CommandButton5_Click()
    
    If MsgBox("Voulez vous importer un fichier ""txt"" du bus USB_CAN ?", vbYesNo + vbQuestion, "IMPORT SUR LA FEUILLE " & FEUIL_NMEA & ", vient du programme USB-CAN") = vbYes Then
       On Error Resume Next
       cImport.Txt (UseFileDialogOpen())
       
    End If
    
End Sub

Private Sub CommandButton6_Click()
If Aide.Visible Then
    Aide.Hide
Else
    Aide.Show
End If
End Sub

Private Sub CommandButton7_Click()
If Not (MMSI.Visible) Then
    cMMSI.Tri
    MMSI.Show
Else
    MMSI.Hide
End If
End Sub
'Montre l'�cran de param�tres
Private Sub CommandButton8_Click()
        User_Form_COM.Show

    User_Form_COM.StrFichier.Value = A_StrFichier
    User_Form_COM.PORT_ID.Value = A_PORT_ID
    User_Form_COM.COM.Value = A_COM
    User_Form_COM.VITESSE.Value = A_VITESSE
    User_Form_COM.NOMBRE_CARACTERE.Value = A_NOMBRE_CARACTERE
    User_Form_COM.Check_Exemple.Value = A_Check_Exemple
    User_Form_COM.Check_MMSI.Value = A_Check_MMSI
    User_Form_COM.CheckBox1.Value = A_CheckBox1
    
End Sub

Private Sub CommandButton9_Click()
    If MsgBox("Voulez vous importer un fichier ""txt"" r�alis� depuis cette application ?", vbYesNo + vbQuestion, "IMPORT SUR LA FEUILLE " & FEUIL_NMEA & ", r�alis� par cette application") = vbYes Then
       On Error Resume Next
       cImportA.Txt (UseFileDialogOpen())
       
    End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        If MsgBox("Voulez vous fermer " & ThisWorkbook.Name & " sous Excel ?", vbYesNo + vbQuestion, FEUIL_NMEA) = vbYes Then
            Cancel = True
            'Call CommandButton3_Click
            Application.Quit
            Exit Sub
        Else
            Cancel = True
        End If
End Sub
