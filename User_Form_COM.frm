VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} User_Form_COM 
   Caption         =   "Paramètres de communications série"
   ClientHeight    =   3456
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "User_Form_COM.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "User_Form_COM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Box d'enregistrement du fichier
Private Sub CheckBox1_AfterUpdate()
    If CheckBox1 = True Then
        If StrFichier = "" Then
            MsgBox "Choisissez le fichier avant de cliquer", vbOKOnly + vbInformation, "ENREGISTREMENT DE LA COMMUNICATION"
            CheckBox1 = False
        End If
    End If

End Sub

'Bouton Annuler
Private Sub CommandButton10_Click()

            'on conserve les valeaur initiales
            StrFichier.Value = A_StrFichier
            PORT_ID.Value = A_PORT_ID
            COM.Value = A_COM
            VITESSE.Value = A_VITESSE
            NOMBRE_CARACTERE.Value = A_NOMBRE_CARACTERE
            Check_Exemple.Value = A_Check_Exemple
            Check_MMSI.Value = A_Check_MMSI
            CheckBox1.Value = A_CheckBox1

User_Form_COM.Hide
End Sub

'Bouton OK
Private Sub CommandButton11_Click()
If PORT_ID.Value = "" Or COM.Value = "" Or VITESSE.Value = "" Or NOMBRE_CARACTERE.Value = "" Then
    MsgBox "Il manque certains paramètres, veuillez compléter", vbOKOnly + vbInformation, "ENREGISTRER LA COMMUNICATION"
    Exit Sub
End If

            'On range les valeur dana la feuille "Memoires"
            Sheets("Memoires").Range("A1") = StrFichier.Value
            Sheets("Memoires").Range("A2") = PORT_ID.Value
            Sheets("Memoires").Range("A3") = COM.Value
            Sheets("Memoires").Range("A4") = VITESSE.Value
            Sheets("Memoires").Range("A5") = NOMBRE_CARACTERE.Value
            Sheets("Memoires").Range("A6") = Check_Exemple.Value
            Sheets("Memoires").Range("A7") = Check_MMSI.Value
            Sheets("Memoires").Range("A8") = CheckBox1.Value
            
            'On conserve les valeurrs pour le prochain
            A_StrFichier = StrFichier.Value
            A_PORT_ID = PORT_ID.Value
            A_COM = COM.Value
            A_VITESSE = VITESSE.Value
            A_NOMBRE_CARACTERE = NOMBRE_CARACTERE.Value
            A_Check_Exemple = Check_Exemple.Value
            A_Check_MMSI = Check_MMSI.Value
            A_CheckBox1 = CheckBox1.Value

User_Form_COM.Hide
End Sub

'Bouton Choisir un fichier
Private Sub CommandButton8_Click()
Dim NomFichier_lu
    
    NomFichier_lu = CreerFichierTexte
    If Not (NomFichier_lu = "Faux") Then
        NomFichier = NomFichier_lu
        CheckBox1 = True
    End If
    
    StrFichier.Value = NomFichier
    
End Sub
'

'Bonton Lancer
Private Sub CommandButton9_Click()

If PORT_ID.Value = "" Or COM.Value = "" Or VITESSE.Value = "" Or NOMBRE_CARACTERE.Value = "" Then
    MsgBox "Il manque certains paramètres, veuillez compléter", vbOKOnly + vbInformation, "LANCER LA COMMUNICATION"
    Exit Sub
End If


    User_Form_PGN.BufferA = False
    User_Form_COM.Hide
    
    If CheckBox1 = True Then
        If StrFichier <> "" Then
            NomFichier = StrFichier
            If MsgBox("Vous allez enregistrer les trames reçus dans le fichier: " & vbCr & NomFichier & vbCr & vbCr & "Etes vous prêt à lancer la communication ?" & vbCr & vbCr & "Dévalider la case à cocher ""Rafraichir en temps réel"" pour terminer", vbOKCancel + vbQuestion, "ENREGISTREMENT DU BUS CAN") = vbOK Then
                On Error Resume Next
                Open StrFichier For Append As #1
                On Error GoTo 0
                User_Form_PGN.BufferA.Enabled = True
                User_Form_PGN.BufferA = True
            Else
                User_Form_COM.Show
                User_Form_PGN.BufferA.Enabled = False
            End If
        End If
    Else
            If MsgBox("Vous allez rafraichir les écrans cochés" & vbCr & vbCr & "Etes vous prêt à lancer la communication ?" & vbCr & vbCr & "Dévalider la case à cocher ""Rafraichir en temps réel"" pour arrêter", vbOKCancel + vbQuestion, "ENREGISTREMENT DU BUS CAN") = vbOK Then
                User_Form_PGN.BufferA.Enabled = True
                User_Form_PGN.BufferA = True
            Else
                User_Form_COM.Show
                User_Form_PGN.BufferA.Enabled = False
            End If
    End If
            
            'Sauvegarder les paramètres
            A_PORT_ID = PORT_ID.Value
            A_COM = COM.Value
            A_VITESSE = VITESSE.Value
            A_NOMBRE_CARACTERE = NOMBRE_CARACTERE.Value
            A_Check_MMSI = Check_MMSI.Value
            A_Check_Exemple = Check_Exemple.Value
            A_StrFichier = StrFichier.Value
            
            'Sauvegarde en mémoire
            Sheets("Memoires").Range("A2") = PORT_ID.Value
            Sheets("Memoires").Range("A3") = COM.Value
            Sheets("Memoires").Range("A4") = VITESSE.Value
            Sheets("Memoires").Range("A5") = NOMBRE_CARACTERE.Value
            Sheets("Memoires").Range("A6") = Check_Exemple.Value
            Sheets("Memoires").Range("A7") = Check_MMSI.Value
            Sheets("Memoires").Range("A8") = CheckBox1.Value
            Sheets("Memoires").Range("A1") = StrFichier.Value

    
    
End Sub

'Boite d'affichage du fichier
Private Sub StrFichier_AfterUpdate()
    If StrFichier = "" Then
        CheckBox1 = False
    End If
End Sub

Private Sub UserForm_Initialize()

            'Récupère les anciennes valeurs
            StrFichier.Value = A_StrFichier
            PORT_ID.Value = A_PORT_ID
            COM.Value = A_COM
            VITESSE.Value = A_VITESSE
            NOMBRE_CARACTERE.Value = A_NOMBRE_CARACTERE
            Check_Exemple.Value = A_Check_Exemple
            Check_MMSI.Value = A_Check_MMSI
            CheckBox1.Value = A_CheckBox1


    ' Ajouter des éléments un par un à la ComboBox
    COM.AddItem "COM1"
    COM.AddItem "COM2"
    COM.AddItem "COM3"
    COM.AddItem "COM4"
    COM.AddItem "COM5"
    COM.AddItem "COM6"
    COM.AddItem "COM7"
    
    VITESSE.AddItem "4800"
    VITESSE.AddItem "9600"
    VITESSE.AddItem "115200"
    VITESSE.AddItem "500000"
    VITESSE.AddItem "1000000"
    VITESSE.AddItem "2000000"
    
    NomFichier = ""
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Cancel = True

Call CommandButton10_Click
End Sub

