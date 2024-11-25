VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========= Fonctions de r�cup�ration des infos dans les tables =========
'Ex. la fonction Temperature(Num�ro inscrit dans les octets) retourne le texte correspondant
'On peut en rajouter, seules les tableaux utiles sont d�finis
'=======================================================================

' R�cup�re les info dans la table Temp�rature
' Num -> Le num�ro inscit dans la table
Public Function Temperature(Num As String) As String

Select Case Num
    Case "0"
        Temperature = "Temperature Mer"
    Case "1"
        Temperature = "Temperature Exterieur"
    Case "2"
        Temperature = "Temperature Interieur"
    Case "3"
        Temperature = "Temperature Machine"
    Case "4"
        Temperature = "Temperature Cabine"
    Case "5"
        Temperature = "Temperature Vivier"
    Case "6"
        Temperature = "Temperature puits d�app�t"
    Case "7"
        Temperature = "Temperature r�frig�ration"
    Case "8"
        Temperature = "Temperature chauffage"
    Case "9"
        Temperature = "Temperature point de ros�e"
'    Case "10"
'        Source_Temperature = "Temperature puits d�app�t"
'    Case "11"
'        Source_Temperature = "Temperature puits d�app�t"
'    Case "12"
'        Source_Temperature = "Temperature puits d�app�t"
'    Case "13"
'        Source_Temperature = "Temperature puits d�app�t"
    Case "14"
        Temperature = "Temperature gaz d��chappement"
    Case "15"
        Temperature = "Temperature joint d�arbre"
    Case Else
        Temperature = "Inconnu"
       
End Select

End Function

' R�cup�re les info dans la table Speed
Public Function Speed(Num As String) As String

Select Case Num
    Case "0"
        Speed = "Roue � aubes"
    Case "1"
        Speed = "Tube de Pitot"
    Case "2"
        Speed = "Doppler"
    Case "3"
        Speed = "Ultrasons"
    Case "4"
        Speed = "�lectromagn�tique"
    Case Else
        Speed = "Inconnu"
End Select

End Function

' R�cup�re les info dans la table Vent
Public Function Wind(Num As String) As String

Select Case Num
    Case "0"
        Wind = "Vent Ground"
    Case "1"
        Wind = "Vent Magn�tique"
    Case "2"
        Wind = "Vent Apparent"
    Case "3"
        Wind = "Vent Vrai Fond"
    Case "4"
        Wind = "Vent Vrai Mer"
    Case Else
        Wind = "Inconnu"
End Select

End Function

' R�cup�re les info dans la table Batterie
Public Function Batterie(Num As String) As String

Select Case Num
    Case "0"
        Batterie = "Batterie"
    Case "1"
        Batterie = "Alternateur"
    Case "2"
        Batterie = "Photovolta�que"
    Case "3"
        Batterie = "A�rog�n�rateur"
    Case Else
        Batterie = "Inconnu"
End Select

End Function

' R�cup�re les info dans la table R�servoirs
Public Function Tank(Num As String) As String

Select Case Num
    Case "0"
        Tank = "Niveau Combustible"
    Case "1"
        Tank = "Niveau Eau"
    Case "2"
        Tank = "Noveau Eaux grises"
    Case "3"
        Tank = "Niveau Eaux Courante"
    Case "4"
        Tank = "Niveau Huile"
    Case "5"
        Tank = "Niveau Eaux noires"
    Case Else
        Tank = "Inconnu"
End Select

End Function

' R�cup�re les info dans la table Pression
Public Function PRESSION(Num As String) As String

Select Case Num
    Case "0"
        PRESSION = "Pression Atmosph�rique"
    Case "1"
        PRESSION = "Pression Eau"
    Case "2"
        PRESSION = "Pression Vapeur"
    Case "3"
        PRESSION = "Pression Air comprim�"
    Case "4"
        PRESSION = "Pression Hydraulique"
    Case "5"
        PRESSION = "Pression Filtre"
    Case "6"
        PRESSION = "Pression Altim�treR�glage"
    Case "7"
        PRESSION = "Pression Huile"
    Case "8"
        PRESSION = "Pression Combustible"
End Select

End Function

'R�cup�re les d�signations des adresses d�finis par l'utilisateur dans les colonnes AD et AE
Public Function Adresses(Num As String) As String

Dim i As Integer

If Num >= 240 Then
    Adresses = "TOUS"
    Exit Function
End If
'Les adresses sont d�fini dans le tableau "Participants"
For i = 0 To 32
    If Num = Participants(i, ADRESSE_PARTICIPANT) Then
        Adresses = Participants(i, NOM_PARTICIPANT)
        Exit For
   End If
Next i

End Function
