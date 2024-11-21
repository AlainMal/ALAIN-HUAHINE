VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================ FONCTION PGN ==================
'
'La Fonction PGN(Numéro de la ligne) retourne le PGN de l'ID du bus CAN
'La fonction SourceID(Numéro de la ligne) retourne l'adresse source
'La fonction DestinationID(Numéro de la ligne) retourne l'adresse de destination
'
'===============================================

Public Function Pgn(ligne As Long) As String

Dim ID As String   'ID du bus CAN

'Récupère l'ID sur la colonne 8 (H)
ID = IIf(Worksheets(FEUIL_NMEA).Cells(ligne, COLONNE_ID) = "", "00", Worksheets(FEUIL_NMEA).Cells(ligne, COLONNE_ID))

'Fait le masque pour récuperer le PGN
Pgn = CDbl("&H" & ID) And CDbl("&H0FF0000")

If WorksheetFunction.Bitrshift(Arg1:=Pgn, Arg2:=16) < 240 Then   'Si le PS < 240
    Pgn = CDbl("&H" & ID) And CDbl("&H1FF0000")
    Pgn = WorksheetFunction.Bitrshift(Arg1:=Pgn, Arg2:=8)
Else
    Pgn = CDbl("&H" & ID) And CDbl("&H1FFFF00")
    Pgn = WorksheetFunction.Bitrshift(Arg1:=Pgn, Arg2:=8)
End If

End Function

' Fonction de Récupèration de l'adresse source dans l'ID
Public Function SourceID(ligne As Long) As String
Dim ID As String  'ID du bus CAN

'Récupère l'ID
ID = Worksheets(FEUIL_NMEA).Cells(ligne, COLONNE_ID)

'Fait le masque pour récuperer l'adresse source
SourceID = CDbl("&H" & ID) And CDbl("&H00000FF")

End Function

' Fonction de  Récupèration de l'adresse destination dans l'ID
Public Function DestinationID(ligne As Long) As String
Dim ID As String  'ID du bus CAN

'Récupère l'ID dans sa colonne
ID = Worksheets(FEUIL_NMEA).Cells(ligne, COLONNE_ID)

'Calcul l'adresse destination
DestinationID = CDbl("&H" & ID) And CDbl("&H0FF0000")

If WorksheetFunction.Bitrshift(Arg1:=DestinationID, Arg2:=16) < 240 Then       'Si le PS < 240 --> C'est une adresse destination
    DestinationID = CDbl("&H" & ID) And CDbl("&H0FF00")
    DestinationID = WorksheetFunction.Bitrshift(Arg1:=DestinationID, Arg2:=8)
Else                                                                         'Sinon >=240 c'est tout le monde , c'est la continuité du PGN
    DestinationID = CDbl("&H" & ID) And CDbl("&H0FF0000")
    DestinationID = WorksheetFunction.Bitrshift(Arg1:=DestinationID, Arg2:=16)
End If

End Function



