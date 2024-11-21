VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cIDa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'================ FONCTION PGN ==================
'
'La Fonction PGN(ID) retourne le PGN de l'ID du bus CAN
'La fonction SourceID(Numéro de la ligne) retourne l'adresse source
'La fonction DestinationID(Numéro de la ligne) retourne l'adresse de destination
'
'===============================================

Public Function Pgn(ID As String) As String
Dim PgnDbl As Double

Err.Clear
On Error Resume Next
'Fait le masque pour récuperer le PGN
PgnDbl = CDbl("&H" & ID) And CDbl("&H0FF0000")
If Err.Number Then Exit Function
On Error GoTo 0
If WorksheetFunction.Bitrshift(Arg1:=PgnDbl, Arg2:=16) < 240 Then   'Si le PS < 240
    PgnDbl = CDbl("&H" & ID) And CDbl("&H1FF0000")
    PgnDbl = WorksheetFunction.Bitrshift(Arg1:=PgnDbl, Arg2:=8)
Else
    PgnDbl = CDbl("&H" & ID) And CDbl("&H1FFFF00")
    PgnDbl = WorksheetFunction.Bitrshift(Arg1:=PgnDbl, Arg2:=8)
End If
    Pgn = CStr(PgnDbl)
End Function

' Fonction de Récupèration de l'adresse source dans l'ID
Public Function SourceID(ID As String) As String

'Fait le masque pour récuperer l'adresse source
SourceID = CDbl("&H" & ID) And CDbl("&H00000FF")

End Function

' Fonction de  Récupèration de l'adresse destination dans l'ID
Public Function DestinationID(ID As String) As String

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




