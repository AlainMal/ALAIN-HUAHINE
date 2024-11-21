VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cMMSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Position_MMSI As String 'Position dans la feuille du MMSI ex.$A$2)
Dim Nom_Bateau As String    'Nom du bateau sur plusieurs lignes
Dim Ligne_mmsi As Integer  'Numéro de ligne du MMSI

Public Sub Affiche_MMSI(ligne As Long)
Dim MMSI As Range

'Compte le nombre de ligne existante
Ligne_mmsi = Worksheets(FEUIL_MMSI).Cells(Rows.Count, 1).End(xlUp).Row - 1

'En fonction du PGN issu du bus CAN sur la ligne en cours
Select Case cID.Pgn(ligne)
    
    Case 129038 'Position MMSI Class A
        If ValeurChoisie = 0 Then       'Dans VeleurChoisie on a le numéro de ligne donc la ligne est "0"
            If Trouve_MMSI(Format(ValeurChoisie2, "###\ ###\ ###")) Is Nothing Then  'Crée la ligne MMSI
                Worksheets(FEUIL_MMSI).Range(COLONNE_NUM_MMSI & Ligne_mmsi + 2).Value = Format(ValeurChoisie2, "###\ ###\ ###") 'MMSI
                Worksheets(FEUIL_MMSI).Range(COLONNE_CLASSE_MMSI & Ligne_mmsi + 2).Value = "A" 'Classe A
                Ligne_mmsi = Ligne_mmsi + 1
            End If
            'On a trouvé le MMSI et on calcul sa position
            Position_MMSI = Trouve_MMSI(Format(ValeurChoisie2, "###\ ###\ ###")).Address  'Ex. Position_MMSI = $A$10 on remplace A par B
            'On met à jour l'heure sur la colonne "Heure"
            Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_HEURE_MMSI)).Value = Time
            'Tri
        End If
          
        If ValeurChoisie = 1 And Position_MMSI <> "" Then  'Ligne "1", si on a un MMSI
            If ValeurChoisie2 <> "" Then  'S'il y a quelque chose sur ValeurChoisie2 on ecrit la lattitude
                Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_LATITUDE_MMSI)).Value = Format(ValeurChoisie2, "0.0000") & IIf(ValeurChoisie2 < 180, " N", " S")
            End If
            If ValeurChoisie3 <> "" Then   'S'il y a quelque chose sur ValeurChoisie3 on ecrit la longitude
                Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_LONGITUDE_MMSI)).Value = Format(ValeurChoisie3, "0.0000") & IIf(ValeurChoisie3 < 180, " E", " O")
                'Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, "j")).Value = ((Ma_position_longitude * Cos(ValeurChoisie3)) ^ 2) ^ 0.5
            End If
        End If
        
        If ValeurChoisie = 2 And Position_MMSI <> "" Then 'Ligne 2, si on a la position du MMSI dans la feuille
            If ValeurChoisie2 <> "" Then 'On écrit le COG
                Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_COG_MMSI)).Value = Format(ValeurChoisie2, "0.00")
            End If
            If ValeurChoisie3 <> "" Then 'On ecrit le SOG
                Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_SOG_MMSI)).Value = Format(ValeurChoisie3, "0.00")
            End If
      End If
    
    Case 129794 'Données MMSI Class A
        If ValeurChoisie = 0 Then
            If Trouve_MMSI(Format(ValeurChoisie2, "###\ ###\ ###")) Is Nothing Then
                Worksheets(FEUIL_MMSI).Range(PREMIERE_COLONNE_MMSI & Ligne_mmsi + 2).Value = Format(ValeurChoisie2, "###\ ###\ ###")
                Worksheets(FEUIL_MMSI).Range(COLONNE_CLASSE_MMSI & Ligne_mmsi + 2).Value = PREMIERE_COLONNE_MMSI
                Ligne_mmsi = Ligne_mmsi + 1
            End If
            Position_MMSI = Trouve_MMSI(Format(ValeurChoisie2, "###\ ###\ ###")).Address
            Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_HEURE_MMSI)).Value = Time
        End If
          
        If ValeurChoisie = 2 And Position_MMSI <> "" Then 'Ligne 2
            If ValeurChoisie2 <> "" Then 'On commence à ecrire le nom du bateau
                Nom_Bateau = Format(ValeurChoisie2, "#")
                Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_NOM_MMSI)).Value = Trim(Format(ValeurChoisie2, "#"))
            End If
        End If
        If ValeurChoisie = 3 And Position_MMSI <> "" Then  'Ligne 3
            If ValeurChoisie2 <> "" Then 'On ecrit la suite du nom du bateau
                Nom_Bateau = Nom_Bateau & Format(ValeurChoisie2, "#")
                Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_NOM_MMSI)).Value = Trim(Nom_Bateau)
            End If
        End If
        If ValeurChoisie = 4 And Position_MMSI <> "" Then  'Ligne 4
            If ValeurChoisie2 <> "" Then 'On ecrit la suite du nom du bateau
                Nom_Bateau = Nom_Bateau & Format(ValeurChoisie2, "#")
                Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_NOM_MMSI)).Value = Trim(Nom_Bateau)
            End If
        End If

        If ValeurChoisie = 6 And Position_MMSI <> "" Then
            If ValeurChoisie2 <> "" Then 'On ecrit la distance du bateau
                Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_DISTANCE_MMSI)).Value = Format(ValeurChoisie2, "0.00")
            End If
        End If
      
    Case 129039 'Position MMSI Class B
        If ValeurChoisie = 0 Then
            If Trouve_MMSI(Format(ValeurChoisie2, "###\ ###\ ###")) Is Nothing Then  'MMSI
                Worksheets(FEUIL_MMSI).Range(PREMIERE_COLONNE_MMSI & Ligne_mmsi + 2).Value = Format(ValeurChoisie2, "###\ ###\ ###")
                Worksheets(FEUIL_MMSI).Range(COLONNE_CLASSE_MMSI & Ligne_mmsi + 2).Value = "B"
                Ligne_mmsi = Ligne_mmsi + 1
            End If
            Position_MMSI = Trouve_MMSI(Format(ValeurChoisie2, "###\ ###\ ###")).Address
            Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_HEURE_MMSI)).Value = Time
            'Tri
        End If
        
        If ValeurChoisie = 1 And Position_MMSI <> "" Then
            'IL y a une fausse valeur >80
            If ValeurChoisie2 <> "" Then 'Longitude             .... Si la positon et en dehors des clous (+/- 1 degrés)(60 milles d'écart)
                If Not (ValeurChoisie2 > Ma_position_longitude + 1) And Not (ValeurChoisie2 < Ma_position_longitude - 1) Then
                    Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_LONGITUDE_MMSI)).Value = Format(ValeurChoisie2, "0.0000") & IIf(ValeurChoisie2 < 180, " E", " O")
                End If
            End If
            If ValeurChoisie3 <> "" Then 'Latitude
                Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_LATITUDE_MMSI)).Value = Format(ValeurChoisie3, "0.0000") & IIf(ValeurChoisie3 < 180, " N", " S")
            End If
        End If

        
        If ValeurChoisie = 2 And Position_MMSI <> "" Then
            If ValeurChoisie2 <> "" Then  'COG
                Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_COG_MMSI)).Value = Format(ValeurChoisie2, "0.00")
            End If
            If ValeurChoisie3 <> "" Then 'SOG
                Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_SOG_MMSI)).Value = Format(ValeurChoisie3, "0.00")
            End If
        End If


    Case 129809 'Données MMSI Class B
        If ValeurChoisie = 0 Then  'MMSI
            If Trouve_MMSI(Format(ValeurChoisie2, "###\ ###\ ###")) Is Nothing Then
                Worksheets(FEUIL_MMSI).Range(PREMIERE_COLONNE_MMSI & Ligne_mmsi + 2).Value = Format(ValeurChoisie2, "###\ ###\ ###")
                Nom_Bateau = Format(ValeurChoisie3, "#")
                Worksheets(FEUIL_MMSI).Range(COLONNE_CLASSE_MMSI & Ligne_mmsi + 2).Value = "B"
                Ligne_mmsi = Ligne_mmsi + 1
            End If
            Position_MMSI = Trouve_MMSI(Format(ValeurChoisie2, "###\ ###\ ###")).Address
            
            'Nom du bateau, les autres sont contaténés (&)
            Nom_Bateau = Format(ValeurChoisie3, "#")
            
            Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_NOM_MMSI)).Value = Trim(Format(ValeurChoisie3, "#"))
            Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_HEURE_MMSI)).Value = Time
        End If
          
        If ValeurChoisie = 1 And Position_MMSI <> "" Then  'Nom du bateau
            If ValeurChoisie3 <> "" Then
                Nom_Bateau = Nom_Bateau & Format(ValeurChoisie3, "#")
                Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_NOM_MMSI)).Value = Trim(Nom_Bateau)
            End If
        End If
        If ValeurChoisie = 2 And Position_MMSI <> "" Then  'Nom du bateau
            If ValeurChoisie3 <> "" Then
                Nom_Bateau = Nom_Bateau & Format(ValeurChoisie3, "#")
                Worksheets(FEUIL_MMSI).Range(Replace(Position_MMSI, COLONNE_NUM_MMSI, COLONNE_NOM_MMSI)).Value = Trim(Nom_Bateau)
            End If
      End If

    End Select

End Sub

'Trouve le MMSI
Private Function Trouve_MMSI(Str As String) As Range

Dim wb As Workbook
Dim s1 As Worksheet
Dim rng As Range

Set wb = ThisWorkbook
Set s1 = wb.Sheets(FEUIL_MMSI)

Set rng = s1.Columns("A:A")
Set Trouve_MMSI = rng.Find(What:=Str, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False) 'donne la veleur dans Range.Address ex. $A$2

End Function

'Réalise le trie du plus récent au plus ancien
Public Sub Tri()
On Error Resume Next 'Il peut avoir des erreurs lorsqu'on entre dans le programme sinon il n'y a pas.
Worksheets(FEUIL_MMSI).Sort.SortFields.Clear
Range("A1:I300").Sort Key1:=Range("B1"), Header:=xlYes
On Error GoTo 0
End Sub
