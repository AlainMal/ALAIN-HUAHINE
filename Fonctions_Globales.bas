Attribute VB_Name = "Fonctions_Globales"
Option Explicit

'Fonction d'ouverture "FileDialog"
'Vient d'internet
Public Function UseFileDialogOpen() As String

Dim lngCount As Long
 Dim tableau()  As String
 Dim nom As String
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        '.AllowMultiSelect = True            On n'en veut pas
        .Title = "Selectionnez un fichier sur USB CAN"
        .Filters.Clear
        .Filters.Add "Texte", "*.txt"
        .Show
 
 'Ne traite pas les erreurs
 On Error Resume Next
        
        ' Display paths of each file selected
        'For lngCount = 1 To .SelectedItems.Count
            UseFileDialogOpen = .SelectedItems(1)
        'Next lngCount
 
tableau = Split(UseFileDialogOpen, "\")
nom = tableau(UBound(tableau, 1))
User_Form_PGN.Txt = nom
Sheets("Memoires").Range("A1") = nom
    End With
 
End Function

'Fonction pour ecrire un fichier texte
Function CreerFichierTexte() As String
Dim fso As Object
Dim FichierTexte As Object

CreerFichierTexte = Application.GetSaveAsFilename(FileFilter:="Text Files (*.txt), *.txt", Title:="ENREGISTRER UN FICHIER TEXTE POUR LE BUS CAN")

'Set fso = CreateObject("Scripting.FileSystemObject")
'On Error Resume Next
'Set FichierTexte = fso.CreateTextFile(CreerFichierTexte, True)
'On Error GoTo 0

End Function

