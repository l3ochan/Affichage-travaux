Sub Allzones()
'=============================Init==========================================
    ' Masquer l'interface utilisateur pour afficher le classeur en plein écran
    Application.DisplayFullScreen = True
    Application.CommandBars("Worksheet Menu Bar").Enabled = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayScrollBars = False
    Application.DisplayAlerts = False
    Sheets("AllZones Affichage").Activate
    

    ' Définir la source et la destination
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("Source Affichage")
    Set destinationSheet = ThisWorkbook.Sheets("AllZones Affichage")
    ' Trouver la dernière cellule non vide dans la colonne spécifiée
    Dim lastCell As Range
    Set lastCell = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp)

    ' Obtenir le numéro de la dernière ligne
    Dim lastRow As Long
    lastRow = lastCell.row

    ' Définir la plage de cellules visible dans la feuille de destination
    Dim visibleRange As Range
    Set visibleRange = destinationSheet.Range("A2:M33")

    ' Définir la ligne de début pour la copie des données
    Dim startRow As Long
    startRow = 2
    sourceSheet.Cells.Font.Size = 20
    sourceSheet.Range("G1").MergeArea.Copy Destination:=destinationSheet.Range("F2")
    '=======================================================================
    ' Boucle pour mettre à jour la plage visible
    Do While True
        DoEvents
        ' Vider cellules
        With destinationSheet.Cells.Range("A4:F33")
            .ClearContents
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlNone ' supprimer les bordures
        End With
        ' Copier les données dans la plage visible
        sourceSheet.Range("A" & startRow & ":M" & (startRow + 30)).Copy Destination:=visibleRange

        ' Incrementer la ligne de départ
        startRow = startRow + 33
        If startRow > lastRow Then
            startRow = 4
        End If

        ' Attendre 1 seconde avant de mettre à jour à nouveau la plage visible
        Application.Wait (Now + TimeValue("0:00:10"))
    Loop

End Sub



