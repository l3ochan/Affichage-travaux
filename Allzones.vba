'Module dev par Leonard
Sub Allzones()
'=============================Init==========================================
    ' Masquer l'interface utilisateur pour afficher le classeur en plein écran
    Application.DisplayFullScreen = True
    Application.CommandBars("Worksheet Menu Bar").Enabled = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayScrollBars = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    Sheets("Multibat Affichage").Activate
    

    ' Définir la source et la destination
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("Source Affichage")
    Set destinationSheet = ThisWorkbook.Sheets("Multibat Affichage")
    ' Trouver la dernière cellule non vide dans la colonne spécifiée
    Dim lastCell As Range
    Set lastCell = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp)

    ' Obtenir le numéro de la dernière ligne
    Dim lastRow As Long
    lastRow = lastCell.row

    ' Définir la plage de cellules visible dans la feuille de destination
    Dim visibleRange As Range
    Set visibleRange = destinationSheet.Range("A4:M33")
    'Afficher le nom de la zone qui concerne les données
    With destinationSheet.Range("A1:M1")
        .Merge
        .Value = "Données pour toute les zones"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 26
    End With

    ' Définir la ligne de début pour la copie des données
    Dim startRow As Long
    startRow = 4
    StopCodeAcc = False
    'Mettre la police a 20 de la feuille de source
    sourceSheet.Cells.Font.Size = 20
    'Affichage du numéro & jours de la semaine
    sourceSheet.Range("G1").MergeArea.Copy Destination:=destinationSheet.Range("G2")
    sourceSheet.Range("G3:M3").Copy Destination:=destinationSheet.Range("G4:M4")
    '=======================================================================
    ' Boucle pour mettre à jour la plage visible
    Do While True
        DoEvents
        If StopCodeAcc Then Exit Do
        ' Vider cellules
        With destinationSheet.Cells.Range("A5:F33")
            .ClearContents
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlNone ' supprimer les bordures
        End With
        ' Copier les données dans la plage visible
        sourceSheet.Range("A" & startRow & ":M" & (startRow + 30)).Copy Destination:=visibleRange
        DoEvents
        ' Incrementer la ligne de départ
        startRow = startRow + 33
        If startRow > lastRow Then
            startRow = 4
        End If

        ' Attendre 1 seconde avant de mettre à jour à nouveau la plage visible
        Application.Wait (Now + TimeValue("0:00:10"))
    Loop
ThisWorkbook.RefreshAll
End Sub




