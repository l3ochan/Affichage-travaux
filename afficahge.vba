Sub Affichage()
    ' Masquer l'interface utilisateur pour afficher le classeur en plein écran
    Application.DisplayFullScreen = True
    Application.CommandBars("Worksheet Menu Bar").Enabled = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayScrollBars = False
    Application.DisplayAlerts = False

    ' Définir la source et la destination
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("Source Affichage")
    Set destinationSheet = ThisWorkbook.Sheets("Affichage")

    ' Trouver la dernière cellule non vide dans la colonne spécifiée
    Dim lastCell As Range
    Set lastCell = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp)

    ' Obtenir le numéro de la dernière ligne
    Dim lastRow As Long
    lastRow = lastCell.row

    ' Définir la ligne de début pour la copie des données
    Dim startRow As Long
    startRow = 3

    ' Définir la ligne de destination
    Dim destRow As Long
    destRow = 4

    ' Afficher "Données pour le bâtiment: ValChosenBat" en haut de la feuille de destination
    destinationSheet.Range("A1:K1").Merge
    destinationSheet.Range("A1").Value = "Données pour la zone: " & ValChosenBat
    destinationSheet.Range("A1").HorizontalAlignment = xlCenter
    destinationSheet.Range("A1").Font.Bold = True
    destinationSheet.Range("A1").Font.Size = 26
    
    ' Boucle pour mettre à jour la plage visible
    Do While True
        Dim dataFound As Boolean
        dataFound = False ' Nous supposons qu'aucune donnée ne sera trouvée à chaque début d'itération

        Dim rowCounter As Integer
        rowCounter = 0
        ' Parcourir chaque ligne de la feuille de source
        For i = startRow To lastRow
            ' Si la valeur de la colonne A correspond à ValChosenBat, copier les données
            If UCase(sourceSheet.Cells(i, "A").Value) = UCase(ValChosenBat) Then
                If destRow <= 38 Then
                    With destinationSheet.Range("A" & destRow & ":L38")
                        .ClearContents
                        .Interior.Color = RGB(255, 255, 255)
                        .Borders.LineStyle = xlNone ' supprimer les bordures
                    End With
                End If
                ' Défusionner les cellules avant de copier
                destinationSheet.Range("A" & destRow & ":L" & destRow).UnMerge
                sourceSheet.Range("B" & i & ":K" & i).Copy Destination:=destinationSheet.Range("A" & destRow & ":J" & destRow)
                destRow = destRow + 1
                startRow = i + 1
                rowCounter = rowCounter + 1
                dataFound = True ' Des données correspondantes ont été trouvées
            If destRow > 38 Then
                destRow = 4 ' Une fois la ligne 38 atteinte, repartir à la ligne 4
                destinationSheet.Range("A4:L38").UnMerge ' Défusionner avant de re-remplir
                Application.Wait (Now + TimeValue("0:00:10")) ' Attendre 10 secondes
                DoEvents
                Exit For ' Sortir de la boucle For
            End If
        Next i

        ' Si aucune donnée n'a été trouvée, afficher le message d'erreur
        If Not dataFound Then
            destinationSheet.Range("A" & destRow & ":L38").UnMerge
            destinationSheet.Range("A" & destRow).Value = "Aucune entrée pour la zone: " & ValChosenBat
            destinationSheet.Range("A" & destRow).HorizontalAlignment = xlCenter
            destinationSheet.Range("A" & destRow).Font.Bold = True
            destinationSheet.Range("A" & destRow).Font.Color = RGB(255, 0, 0)
            destinationSheet.Range("A" & destRow).Font.Size = 26
        End If

        ' Si startRow dépasse lastRow, réinitialiser startRow et destRow
        If startRow > lastRow Then
            startRow = 3
            destRow = 4
        End If

        ' Refresh le document
        Workbooks(1).RefreshAll

        If StopCodeAcc Then Exit Do ' si StopCodeAcc est True, sortir de la boucle
    Loop
    StopCodeAcc = False
End Sub

