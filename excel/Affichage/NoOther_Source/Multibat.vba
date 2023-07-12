'Module dev par Leonard
Sub Multibat()
    '=================================Init=================================
    ' Masquer l'interface utilisateur pour afficher le classeur en plein écran
    Application.DisplayFullScreen = True
    Application.CommandBars("Worksheet Menu Bar").Enabled = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayScrollBars = False
    Application.DisplayAlerts = False
    Application.CommandBars("Full Screen").Visible = False
    '=================================Init Var=================================
    ' Définir la source et la destination
    Dim destColumn As Integer 'Colonne de destination pour les jours et semaines
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("Planning commun des travaux DDP")
    Set destinationSheet = ThisWorkbook.Sheets("Multibat Affichage")
    ' Trouver la dernière cellule non vide dans la colonne spécifiée
    Dim lastCell As Range
    Set lastCell = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp)
    ' Obtenir le numéro de la dernière ligne
    Dim lastRow As Long
    lastRow = lastCell.Row '(Dernière ligne remplie)
    ' Définir la ligne de début pour la copie des données
    Dim startRow As Long
    startRow = 3
    ' Définir la ligne de destination
    Dim destRow As Long
    destRow = 5
    'Compteur de lignes de la feuille de destination
    Dim rowCounter As Integer
    rowCounter = 0
    'état tableau plein
    Dim fullCells As Boolean
    fullCells = False
    'nombre de cellules correspondantes au batiment choisi
    Dim corespondingRow As Integer
    corespondingRow = 0
    'Comptage des cellules correspondantes
    Dim cell As Range
    Dim cellD As Range
    For i = 3 To lastRow
        Set cell = sourceSheet.Cells(i, "A")
        Set cellD = sourceSheet.Cells(i, "D")
    If LCase(cell.Value) Like "*" & LCase(ValChosenBat) & "*" Then
            If UCase(cellD.Value) = "EN COURS" Or UCase(cellD.Value) = "A LANCER" Then ' Utilise UCase pour ignorer la casse
                corespondingRow = corespondingRow + 1
            End If
        End If
    Next i
    StopCodeAcc = False
    'Mettre la police a 20 et centrer le texte des cellules la feuille de source
    With sourceSheet.Cells
        .Font.Size = 20
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    '=======================================================================
    'Clear tout
    With destinationSheet.Range("A" & destRow & ":M33")
        .UnMerge
        .ClearContents
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlNone ' supprimer les bordures
        'Défusionner les cellules avant de copier
    End With
    
    ' Afficher "Données pour le bâtiment: ValChosenBat" en haut de la feuille de destination
    With destinationSheet.Range("A1:K1")
        .Merge
        .Value = "Données pour la zone: " & ValChosenBat
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 26
    End With
    ' Si aucune donnée n'a été trouvée, afficher le message d'erreur
    If corespondingRow = 0 Then
        With destinationSheet.Range("A" & destRow & ":M33")
            .Merge
            .Value = "Aucune entrée pour la zone: " & ValChosenBat
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
            .Font.Size = 26
            .Interior.Color = RGB(217, 217, 217)
        End With
    End If
    'Afficher le numéro de semaine
    Set sourceRange = sourceSheet.Range("M1:NS1")
    destColumn = 7 ' Start at column F Le numéro correspond a la lettre
    destinationSheet.Range("G2:M2").UnMerge
    For Each cell In sourceRange.SpecialCells(xlCellTypeVisible)
        cell.Copy Destination:=destinationSheet.Cells(2, destColumn)
        destColumn = destColumn + 1
    Next cell
    destinationSheet.Range("G2:M2").Merge
    'Afficher les jours de la semaine
    Set sourceRange = sourceSheet.Range("M3:NS3")
    destColumn = 7 ' Start at column F Le numéro correspond a la lettre
    For Each cell In sourceRange.SpecialCells(xlCellTypeVisible)
        cell.Copy Destination:=destinationSheet.Cells(4, destColumn)
        destColumn = destColumn + 1
    Next cell
        ' Parcourir chaque ligne de la feuille de source
        Do While True
            ' Parcourir chaque ligne de la feuille de source
            For i = startRow To lastRow
                DoEvents
                If StopCodeAcc Then Exit Do ' si StopCodeAcc est True, sortir de la boucle
                ' Si la valeur de la colonne A correspond partiellement à ValChosenBat, copier les données
                If LCase(sourceSheet.Cells(i, "A").Value) Like "*" & LCase(ValChosenBat) & "*" Then
                    If UCase(sourceSheet.Cells(i, "D").Value) = "EN COURS" Or UCase(sourceSheet.Cells(i, "D").Value) = "A LANCER" Then
                        If destRow <= 33 Then
                            With destinationSheet.Range("A" & destRow & ":L" & destRow)
                                sourceSheet.Range("A" & i & ":F" & i).Copy Destination:=destinationSheet.Range("A" & destRow)
                                ' Copy visible cells from M:NS
                                Set sourceRange = sourceSheet.Range("M" & i & ":NS" & i)
                                destColumn = 7 ' Start at column F Le numéro correspond a la lettre
                                For Each cell In sourceRange.SpecialCells(xlCellTypeVisible)
                                    cell.Copy Destination:=destinationSheet.Cells(destRow, destColumn)
                                    destColumn = destColumn + 1
                                Next cell
                                destRow = destRow + 1
                                rowCounter = rowCounter + 1
                            End With
                        Else
                            Application.Wait (Now + TimeValue("0:00:15")) ' Attendre 15 secondes
                            destRow = 5
                            With destinationSheet.Range("A" & destRow & ":L" & destRow)
                                .ClearContents
                                .Interior.Color = RGB(255, 255, 255)
                                .Borders.LineStyle = xlNone
                                sourceSheet.Range("A" & i & ":F" & i).Copy Destination:=destinationSheet.Range("A" & destRow)
                                ' Copy visible cells from M:NS
                                Set sourceRange = sourceSheet.Range("M" & i & ":NS" & i)
                                destColumn = 7 ' Start at column F Le numéro correspond a la lettre
                                For Each cell In sourceRange.SpecialCells(xlCellTypeVisible)
                                    cell.Copy Destination:=destinationSheet.Cells(destRow, destColumn)
                                    destColumn = destColumn + 1
                                Next cell
                                destRow = destRow + 1
                                rowCounter = rowCounter + 1
                            End With
                        End If
                    End If
                End If
                If corespondingRow > 33 Then
                    If (rowCounter = corespondingRow) Then
                        fullCells = True
                        Application.Wait (Now + TimeValue("0:00:15")) ' Attendre 15 secondes
                        Exit Do
                    End If
                Else
                    If rowCounter = corespondingRow Then Exit Do
                End If
            Next i
        Loop
If fullCells = True Then
    startRow = 3
    destRow = 5
    'Tout supprimer avant d'afficher de nouvelles données
    With destinationSheet.Range("A" & destRow & ":M33")
        .ClearContents
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlNone ' supprimer les bordures
        'Défusionner les cellules avant de copier
        destinationSheet.Range("A" & destRow & ":E" & destRow).UnMerge
        fullCells = False
        Multibat
    End With
End If
StopCodeAcc = False
fullCells = False
ThisWorkbook.RefreshAll ' Refresh le document
End Sub
