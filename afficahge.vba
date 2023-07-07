'Module dev par Leonard
Sub Affichage()
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
    Dim sourceSheet As Worksheet
    Dim destinationSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("Source Affichage")
    Set destinationSheet = ThisWorkbook.Sheets("Affichage")
    ' Trouver la dernière cellule non vide dans la colonne spécifiée
    Dim lastCell As Range
    Set lastCell = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp)
    ' Obtenir le numéro de la dernière ligne
    Dim lastRow As Long
    lastRow = lastCell.row '(Dernière ligne remplie)
    ' Définir la ligne de début pour la copie des données
    Dim startRow As Long
    startRow = 3
    ' Définir la ligne de destination
    Dim destRow As Long
    destRow = 5
    'Compteur de lignes de la feuille d'origine
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
    For Each cell In sourceSheet.Range("A3:A" & lastRow)
        If UCase(cell.Value) = UCase(ValChosenBat) Then ' Utilise UCase pour ignorer la casse
            corespondingRow = corespondingRow + 1
        End If
    Next cell
    StopCodeAcc = False
    'Mettre la police a 20 de la feuille de source
    sourceSheet.Cells.Font.Size = 20
    'Affichage du numéro & jours de la semaine
    sourceSheet.Range("G1").MergeArea.Copy Destination:=destinationSheet.Range("F2")
    sourceSheet.Range("G3:M3").Copy Destination:=destinationSheet.Range("F4:L4")
    '=======================================================================
    'Clear tout
    With destinationSheet.Range("A" & destRow & ":L33")
        .ClearContents
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlNone ' supprimer les bordures
        'Défusionner les cellules avant de copier
        .UnMerge
        
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
        With destinationSheet.Range("A" & destRow & ":L33")
            .Merge
            .Value = "Aucune entrée pour la zone: " & ValChosenBat
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
            .Font.Size = 26
            .Interior.Color = RGB(217, 217, 217)
        End With
    End If
    ' Boucle pour mettre à jour la plage visible
    Do While True
        ' Parcourir chaque ligne de la feuille de source
        For i = startRow To lastRow
            DoEvents
            If StopCodeAcc Then Exit Do ' si StopCodeAcc est True, sortir de la boucle
            ' Si la valeur de la colonne A correspond à ValChosenBat, copier les données
            If UCase(sourceSheet.Cells(i, "A").Value) = UCase(ValChosenBat) Then
                If destRow <= 33 Then
                    With destinationSheet.Range("A" & destRow & ":L33")
                        sourceSheet.Range("B" & i & ":M" & i).Copy Destination:=destinationSheet.Range("A" & destRow & ":J" & destRow)
                        destRow = destRow + 1
                        rowCounter = rowCounter + 1
                    End With
                Else
                    Application.Wait (Now + TimeValue("0:00:10")) ' Attendre 10 secondes
                    destRow = 5
                    With destinationSheet.Range("A" & destRow & ":L33")
                        sourceSheet.Range("B" & i & ":M" & i).Copy Destination:=destinationSheet.Range("A" & destRow & ":J" & destRow)
                        destRow = destRow + 1
                        rowCounter = rowCounter + 1
                    End With
                End If
            End If
                If corespondingRow > 33 Then
                    If (rowCounter = corespondingRow) Then
                        fullCells = True
                        Application.Wait (Now + TimeValue("0:00:10")) ' Attendre 10 secondes
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
    With destinationSheet.Range("A" & destRow & ":L33")
        .ClearContents
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlNone ' supprimer les bordures
        'Défusionner les cellules avant de copier
        destinationSheet.Range("A" & destRow & ":L" & destRow).UnMerge
        fullCells = False
        Affichage
    End With
End If
StopCodeAcc = False
fullCells = False
ThisWorkbook.RefreshAll ' Refresh le document
End Sub


