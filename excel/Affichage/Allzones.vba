'Module dev par Leonard
Sub Allzones()
    '=================================Init=================================
    ' Masquer l'interface utilisateur pour afficher le classeur en plein écran
    Application.DisplayFullScreen = True
    Application.CommandBars("Worksheet Menu Bar").Enabled = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayScrollBars = False
    Application.DisplayAlerts = False
    Application.CommandBars("Full Screen").Visible = False
    '=================================Init Var=================================
    'Changer de feuille
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
    Dim lastRow As Integer
    lastRow = lastCell.row '(Dernière ligne remplie)
    ' Définir la ligne de début pour la copie des données
    Dim startRow As Long
    startRow = 4
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
    StopCodeAcc = False
    'Mettre la police a 20 et centrer le texte des cellules la feuille de source
    With sourceSheet.Cells
        .Font.Size = 20
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    'Affichage du
    'Affichage du numéro & jours de la semaine
    sourceSheet.Range("G1").MergeArea.Copy Destination:=destinationSheet.Range("G2")
    sourceSheet.Range("G3:M3").Copy Destination:=destinationSheet.Range("G4:M4")
    '=======================================================================
    'Clear tout
    With destinationSheet.Range("A" & destRow & ":M33")
        .ClearContents
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlNone ' supprimer les bordures
        'Défusionner les cellules avant de copier
        .UnMerge
    End With
    
    ' Afficher "Données de toute les zones" en haut de la feuille de destination
    With destinationSheet.Range("A1:M1")
        .Merge
        .Value = "Données de toute les zones"
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 26
    End With

    ' Boucle pour mettre à jour la plage visible
    Do While True
        ' Parcourir chaque ligne de la feuille de source
        For i = startRow To lastRow
            DoEvents
            If StopCodeAcc Then Exit Do ' si StopCodeAcc est True, sortir de la boucle
            ' Si la valeur de la colonne A correspond à ValChosenBat, copier les données
            
                If destRow <= 33 Then
                    With destinationSheet.Range("A" & destRow & ":M33")
                        sourceSheet.Range("A" & i & ":M" & i).Copy Destination:=destinationSheet.Range("A" & destRow & ":M" & destRow)
                        destRow = destRow + 1
                        rowCounter = rowCounter + 1
                    End With
               Else
                    Application.Wait (Now + TimeValue("0:00:15")) ' Attendre 15 secondes
                    destRow = 5
                    With destinationSheet.Range("A" & destRow & ":M33")
                        .ClearContents
                        .Interior.Color = RGB(255, 255, 255)
                        .Borders.LineStyle = xlNone
                        sourceSheet.Range("A" & i & ":M" & i).Copy Destination:=destinationSheet.Range("A" & destRow & ":M" & destRow)
                        destRow = destRow + 1
                        rowCounter = rowCounter + 1
                    End With
            End If
            
            If lastRow - 3 > 33 Then
                If rowCounter = lastRow - 3 Then
                    fullCells = True
                    Application.Wait (Now + TimeValue("0:00:15")) ' Attendre 15 secondes
                    Exit Do
                End If
            Else
                If rowCounter = lastRow Then Exit Do
            End If
        Next i
    Loop
    
    If fullCells = True Then
        startRow = 4
        destRow = 5
        ' Tout supprimer avant d'afficher de nouvelles données
        With destinationSheet.Range("A" & destRow & ":M33")
            .ClearContents
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlNone ' supprimer les bordures
            ' Défusionner les cellules avant de copier
            destinationSheet.Range("A" & destRow & ":M" & destRow).UnMerge
            fullCells = False
            Allzones
        End With
    End If
    
    StopCodeAcc = False
    fullCells = False
    Workbooks(1).RefreshAll ' Refresh le document
End Sub

