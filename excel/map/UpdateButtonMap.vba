'module dev par Leonard
Sub UpdateButtonMap()
    Dim homeSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim btn As Object
    Dim cell As Range
    Dim lastRow As Long
    Dim hasWork As Boolean

    Set homeSheet = ThisWorkbook.Sheets("Accueil Affichage")
    Set sourceSheet = ThisWorkbook.Sheets("Planning commun des travaux DDP")
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Parcourir tous les boutons sur la feuille d'accueil
    For Each btn In homeSheet.Buttons
        hasWork = False
        ' Vérifier si le nom du bouton correspond à un bâtiment avec du travail
        For Each cell In sourceSheet.Range("A3:A" & lastRow)
            ' Utilise InStr pour une correspondance partielle sensible à la casse
            ' Ajoute des espaces autour du nom du bâtiment pour éviter les correspondances partielles indésirables
            If InStr(1, " " & cell.Value & " ", " " & btn.Text & " ", vbBinaryCompare) > 0 And _
               (UCase(sourceSheet.Cells(cell.Row, "D").Value) = "EN COURS" Or _
                UCase(sourceSheet.Cells(cell.Row, "D").Value) = "A LANCER") Then
                hasWork = True
                Exit For
            End If
        Next cell
        ' Changer la mise en forme du bouton en fonction de la variable hasWork
        If hasWork Then
            With btn.Font
                .Bold = True
                .Underline = True
                .ColorIndex = 3 ' Rouge
            End With
        Else
            With btn.Font
                .Bold = False
                .Underline = False
                .ColorIndex = 1 ' Noir
            End With
        End If
    Next btn
End Sub


