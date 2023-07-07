Public ValChosenBat As String
Public StopCodeAcc As Boolean

Sub Boutons_choix_bat()
    Dim btn As Object
    Set btn = ActiveSheet.Buttons(Application.Caller)
    ValChosenBat = btn.Text
    Sheets("Affichage").Activate
    Affichage
End Sub
Sub StopCode()
    StopCodeAcc = True
    Sheets("Acceuil").Activate
End Sub
Sub Multibatbut()
    Dim btn As Object
    Set btn = ActiveSheet.Buttons(Application.Caller)
    ValChosenBat = btn.Text
    Sheets("Multibat Affichage").Activate
    Multibat
End Sub

