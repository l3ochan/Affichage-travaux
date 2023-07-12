'Module dev par Leonard
Public ValChosenBat As String
Public StopCodeAcc As Boolean

'Bouton de choix de bâtiments simple (Ex Bat DA)
Sub Boutons_choix_bat()
    Dim btn As Object
    Set btn = ActiveSheet.Buttons(Application.Caller)
    ValChosenBat = btn.Text
    Sheets("Affichage").Activate
    Affichage
End Sub
'Bouton d'arrêt et retour a l'acceuil
Sub StopCode()
    StopCodeAcc = True
    Sheets("Accueil Affichage").Activate
    ActiveWindow.Zoom = 88
End Sub
'Bouton de choix de bâtiments double (Ex BA-BB)
Sub Multibatbut()
    Dim btn As Object
    Set btn = ActiveSheet.Buttons(Application.Caller)
    ValChosenBat = btn.Text
    Sheets("Multibat Affichage").Activate
    Multibat
End Sub


