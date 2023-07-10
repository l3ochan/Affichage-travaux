Sub Resetdisplay()
    Application.DisplayFullScreen = False
    Application.CommandBars("Worksheet Menu Bar").Enabled = True
    ActiveWindow.DisplayHeadings = True
    Application.DisplayScrollBars = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
End Sub
Sub Fullscreen()
    Application.DisplayFullScreen = True
    Application.CommandBars("Worksheet Menu Bar").Enabled = False
    ActiveWindow.DisplayHeadings = False
    Application.DisplayScrollBars = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
End Sub

