Imports System.Drawing

Public Class Forms


    Public Function NewWindow() As Window
        Return New Window()
    End Function


    Private WithEvents pSubform As TransparentWindow

    Public Sub ShowTransparentWindow(ByVal message As String, Optional ByVal fadingTime As Long = 2)
        pSubform = New TransparentWindow()
        With pSubform
            .ControlBox = False
            .FormBorderStyle = Windows.Forms.FormBorderStyle.None
            .Text = String.Empty
            .BackColor = Color.Red
            Call .setTransparentLayeredWindow()
            Call .Show()
            '.ShowInTaskbar = False
            .SeekSpeed = 50
            .StepsToFade = 20 '(fadingTime * 1000) / 50
            Call .seekTo(0)
            'Call .updateOpacity(100, True)
        End With
    End Sub


    Private Sub pSubform_MouseEnter(sender As Object, e As EventArgs) Handles pSubform.MouseEnter
        Call pSubform.stopTimer()
    End Sub

    Private Sub pSubform_MouseLeave(sender As Object, e As EventArgs) Handles pSubform.MouseLeave
        Call pSubform.seekTo(0)
    End Sub

    Private Sub pSubform_OnFadeComplete() Handles pSubform.OnFadeComplete
        Call pSubform.Hide()
        pSubform = Nothing
    End Sub


End Class