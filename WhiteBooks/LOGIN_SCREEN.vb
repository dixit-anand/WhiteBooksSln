Public Class LOGIN_SCREEN
    Private Sub OK_BUTTON_Click(sender As Object, e As EventArgs) Handles OK_BUTTON.Click
        MsgBox("hello world")
        'MAIN_SCREEN.Show()
    End Sub

    Private Sub CANCEL_Click(sender As Object, e As EventArgs) Handles CANCEL.Click
        Application.Exit()
    End Sub
End Class