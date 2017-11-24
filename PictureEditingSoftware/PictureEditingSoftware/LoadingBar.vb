Imports System.Windows.Forms

Public Class LoadingBar

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub LoadingBar_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Location = New Point((Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2, (Screen.PrimaryScreen.Bounds.Height - Me.Height) / 2)
    End Sub
End Class
