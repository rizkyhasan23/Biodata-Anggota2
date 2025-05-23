Public Class Form1
    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        If txtUser.Text = "rizkyhasan" And txtPass.Text = "rizky123" Then
            Me.Hide()
            Form2.Show()
        Else
            MessageBox.Show("Username atau Password salah!")
        End If
    End Sub
End Class
