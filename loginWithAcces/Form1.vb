Public Class Form1

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Environment.Exit(10)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = conn
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = "SELECT * FROM tb_login WHERE username = '" & TextBox1.Text & "'AND password = '" & TextBox2.Text & "'"
        RD = objcmd.ExecuteReader()
        If RD.HasRows Then
            MsgBox("Login berhasil", vbInformation, "Aplikasi input data siswa")
            Form2.Show()
            Me.Hide()
        Else
            MsgBox("Maaf Username atau password salah")
        End If
        TextBox1.Text = ""
        TextBox2.Text = ""

    End Sub
End Class
