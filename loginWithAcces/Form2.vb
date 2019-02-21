Public Class Form2
    Dim sqlnya As String

    Sub panggildata()
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_siswa", conn)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_siswa")
        DataGridView1.DataSource = DS.Tables("tb_siswa")
        DataGridView1.Enabled = True
    End Sub


    Sub jalan()
        Dim odjcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        odjcmd.Connection = conn
        odjcmd.CommandType = CommandType.Text
        odjcmd.CommandText = sqlnya
        odjcmd.ExecuteNonQuery()
        odjcmd.Dispose()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""
    End Sub


    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ComboBox1.Items.Add("Laki-laki")
        ComboBox1.Items.Add("Perempuan")

        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        Call panggildata()

        Dim btn As New DataGridViewButtonColumn()
        DataGridView1.Columns.Add(btn)
        btn.HeaderText = "Hapus"
        btn.Text = "Hapus"
        btn.Name = "btn"
        btn.UseColumnTextForButtonValue = True

        Dim btn2 As New DataGridViewButtonColumn()
        DataGridView1.Columns.Add(btn2)
        btn2.HeaderText = "Edit"
        btn2.Text = "Edit"
        btn2.Name = "btn2"
        btn2.UseColumnTextForButtonValue = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        sqlnya = "INSERT INTO tb_siswa (nis, nama, jk, email) values('" & TextBox1.Text & "', '" & TextBox2.Text & "', '" & ComboBox1.Text & "', '" & TextBox4.Text & "')"
        Call jalan()
        MsgBox("Data Berhasil Tersimpan")
        Call panggildata()
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_siswa WHERE nama like '%" & TextBox3.Text & "%' ", conn)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_siswa")
        DataGridView1.DataSource = DS.Tables("tb_siswa")
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = 4 Then
            Dim i As Integer
            Dim hap As Integer
            i = DataGridView1.CurrentRow.Index
            TextBox1.Text = DataGridView1.Item(0, i).Value
            hap = MsgBox("Apakah Yakin Ingin Di Hapus?", vbYesNo, "Informasi")
            If hap = vbYes Then
                sqlnya = "DELETE FROM tb_siswa WHERE nis='" & TextBox1.Text & "'"
                Call jalan()
                MsgBox("Data Berhasil Terhapus")
                Call panggildata()
            End If
        End If

        If e.ColumnIndex = 5 Then
            Dim a As Integer
            Dim has As Integer
            a = DataGridView1.CurrentRow.Index
            TextBox1.Text = DataGridView1.Item(0, a).Value
            TextBox2.Text = DataGridView1.Item(1, a).Value
            ComboBox1.Text = DataGridView1.Item(2, a).Value
            TextBox4.Text = DataGridView1.Item(3, a).Value
            has = MsgBox("Apakah Data Ini Ingin di Ubah?", vbYesNo, "Informasi")
            If has = vbYes Then
                sqlnya = "UPDATE FROM tb_siswa SET nis='" & TextBox1.Text & "', nama='" & TextBox2.Text & "', jk='" & ComboBox1.Text & "', email='" & TextBox4.Text & "'"
                Call jalan()
                MsgBox("Data Berhasil di Ubah")
                Call panggildata()
            End If
        End If
    End Sub
End Class