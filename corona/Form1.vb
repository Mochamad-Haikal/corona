Public Class Form1
    Dim sqlnya As String
    Sub panggildata()
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_corona", conn)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_corona")
        DataGridView1.DataSource = DS.Tables("tb_corona")
        DataGridView1.Enabled = True
    End Sub
    Sub jalan()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = conn
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlnya
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        sqlnya = "insert into tb_corona(nama,umur,gol_darah,point)values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "')"
        Call jalan()
        MsgBox("Data berhasil tersimpan")
        Call panggildata()
        '---------------------------------------------------------checkbox---------------------------------------------------'
        If CheckBox1.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox2.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox3.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox4.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox5.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox6.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox7.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox8.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox9.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox10.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox11.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox12.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox13.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox14.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox15.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox16.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox17.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox18.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox19.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox20.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
        If CheckBox21.Checked = True Then
            TextBox4.Text = Val(TextBox4.Text) + 1
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        TextBox3.Text = DataGridView1.Item(2, i).Value
        TextBox4.Text = DataGridView1.Item(3, i).Value
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        sqlnya = "delete from tb_corona where nama='" & TextBox1.Text & "'"
        Call jalan()
        MsgBox("data berhasil terhapus")
        Call panggildata()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_corona where nama like '%" & TextBox5.Text & "%'", conn)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_corona")
        DataGridView1.DataSource = DS.Tables("tb_corona")
        DataGridView1.Enabled = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        sqlnya = "UPDATE tb_corona set umur='" & TextBox2.Text & "',gol_darah='" & TextBox3.Text & "',point='" & TextBox4.Text & "'where nama='" & TextBox1.Text & "'"
        Call jalan()
        MsgBox("data berhasil terubah")
        Call panggildata()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call panggildata()
    End Sub
End Class
