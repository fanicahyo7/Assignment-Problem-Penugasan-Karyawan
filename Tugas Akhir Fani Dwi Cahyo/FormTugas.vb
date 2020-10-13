Imports System.Data.Odbc
Public Class FormTugas
    Sub isicombo()
        CMD = New OdbcCommand("select kd_karyawan FROM karyawan", CONN)
        RD = CMD.ExecuteReader
        Do While RD.Read
            ComboBox1.Items.Add(RD.Item(0))
        Loop
    End Sub
    Sub isicombo2()
        CMD = New OdbcCommand("select kd_karyawan FROM karyawan", CONN)
        RD = CMD.ExecuteReader
        Do While RD.Read
            Dim stts As String = "NO"
            For a = 0 To DataGridView1.RowCount - 1
                If RD.Item(0) = DataGridView1.Item(0, a).Value Then
                    stts = "YES"
                End If
            Next
            If stts = "NO" Then
                ComboBox1.Items.Add(RD.Item(0))
            End If
        Loop
    End Sub
    Sub segarkan()
        Call Koneksi()
        Call TampilGrid()
        Call bersih()
        Call isimati()
        Call tombolhidup()
        ComboBox1.Items.Clear()
        Call isicombo()
    End Sub
    Sub TampilGrid()
        DA = New OdbcDataAdapter("select * From tb_produksi", CONN)
        DS = New DataSet
        DA.Fill(DS, "tb_produksi")
        DataGridView1.DataSource = DS.Tables("tb_produksi")
        DataGridView1.ReadOnly = True
    End Sub
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
    End Sub
    Sub tombolhidup()
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        Button4.Enabled = False
        Button5.Enabled = True
        Button1.Text = "Tambah"
        Button2.Text = "Ubah"
        Button3.Text = "Hapus"
    End Sub
    Sub isimati()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = False
        TextBox6.Enabled = False
        TextBox7.Enabled = False
        TextBox8.Enabled = False
        ComboBox1.Enabled = False
    End Sub
    Sub isihidup()
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = True
        TextBox5.Enabled = True
        TextBox6.Enabled = True
        TextBox7.Enabled = True
        TextBox8.Enabled = True
        ComboBox1.Enabled = True
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If Button1.Text = "Tambah" Then
            Button1.Text = "Simpan"
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = True
            Button5.Enabled = False
            Call bersih()
            Call isihidup()
            ComboBox1.Items.Clear()
            Call isicombo2()
        ElseIf Button1.Text = "Simpan" Then
        If ComboBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Then
            MsgBox("Data Yang Anda Masukkan Belum Lengkap!", vbCritical + vbOKOnly, "Peringatan")
        Else
            Dim simpan As String = "insert into tb_produksi values ('" & ComboBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & TextBox7.Text & "','" & TextBox8.Text & "')"
            CMD = New OdbcCommand(simpan, CONN)
            CMD.ExecuteNonQuery()
            MsgBox("Data Berhasil Disimpan!", vbInformation + vbOKOnly, "Informasi")
            ComboBox1.Items.Clear()
            Call segarkan()
        End If
        End If
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        ComboBox1.Items.Clear()
        Call segarkan()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        CMD = New OdbcCommand("select nama FROM karyawan where kd_karyawan= '" & ComboBox1.Text & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        TextBox1.Text = RD.Item(0)
    End Sub

    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick
        ComboBox1.Text = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
        TextBox2.Text = DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value
        TextBox3.Text = DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value
        TextBox4.Text = DataGridView1.Item(3, DataGridView1.CurrentRow.Index).Value
        TextBox5.Text = DataGridView1.Item(4, DataGridView1.CurrentRow.Index).Value
        TextBox6.Text = DataGridView1.Item(5, DataGridView1.CurrentRow.Index).Value
        TextBox7.Text = DataGridView1.Item(6, DataGridView1.CurrentRow.Index).Value
        TextBox8.Text = DataGridView1.Item(7, DataGridView1.CurrentRow.Index).Value
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If ComboBox1.Text = "" Then
            MsgBox("Pilih Data Yang Akan Diubah!", vbCritical + vbOKOnly, "Peringatan")
        ElseIf Button2.Text = "Ubah" Then
            Button2.Text = "Simpan"
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = True
            Button5.Enabled = False
            Call isihidup()
            ComboBox1.Enabled = False
        ElseIf Button2.Text = "Simpan" Then
            If ComboBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Or TextBox8.Text = "" Then
                MsgBox("Data Yang Anda Masukkan Belum Lengkap!", vbCritical + vbOKOnly, "Peringatan")
            Else
                Dim edit As String = "update tb_produksi set Tugas_A='" & TextBox2.Text & "',Tugas_B='" & TextBox3.Text & "',Tugas_C='" & TextBox4.Text & "',Tugas_D='" & TextBox5.Text & "',Tugas_E='" & TextBox6.Text & "',Tugas_F='" & TextBox7.Text & "',Tugas_G='" & TextBox8.Text & "' where kd_karyawan='" & ComboBox1.Text & "'"
                CMD = New OdbcCommand(edit, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Data Berhasil Diubah", vbInformation + vbOKOnly, "Informasi")
                ComboBox1.Items.Clear()
                Call segarkan()
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        If ComboBox1.Text = "" Then
            MsgBox("Pilih Data Yang Akan Dihapus!", vbCritical + vbOKOnly, "Peringatan")
        Else
            If MsgBox("Anda Yakin Ingin Menghapus?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
                Dim hapus As String = "delete From tb_produksi where kd_karyawan='" & ComboBox1.Text & "'"
                CMD = New OdbcCommand(hapus, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Data Berhasil Dihapus", vbInformation + vbOKOnly, "Informasi")
                ComboBox1.Items.Clear()
                Call segarkan()
            End If
        End If
    End Sub

    Private Sub FormTugas_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = FormUtama
        keTengah(FormUtama, Me)
        ComboBox1.Items.Clear()
        Call segarkan()
    End Sub
    Sub keluar()
        Close()
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If (e.KeyChar = Chr(13)) Then
            TextBox3.Focus()
        End If
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If (e.KeyChar = Chr(13)) Then
            TextBox4.Focus()
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox4.KeyPress
        If (e.KeyChar = Chr(13)) Then
            TextBox5.Focus()
        End If
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        If (e.KeyChar = Chr(13)) Then
            TextBox6.Focus()
        End If
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox6.KeyPress
        If (e.KeyChar = Chr(13)) Then
            TextBox7.Focus()
        End If
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox7.KeyPress
        If (e.KeyChar = Chr(13)) Then
            TextBox8.Focus()
        End If
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        laporantugaskaryawan.Show()
    End Sub
End Class