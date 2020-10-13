Imports System.Data.Odbc
Public Class FormKaryawan
    Sub segarkan()
        Call Koneksi()
        Call TampilGrid()
        Call bersih()
        Call tombolhidup()
        Call isimati()
    End Sub
    Sub TampilGrid()
        DA = New OdbcDataAdapter("select * From karyawan", CONN)
        DS = New DataSet
        DA.Fill(DS, "karyawan")
        DataGridView1.DataSource = DS.Tables("karyawan")
        DataGridView1.ReadOnly = True
    End Sub
    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        DateTimePicker1.Value = Now
        RadioButton1.Checked = False
        RadioButton2.Checked = False
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
        RadioButton1.Enabled = False
        RadioButton2.Enabled = False
        DateTimePicker1.Enabled = False
    End Sub
    Sub isihidup()
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        RadioButton1.Enabled = True
        RadioButton2.Enabled = True
        DateTimePicker1.Enabled = True
    End Sub
    Sub kodeotomatis()
        CMD = New OdbcCommand("select * from karyawan order by kd_karyawan desc", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        If Not RD.HasRows Then
            TextBox1.Text = "KYW" + "001"
        Else
            TextBox1.Text = Val(Microsoft.VisualBasic.Mid(RD.Item("kd_Karyawan").ToString, 4, 3)) + 1
            If Len(TextBox1.Text) = 1 Then
                TextBox1.Text = "KYW00" & TextBox1.Text & ""
            ElseIf Len(TextBox1.Text) = 2 Then
                TextBox1.Text = "KYW0" & TextBox1.Text & ""
            ElseIf Len(TextBox1.Text) = 3 Then
                TextBox1.Text = "KYW" & TextBox1.Text & ""
            End If
        End If
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If Button1.Text = "Tambah" Then
            Button1.Text = "Simpan"
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = True
            Button5.Enabled = False
            Call bersih()
            Call kodeotomatis()
            Call isihidup()
        ElseIf Button1.Text = "Simpan" Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
                MsgBox("Data Yang Anda Masukkan Belum Lengkap!", vbCritical + vbOKOnly, "Peringatan")
            Else
                Dim jk As String = ""
                If RadioButton1.Checked = True Then
                    jk = "Laki-Laki"
                ElseIf RadioButton2.Checked = True Then
                    jk = "Perempuan"
                End If
                Dim simpan As String = "insert into karyawan values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & jk & "','" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "','" & TextBox3.Text & "')"
                CMD = New OdbcCommand(simpan, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Data Berhasil Disimpan!", vbInformation + vbOKOnly, "Informasi")
                Call segarkan()
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Call segarkan()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("Pilih Data Yang Akan Diubah!", vbCritical + vbOKOnly, "Peringatan")
        ElseIf Button2.Text = "Ubah" Then
            Button2.Text = "Simpan"
            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = True
            Button5.Enabled = False
            Call isihidup()
            TextBox1.Enabled = False
        ElseIf Button2.Text = "Simpan" Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
                MsgBox("Data Yang Anda Masukkan Belum Lengkap!", vbCritical + vbOKOnly, "Peringatan")
            Else
                Dim jk As String = ""
                If RadioButton1.Checked = True Then
                    jk = "Laki-Laki"
                ElseIf RadioButton2.Checked = True Then
                    jk = "Perempuan"
                End If
                Dim edit As String = "update karyawan set nama='" & TextBox2.Text & "',jenis_kelamin='" & jk & "',tanggal_lahir='" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "',alamat='" & TextBox3.Text & "' where kd_karyawan='" & TextBox1.Text & "'"
                CMD = New OdbcCommand(edit, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Data Berhasil Diubah", vbInformation + vbOKOnly, "Informasi")
                Call segarkan()
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellMouseDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick

        For f As Integer = 0 To DataGridView1.RowCount - 1
            If IsDBNull(DataGridView1.CurrentRow.Cells(0).Value) Then
                TextBox1.Text = ""
            Else
                TextBox1.Text = DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(1).Value) Then
                TextBox2.Text = ""
            Else
                TextBox2.Text = DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(2).Value) Then
                RadioButton1.Checked = False
                RadioButton2.Checked = False
            Else
                If DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value = "Laki-Laki" Then
                    RadioButton1.Checked = True
                    RadioButton2.Checked = False
                ElseIf DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value = "Perempuan" Then
                    RadioButton1.Checked = False
                    RadioButton2.Checked = True
                End If
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(3).Value) Then
                DateTimePicker1.Value = Now
            Else
                DateTimePicker1.Value = DataGridView1.Item(3, DataGridView1.CurrentRow.Index).Value
            End If

            If IsDBNull(DataGridView1.CurrentRow.Cells(4).Value) Then
                TextBox3.Text = ""
            Else
                TextBox3.Text = DataGridView1.Item(4, DataGridView1.CurrentRow.Index).Value
            End If
        Next
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Then
            MsgBox("Pilih Data Yang Akan Dihapus!", vbCritical + vbOKOnly, "Peringatan")
        Else
            If MsgBox("Anda Yakin Ingin Menghapus?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
                Dim hapus As String = "delete From karyawan where kd_karyawan='" & TextBox1.Text & "'"
                Dim hapus1 As String = "delete From tb_produksi where kd_karyawan='" & TextBox1.Text & "'"
                CMD = New OdbcCommand(hapus1, CONN)
                CMD.ExecuteNonQuery()
                CMD = New OdbcCommand(hapus, CONN)
                CMD.ExecuteNonQuery()
                MsgBox("Data Berhasil Dihapus", vbInformation + vbOKOnly, "Informasi")
                Call segarkan()
            End If
        End If
    End Sub

    Private Sub FormKaryawan_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = FormUtama
        keTengah(FormUtama, Me)
        Call segarkan()
    End Sub
    Sub keluar()
        Close()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If (e.KeyChar = Chr(13)) Then
            TextBox2.Focus()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If (e.KeyChar = Chr(13)) Then
            TextBox3.Focus()
        End If
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        laporankaryawan.Show()
    End Sub
End Class