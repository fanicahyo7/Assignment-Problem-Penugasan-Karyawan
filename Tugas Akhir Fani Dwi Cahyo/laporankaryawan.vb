Public Class laporankaryawan

    Private Sub laporankaryawan_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.MdiParent = FormUtama
        keTengah(FormUtama, Me)
        FormKaryawan.Close()
    End Sub
End Class