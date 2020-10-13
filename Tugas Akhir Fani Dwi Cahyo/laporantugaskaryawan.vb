Public Class laporantugaskaryawan

    Private Sub laporantugaskaryawan_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = FormUtama
        keTengah(FormUtama, Me)
        FormTugas.Close()
    End Sub
End Class