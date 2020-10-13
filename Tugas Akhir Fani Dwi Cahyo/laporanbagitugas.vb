Public Class laporanbagitugas

    Private Sub laporanbagitugas_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.MdiParent = FormUtama
        keTengah(FormUtama, Me)
        FormPenugasanKaryawan.Close()
    End Sub
End Class