Public Class FormUtama

    Private Sub FormUtama_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        End
    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.Left = (Screen.PrimaryScreen.WorkingArea.Width - Me.Width) / 2
        Me.Top = (Screen.PrimaryScreen.WorkingArea.Height - Me.Height) / 2
        Me.IsMdiContainer = True
    End Sub

    Private Sub KaryawanToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles KaryawanToolStripMenuItem.Click
        FormKaryawan.MdiParent = Me
        FormKaryawan.Show()

        FormTugas.keluar()
        FormPenugasanKaryawan.keluar()
        FormGantiPassword.keluar()
        laporanbagitugas.Close()
        laporankaryawan.Close()
        laporantugaskaryawan.Close()
    End Sub

    Private Sub PengerjaanTugasToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PengerjaanTugasToolStripMenuItem.Click
        FormTugas.MdiParent = Me
        FormTugas.Show()

        FormKaryawan.keluar()
        FormPenugasanKaryawan.keluar()
        FormGantiPassword.keluar()
        laporanbagitugas.Close()
        laporankaryawan.Close()
        laporantugaskaryawan.Close()
    End Sub

    Private Sub RekomendasiPenugasanKaryawanToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RekomendasiPenugasanKaryawanToolStripMenuItem.Click
        FormPenugasanKaryawan.MdiParent = Me
        FormPenugasanKaryawan.Show()

        FormKaryawan.keluar()
        FormTugas.keluar()
        FormGantiPassword.keluar()
        laporanbagitugas.Close()
        laporankaryawan.Close()
        laporantugaskaryawan.Close()
    End Sub

    Private Sub GantiPasswordToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles GantiPasswordToolStripMenuItem.Click
        FormGantiPassword.MdiParent = Me
        FormGantiPassword.Show()

        FormKaryawan.keluar()
        FormTugas.keluar()
        FormPenugasanKaryawan.keluar()
        laporanbagitugas.Close()
        laporankaryawan.Close()
        laporantugaskaryawan.Close()
    End Sub
End Class
