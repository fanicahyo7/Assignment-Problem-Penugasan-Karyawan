Imports System.Data.Odbc
Public Class FormPenugasanKaryawan
    Dim karyawan1, karyawan2, karyawan3, karyawan4, karyawan5, karyawan6, karyawan7 As String
    Dim baris As Integer = 0
    Private Sub FormPenugasanKaryawan_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.MdiParent = FormUtama
        keTengah(FormUtama, Me)
        Call Koneksi()
        Call TampilGrid()
        Call jumlah7()
        Call isicombo()
        Button5.Enabled = False
        Dim truncate As String = "TRUNCATE tb_tugasbagi"
        CMD = New OdbcCommand(truncate, CONN)
        CMD.ExecuteNonQuery()
    End Sub
    Sub keluar()
        Close()
    End Sub
    Sub TampilGrid()
        DA = New OdbcDataAdapter("select * From tb_produksi", CONN)
        DS = New DataSet
        DA.Fill(DS, "tb_produksi")
        DataGridView1.DataSource = DS.Tables("tb_produksi")
        DataGridView1.ReadOnly = True
    End Sub
    Sub jumlah7()
        If Not DataGridView1.RowCount >= 8 Then
            MsgBox("Jumlah Data Belum Berjumlah 7 Karyawan", vbCritical + vbOKOnly, "Peringatan")
            Close()
        End If
    End Sub
    Sub tampilsalin()
        'menyalin datagrid1 ke datagrid3
        For a = 0 To 6
            DataGridView3.Rows.Add(DataGridView6.Item(0, a).Value)
            For b = 0 To 7
                DataGridView3.Item(b, a).Value = DataGridView6.Item(b, a).Value
            Next
        Next
    End Sub
    Sub tampilsalin2()
        For a = 0 To 6
            DataGridView4.Rows.Add(DataGridView3.Item(0, a).Value)
            For b = 0 To 7
                DataGridView4.Item(b, a).Value = DataGridView3.Item(b, a).Value
            Next
        Next
    End Sub
    Sub tampilsalin3()
        For a = 0 To 6
            DataGridView5.Rows.Add(DataGridView4.Item(0, a).Value)
            For b = 0 To 7
                DataGridView5.Item(b, a).Value = DataGridView4.Item(b, a).Value
            Next
        Next
    End Sub
    Sub baristerkecil()
        For a = 0 To 6
            'mencari nilai terkecil pada baris
            Dim nilai As Integer = CInt(DataGridView6.Item(1, a).Value)
            For b = 1 To 7
                If CInt(DataGridView6.Item(b, a).Value) < nilai Then
                    nilai = DataGridView6.Item(b, a).Value
                End If
            Next
            'mengurangi nilai pada baris dengan nilai terkecil pada baris
            For b = 1 To 7
                DataGridView3.Item(b, a).Value = DataGridView6.Item(b, a).Value - nilai
            Next
        Next
    End Sub
    Sub kolomterkecil()
        For a = 1 To 7
            Dim deteksinol As String = ""
            Dim kolom As Integer
            Dim nilai As Integer
            'mendeteksi kolom sudah terdapat nilai 0 atau tidak
            For b = 0 To 6
                If CInt(DataGridView3.Item(a, b).Value) = 0 Then
                    deteksinol = "Ada"
                    b = 7
                Else
                    deteksinol = "Tidak Ada"
                    kolom = a
                End If
            Next

            If deteksinol = "Tidak Ada" Then
                'mencari nilai terkecil pada kolom
                nilai = CInt(DataGridView3.Item(kolom, 0).Value)
                For c = 0 To kolom
                    If CInt(DataGridView3.Item(kolom, c).Value) < nilai Then
                        nilai = DataGridView3.Item(kolom, c).Value
                    End If
                Next
                'mengurangi nilai kolom dengan nilai terkecil pada kolom
                For c = 0 To 6
                    DataGridView3.Item(kolom, c).Value = DataGridView3.Item(kolom, c).Value - nilai
                Next
            End If
        Next
    End Sub
    Sub tarikgaris1()
        Dim K1, K2, K3, K4, K5, K6, K7 As Integer
        Dim B1, B2, B3, B4, B5, B6, B7 As Integer
        Dim totalgaris As Integer = 0
        Dim totalgarisb As Integer = 0
        Dim totalgarisk As Integer = 0

        'mencari total nol pada kolom
        For a = 1 To 7
            For b = 0 To 6
                If CInt(DataGridView3.Item(a, b).Value) = 0 Then
                    If a = 1 Then
                        K1 += 1
                    ElseIf a = 2 Then
                        K2 += 1
                    ElseIf a = 3 Then
                        K3 += 1
                    ElseIf a = 4 Then
                        K4 += 1
                    ElseIf a = 5 Then
                        K5 += 1
                    ElseIf a = 6 Then
                        K6 += 1
                    ElseIf a = 7 Then
                        K7 += 1
                    End If
                End If
            Next
        Next

        'mencari total 0 pada baris
        For a = 0 To 6
            For b = 1 To 7
                If CInt(DataGridView3.Item(b, a).Value) = 0 Then
                    If a = 0 Then
                        B1 += 1
                    ElseIf a = 1 Then
                        B2 += 1
                    ElseIf a = 2 Then
                        B3 += 1
                    ElseIf a = 3 Then
                        B4 += 1
                    ElseIf a = 4 Then
                        B5 += 1
                    ElseIf a = 5 Then
                        B6 += 1
                    ElseIf a = 6 Then
                        B7 += 1
                    End If
                End If
            Next
        Next

        Dim besark As Integer = 0
        Dim keterangank As String = ""
        Dim besarb As Integer = 0
        Dim keteranganb As String = ""

        Do Until (totalgaris = 6)
            For a = 0 To 6
                If K1 > besark Then
                    besark = K1
                    keterangank = "K1"
                ElseIf K2 > besark Then
                    besark = K2
                    keterangank = "K2"
                ElseIf K3 > besark Then
                    besark = K3
                    keterangank = "K3"
                ElseIf K4 > besark Then
                    besark = K4
                    keterangank = "K4"
                ElseIf K5 > besark Then
                    besark = K5
                    keterangank = "K5"
                ElseIf K6 > besark Then
                    besark = K6
                    keterangank = "K6"
                ElseIf K7 > besark Then
                    besark = K7
                    keterangank = "K7"
                End If


                If B1 > besarb Then
                    besarb = B1
                    keteranganb = "B1"
                ElseIf B2 > besarb Then
                    besarb = B2
                    keteranganb = "B2"
                ElseIf B3 > besarb Then
                    besarb = B3
                    keteranganb = "B3"
                ElseIf B4 > besarb Then
                    besarb = B4
                    keteranganb = "B4"
                ElseIf B5 > besarb Then
                    besarb = B5
                    keteranganb = "B5"
                ElseIf B6 > besarb Then
                    besarb = B6
                    keteranganb = "B6"
                ElseIf B7 > besarb Then
                    besarb = B7
                    keteranganb = "B7"
                End If
            Next

            If besark > besarb Then
                Dim a As Integer = 0
                If keterangank = "K1" Then
                    a = 1
                    K1 = 0
                ElseIf keterangank = "K2" Then
                    a = 2
                    K2 = 0
                ElseIf keterangank = "K3" Then
                    a = 3
                    K3 = 0
                ElseIf keterangank = "K4" Then
                    a = 4
                    K4 = 0
                ElseIf keterangank = "K5" Then
                    a = 5
                    K5 = 0
                ElseIf keterangank = "K6" Then
                    a = 6
                    K6 = 0
                ElseIf keterangank = "K7" Then
                    a = 7
                    K7 = 0
                End If

                For b = 0 To 6
                    If DataGridView3.Rows(b).Cells(a).Style.BackColor = Color.Aqua Then
                        DataGridView3.Rows(b).Cells(a).Style.BackColor = Color.Brown
                    Else
                        DataGridView3.Rows(b).Cells(a).Style.BackColor = Color.Aqua
                    End If
                Next
                totalgaris += 1
                totalgarisk += 1

            ElseIf besarb > besark Then
                Dim a As Integer = 0
                If keteranganb = "B1" Then
                    a = 0
                    B1 = 0
                ElseIf keteranganb = "B2" Then
                    a = 1
                    B2 = 0
                ElseIf keteranganb = "B3" Then
                    a = 2
                    B3 = 0
                ElseIf keteranganb = "B4" Then
                    a = 3
                    B4 = 0
                ElseIf keteranganb = "B5" Then
                    a = 4
                    B5 = 0
                ElseIf keteranganb = "B6" Then
                    a = 5
                    B6 = 0
                ElseIf keteranganb = "B7" Then
                    a = 6
                    B7 = 0
                End If

                For b = 1 To 7
                    If DataGridView3.Rows(a).Cells(b).Style.BackColor = Color.Aqua Then
                        DataGridView3.Rows(a).Cells(b).Style.BackColor = Color.Brown
                    Else
                        DataGridView3.Rows(a).Cells(b).Style.BackColor = Color.Aqua
                    End If
                Next
                totalgaris += 1
                totalgarisb += 1
            ElseIf besarb = besark Then
                If totalgarisb >= totalgarisk Then
                    Dim a As Integer = 0
                    If keteranganb = "B1" Then
                        a = 0
                        B1 = 0
                    ElseIf keteranganb = "B2" Then
                        a = 1
                        B2 = 0
                    ElseIf keteranganb = "B3" Then
                        a = 2
                        B3 = 0
                    ElseIf keteranganb = "B4" Then
                        a = 3
                        B4 = 0
                    ElseIf keteranganb = "B5" Then
                        a = 4
                        B5 = 0
                    ElseIf keteranganb = "B6" Then
                        a = 5
                        B6 = 0
                    ElseIf keteranganb = "B7" Then
                        a = 6
                        B7 = 0
                    End If
                    For b = 1 To 7
                        If DataGridView3.Rows(a).Cells(b).Style.BackColor = Color.Aqua Then
                            DataGridView3.Rows(a).Cells(b).Style.BackColor = Color.Brown
                        Else
                            DataGridView3.Rows(a).Cells(b).Style.BackColor = Color.Aqua
                        End If
                    Next
                    totalgaris += 1
                    totalgarisb += 1
                Else
                    Dim a As Integer = 0
                    If keterangank = "K1" Then
                        a = 1
                        K1 = 0
                    ElseIf keterangank = "K2" Then
                        a = 2
                        K2 = 0
                    ElseIf keterangank = "K3" Then
                        a = 3
                        K3 = 0
                    ElseIf keterangank = "K4" Then
                        a = 4
                        K4 = 0
                    ElseIf keterangank = "K5" Then
                        a = 5
                        K5 = 0
                    ElseIf keterangank = "K6" Then
                        a = 6
                        K6 = 0
                    ElseIf keterangank = "K7" Then
                        a = 7
                        K7 = 0
                    End If

                    For b = 0 To 6
                        If DataGridView3.Rows(b).Cells(a).Style.BackColor = Color.Aqua Then
                            DataGridView3.Rows(b).Cells(a).Style.BackColor = Color.Brown
                        Else
                            DataGridView3.Rows(b).Cells(a).Style.BackColor = Color.Aqua
                        End If
                    Next
                    totalgaris += 1
                    totalgarisk += 1
                End If
            End If
            besark = 0
            keterangank = ""
            besarb = 0
            keteranganb = ""
        Loop
    End Sub
    Sub carikesempatanlagi()
        Dim nilaikecil As String = ""
        'cari nilai kecil yang tidak terkena garis
        For a = 0 To 6
            For b = 1 To 7
                If Not DataGridView3.Rows(a).Cells(b).Style.BackColor = Color.Aqua And Not DataGridView3.Rows(a).Cells(b).Style.BackColor = Color.Brown Then
                    If nilaikecil = "" Then
                        nilaikecil = DataGridView3.Item(b, a).Value
                    ElseIf CInt(DataGridView3.Item(b, a).Value) < CInt(nilaikecil) Then
                        nilaikecil = DataGridView3.Item(b, a).Value
                    End If
                End If
            Next
        Next

        'mengurangkan nilai yang tidak terkena garis dengan nilai terkecil yang tidak terkena garis
        For a = 0 To 6
            For b = 1 To 7
                If Not DataGridView3.Rows(a).Cells(b).Style.BackColor = Color.Aqua And Not DataGridView3.Rows(a).Cells(b).Style.BackColor = Color.Brown Then
                    DataGridView4.Item(b, a).Value = DataGridView3.Item(b, a).Value - CInt(nilaikecil)
                End If
            Next
        Next

        For a = 0 To 6
            For b = 1 To 7
                If DataGridView3.Rows(a).Cells(b).Style.BackColor = Color.Brown Then
                    DataGridView4.Item(b, a).Value = CInt(DataGridView3.Item(b, a).Value) + nilaikecil
                End If
            Next
        Next
        'menambahkan nilai yang berpotongan garis dengan nilai terkecil
    End Sub
    Sub tarikgaris2()
        Dim K1, K2, K3, K4, K5, K6, K7 As Integer
        Dim B1, B2, B3, B4, B5, B6, B7 As Integer
        Dim totalgaris As Integer = 0
        Dim totalgarisb As Integer = 0
        Dim totalgarisk As Integer = 0

        'mencari total nol pada kolom
        For a = 1 To 7
            For b = 0 To 6
                If CInt(DataGridView4.Item(a, b).Value) = 0 Then
                    If a = 1 Then
                        K1 += 1
                    ElseIf a = 2 Then
                        K2 += 1
                    ElseIf a = 3 Then
                        K3 += 1
                    ElseIf a = 4 Then
                        K4 += 1
                    ElseIf a = 5 Then
                        K5 += 1
                    ElseIf a = 6 Then
                        K6 += 1
                    ElseIf a = 7 Then
                        K7 += 1
                    End If
                End If
            Next
        Next

        'mencari total 0 pada baris
        For a = 0 To 6
            For b = 1 To 7
                If CInt(DataGridView4.Item(b, a).Value) = 0 Then
                    If a = 0 Then
                        B1 += 1
                    ElseIf a = 1 Then
                        B2 += 1
                    ElseIf a = 2 Then
                        B3 += 1
                    ElseIf a = 3 Then
                        B4 += 1
                    ElseIf a = 4 Then
                        B5 += 1
                    ElseIf a = 5 Then
                        B6 += 1
                    ElseIf a = 6 Then
                        B7 += 1
                    End If
                End If
            Next
        Next

        Dim besark As Integer = 0
        Dim keterangank As String = ""
        Dim besarb As Integer = 0
        Dim keteranganb As String = ""

        Do Until (totalgaris = 7)
            For a = 0 To 6
                If K1 > besark Then
                    besark = K1
                    keterangank = "K1"
                ElseIf K2 > besark Then
                    besark = K2
                    keterangank = "K2"
                ElseIf K3 > besark Then
                    besark = K3
                    keterangank = "K3"
                ElseIf K4 > besark Then
                    besark = K4
                    keterangank = "K4"
                ElseIf K5 > besark Then
                    besark = K5
                    keterangank = "K5"
                ElseIf K6 > besark Then
                    besark = K6
                    keterangank = "K6"
                ElseIf K7 > besark Then
                    besark = K7
                    keterangank = "K7"
                End If


                If B1 > besarb Then
                    besarb = B1
                    keteranganb = "B1"
                ElseIf B2 > besarb Then
                    besarb = B2
                    keteranganb = "B2"
                ElseIf B3 > besarb Then
                    besarb = B3
                    keteranganb = "B3"
                ElseIf B4 > besarb Then
                    besarb = B4
                    keteranganb = "B4"
                ElseIf B5 > besarb Then
                    besarb = B5
                    keteranganb = "B5"
                ElseIf B6 > besarb Then
                    besarb = B6
                    keteranganb = "B6"
                ElseIf B7 > besarb Then
                    besarb = B7
                    keteranganb = "B7"
                End If
            Next

            If besark > besarb Then
                Dim a As Integer = 0
                If keterangank = "K1" Then
                    a = 1
                    K1 = 0
                ElseIf keterangank = "K2" Then
                    a = 2
                    K2 = 0
                ElseIf keterangank = "K3" Then
                    a = 3
                    K3 = 0
                ElseIf keterangank = "K4" Then
                    a = 4
                    K4 = 0
                ElseIf keterangank = "K5" Then
                    a = 5
                    K5 = 0
                ElseIf keterangank = "K6" Then
                    a = 6
                    K6 = 0
                ElseIf keterangank = "K7" Then
                    a = 7
                    K7 = 0
                End If

                For b = 0 To 6
                    DataGridView4.Rows(b).Cells(a).Style.BackColor = Color.Aqua
                Next
                totalgaris += 1
                totalgarisk += 1

            ElseIf besarb > besark Then
                Dim a As Integer = 0
                If keteranganb = "B1" Then
                    a = 0
                    B1 = 0
                ElseIf keteranganb = "B2" Then
                    a = 1
                    B2 = 0
                ElseIf keteranganb = "B3" Then
                    a = 2
                    B3 = 0
                ElseIf keteranganb = "B4" Then
                    a = 3
                    B4 = 0
                ElseIf keteranganb = "B5" Then
                    a = 4
                    B5 = 0
                ElseIf keteranganb = "B6" Then
                    a = 5
                    B6 = 0
                ElseIf keteranganb = "B7" Then
                    a = 6
                    B7 = 0
                End If

                For b = 1 To 7
                    DataGridView4.Rows(a).Cells(b).Style.BackColor = Color.Aqua
                Next
                totalgaris += 1
                totalgarisb += 1
            ElseIf besarb = besark Then
                If totalgarisb >= totalgarisk Then
                    Dim a As Integer = 0
                    If keteranganb = "B1" Then
                        a = 0
                        B1 = 0
                    ElseIf keteranganb = "B2" Then
                        a = 1
                        B2 = 0
                    ElseIf keteranganb = "B3" Then
                        a = 2
                        B3 = 0
                    ElseIf keteranganb = "B4" Then
                        a = 3
                        B4 = 0
                    ElseIf keteranganb = "B5" Then
                        a = 4
                        B5 = 0
                    ElseIf keteranganb = "B6" Then
                        a = 5
                        B6 = 0
                    ElseIf keteranganb = "B7" Then
                        a = 6
                        B7 = 0
                    End If
                    For b = 1 To 7
                        DataGridView4.Rows(a).Cells(b).Style.BackColor = Color.Aqua
                    Next
                    totalgaris += 1
                    totalgarisb += 1
                Else
                    Dim a As Integer = 0
                    If keterangank = "K1" Then
                        a = 1
                        K1 = 0
                    ElseIf keterangank = "K2" Then
                        a = 2
                        K2 = 0
                    ElseIf keterangank = "K3" Then
                        a = 3
                        K3 = 0
                    ElseIf keterangank = "K4" Then
                        a = 4
                        K4 = 0
                    ElseIf keterangank = "K5" Then
                        a = 5
                        K5 = 0
                    ElseIf keterangank = "K6" Then
                        a = 6
                        K6 = 0
                    ElseIf keterangank = "K7" Then
                        a = 7
                        K7 = 0
                    End If

                    For b = 0 To 6
                        DataGridView4.Rows(b).Cells(a).Style.BackColor = Color.Aqua
                    Next
                    totalgaris += 1
                    totalgarisk += 1
                End If
            End If
            besark = 0
            keterangank = ""
            besarb = 0
            keteranganb = ""
        Loop
    End Sub
    Sub tampilsalinkode()
        For a = 0 To 6
            DataGridView2.Rows.Add(DataGridView6.Item(0, a).Value)
            DataGridView2.Item(0, a).Value = DataGridView6.Item(0, a).Value
            Dim sql As String = "select nama from karyawan where kd_karyawan='" & DataGridView2.Item(0, a).Value & "'"
            CMD = New OdbcCommand(sql, CONN)
            RD = CMD.ExecuteReader()
            RD.Read()
            DataGridView2.Item(1, a).Value = RD.Item(0)
        Next
    End Sub
    Sub tandanol()
        For a = 0 To 6
            For b = 1 To 7
                If DataGridView5.Item(b, a).Value = 0 Then
                    DataGridView5.Rows(a).Cells(b).Style.BackColor = Color.Aqua
                End If
            Next
        Next
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If DataGridView2.RowCount = 8 Then
            MsgBox("Bersihkan Terlebih Dahulu", vbCritical + vbOKOnly, "Peringatan")
        ElseIf DataGridView6.RowCount = 8 Then
            Call tampilsalin()
            Call baristerkecil()
            Call kolomterkecil()
            Call tarikgaris1()
            Call tampilsalin2()
            Call carikesempatanlagi()
            Call tarikgaris2()
            Call tampilsalin3()
            Call tampilsalinkode()
            Call tandanol()
            Call seleksi()
            Button5.Enabled = True
        ElseIf DataGridView6.RowCount < 8 Then
            MsgBox("Jumlah Karyawan Kurang Dari Jumlah Tugas", vbCritical + vbOKOnly, "Peringatan")
        ElseIf DataGridView6.RowCount > 8 Then
            MsgBox("Jumlah Karyawan Lebih Dari Jumlah Tugas", vbCritical + vbOKOnly, "Peringatan")
        End If
    End Sub

    Sub seleksi()
        For a = 0 To 6
            For b = 1 To 7
                'merekap tugas yang cocok diambil setiap karyawan atau yang memiliki nilai 0
                If DataGridView5.Rows(a).Cells(b).Style.BackColor = Color.Aqua Then
                    If a = 0 And b = 1 Then
                        karyawan1 += "A,"
                    ElseIf a = 0 And b = 2 Then
                        karyawan1 += "B,"
                    ElseIf a = 0 And b = 3 Then
                        karyawan1 += "C,"
                    ElseIf a = 0 And b = 4 Then
                        karyawan1 += "D,"
                    ElseIf a = 0 And b = 5 Then
                        karyawan1 += "E,"
                    ElseIf a = 0 And b = 6 Then
                        karyawan1 += "F,"
                    ElseIf a = 0 And b = 7 Then
                        karyawan1 += "G,"
                    ElseIf a = 1 And b = 1 Then
                        karyawan2 += "A,"
                    ElseIf a = 1 And b = 2 Then
                        karyawan2 += "B,"
                    ElseIf a = 1 And b = 3 Then
                        karyawan2 += "C,"
                    ElseIf a = 1 And b = 4 Then
                        karyawan2 += "D,"
                    ElseIf a = 1 And b = 5 Then
                        karyawan2 += "E,"
                    ElseIf a = 1 And b = 6 Then
                        karyawan2 += "F,"
                    ElseIf a = 1 And b = 7 Then
                        karyawan2 += "G,"
                    ElseIf a = 2 And b = 1 Then
                        karyawan3 += "A,"
                    ElseIf a = 2 And b = 2 Then
                        karyawan3 += "B,"
                    ElseIf a = 2 And b = 3 Then
                        karyawan3 += "C,"
                    ElseIf a = 2 And b = 4 Then
                        karyawan3 += "D,"
                    ElseIf a = 2 And b = 5 Then
                        karyawan3 += "E,"
                    ElseIf a = 2 And b = 6 Then
                        karyawan3 += "F,"
                    ElseIf a = 2 And b = 7 Then
                        karyawan3 += "G,"
                    ElseIf a = 3 And b = 1 Then
                        karyawan4 += "A,"
                    ElseIf a = 3 And b = 2 Then
                        karyawan4 += "B,"
                    ElseIf a = 3 And b = 3 Then
                        karyawan4 += "C,"
                    ElseIf a = 3 And b = 4 Then
                        karyawan4 += "D,"
                    ElseIf a = 3 And b = 5 Then
                        karyawan4 += "E,"
                    ElseIf a = 3 And b = 6 Then
                        karyawan4 += "F,"
                    ElseIf a = 3 And b = 7 Then
                        karyawan4 += "G,"
                    ElseIf a = 4 And b = 1 Then
                        karyawan5 += "A,"
                    ElseIf a = 4 And b = 2 Then
                        karyawan5 += "B,"
                    ElseIf a = 4 And b = 3 Then
                        karyawan5 += "C,"
                    ElseIf a = 4 And b = 4 Then
                        karyawan5 += "D,"
                    ElseIf a = 4 And b = 5 Then
                        karyawan5 += "E,"
                    ElseIf a = 4 And b = 6 Then
                        karyawan5 += "F,"
                    ElseIf a = 4 And b = 7 Then
                        karyawan5 += "G,"
                    ElseIf a = 5 And b = 1 Then
                        karyawan6 += "A,"
                    ElseIf a = 5 And b = 2 Then
                        karyawan6 += "B,"
                    ElseIf a = 5 And b = 3 Then
                        karyawan6 += "C,"
                    ElseIf a = 5 And b = 4 Then
                        karyawan6 += "D,"
                    ElseIf a = 5 And b = 5 Then
                        karyawan6 += "E,"
                    ElseIf a = 5 And b = 6 Then
                        karyawan6 += "F,"
                    ElseIf a = 5 And b = 7 Then
                        karyawan6 += "G,"
                    ElseIf a = 6 And b = 1 Then
                        karyawan7 += "A,"
                    ElseIf a = 6 And b = 2 Then
                        karyawan7 += "B,"
                    ElseIf a = 6 And b = 3 Then
                        karyawan7 += "C,"
                    ElseIf a = 6 And b = 4 Then
                        karyawan7 += "D,"
                    ElseIf a = 6 And b = 5 Then
                        karyawan7 += "E,"
                    ElseIf a = 6 And b = 6 Then
                        karyawan7 += "F,"
                    ElseIf a = 6 And b = 7 Then
                        karyawan7 += "G,"
                    End If
                End If
            Next
        Next

        'menampilkan rekapan setiap karyawan di datagrid2
        For a = 0 To 6
            If a = 0 Then
                DataGridView2.Item(2, a).Value = karyawan1
            ElseIf a = 1 Then
                DataGridView2.Item(2, a).Value = karyawan2
            ElseIf a = 2 Then
                DataGridView2.Item(2, a).Value = karyawan3
            ElseIf a = 3 Then
                DataGridView2.Item(2, a).Value = karyawan4
            ElseIf a = 4 Then
                DataGridView2.Item(2, a).Value = karyawan5
            ElseIf a = 5 Then
                DataGridView2.Item(2, a).Value = karyawan6
            ElseIf a = 6 Then
                DataGridView2.Item(2, a).Value = karyawan7
            End If
        Next

        Dim jmlsatu, jmldua, jmltiga, jmlempat, jmllima, jmlenam, jmltujuh As Integer

        Dim kecilstr As String = ""

        jmlsatu = Strings.Split(karyawan1, ",").Length
        jmldua = Strings.Split(karyawan2, ",").Length
        jmltiga = Strings.Split(karyawan3, ",").Length
        jmlempat = Strings.Split(karyawan4, ",").Length
        jmllima = Strings.Split(karyawan5, ",").Length
        jmlenam = Strings.Split(karyawan6, ",").Length
        jmltujuh = Strings.Split(karyawan7, ",").Length

        Dim tugasa, tugasb, tugasc, tugasd, tugase, tugasf, tugasg As Integer

        Dim sttugasa As String = "BELUM"
        Dim sttugasb As String = "BELUM"
        Dim sttugasc As String = "BELUM"
        Dim sttugasd As String = "BELUM"
        Dim sttugase As String = "BELUM"
        Dim sttugasf As String = "BELUM"
        Dim sttugasg As String = "BELUM"

        Do Until (jmlsatu = 1 And jmldua = 1 And jmltiga = 1 And jmlempat = 1 And jmllima = 1 And jmlenam = 1 And jmltujuh = 1)
            jmlsatu = Strings.Split(karyawan1, ",").Length
            jmldua = Strings.Split(karyawan2, ",").Length
            jmltiga = Strings.Split(karyawan3, ",").Length
            jmlempat = Strings.Split(karyawan4, ",").Length
            jmllima = Strings.Split(karyawan5, ",").Length
            jmlenam = Strings.Split(karyawan6, ",").Length
            jmltujuh = Strings.Split(karyawan7, ",").Length

            'seleksi jika karyawan hanya memiliki satu tugas
            Dim slsi As String
            For a = 0 To 6
                If jmlsatu = 2 Then
                    slsi = Strings.Split(karyawan1, ",")(0)
                    If slsi = "A" Then
                        karyawan1 = "A"
                        sttugasa = "SUDAH"
                        Call seleksikerjaan("A")
                    ElseIf slsi = "B" Then
                        karyawan1 = "B"
                        sttugasb = "SUDAH"
                        Call seleksikerjaan("B")
                    ElseIf slsi = "C" Then
                        karyawan1 = "C"
                        sttugasc = "SUDAH"
                        Call seleksikerjaan("C")
                    ElseIf slsi = "D" Then
                        karyawan1 = "D"
                        sttugasd = "SUDAH"
                        Call seleksikerjaan("D")
                    ElseIf slsi = "E" Then
                        karyawan1 = "E"
                        sttugase = "SUDAH"
                        Call seleksikerjaan("E")
                    ElseIf slsi = "F" Then
                        karyawan1 = "F"
                        sttugasf = "SUDAH"
                        Call seleksikerjaan("F")
                    ElseIf slsi = "G" Then
                        karyawan1 = "G"
                        sttugasg = "SUDAH"
                        Call seleksikerjaan("G")
                    End If
                ElseIf jmldua = 2 Then
                    slsi = Strings.Split(karyawan2, ",")(0)
                    If slsi = "A" Then
                        karyawan2 = "A"
                        sttugasa = "SUDAH"
                        Call seleksikerjaan("A")
                    ElseIf slsi = "B" Then
                        karyawan2 = "B"
                        sttugasb = "SUDAH"
                        Call seleksikerjaan("B")
                    ElseIf slsi = "C" Then
                        karyawan2 = "C"
                        sttugasc = "SUDAH"
                        Call seleksikerjaan("C")
                    ElseIf slsi = "D" Then
                        karyawan2 = "D"
                        sttugasd = "SUDAH"
                        Call seleksikerjaan("D")
                    ElseIf slsi = "E" Then
                        karyawan2 = "E"
                        sttugase = "SUDAH"
                        Call seleksikerjaan("E")
                    ElseIf slsi = "F" Then
                        karyawan2 = "F"
                        sttugasf = "SUDAH"
                        Call seleksikerjaan("F")
                    ElseIf slsi = "G" Then
                        karyawan2 = "G"
                        sttugasg = "SUDAH"
                        Call seleksikerjaan("G")
                    End If
                ElseIf jmltiga = 2 Then
                    slsi = Strings.Split(karyawan3, ",")(0)
                    If slsi = "A" Then
                        karyawan3 = "A"
                        sttugasa = "SUDAH"
                        Call seleksikerjaan("A")
                    ElseIf slsi = "B" Then
                        karyawan3 = "B"
                        sttugasb = "SUDAH"
                        Call seleksikerjaan("B")
                    ElseIf slsi = "C" Then
                        karyawan3 = "C"
                        sttugasc = "SUDAH"
                        Call seleksikerjaan("C")
                    ElseIf slsi = "D" Then
                        karyawan3 = "D"
                        sttugasd = "SUDAH"
                        Call seleksikerjaan("D")
                    ElseIf slsi = "E" Then
                        karyawan3 = "E"
                        sttugase = "SUDAH"
                        Call seleksikerjaan("E")
                    ElseIf slsi = "F" Then
                        karyawan3 = "F"
                        sttugasf = "SUDAH"
                        Call seleksikerjaan("F")
                    ElseIf slsi = "G" Then
                        karyawan3 = "G"
                        sttugasg = "SUDAH"
                        Call seleksikerjaan("G")
                    End If
                ElseIf jmlempat = 2 Then
                    slsi = Strings.Split(karyawan4, ",")(0)
                    If slsi = "A" Then
                        karyawan4 = "A"
                        sttugasa = "SUDAH"
                        Call seleksikerjaan("A")
                    ElseIf slsi = "B" Then
                        karyawan4 = "B"
                        sttugasb = "SUDAH"
                        Call seleksikerjaan("B")
                    ElseIf slsi = "C" Then
                        karyawan4 = "C"
                        sttugasc = "SUDAH"
                        Call seleksikerjaan("C")
                    ElseIf slsi = "D" Then
                        karyawan4 = "D"
                        sttugasd = "SUDAH"
                        Call seleksikerjaan("D")
                    ElseIf slsi = "E" Then
                        karyawan4 = "E"
                        sttugase = "SUDAH"
                        Call seleksikerjaan("E")
                    ElseIf slsi = "F" Then
                        karyawan4 = "F"
                        sttugasf = "SUDAH"
                        Call seleksikerjaan("F")
                    ElseIf slsi = "G" Then
                        karyawan4 = "G"
                        sttugasg = "SUDAH"
                        Call seleksikerjaan("G")
                    End If
                ElseIf jmllima = 2 Then
                    slsi = Strings.Split(karyawan5, ",")(0)
                    If slsi = "A" Then
                        karyawan5 = "A"
                        sttugasa = "SUDAH"
                        Call seleksikerjaan("A")
                    ElseIf slsi = "B" Then
                        karyawan5 = "B"
                        sttugasb = "SUDAH"
                        Call seleksikerjaan("B")
                    ElseIf slsi = "C" Then
                        karyawan5 = "C"
                        sttugasc = "SUDAH"
                        Call seleksikerjaan("C")
                    ElseIf slsi = "D" Then
                        karyawan5 = "D"
                        sttugasd = "SUDAH"
                        Call seleksikerjaan("D")
                    ElseIf slsi = "E" Then
                        karyawan5 = "E"
                        sttugase = "SUDAH"
                        Call seleksikerjaan("E")
                    ElseIf slsi = "F" Then
                        karyawan5 = "F"
                        sttugasf = "SUDAH"
                        Call seleksikerjaan("F")
                    ElseIf slsi = "G" Then
                        karyawan5 = "G"
                        sttugasg = "SUDAH"
                        Call seleksikerjaan("G")
                    End If
                ElseIf jmlenam = 2 Then
                    slsi = Strings.Split(karyawan6, ",")(0)
                    If slsi = "A" Then
                        karyawan6 = "A"
                        sttugasa = "SUDAH"
                        Call seleksikerjaan("A")
                    ElseIf slsi = "B" Then
                        karyawan6 = "B"
                        sttugasb = "SUDAH"
                        Call seleksikerjaan("B")
                    ElseIf slsi = "C" Then
                        karyawan6 = "C"
                        sttugasc = "SUDAH"
                        Call seleksikerjaan("C")
                    ElseIf slsi = "D" Then
                        karyawan6 = "D"
                        sttugasd = "SUDAH"
                        Call seleksikerjaan("D")
                    ElseIf slsi = "E" Then
                        karyawan6 = "E"
                        sttugase = "SUDAH"
                        Call seleksikerjaan("E")
                    ElseIf slsi = "F" Then
                        karyawan6 = "F"
                        sttugasf = "SUDAH"
                        Call seleksikerjaan("F")
                    ElseIf slsi = "G" Then
                        karyawan6 = "G"
                        sttugasg = "SUDAH"
                        Call seleksikerjaan("G")
                    End If
                ElseIf jmltujuh = 2 Then
                    slsi = Strings.Split(karyawan7, ",")(0)
                    If slsi = "A" Then
                        karyawan7 = "A"
                        sttugasa = "SUDAH"
                        Call seleksikerjaan("A")
                    ElseIf slsi = "B" Then
                        karyawan7 = "B"
                        sttugasb = "SUDAH"
                        Call seleksikerjaan("B")
                    ElseIf slsi = "C" Then
                        karyawan7 = "C"
                        sttugasc = "SUDAH"
                        Call seleksikerjaan("C")
                    ElseIf slsi = "D" Then
                        karyawan7 = "D"
                        sttugasd = "SUDAH"
                        Call seleksikerjaan("D")
                    ElseIf slsi = "E" Then
                        karyawan7 = "E"
                        sttugase = "SUDAH"
                        Call seleksikerjaan("E")
                    ElseIf slsi = "F" Then
                        karyawan7 = "F"
                        sttugasf = "SUDAH"
                        Call seleksikerjaan("F")
                    ElseIf slsi = "G" Then
                        karyawan7 = "G"
                        sttugasg = "SUDAH"
                        Call seleksikerjaan("G")
                    End If
                End If
                jmlsatu = Strings.Split(karyawan1, ",").Length
                jmldua = Strings.Split(karyawan2, ",").Length
                jmltiga = Strings.Split(karyawan3, ",").Length
                jmlempat = Strings.Split(karyawan4, ",").Length
                jmllima = Strings.Split(karyawan5, ",").Length
                jmlenam = Strings.Split(karyawan6, ",").Length
                jmltujuh = Strings.Split(karyawan7, ",").Length
            Next
            jmlsatu = Strings.Split(karyawan1, ",").Length
            jmldua = Strings.Split(karyawan2, ",").Length
            jmltiga = Strings.Split(karyawan3, ",").Length
            jmlempat = Strings.Split(karyawan4, ",").Length
            jmllima = Strings.Split(karyawan5, ",").Length
            jmlenam = Strings.Split(karyawan6, ",").Length
            jmltujuh = Strings.Split(karyawan7, ",").Length

            tugasa = 0
            tugasb = 0
            tugasc = 0
            tugasd = 0
            tugase = 0
            tugasf = 0
            tugasg = 0

            'merekap jumlah tugas
            For a = 0 To jmlsatu - 2
                If Strings.Split(karyawan1, ",")(a) = "A" Then
                    tugasa += 1
                ElseIf Strings.Split(karyawan1, ",")(a) = "B" Then
                    tugasb += 1
                ElseIf Strings.Split(karyawan1, ",")(a) = "C" Then
                    tugasc += 1
                ElseIf Strings.Split(karyawan1, ",")(a) = "D" Then
                    tugasd += 1
                ElseIf Strings.Split(karyawan1, ",")(a) = "E" Then
                    tugase += 1
                ElseIf Strings.Split(karyawan1, ",")(a) = "F" Then
                    tugasf += 1
                ElseIf Strings.Split(karyawan1, ",")(a) = "G" Then
                    tugasg += 1
                End If
            Next

            For a = 0 To jmldua - 2
                If Strings.Split(karyawan2, ",")(a) = "A" Then
                    tugasa += 1
                ElseIf Strings.Split(karyawan2, ",")(a) = "B" Then
                    tugasb += 1
                ElseIf Strings.Split(karyawan2, ",")(a) = "C" Then
                    tugasc += 1
                ElseIf Strings.Split(karyawan2, ",")(a) = "D" Then
                    tugasd += 1
                ElseIf Strings.Split(karyawan2, ",")(a) = "E" Then
                    tugase += 1
                ElseIf Strings.Split(karyawan2, ",")(a) = "F" Then
                    tugasf += 1
                ElseIf Strings.Split(karyawan2, ",")(a) = "G" Then
                    tugasg += 1
                End If
            Next

            For a = 0 To jmltiga - 2
                If Strings.Split(karyawan3, ",")(a) = "A" Then
                    tugasa += 1
                ElseIf Strings.Split(karyawan3, ",")(a) = "B" Then
                    tugasb += 1
                ElseIf Strings.Split(karyawan3, ",")(a) = "C" Then
                    tugasc += 1
                ElseIf Strings.Split(karyawan3, ",")(a) = "D" Then
                    tugasd += 1
                ElseIf Strings.Split(karyawan3, ",")(a) = "E" Then
                    tugase += 1
                ElseIf Strings.Split(karyawan3, ",")(a) = "F" Then
                    tugasf += 1
                ElseIf Strings.Split(karyawan3, ",")(a) = "G" Then
                    tugasg += 1
                End If
            Next

            For a = 0 To jmlempat - 2
                If Strings.Split(karyawan4, ",")(a) = "A" Then
                    tugasa += 1
                ElseIf Strings.Split(karyawan4, ",")(a) = "B" Then
                    tugasb += 1
                ElseIf Strings.Split(karyawan4, ",")(a) = "C" Then
                    tugasc += 1
                ElseIf Strings.Split(karyawan4, ",")(a) = "D" Then
                    tugasd += 1
                ElseIf Strings.Split(karyawan4, ",")(a) = "E" Then
                    tugase += 1
                ElseIf Strings.Split(karyawan4, ",")(a) = "F" Then
                    tugasf += 1
                ElseIf Strings.Split(karyawan4, ",")(a) = "G" Then
                    tugasg += 1
                End If
            Next

            For a = 0 To jmllima - 2
                If Strings.Split(karyawan5, ",")(a) = "A" Then
                    tugasa += 1
                ElseIf Strings.Split(karyawan5, ",")(a) = "B" Then
                    tugasb += 1
                ElseIf Strings.Split(karyawan5, ",")(a) = "C" Then
                    tugasc += 1
                ElseIf Strings.Split(karyawan5, ",")(a) = "D" Then
                    tugasd += 1
                ElseIf Strings.Split(karyawan5, ",")(a) = "E" Then
                    tugase += 1
                ElseIf Strings.Split(karyawan5, ",")(a) = "F" Then
                    tugasf += 1
                ElseIf Strings.Split(karyawan5, ",")(a) = "G" Then
                    tugasg += 1
                End If
            Next

            For a = 0 To jmlenam - 2
                If Strings.Split(karyawan6, ",")(a) = "A" Then
                    tugasa += 1
                ElseIf Strings.Split(karyawan6, ",")(a) = "B" Then
                    tugasb += 1
                ElseIf Strings.Split(karyawan6, ",")(a) = "C" Then
                    tugasc += 1
                ElseIf Strings.Split(karyawan6, ",")(a) = "D" Then
                    tugasd += 1
                ElseIf Strings.Split(karyawan6, ",")(a) = "E" Then
                    tugase += 1
                ElseIf Strings.Split(karyawan6, ",")(a) = "F" Then
                    tugasf += 1
                ElseIf Strings.Split(karyawan6, ",")(a) = "G" Then
                    tugasg += 1
                End If
            Next

            For a = 0 To jmltujuh - 2
                If Strings.Split(karyawan7, ",")(a) = "A" Then
                    tugasa += 1
                ElseIf Strings.Split(karyawan7, ",")(a) = "B" Then
                    tugasb += 1
                ElseIf Strings.Split(karyawan7, ",")(a) = "C" Then
                    tugasc += 1
                ElseIf Strings.Split(karyawan7, ",")(a) = "D" Then
                    tugasd += 1
                ElseIf Strings.Split(karyawan7, ",")(a) = "E" Then
                    tugase += 1
                ElseIf Strings.Split(karyawan7, ",")(a) = "F" Then
                    tugasf += 1
                ElseIf Strings.Split(karyawan7, ",")(a) = "G" Then
                    tugasg += 1
                End If
            Next

            'menyimpan ke array
            Dim anu(8) As Integer
            For a = 0 To 6
                If a = 0 Then
                    anu(a) = tugasa
                ElseIf a = 1 Then
                    anu(a) = tugasb
                ElseIf a = 2 Then
                    anu(a) = tugasc
                ElseIf a = 3 Then
                    anu(a) = tugasd
                ElseIf a = 4 Then
                    anu(a) = tugase
                ElseIf a = 5 Then
                    anu(a) = tugasf
                ElseIf a = 6 Then
                    anu(a) = tugasg
                End If
            Next
            Dim kecil As Integer = 0
            Dim keterangan As String = ""

            'mencari jumlah tugas yang paling sedikit
            If Not anu(0) <= 1 Then
                kecil = anu(0)
                keterangan = "A"
            ElseIf Not anu(1) <= 1 Then
                kecil = anu(1)
                keterangan = "B"
            ElseIf Not anu(2) <= 1 Then
                kecil = anu(2)
                keterangan = "C"
            ElseIf Not anu(3) <= 1 Then
                kecil = anu(3)
                keterangan = "D"
            ElseIf Not anu(4) <= 1 Then
                kecil = anu(4)
                keterangan = "E"
            ElseIf Not anu(5) <= 1 Then
                kecil = anu(5)
                keterangan = "F"
            ElseIf Not anu(6) <= 1 Then
                kecil = anu(6)
                keterangan = "G"
            End If

            For a = 0 To 6
                'jika jumlah tugas A lebih kecil dari nilai terkecil dan tugas A belum memiliki 1 pemilik(karyawan) 
                If anu(0) < kecil And sttugasa = "BELUM" Then
                    'jika jumlah tugas A = 1 maka akan dicari nilai A disetiap karyawan. jika tidak dan nilai A lebih kecil dari kecil akan menjadi nilai terkecil
                    If anu(0) = 1 Then
                        'jika jumlah tugas pada karyawan lebih dari satu maka akan dicari Tugas A pada setiap karyawan.
                        For b = 0 To jmlsatu - 2
                            If Strings.Split(karyawan1, ",")(b) = "A" Then
                                karyawan1 = "A"
                                sttugasa = "SUDAH"
                                b = jmlsatu
                                Call seleksikerjaan("A")
                            End If
                        Next

                        For b = 0 To jmldua - 2
                            If Strings.Split(karyawan2, ",")(b) = "A" Then
                                karyawan2 = "A"
                                sttugasa = "SUDAH"
                                b = jmldua
                                Call seleksikerjaan("A")
                            End If
                        Next

                        For b = 0 To jmltiga - 2
                            If Strings.Split(karyawan3, ",")(b) = "A" Then
                                karyawan3 = "A"
                                sttugasa = "SUDAH"
                                b = jmltiga
                                Call seleksikerjaan("A")
                            End If
                        Next

                        For b = 0 To jmlempat - 2
                            If Strings.Split(karyawan4, ",")(b) = "A" Then
                                karyawan4 = "A"
                                sttugasa = "SUDAH"
                                b = jmlempat
                                Call seleksikerjaan("A")
                            End If
                        Next

                        For b = 0 To jmllima - 2
                            If Strings.Split(karyawan5, ",")(b) = "A" Then
                                karyawan5 = "A"
                                sttugasa = "SUDAH"
                                b = jmllima
                                Call seleksikerjaan("A")
                            End If
                        Next

                        For b = 0 To jmlenam - 2
                            If Strings.Split(karyawan6, ",")(b) = "A" Then
                                karyawan6 = "A"
                                sttugasa = "SUDAH"
                                b = jmlenam
                                Call seleksikerjaan("A")
                            End If
                        Next

                        For b = 0 To jmltujuh - 2
                            If Strings.Split(karyawan7, ",")(b) = "A" Then
                                karyawan7 = "A"
                                sttugasa = "SUDAH"
                                b = jmltujuh
                                Call seleksikerjaan("A")
                            End If
                        Next
                    Else
                        kecil = anu(0)
                        keterangan = "A"
                    End If
                Else
                    kecil = kecil
                    keterangan = keterangan
                End If
                jmlsatu = Strings.Split(karyawan1, ",").Length
                jmldua = Strings.Split(karyawan2, ",").Length
                jmltiga = Strings.Split(karyawan3, ",").Length
                jmlempat = Strings.Split(karyawan4, ",").Length
                jmllima = Strings.Split(karyawan5, ",").Length
                jmlenam = Strings.Split(karyawan6, ",").Length
                jmltujuh = Strings.Split(karyawan7, ",").Length

                If anu(1) < kecil And sttugasb = "BELUM" Then
                    If anu(1) = 1 Then
                        For b = 0 To jmlsatu - 2
                            If Strings.Split(karyawan1, ",")(b) = "B" Then
                                karyawan1 = "B"
                                sttugasb = "SUDAH"
                                b = jmlsatu
                                Call seleksikerjaan("B")
                            End If
                        Next

                        For b = 0 To jmldua - 2
                            If Strings.Split(karyawan2, ",")(b) = "B" Then
                                karyawan2 = "B"
                                sttugasb = "SUDAH"
                                b = jmldua
                                Call seleksikerjaan("B")
                            End If
                        Next

                        For b = 0 To jmltiga - 2
                            If Strings.Split(karyawan3, ",")(b) = "B" Then
                                karyawan3 = "B"
                                sttugasb = "SUDAH"
                                b = jmltiga
                                Call seleksikerjaan("B")
                            End If
                        Next

                        For b = 0 To jmlempat - 2
                            If Strings.Split(karyawan4, ",")(b) = "B" Then
                                karyawan4 = "B"
                                sttugasb = "SUDAH"
                                b = jmlempat
                                Call seleksikerjaan("B")
                            End If
                        Next

                        For b = 0 To jmllima - 2
                            If Strings.Split(karyawan5, ",")(b) = "B" Then
                                karyawan5 = "B"
                                sttugasb = "SUDAH"
                                b = jmllima
                                Call seleksikerjaan("B")
                            End If
                        Next

                        For b = 0 To jmlenam - 2
                            If Strings.Split(karyawan6, ",")(b) = "B" Then
                                karyawan6 = "B"
                                sttugasb = "SUDAH"
                                b = jmlenam
                                Call seleksikerjaan("B")
                            End If
                        Next

                        For b = 0 To jmltujuh - 2
                            If Strings.Split(karyawan7, ",")(b) = "B" Then
                                karyawan7 = "B"
                                sttugasb = "SUDAH"
                                b = jmltujuh
                                Call seleksikerjaan("B")
                            End If
                        Next
                    Else
                        kecil = anu(1)
                        keterangan = "B"
                    End If
                Else
                    kecil = kecil
                    keterangan = keterangan
                End If
                jmlsatu = Strings.Split(karyawan1, ",").Length
                jmldua = Strings.Split(karyawan2, ",").Length
                jmltiga = Strings.Split(karyawan3, ",").Length
                jmlempat = Strings.Split(karyawan4, ",").Length
                jmllima = Strings.Split(karyawan5, ",").Length
                jmlenam = Strings.Split(karyawan6, ",").Length
                jmltujuh = Strings.Split(karyawan7, ",").Length

                If anu(2) < kecil And sttugasc = "BELUM" Then
                    If anu(2) = 1 Then
                        For b = 0 To jmlsatu - 2
                            If Strings.Split(karyawan1, ",")(b) = "C" Then
                                karyawan1 = "C"
                                sttugasc = "SUDAH"
                                b = jmlsatu
                                Call seleksikerjaan("C")
                            End If
                        Next

                        For b = 0 To jmldua - 2
                            If Strings.Split(karyawan2, ",")(b) = "C" Then
                                karyawan2 = "C"
                                sttugasc = "SUDAH"
                                b = jmldua
                                Call seleksikerjaan("C")
                            End If
                        Next

                        For b = 0 To jmltiga - 2
                            If Strings.Split(karyawan3, ",")(b) = "C" Then
                                karyawan3 = "C"
                                sttugasc = "SUDAH"
                                b = jmltiga
                                Call seleksikerjaan("C")
                            End If
                        Next

                        For b = 0 To jmlempat - 2
                            If Strings.Split(karyawan4, ",")(b) = "C" Then
                                karyawan4 = "C"
                                sttugasc = "SUDAH"
                                b = jmlempat
                                Call seleksikerjaan("C")
                            End If
                        Next

                        For b = 0 To jmllima - 2
                            If Strings.Split(karyawan5, ",")(b) = "C" Then
                                karyawan5 = "C"
                                sttugasc = "SUDAH"
                                b = jmllima
                                Call seleksikerjaan("C")
                            End If
                        Next

                        For b = 0 To jmlenam - 2
                            If Strings.Split(karyawan6, ",")(b) = "C" Then
                                karyawan6 = "C"
                                sttugasc = "SUDAH"
                                b = jmlenam
                                Call seleksikerjaan("C")
                            End If
                        Next

                        For b = 0 To jmltujuh - 2
                            If Strings.Split(karyawan7, ",")(b) = "C" Then
                                karyawan7 = "C"
                                sttugasc = "SUDAH"
                                b = jmltujuh
                                Call seleksikerjaan("C")
                            End If
                        Next
                    Else
                        kecil = anu(2)
                        keterangan = "C"
                    End If
                Else
                    kecil = kecil
                    keterangan = keterangan
                End If
                jmlsatu = Strings.Split(karyawan1, ",").Length
                jmldua = Strings.Split(karyawan2, ",").Length
                jmltiga = Strings.Split(karyawan3, ",").Length
                jmlempat = Strings.Split(karyawan4, ",").Length
                jmllima = Strings.Split(karyawan5, ",").Length
                jmlenam = Strings.Split(karyawan6, ",").Length
                jmltujuh = Strings.Split(karyawan7, ",").Length


                If anu(3) < kecil And sttugasd = "BELUM" Then
                    If anu(3) = 1 Then
                        For b = 0 To jmlsatu - 2
                            If Strings.Split(karyawan1, ",")(b) = "D" Then
                                karyawan1 = "D"
                                sttugasd = "SUDAH"
                                b = jmlsatu
                                Call seleksikerjaan("D")
                            End If
                        Next

                        For b = 0 To jmldua - 2
                            If Strings.Split(karyawan2, ",")(b) = "D" Then
                                karyawan2 = "D"
                                sttugasd = "SUDAH"
                                b = jmldua
                                Call seleksikerjaan("D")
                            End If
                        Next

                        For b = 0 To jmltiga - 2
                            If Strings.Split(karyawan3, ",")(b) = "D" Then
                                karyawan3 = "D"
                                sttugasd = "SUDAH"
                                b = jmltiga
                                Call seleksikerjaan("D")
                            End If
                        Next

                        For b = 0 To jmlempat - 2
                            If Strings.Split(karyawan4, ",")(b) = "D" Then
                                karyawan4 = "D"
                                sttugasd = "SUDAH"
                                b = jmlempat
                                Call seleksikerjaan("D")
                            End If
                        Next

                        For b = 0 To jmllima - 2
                            If Strings.Split(karyawan5, ",")(b) = "D" Then
                                karyawan5 = "D"
                                sttugasd = "SUDAH"
                                b = jmllima
                                Call seleksikerjaan("D")
                            End If
                        Next

                        For b = 0 To jmlenam - 2
                            If Strings.Split(karyawan6, ",")(b) = "D" Then
                                karyawan6 = "D"
                                sttugasd = "SUDAH"
                                b = jmlenam
                                Call seleksikerjaan("D")
                            End If
                        Next

                        For b = 0 To jmltujuh - 2
                            If Strings.Split(karyawan7, ",")(b) = "D" Then
                                karyawan7 = "D"
                                sttugasd = "SUDAH"
                                b = jmltujuh
                                Call seleksikerjaan("D")
                            End If
                        Next
                    Else
                        kecil = anu(3)
                        keterangan = "D"
                    End If
                Else
                    kecil = kecil
                    keterangan = keterangan
                End If
                jmlsatu = Strings.Split(karyawan1, ",").Length
                jmldua = Strings.Split(karyawan2, ",").Length
                jmltiga = Strings.Split(karyawan3, ",").Length
                jmlempat = Strings.Split(karyawan4, ",").Length
                jmllima = Strings.Split(karyawan5, ",").Length
                jmlenam = Strings.Split(karyawan6, ",").Length
                jmltujuh = Strings.Split(karyawan7, ",").Length


                If anu(4) < kecil And sttugase = "BELUM" Then
                    If anu(4) = 1 Then
                        For b = 0 To jmlsatu - 2
                            If Strings.Split(karyawan1, ",")(b) = "E" Then
                                karyawan1 = "E"
                                sttugase = "SUDAH"
                                b = jmlsatu
                                Call seleksikerjaan("E")
                            End If
                        Next

                        For b = 0 To jmldua - 2
                            If Strings.Split(karyawan2, ",")(b) = "E" Then
                                karyawan2 = "E"
                                sttugase = "SUDAH"
                                b = jmldua
                                Call seleksikerjaan("E")
                            End If
                        Next

                        For b = 0 To jmltiga - 2
                            If Strings.Split(karyawan3, ",")(b) = "E" Then
                                karyawan3 = "E"
                                sttugase = "SUDAH"
                                b = jmltiga
                                Call seleksikerjaan("E")
                            End If
                        Next

                        For b = 0 To jmlempat - 2
                            If Strings.Split(karyawan4, ",")(b) = "E" Then
                                karyawan4 = "E"
                                sttugase = "SUDAH"
                                b = jmlempat
                                Call seleksikerjaan("E")
                            End If
                        Next

                        For b = 0 To jmllima - 2
                            If Strings.Split(karyawan5, ",")(b) = "E" Then
                                karyawan5 = "E"
                                sttugase = "SUDAH"
                                b = jmllima
                                Call seleksikerjaan("E")
                            End If
                        Next

                        For b = 0 To jmlenam - 2
                            If Strings.Split(karyawan6, ",")(b) = "E" Then
                                karyawan6 = "E"
                                sttugase = "SUDAH"
                                b = jmlenam
                                Call seleksikerjaan("E")
                            End If
                        Next

                        For b = 0 To jmltujuh - 2
                            If Strings.Split(karyawan7, ",")(b) = "E" Then
                                karyawan7 = "E"
                                sttugase = "SUDAH"
                                b = jmltujuh
                                Call seleksikerjaan("E")
                            End If
                        Next
                    Else
                        kecil = anu(4)
                        keterangan = "E"
                    End If
                Else
                    kecil = kecil
                    keterangan = keterangan
                End If
                jmlsatu = Strings.Split(karyawan1, ",").Length
                jmldua = Strings.Split(karyawan2, ",").Length
                jmltiga = Strings.Split(karyawan3, ",").Length
                jmlempat = Strings.Split(karyawan4, ",").Length
                jmllima = Strings.Split(karyawan5, ",").Length
                jmlenam = Strings.Split(karyawan6, ",").Length
                jmltujuh = Strings.Split(karyawan7, ",").Length

                If anu(5) < kecil And sttugasf = "BELUM" Then
                    If anu(5) = 1 Then
                        For b = 0 To jmlsatu - 2
                            If Strings.Split(karyawan1, ",")(b) = "F" Then
                                karyawan1 = "F"
                                sttugasf = "SUDAH"
                                b = jmlsatu
                                Call seleksikerjaan("F")
                            End If
                        Next

                        For b = 0 To jmldua - 2
                            If Strings.Split(karyawan2, ",")(b) = "F" Then
                                karyawan2 = "F"
                                sttugasf = "SUDAH"
                                b = jmldua
                                Call seleksikerjaan("F")
                            End If
                        Next

                        For b = 0 To jmltiga - 2
                            If Strings.Split(karyawan3, ",")(b) = "F" Then
                                karyawan3 = "F"
                                sttugasf = "SUDAH"
                                b = jmltiga
                                Call seleksikerjaan("F")
                            End If
                        Next

                        For b = 0 To jmlempat - 2
                            If Strings.Split(karyawan4, ",")(b) = "F" Then
                                karyawan4 = "F"
                                sttugasf = "SUDAH"
                                b = jmlempat
                                Call seleksikerjaan("F")
                            End If
                        Next

                        For b = 0 To jmllima - 2
                            If Strings.Split(karyawan5, ",")(b) = "F" Then
                                karyawan5 = "F"
                                sttugasf = "SUDAH"
                                b = jmllima
                                Call seleksikerjaan("F")
                            End If
                        Next

                        For b = 0 To jmlenam - 2
                            If Strings.Split(karyawan6, ",")(b) = "F" Then
                                karyawan6 = "F"
                                sttugasf = "SUDAH"
                                b = jmlenam
                                Call seleksikerjaan("F")
                            End If
                        Next

                        For b = 0 To jmltujuh - 2
                            If Strings.Split(karyawan7, ",")(b) = "F" Then
                                karyawan7 = "F"
                                sttugasf = "SUDAH"
                                b = jmltujuh
                                Call seleksikerjaan("F")
                            End If
                        Next
                    Else
                        kecil = anu(5)
                        keterangan = "F"
                    End If
                Else
                    kecil = kecil
                    keterangan = keterangan
                End If
                jmlsatu = Strings.Split(karyawan1, ",").Length
                jmldua = Strings.Split(karyawan2, ",").Length
                jmltiga = Strings.Split(karyawan3, ",").Length
                jmlempat = Strings.Split(karyawan4, ",").Length
                jmllima = Strings.Split(karyawan5, ",").Length
                jmlenam = Strings.Split(karyawan6, ",").Length
                jmltujuh = Strings.Split(karyawan7, ",").Length


                If anu(6) < kecil And sttugasg = "BELUM" Then
                    If anu(6) = 1 Then
                        For b = 0 To jmlsatu - 2
                            If Strings.Split(karyawan1, ",")(b) = "G" Then
                                karyawan1 = "G"
                                sttugasg = "SUDAH"
                                b = jmlsatu
                                Call seleksikerjaan("G")
                            End If
                        Next

                        For b = 0 To jmldua - 2
                            If Strings.Split(karyawan2, ",")(b) = "G" Then
                                karyawan2 = "G"
                                sttugasg = "SUDAH"
                                b = jmldua
                                Call seleksikerjaan("G")
                            End If
                        Next

                        For b = 0 To jmltiga - 2
                            If Strings.Split(karyawan3, ",")(b) = "G" Then
                                karyawan3 = "G"
                                sttugasg = "SUDAH"
                                b = jmltiga
                                Call seleksikerjaan("G")
                            End If
                        Next

                        For b = 0 To jmlempat - 2
                            If Strings.Split(karyawan4, ",")(b) = "G" Then
                                karyawan4 = "G"
                                sttugasg = "SUDAH"
                                b = jmlempat
                                Call seleksikerjaan("G")
                            End If
                        Next

                        For b = 0 To jmllima - 2
                            If Strings.Split(karyawan5, ",")(b) = "G" Then
                                karyawan5 = "G"
                                sttugasg = "SUDAH"
                                b = jmllima
                                Call seleksikerjaan("G")
                            End If
                        Next

                        For b = 0 To jmlenam - 2
                            If Strings.Split(karyawan6, ",")(b) = "G" Then
                                karyawan6 = "G"
                                sttugasg = "SUDAH"
                                b = jmlenam
                                Call seleksikerjaan("G")
                            End If
                        Next

                        For b = 0 To jmltujuh - 2
                            If Strings.Split(karyawan7, ",")(b) = "G" Then
                                karyawan7 = "G"
                                sttugasg = "SUDAH"
                                b = jmltujuh
                                Call seleksikerjaan("G")
                            End If
                        Next
                    Else
                        kecil = anu(6)
                        keterangan = "G"
                    End If
                Else
                    kecil = kecil
                    keterangan = keterangan
                End If
                jmlsatu = Strings.Split(karyawan1, ",").Length
                jmldua = Strings.Split(karyawan2, ",").Length
                jmltiga = Strings.Split(karyawan3, ",").Length
                jmlempat = Strings.Split(karyawan4, ",").Length
                jmllima = Strings.Split(karyawan5, ",").Length
                jmlenam = Strings.Split(karyawan6, ",").Length
                jmltujuh = Strings.Split(karyawan7, ",").Length
            Next
            jmlsatu = Strings.Split(karyawan1, ",").Length
            jmldua = Strings.Split(karyawan2, ",").Length
            jmltiga = Strings.Split(karyawan3, ",").Length
            jmlempat = Strings.Split(karyawan4, ",").Length
            jmllima = Strings.Split(karyawan5, ",").Length
            jmlenam = Strings.Split(karyawan6, ",").Length
            jmltujuh = Strings.Split(karyawan7, ",").Length
            Dim kecil2 As Integer = 0
            Dim target1 As String = ""
            Dim target2 As String = ""

            'mencari nilai sesuai keterangan, jika nilai yang sesuai ada di karyawan1 maka nilai target bernilai karyawan1 
            For a = 0 To jmlsatu - 2
                If Strings.Split(karyawan1, ",")(a) = keterangan Then
                    If target1 = "" Then
                        target1 = "karyawan1"
                    Else
                        target2 = "karyawan1"
                    End If
                End If
            Next

            For a = 0 To jmldua - 2
                If Strings.Split(karyawan2, ",")(a) = keterangan Then
                    If target1 = "" Then
                        target1 = "karyawan2"
                    Else
                        target2 = "karyawan2"
                    End If
                End If
            Next

            For a = 0 To jmltiga - 2
                If Strings.Split(karyawan3, ",")(a) = keterangan Then
                    If target1 = "" Then
                        target1 = "karyawan3"
                    Else
                        target2 = "karyawan3"
                    End If
                End If
            Next

            For a = 0 To jmlempat - 2
                If Strings.Split(karyawan4, ",")(a) = keterangan Then
                    If target1 = "" Then
                        target1 = "karyawan4"
                    Else
                        target2 = "karyawan4"
                    End If
                End If
            Next

            For a = 0 To jmllima - 2
                If Strings.Split(karyawan5, ",")(a) = keterangan Then
                    If target1 = "" Then
                        target1 = "karyawan5"
                    Else
                        target2 = "karyawan5"
                    End If
                End If
            Next

            For a = 0 To jmlenam - 2
                If Strings.Split(karyawan6, ",")(a) = keterangan Then
                    If target1 = "" Then
                        target1 = "karyawan6"
                    Else
                        target2 = "karyawan6"
                    End If
                End If
            Next

            For a = 0 To jmltujuh - 2
                If Strings.Split(karyawan7, ",")(a) = keterangan Then
                    If target1 = "" Then
                        target1 = "karyawan7"
                    Else
                        target2 = "karyawan7"
                    End If
                End If
            Next

            Dim jml1 As Integer = 0
            Dim jml2 As Integer = 0
            'mencari jumlah tugas setiap karyawan yang sesuai target 1 dan 2
            If target1 = "karyawan1" Then
                jml1 = Strings.Split(karyawan1, ",").Length - 2
            ElseIf target1 = "karyawan2" Then
                jml1 = Strings.Split(karyawan2, ",").Length - 2
            ElseIf target1 = "karyawan3" Then
                jml1 = Strings.Split(karyawan3, ",").Length - 2
            ElseIf target1 = "karyawan4" Then
                jml1 = Strings.Split(karyawan4, ",").Length - 2
            ElseIf target1 = "karyawan5" Then
                jml1 = Strings.Split(karyawan5, ",").Length - 2
            ElseIf target1 = "karyawan6" Then
                jml1 = Strings.Split(karyawan6, ",").Length - 2
            ElseIf target1 = "karyawan7" Then
                jml1 = Strings.Split(karyawan7, ",").Length - 2
            End If

            If target2 = "" Then
                jml2 = 100
            ElseIf target2 = "karyawan1" Then
                jml2 = Strings.Split(karyawan1, ",").Length - 2
            ElseIf target2 = "karyawan2" Then
                jml2 = Strings.Split(karyawan2, ",").Length - 2
            ElseIf target2 = "karyawan3" Then
                jml2 = Strings.Split(karyawan3, ",").Length - 2
            ElseIf target2 = "karyawan4" Then
                jml2 = Strings.Split(karyawan4, ",").Length - 2
            ElseIf target2 = "karyawan5" Then
                jml2 = Strings.Split(karyawan5, ",").Length - 2
            ElseIf target2 = "karyawan6" Then
                jml2 = Strings.Split(karyawan6, ",").Length - 2
            ElseIf target2 = "karyawan7" Then
                jml2 = Strings.Split(karyawan7, ",").Length - 2
            End If

            'membandingkan target 1 dan 2, jika target1 kurang dari target2 maka karyawan yang sesuai target satu akan memperoleh tugas sesuai keterangan
            If jml1 <= jml2 Then
                If target1 = "karyawan1" Then
                    karyawan1 = keterangan
                ElseIf target1 = "karyawan2" Then
                    karyawan2 = keterangan
                ElseIf target1 = "karyawan3" Then
                    karyawan3 = keterangan
                ElseIf target1 = "karyawan4" Then
                    karyawan4 = keterangan
                ElseIf target1 = "karyawan5" Then
                    karyawan5 = keterangan
                ElseIf target1 = "karyawan6" Then
                    karyawan6 = keterangan
                ElseIf target1 = "karyawan7" Then
                    karyawan7 = keterangan
                End If
            Else
                If target2 = "karyawan1" Then
                    karyawan1 = keterangan
                ElseIf target2 = "karyawan2" Then
                    karyawan2 = keterangan
                ElseIf target2 = "karyawan3" Then
                    karyawan3 = keterangan
                ElseIf target2 = "karyawan4" Then
                    karyawan4 = keterangan
                ElseIf target2 = "karyawan5" Then
                    karyawan5 = keterangan
                ElseIf target2 = "karyawan6" Then
                    karyawan6 = keterangan
                ElseIf target2 = "karyawan7" Then
                    karyawan7 = keterangan
                End If
            End If
            'setelah tugas sudah dimiliki karyawan maka akan bernilai SUDAH agar tidak diproses kembali di pencarian nilai
            If keterangan = "A" Then
                sttugasa = "SUDAH"
            ElseIf keterangan = "B" Then
                sttugasb = "SUDAH"
            ElseIf keterangan = "C" Then
                sttugasc = "SUDAH"
            ElseIf keterangan = "D" Then
                sttugasd = "SUDAH"
            ElseIf keterangan = "E" Then
                sttugase = "SUDAH"
            ElseIf keterangan = "F" Then
                sttugasf = "SUDAH"
            ElseIf keterangan = "G" Then
                sttugasg = "SUDAH"
            End If

            Dim numpang11 As String = ""
            Dim numpang222 As String = ""
            Dim numpang333 As String = ""
            Dim numpang444 As String = ""
            Dim numpang555 As String = ""
            Dim numpang666 As String = ""
            Dim numpang777 As String = ""

            'proses pengambilan tugas yang tidak sesuai dengan keterangan
            If Not Strings.Split(karyawan1, ",").Length = 1 Then
                For a = 0 To jmlsatu - 2
                    If Not Strings.Split(karyawan1, ",")(a) = keterangan Then
                        If Strings.Split(karyawan1, ",").Length = 2 Then
                            numpang11 += Strings.Split(karyawan1, ",")(a)
                        Else
                            numpang11 += Strings.Split(karyawan1, ",")(a) + ","
                        End If
                    End If
                Next
                If numpang11 = "" Then
                    numpang11 = karyawan1
                Else
                    karyawan1 = numpang11
                End If
            End If

            If Not Strings.Split(karyawan2, ",").Length = 1 Then
                For a = 0 To jmldua - 2
                    If Not Strings.Split(karyawan2, ",")(a) = keterangan Then
                        If Strings.Split(karyawan2, ",").Length = 2 Then
                            numpang222 += Strings.Split(karyawan2, ",")(a)
                        Else
                            numpang222 += Strings.Split(karyawan2, ",")(a) + ","
                        End If
                    End If
                Next
                If numpang222 = "" Then
                    numpang222 = karyawan2
                Else
                    karyawan2 = numpang222
                End If
            End If

            If Not Strings.Split(karyawan3, ",").Length = 1 Then
                For a = 0 To jmltiga - 2
                    If Not Strings.Split(karyawan3, ",")(a) = keterangan Then
                        If Strings.Split(karyawan3, ",").Length = 2 Then
                            numpang333 += Strings.Split(karyawan3, ",")(a)
                        Else
                            numpang333 += Strings.Split(karyawan3, ",")(a) + ","
                        End If
                    End If
                Next
                If numpang333 = "" Then
                    numpang333 = karyawan3
                Else
                    karyawan3 = numpang333
                End If
            End If

            If Not Strings.Split(karyawan4, ",").Length = 1 Then
                For a = 0 To jmlempat - 2
                    If Not Strings.Split(karyawan4, ",")(a) = keterangan Then
                        If Strings.Split(karyawan4, ",").Length = 2 Then
                            numpang444 += Strings.Split(karyawan4, ",")(a)
                        Else
                            numpang444 += Strings.Split(karyawan4, ",")(a) + ","
                        End If
                    End If
                Next
                If numpang444 = "" Then
                    numpang444 = karyawan4
                Else
                    karyawan4 = numpang444
                End If
            End If

            If Not Strings.Split(karyawan5, ",").Length = 1 Then
                For a = 0 To jmllima - 2
                    If Not Strings.Split(karyawan5, ",")(a) = keterangan Then
                        If Strings.Split(karyawan5, ",").Length = 2 Then
                            numpang555 += Strings.Split(karyawan5, ",")(a)
                        Else
                            numpang555 += Strings.Split(karyawan5, ",")(a) + ","
                        End If
                    End If
                Next
                If numpang555 = "" Then
                    numpang555 = karyawan5
                Else
                    karyawan5 = numpang555
                End If
            End If

            If Not Strings.Split(karyawan6, ",").Length = 1 Then
                For a = 0 To jmlenam - 2
                    If Not Strings.Split(karyawan6, ",")(a) = keterangan Then
                        If Strings.Split(karyawan6, ",").Length = 2 Then
                            numpang666 += Strings.Split(karyawan6, ",")(a)
                        Else
                            numpang666 += Strings.Split(karyawan6, ",")(a) + ","
                        End If
                    End If
                Next
                If numpang666 = "" Then
                    numpang666 = karyawan6
                Else
                    karyawan6 = numpang666
                End If
            End If

            If Not Strings.Split(karyawan7, ",").Length = 1 Then
                For a = 0 To jmltujuh - 2
                    If Not Strings.Split(karyawan7, ",")(a) = keterangan Then
                        If Strings.Split(karyawan7, ",").Length = 2 Then
                            numpang777 += Strings.Split(karyawan7, ",")(a)
                        Else
                            numpang777 += Strings.Split(karyawan7, ",")(a) + ","
                        End If
                    End If
                Next
                If numpang777 = "" Then
                    numpang777 = karyawan7
                Else
                    karyawan7 = numpang777
                End If
            End If

            jmlsatu = Strings.Split(karyawan1, ",").Length
            jmldua = Strings.Split(karyawan2, ",").Length
            jmltiga = Strings.Split(karyawan3, ",").Length
            jmlempat = Strings.Split(karyawan4, ",").Length
            jmllima = Strings.Split(karyawan5, ",").Length
            jmlenam = Strings.Split(karyawan6, ",").Length
            jmltujuh = Strings.Split(karyawan7, ",").Length
        Loop

        'menampilkan tugas final setiap karyawan di datagrid2
        For a = 0 To 6
            If a = 0 Then
                DataGridView2.Item(3, a).Value = Strings.Left(karyawan1, 1)
            ElseIf a = 1 Then
                DataGridView2.Item(3, a).Value = Strings.Left(karyawan2, 1)
            ElseIf a = 2 Then
                DataGridView2.Item(3, a).Value = Strings.Left(karyawan3, 1)
            ElseIf a = 3 Then
                DataGridView2.Item(3, a).Value = Strings.Left(karyawan4, 1)
            ElseIf a = 4 Then
                DataGridView2.Item(3, a).Value = Strings.Left(karyawan5, 1)
            ElseIf a = 5 Then
                DataGridView2.Item(3, a).Value = Strings.Left(karyawan6, 1)
            ElseIf a = 6 Then
                DataGridView2.Item(3, a).Value = Strings.Left(karyawan7, 1)
            End If
        Next
        For a = 0 To 6
            Dim simpan As String = "insert into tb_tugasbagi values ('" & DataGridView2.Item(0, a).Value & "','" & DataGridView2.Item(1, a).Value & "','" & DataGridView2.Item(2, a).Value & "','" & DataGridView2.Item(3, a).Value & "')"
            CMD = New OdbcCommand(simpan, CONN)
            CMD.ExecuteNonQuery()
        Next

        karyawan1 = ""
        karyawan2 = ""
        karyawan3 = ""
        karyawan4 = ""
        karyawan5 = ""
        karyawan6 = ""
        karyawan7 = ""
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If DataGridView6.RowCount < 8 Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or ComboBox1.Text = "" Then
                MsgBox("Karyawan Belum Terpilih", vbCritical + vbOKOnly, "Peringatan")
            Else
                CMD = New OdbcCommand("select * FROM tb_produksi where kd_karyawan= '" & ComboBox1.Text & "'", CONN)
                RD = CMD.ExecuteReader
                RD.Read()
                DataGridView6.Rows.Add()
                For a = 0 To 7
                    DataGridView6.Item(a, baris).Value = RD.Item(a)
                Next
                baris += 1
                ComboBox1.Items.Clear()
                ComboBox1.Text = ""
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                isicombo()
            End If
        Else
            MsgBox("Jumlah Karyawan Melebihi Jumlah Tugas", vbCritical + vbOKOnly, "Peringatan")
            ComboBox1.Items.Clear()
            ComboBox1.Text = ""
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            isicombo()
        End If
    End Sub
    Sub isicombo()
        CMD = New OdbcCommand("select kd_karyawan FROM karyawan", CONN)
        RD = CMD.ExecuteReader
        Do While RD.Read
            Dim stts As String = "NO"
            For a = 0 To DataGridView6.RowCount - 1
                If RD.Item(0) = DataGridView6.Item(0, a).Value Then
                    stts = "YES"
                End If
            Next
            If stts = "NO" Then
                ComboBox1.Items.Add(RD.Item(0))
            End If
        Loop
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        CMD = New OdbcCommand("select * FROM karyawan where kd_karyawan= '" & ComboBox1.Text & "'", CONN)
        RD = CMD.ExecuteReader
        RD.Read()
        TextBox1.Text = RD.Item(1)
        TextBox2.Text = RD.Item(2)
        TextBox3.Text = RD.Item(3)
        TextBox4.Text = RD.Item(4)
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        DataGridView2.Rows.Clear()
        DataGridView3.Rows.Clear()
        DataGridView4.Rows.Clear()
        DataGridView5.Rows.Clear()
        DataGridView6.Rows.Clear()
        ComboBox1.Items.Clear()
        isicombo()
        baris = 0
        Button5.Enabled = False
        Dim truncate As String = "TRUNCATE tb_tugasbagi"
        CMD = New OdbcCommand(truncate, CONN)
        CMD.ExecuteNonQuery()
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        laporanbagitugas.Show()
    End Sub
    Sub seleksikerjaan(ByVal keterangan As String)
        Dim jmlsatu, jmldua, jmltiga, jmlempat, jmllima, jmlenam, jmltujuh As Integer
        jmlsatu = Strings.Split(karyawan1, ",").Length
        jmldua = Strings.Split(karyawan2, ",").Length
        jmltiga = Strings.Split(karyawan3, ",").Length
        jmlempat = Strings.Split(karyawan4, ",").Length
        jmllima = Strings.Split(karyawan5, ",").Length
        jmlenam = Strings.Split(karyawan6, ",").Length
        jmltujuh = Strings.Split(karyawan7, ",").Length
        Dim numpang11 As String = ""
        Dim numpang222 As String = ""
        Dim numpang333 As String = ""
        Dim numpang444 As String = ""
        Dim numpang555 As String = ""
        Dim numpang666 As String = ""
        Dim numpang777 As String = ""

        'proses pengambilan tugas yang tidak sesuai dengan keterangan
        If Not Strings.Split(karyawan1, ",").Length = 1 Then
            For a = 0 To jmlsatu - 2
                If Not Strings.Split(karyawan1, ",")(a) = keterangan Then
                        numpang11 += Strings.Split(karyawan1, ",")(a) + ","
                End If
            Next
            If numpang11 = "" Then
                numpang11 = karyawan1
            Else
                karyawan1 = numpang11
            End If
        End If

        If Not Strings.Split(karyawan2, ",").Length = 1 Then
            For a = 0 To jmldua - 2
                If Not Strings.Split(karyawan2, ",")(a) = keterangan Then
                        numpang222 += Strings.Split(karyawan2, ",")(a) + ","
                End If
            Next
            If numpang222 = "" Then
                numpang222 = karyawan2
            Else
                karyawan2 = numpang222
            End If
        End If

        If Not Strings.Split(karyawan3, ",").Length = 1 Then
            For a = 0 To jmltiga - 2
                If Not Strings.Split(karyawan3, ",")(a) = keterangan Then
                        numpang333 += Strings.Split(karyawan3, ",")(a) + ","
                End If
            Next
            If numpang333 = "" Then
                numpang333 = karyawan3
            Else
                karyawan3 = numpang333
            End If
        End If

        If Not Strings.Split(karyawan4, ",").Length = 1 Then
            For a = 0 To jmlempat - 2
                If Not Strings.Split(karyawan4, ",")(a) = keterangan Then
                        numpang444 += Strings.Split(karyawan4, ",")(a) + ","
                End If
            Next
            If numpang444 = "" Then
                numpang444 = karyawan4
            Else
                karyawan4 = numpang444
            End If
        End If

        If Not Strings.Split(karyawan5, ",").Length = 1 Then
            For a = 0 To jmllima - 2
                If Not Strings.Split(karyawan5, ",")(a) = keterangan Then
                        numpang555 += Strings.Split(karyawan5, ",")(a) + ","
                End If
            Next
            If numpang555 = "" Then
                numpang555 = karyawan5
            Else
                karyawan5 = numpang555
            End If
        End If

        If Not Strings.Split(karyawan6, ",").Length = 1 Then
            For a = 0 To jmlenam - 2
                If Not Strings.Split(karyawan6, ",")(a) = keterangan Then
                        numpang666 += Strings.Split(karyawan6, ",")(a) + ","
                End If
            Next
            If numpang666 = "" Then
                numpang666 = karyawan6
            Else
                karyawan6 = numpang666
            End If
        End If

        If Not Strings.Split(karyawan7, ",").Length = 1 Then
            For a = 0 To jmltujuh - 2
                If Not Strings.Split(karyawan7, ",")(a) = keterangan Then
                        numpang777 += Strings.Split(karyawan7, ",")(a) + ","
                End If
            Next
            If numpang777 = "" Then
                numpang777 = karyawan7
            Else
                karyawan7 = numpang777
            End If
        End If
    End Sub
End Class