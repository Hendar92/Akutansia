Imports System.Data.SqlClient
Public Class formJurnal
    Dim str = "Data Source=DESKTOP-KC1E6QK\SQLEXPRESS;Initial Catalog=aplikasiSIA;Integrated Security=True"
    Dim koneksi As New SqlConnection(str)

    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox2.Text = ""
        TextBox5.Text = ""
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = True
        TextBox1.Focus()
    End Sub

    Protected Sub tampilJurnal(ByVal xGrid As DataGridView)
        With xGrid
            .ColumnCount = 8
            .Columns(0).Name = "Kode Jurnal"
            .Columns(1).Name = "Tanggal"
            .Columns(2).Name = "No Bukti"
            .Columns(3).Name = "Keterangan"
            .Columns(4).Name = "Kode Akun"
            .Columns(5).Name = "Nama Akun"
            .Columns(6).Name = "Debit"
            .Columns(7).Name = "Kredit"
            .Rows.Clear()
        End With
        Dim sSql As String
        sSql = "Select * from tbJurnal"
        Dim cmd As New SqlCommand(sSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    Dim baris(7) As String
                    baris(0) = rd(0) 'Kode Jurnal
                    baris(1) = rd(1) 'Tanggal
                    baris(2) = rd(2) 'No Bukti
                    baris(3) = rd(3) 'Keterangan
                    baris(4) = rd(4) 'Kode Akun
                    baris(5) = rd(5) 'Nama Akun
                    baris(6) = rd(6) 'Debit
                    baris(7) = rd(7) 'Kredit
                    xGrid.Rows.Add(baris)
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()
    End Sub

    Private Sub formJurnal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call tampilJurnal(DataGridView1)
        Dim sSql As String
        sSql = "select * from tbBagan"
        Dim cmd As New SqlCommand(sSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    ComboBox1.Items.Add(rd.Item("namaAkun"))
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call bersih()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or
            ComboBox1.Text = "" Or ComboBox2.Text = "" Or TextBox5.Text = "" Then
            MsgBox("Data Masih Kosong", MsgBoxStyle.Information, "Perhatian!")
        Else
            Dim cmd As New SqlCommand("tambahJurnal", koneksi)
            Dim xParam As New SqlParameter
            xParam.Direction = ParameterDirection.Input
            Try
                koneksi.Open()
                With cmd
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@kodeJurnal", SqlDbType.Char, 10).Value = TextBox1.Text
                    .Parameters.Add("@tanggal", SqlDbType.DateTime).Value = DateTimePicker1.Text
                    .Parameters.Add("@noBukti", SqlDbType.Char, 10).Value = TextBox2.Text
                    .Parameters.Add("@keterangan", SqlDbType.VarChar, 50).Value = TextBox3.Text
                    .Parameters.Add("@kodeAkun", SqlDbType.Char, 10).Value = TextBox4.Text
                    .Parameters.Add("@namaAkun", SqlDbType.VarChar, 50).Value = ComboBox1.Text
                    If ComboBox2.Text = "Debit" Then
                        .Parameters.Add("@debit", SqlDbType.Money).Value = TextBox5.Text
                        .Parameters.Add("@kredit", SqlDbType.Money).Value = "0"
                    Else
                        .Parameters.Add("@debit", SqlDbType.Money).Value = "0"
                        .Parameters.Add("@kredit", SqlDbType.Money).Value = TextBox5.Text
                    End If
                    .ExecuteNonQuery()
                End With
                MsgBox("Disimpan!", MsgBoxStyle.Information, "Perhatian!")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
            Finally
                koneksi.Close()
            End Try
            Call tampilJurnal(DataGridView1)
        End If
        Call bersih()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        formMenuUtama.Show()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim sSql As String
        sSql = "select * from tbBagan where namaAkun='" & ComboBox1.Text & "'"
        Dim cmd As New SqlCommand(sSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    TextBox4.Text = rd.Item("kodeAkun")
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()
    End Sub
End Class