Imports System.Data.SqlClient
Public Class formPiutang
    Dim str = "Data Source=DESKTOP-KC1E6QK\SQLEXPRESS;Initial Catalog=aplikasiSIA;Integrated Security=True"
    Dim koneksi As New SqlConnection(str)

    Sub bersih()
        TextBox1.Text = ""
        ComboBox1.Text = ""
        TextBox2.Text = ""
        ComboBox2.Text = ""
        TextBox3.Text = ""
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = True
        TextBox1.Focus()
    End Sub

    Protected Sub tampilPiutang(ByVal xGrid As DataGridView)
        With xGrid
            .ColumnCount = 6
            .Columns(0).Name = "No Faktur"
            .Columns(1).Name = "Tanggal Penerimaan"
            .Columns(2).Name = "ID User"
            .Columns(3).Name = "ID Pelanggan"
            .Columns(4).Name = "Nama Pelanggan"
            .Columns(5).Name = "Piutang"
            .Rows.Clear()
        End With
        Dim sSql As String
        sSql = "Select * from tbPiutang"
        Dim cmd As New SqlCommand(sSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    Dim baris(5) As String
                    baris(0) = rd(0) 'No Faktur
                    baris(1) = rd(1) 'Tanggal Penerimaan
                    baris(2) = rd(2) 'ID User
                    baris(3) = rd(3) 'ID Pelanggan
                    baris(4) = rd(4) 'Nama Pelanggan
                    baris(5) = rd(5) 'Piutang
                    xGrid.Rows.Add(baris)
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()
    End Sub

    Private Sub formPiutang_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call tampilPiutang(DataGridView1)

        Dim sSql As String
        sSql = "select * from tbUser"
        Dim cmd As New SqlCommand(sSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    ComboBox1.Items.Add(rd.Item("idUSer"))
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()

        Dim aSql As String
        aSql = "select * from tbPelanggan"
        Dim com As New SqlCommand(aSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = com.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    ComboBox2.Items.Add(rd.Item("namaPelanggan"))
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Me.TextBox1.Text = Me.DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
        Me.DateTimePicker1.Text = Me.DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value
        Me.ComboBox1.Text = Me.DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value
        Me.TextBox2.Text = Me.DataGridView1.Item(3, DataGridView1.CurrentRow.Index).Value
        Me.ComboBox2.Text = Me.DataGridView1.Item(4, DataGridView1.CurrentRow.Index).Value
        Me.TextBox3.Text = Me.DataGridView1.Item(5, DataGridView1.CurrentRow.Index).Value
        TextBox1.Enabled = False
        Button1.Enabled = True
        Button2.Enabled = False
        Button3.Enabled = True
        Button4.Enabled = True
        Button5.Enabled = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call bersih()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or ComboBox1.Text = "" Or TextBox2.Text = "" Or ComboBox2.Text = "" Or TextBox3.Text = "" Then
            MsgBox("Data Masih Kosong", MsgBoxStyle.Information, "Perhatian!")
        Else
            Dim cmd As New SqlCommand("tambahPiutang", koneksi)
            Dim xParam As New SqlParameter
            xParam.Direction = ParameterDirection.Input
            Try
                koneksi.Open()
                With cmd
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@noFaktur", SqlDbType.Char, 10).Value = TextBox1.Text
                    .Parameters.Add("@tglTerima", SqlDbType.DateTime).Value = DateTimePicker1.Text
                    .Parameters.Add("@idUser", SqlDbType.Char, 10).Value = ComboBox1.Text
                    .Parameters.Add("@idPelanggan", SqlDbType.Char, 10).Value = TextBox2.Text
                    .Parameters.Add("@namaPelanggan", SqlDbType.VarChar, 50).Value = ComboBox2.Text
                    .Parameters.Add("@piutang", SqlDbType.Money).Value = TextBox3.Text
                    .ExecuteNonQuery()
                End With
                MsgBox("Disimpan!", MsgBoxStyle.Information, "Perhatian!")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
            Finally
                koneksi.Close()
            End Try
            Call tampilPiutang(DataGridView1)
        End If
        Call bersih()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim cmd As New SqlCommand("ubahPiutang", koneksi)
        Dim xParam As New SqlParameter
        xParam.Direction = ParameterDirection.Input
        Try
            koneksi.Open()
            With cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add("@noFaktur", SqlDbType.Char, 10).Value = TextBox1.Text
                .Parameters.Add("@tglTerima", SqlDbType.DateTime).Value = DateTimePicker1.Text
                .Parameters.Add("@idUser", SqlDbType.Char, 10).Value = ComboBox1.Text
                .Parameters.Add("@idPelanggan", SqlDbType.Char, 10).Value = TextBox2.Text
                .Parameters.Add("@namaPelanggan", SqlDbType.VarChar, 50).Value = ComboBox2.Text
                .Parameters.Add("@piutang", SqlDbType.Money).Value = TextBox3.Text
                .ExecuteNonQuery()
            End With
            MsgBox("Diubah!", MsgBoxStyle.Information, "Perhatian!")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
        Finally
            koneksi.Close()
        End Try
        Call tampilPiutang(DataGridView1)
        Call bersih()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If MessageBox.Show("Yakin Akan Dihapus?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Dim cmd As New SqlCommand("hapusPiutang", koneksi)
            Dim xParam As New SqlParameter
            xParam.Direction = ParameterDirection.Input
            Try
                koneksi.Open()
                With cmd
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@noFaktur", SqlDbType.Char, 10).Value = TextBox1.Text
                    .ExecuteNonQuery()
                End With
                MsgBox("Dihapus!", MsgBoxStyle.Information, "Perhatian!")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
            Finally
                koneksi.Close()
            End Try
            Call tampilPiutang(DataGridView1)
            Call bersih()
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim sSql As String
        sSql = "select * from tbPelanggan where namaPelanggan='" & ComboBox2.Text & "'"
        Dim cmd As New SqlCommand(sSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    TextBox2.Text = rd.Item("idPelanggan")
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        formMenuUtama.Show()
    End Sub
End Class