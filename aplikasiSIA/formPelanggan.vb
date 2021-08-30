Imports System.Data.SqlClient
Public Class formPelanggan
    Dim str = "Data Source=DESKTOP-KC1E6QK\SQLEXPRESS;Initial Catalog=aplikasiSIA;Integrated Security=True"
    Dim koneksi As New SqlConnection(str)

    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = True
        TextBox1.Focus()
    End Sub

    Protected Sub tampilPelanggan(ByVal xGrid As DataGridView)
        With xGrid
            .ColumnCount = 4
            .Columns(0).Name = "ID Pelanggan"
            .Columns(1).Name = "Nama Pelanggan"
            .Columns(2).Name = "Alamat"
            .Columns(3).Name = "Telepon"
            .Rows.Clear()
        End With
        Dim sSql As String
        sSql = "Select * from tbPelanggan"
        Dim cmd As New SqlCommand(sSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    Dim baris(3) As String
                    baris(0) = rd(0) 'ID Pelanggan
                    baris(1) = rd(1) 'Nama Pelanggan
                    baris(2) = rd(2) 'Alamat
                    baris(3) = rd(3) 'Telepon
                    xGrid.Rows.Add(baris)
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()
    End Sub

    Private Sub formPelanggan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call tampilPelanggan(DataGridView1)
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Me.TextBox1.Text = Me.DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
        Me.TextBox2.Text = Me.DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value
        Me.TextBox3.Text = Me.DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value
        Me.TextBox4.Text = Me.DataGridView1.Item(3, DataGridView1.CurrentRow.Index).Value
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
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            MsgBox("Data Masih Kosong", MsgBoxStyle.Information, "Perhatian!")
        Else
            Dim cmd As New SqlCommand("tambahPelanggan", koneksi)
            Dim xParam As New SqlParameter
            xParam.Direction = ParameterDirection.Input
            Try
                koneksi.Open()
                With cmd
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@idPelanggan", SqlDbType.Char, 10).Value = TextBox1.Text
                    .Parameters.Add("@namaPelanggan", SqlDbType.VarChar, 50).Value = TextBox2.Text
                    .Parameters.Add("@alamat", SqlDbType.VarChar, 50).Value = TextBox3.Text
                    .Parameters.Add("@telepon", SqlDbType.VarChar, 50).Value = TextBox4.Text
                    .ExecuteNonQuery()
                End With
                MsgBox("Disimpan!", MsgBoxStyle.Information, "Perhatian!")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
            Finally
                koneksi.Close()
            End Try
            Call tampilPelanggan(DataGridView1)
        End If
        Call bersih()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim cmd As New SqlCommand("ubahPelanggan", koneksi)
        Dim xParam As New SqlParameter
        xParam.Direction = ParameterDirection.Input
        Try
            koneksi.Open()
            With cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add("@idPelanggan", SqlDbType.Char, 10).Value = TextBox1.Text
                .Parameters.Add("@namaPelanggan", SqlDbType.VarChar, 50).Value = TextBox2.Text
                .Parameters.Add("@alamat", SqlDbType.VarChar, 50).Value = TextBox3.Text
                .Parameters.Add("@telepon", SqlDbType.VarChar, 50).Value = TextBox4.Text
                .ExecuteNonQuery()
            End With
            MsgBox("Diubah!", MsgBoxStyle.Information, "Perhatian!")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
        Finally
            koneksi.Close()
        End Try
        Call tampilPelanggan(DataGridView1)
        Call bersih()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If MessageBox.Show("Yakin Akan Dihapus?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Dim cmd As New SqlCommand("hapusPelanggan", koneksi)
            Dim xParam As New SqlParameter
            xParam.Direction = ParameterDirection.Input
            Try
                koneksi.Open()
                With cmd
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@idPelanggan", SqlDbType.Char, 10).Value = TextBox1.Text
                    .ExecuteNonQuery()
                End With
                MsgBox("Dihapus!", MsgBoxStyle.Information, "Perhatian!")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
            Finally
                koneksi.Close()
            End Try
            Call tampilPelanggan(DataGridView1)
            Call bersih()
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        formMenuUtama.Show()
    End Sub
End Class