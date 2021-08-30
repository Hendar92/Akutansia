Imports System.Data.SqlClient
Public Class formBarang
    Dim str = "Data Source=DESKTOP-KC1E6QK\SQLEXPRESS;Initial Catalog=aplikasiSIA;Integrated Security=True"
    Dim koneksi As New SqlConnection(str)

    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = True
        TextBox1.Focus()
    End Sub

    Protected Sub tampilBarang(ByVal xGrid As DataGridView)
        With xGrid
            .ColumnCount = 5
            .Columns(0).Name = "ID Barang"
            .Columns(1).Name = "Nama Barang"
            .Columns(2).Name = "Harga Beli"
            .Columns(3).Name = "Harga Jual"
            .Columns(4).Name = "Stok"
            .Rows.Clear()
        End With
        Dim sSql As String
        sSql = "Select * from tbBarang"
        Dim cmd As New SqlCommand(sSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    Dim baris(4) As String
                    baris(0) = rd(0) 'ID Barang
                    baris(1) = rd(1) 'Nama Barang
                    baris(2) = rd(2) 'Harga Beli
                    baris(3) = rd(3) 'Harga Jual
                    baris(4) = rd(4) 'Stok
                    xGrid.Rows.Add(baris)
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()
    End Sub

    Private Sub formBarang_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call tampilBarang(DataGridView1)
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Me.TextBox1.Text = Me.DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
        Me.TextBox2.Text = Me.DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value
        Me.TextBox3.Text = Me.DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value
        Me.TextBox4.Text = Me.DataGridView1.Item(3, DataGridView1.CurrentRow.Index).Value
        Me.TextBox5.Text = Me.DataGridView1.Item(4, DataGridView1.CurrentRow.Index).Value
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
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then
            MsgBox("Data Masih Kosong", MsgBoxStyle.Information, "Perhatian!")
        Else
            Dim cmd As New SqlCommand("tambahBarang", koneksi)
            Dim xParam As New SqlParameter
            xParam.Direction = ParameterDirection.Input
            Try
                koneksi.Open()
                With cmd
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@idBarang", SqlDbType.Char, 10).Value = TextBox1.Text
                    .Parameters.Add("@namaBarang", SqlDbType.VarChar, 50).Value = TextBox2.Text
                    .Parameters.Add("@hargaBeli", SqlDbType.Money).Value = TextBox3.Text
                    .Parameters.Add("@hargaJual", SqlDbType.Money).Value = TextBox4.Text
                    .Parameters.Add("@stok", SqlDbType.Int).Value = TextBox5.Text
                    .ExecuteNonQuery()
                End With
                MsgBox("Disimpan!", MsgBoxStyle.Information, "Perhatian!")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
            Finally
                koneksi.Close()
            End Try
            Call tampilBarang(DataGridView1)
        End If
        Call bersih()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim cmd As New SqlCommand("ubahBarang", koneksi)
        Dim xParam As New SqlParameter
        xParam.Direction = ParameterDirection.Input
        Try
            koneksi.Open()
            With cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add("@idBarang", SqlDbType.Char, 10).Value = TextBox1.Text
                .Parameters.Add("@namaBarang", SqlDbType.VarChar, 50).Value = TextBox2.Text
                .Parameters.Add("@hargaBeli", SqlDbType.Money).Value = TextBox3.Text
                .Parameters.Add("@hargaJual", SqlDbType.Money).Value = TextBox4.Text
                .Parameters.Add("@stok", SqlDbType.Int).Value = TextBox5.Text
                .ExecuteNonQuery()
            End With
            MsgBox("Diubah!", MsgBoxStyle.Information, "Perhatian!")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
        Finally
            koneksi.Close()
        End Try
        Call tampilBarang(DataGridView1)
        Call bersih()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If MessageBox.Show("Yakin Akan Dihapus?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Dim cmd As New SqlCommand("hapusBarang", koneksi)
            Dim xParam As New SqlParameter
            xParam.Direction = ParameterDirection.Input
            Try
                koneksi.Open()
                With cmd
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@idBarang", SqlDbType.Char, 10).Value = TextBox1.Text
                    .ExecuteNonQuery()
                End With
                MsgBox("Dihapus!", MsgBoxStyle.Information, "Perhatian!")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
            Finally
                koneksi.Close()
            End Try
            Call tampilBarang(DataGridView1)
            Call bersih()
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        formMenuUtama.Show()
    End Sub
End Class