Imports System.Data.SqlClient
Public Class formPembelian
    Dim str = "Data Source=DESKTOP-KC1E6QK\SQLEXPRESS;Initial Catalog=aplikasiSIA;Integrated Security=True"
    Dim koneksi As New SqlConnection(str)

    Sub bersih()
        TextBox1.Text = ""
        ComboBox1.Text = ""
        TextBox2.Text = ""
        ComboBox2.Text = ""
        TextBox3.Text = ""
        ComboBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = True
        TextBox1.Focus()
    End Sub

    Protected Sub tampilBeli(ByVal xGrid As DataGridView)
        With xGrid
            .ColumnCount = 10
            .Columns(0).Name = "No Faktur"
            .Columns(1).Name = "Tanggal Pembelian"
            .Columns(2).Name = "ID User"
            .Columns(3).Name = "ID Pemasok"
            .Columns(4).Name = "Nama Pemasok"
            .Columns(5).Name = "ID Barang"
            .Columns(6).Name = "Nama Barang"
            .Columns(7).Name = "Harga Barang"
            .Columns(8).Name = "Jumlah Terima"
            .Columns(9).Name = "Total Bayar"
            .Rows.Clear()
        End With
        Dim sSql As String
        sSql = "Select * from tbBeli"
        Dim cmd As New SqlCommand(sSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    Dim baris(9) As String
                    baris(0) = rd(0) 'No Faktur
                    baris(1) = rd(1) 'Tanggal Pembelian
                    baris(2) = rd(2) 'ID User
                    baris(3) = rd(3) 'ID Pemasok
                    baris(4) = rd(4) 'Nama Pemasok
                    baris(5) = rd(5) 'ID Barang
                    baris(6) = rd(6) 'Nama Barang
                    baris(7) = rd(7) 'Harga Barang
                    baris(8) = rd(8) 'Jumlah Terima
                    baris(9) = rd(9) 'Total Bayar
                    xGrid.Rows.Add(baris)
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()
    End Sub

    Private Sub formPembelian_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call tampilBeli(DataGridView1)

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
        aSql = "select * from tbPemasok"
        Dim com As New SqlCommand(aSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = com.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    ComboBox2.Items.Add(rd.Item("namaPemasok"))
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()

        Dim Sql As String
        Sql = "select * from tbBarang"
        Dim cmm As New SqlCommand(Sql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = cmm.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    ComboBox3.Items.Add(rd.Item("namaBarang"))
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
        Me.ComboBox3.Text = Me.DataGridView1.Item(6, DataGridView1.CurrentRow.Index).Value
        Me.TextBox4.Text = Me.DataGridView1.Item(7, DataGridView1.CurrentRow.Index).Value
        Me.TextBox5.Text = Me.DataGridView1.Item(8, DataGridView1.CurrentRow.Index).Value
        Me.TextBox6.Text = Me.DataGridView1.Item(9, DataGridView1.CurrentRow.Index).Value
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
        If TextBox1.Text = "" Or ComboBox1.Text = "" Or TextBox2.Text = "" Or ComboBox2.Text = "" Or TextBox3.Text = "" Or
           ComboBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
            MsgBox("Data Masih Kosong", MsgBoxStyle.Information, "Perhatian!")
        Else
            Dim cmd As New SqlCommand("tambahBeli", koneksi)
            Dim xParam As New SqlParameter
            xParam.Direction = ParameterDirection.Input
            Try
                koneksi.Open()
                With cmd
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@noFaktur", SqlDbType.Char, 10).Value = TextBox1.Text
                    .Parameters.Add("@tglBeli", SqlDbType.DateTime).Value = DateTimePicker1.Text
                    .Parameters.Add("@idUser", SqlDbType.Char, 10).Value = ComboBox1.Text
                    .Parameters.Add("@idPemasok", SqlDbType.Char, 10).Value = TextBox2.Text
                    .Parameters.Add("@namaPemasok", SqlDbType.VarChar, 50).Value = ComboBox2.Text
                    .Parameters.Add("@idBarang", SqlDbType.Char, 10).Value = TextBox3.Text
                    .Parameters.Add("@namaBarang", SqlDbType.VarChar, 50).Value = ComboBox3.Text
                    .Parameters.Add("@hargaBeli", SqlDbType.Money).Value = TextBox4.Text
                    .Parameters.Add("@jlhTerima", SqlDbType.Int).Value = TextBox5.Text
                    .Parameters.Add("@total", SqlDbType.Money).Value = TextBox6.Text
                    .ExecuteNonQuery()
                End With
                MsgBox("Disimpan!", MsgBoxStyle.Information, "Perhatian!")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
            Finally
                koneksi.Close()
            End Try
            Call tampilBeli(DataGridView1)
        End If
        Call bersih()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim cmd As New SqlCommand("ubahBeli", koneksi)
        Dim xParam As New SqlParameter
        xParam.Direction = ParameterDirection.Input
        Try
            koneksi.Open()
            With cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add("@noFaktur", SqlDbType.Char, 10).Value = TextBox1.Text
                .Parameters.Add("@tglBeli", SqlDbType.DateTime).Value = DateTimePicker1.Text
                .Parameters.Add("@idUser", SqlDbType.Char, 10).Value = ComboBox1.Text
                .Parameters.Add("@idPemasok", SqlDbType.Char, 10).Value = TextBox2.Text
                .Parameters.Add("@namaPemasok", SqlDbType.VarChar, 50).Value = ComboBox2.Text
                .Parameters.Add("@idBarang", SqlDbType.Char, 10).Value = TextBox3.Text
                .Parameters.Add("@namaBarang", SqlDbType.VarChar, 50).Value = ComboBox3.Text
                .Parameters.Add("@hargaBeli", SqlDbType.Money).Value = TextBox4.Text
                .Parameters.Add("@jlhTerima", SqlDbType.Int).Value = TextBox5.Text
                .Parameters.Add("@total", SqlDbType.Money).Value = TextBox6.Text
                .ExecuteNonQuery()
            End With
            MsgBox("Diubah!", MsgBoxStyle.Information, "Perhatian!")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
        Finally
            koneksi.Close()
        End Try
        Call tampilBeli(DataGridView1)
        Call bersih()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If MessageBox.Show("Yakin Akan Dihapus?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Dim cmd As New SqlCommand("hapusBeli", koneksi)
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
            Call tampilBeli(DataGridView1)
            Call bersih()
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim sSql As String
        sSql = "select * from tbPemasok where namaPemasok='" & ComboBox2.Text & "'"
        Dim cmd As New SqlCommand(sSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    TextBox2.Text = rd.Item("idPemasok")
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim sSql As String
        sSql = "select * from tbBarang where namaBarang='" & ComboBox3.Text & "'"
        Dim cmd As New SqlCommand(sSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    TextBox3.Text = rd.Item("idBarang")
                    TextBox4.Text = rd.Item("hargaBeli")
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()
    End Sub

    Private Sub TextBox5_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox5.KeyPress
        TextBox6.Text = Val(TextBox4.Text) * Val(TextBox5.Text)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        formMenuUtama.Show()
    End Sub
End Class