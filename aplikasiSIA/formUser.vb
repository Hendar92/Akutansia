Imports System.Data.SqlClient
Public Class formUser
    Dim str = "Data Source=DESKTOP-KC1E6QK\SQLEXPRESS;Initial Catalog=aplikasiSIA;Integrated Security=True"
    Dim koneksi As New SqlConnection(str)

    Sub bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        Button1.Enabled = True
        Button2.Enabled = True
        Button3.Enabled = False
        Button4.Enabled = False
        Button5.Enabled = True
        TextBox1.Focus()
    End Sub

    Protected Sub tampil(ByVal xGrid As DataGridView)
        With xGrid
            .ColumnCount = 3
            .Columns(0).Name = "ID User"
            .Columns(1).Name = "Username"
            .Columns(2).Name = "Password"
            .Rows.Clear()
        End With
        Dim sSql As String
        sSql = "Select * from tbUser"
        Dim cmd As New SqlCommand(sSql, koneksi)
        Try
            koneksi.Open()
            Dim rd As SqlDataReader = cmd.ExecuteReader
            If rd.HasRows Then
                While rd.Read
                    Dim baris(2) As String
                    baris(0) = rd(0) 'ID User
                    baris(1) = rd(1) 'Username
                    baris(2) = rd(2) 'Password
                    xGrid.Rows.Add(baris)
                End While
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        koneksi.Close()
    End Sub

    Private Sub formUser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call tampil(DataGridView1)
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Me.TextBox1.Text = Me.DataGridView1.Item(0, DataGridView1.CurrentRow.Index).Value
        Me.TextBox2.Text = Me.DataGridView1.Item(1, DataGridView1.CurrentRow.Index).Value
        Me.TextBox3.Text = Me.DataGridView1.Item(2, DataGridView1.CurrentRow.Index).Value
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
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
            MsgBox("Data Masih Kosong", MsgBoxStyle.Information, "Perhatian!")
        Else
            Dim cmd As New SqlCommand("tambah", koneksi)
            Dim xParam As New SqlParameter
            xParam.Direction = ParameterDirection.Input
            Try
                koneksi.Open()
                With cmd
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@idUser", SqlDbType.Char, 10).Value = TextBox1.Text
                    .Parameters.Add("@namaUser", SqlDbType.VarChar, 50).Value = TextBox2.Text
                    .Parameters.Add("@pass", SqlDbType.VarChar, 50).Value = TextBox3.Text
                    .ExecuteNonQuery()
                End With
                MsgBox("Disimpan!", MsgBoxStyle.Information, "Perhatian!")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
            Finally
                koneksi.Close()
            End Try
            Call tampil(DataGridView1)
        End If
        Call bersih()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim cmd As New SqlCommand("ubah", koneksi)
        Dim xParam As New SqlParameter
        xParam.Direction = ParameterDirection.Input
        Try
            koneksi.Open()
            With cmd
                .CommandType = CommandType.StoredProcedure
                .Parameters.Add("@idUser", SqlDbType.Char, 10).Value = TextBox1.Text
                .Parameters.Add("@namaUser", SqlDbType.VarChar, 50).Value = TextBox2.Text
                .Parameters.Add("@pass", SqlDbType.VarChar, 50).Value = TextBox3.Text
                .ExecuteNonQuery()
            End With
            MsgBox("Diubah!", MsgBoxStyle.Information, "Perhatian!")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
        Finally
            koneksi.Close()
        End Try
        Call tampil(DataGridView1)
        Call bersih()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If MessageBox.Show("Yakin Akan Dihapus?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
            Dim cmd As New SqlCommand("hapus", koneksi)
            Dim xParam As New SqlParameter
            xParam.Direction = ParameterDirection.Input
            Try
                koneksi.Open()
                With cmd
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@idUser", SqlDbType.Char, 10).Value = TextBox1.Text
                    .ExecuteNonQuery()
                End With
                MsgBox("Dihapus!", MsgBoxStyle.Information, "Perhatian!")
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
            Finally
                koneksi.Close()
            End Try
            Call tampil(DataGridView1)
            Call bersih()
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        formMenuUtama.Show()
    End Sub
End Class