Imports System.Data.SqlClient
Module Module1
    Public conn As SqlConnection
    Public da As SqlDataAdapter
    Public ds As New DataSet
    Public cmd As SqlCommand
    Public rd As SqlDataReader
    Public str As Strin
    Public Sub koneksi()aplikasiSIA;Integrated Security=True"
        str = "Data Source=DESKTOP-KC1E6QK\SQLEXPRESS;Initial Catalog=apl
        conn = New SqlConnection(str)
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
    End Sub
End Module