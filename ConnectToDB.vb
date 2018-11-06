Imports System.Data.SqlClient

Public Class ConnectToDB
    Public reader As SqlDataReader
    Public sqlConn As SqlConnection
    Public cmd As SqlCommand
    Public table As New DataTable
    Public adapter As New SqlDataAdapter
    Public bindingSource As New BindingSource

    Public Function GetConnection(query As String) As SqlDataReader
        sqlConn = New SqlConnection()
        sqlConn.ConnectionString = "server=(local);database=TaxInvoice;Trusted_Connection=True"
        Try
            sqlConn.Open()
            cmd = New SqlCommand(query, sqlConn)

            reader = cmd.ExecuteReader
            'adapter.SelectCommand = cmd
            'adapter.Fill(table)
            'BindingSource.DataSource = table
            ''DataGridView1.DataSource = BindingSource
            Return reader
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
End Class

